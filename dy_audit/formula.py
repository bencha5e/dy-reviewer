"""Formula parsing, reference resolution, and a restricted evaluator.

Three jobs, all in service of the blocker checks:

1. **Reference extraction** - what cells does this formula touch?
2. **Pass-through resolution** - an OSAR line often does not hold its own logic.
   Ares routes `Comm OSAR!I39` through a two-level rollup on `Lender Calc`
   (`I39 -> E20 =SUM(E21) -> E21 =MAX(...)`), so a check that reads only the
   OSAR cell sees a bare reference and misses the defect underneath it.
3. **Evaluation under substitution** - `CHK_VACANCY_SIGN` cannot be judged at
   today's occupancy. Strada's inverted branch returns 0 at 91.95% occupancy and
   looks perfectly healthy; the fault only appears above 95%. The evaluator lets
   a check re-run a formula with a synthetic occupancy.

The evaluator deliberately supports a small grammar. Anything outside it raises
`UnsupportedFormula` so the caller can report MANUAL_REVIEW rather than guess.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from typing import Callable

from openpyxl.utils import column_index_from_string, get_column_letter

from .workbook import ERROR_VALUES, Workbook


class UnsupportedFormula(Exception):
    """The evaluator met syntax or a function outside its restricted grammar."""


class EvalError(Exception):
    """The formula parsed but could not produce a value (bad ref, error cell)."""


# --------------------------------------------------------------------------
# Reference extraction
# --------------------------------------------------------------------------

_SHEET = r"(?:'(?:[^']|'')*'|\[\d+\][A-Za-z0-9_. ()&-]*|[A-Za-z0-9_.]+)"
_CELL = r"\$?[A-Za-z]{1,3}\$?\d{1,7}"
_COLRANGE = r"\$?[A-Za-z]{1,3}:\$?[A-Za-z]{1,3}"

#: A reference, optionally sheet-qualified. The lookbehind stops `1E5` from
#: being read as a reference to cell E5.
_REF_RE = re.compile(
    rf"(?<![A-Za-z0-9_.])(?:(?P<sheet>{_SHEET})!)?"
    rf"(?P<body>{_CELL}(?::{_CELL})?|{_COLRANGE})"
)

_STRING_RE = re.compile(r'"(?:[^"]|"")*"')
_EXTERNAL_RE = re.compile(r"\[(\d+)\]")
_CELL_ONLY_RE = re.compile(rf"^{_CELL}$")


@dataclass(frozen=True)
class Ref:
    """One reference found in a formula."""

    sheet: str | None
    body: str
    external_index: int | None = None

    @property
    def is_range(self) -> bool:
        return ":" in self.body

    @property
    def is_whole_column(self) -> bool:
        return self.is_range and not any(ch.isdigit() for ch in self.body)

    @property
    def coord(self) -> str:
        """Single-cell coordinate with `$` stripped; empty for ranges."""
        return "" if self.is_range else self.body.replace("$", "").upper()

    def resolved_sheet(self, default: str) -> str:
        return self.sheet or default

    def __str__(self) -> str:
        return f"{self.sheet}!{self.body}" if self.sheet else self.body


def strip_strings(formula: str) -> str:
    """Blank out double-quoted literals so refs inside text are not matched."""
    return _STRING_RE.sub(lambda m: '"' + " " * (len(m.group(0)) - 2) + '"', formula)


def normalize(formula: str | None) -> str:
    """Drop the leading `=`, the template's cosmetic `+`, and `_xlfn.` prefixes.

    Excel stores newer functions with an `_xlfn.` marker, so Campus's XLOOKUP
    arrives as `_xlfn.XLOOKUP(...)`.
    """
    if not formula:
        return ""
    text = formula.strip()
    if text.startswith("="):
        text = text[1:]
    text = text.replace("_xlfn._xlws.", "").replace("_xlfn.", "")
    return text.strip()


def unquote_sheet(name: str | None) -> str | None:
    if name is None:
        return None
    if name.startswith("'") and name.endswith("'"):
        name = name[1:-1].replace("''", "'")
    return name


def iter_refs(formula: str | None) -> list[Ref]:
    """Every reference in a formula, in source order."""
    text = strip_strings(normalize(formula))
    refs: list[Ref] = []
    for m in _REF_RE.finditer(text):
        sheet_raw = m.group("sheet")
        external = None
        if sheet_raw:
            ext = _EXTERNAL_RE.search(sheet_raw)
            if ext:
                external = int(ext.group(1))
                sheet_raw = _EXTERNAL_RE.sub("", sheet_raw)
        refs.append(Ref(unquote_sheet(sheet_raw), m.group("body"), external))
    return refs


def function_names(formula: str | None) -> list[str]:
    """Uppercase names of functions called in the formula."""
    text = strip_strings(normalize(formula))
    return [m.group(1).upper() for m in re.finditer(r"([A-Za-z][A-Za-z0-9_.]*)\s*\(", text)]


def has_external_refs(formula: str | None) -> bool:
    """True when the formula links to another workbook (`[1]Sheet!A1`)."""
    return any(r.external_index is not None for r in iter_refs(formula))


# --------------------------------------------------------------------------
# Pass-through resolution
# --------------------------------------------------------------------------

#: `=+X`, `=X`, `='Tab'!X` - a bare reference and nothing else.
_PASSTHROUGH_REF_RE = re.compile(
    rf"^\+?\s*(?:(?P<sheet>{_SHEET})!)?(?P<body>{_CELL})$"
)
#: `=SUM(X)` over exactly one cell. Ares's Lender Calc uses this as its rollup
#: layer, so it is a hop rather than a computation.
_PASSTHROUGH_SUM_RE = re.compile(
    rf"^\+?\s*SUM\(\s*(?:(?P<sheet>{_SHEET})!)?(?P<body>{_CELL})\s*\)$",
    re.IGNORECASE,
)


@dataclass
class Resolution:
    """Where a line's real logic lives, and how we got there."""

    sheet: str
    coord: str
    formula: str | None
    #: `Sheet!Cell` hops walked, starting at the cell originally asked about.
    path: list[str] = field(default_factory=list)
    value: object = None
    truncated: bool = False

    @property
    def body(self) -> str:
        return normalize(self.formula)

    @property
    def ref(self) -> str:
        return f"{self.sheet}!{self.coord}"

    @property
    def hopped(self) -> bool:
        return len(self.path) > 1


def _passthrough_target(body: str) -> tuple[str | None, str] | None:
    for pattern in (_PASSTHROUGH_REF_RE, _PASSTHROUGH_SUM_RE):
        m = pattern.match(body)
        if m:
            return unquote_sheet(m.group("sheet")), m.group("body").replace("$", "").upper()
    return None


def resolve_defining_formula(
    wb: Workbook, sheet: str, coord: str, max_hops: int = 8
) -> Resolution:
    """Follow pass-through references to the cell that actually does the work.

    Stops at the first formula that computes something - a MAX, an IF, a lookup,
    an arithmetic expression - or at a literal value.
    """
    path: list[str] = []
    cur_sheet, cur_coord = sheet, coord.replace("$", "").upper()

    for _ in range(max_hops):
        path.append(f"{cur_sheet}!{cur_coord}")
        formula = wb.formula(cur_sheet, cur_coord)
        if formula is None:
            return Resolution(cur_sheet, cur_coord, None, path, wb.value(cur_sheet, cur_coord))
        target = _passthrough_target(normalize(formula))
        if target is None:
            return Resolution(
                cur_sheet, cur_coord, formula, path, wb.value(cur_sheet, cur_coord)
            )
        next_sheet, next_coord = target
        next_sheet = next_sheet or cur_sheet
        if f"{next_sheet}!{next_coord}" in path:  # cycle guard
            break
        if not wb.has_sheet(next_sheet):
            return Resolution(
                cur_sheet, cur_coord, formula, path, wb.value(cur_sheet, cur_coord)
            )
        cur_sheet, cur_coord = next_sheet, next_coord

    return Resolution(
        cur_sheet,
        cur_coord,
        wb.formula(cur_sheet, cur_coord),
        path,
        wb.value(cur_sheet, cur_coord),
        truncated=True,
    )


def dependencies(
    wb: Workbook, sheet: str, coord: str, max_depth: int = 6
) -> dict[str, tuple[str, str]]:
    """Transitive single-cell dependencies of a formula.

    Returns `{"Sheet!Coord": (sheet, coord)}`. Ranges are skipped: nothing on the
    occupancy path in these models depends on one, and expanding whole-column
    references would be both slow and meaningless.
    """
    found: dict[str, tuple[str, str]] = {}
    frontier = [(sheet, coord.replace("$", "").upper(), 0)]
    seen: set[str] = set()

    while frontier:
        cur_sheet, cur_coord, depth = frontier.pop()
        key = f"{cur_sheet}!{cur_coord}"
        if key in seen or depth > max_depth:
            continue
        seen.add(key)
        formula = wb.formula(cur_sheet, cur_coord)
        if formula is None:
            continue
        for ref in iter_refs(formula):
            if ref.is_range or ref.external_index is not None:
                continue
            dep_sheet = ref.resolved_sheet(cur_sheet)
            if not wb.has_sheet(dep_sheet):
                continue
            dep_key = f"{dep_sheet}!{ref.coord}"
            found[dep_key] = (dep_sheet, ref.coord)
            frontier.append((dep_sheet, ref.coord, depth + 1))
    return found


# --------------------------------------------------------------------------
# Restricted evaluator
# --------------------------------------------------------------------------

#: Value returned for a blank cell or an empty-string formula result. Excel
#: treats a blank vacancy line as no deduction, so this must behave as 0.
BLANK = ""

Resolver = Callable[[str, str], object]


def to_number(value) -> float:
    """Coerce a cell value to a number the way Excel would in arithmetic."""
    if value is None or value == BLANK:
        return 0.0
    if isinstance(value, bool):
        return 1.0 if value else 0.0
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.strip()
        if text in ERROR_VALUES:
            raise EvalError(f"formula depends on an error cell ({text})")
        if not text:
            return 0.0
        try:
            return float(text.replace(",", "").replace("$", "").rstrip("%"))
        except ValueError as exc:
            raise EvalError(f"non-numeric value {value!r}") from exc
    raise EvalError(f"cannot coerce {value!r} to a number")


class _Parser:
    """Recursive-descent evaluator over the restricted grammar."""

    _FUNCS = {"IF", "MAX", "MIN", "SUM", "ABS", "ROUND", "IFERROR", "AND", "OR", "NOT"}

    def __init__(self, text: str, sheet: str, resolver: Resolver):
        self.text = text
        self.pos = 0
        self.sheet = sheet
        self.resolve = resolver

    # -- lexer helpers --

    def _skip(self) -> None:
        while self.pos < len(self.text) and self.text[self.pos] in " \t\n\r":
            self.pos += 1

    def _peek(self, *literals: str) -> str | None:
        self._skip()
        for lit in literals:
            if self.text.startswith(lit, self.pos):
                return lit
        return None

    def _take(self, literal: str) -> None:
        self._skip()
        if not self.text.startswith(literal, self.pos):
            raise UnsupportedFormula(f"expected {literal!r} at offset {self.pos}")
        self.pos += len(literal)

    # -- grammar --

    def parse(self):
        value = self.comparison()
        self._skip()
        if self.pos != len(self.text):
            raise UnsupportedFormula(f"trailing input at offset {self.pos}: {self.text[self.pos:]!r}")
        return value

    def comparison(self):
        left = self.additive()
        while (op := self._peek("<=", ">=", "<>", "=", "<", ">")) is not None:
            self.pos += len(op)
            right = self.additive()
            left = self._compare(op, left, right)
        return left

    @staticmethod
    def _compare(op: str, left, right) -> bool:
        if isinstance(left, str) and isinstance(right, str) and op in ("=", "<>"):
            same = left.strip().lower() == right.strip().lower()
            return same if op == "=" else not same
        a, b = to_number(left), to_number(right)
        return {
            "=": a == b,
            "<>": a != b,
            "<": a < b,
            ">": a > b,
            "<=": a <= b,
            ">=": a >= b,
        }[op]

    def additive(self):
        value = to_number(self.multiplicative())
        while (op := self._peek("+", "-")) is not None:
            self.pos += 1
            rhs = to_number(self.multiplicative())
            value = value + rhs if op == "+" else value - rhs
        return value

    def multiplicative(self):
        value = to_number(self.unary())
        while (op := self._peek("*", "/")) is not None:
            self.pos += 1
            rhs = to_number(self.unary())
            if op == "*":
                value *= rhs
            else:
                if rhs == 0:
                    raise EvalError("division by zero")
                value /= rhs
        return value

    def unary(self):
        if (op := self._peek("-", "+")) is not None:
            self.pos += 1
            value = to_number(self.unary())
            return -value if op == "-" else value
        return self.power()

    def power(self):
        value = self.primary()
        if self._peek("^") is not None:
            self.pos += 1
            value = to_number(value) ** to_number(self.unary())
        return value

    def primary(self):
        self._skip()
        if self.pos >= len(self.text):
            raise UnsupportedFormula("unexpected end of formula")
        ch = self.text[self.pos]

        if ch == "(":
            self._take("(")
            value = self.comparison()
            self._take(")")
            return value

        if ch == '"':
            return self._string()

        if ch.isdigit() or (ch == "." and self.pos + 1 < len(self.text) and self.text[self.pos + 1].isdigit()):
            return self._number()

        m = _REF_RE.match(self.text, self.pos)
        func = re.match(r"([A-Za-z][A-Za-z0-9_.]*)\s*\(", self.text[self.pos :])
        # A function call wins over a reference match: TRUE(, SUM( and friends
        # can look like the start of a reference.
        if func and (not m or func.start() + self.pos <= m.start()):
            return self._function(func.group(1).upper())
        if m:
            self.pos = m.end()
            return self._reference(m)

        word = re.match(r"(TRUE|FALSE)\b", self.text[self.pos :], re.IGNORECASE)
        if word:
            self.pos += word.end()
            return word.group(1).upper() == "TRUE"

        raise UnsupportedFormula(f"cannot parse at offset {self.pos}: {self.text[self.pos:][:40]!r}")

    def _string(self):
        m = _STRING_RE.match(self.text, self.pos)
        if not m:
            raise UnsupportedFormula("unterminated string literal")
        self.pos = m.end()
        return m.group(0)[1:-1].replace('""', '"')

    def _number(self):
        m = re.match(r"\d*\.?\d+(?:[eE][+-]?\d+)?", self.text[self.pos :])
        if not m:
            raise UnsupportedFormula("bad number")
        self.pos += m.end()
        value = float(m.group(0))
        if self.pos < len(self.text) and self.text[self.pos] == "%":
            self.pos += 1
            value /= 100.0
        return value

    def _reference(self, m: re.Match):
        body = m.group("body")
        if ":" in body:
            raise UnsupportedFormula(f"range reference {body!r} outside SUM")
        sheet = unquote_sheet(m.group("sheet")) or self.sheet
        return self.resolve(sheet, body.replace("$", "").upper())

    def _args(self) -> list:
        self._take("(")
        args: list = []
        if self._peek(")") is not None:
            self._take(")")
            return args
        while True:
            args.append(self.comparison())
            if self._peek(",") is not None:
                self._take(",")
                continue
            self._take(")")
            return args

    def _function(self, name: str):
        if name not in self._FUNCS:
            raise UnsupportedFormula(f"unsupported function {name}()")
        self.pos += self.text[self.pos :].index("(")

        # IF and IFERROR must not evaluate the branch they do not take: the
        # untaken branch may reference an error cell or an unsupported function.
        if name == "IF":
            self._take("(")
            condition = self.comparison()
            self._take(",")
            if _truthy(condition):
                value = self.comparison()
                self._skip_remaining_args()
            else:
                self._skip_one_arg()
                if self._peek(",") is not None:
                    self._take(",")
                    value = self.comparison()
                    self._skip_remaining_args()
                else:
                    value = False
                    self._take(")")
            return value

        if name == "IFERROR":
            self._take("(")
            start = self.pos
            try:
                value = self.comparison()
            except (EvalError, UnsupportedFormula):
                self.pos = start
                self._skip_one_arg()
                self._take(",")
                value = self.comparison()
                self._skip_remaining_args()
                return value
            self._take(",")
            self._skip_one_arg()
            self._take(")")
            return value

        args = self._args()
        if name == "MAX":
            return max(to_number(a) for a in args)
        if name == "MIN":
            return min(to_number(a) for a in args)
        if name == "SUM":
            return sum(to_number(a) for a in args)
        if name == "ABS":
            return abs(to_number(args[0]))
        if name == "ROUND":
            digits = int(to_number(args[1])) if len(args) > 1 else 0
            return round(to_number(args[0]), digits)
        if name == "AND":
            return all(_truthy(a) for a in args)
        if name == "OR":
            return any(_truthy(a) for a in args)
        if name == "NOT":
            return not _truthy(args[0])
        raise UnsupportedFormula(f"unsupported function {name}()")

    def _skip_one_arg(self) -> None:
        """Advance past one argument without evaluating it."""
        depth = 0
        while self.pos < len(self.text):
            ch = self.text[self.pos]
            if ch == '"':
                m = _STRING_RE.match(self.text, self.pos)
                self.pos = m.end() if m else self.pos + 1
                continue
            if ch == "(":
                depth += 1
            elif ch == ")":
                if depth == 0:
                    return
                depth -= 1
            elif ch == "," and depth == 0:
                return
            self.pos += 1

    def _skip_remaining_args(self) -> None:
        while self._peek(",") is not None:
            self._take(",")
            self._skip_one_arg()
        self._take(")")


def _truthy(value) -> bool:
    if isinstance(value, bool):
        return value
    if isinstance(value, str):
        return value.strip().lower() == "true"
    return to_number(value) != 0


def evaluate(formula: str, sheet: str, resolver: Resolver):
    """Evaluate `formula` with references resolved through `resolver`."""
    return _Parser(normalize(formula), sheet, resolver).parse()


def make_resolver(
    wb: Workbook,
    overrides: dict[str, object] | None = None,
    recompute: set[str] | None = None,
    max_depth: int = 8,
) -> Resolver:
    """Build a resolver for `evaluate`.

    `overrides` pins specific cells to a supplied value (the synthetic occupancy).
    `recompute` names cells whose formulas must be re-evaluated because they
    depend on an override; every other cell uses the value Excel cached, which
    keeps lookups and other unsupported functions off the evaluation path.
    """
    overrides = overrides or {}
    recompute = recompute or set()
    cache: dict[str, object] = {}

    def resolve(sheet: str, coord: str, _depth: int = 0) -> object:
        key = f"{sheet}!{coord}"
        if key in overrides:
            return overrides[key]
        if key in cache:
            return cache[key]
        if key in recompute and _depth < max_depth:
            formula = wb.formula(sheet, coord)
            if formula is not None:
                cache[key] = 0.0  # cycle guard while this cell is in flight
                value = _Parser(
                    normalize(formula), sheet, lambda s, c: resolve(s, c, _depth + 1)
                ).parse()
                cache[key] = value
                return value
        value = wb.value(sheet, coord)
        if isinstance(value, str) and value.strip() in ERROR_VALUES:
            raise EvalError(f"{key} holds {value.strip()}")
        cache[key] = value
        return value

    return lambda sheet, coord: resolve(sheet, coord)


def extract_call_args(formula: str | None, func: str) -> list[str] | None:
    """Raw argument text of the first `func(...)` call, split at top level.

    Returns None when the function is not called. Used by the tax and insurance
    checks to look at each side of a `MAX(...)` separately.
    """
    text = normalize(formula)
    pattern = re.compile(rf"(?<![A-Za-z0-9_.]){re.escape(func)}\s*\(", re.IGNORECASE)
    m = pattern.search(text)
    if not m:
        return None

    args: list[str] = []
    depth = 0
    start = m.end()
    i = start
    while i < len(text):
        ch = text[i]
        if ch == '"':
            sm = _STRING_RE.match(text, i)
            i = sm.end() if sm else i + 1
            continue
        if ch == "(":
            depth += 1
        elif ch == ")":
            if depth == 0:
                args.append(text[start:i].strip())
                return [a for a in args if a]
            depth -= 1
        elif ch == "," and depth == 0:
            args.append(text[start:i].strip())
            start = i + 1
        i += 1
    return None


def expand_range(ref: Ref, default_sheet: str, limit: int = 500) -> list[tuple[str, str]]:
    """Expand a bounded range into cells; empty for whole-column references."""
    if not ref.is_range or ref.is_whole_column:
        return []
    start, end = ref.body.replace("$", "").upper().split(":")
    sm = re.match(r"([A-Z]+)(\d+)", start)
    em = re.match(r"([A-Z]+)(\d+)", end)
    if not sm or not em:
        return []
    c1, r1 = column_index_from_string(sm.group(1)), int(sm.group(2))
    c2, r2 = column_index_from_string(em.group(1)), int(em.group(2))
    if (abs(c2 - c1) + 1) * (abs(r2 - r1) + 1) > limit:
        return []
    sheet = ref.resolved_sheet(default_sheet)
    return [
        (sheet, f"{get_column_letter(c)}{r}")
        for c in range(min(c1, c2), max(c1, c2) + 1)
        for r in range(min(r1, r2), max(r1, r2) + 1)
    ]
