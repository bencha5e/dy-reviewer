"""Stage 0 - locate the live OSAR tab, its DY-test column, and the period.

Row numbers drift every quarter as columns and tabs change, so nothing here is
hardcoded to a cell address. Lines are found by their **label**, which is stable
across the four models even though the label column is not: Strada, Campus and
Hialeah put labels in column C, Ares puts them in column B.
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field
from enum import Enum

from openpyxl.utils import column_index_from_string, get_column_letter

from .workbook import Workbook


class Line(str, Enum):
    """Canonical OSAR lines the checks refer to."""

    UPB = "UPB"
    NRSF = "NRSF"
    RESERVE_RATE = "RESERVE_RATE"
    STMT_END = "STMT_END"
    OCCUPANCY = "OCCUPANCY"
    GPR = "GPR"
    VACANCY = "VACANCY"
    BASE_RENT = "BASE_RENT"
    REIMBURSEMENT = "REIMBURSEMENT"
    PERCENTAGE_RENT = "PERCENTAGE_RENT"
    PARKING = "PARKING"
    OTHER_INCOME = "OTHER_INCOME"
    EGI = "EGI"
    TAX = "TAX"
    INSURANCE = "INSURANCE"
    MGMT_FEE = "MGMT_FEE"
    TOTAL_OPEX = "TOTAL_OPEX"
    NOI = "NOI"
    CAPEX = "CAPEX"
    TOTAL_CAPITAL = "TOTAL_CAPITAL"
    NCF = "NCF"
    DEBT_YIELD = "DEBT_YIELD"


#: Ordered (line, pattern) pairs searched against the normalized label.
#: Patterns are applied with `search`, so `^` is written explicitly wherever the
#: label must start with the phrase. Vacancy is deliberately unanchored: the
#: template writes it as "Less: Vacancy Loss".
#: Order matters where one label is a prefix of another.
_LINE_PATTERNS: list[tuple[Line, re.Pattern]] = [
    (Line.UPB, re.compile(r"^note a-?\s*scheduled loan balance")),
    (Line.NRSF, re.compile(r"^current net rentable")),
    (Line.RESERVE_RATE, re.compile(r"^cap ?ex reserve")),
    (Line.STMT_END, re.compile(r"^statement ending date")),
    (Line.OCCUPANCY, re.compile(r"^occupancy rate")),
    (Line.GPR, re.compile(r"^gross potential rent")),
    (Line.VACANCY, re.compile(r"vacancy loss")),
    (Line.BASE_RENT, re.compile(r"^base rent")),
    (Line.REIMBURSEMENT, re.compile(r"^expense reimbursement")),
    (Line.PERCENTAGE_RENT, re.compile(r"^percentage rent")),
    (Line.PARKING, re.compile(r"^parking income")),
    (Line.OTHER_INCOME, re.compile(r"^other income")),
    (Line.EGI, re.compile(r"^effective gross income")),
    (Line.TAX, re.compile(r"^real estate taxes")),
    (Line.INSURANCE, re.compile(r"^property insurance")),
    (Line.MGMT_FEE, re.compile(r"^management fees?")),
    (Line.TOTAL_OPEX, re.compile(r"^total operating expenses")),
    # Must not match the DSCR rows, which mention NOI inside parentheses.
    (Line.NOI, re.compile(r"^net operating income")),
    (Line.CAPEX, re.compile(r"^capital expenditures")),
    (Line.TOTAL_CAPITAL, re.compile(r"^total capital items")),
    (Line.NCF, re.compile(r"^net cash flow$")),
    (Line.DEBT_YIELD, re.compile(r"^debt yield")),
]

#: Sheet names that can carry the OSAR output grid.
_OSAR_RE = re.compile(r"\bosar\b", re.IGNORECASE)

#: Columns that can hold figures; A-C are labels, and nothing useful sits past N.
_FIGURE_COLS = [get_column_letter(i) for i in range(column_index_from_string("D"), column_index_from_string("N") + 1)]

_LABEL_COLS = ("B", "C")

_EXCEL_EPOCH = dt.date(1899, 12, 30)


def normalize_label(raw) -> str:
    """Reduce an OSAR label to a comparable form.

    Strips the leading `*` emphasis the template uses for subtotals, footnote
    markers like `(3)`, indentation, and trailing punctuation.
    """
    if raw is None:
        return ""
    text = str(raw)
    text = text.replace("*", " ")
    text = re.sub(r"\(\d+\)", " ", text)          # footnote markers
    text = text.replace(" ", " ")
    text = re.sub(r"\s+", " ", text).strip()
    text = text.strip(" :.-")
    return text.lower()


def excel_to_date(value) -> dt.date | None:
    """Coerce a cached cell value to a date, accepting serials and datetimes."""
    if isinstance(value, dt.datetime):
        return value.date()
    if isinstance(value, dt.date):
        return value
    if isinstance(value, (int, float)) and 20000 < float(value) < 80000:
        return _EXCEL_EPOCH + dt.timedelta(days=int(value))
    if isinstance(value, str):
        m = re.search(r"(\d{1,2})/(\d{1,2})/(\d{4})", value)
        if m:
            month, day, year = (int(g) for g in m.groups())
            try:
                return dt.date(year, month, day)
            except ValueError:
                return None
    return None


def quarter_end(date: dt.date) -> dt.date:
    """The quarter-end date for the quarter containing `date`."""
    q_end_month = ((date.month - 1) // 3) * 3 + 3
    last_day = {3: 31, 6: 30, 9: 30, 12: 31}[q_end_month]
    return dt.date(date.year, q_end_month, last_day)


def is_quarter_end(date: dt.date | None) -> bool:
    return date is not None and date == quarter_end(date)


@dataclass
class OsarTab:
    """The resolved output tab plus everything Stage 0 learned about it."""

    sheet: str
    dy_column: str
    label_column: str
    rows: dict[Line, int] = field(default_factory=dict)
    labels: dict[int, str] = field(default_factory=dict)
    raw_labels: dict[int, str] = field(default_factory=dict)
    period_end: dt.date | None = None
    period_source: str | None = None
    notes: list[str] = field(default_factory=list)
    #: Other visible OSAR tabs that were considered and rejected.
    rejected: list[str] = field(default_factory=list)
    hidden_osar_tabs: list[str] = field(default_factory=list)

    def cell(self, line: Line, column: str | None = None) -> str | None:
        """`I25`-style coordinate for a line in the DY column (or another column)."""
        row = self.rows.get(line)
        if row is None:
            return None
        return f"{column or self.dy_column}{row}"

    def has(self, line: Line) -> bool:
        return line in self.rows


def find_osar_candidates(wb: Workbook) -> tuple[list[str], list[str]]:
    """Return (visible OSAR tabs, hidden OSAR tabs)."""
    visible, hidden = [], []
    for name in wb.sheet_names:
        if not _OSAR_RE.search(name):
            continue
        (visible if wb.is_visible(name) else hidden).append(name)
    return visible, hidden


def _build_label_map(wb: Workbook, sheet: str, max_row: int = 140) -> tuple[str, dict[int, str], dict[int, str]]:
    """Find the label column and read the labels out of it.

    Scans both candidate columns and keeps whichever yields more recognizable
    OSAR lines - Ares labels in column B, the others in column C.
    """
    best_col, best_hits, best_norm, best_raw = _LABEL_COLS[0], -1, {}, {}
    for col in _LABEL_COLS:
        norm: dict[int, str] = {}
        raw: dict[int, str] = {}
        hits = 0
        for row in range(1, max_row + 1):
            text = wb.text(sheet, f"{col}{row}")
            if not text:
                continue
            n = normalize_label(text)
            if not n:
                continue
            norm[row] = n
            raw[row] = text
            if any(p.search(n) for _, p in _LINE_PATTERNS):
                hits += 1
        if hits > best_hits:
            best_col, best_hits, best_norm, best_raw = col, hits, norm, raw
    return best_col, best_norm, best_raw


def _resolve_rows(labels: dict[int, str]) -> tuple[dict[Line, int], list[str]]:
    """Map each canonical line to its row, first match wins."""
    rows: dict[Line, int] = {}
    notes: list[str] = []
    for row in sorted(labels):
        n = labels[row]
        for line, pattern in _LINE_PATTERNS:
            if line in rows:
                continue
            if pattern.search(n):
                rows[line] = row
                break
    missing = [
        line.value
        for line in (Line.GPR, Line.EGI, Line.NOI, Line.NCF, Line.DEBT_YIELD)
        if line not in rows
    ]
    if missing:
        notes.append(f"OSAR labels not found for: {', '.join(missing)}")
    return rows, notes


#: A debt yield outside this band on the Debt Yield row marks the column as a
#: variance or helper column rather than the test figure itself.
_PLAUSIBLE_DY = (0.005, 0.5)


def _is_self_row_delta(wb: Workbook, sheet: str, coord: str, row: int) -> bool:
    """True when a formula only combines other cells on its own row.

    Analysts add helper columns like `N73 = I73-G73` (this quarter vs last) to
    the right of the figures; such a column carries a formula on the Debt Yield
    row without being a debt-yield figure, and must not win the rightmost rule.
    """
    from . import formula as F  # local import: formula.py is check-layer code

    refs = [r for r in F.iter_refs(wb.formula(sheet, coord)) if not r.is_range]
    if not refs:
        return False
    for ref in refs:
        m = re.match(r"^[A-Z]+(\d+)$", ref.coord)
        if ref.sheet is not None or not m or int(m.group(1)) != row:
            return False
    return True


def _detect_dy_column(wb: Workbook, sheet: str, rows: dict[Line, int]) -> tuple[str, list[str]]:
    """Identify the current DY-test figure column.

    Spec section 1 says the current figures live in the rightmost figure column
    (column I in the Q1 2026 set), but the column is derived rather than assumed:
    it is the rightmost figure column carrying a formula on the Debt Yield row.
    Columns that merely compare other columns on that row, or whose value is not
    a plausible debt yield, are variance helpers and are passed over.
    """
    notes: list[str] = []
    dy_row = rows.get(Line.DEBT_YIELD)
    candidates: list[str] = []
    if dy_row is not None:
        for col in _FIGURE_COLS:
            if wb.formula(sheet, f"{col}{dy_row}") is None:
                continue
            if _is_self_row_delta(wb, sheet, f"{col}{dy_row}", dy_row):
                notes.append(
                    f"Column {col} skipped: its Debt Yield formula only compares other "
                    f"columns on the same row (a variance helper, not a figure)."
                )
                continue
            candidates.append(col)
        plausible = [
            col
            for col in candidates
            if (value := wb.number(sheet, f"{col}{dy_row}")) is not None
            and _PLAUSIBLE_DY[0] <= abs(value) <= _PLAUSIBLE_DY[1]
        ]
        if plausible and len(plausible) < len(candidates):
            skipped = [c for c in candidates if c not in plausible]
            notes.append(
                f"Column(s) {', '.join(skipped)} skipped: Debt Yield value outside the "
                f"plausible band, so they read as variance columns."
            )
            candidates = plausible
    if not candidates:
        # Fall back to the NCF row, then to the spec's stated default.
        ncf_row = rows.get(Line.NCF)
        if ncf_row is not None:
            for col in _FIGURE_COLS:
                if wb.formula(sheet, f"{col}{ncf_row}") is not None:
                    candidates.append(col)
        if candidates:
            notes.append("Debt Yield row carried no formula; DY column derived from the NCF row.")
    if not candidates:
        notes.append("Could not derive the DY column from formulas; defaulting to column I per spec section 1.")
        return "I", notes

    chosen = candidates[-1]
    if chosen != "I":
        notes.append(
            f"DY-test column resolved to {chosen}, not the column I the spec's Q1 2026 "
            f"appendix describes - the cell map has drifted."
        )
    return chosen, notes


def _identify_period(wb: Workbook, sheet: str, rows: dict[Line, int], dy_col: str) -> tuple[dt.date | None, str | None]:
    """Read the statement ending date from the DY column."""
    row = rows.get(Line.STMT_END)
    if row is None:
        return None, None
    for col in (dy_col, "H", "F"):
        coord = f"{col}{row}"
        date = excel_to_date(wb.value(sheet, coord))
        if date is not None:
            return date, coord
    return None, f"{dy_col}{row}"


def select_osar(wb: Workbook) -> OsarTab:
    """Pick the live OSAR tab and read its structure.

    Tab-selection rule (spec section 3, Stage 0): use the OSAR tab that is
    visible and ignore hidden ones. Only when more than one is visible do we
    disambiguate by a populated DY column at quarter-end.
    """
    visible, hidden = find_osar_candidates(wb)
    if not visible:
        raise ValueError(
            f"No visible OSAR tab in {wb.path.name}; sheets: {', '.join(wb.sheet_names)}"
        )

    notes: list[str] = []
    rejected: list[str] = []

    if len(visible) == 1:
        chosen = visible[0]
    else:
        scored: list[tuple[int, str]] = []
        for name in visible:
            col, labels, _ = _build_label_map(wb, name)
            rows, _ = _resolve_rows(labels)
            dy_col, _ = _detect_dy_column(wb, name, rows)
            populated = sum(
                1
                for line in (Line.GPR, Line.EGI, Line.NOI, Line.NCF)
                if (c := rows.get(line)) and wb.value(name, f"{dy_col}{c}") is not None
            )
            period, _ = _identify_period(wb, name, rows, dy_col)
            score = populated * 2 + (3 if is_quarter_end(period) else 0)
            scored.append((score, name))
        scored.sort(key=lambda t: t[0], reverse=True)
        chosen = scored[0][1]
        rejected = [n for _, n in scored[1:]]
        notes.append(
            f"{len(visible)} visible OSAR tabs; selected {chosen!r} on populated "
            f"DY column at quarter-end (also considered: {', '.join(rejected)})."
        )

    label_col, labels, raw_labels = _build_label_map(wb, chosen)
    rows, row_notes = _resolve_rows(labels)
    notes.extend(row_notes)
    dy_col, col_notes = _detect_dy_column(wb, chosen, rows)
    notes.extend(col_notes)
    period, period_src = _identify_period(wb, chosen, rows, dy_col)

    return OsarTab(
        sheet=chosen,
        dy_column=dy_col,
        label_column=label_col,
        rows=rows,
        labels=labels,
        raw_labels=raw_labels,
        period_end=period,
        period_source=period_src,
        notes=notes,
        rejected=rejected,
        hidden_osar_tabs=hidden,
    )
