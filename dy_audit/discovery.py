"""Stage 0 - discover loan pairs in the input folder.

Spec section 0: match each Excel file to its definitions file by loan name parsed
out of the filename. Do not assume a fixed pairing order or a fixed count of
files; the loan roster changes as loans are added and paid off.

Real filenames in the Q1 2026 set use spaces and a "N. " prefix
(`1. Strada DY Test_1Q26_vF vBCS.xlsx`) while the spec's examples use
underscores (`1__Strada_DY_Test_1Q26_vF_vBCS.xlsx`). Both normalize the same way.
"""

from __future__ import annotations

import re
from difflib import SequenceMatcher
from pathlib import Path

from .model import LoanFiles

#: Filename tokens that carry no loan identity - version marks, quarter tags,
#: and the boilerplate that appears on every file in the folder.
_NOISE_TOKENS = {
    "dy",
    "test",
    "tests",
    "dydefinitions",
    "definitions",
    "definition",
    "vf",
    "vbcs",
    "bcs",
    "final",
    "draft",
    "copy",
}

#: Version marks (v2, v10) and period tags (1q26, 1q, 2026, 1q2026, q1).
_NOISE_PATTERNS = [
    re.compile(r"^v\d+$"),
    re.compile(r"^\d{1,2}q\d{0,4}$"),
    re.compile(r"^q\d{1,2}$"),
    re.compile(r"^\d{4}$"),
]

_DEFS_SUFFIXES = ("_dydefinitions", "-dydefinitions", " dydefinitions", "dydefinitions")

#: Leading "1. ", "1__", "1-", "1_" filename prefix.
_PREFIX_RE = re.compile(r"^\s*(\d+)\s*(?:[.)_-]+|__)\s*")

#: Loan name header inside the definitions file, e.g. `*****Strada*****`.
#: Applied after backslash unescaping - the source documents write the header as
#: `\*\*\*\*\*Strada\*\*\*\*\*`, so no two asterisks are adjacent until the
#: backslashes come out.
_HEADER_RE = re.compile(r"^\*{2,}\s*(.+?)\s*\*{2,}$")

#: The template placeholder that heads every definitions file before the real name.
_PLACEHOLDER_NAMES = {"loan_name", "loanname", "definition from loan agreement"}


def _strip_prefix(stem: str) -> tuple[str | None, str]:
    """Split a leading numeric filename prefix off the stem."""
    m = _PREFIX_RE.match(stem)
    if m:
        return m.group(1), stem[m.end() :]
    return None, stem


def _split_camel(text: str) -> list[str]:
    """Split camelCase / PascalCase: CampusAtVillaLaJolla -> [campus, at, villa, la, jolla]."""
    spaced = re.sub(r"(?<=[a-z0-9])(?=[A-Z])", " ", text)
    return [t for t in re.split(r"[^A-Za-z0-9]+", spaced.lower()) if t]


def _tokenize(text: str) -> list[str]:
    """Identifying tokens from a filename fragment.

    Noise is dropped on the *punctuation-split* tokens before camelCase
    splitting, because splitting first would shred version marks into letters -
    `vF vBCS` becomes `v f v bcs`, none of which look like noise on their own.
    """
    tokens: list[str] = []
    for raw in re.split(r"[^A-Za-z0-9]+", text):
        if not raw or _is_noise(raw.lower()):
            continue
        tokens.extend(_split_camel(raw))
    return tokens


def _is_noise(token: str) -> bool:
    if token in _NOISE_TOKENS:
        return True
    return any(p.match(token) for p in _NOISE_PATTERNS)


def _core_name(stem: str) -> str:
    """Reduce a filename stem to its identifying characters.

    Strips the numeric prefix, the DYDefinitions suffix, version/period marks,
    and all punctuation, leaving a lowercase alphanumeric run for comparison.
    """
    _, rest = _strip_prefix(stem)
    lowered = rest.lower()
    for suffix in _DEFS_SUFFIXES:
        idx = lowered.find(suffix)
        if idx != -1:
            rest = rest[:idx]
            break
    return "".join(_tokenize(rest))


def _similarity(a: str, b: str) -> float:
    """Score two core names in [0, 1]; containment counts as a strong match."""
    if not a or not b:
        return 0.0
    if a == b:
        return 1.0
    if a in b or b in a:
        shorter, longer = (a, b) if len(a) <= len(b) else (b, a)
        # A short name fully contained in a longer one (hialeah in
        # hialeahindustrialpark) is a confident match, scaled by how much of the
        # longer string it explains so that stray one-token hits stay low.
        return 0.75 + 0.25 * (len(shorter) / len(longer))
    return SequenceMatcher(None, a, b).ratio()


def loan_name_from_definitions(path: Path) -> str | None:
    """Read the authoritative loan name from the definitions file header.

    Each definitions file carries a `*****<Loan Name>*****` header after a
    template placeholder block. That name is preferred over anything parsed from
    a filename because it is what the loan agreement itself calls the property.
    """
    try:
        text = path.read_text(encoding="utf-8", errors="replace")
    except OSError:
        return None
    for line in text.splitlines():
        stripped = line.replace("\\", "").strip()
        if not stripped:
            continue
        m = _HEADER_RE.match(stripped)
        if not m:
            continue
        name = m.group(1).strip()
        if not name or name.lower() in _PLACEHOLDER_NAMES:
            continue
        if len(name) > 80:
            continue
        return name
    return None


def _prettify(core_stem: str) -> str:
    """Fallback display name when the definitions file has no usable header."""
    _, rest = _strip_prefix(core_stem)
    return " ".join(t.capitalize() for t in _tokenize(rest)) or core_stem


def find_candidates(input_dir: Path) -> tuple[list[Path], list[Path]]:
    """Return (excel files, definitions files) present in the input folder."""
    excels: list[Path] = []
    defs: list[Path] = []
    for p in sorted(input_dir.iterdir()):
        if not p.is_file() or p.name.startswith("~$") or p.name.startswith("."):
            continue
        suffix = p.suffix.lower()
        if suffix in (".xlsx", ".xlsm"):
            excels.append(p)
        elif suffix in (".md", ".txt"):
            # Only definition files pair with a model; the build spec and any
            # stray notes in the folder are not loan inputs.
            if "definition" in p.stem.lower():
                defs.append(p)
    return excels, defs


def discover(input_dir: Path, min_score: float = 0.45) -> tuple[list[LoanFiles], list[str]]:
    """Pair every Excel model with its definitions file.

    Returns the pairs plus a list of human-readable problems (unpaired files,
    ambiguous matches) for the run-level log.
    """
    excels, defs = find_candidates(input_dir)
    problems: list[str] = []

    if not excels:
        problems.append(f"No .xlsx models found in {input_dir}")
    if not defs:
        problems.append(f"No *Definitions* files found in {input_dir}")

    scored: list[tuple[float, Path, Path]] = []
    for x in excels:
        x_prefix, _ = _strip_prefix(x.stem)
        x_core = _core_name(x.stem)
        for d in defs:
            d_prefix, _ = _strip_prefix(d.stem)
            d_core = _core_name(d.stem)
            score = _similarity(x_core, d_core)
            # A shared numeric prefix is an explicit shared key, not an ordering
            # assumption - it lifts a weak name match but cannot create one.
            if x_prefix and d_prefix and x_prefix == d_prefix:
                score = min(1.0, score + 0.35)
            scored.append((score, x, d))

    scored.sort(key=lambda t: t[0], reverse=True)
    used_x: set[Path] = set()
    used_d: set[Path] = set()
    pairs: list[LoanFiles] = []

    for score, x, d in scored:
        if score < min_score or x in used_x or d in used_d:
            continue
        used_x.add(x)
        used_d.add(d)
        name = loan_name_from_definitions(d) or _prettify(x.stem)
        prefix, _ = _strip_prefix(x.stem)
        pairs.append(LoanFiles(loan_name=name, xlsx_path=x, defs_path=d, prefix=prefix))

    for x in excels:
        if x not in used_x:
            problems.append(f"Unpaired model (no definitions file matched): {x.name}")
    for d in defs:
        if d not in used_d:
            problems.append(f"Unpaired definitions file (no model matched): {d.name}")

    pairs.sort(key=lambda p: (p.prefix or "", p.loan_name))
    return pairs, problems
