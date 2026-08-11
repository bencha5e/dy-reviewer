"""Parse a loan's `*_DYDefinitions.md` into LoanParams.

The definitions files are extracts from the loan agreements, lightly escaped for
markdown (`\\*` for `*`). Every number the checks rely on is pulled from the text
rather than assumed, because the parameters differ per loan and change as loans
are added: the vacancy floor, the reserve rate, the management-fee percentage
and the delinquency window are all loan-specific.

Numbers are written inconsistently across agreements - "5.00%", "five percent
(5.0%)", "forty-five (45) days", "60-days" - so each pattern accepts the spelled
form alongside the numeral, and hyphenated number words must not break the match.
"""

from __future__ import annotations

import re
from pathlib import Path

from .params import CURRENT, LoanParams, Parsed

#: Spec section 2.3's parameter table, kept only to detect disagreement with the
#: loan text. The definitions file always wins; a mismatch is reported.
SPEC_TABLE = {
    "Strada": {"vacancy_floor": 0.05, "reserve_rate": 250.0, "delinquency": 60},
    "Campus at Villa La Jolla": {"vacancy_floor": 0.05, "reserve_rate": 0.25, "delinquency": 45},
    "Hialeah": {"vacancy_floor": 0.05, "reserve_rate": 0.25, "delinquency": 60},
    "Ares55thAve": {"vacancy_floor": 0.05, "reserve_rate": 0.10, "delinquency": 60},
}

#: Any character except a sentence-ending period. A period followed by a digit is
#: a decimal point and must be allowed through, or a clause containing "(5.0%)"
#: becomes unreachable.
_NOSENT = r"(?:[^.]|\.(?=\d))"

#: `(y) 5.00%` / `(y) five percent (5.0%)` / `(y) 5%` after "vacancy factor".
_VACANCY_FLOOR = re.compile(
    rf"vacancy factor{_NOSENT}{{0,200}}?\(y\)\s*(?:[\w\s-]*?\(\s*)?(\d+(?:\.\d+)?)\s*%",
    re.IGNORECASE,
)
_IG_CARVEOUT = re.compile(
    rf"vacancy factor{_NOSENT}{{0,300}}?\(\s*excluding\s+investment\s+grade", re.IGNORECASE
)

#: `$250 per unit`, `$0.25 per square foot`, `$0.10 per square feet`.
_RESERVE = re.compile(
    r"\$\s*([\d,]+(?:\.\d+)?)\s*per\s+(unit|square\s*(?:foot|feet|ft))", re.IGNORECASE
)

#: `3.00% of Gross Revenues`, `(3.0%) of Gross Revenue`, `3.0% of gross operating income`.
_MGMT_FEE = re.compile(
    r"(\d+(?:\.\d+)?)\s*%\s*\)?\s*of\s+(gross\s+\w+(?:\s+\w+)?)", re.IGNORECASE
)

#: `more than forty-five (45) days delinquent` - the number word may be hyphenated.
_DELINQ_NUMBERED = re.compile(
    r"more than\s+[\w\s()-]{0,40}?\((\d+)\)\s*days?\s+delinquent", re.IGNORECASE
)
#: `60-days past due`, `60 days past due`.
_DELINQ_PAST_DUE = re.compile(r"(\d+)\s*-?\s*days?\s+past\s+due", re.IGNORECASE)
#: Strada: no day count at all, only a requirement to be current.
_DELINQ_CURRENT = re.compile(r"current on their rental obligations", re.IGNORECASE)

_NEW_LEASE = re.compile(
    r"occupancy is expected to occur within\s*(\d+)\s*-?\s*days?", re.IGNORECASE
)
_NEW_LEASE_IG = re.compile(
    r"investment grade[^.]{0,80}?(\d+)\s*-?\s*days?|(\d+)\s*-?\s*days?\s+for\s+investment grade",
    re.IGNORECASE,
)
#: `over the following twelve (12) months` / `over the following ninety (90) days`.
_RENT_STEP = re.compile(
    r"over the following\s+[\w\s-]{0,25}?\((\d+)\)\s*(months?|days?)", re.IGNORECASE
)
#: Campus writes "within the next ninety (90) days"; Hialeah writes "within
#: either (1) three (3) months of their notified vacate date".
_KNOWN_VACATE = re.compile(
    rf"(?:terminate,?\s*cancel or vacate|intent to vacate){_NOSENT}{{0,90}}?"
    rf"within\s+(?:the next|either\s*\(\d+\))\s*"
    rf"[\w\s-]{{0,20}}?\((\d+)\)\s*(days?|months?)",
    re.IGNORECASE,
)


def _clean(text: str) -> str:
    """Undo the markdown escaping and collapse whitespace for matching."""
    return re.sub(r"\s+", " ", text.replace("\\", ""))


def _quote(match: re.Match, width: int = 130) -> str:
    """A readable snippet of the agreement around a match."""
    text = match.string
    start = max(0, match.start() - 30)
    end = min(len(text), match.end() + 30)
    snippet = text[start:end].strip()
    return (("..." if start else "") + snippet + ("..." if end < len(text) else ""))[:width]


def parse_definitions(path: Path, loan_name: str) -> LoanParams:
    """Read a definitions file into LoanParams, recording every source quote."""
    raw = Path(path).read_text(encoding="utf-8", errors="replace")
    text = _clean(raw)
    params = LoanParams(loan_name=loan_name)

    if m := _VACANCY_FLOOR.search(text):
        params.vacancy_floor = Parsed(float(m.group(1)) / 100.0, _quote(m))
    params.ig_vacancy_carveout = Parsed(bool(_IG_CARVEOUT.search(text)))

    if m := _RESERVE.search(text):
        basis = "unit" if m.group(2).lower().startswith("unit") else "sf"
        params.reserve_rate = Parsed(float(m.group(1).replace(",", "")), _quote(m))
        params.reserve_basis = Parsed(basis, _quote(m))

    if m := _MGMT_FEE.search(text):
        params.mgmt_fee_pct = Parsed(float(m.group(1)) / 100.0, _quote(m))
        params.mgmt_stated_base = Parsed(m.group(2).strip(), _quote(m))

    if m := _DELINQ_NUMBERED.search(text):
        params.delinquency = Parsed(int(m.group(1)), _quote(m))
    elif m := _DELINQ_PAST_DUE.search(text):
        params.delinquency = Parsed(int(m.group(1)), _quote(m))
    elif m := _DELINQ_CURRENT.search(text):
        params.delinquency = Parsed(CURRENT, _quote(m))

    if m := _NEW_LEASE.search(text):
        params.new_lease_days = Parsed(int(m.group(1)), _quote(m))
    if m := _NEW_LEASE_IG.search(text):
        value = m.group(1) or m.group(2)
        if value:
            params.new_lease_ig_days = Parsed(int(value), _quote(m))
    if m := _RENT_STEP.search(text):
        unit = "months" if m.group(2).lower().startswith("month") else "days"
        params.rent_step_window = Parsed(f"{m.group(1)} {unit}", _quote(m))
    if m := _KNOWN_VACATE.search(text):
        unit = "months" if m.group(2).lower().startswith("month") else "days"
        params.known_vacate_days = Parsed(f"{m.group(1)} {unit}", _quote(m))

    params.conflicts = _spec_conflicts(params)
    return params


def _spec_conflicts(params: LoanParams) -> list[str]:
    """Note where the build spec's table disagrees with the loan text."""
    expected = SPEC_TABLE.get(params.loan_name)
    if not expected:
        return []
    conflicts: list[str] = []

    if params.vacancy_floor.found and abs(params.vacancy_floor.value - expected["vacancy_floor"]) > 1e-9:
        conflicts.append(
            f"vacancy floor: loan agreement says {params.vacancy_floor.value:.2%}, "
            f"spec table says {expected['vacancy_floor']:.2%}"
        )
    if params.reserve_rate.found and abs(params.reserve_rate.value - expected["reserve_rate"]) > 1e-9:
        conflicts.append(
            f"reserve rate: loan agreement says {params.reserve_rate.value}, "
            f"spec table says {expected['reserve_rate']}"
        )
    if params.delinquency.found and params.delinquency.value != expected["delinquency"]:
        if params.delinquency.value == CURRENT:
            conflicts.append(
                f"delinquency window: the loan agreement states no day count - only that tenants "
                f"be \"current on their rental obligations\" - while the spec table says "
                f"{expected['delinquency']} days. The agreement is stricter and governs; tenants "
                f"with any past-due balance in the AR aging are excluded."
            )
        else:
            conflicts.append(
                f"delinquency window: loan agreement says {params.delinquency.value} days, "
                f"spec table says {expected['delinquency']} days"
            )
    return conflicts
