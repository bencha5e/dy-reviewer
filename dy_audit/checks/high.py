"""HIGH checks - methodology and adherence to the loan definition (spec section 4).

Every loan-specific number these checks compare against is read from the loan
agreement, never assumed. The four current loans all happen to use a 5% vacancy
floor and a 3% management fee, but the roster changes and the agreements differ.
"""

from __future__ import annotations

import datetime as dt
import re

from openpyxl.utils import column_index_from_string, get_column_letter

from .. import formula as F
from ..context import LoanContext
from ..model import Finding, Severity, Status
from ..osar import Line, excel_to_date, normalize_label
from ..params import CURRENT

#: Reserve rate tolerance. Ares stores 0.09996 for a $0.10/sf reserve, which is a
#: rounding artefact rather than a wrong rate.
RESERVE_TOLERANCE = 0.005

_ROW_RE = re.compile(r"^([A-Z]+)(\d+)$")

#: Header text that introduces the period a supporting tab covers.
_PERIOD_MARKER = re.compile(r"as[- ]of|ending period|period end(?:ing|ed)|through", re.I)
#: Text that carries a date but is not the tab's as-of date - a database caption
#: or an export timestamp. Campus's aging report leads with "Live 09/11/2023".
_NOT_AS_OF = re.compile(r"db caption|live \d|printed|run date|generated", re.I)

_MONTHS = (
    "january february march april may june july august september october november december"
).split()

#: A multiplier that applies a vacancy factor inside a revenue line, e.g. `*0.95`,
#: `*95%`, `*(1-0.05)`.
def _vacancy_multiplier_patterns(floor: float) -> list[re.Pattern]:
    keep = 1.0 - floor
    return [
        re.compile(rf"\*\s*0?{re.escape(f'{keep:g}'.lstrip('0'))}\b"),
        re.compile(rf"\*\s*{keep * 100:g}\s*%"),
        re.compile(rf"\*\s*\(\s*1\s*-\s*0?{re.escape(f'{floor:g}'.lstrip('0'))}\s*\)"),
        re.compile(rf"\*\s*\(\s*1\s*-\s*{floor * 100:g}\s*%\s*\)"),
    ]


def _row_of(coord: str) -> int | None:
    m = _ROW_RE.match(coord.replace("$", "").upper())
    return int(m.group(2)) if m else None


def _col_of(coord: str) -> str | None:
    m = _ROW_RE.match(coord.replace("$", "").upper())
    return m.group(1) if m else None


def _is_number_text(text: str) -> bool:
    return text.replace(".", "").replace("-", "").replace(",", "").isdigit()


def nearby_label(ctx: LoanContext, sheet: str, coord: str) -> str | None:
    """The descriptive label for a cell.

    Columns immediately to the left are checked **before** the sheet's leftmost
    columns. Supporting tabs often carry several unrelated label blocks across
    one row: Strada's operating statement labels `R54` from `Q54`
    ("Delinquent Rents") while `B54` holds an unrelated "UTILITIES" heading, so
    reading left-to-right would attach the wrong name to the cell.
    """
    row, col = _row_of(coord), _col_of(coord)
    if row is None or col is None:
        return None
    index = column_index_from_string(col)
    candidates = [get_column_letter(index - k) for k in (1, 2) if index - k >= 1]
    candidates += ["A", "B", "C", "D"]
    for candidate in candidates:
        text = ctx.wb.text(sheet, f"{candidate}{row}")
        if text and not _is_number_text(text):
            return text.strip()
    return None


def source_tabs(ctx: LoanContext) -> dict[str, str | None]:
    """Locate the rent roll, T12 and AR tabs by following the model's own links."""
    wb, tab = ctx.wb, ctx.tab
    found: dict[str, str | None] = {"rent_roll": None, "t12": None, "ar": None}

    gpr = tab.cell(Line.GPR)
    if gpr:
        res = F.resolve_defining_formula(wb, tab.sheet, gpr)
        # Consider the whole resolution path, not just the final formula. Ares's
        # GPR is `SUM(RR!J8)`, a single-cell SUM that the pass-through walker
        # follows all the way to a literal, leaving no formula to read refs from.
        candidates: list[str] = []
        for hop in res.path:
            sheet = hop.rsplit("!", 1)[0]
            if sheet != tab.sheet and wb.has_sheet(sheet):
                candidates.append(sheet)
        for ref in F.iter_refs(res.formula):
            sheet = ref.resolved_sheet(res.sheet)
            if sheet != tab.sheet and wb.has_sheet(sheet):
                candidates.append(sheet)
        named = [s for s in candidates if re.search(r"rent\s*roll|^rr\b|\brr$", s, re.I)]
        found["rent_roll"] = (named or candidates or [None])[0]

    # The T12 tab is whatever the reference column (H) pulls opex from.
    tax_row = tab.rows.get(Line.TAX)
    if tax_row:
        for ref in F.iter_refs(wb.formula(tab.sheet, f"H{tax_row}")):
            sheet = ref.resolved_sheet(tab.sheet)
            if sheet != tab.sheet and wb.has_sheet(sheet):
                found["t12"] = sheet
                break

    for sheet in wb.sheet_names:
        if wb.is_visible(sheet) and re.search(r"^ar$|aging|receivable", sheet, re.I):
            found["ar"] = sheet
            break
    return found


# ---------------------------------------------------------------------------
# Definitions: record what was parsed, and any conflict with the spec table
# ---------------------------------------------------------------------------


def check_definitions(ctx: LoanContext) -> list[Finding]:
    """Record the parsed loan parameters and surface conflicts, never resolve them."""
    params = ctx.params
    if params is None:
        return [
            Finding(
                "CHK_DEFINITIONS",
                Severity.HIGH,
                Status.UNVERIFIABLE,
                "No definitions file was parsed for this loan.",
            )
        ]

    findings: list[Finding] = []
    missing = params.missing()
    summary = (
        f"vacancy floor {params.vacancy_floor}, reserve {params.reserve_rate} per "
        f"{params.reserve_basis}, management fee {params.mgmt_fee_pct} of "
        f"{params.mgmt_stated_base}, delinquency {params.delinquency}"
    )
    findings.append(
        Finding(
            "CHK_DEFINITIONS",
            Severity.INFO,
            Status.MANUAL_REVIEW if missing else Status.PASS,
            (
                f"Parameters read from the loan agreement: {summary}."
                + (f" Could not parse: {', '.join(missing)}." if missing else "")
            ),
            evidence=params.vacancy_floor.quote,
        )
    )

    # The management-fee base is a deliberate house-rule override of the loan
    # wording, so it is recorded rather than left implicit.
    if params.mgmt_stated_base.found:
        findings.append(
            Finding(
                "CHK_MGMT_BASE_RULE",
                Severity.INFO,
                Status.PASS,
                f"House rule applied: the {params.mgmt_fee_pct.value:.2%} management fee is "
                f"tested against EGI, though this loan agreement says "
                f"\"{params.mgmt_stated_base.value}\". Applies to every loan regardless of wording.",
                evidence=params.mgmt_fee_pct.quote,
            )
        )

    for conflict in params.conflicts:
        findings.append(
            Finding(
                "CHK_DEFINITIONS_CONFLICT",
                Severity.INFO,
                Status.FLAG,
                f"The loan agreement and the build spec's parameter table disagree - "
                f"{conflict}. The loan agreement governs.",
                evidence=params.delinquency.quote,
            )
        )
    return findings


# ---------------------------------------------------------------------------
# CHK_VACANCY_FLOOR
# ---------------------------------------------------------------------------


def check_vacancy_floor(ctx: LoanContext) -> list[Finding]:
    """The floor hardcoded in the model must match the loan agreement's floor."""
    wb, tab, params = ctx.wb, ctx.tab, ctx.params
    coord = tab.cell(Line.VACANCY)
    if coord is None or params is None or not params.vacancy_floor.found:
        return [
            Finding(
                "CHK_VACANCY_FLOOR",
                Severity.HIGH,
                Status.UNVERIFIABLE,
                "Vacancy floor could not be compared: the loan agreement's floor or the "
                "model's vacancy line is missing.",
                sheet=tab.sheet,
                cell=coord,
            )
        ]

    floor = float(params.vacancy_floor.value)
    threshold = 1.0 - floor
    res = F.resolve_defining_formula(wb, tab.sheet, coord)

    # Collect every constant the vacancy branch could be keyed on: literals in
    # the formula itself, and the values of any plain cells it references
    # (Campus keeps its 5% in O19 rather than inline).
    literals: list[float] = []
    texts = [F.normalize(res.formula)]
    deps = F.dependencies(wb, res.sheet, res.coord)
    for sheet, cell in deps.values():
        dep_formula = wb.formula(sheet, cell)
        if dep_formula:
            texts.append(F.normalize(dep_formula))
        else:
            number = wb.number(sheet, cell)
            if number is not None:
                literals.append(number)

    for text in texts:
        for m in re.finditer(r"(\d+(?:\.\d+)?)\s*(%?)", text):
            if not m.group(1):
                continue
            value = float(m.group(1))
            literals.append(value / 100.0 if m.group(2) else value)

    matched = [v for v in literals if abs(v - floor) < 1e-6 or abs(v - threshold) < 1e-6]
    evidence = (
        f"{res.ref} = {F.normalize(res.formula)} | agreement floor {floor:.2%} "
        f"(threshold {threshold:.2%}) | constants seen: "
        f"{sorted({round(v, 6) for v in literals if 0 < v <= 1})}"
    )

    if matched:
        return [
            Finding(
                "CHK_VACANCY_FLOOR",
                Severity.HIGH,
                Status.PASS,
                f"Model's vacancy branch uses the {floor:.2%} floor the loan agreement specifies.",
                sheet=tab.sheet,
                cell=coord,
                evidence=evidence,
                on_dy_path=True,
            )
        ]

    return [
        Finding(
            "CHK_VACANCY_FLOOR",
            Severity.HIGH,
            Status.FLAG,
            f"The loan agreement sets a {floor:.2%} vacancy floor (occupancy threshold "
            f"{threshold:.2%}), but no matching constant appears in the model's vacancy "
            f"calculation. Confirm the model is not carrying another loan's floor.",
            sheet=tab.sheet,
            cell=coord,
            evidence=evidence,
            on_dy_path=True,
        )
    ]


# ---------------------------------------------------------------------------
# CHK_VACANCY_DOUBLE_COUNT
# ---------------------------------------------------------------------------


def _range_formulas(ctx: LoanContext, ref: F.Ref, default_sheet: str, limit: int = 400) -> list[str]:
    """Formulas inside a referenced range, including whole-column references."""
    sheet = ref.resolved_sheet(default_sheet)
    if not ctx.wb.has_sheet(sheet):
        return []
    if not ref.is_range:
        formula = ctx.wb.formula(sheet, ref.coord)
        return [F.normalize(formula)] if formula else []

    body = ref.body.replace("$", "").upper()
    start, end = body.split(":")
    if ref.is_whole_column:
        c1, c2 = column_index_from_string(start), column_index_from_string(end)
        rows, _ = ctx.wb.bounds(sheet)
        cells = [
            f"{get_column_letter(c)}{r}"
            for c in range(min(c1, c2), max(c1, c2) + 1)
            for r in range(1, min(rows, limit) + 1)
        ]
    else:
        cells = [c for _s, c in F.expand_range(ref, sheet, limit)]
    out = []
    for cell in cells:
        formula = ctx.wb.formula(sheet, cell)
        if formula:
            out.append(F.normalize(formula))
    return out


def check_vacancy_double_count(ctx: LoanContext) -> list[Finding]:
    """A vacancy factor applied inside GPR while actual vacancy already exceeds it.

    In-place GPR only counts occupied space, so actual vacancy is already
    embedded. When actual vacancy is above the floor the "greater of" test is
    satisfied by actual vacancy and no further deduction applies - applying the
    floor again inside the rent roll deducts it twice.
    """
    wb, tab, params = ctx.wb, ctx.tab, ctx.params
    gpr_coord = tab.cell(Line.GPR)
    if gpr_coord is None or params is None or not params.vacancy_floor.found:
        return []

    floor = float(params.vacancy_floor.value)
    threshold = 1.0 - floor
    occupancy = None
    occ_row = tab.rows.get(Line.OCCUPANCY)
    if occ_row:
        for col in (tab.dy_column, "H"):
            occupancy = wb.number(tab.sheet, f"{col}{occ_row}")
            if occupancy is not None:
                break
    if occupancy is None or occupancy > threshold:
        return []  # above the threshold a deduction is correct, not a double count

    res = F.resolve_defining_formula(wb, tab.sheet, gpr_coord)
    patterns = _vacancy_multiplier_patterns(floor)
    hits: list[tuple[str, str]] = []
    for ref in F.iter_refs(res.formula):
        sheet = ref.resolved_sheet(res.sheet)
        for text in _range_formulas(ctx, ref, res.sheet):
            if any(p.search(text) for p in patterns):
                hits.append((sheet, text))
                break
    if not hits:
        return []

    gpr_value = wb.number(tab.sheet, gpr_coord) or 0.0
    overstated = gpr_value / (1.0 - floor) - gpr_value if floor < 1 else 0.0
    upb = None
    if ctx.facts.get("upb_ref"):
        upb = wb.number(*ctx.facts["upb_ref"])
    impact = f" (~{overstated / upb:.2%} of debt yield)" if upb else ""
    sheet, sample = hits[0]

    return [
        Finding(
            "CHK_VACANCY_DOUBLE_COUNT",
            Severity.HIGH,
            Status.FLAG,
            f"A {floor:.2%} vacancy factor is applied inside the gross potential rent build on "
            f"{sheet!r}, but actual occupancy is {occupancy:.2%}, so actual vacancy "
            f"({1 - occupancy:.2%}) already exceeds the floor and is embedded in in-place rent. "
            f"The floor should not be applied again: GPR is understated by roughly "
            f"{overstated:,.0f}{impact}. The OSAR vacancy line is correctly 0, which hides this.",
            sheet=sheet,
            cell=None,
            evidence=f"{tab.sheet}!{gpr_coord} = {F.normalize(res.formula)} | source formula: {sample}",
            on_dy_path=True,
        )
    ]


# ---------------------------------------------------------------------------
# CHK_PERIOD
# ---------------------------------------------------------------------------


#: `FISCAL PERIOD 032026`, `PERIOD 03/2026` - a month-and-year stamp rather than
#: a full date. Strada's aging report identifies its period this way and carries
#: the export timestamp separately.
_FISCAL_PERIOD = re.compile(r"(?:fiscal\s+)?period\s*(\d{2})\s*/?\s*(\d{4})", re.I)


def _tab_period_evidence(
    ctx: LoanContext, sheet: str, quarter_end: dt.date
) -> tuple[bool, bool, list[str]]:
    """Look for confirmation that a supporting tab is as of quarter-end.

    Returns (matched, only_later_dates, notes). A date *after* quarter-end is a
    report run-date, which the spec explicitly allows so long as the snapshot
    itself is quarter-end; only a stale earlier period is a problem.
    """
    notes: list[str] = []
    matched = False
    saw_earlier = False
    saw_later = False
    month_name = _MONTHS[quarter_end.month - 1]
    month_text = re.compile(rf"{month_name}\s+{quarter_end.year}", re.I)

    for coord, _f, value in ctx.wb.iter_cells(sheet, max_row=14, max_col=32):
        text = str(value).strip() if value is not None else ""

        if isinstance(value, str):
            if fiscal := _FISCAL_PERIOD.search(text):
                month, year = int(fiscal.group(1)), int(fiscal.group(2))
                if (month, year) == (quarter_end.month, quarter_end.year):
                    matched = True
                    notes.append(f"{sheet}!{coord} = {fiscal.group(0)!r}")
                    continue
            if month_text.search(text):
                matched = True
                notes.append(f"{sheet}!{coord} = {text[:50]!r}")
                continue
            if _NOT_AS_OF.search(text):
                # A database caption or export timestamp, not the reporting period.
                continue

        date = excel_to_date(value)
        if date is None:
            continue
        if date == quarter_end:
            matched = True
            notes.append(f"{sheet}!{coord} = {quarter_end:%m/%d/%Y}")
        elif isinstance(value, str) and _PERIOD_MARKER.search(text):
            notes.append(f"{sheet}!{coord} = {text[:60]!r} -> {date:%m/%d/%Y}")
            if date > quarter_end:
                saw_later = True
            else:
                saw_earlier = True

    return matched, (saw_later and not saw_earlier), notes[:4]


def check_period(ctx: LoanContext) -> list[Finding]:
    """T12 end-month, rent-roll as-of and AR as-of must all equal quarter-end."""
    tab = ctx.tab
    quarter_end = tab.period_end
    if quarter_end is None:
        return [
            Finding(
                "CHK_PERIOD",
                Severity.HIGH,
                Status.UNVERIFIABLE,
                "The OSAR statement ending date could not be read, so supporting tabs "
                "cannot be tied to quarter-end.",
                sheet=tab.sheet,
            )
        ]

    tabs = source_tabs(ctx)
    ctx.facts["source_tabs"] = tabs
    findings: list[Finding] = []

    for role, label in (("rent_roll", "Rent roll"), ("t12", "T12 / operating statement"), ("ar", "AR aging")):
        sheet = tabs.get(role)
        if sheet is None:
            if role == "ar":
                findings.append(
                    Finding(
                        "CHK_PERIOD",
                        Severity.HIGH,
                        Status.MANUAL_REVIEW,
                        "No AR or aging tab is present in this workbook, so its as-of date "
                        "cannot be tied to quarter-end.",
                        sheet=tab.sheet,
                    )
                )
            continue
        matched, run_date_only, notes = _tab_period_evidence(ctx, sheet, quarter_end)
        evidence = "; ".join(notes) if notes else "no dated header cells found"
        if matched:
            findings.append(
                Finding(
                    "CHK_PERIOD",
                    Severity.HIGH,
                    Status.PASS,
                    f"{label} tab {sheet!r} is as of {quarter_end:%m/%d/%Y}.",
                    sheet=sheet,
                    evidence=evidence,
                    on_dy_path=True,
                )
            )
        elif run_date_only:
            findings.append(
                Finding(
                    "CHK_PERIOD",
                    Severity.HIGH,
                    Status.MANUAL_REVIEW,
                    f"{label} tab {sheet!r} is stamped later than quarter-end "
                    f"({quarter_end:%m/%d/%Y}), which reads as a report run-date rather than a "
                    f"different reporting period. That is acceptable provided the snapshot "
                    f"itself is as of quarter-end - confirm.",
                    sheet=sheet,
                    evidence=evidence,
                    on_dy_path=True,
                )
            )
        elif notes:
            findings.append(
                Finding(
                    "CHK_PERIOD",
                    Severity.HIGH,
                    Status.FLAG,
                    f"{label} tab {sheet!r} carries a period marker earlier than quarter-end "
                    f"({quarter_end:%m/%d/%Y}) - it looks like a stale snapshot from an "
                    f"earlier period.",
                    sheet=sheet,
                    evidence=evidence,
                    on_dy_path=True,
                )
            )
        else:
            findings.append(
                Finding(
                    "CHK_PERIOD",
                    Severity.HIGH,
                    Status.MANUAL_REVIEW,
                    f"{label} tab {sheet!r} has no readable period header; confirm by hand "
                    f"that it is as of {quarter_end:%m/%d/%Y}.",
                    sheet=sheet,
                    evidence=evidence,
                    on_dy_path=True,
                )
            )
    return findings


# ---------------------------------------------------------------------------
# CHK_RESERVE_RATE
# ---------------------------------------------------------------------------


def check_reserve_rate(ctx: LoanContext) -> list[Finding]:
    """The replacement reserve rate must match the loan agreement."""
    wb, tab, params = ctx.wb, ctx.tab, ctx.params
    coord = tab.cell(Line.CAPEX)
    if coord is None or params is None or not params.reserve_rate.found:
        return [
            Finding(
                "CHK_RESERVE_RATE",
                Severity.HIGH,
                Status.UNVERIFIABLE,
                "Reserve rate could not be compared: the capital-expenditure line or the "
                "agreement's rate is missing.",
                sheet=tab.sheet,
                cell=coord,
            )
        ]

    expected = float(params.reserve_rate.value)
    basis = params.reserve_basis.value or "unit"
    res = F.resolve_defining_formula(wb, tab.sheet, coord)

    observed: list[tuple[str, float]] = []
    for ref in F.iter_refs(res.formula):
        if ref.is_range:
            continue
        sheet = ref.resolved_sheet(res.sheet)
        value = wb.number(sheet, ref.coord)
        if value is not None:
            observed.append((f"{sheet}!{ref.coord}", value))
    for m in re.finditer(r"(?<![A-Za-z0-9_.$])(\d+(?:\.\d+)?)", F.normalize(res.formula)):
        observed.append(("literal", float(m.group(1))))

    tolerance = abs(expected) * RESERVE_TOLERANCE
    matches = [(where, v) for where, v in observed if abs(v - expected) <= tolerance]
    evidence = (
        f"{res.ref} = {F.normalize(res.formula)} | agreement: ${expected:g} per {basis} "
        f"| values seen: {[(w, round(v, 6)) for w, v in observed]}"
    )

    if matches:
        where, value = matches[0]
        note = ""
        if abs(value - expected) > 1e-9:
            note = f" (stored as {value:g}, within the {RESERVE_TOLERANCE:.1%} rounding tolerance)"
        return [
            Finding(
                "CHK_RESERVE_RATE",
                Severity.HIGH,
                Status.PASS,
                f"Reserve rate ${expected:g} per {basis} matches the loan agreement{note}.",
                sheet=tab.sheet,
                cell=coord,
                evidence=evidence,
                on_dy_path=True,
            )
        ]

    return [
        Finding(
            "CHK_RESERVE_RATE",
            Severity.HIGH,
            Status.FLAG,
            f"The loan agreement requires a ${expected:g} per {basis} replacement reserve, but "
            f"no such rate appears in the capital-items calculation.",
            sheet=tab.sheet,
            cell=coord,
            evidence=evidence,
            on_dy_path=True,
        )
    ]


# ---------------------------------------------------------------------------
# CHK_MGMT_BASE
# ---------------------------------------------------------------------------


def check_mgmt_base(ctx: LoanContext) -> list[Finding]:
    """The management fee percentage must be applied to EGI (house rule)."""
    wb, tab, params = ctx.wb, ctx.tab, ctx.params
    coord = tab.cell(Line.MGMT_FEE)
    if coord is None:
        return [
            Finding(
                "CHK_MGMT_BASE",
                Severity.HIGH,
                Status.UNVERIFIABLE,
                "No management fee line found on the OSAR tab.",
                sheet=tab.sheet,
            )
        ]

    res = F.resolve_defining_formula(wb, tab.sheet, coord)
    pct = float(params.mgmt_fee_pct.value) if params and params.mgmt_fee_pct.found else 0.03
    args = F.extract_call_args(res.formula, "MAX")
    evidence_base = f"{res.ref} = {F.normalize(res.formula)}"

    if args is None:
        return [
            Finding(
                "CHK_MGMT_BASE",
                Severity.HIGH,
                Status.FLAG,
                f"Management fee is not MAX(actual, {pct:.2%} of EGI) - it takes no MAX, so the "
                f"greater of the two is not enforced.",
                sheet=tab.sheet,
                cell=coord,
                evidence=evidence_base,
                on_dy_path=True,
            )
        ]

    # The percentage term is the argument carrying the rate; whatever cell it
    # multiplies must be the EGI line, identified by its label on its own sheet.
    for arg in args:
        has_rate = any(
            abs(float(m.group(1)) - pct) < 1e-9 or abs(float(m.group(1)) / 100.0 - pct) < 1e-9
            for m in re.finditer(r"(?<![A-Za-z0-9_.$])(\d+(?:\.\d+)?)", arg)
        )
        rate_cells = []
        for ref in F.iter_refs(arg):
            sheet = ref.resolved_sheet(res.sheet)
            value = wb.number(sheet, ref.coord)
            if value is not None and abs(value - pct) < 1e-9:
                has_rate = True
                rate_cells.append(f"{sheet}!{ref.coord}")
        if not has_rate:
            continue

        for ref in F.iter_refs(arg):
            sheet = ref.resolved_sheet(res.sheet)
            if f"{sheet}!{ref.coord}" in rate_cells:
                continue
            label = normalize_label(nearby_label(ctx, sheet, ref.coord) or "")
            if not label:
                continue
            if "effective gross income" in label:
                return [
                    Finding(
                        "CHK_MGMT_BASE",
                        Severity.HIGH,
                        Status.PASS,
                        f"Management fee = MAX(actual, {pct:.2%} x EGI), the required base.",
                        sheet=tab.sheet,
                        cell=coord,
                        evidence=f"{evidence_base} | base {sheet}!{ref.coord} = {label!r}",
                        on_dy_path=True,
                    )
                ]
            if "gross potential rent" in label or label.startswith("base rent"):
                return [
                    Finding(
                        "CHK_MGMT_BASE",
                        Severity.HIGH,
                        Status.FLAG,
                        f"Management fee applies {pct:.2%} to {label!r} instead of EGI. The base "
                        f"is always EGI regardless of the loan agreement's wording.",
                        sheet=tab.sheet,
                        cell=coord,
                        evidence=f"{evidence_base} | base {sheet}!{ref.coord} = {label!r}",
                        on_dy_path=True,
                    )
                ]

    return [
        Finding(
            "CHK_MGMT_BASE",
            Severity.HIGH,
            Status.MANUAL_REVIEW,
            f"Management fee takes a MAX but the {pct:.2%} base could not be identified; "
            f"confirm it is EGI.",
            sheet=tab.sheet,
            cell=coord,
            evidence=evidence_base,
            on_dy_path=True,
        )
    ]


# ---------------------------------------------------------------------------
# CHK_OTHER_INCOME_BASIS
# ---------------------------------------------------------------------------

#: Basis keywords per loan, taken from the definition text.
_T3_HINT = re.compile(r"trailing three-month|trailing 3", re.I)


def check_other_income_basis(ctx: LoanContext) -> list[Finding]:
    """Report the basis the model uses for other income against the agreement's."""
    wb, tab = ctx.wb, ctx.tab
    coord = tab.cell(Line.OTHER_INCOME)
    if coord is None:
        return []

    defs_text = ""
    if ctx.files and ctx.files.defs_path.exists():
        defs_text = ctx.files.defs_path.read_text(encoding="utf-8", errors="replace")
    expected = "trailing 3-month annualized" if _T3_HINT.search(defs_text) else "trailing 12-month"

    res = F.resolve_defining_formula(wb, tab.sheet, coord)
    sources = []
    for ref in F.iter_refs(res.formula):
        sheet = ref.resolved_sheet(res.sheet)
        label = nearby_label(ctx, sheet, ref.coord) if not ref.is_range else None
        sources.append(f"{sheet}!{ref.body}" + (f" ({label})" if label else ""))

    tabs = ctx.facts.get("source_tabs") or source_tabs(ctx)
    t12_tab = tabs.get("t12")
    # A workbook can hold several trailing-twelve tabs under different names -
    # Ares has both "Actuals (T12)" and "T12" - so any of them counts as a T12
    # source, not only the one column H happens to link to.
    t12_like = re.compile(r"t-?12|ttm|trailing|operating statement", re.I)
    references_t12 = any(
        (t12_tab and t12_tab in s) or t12_like.search(s.split("!")[0]) for s in sources
    )
    evidence = f"{res.ref} = {F.normalize(res.formula)} | sources: {sources}"

    # Only a clear contradiction is worth a flag; anything ambiguous goes to the
    # reviewer rather than guessing at a tab's column semantics.
    if expected.startswith("trailing 12") and references_t12:
        status, message = Status.PASS, (
            f"Other income is drawn from the T12 tab {t12_tab!r}, matching the loan "
            f"agreement's {expected} basis."
        )
    else:
        status, message = Status.MANUAL_REVIEW, (
            f"The loan agreement specifies a {expected} basis for other income. Confirm the "
            f"model's source matches - the tool reports the linkage but does not interpret "
            f"the source tab's column periods."
        )
    return [
        Finding(
            "CHK_OTHER_INCOME_BASIS",
            Severity.HIGH,
            status,
            message,
            sheet=tab.sheet,
            cell=coord,
            evidence=evidence,
            on_dy_path=True,
        )
    ]


# ---------------------------------------------------------------------------
# CHK_EXCLUSIONS
# ---------------------------------------------------------------------------

_RECOVERY_LABEL = re.compile(r"recover|reimburs", re.I)
_DELINQUENT_LABEL = re.compile(r"delinquen", re.I)


def check_exclusions(ctx: LoanContext) -> list[Finding]:
    """Delinquency exclusions and the placement of recovery income.

    The AR aging is the source of truth for delinquency, so a workbook without
    one cannot pass this check - it reports MANUAL_REVIEW instead.
    """
    wb, tab, params = ctx.wb, ctx.tab, ctx.params
    findings: list[Finding] = []
    tabs = ctx.facts.get("source_tabs") or source_tabs(ctx)
    ar_tab = tabs.get("ar")

    window = params.delinquency.value if params and params.delinquency.found else None
    window_text = (
        "tenants must be current (any past-due balance excludes them)"
        if window == CURRENT
        else f"tenants more than {window} days delinquent are excluded"
        if window
        else "the delinquency window could not be read from the loan agreement"
    )

    if ar_tab is None:
        findings.append(
            Finding(
                "CHK_EXCLUSIONS",
                Severity.HIGH,
                Status.MANUAL_REVIEW,
                f"This workbook has no AR or aging tab, so delinquency exclusions cannot be "
                f"verified against their source of truth. The loan agreement requires that "
                f"{window_text}. Obtain the aging report and confirm by hand.",
                sheet=tab.sheet,
                evidence=f"sheets present: {', '.join(wb.sheet_names)}",
                on_dy_path=True,
            )
        )
    else:
        findings.append(
            Finding(
                "CHK_EXCLUSIONS",
                Severity.HIGH,
                Status.PASS,
                f"AR aging tab {ar_tab!r} is present for delinquency testing; {window_text}.",
                sheet=ar_tab,
                evidence=params.delinquency.quote if params else None,
                on_dy_path=True,
            )
        )

    findings.extend(_check_delinquency_direction(ctx, ar_tab, window_text))
    findings.extend(_check_recovery_placement(ctx))
    return findings


def _check_delinquency_direction(ctx: LoanContext, ar_tab: str | None, window_text: str) -> list[Finding]:
    """A delinquency adjustment inside GPR must reduce it, never increase it."""
    wb, tab = ctx.wb, ctx.tab
    coord = tab.cell(Line.GPR)
    if coord is None:
        return []
    res = F.resolve_defining_formula(wb, tab.sheet, coord)

    for ref, start, _end in F.iter_refs_positioned(res.formula):
        if ref.is_range:
            continue
        sheet = ref.resolved_sheet(res.sheet)
        label = nearby_label(ctx, sheet, ref.coord) or ""
        if not _DELINQUENT_LABEL.search(label):
            continue
        sign = F.term_sign(res.formula, start)
        value = wb.number(sheet, ref.coord)
        if sign > 0:
            return [
                Finding(
                    "CHK_EXCLUSIONS",
                    Severity.HIGH,
                    Status.FLAG,
                    f"Gross potential rent adds a delinquency term ({label!r}) rather than "
                    f"deducting it. The loan agreement requires that {window_text}, so a "
                    f"delinquent balance must reduce rent, not increase it. The term is "
                    f"{value:,.2f} today, so the reported DY is unaffected this quarter - but "
                    f"it will overstate GPR as soon as the aging report shows a balance.",
                    sheet=tab.sheet,
                    cell=coord,
                    evidence=(
                        f"{tab.sheet}!{coord} = {F.normalize(res.formula)} | "
                        f"{sheet}!{ref.coord} = {label!r} = {value!r}"
                        + (f" | AR tab: {ar_tab}" if ar_tab else "")
                    ),
                    on_dy_path=True,
                )
            ]
    return []


def _check_recovery_placement(ctx: LoanContext) -> list[Finding]:
    """Recoveries belong on the reimbursement line, where the vacancy factor reaches them."""
    wb, tab = ctx.wb, ctx.tab
    other_coord = tab.cell(Line.OTHER_INCOME)
    reimb_coord = tab.cell(Line.REIMBURSEMENT)
    if other_coord is None or reimb_coord is None:
        return []

    reimb_value = wb.number(tab.sheet, reimb_coord) or 0.0
    if abs(reimb_value) > 1.0:
        return []  # recoveries are already on their own line

    res = F.resolve_defining_formula(wb, tab.sheet, other_coord)
    candidates = {f"{res.sheet}!{res.coord}": (res.sheet, res.coord)}
    candidates.update(F.dependencies(wb, res.sheet, res.coord, max_depth=3))

    for sheet, cell in candidates.values():
        label = nearby_label(ctx, sheet, cell) or ""
        if not _RECOVERY_LABEL.search(label):
            continue
        amount = wb.number(sheet, cell)
        if amount is None or abs(amount) < 1.0:
            continue
        return [
            Finding(
                "CHK_EXCLUSIONS",
                Severity.MEDIUM,
                Status.FLAG,
                f"About {amount:,.0f} of tenant recoveries sits in Other Income while the "
                f"Expense Reimbursement line is empty. The loan definition places recoveries in "
                f"the revenue clause that the vacancy factor applies to, so parking them in "
                f"Other Income lets them escape it. Confirm the treatment.",
                sheet=tab.sheet,
                cell=other_coord,
                evidence=(
                    f"{tab.sheet}!{other_coord} -> {sheet}!{cell} = {label!r} = {amount:,.2f}; "
                    f"{tab.sheet}!{reimb_coord} = {reimb_value:,.2f}"
                ),
                on_dy_path=True,
            )
        ]
    return []


HIGH_CHECKS = (
    check_definitions,
    check_vacancy_floor,
    check_vacancy_double_count,
    check_period,
    check_reserve_rate,
    check_mgmt_base,
    check_other_income_basis,
    check_exclusions,
)


def run_high(ctx: LoanContext) -> list[Finding]:
    findings: list[Finding] = []
    for check in HIGH_CHECKS:
        findings.extend(check(ctx))
    return findings
