"""Independent recompute of the two lines the spec requires rebuilt: GPR and vacancy.

Everything else in the tool verifies the model's own formulas. These two are
rebuilt from the rent roll and compared back, because they are where the
definitional rules actually bite.

The governing rule is **never guess**. A rent roll the parser cannot read
confidently produces MANUAL_REVIEW, never a variance flag: a half-parsed rent
roll reporting a spurious GPR difference would be worse than no check at all.
Rent steps and free rent are deliberately left to the reviewer (spec section 5,
rule 12), so a model carrying those adjustments is reported with both figures
rather than flagged.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

from openpyxl.utils import column_index_from_string, get_column_letter

from . import formula as F
from .context import LoanContext
from .model import Finding, Severity, Status
from .osar import Line

#: Individual revenue lines flag above 0.1% difference (spec section 4).
REVENUE_TOLERANCE = 0.001

#: A status that means the tenant is in occupancy and paying.
_OCCUPIED = re.compile(r"^occupied", re.I)
#: Columns whose header marks an analyst include/exclude decision.
_INCLUDE_HEADER = re.compile(r"include.*(dy|test)|(dy|test).*include", re.I)
#: Headers signalling that the model applies step-up or free-rent adjustments,
#: which the definition permits and the tool does not attempt to verify.
_ADJUSTMENT_HEADER = re.compile(
    r"free rent|step[- ]?up|months? included|months? excluded|adj\w*\s+annual", re.I
)
_MONTHLY_HEADER = re.compile(r"month", re.I)
_ANNUAL_HEADER = re.compile(r"annual|annualiz", re.I)

_CELL_RE = re.compile(r"^([A-Z]+)(\d+)$")


def _split(coord: str) -> tuple[str, int]:
    m = _CELL_RE.match(coord.replace("$", "").upper())
    return (m.group(1), int(m.group(2))) if m else ("A", 0)


@dataclass
class RentRollParse:
    """What the parser could establish about a rent roll."""

    sheet: str | None = None
    rent_column: str | None = None
    rows: list[int] = field(default_factory=list)
    values: dict[int, float] = field(default_factory=dict)
    status_column: str | None = None
    statuses: dict[int, str] = field(default_factory=dict)
    monthly: bool = True
    has_adjustments: bool = False
    adjustment_headers: list[str] = field(default_factory=list)
    #: True when the summed column is itself computed by the model (an "adjusted
    #: rent" column) rather than raw contractual rent. The recompute then
    #: re-sums the model's own working rather than rebuilding from scratch.
    rent_column_is_derived: bool = False
    notes: list[str] = field(default_factory=list)

    @property
    def confident(self) -> bool:
        return bool(self.sheet and self.rent_column and self.values)

    def included_rows(self) -> list[int]:
        """Rows counted as in-place, current tenants."""
        if not self.statuses:
            return sorted(self.values)
        return sorted(r for r in self.values if _OCCUPIED.match(self.statuses.get(r, "")))

    def annualized(self) -> float:
        total = sum(self.values[r] for r in self.included_rows())
        return total * 12 if self.monthly else total


def _header_text(ctx: LoanContext, sheet: str, column: str, first_row: int) -> str:
    """Join the header cells sitting above a column's data."""
    parts = []
    for row in range(max(1, first_row - 6), first_row):
        text = ctx.wb.text(sheet, f"{column}{row}")
        if text:
            parts.append(text)
    return " ".join(parts)


#: A referenced range no bigger than this whose cells are themselves formulas is
#: treated as a summary block and its member formulas are followed. Tenant data
#: ranges run to hundreds of rows and are never expanded.
_SUMMARY_BLOCK_LIMIT = 40


def _candidate_ranges(ctx: LoanContext, sheet: str, start_sheet: str, formula: str | None,
                      depth: int = 0) -> list[F.Ref]:
    """Ranges on the rent-roll sheet reachable from a formula.

    Strada's GPR reaches its tenant rows through a summary block of `SUMIF`s
    referenced cell by cell (`M540-M532-M537-M536`), so the search follows
    single-cell references on the rent roll one more hop. Lydian wraps the same
    block in a range (`SUM(G251:G257)` of per-status `SUMIF`s), which would stop
    the walk at the block itself - re-summing the model's own subtotals and
    proving nothing - so members of a small formula-bearing range are followed
    too, down to the tenant rows the block aggregates.
    """
    if not formula or depth > 3:
        return []
    found: list[F.Ref] = []
    for ref in F.iter_refs(formula):
        ref_sheet = ref.resolved_sheet(start_sheet)
        if ref_sheet != sheet:
            continue
        if ref.is_range:
            found.append(ref)
            for member_sheet, member in F.expand_range(ref, sheet, _SUMMARY_BLOCK_LIMIT):
                member_formula = ctx.wb.formula(member_sheet, member)
                if member_formula:
                    found.extend(
                        _candidate_ranges(ctx, sheet, sheet, member_formula, depth + 1)
                    )
        else:
            found.extend(
                _candidate_ranges(ctx, sheet, sheet, ctx.wb.formula(sheet, ref.coord), depth + 1)
            )
    return found


def parse_rent_roll(ctx: LoanContext) -> RentRollParse:
    """Locate the tenant rows and the in-place rent column feeding GPR."""
    wb, tab = ctx.wb, ctx.tab
    parse = RentRollParse()
    tabs = ctx.facts.get("source_tabs") or {}
    sheet = tabs.get("rent_roll")
    if not sheet or not wb.has_sheet(sheet):
        parse.notes.append("could not identify the rent-roll tab from the GPR formula")
        return parse
    parse.sheet = sheet

    gpr_coord = tab.cell(Line.GPR)
    if gpr_coord is None:
        parse.notes.append("no GPR line on the OSAR tab")
        return parse
    res = F.resolve_defining_formula(wb, tab.sheet, gpr_coord)

    ranges = _candidate_ranges(ctx, sheet, res.sheet, res.formula)
    # The same tenant range is reachable through several summary cells; keep one.
    unique: dict[str, F.Ref] = {}
    for ref in ranges:
        unique.setdefault(ref.body.replace("$", "").upper(), ref)
    ranges = list(unique.values())
    # A pass-through that lands directly on a rent-roll cell (Ares: SUM(RR!J8))
    # leaves no range; treat the landing cell's column as the rent column.
    if not ranges and res.sheet == sheet:
        col, row = _split(res.coord)
        parse.rent_column, parse.rows = col, [row]
        value = wb.number(sheet, res.coord)
        if value is not None:
            parse.values[row] = value

    numeric_ranges: list[tuple[F.Ref, dict[int, float]]] = []
    text_ranges: list[tuple[F.Ref, dict[int, str]]] = []
    rows_bound, _cols = wb.bounds(sheet)

    for ref in ranges:
        body = ref.body.replace("$", "").upper()
        start, end = body.split(":")
        if ref.is_whole_column:
            col = re.sub(r"\d", "", start)
            row_span = range(1, min(rows_bound, 2000) + 1)
        else:
            c1, r1 = _split(start)
            c2, r2 = _split(end)
            if c1 != c2:
                continue
            col, row_span = c1, range(min(r1, r2), max(r1, r2) + 1)
        numbers: dict[int, float] = {}
        texts: dict[int, str] = {}
        for row in row_span:
            value = wb.value(sheet, f"{col}{row}")
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                numbers[row] = float(value)
            elif isinstance(value, str) and value.strip():
                texts[row] = value.strip()
        if numbers:
            numeric_ranges.append((ref, numbers))
        if texts and len(texts) > len(numbers):
            text_ranges.append((ref, texts))

    # Resolve the status column before the rent column: on a unit-level rent roll
    # the two must describe the same rows. Strada reaches its tenant rents
    # through a summary block of SUMIFs, so both the tenant range (T7:T517) and
    # the summary range (M531:M539) are reachable and sum to the same total -
    # only the tenant range lines up with the statuses in H7:H517.
    for ref, texts in text_ranges:
        if any(_OCCUPIED.match(s) for s in list(texts.values())[:60]):
            parse.status_column = re.sub(
                r"\d", "", ref.body.replace("$", "").upper().split(":")[0]
            )
            parse.statuses = texts
            break

    if numeric_ranges and not parse.rent_column:
        if parse.statuses:
            status_rows = set(parse.statuses)
            aligned = [
                (ref, numbers) for ref, numbers in numeric_ranges if status_rows & set(numbers)
            ]
            numeric_ranges = aligned or numeric_ranges

        # Pick the numeric range that best explains the model's GPR, testing both
        # a monthly and an already-annual reading. Scoring uses the filtered
        # total so the definitional inclusion rule is part of the match.
        model_gpr = wb.number(tab.sheet, gpr_coord) or 0.0
        best = None
        for ref, numbers in numeric_ranges:
            if parse.statuses:
                rows = [r for r in numbers if _OCCUPIED.match(parse.statuses.get(r, ""))]
                total = sum(numbers[r] for r in rows) if rows else sum(numbers.values())
            else:
                total = sum(numbers.values())
            for monthly, scaled in ((True, total * 12), (False, total)):
                if scaled == 0:
                    continue
                error = abs(scaled - model_gpr) / max(abs(model_gpr), 1.0)
                if best is None or error < best[0]:
                    best = (error, ref, numbers, monthly)
        if best:
            _err, ref, numbers, monthly = best
            parse.rent_column = re.sub(r"\d", "", ref.body.replace("$", "").upper().split(":")[0])
            parse.values = numbers
            parse.rows = sorted(numbers)
            parse.monthly = monthly
            if parse.statuses and not (set(parse.statuses) & set(numbers)):
                # The chosen rents do not line up with the statuses, so the
                # inclusion filter would silently drop every row.
                parse.status_column, parse.statuses = None, {}
                parse.notes.append("status column ignored: rows do not align with the rent range")

    if parse.rows:
        first_row = min(parse.rows)
        derived = sum(1 for r in parse.rows if wb.formula(sheet, f"{parse.rent_column}{r}"))
        parse.rent_column_is_derived = derived > len(parse.rows) / 2
        header = _header_text(ctx, sheet, parse.rent_column, first_row)
        if _ANNUAL_HEADER.search(header) and not _MONTHLY_HEADER.search(header):
            parse.monthly = False
        elif _MONTHLY_HEADER.search(header):
            parse.monthly = True
        parse.notes.append(f"rent column {parse.rent_column} header: {header[:70]!r}")

        # Note any adjustment columns; their presence means the model's GPR is
        # not a plain annualisation and the difference is for a human to judge.
        index = column_index_from_string(parse.rent_column)
        for offset in range(-12, 13):
            col_index = index + offset
            if col_index < 1:
                continue
            column = get_column_letter(col_index)
            head = _header_text(ctx, sheet, column, first_row)
            if _ADJUSTMENT_HEADER.search(head):
                parse.has_adjustments = True
                parse.adjustment_headers.append(f"{column}: {head[:44]}")
            elif _INCLUDE_HEADER.search(head):
                parse.has_adjustments = True
                parse.adjustment_headers.append(f"{column}: {head[:44]} (analyst include flag)")
    else:
        parse.notes.append("no numeric tenant rent range reachable from the GPR formula")

    return parse


def check_gpr_recompute(ctx: LoanContext) -> list[Finding]:
    """Rebuild annualized in-place rent from the rent roll and compare."""
    wb, tab = ctx.wb, ctx.tab
    coord = tab.cell(Line.GPR)
    if coord is None:
        return []
    model_gpr = wb.number(tab.sheet, coord)
    parse = parse_rent_roll(ctx)
    ctx.facts["rent_roll_parse"] = parse

    if not parse.confident or model_gpr is None:
        return [
            Finding(
                "CHK_GPR_RECOMPUTE",
                Severity.HIGH,
                Status.MANUAL_REVIEW,
                "Gross potential rent could not be independently rebuilt from the rent roll, so "
                "no variance is reported. Rebuild by hand: "
                + ("; ".join(parse.notes) if parse.notes else "rent roll not readable."),
                sheet=tab.sheet,
                cell=coord,
                evidence=f"rent-roll tab: {parse.sheet!r}",
                on_dy_path=True,
            )
        ]

    recomputed = parse.annualized()
    included = len(parse.included_rows())
    basis = "monthly x 12" if parse.monthly else "annual"
    detail = (
        f"{parse.sheet}!{parse.rent_column} over {included} included tenant row(s), {basis}; "
        f"recomputed {recomputed:,.2f} vs model {model_gpr:,.2f}"
        + (f"; status column {parse.status_column}" if parse.status_column else "")
    )
    difference = recomputed - model_gpr
    relative = abs(difference) / max(abs(model_gpr), 1.0)

    if relative <= REVENUE_TOLERANCE:
        # Be explicit about how strong the tie actually is. Where the summed
        # column is computed by the model itself, this confirms the rent roll
        # foots to the OSAR but does not re-derive the per-tenant adjustments
        # inside that column.
        caveat = ""
        if parse.rent_column_is_derived:
            caveat = (
                f" Note: column {parse.rent_column} is calculated by the model rather than raw "
                f"contractual rent, so this confirms the rent roll foots to the OSAR but does "
                f"not independently verify the per-tenant adjustments inside it"
                + (f" ({', '.join(parse.adjustment_headers[:2])})" if parse.adjustment_headers else "")
                + "."
            )
        return [
            Finding(
                "CHK_GPR_RECOMPUTE",
                Severity.HIGH,
                Status.PASS,
                f"Independently rebuilt gross potential rent ties to the model within "
                f"{REVENUE_TOLERANCE:.1%} ({recomputed:,.0f} vs {model_gpr:,.0f})." + caveat,
                sheet=tab.sheet,
                cell=coord,
                evidence=detail,
                on_dy_path=True,
            )
        ]

    if parse.has_adjustments:
        return [
            Finding(
                "CHK_GPR_RECOMPUTE",
                Severity.HIGH,
                Status.MANUAL_REVIEW,
                f"A plain annualisation of in-place rent gives {recomputed:,.0f} against the "
                f"model's {model_gpr:,.0f} ({difference:+,.0f}, {relative:.2%}). The rent roll "
                f"carries step-up, free-rent or include-flag adjustments, which the loan "
                f"definition permits and this tool does not verify - review the difference by "
                f"hand rather than treating it as an error.",
                sheet=tab.sheet,
                cell=coord,
                evidence=f"{detail}; adjustment columns: {parse.adjustment_headers}",
                on_dy_path=True,
            )
        ]

    return [
        Finding(
            "CHK_GPR_RECOMPUTE",
            Severity.HIGH,
            Status.FLAG,
            f"Independently rebuilt gross potential rent differs from the model by "
            f"{difference:+,.0f} ({relative:.2%}), beyond the {REVENUE_TOLERANCE:.1%} tolerance "
            f"for an individual revenue line, and the rent roll shows no step-up or free-rent "
            f"adjustment that would explain it.",
            sheet=tab.sheet,
            cell=coord,
            evidence=detail,
            on_dy_path=True,
        )
    ]


# A status that must never contribute rent: the unit is empty, or the person
# named has not taken occupancy (a pending applicant, an unsigned renewal). A
# "Pending renewal" row in particular duplicates a unit that already appears as
# Occupied, so counting its rent counts the unit twice.
_VACANT_MARKER = re.compile(r"^vacant(\s*-?\s*\w+)?\s*$", re.I)
_NOT_IN_OCCUPANCY = re.compile(r"vacant|applicant|pending|future|model|down", re.I)


def _tenant_header_row(parse: RentRollParse) -> int:
    return min(parse.rows) - 1 if parse.rows else 0


def _find_unit_column(ctx: LoanContext, parse: RentRollParse) -> str | None:
    """The column headed `Unit` (not `Unit Designation` / `Unit/Lease Status`)."""
    header_row = _tenant_header_row(parse)
    if not parse.sheet or header_row < 1:
        return None
    for row in (header_row, header_row - 1):
        if row < 1:
            continue
        for index in range(1, 41):
            column = get_column_letter(index)
            text = ctx.wb.text(parse.sheet, f"{column}{row}")
            if text and re.fullmatch(r"unit(\(s\))?\s*#?|suite\s*(id)?", text.strip(), re.I):
                return column
    return None


def _vacant_marked_rows(ctx: LoanContext, parse: RentRollParse) -> dict[int, str]:
    """Rows whose status or name cell reads as a bare `VACANT` marker.

    The marker must be the whole cell (`Vacant`, `VACANT`, `Vacant-Leased`), so a
    summary caption like `Vacant Sqft:` on Hialeah never marks a tenant row.
    """
    marked: dict[int, str] = {}
    if not parse.sheet or not parse.rows:
        return marked
    rent_index = column_index_from_string(parse.rent_column) if parse.rent_column else 30
    for index in range(1, rent_index + 2):
        column = get_column_letter(index)
        for row in parse.rows:
            text = ctx.wb.text(parse.sheet, f"{column}{row}")
            if text and _VACANT_MARKER.match(text.strip()):
                marked.setdefault(row, f"{column}{row}={text.strip()!r}")
    return marked


def check_rent_status(ctx: LoanContext) -> list[Finding]:
    """No rent from units that are vacant, pending, or counted twice.

    Two legs, both driven off the same tenant-level parse as the GPR recompute:

    - a row marked `VACANT` (by status or by name) must carry zero rent;
    - the model's GPR must not exceed the occupied-status rent total - the
      excess is rent picked up from Applicant / Pending / Vacant rows, which the
      loan definitions exclude (a pending row also duplicates a unit already
      counted as occupied, so the same unit's rent lands twice).
    """
    wb, tab = ctx.wb, ctx.tab
    coord = tab.cell(Line.GPR)
    parse = ctx.facts.get("rent_roll_parse")
    if parse is None or coord is None:
        return []
    if not parse.confident:
        return []  # the GPR recompute already reported MANUAL_REVIEW

    findings: list[Finding] = []
    scale = 12.0 if parse.monthly else 1.0

    # Leg 1: vacant-marked rows carrying rent.
    vacant_rows = _vacant_marked_rows(ctx, parse)
    vacant_with_rent = {
        row: marker
        for row, marker in vacant_rows.items()
        if abs(parse.values.get(row, 0.0)) > 1.0
    }
    if vacant_with_rent:
        sample = "; ".join(
            f"row {row} ({marker}) rent {parse.values[row]:,.0f}"
            for row, marker in sorted(vacant_with_rent.items())[:6]
        )
        total = sum(parse.values[r] for r in vacant_with_rent) * scale
        findings.append(
            Finding(
                "CHK_RENT_STATUS",
                Severity.HIGH,
                Status.FLAG,
                f"{len(vacant_with_rent)} row(s) marked VACANT on {parse.sheet!r} carry rent in "
                f"column {parse.rent_column} ({total:,.0f} annualized). A vacant unit must not "
                f"pick up rent.",
                sheet=parse.sheet,
                cell=f"{parse.rent_column}{min(vacant_with_rent)}",
                evidence=sample,
                on_dy_path=True,
            )
        )

    # Leg 2: rent from non-occupied statuses reaching the model's GPR.
    if parse.statuses:
        model_gpr = wb.number(tab.sheet, coord)
        included_rows = set(parse.included_rows())
        included_total = sum(parse.values[r] for r in included_rows)

        excluded: dict[str, tuple[int, float]] = {}
        excluded_rows: list[int] = []
        for row, status in sorted(parse.statuses.items()):
            if row in included_rows or row not in parse.values:
                continue
            rent = parse.values[row]
            if abs(rent) <= 1.0:
                continue
            count, total = excluded.get(status, (0, 0.0))
            excluded[status] = (count + 1, total + rent)
            excluded_rows.append(row)
        excluded_total = sum(total for _n, total in excluded.values())

        unit_column = _find_unit_column(ctx, parse)
        duplicated = 0
        if unit_column and excluded_rows:
            included_units = {
                ctx.wb.value(parse.sheet, f"{unit_column}{r}") for r in included_rows
            }
            included_units.discard(None)
            duplicated = sum(
                1
                for r in excluded_rows
                if ctx.wb.value(parse.sheet, f"{unit_column}{r}") in included_units
            )

        breakdown = "; ".join(
            f"{status!r}: {count} row(s), {total * scale:,.0f} annualized"
            for status, (count, total) in sorted(excluded.items())
        )
        dup_note = (
            f" {duplicated} of these rows share a unit with a row already counted as occupied, "
            f"so those units' rent is in the total twice."
            if duplicated
            else ""
        )

        if model_gpr is not None:
            gap = model_gpr - included_total * scale
            explained = excluded_total * scale
            if gap > max(1.0, abs(model_gpr) * REVENUE_TOLERANCE):
                if explained > 0 and gap >= 0.5 * explained:
                    findings.append(
                        Finding(
                            "CHK_RENT_STATUS",
                            Severity.HIGH,
                            Status.FLAG,
                            f"Gross potential rent includes {gap:,.0f} of rent from rows whose "
                            f"status says the tenant is not in occupancy ({breakdown}). The loan "
                            f"definitions count rent only from tenants in place, and the rent "
                            f"roll's own unit count excludes these rows.{dup_note}",
                            sheet=tab.sheet,
                            cell=coord,
                            evidence=(
                                f"model GPR {model_gpr:,.2f} vs occupied-status rent "
                                f"{included_total * scale:,.2f} ({len(included_rows)} row(s), "
                                f"status column {parse.status_column})"
                            ),
                            on_dy_path=True,
                        )
                    )
                else:
                    findings.append(
                        Finding(
                            "CHK_RENT_STATUS",
                            Severity.HIGH,
                            Status.MANUAL_REVIEW,
                            f"The model's GPR exceeds the occupied-status rent total by "
                            f"{gap:,.0f}, which rent on non-occupied rows does not explain "
                            f"({breakdown or 'no rent on non-occupied rows'}). Trace the "
                            f"difference by hand.",
                            sheet=tab.sheet,
                            cell=coord,
                            evidence=(
                                f"model GPR {model_gpr:,.2f} vs occupied-status rent "
                                f"{included_total * scale:,.2f}"
                            ),
                            on_dy_path=True,
                        )
                    )
            elif not vacant_with_rent:
                detail = (
                    f"rent on non-occupied rows is correctly excluded ({breakdown})"
                    if excluded
                    else "no rent sits on any non-occupied row"
                )
                findings.append(
                    Finding(
                        "CHK_RENT_STATUS",
                        Severity.HIGH,
                        Status.PASS,
                        f"Gross potential rent draws only on occupied-status rows: {detail}.",
                        sheet=tab.sheet,
                        cell=coord,
                        evidence=(
                            f"model GPR {model_gpr:,.2f} = occupied-status rent over "
                            f"{len(included_rows)} row(s), status column {parse.status_column}"
                        ),
                        on_dy_path=True,
                    )
                )

    return findings


def check_vacancy_recompute(ctx: LoanContext) -> list[Finding]:
    """Rebuild the vacancy line from occupancy and the agreement's floor."""
    wb, tab, params = ctx.wb, ctx.tab, ctx.params
    coord = tab.cell(Line.VACANCY)
    gpr_coord = tab.cell(Line.GPR)
    if coord is None or gpr_coord is None or params is None or not params.vacancy_floor.found:
        return [
            Finding(
                "CHK_VACANCY_RECOMPUTE",
                Severity.HIGH,
                Status.UNVERIFIABLE,
                "Vacancy could not be recomputed: the vacancy line, GPR line or the "
                "agreement's floor is missing.",
                sheet=tab.sheet,
                cell=coord,
            )
        ]

    floor = float(params.vacancy_floor.value)
    threshold = 1.0 - floor
    occupancy = None
    occ_row = tab.rows.get(Line.OCCUPANCY)
    if occ_row:
        for column in (tab.dy_column, "H"):
            occupancy = wb.number(tab.sheet, f"{column}{occ_row}")
            if occupancy is not None:
                break
    gpr = wb.number(tab.sheet, gpr_coord)
    if occupancy is None or gpr is None:
        return [
            Finding(
                "CHK_VACANCY_RECOMPUTE",
                Severity.HIGH,
                Status.MANUAL_REVIEW,
                "Occupancy or GPR could not be read, so the vacancy line was not recomputed.",
                sheet=tab.sheet,
                cell=coord,
                on_dy_path=True,
            )
        ]

    # In-place GPR already reflects actual vacancy, so only the shortfall to the
    # floor is deducted, and only when occupancy is above the threshold.
    if occupancy > threshold and occupancy > 0:
        expected = -(gpr / occupancy) * (floor - (1 - occupancy))
    else:
        expected = 0.0

    model_value = wb.number(tab.sheet, coord)
    if model_value is None:
        model_value = 0.0  # a blank vacancy line is no deduction
    difference = expected - model_value
    scale = max(abs(gpr), 1.0)
    evidence = (
        f"occupancy {occupancy:.4%}, floor {floor:.2%}, GPR {gpr:,.2f} -> expected "
        f"{expected:,.2f}; model {tab.sheet}!{coord} = {model_value:,.2f}"
    )

    if abs(difference) / scale <= REVENUE_TOLERANCE:
        return [
            Finding(
                "CHK_VACANCY_RECOMPUTE",
                Severity.HIGH,
                Status.PASS,
                (
                    f"Vacancy line ties to an independent recompute ({expected:,.0f})."
                    if expected
                    else f"Vacancy is correctly 0: actual vacancy {1 - occupancy:.2%} already "
                    f"exceeds the {floor:.2%} floor and is embedded in in-place rent."
                ),
                sheet=tab.sheet,
                cell=coord,
                evidence=evidence,
                on_dy_path=True,
            )
        ]

    return [
        Finding(
            "CHK_VACANCY_RECOMPUTE",
            Severity.HIGH,
            Status.FLAG,
            f"Vacancy line differs from an independent recompute by {difference:+,.0f}: expected "
            f"{expected:,.0f} at {occupancy:.2%} occupancy against the {floor:.2%} floor, model "
            f"shows {model_value:,.0f}.",
            sheet=tab.sheet,
            cell=coord,
            evidence=evidence,
            on_dy_path=True,
        )
    ]


#: `check_rent_status` reads the parse that `check_gpr_recompute` stores, so it
#: must run after it.
RECOMPUTE_CHECKS = (check_gpr_recompute, check_rent_status, check_vacancy_recompute)


def run_recompute(ctx: LoanContext) -> list[Finding]:
    findings: list[Finding] = []
    for check in RECOMPUTE_CHECKS:
        findings.extend(check(ctx))
    return findings
