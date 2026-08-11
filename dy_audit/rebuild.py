"""Independent rent-roll rebuild (the top-line revenue deep dive).

Every rent roll carries a column of tenant rent. This module finds that column
on its own - by reading the export's headers, not by trusting the model's
formulas - decides which tenants belong in base rent, rebuilds the annualized
figure, and reconciles it against the OSAR's Gross Potential Rent line. The
GPR line is base rent only: supplementary billing (pet rent, utility
reimbursements, parking add-ons) does not belong in it.

The rebuild also captures the forensic signals a tie-out alone would miss:

- formula cells sitting inside an otherwise-literal exported rent column
  (an analyst typing over the export - e.g. backfilling last quarter's rent);
- one formula shaped differently from the rest of a derived column;
- the export's own Total rows disagreeing with the sum of the rows above them
  (the sheet was edited after export);
- rent kept on tenants with delinquent balances or scheduled move-outs while
  the OSAR's own note claims they were excluded.

Where the header scan fails (Campus's header row is placeholder `x`es on one
row - the real headers sit two rows up), the model-chain parse from
`recompute.py` supplies the rows instead, and the sheet says so.
"""

from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field

from openpyxl.utils import column_index_from_string, get_column_letter

from . import formula as F
from .context import LoanContext
from .model import Finding, Severity, Status
from .osar import Line, excel_to_date

#: Tie tolerance, matching the individual-revenue-line rule (spec section 4).
TOLERANCE = 0.001

# ---------------------------------------------------------------------------
# Header vocabulary
# ---------------------------------------------------------------------------

#: Rent-column priorities, first match wins. "Market Rent" is never the tenant
#: rent column - it is what a vacant unit *would* fetch.
_RENT_PRIORITY = (
    re.compile(r"dy\s*test.*rent|rent.*dy\s*test", re.I),
    re.compile(r"^rent$", re.I),
    re.compile(r"^actual\s*rent$", re.I),
    re.compile(r"^lease\s*rent$", re.I),
    re.compile(r"^monthly\s*(base\s*)?rent$", re.I),
    re.compile(r"^(monthly\s*)?base\s*rent$", re.I),
    re.compile(r"^annual\s*(base\s*)?rent$", re.I),
    re.compile(r"adj\w*\s*ann\w*\s*rent", re.I),
)
_ANY_RENT = re.compile(r"\brent\b", re.I)
_MARKET = re.compile(r"market", re.I)
_UNIT_HEADER = re.compile(r"^unit(\(s\))?\s*#?$|^suite(\s*id)?$|^apt\b", re.I)
_NAME_HEADER = re.compile(
    r"^name$|^(resident|tenant|occupant|lessee)(\s*name)?$|^lease$|^occupant$", re.I
)
_STATUS_HEADER = re.compile(r"status", re.I)
_BALANCE_HEADER = re.compile(r"^balance$", re.I)
_MOVE_OUT_HEADER = re.compile(r"move\s*-?\s*out", re.I)
_SQFT_HEADER = re.compile(r"sq\s*ft|sqft|^area$|^gla$", re.I)

#: Statuses (or names) that exclude a row from base rent. Everything else is
#: included - an unfamiliar but occupied-sounding status must not silently
#: drop rent.
_EXCLUDE_STATUS = re.compile(
    r"vacant|applicant|pending|future|former|down|admin|model\b", re.I
)
_VACANT_NAME = re.compile(r"^vacant(\s*-?\s*\w+)?\s*$", re.I)

_TOTAL_LABEL = re.compile(r"^total", re.I)
#: Headers that mark documented per-tenant adjustments the model may apply
#: (free-rent proration, months included, an analyst include flag). Their
#: presence means raw contractual rent legitimately differs from the model.
_ADJUSTMENT_HEADER = re.compile(
    r"free\s*rent|months?\s*(in|ex)cluded|include\s*in\s*dy|adj\w*\s*ann", re.I
)
_SECTION_APPLICANT = re.compile(r"future|applicant|pending", re.I)
_SECTION_RESET = re.compile(r"current|occupied", re.I)
_AS_OF = re.compile(r"as\s*.?of\s*(date)?\s*[:=]?\s*(.+)", re.I)

#: The OSAR's note beside GPR claiming exclusions the formula must then apply.
_NOTE_DELINQ = re.compile(r"delinq", re.I)
_NOTE_KV = re.compile(r"\bkv\b|known\s*vacate", re.I)

#: Non-base-rent billing codes (supplementary income / reimbursements).
SUPPLEMENTARY_LABEL = re.compile(
    r"pet|gas\b|util|trash|package|amen|carport|storage|tech\s*fee|parking|valet|"
    r"laundry|late\s*fee|reimb|trv\b|smarthome|renters?ins|insur|deposit",
    re.I,
)
CONCESSION_LABEL = re.compile(r"concession|free\s*rent|write\s*-?off", re.I)


def _clean(text) -> str:
    return re.sub(r"\s+", " ", str(text)).strip() if text is not None else ""


def _to_number(value) -> float | None:
    """A number, tolerating the comma-formatted text of static export totals."""
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        text = value.replace(",", "").replace("$", "").strip()
        if re.fullmatch(r"-?\d+(\.\d+)?", text):
            return float(text)
    return None


@dataclass
class TenantRow:
    row: int
    unit: str = ""
    name: str = ""
    status: str = ""
    sqft: float | None = None
    monthly_rent: float = 0.0
    market_rent: float | None = None
    balance: float | None = None
    move_out: dt.date | None = None
    rent_is_formula: bool = False
    rent_formula: str | None = None
    included: bool = True
    reason: str = ""


@dataclass
class TotalRow:
    row: int
    label: str
    reported: float
    block_sum: float
    grand_sum: float
    #: The same sums over included (occupied) rows only - some exports total on
    #: an occupied-only basis, which is a basis difference, not tampering.
    block_included: float = 0.0
    grand_included: float = 0.0

    @property
    def ties(self) -> bool:
        scale = max(abs(self.reported), 1.0)
        return any(
            abs(self.reported - basis) / scale <= 0.005
            for basis in (self.block_sum, self.grand_sum, self.block_included, self.grand_included)
        )


@dataclass
class Rebuild:
    """One rent-roll tab rebuilt from its own tenant table."""

    sheet: str | None = None
    line: Line = Line.GPR
    source: str = "headers"  # or "model-chain"
    header_row: int | None = None
    columns: dict[str, str] = field(default_factory=dict)
    rent_header: str = ""
    monthly: bool = True
    rows: list[TenantRow] = field(default_factory=list)
    totals: list[TotalRow] = field(default_factory=list)
    as_of: str | None = None
    notes: list[str] = field(default_factory=list)
    adjustment_headers: list[str] = field(default_factory=list)
    #: Reconciliation lines built by `reconcile`: (item, annual amount or None).
    reconciliation: list[tuple[str, float | None]] = field(default_factory=list)
    model_gpr: float | None = None
    model_note: str | None = None

    @property
    def confident(self) -> bool:
        return self.sheet is not None and any(r.included for r in self.rows)

    @property
    def scale(self) -> float:
        return 12.0 if self.monthly else 1.0

    def included_rows(self) -> list[TenantRow]:
        return [r for r in self.rows if r.included]

    def annualized(self) -> float:
        return sum(r.monthly_rent for r in self.included_rows()) * self.scale

    def excluded_with_rent(self) -> dict[str, tuple[int, float]]:
        out: dict[str, tuple[int, float]] = {}
        for r in self.rows:
            if r.included or abs(r.monthly_rent) <= 1.0:
                continue
            key = r.status or r.reason or "excluded"
            n, total = out.get(key, (0, 0.0))
            out[key] = (n + 1, total + r.monthly_rent)
        return out


# ---------------------------------------------------------------------------
# Header detection
# ---------------------------------------------------------------------------


def _joined_headers(ctx: LoanContext, sheet: str, row: int, max_col: int) -> dict[int, str]:
    """Header text per column, joining a two-row header vertically."""
    out: dict[int, str] = {}
    for col in range(1, max_col + 1):
        top = _clean(ctx.wb.text(sheet, f"{get_column_letter(col)}{row}"))
        below = _clean(ctx.wb.text(sheet, f"{get_column_letter(col)}{row + 1}"))
        joined = " ".join(part for part in (top, below) if part)
        if joined:
            out[col] = joined.replace("\n", " ")
    return out


def _pick_rent_column(headers: dict[int, str]) -> tuple[int | None, str, int]:
    for rank, pattern in enumerate(_RENT_PRIORITY):
        for col in sorted(headers):
            text = headers[col]
            if _MARKET.search(text):
                continue
            if pattern.search(text):
                return col, text, rank
    return None, "", len(_RENT_PRIORITY)


def _single_row_headers(ctx: LoanContext, sheet: str, row: int, max_col: int) -> dict[int, str]:
    out: dict[int, str] = {}
    for col in range(1, max_col + 1):
        text = _clean(ctx.wb.text(sheet, f"{get_column_letter(col)}{row}"))
        if text:
            out[col] = text.replace("\n", " ")
    return out


def _interpret(headers: dict[int, str]) -> tuple[dict[str, int], str, int] | None:
    unit_col = next((c for c in sorted(headers) if _UNIT_HEADER.match(headers[c])), None)
    if unit_col is None:
        return None
    name_col = next((c for c in sorted(headers) if _NAME_HEADER.match(headers[c])), None)
    status_col = next((c for c in sorted(headers) if _STATUS_HEADER.search(headers[c])), None)
    if name_col is None and status_col is None:
        return None
    rent_col, rent_header, rank = _pick_rent_column(headers)
    if rent_col is None:
        return None
    columns = {"unit": unit_col, "rent": rent_col}
    if name_col is not None:
        columns["name"] = name_col
    if status_col is not None:
        columns["status"] = status_col
    for role, pattern in (
        ("balance", _BALANCE_HEADER),
        ("move_out", _MOVE_OUT_HEADER),
        ("sqft", _SQFT_HEADER),
    ):
        found = next((c for c in sorted(headers) if pattern.search(headers[c])), None)
        if found is not None and found not in columns.values():
            columns[role] = found
    market = next(
        (
            c
            for c in sorted(headers)
            if _MARKET.search(headers[c]) and _ANY_RENT.search(headers[c])
        ),
        None,
    )
    if market is not None:
        columns["market"] = market
    return columns, rent_header, rank


def find_tenant_table(
    ctx: LoanContext, sheet: str
) -> tuple[int | None, dict[str, int], str, bool]:
    """Locate the header row and map roles to columns.

    Returns (header_row, {role: column index}, rent_header, two_row_header).
    Each candidate row is read two ways - alone, and joined with the row below
    (two-row headers like `Actual` / `Rent`) - and the reading that yields the
    stronger rent column wins. A single-row header must not be joined with the
    first data row: `RENT` + a rent figure no longer reads as a header.
    """
    rows_bound, cols_bound = ctx.wb.bounds(sheet)
    max_col = min(cols_bound, 60)
    for row in range(1, min(rows_bound, 40) + 1):
        single = _single_row_headers(ctx, sheet, row, max_col)
        alone = _interpret(single)
        joined = _interpret(_joined_headers(ctx, sheet, row, max_col))
        # The unit header must sit in this row itself, or a title row above the
        # real header would claim the join.
        if joined is not None:
            unit_col = joined[0]["unit"]
            if unit_col not in single or not _UNIT_HEADER.match(single[unit_col]):
                joined = None
        # A rent column must hold numbers below the header. The top half of a
        # two-row header can masquerade as one ("Rent" over "Steps") and would
        # otherwise win on priority while holding prose.
        def _has_numbers(candidate) -> bool:
            if candidate is None:
                return False
            column = get_column_letter(candidate[0]["rent"])
            return any(
                isinstance(ctx.wb.value(sheet, f"{column}{r}"), (int, float))
                for r in range(row + 1, row + 42)
            )

        if not _has_numbers(alone):
            alone = None
        if not _has_numbers(joined):
            joined = None
        best = None
        two_row = False
        if alone and (not joined or alone[2] <= joined[2]):
            best = alone
        elif joined:
            best, two_row = joined, True
        if best is None:
            continue
        columns, rent_header, _rank = best
        return row, columns, rent_header, two_row
    return None, {}, "", False


# ---------------------------------------------------------------------------
# Row extraction
# ---------------------------------------------------------------------------


def _first_text(ctx: LoanContext, sheet: str, row: int, upto_col: int) -> str:
    for col in range(1, upto_col + 1):
        text = ctx.wb.text(sheet, f"{get_column_letter(col)}{row}")
        if text and text.strip():
            return text.strip()
    return ""


def extract_rows(ctx: LoanContext, sheet: str, header_row: int, columns: dict[str, int],
                 rent_header: str, two_row: bool = False,
                 stop_row: int | None = None) -> Rebuild:
    rebuild = Rebuild(sheet=sheet, header_row=header_row, rent_header=rent_header)
    rebuild.columns = {role: get_column_letter(col) for role, col in columns.items()}
    rebuild.monthly = not re.search(r"annual", rent_header, re.I)
    rebuild.adjustment_headers = [
        f"{get_column_letter(col)}: {text[:40]}"
        for col, text in sorted(_joined_headers(ctx, sheet, header_row, 60).items())
        if _ADJUSTMENT_HEADER.search(text)
    ]
    wb = ctx.wb
    rows_bound, _ = wb.bounds(sheet)
    rent_letter = rebuild.columns["rent"]
    unit_letter = rebuild.columns["unit"]

    for row in range(1, header_row):
        text = _first_text(ctx, sheet, row, 12)
        if text and _AS_OF.search(text):
            rebuild.as_of = text[:80]
            break

    applicant_context = False
    block_sum = 0.0
    grand_sum = 0.0
    block_included = 0.0
    grand_included = 0.0
    rows_since_tenant = 0
    start = header_row + 2 if two_row else header_row + 1

    for row in range(start, min(rows_bound, 3000) + 1):
        # Rent rolls end and scrap blocks follow (floorplan averages, #REF!
        # scratch areas). A long run without a tenant row means the table is
        # over, and everything below must not leak in.
        if rebuild.rows and rows_since_tenant > 25:
            break
        if stop_row is not None and row > stop_row:
            break
        rows_since_tenant += 1
        unit_value = wb.value(sheet, f"{unit_letter}{row}")
        rent_raw = wb.value(sheet, f"{rent_letter}{row}")
        rent = _to_number(rent_raw)

        # Total rows: a Total-ish label to the left of the rent column.
        total_label = ""
        for col in range(1, min(columns["rent"], 9)):
            text = wb.text(sheet, f"{get_column_letter(col)}{row}")
            if text and _TOTAL_LABEL.match(text.strip()):
                total_label = text.strip()
                break
        if total_label:
            reported = _to_number(rent_raw)
            if reported is not None:
                rebuild.totals.append(
                    TotalRow(
                        row, total_label[:40], reported,
                        block_sum, grand_sum, block_included, grand_included,
                    )
                )
            block_sum = 0.0
            block_included = 0.0
            applicant_context = False
            rows_since_tenant = 0
            continue

        # Section captions ("Future Residents/Applicants", "Current Leases") -
        # either on an otherwise empty row or sitting in the unit column itself.
        lead = _first_text(ctx, sheet, row, min(columns["rent"], 8))
        unit_text = _clean(unit_value) if unit_value is not None else ""
        wordy_unit = len(unit_text) > 12 and not any(ch.isdigit() for ch in unit_text)
        if (unit_value is None and rent is None and lead) or wordy_unit:
            caption = unit_text or lead
            if _SECTION_APPLICANT.search(caption):
                applicant_context = True
            elif _SECTION_RESET.search(caption):
                applicant_context = False
            continue

        if not unit_text or _TOTAL_LABEL.match(unit_text):
            continue

        name = _clean(wb.value(sheet, f"{rebuild.columns['name']}{row}")) if "name" in rebuild.columns else ""
        status = _clean(wb.value(sheet, f"{rebuild.columns['status']}{row}")) if "status" in rebuild.columns else ""
        # A floorplan-summary block reusing the tenant columns puts numbers
        # where names and statuses belong; those are not tenant rows.
        if _to_number(name) is not None:
            name = ""
        if _to_number(status) is not None:
            status = ""
        if name.lower() == "x":
            name = ""
        # Where the table names its tenants, a row with neither a name nor a
        # status is a summary row reusing the tenant columns, not a tenant.
        if ("name" in rebuild.columns or "status" in rebuild.columns) and not name and not status:
            continue
        if rent is None and not name and not status:
            continue

        tenant = TenantRow(row=row, unit=unit_text, name=name, status=status)
        tenant.monthly_rent = rent or 0.0
        literal = wb.literal(sheet, f"{rent_letter}{row}")
        tenant.rent_is_formula = isinstance(literal, str) and literal.startswith("=")
        if tenant.rent_is_formula:
            tenant.rent_formula = literal
        if "sqft" in rebuild.columns:
            tenant.sqft = _to_number(wb.value(sheet, f"{rebuild.columns['sqft']}{row}"))
        if "market" in rebuild.columns:
            tenant.market_rent = _to_number(wb.value(sheet, f"{rebuild.columns['market']}{row}"))
        if "balance" in rebuild.columns:
            tenant.balance = _to_number(wb.value(sheet, f"{rebuild.columns['balance']}{row}"))
        if "move_out" in rebuild.columns:
            tenant.move_out = excel_to_date(wb.value(sheet, f"{rebuild.columns['move_out']}{row}"))

        if applicant_context:
            tenant.included, tenant.reason = False, "listed under Future Residents/Applicants"
        elif status and _EXCLUDE_STATUS.search(status):
            tenant.included, tenant.reason = False, f"status {status!r}"
        elif name and _VACANT_NAME.match(name):
            tenant.included, tenant.reason = False, "unit is vacant (name reads VACANT)"
            tenant.status = tenant.status or "Vacant"
        else:
            tenant.included = True
            tenant.reason = f"status {status!r}" if status else "in occupancy"
        rebuild.rows.append(tenant)
        rows_since_tenant = 0
        # Export totals sum everything they printed, so the comparison basis is
        # all tenant rows first - with the included-only basis kept alongside.
        block_sum += tenant.monthly_rent
        grand_sum += tenant.monthly_rent
        if tenant.included:
            block_included += tenant.monthly_rent
            grand_included += tenant.monthly_rent

    # Some exports net an adjustment section below their Totals row
    # ("Historically generated Rent - back-dated move-ins/outs..."). Such a
    # total legitimately differs from the plain column sum, so it cannot be
    # used as a tamper check.
    for total in list(rebuild.totals):
        for below in range(total.row + 1, total.row + 9):
            text = _first_text(ctx, sheet, below, 6)
            if text and re.search(r"historically\s+generated|back-?dated", text, re.I):
                rebuild.totals.remove(total)
                rebuild.notes.append(
                    f"export total at row {total.row} nets an adjustment section printed "
                    f"below it ({text[:60]!r}) and is not compared"
                )
                break
    return rebuild


def _rebuild_from_parse(ctx: LoanContext, line: Line) -> Rebuild:
    """Fallback: rows from the model-chain parse (`recompute.parse_rent_roll`)."""
    rebuild = Rebuild(source="model-chain", line=line)
    parse = ctx.facts.get("rent_roll_parse")
    if parse is None or not parse.confident:
        rebuild.notes.append("no header row found and the model-chain parse is not confident")
        return rebuild
    rebuild.sheet = parse.sheet
    rebuild.rent_header = f"column {parse.rent_column} (from the model's own GPR chain)"
    rebuild.columns = {"rent": parse.rent_column}
    rebuild.monthly = parse.monthly
    included = set(parse.included_rows())
    for row in sorted(parse.values):
        status = parse.statuses.get(row, "") if parse.statuses else ""
        tenant = TenantRow(
            row=row,
            status=status,
            monthly_rent=parse.values[row],
            included=row in included,
            reason=f"status {status!r}" if status else "no status column; row summed by the model",
        )
        formula = ctx.wb.formula(parse.sheet, f"{parse.rent_column}{row}")
        tenant.rent_is_formula = formula is not None
        tenant.rent_formula = formula
        rebuild.rows.append(tenant)
    rebuild.notes.append(
        "columns taken from the model's own GPR chain - no independent header row was found"
    )
    return rebuild


# ---------------------------------------------------------------------------
# Reconciliation
# ---------------------------------------------------------------------------


def _read_model_note(ctx: LoanContext, row: int) -> str | None:
    """The Notes text the OSAR carries beside a revenue line."""
    for col_index in range(column_index_from_string(ctx.tab.dy_column) + 1,
                           column_index_from_string(ctx.tab.dy_column) + 6):
        text = ctx.wb.text(ctx.sheet, f"{get_column_letter(col_index)}{row}")
        if text and len(text.strip()) > 8 and not _to_number(text):
            return text.strip()
    return None


def _chain_contributions(ctx: LoanContext, coord: str) -> list[tuple[str, str, float | None]]:
    """(label, ref, exact annual contribution) for each direct term of the GPR formula.

    The contribution is measured by re-evaluating the formula with that term
    zeroed - which handles `(a+b+c)*12 + d` scaling exactly. Terms the
    restricted evaluator cannot handle report None.
    """
    wb, sheet = ctx.wb, ctx.sheet
    formula = wb.formula(sheet, coord)
    if formula is None:
        return []
    out: list[tuple[str, str, float | None]] = []
    try:
        base = F.evaluate(formula, sheet, F.make_resolver(wb))
    except Exception:  # noqa: BLE001 - unsupported grammar means no attribution
        return []
    if not isinstance(base, (int, float)):
        return []
    from .checks.high import nearby_label, _column_header

    for ref in F.iter_refs(formula):
        if ref.is_range or ref.external_index is not None:
            continue
        ref_sheet = ref.resolved_sheet(sheet)
        if not wb.has_sheet(ref_sheet):
            continue
        try:
            zeroed = F.evaluate(
                formula, sheet, F.make_resolver(wb, overrides={f"{ref_sheet}!{ref.coord}": 0.0})
            )
        except Exception:  # noqa: BLE001
            zeroed = None
        contribution = (
            float(base) - float(zeroed) if isinstance(zeroed, (int, float)) else None
        )
        label = nearby_label(ctx, ref_sheet, ref.coord) or _column_header(ctx, ref_sheet, ref.coord) or ""
        out.append((label, f"{ref_sheet}!{ref.coord}", contribution))
    return out


def reconcile(ctx: LoanContext, rebuild: Rebuild, gpr_coord: str) -> None:
    """The second, deeper dive: itemize why the rebuild differs from the model."""
    wb = ctx.wb
    model = wb.number(ctx.sheet, gpr_coord)
    rebuild.model_gpr = model
    row_number = int(re.sub(r"\D", "", gpr_coord) or 0)
    rebuild.model_note = _read_model_note(ctx, row_number)
    if model is None or not rebuild.confident:
        return
    rebuilt = rebuild.annualized()
    difference = model - rebuilt
    items = rebuild.reconciliation
    explained = 0.0

    # 1. Rent on rows the rebuild excluded (the model may still count them).
    excluded = rebuild.excluded_with_rent()
    excluded_total = sum(total for _n, total in excluded.values()) * rebuild.scale
    if excluded_total and difference > max(1.0, abs(model) * TOLERANCE):
        for status, (n, total) in sorted(excluded.items()):
            items.append(
                (f"rent on {n} excluded row(s), {status}", total * rebuild.scale)
            )
        explained += min(excluded_total, difference)

    # 2. The model sums its own derived per-row column (Hialeah's V). Compare
    # that column row by row against the raw rent, so a multiplier applied to
    # one tenant is named with its exact dollars.
    gpr_formula = wb.formula(ctx.sheet, gpr_coord)
    rent_letter = rebuild.columns.get("rent")
    for ref in F.iter_refs(gpr_formula):
        if not ref.is_range or ref.resolved_sheet(ctx.sheet) != rebuild.sheet:
            continue
        derived_letter = re.sub(r"[\d$:]", "", ref.body.split(":")[0]).upper()
        if not derived_letter or derived_letter == rent_letter:
            continue
        for tenant in rebuild.included_rows():
            derived = wb.number(rebuild.sheet, f"{derived_letter}{tenant.row}")
            if derived is None:
                continue
            expected = tenant.monthly_rent * rebuild.scale
            delta = derived - expected
            if abs(delta) > max(1.0, abs(expected) * 0.001):
                cell_formula = F.normalize(
                    wb.formula(rebuild.sheet, f"{derived_letter}{tenant.row}") or ""
                )
                items.append(
                    (
                        f"model's derived column {derived_letter}{tenant.row} carries "
                        f"{derived:,.0f} vs {expected:,.0f} rebuilt"
                        + (f" (formula: {cell_formula[:40]})" if cell_formula else ""),
                        delta,
                    )
                )
                explained += delta

    # 3. GPR-chain terms that are not base rent (supplementary or concession).
    for label, ref, contribution in _chain_contributions(ctx, gpr_coord):
        if contribution is None or abs(contribution) < 1.0:
            continue
        if CONCESSION_LABEL.search(label) or contribution < 0:
            items.append((f"deduction in the GPR formula: {ref} ({label or 'unlabelled'})", contribution))
            explained += contribution
        elif SUPPLEMENTARY_LABEL.search(label):
            items.append(
                (f"supplementary billing in the GPR formula: {ref} ({label})", contribution)
            )
            explained += contribution

    # 3. Analyst edits inside the rent column (suspicious ones only - same-row
    # arithmetic is reported separately as a note, not a reconciliation item).
    formula_rows = [r for r in rebuild.rows if r.rent_is_formula and abs(r.monthly_rent) > 1.0]
    literal_rows = [r for r in rebuild.rows if not r.rent_is_formula and abs(r.monthly_rent) > 1.0]
    if formula_rows and len(literal_rows) >= 4 * len(formula_rows):
        suspicious = [r for r in formula_rows if _suspicious_edit(ctx, rebuild, r)]
        if suspicious:
            total = sum(r.monthly_rent for r in suspicious) * rebuild.scale
            items.append(
                (
                    f"{len(suspicious)} rent cell(s) are formulas inside an otherwise "
                    f"literal export column, pulling rent from outside this export "
                    f"(rows {', '.join(str(r.row) for r in suspicious[:10])})",
                    total,
                )
            )

    # 4. The export's own totals that no longer tie.
    for total in rebuild.totals:
        if not total.ties:
            items.append(
                (
                    f"export total {total.label!r} at row {total.row} reads "
                    f"{total.reported:,.0f} but the rows above sum to {total.block_sum:,.0f} "
                    f"(cumulative {total.grand_sum:,.0f}) - the sheet was edited after export",
                    None,
                )
            )

    # 5. Delinquent balances and scheduled move-outs the OSAR note says are out.
    # Where the workbook carries an AR/aging tab, that tab is the source of
    # truth for delinquency (spec section 5.2) and CHK_EXCLUSIONS tests it; the
    # rent roll's balance column mixes fees and timing and stays a note only.
    note = rebuild.model_note or ""
    has_ar = bool((ctx.facts.get("source_tabs") or {}).get("ar"))
    delinquent = [
        r
        for r in rebuild.included_rows()
        if r.balance is not None and r.balance > max(500.0, r.monthly_rent)
    ]
    if delinquent and has_ar:
        rebuild.notes.append(
            f"{len(delinquent)} included tenant(s) show a rent-roll balance above one "
            f"month's rent; delinquency is tested against the AR aging (see CHK_EXCLUSIONS)"
        )
    elif delinquent:
        total = sum(r.monthly_rent for r in delinquent) * rebuild.scale
        claim = (
            " - the OSAR note says delinquent tenants are excluded, but these are in"
            if _NOTE_DELINQ.search(note)
            else " and no AR aging tab exists to test them against"
        )
        items.append(
            (
                f"{len(delinquent)} included tenant(s) carry a balance above one month's rent"
                f"{claim}",
                total,
            )
        )
    if rebuild.line is Line.GPR:
        movers = [r for r in rebuild.included_rows() if r.move_out is not None]
        if movers and _NOTE_KV.search(note):
            total = sum(r.monthly_rent for r in movers) * rebuild.scale
            items.append(
                (
                    f"{len(movers)} included tenant(s) have a scheduled move-out - the OSAR "
                    f"note says known vacates are excluded, but their rent is in",
                    total,
                )
            )

    residual = difference - explained
    items.append(("difference (model minus rebuild)", difference))
    if abs(residual) > max(1.0, abs(model) * TOLERANCE) and explained:
        items.append(("unexplained after the items above", residual))


# ---------------------------------------------------------------------------
# Companion checks driven by the rebuilt table
# ---------------------------------------------------------------------------


def _strip_rows(formula: str) -> str:
    """Formula shape with row numbers removed, for outlier comparison."""
    return re.sub(r"(?<=[A-Z$])\d+", "#", F.normalize(formula))


#: A column header naming another period: `4Q25 Rent`, `Prior Rent`.
_PRIOR_PERIOD = re.compile(r"\b[1-4]\s*Q\s*'?\d{2,4}\b|prior|prev|last\s+(qtr|quarter|year)", re.I)


def _column_top_header(ctx: LoanContext, sheet: str, column: str, before_row: int) -> str:
    for row in range(1, before_row + 3):
        text = ctx.wb.text(sheet, f"{column}{row}")
        if text and text.strip():
            return text.strip()
    return ""


def _suspicious_edit(ctx: LoanContext, rebuild: Rebuild, tenant: TenantRow) -> str | None:
    """Why a formula in the rent column looks like a backfill, or None if benign.

    Benign: arithmetic over the row's own other columns (a rate times an
    area). Suspicious: a reference into another workbook, or a pull from a
    column whose header names another period - last quarter's rent carried
    into this quarter's column.
    """
    formula = tenant.rent_formula
    if F.has_external_refs(formula):
        return "references another workbook"
    for ref in F.iter_refs(formula):
        if ref.is_range:
            continue
        column = re.sub(r"[\d$]", "", ref.coord)
        if not column or column == rebuild.columns.get("rent"):
            continue
        header = _column_top_header(
            ctx, ref.resolved_sheet(rebuild.sheet or ""), column, tenant.row
        )
        if header and _PRIOR_PERIOD.search(header):
            return f"pulls from column {column} headed {header[:24]!r} (another period)"
    return None


def check_rent_edits(ctx: LoanContext, rebuild: Rebuild) -> list[Finding]:
    """Hand edits inside the rent column.

    An exported rent column is either literals (the export) or a uniform
    formula (a derived column). A formula inside a literal column is an
    analyst typing over the export: it flags when it reaches outside the row's
    own data (Quincy backfilling last quarter's rent onto vacated units), and
    is noted when it is same-row arithmetic (Campus computing rate x area).
    One formula shaped unlike the rest of a derived column is a single-row
    change - a haircut applied to one tenant and nobody else.
    """
    if not rebuild.confident or rebuild.sheet is None:
        return []
    rent_rows = [r for r in rebuild.rows if abs(r.monthly_rent) > 1.0 or r.rent_is_formula]
    formulas = [r for r in rent_rows if r.rent_is_formula]
    literals = [r for r in rent_rows if not r.rent_is_formula]
    column = rebuild.columns.get("rent", "?")

    if formulas and len(literals) >= 4 * len(formulas):
        suspicious = [
            (r, reason)
            for r in formulas
            if (reason := _suspicious_edit(ctx, rebuild, r)) is not None
        ]
        if suspicious:
            total = sum(r.monthly_rent for r, _reason in suspicious) * rebuild.scale
            sample = "; ".join(
                f"{column}{r.row} {F.normalize(r.rent_formula or '')} = "
                f"{r.monthly_rent:,.0f} ({reason})"
                for r, reason in suspicious[:10]
            )
            return [
                Finding(
                    "CHK_RENT_EDITS",
                    Severity.HIGH,
                    Status.FLAG,
                    f"{len(suspicious)} cell(s) in rent column {column} on "
                    f"{rebuild.sheet!r} are formulas inside an otherwise literal export "
                    f"column, and they pull rent from outside this quarter's export - "
                    f"together {total:,.0f} of annualized rent that the export itself does "
                    f"not show. Verify each unit against the source rent roll.",
                    sheet=rebuild.sheet,
                    cell=f"{column}{suspicious[0][0].row}",
                    evidence=sample,
                    on_dy_path=True,
                )
            ]
        sample = "; ".join(
            f"{column}{r.row} = {F.normalize(r.rent_formula or '')}" for r in formulas[:8]
        )
        return [
            Finding(
                "CHK_RENT_EDITS",
                Severity.HIGH,
                Status.MANUAL_REVIEW,
                f"{len(formulas)} cell(s) in rent column {column} on {rebuild.sheet!r} are "
                f"formulas inside an otherwise literal export column ({len(literals)} "
                f"literal rows). They compute from the row's own data, which is usually an "
                f"analyst deriving a figure the export left blank - spot-check them.",
                sheet=rebuild.sheet,
                cell=f"{column}{formulas[0].row}",
                evidence=sample,
                on_dy_path=True,
            )
        ]

    if len(formulas) >= 5 and len(literals) <= len(formulas) // 4:
        shapes: dict[str, list[TenantRow]] = {}
        for r in formulas:
            shapes.setdefault(_strip_rows(r.rent_formula or ""), []).append(r)
        if len(shapes) > 1:
            majority = max(shapes, key=lambda s: len(shapes[s]))
            outliers = [
                (shape, rows)
                for shape, rows in shapes.items()
                if shape != majority and len(rows) <= max(2, len(formulas) // 5)
            ]
            if outliers:
                sample = "; ".join(
                    f"{column}{r.row} = {F.normalize(r.rent_formula or '')}"
                    for _shape, rows in outliers
                    for r in rows[:6]
                )
                count = sum(len(rows) for _s, rows in outliers)
                return [
                    Finding(
                        "CHK_RENT_EDITS",
                        Severity.HIGH,
                        Status.FLAG,
                        f"{count} cell(s) in rent column {column} on {rebuild.sheet!r} use a "
                        f"formula shaped differently from the rest of the column "
                        f"(majority: {majority[:60]}). A single-row change in a derived rent "
                        f"column is how one tenant's rent quietly diverges from the rule "
                        f"applied to every other tenant.",
                        sheet=rebuild.sheet,
                        cell=f"{column}{outliers[0][1][0].row}",
                        evidence=sample,
                        on_dy_path=True,
                    )
                ]

    kind = (
        "all literal export values"
        if not formulas
        else ("a uniform derived formula" if not literals else "mixed")
    )
    return [
        Finding(
            "CHK_RENT_EDITS",
            Severity.HIGH,
            Status.PASS,
            f"Rent column {column} on {rebuild.sheet!r} is {kind} - no hand edits detected.",
            sheet=rebuild.sheet,
            evidence=f"{len(literals)} literal / {len(formulas)} formula cells",
            on_dy_path=True,
        )
    ]


def check_rr_totals(rebuild: Rebuild) -> list[Finding]:
    """The export's own Total rows must equal the rows they summarize.

    A mismatch means cells changed after the export was generated. With hand
    edits found in the rent column it is conclusive and flags; without that
    corroboration it asks for review instead, since some exports total on a
    basis the sheet does not show.
    """
    if not rebuild.confident or not rebuild.totals or rebuild.sheet is None:
        return []
    column = rebuild.columns.get("rent", "?")
    broken = [t for t in rebuild.totals if not t.ties]
    if not broken:
        return [
            Finding(
                "CHK_RR_TOTAL_TIE",
                Severity.HIGH,
                Status.PASS,
                f"The export's {len(rebuild.totals)} Total row(s) on {rebuild.sheet!r} tie "
                f"to the tenant rows they summarize.",
                sheet=rebuild.sheet,
                on_dy_path=True,
            )
        ]
    sample = "; ".join(
        f"row {t.row} {t.label!r}: reported {t.reported:,.0f} vs rows above "
        f"{t.block_sum:,.0f} (cumulative {t.grand_sum:,.0f})"
        for t in broken[:5]
    )
    worst = max(broken, key=lambda t: abs(t.reported - t.grand_sum))
    corroborated = any(
        r.rent_is_formula and abs(r.monthly_rent) > 1.0 for r in rebuild.rows
    ) and any(not r.rent_is_formula and abs(r.monthly_rent) > 1.0 for r in rebuild.rows)
    if corroborated:
        return [
            Finding(
                "CHK_RR_TOTAL_TIE",
                Severity.HIGH,
                Status.FLAG,
                f"{len(broken)} Total row(s) on {rebuild.sheet!r} no longer equal the tenant "
                f"rows they summarize - the sheet was edited after export. The export's own "
                f"total says {worst.reported:,.0f}/period against {worst.grand_sum:,.0f} in "
                f"the cells, so the rent feeding GPR is not what the property manager "
                f"reported.",
                sheet=rebuild.sheet,
                cell=f"{column}{broken[0].row}",
                evidence=sample,
                on_dy_path=True,
            )
        ]
    return [
        Finding(
            "CHK_RR_TOTAL_TIE",
            Severity.HIGH,
            Status.MANUAL_REVIEW,
            f"The export's Total row(s) on {rebuild.sheet!r} differ from the sum of the "
            f"tenant rows above them, and no hand edit explains it. Confirm whether the "
            f"export totals on a basis the sheet does not show, or the rows changed.",
            sheet=rebuild.sheet,
            cell=f"{column}{broken[0].row}",
            evidence=sample,
            on_dy_path=True,
        )
    ]


def check_gpr_supplementary(ctx: LoanContext, gpr_coord: str) -> list[Finding]:
    """The GPR line is base rent only - no supplementary billing codes.

    Memorial adds the PET RENT and GAS billing subtotals into GPR alongside
    RENT. Those belong in Other Income (where the operating statement's own
    coding also puts them, which would count them twice).
    """
    contributions = _chain_contributions(ctx, gpr_coord)
    hits = [
        (label, ref, amount)
        for label, ref, amount in contributions
        if amount is not None
        and amount > 1.0
        and label
        and SUPPLEMENTARY_LABEL.search(label)
        and not _ANY_RENT.fullmatch(label.strip())
    ]
    if not hits:
        return []
    total = sum(amount for _l, _r, amount in hits)
    detail = ", ".join(f"{ref} ({label}) {amount:+,.0f}" for label, ref, amount in hits)

    # Does the T12/operating statement also code the same charge to a revenue
    # line? If so the dollars are counted twice.
    cross = ""
    tabs = ctx.facts.get("source_tabs") or {}
    t12 = tabs.get("t12")
    if t12 and ctx.wb.has_sheet(t12):
        tokens = {
            word.lower()
            for label, _r, _a in hits
            for word in re.findall(r"[A-Za-z]{3,}", label)
        }
        rows_bound, _ = ctx.wb.bounds(t12)
        for row in range(1, min(rows_bound, 400) + 1):
            for col in ("A", "B", "C"):
                text = ctx.wb.text(t12, f"{col}{row}")
                if text and any(tok in text.lower() for tok in tokens) and SUPPLEMENTARY_LABEL.search(text):
                    cross = (
                        f" The operating statement carries the same charge as its own account "
                        f"({t12}!{col}{row} = {text.strip()[:48]!r}), so these dollars are "
                        f"positioned to be counted twice."
                    )
                    break
            if cross:
                break

    return [
        Finding(
            "CHK_GPR_SUPPLEMENTARY",
            Severity.HIGH,
            Status.FLAG,
            f"Gross potential rent includes {total:,.0f} of supplementary billing that is "
            f"not base rent: {detail}. The GPR line is base rent only - supplementary "
            f"income belongs in Other Income.{cross}",
            sheet=ctx.sheet,
            cell=gpr_coord,
            evidence=detail,
            on_dy_path=True,
        )
    ]


# ---------------------------------------------------------------------------
# Entry point + finding
# ---------------------------------------------------------------------------


def _model_stop_row(ctx: LoanContext, line_coord: str, sheet: str) -> int | None:
    """Where the model's own total lands on the rent-roll sheet, when that row
    is a Total row. Rows below it belong to another schedule (Quincy's parking
    block under the retail block) and must not leak into this line's rebuild.
    """
    res = F.resolve_defining_formula(ctx.wb, ctx.sheet, line_coord)
    if res.sheet != sheet:
        return None
    row = int(re.sub(r"\D", "", res.coord) or 0)
    if not row:
        return None
    for col in range(1, 9):
        text = ctx.wb.text(sheet, f"{get_column_letter(col)}{row}")
        if text and _TOTAL_LABEL.match(text.strip()):
            return row
    return None


def _targets(ctx: LoanContext) -> list[tuple[Line, str, str, int | None]]:
    """(line, rent-roll sheet, line coordinate, stop row) tuples to rebuild.

    The GPR line always; the Base Rent line too when it draws on a different
    rent roll (Quincy's commercial schedule).
    """
    out: list[tuple[Line, str, str, int | None]] = []
    tabs = ctx.facts.get("source_tabs") or {}
    gpr = ctx.tab.cell(Line.GPR)
    if gpr and tabs.get("rent_roll"):
        sheet = tabs["rent_roll"]
        out.append((Line.GPR, sheet, gpr, _model_stop_row(ctx, gpr, sheet)))
    base = ctx.tab.cell(Line.BASE_RENT)
    if base:
        value = ctx.wb.number(ctx.sheet, base)
        if value and abs(value) > 1.0:
            res = F.resolve_defining_formula(ctx.wb, ctx.sheet, base)
            for hop in res.path[1:] + [f"{r.resolved_sheet(res.sheet)}!x" for r in F.iter_refs(res.formula)]:
                sheet = hop.rsplit("!", 1)[0]
                if (
                    ctx.wb.has_sheet(sheet)
                    and sheet != ctx.sheet
                    and sheet != tabs.get("rent_roll")
                    and re.search(r"rent\s*roll|\brr\b", sheet, re.I)
                ):
                    out.append((Line.BASE_RENT, sheet, base, _model_stop_row(ctx, base, sheet)))
                    break
    return out


def run_rebuild(ctx: LoanContext) -> list[Finding]:
    """Rebuild each rent roll, reconcile, and report. Stores results in facts."""
    rebuilds: list[Rebuild] = []
    ctx.facts["rebuilds"] = rebuilds
    findings: list[Finding] = []

    targets = _targets(ctx)
    if not targets:
        return [
            Finding(
                "CHK_RENT_REBUILD",
                Severity.HIGH,
                Status.MANUAL_REVIEW,
                "No rent-roll tab could be located from the GPR line, so the rent roll was "
                "not rebuilt. Rebuild it by hand.",
                sheet=ctx.sheet,
                on_dy_path=True,
            )
        ]

    for line, sheet, coord, stop_row in targets:
        header_row, columns, rent_header, two_row = find_tenant_table(ctx, sheet)
        if header_row is not None:
            rebuild = extract_rows(
                ctx, sheet, header_row, columns, rent_header, two_row, stop_row
            )
            rebuild.line = line
        else:
            rebuild = _rebuild_from_parse(ctx, line)
        rebuilds.append(rebuild)

        if not rebuild.confident:
            findings.append(
                Finding(
                    "CHK_RENT_REBUILD",
                    Severity.HIGH,
                    Status.MANUAL_REVIEW,
                    f"The tenant table on {sheet!r} could not be read confidently, so the "
                    f"{line.value} line was not independently rebuilt. "
                    + "; ".join(rebuild.notes),
                    sheet=sheet,
                    on_dy_path=True,
                )
            )
            continue

        reconcile(ctx, rebuild, coord)
        findings.extend(check_rent_edits(ctx, rebuild))
        findings.extend(check_rr_totals(rebuild))
        if line is Line.GPR:
            findings.extend(check_gpr_supplementary(ctx, coord))
        model = rebuild.model_gpr or 0.0
        rebuilt = rebuild.annualized()
        difference = model - rebuilt
        relative = abs(difference) / max(abs(model), 1.0)
        included = len(rebuild.included_rows())
        excluded = len(rebuild.rows) - included
        basis = "monthly x 12" if rebuild.monthly else "annual"
        evidence = (
            f"{sheet}!{rebuild.columns.get('rent')} ({rebuild.rent_header}), {basis}; "
            f"{included} included / {excluded} excluded row(s); rebuilt {rebuilt:,.2f} vs "
            f"model {model:,.2f}"
        )

        # Forensic signals flag even when the headline number ties.
        problems = [
            item
            for item, _amount in rebuild.reconciliation
            if "edited after export" in item
            or "otherwise literal" in item
            or "note says" in item
        ]

        if relative <= TOLERANCE and not problems:
            findings.append(
                Finding(
                    "CHK_RENT_REBUILD",
                    Severity.HIGH,
                    Status.PASS,
                    f"Independent rent-roll rebuild ties to the {line.value} line within "
                    f"{TOLERANCE:.1%} ({rebuilt:,.0f} vs {model:,.0f}). See the Rent Roll "
                    f"Rebuild tab for every row and decision.",
                    sheet=ctx.sheet,
                    cell=coord,
                    evidence=evidence,
                    on_dy_path=True,
                )
            )
            continue

        # Raw contractual rent can legitimately differ where the rent roll
        # documents per-tenant adjustments (Campus's free-rent months and
        # include flags) - provided the model's own adjusted column ties.
        if relative > TOLERANCE and rebuild.adjustment_headers and not problems:
            parse = ctx.facts.get("rent_roll_parse")
            adjusted_ties = (
                parse is not None
                and parse.confident
                and abs(parse.annualized() - model) / max(abs(model), 1.0) <= TOLERANCE
            )
            if adjusted_ties:
                findings.append(
                    Finding(
                        "CHK_RENT_REBUILD",
                        Severity.HIGH,
                        Status.PASS,
                        f"Raw contractual rent rebuilds to {rebuilt:,.0f} against the model's "
                        f"{model:,.0f}; the difference is the rent roll's documented "
                        f"per-tenant adjustments ({', '.join(rebuild.adjustment_headers[:3])}), "
                        f"and the model's adjusted column ties independently (see "
                        f"CHK_GPR_RECOMPUTE). Review the adjustments by hand - the tool does "
                        f"not verify them.",
                        sheet=ctx.sheet,
                        cell=coord,
                        evidence=evidence,
                        on_dy_path=True,
                    )
                )
                continue

        detail: list[str] = []
        if relative > TOLERANCE:
            detail.append(
                f"rebuilt base rent differs from the model by {difference:+,.0f} "
                f"({relative:.2%})"
            )
        detail.extend(problems)
        reconciliation_text = "; ".join(
            f"{item}: {amount:+,.0f}" if amount is not None else item
            for item, amount in rebuild.reconciliation
        )
        findings.append(
            Finding(
                "CHK_RENT_REBUILD",
                Severity.HIGH,
                Status.FLAG,
                f"{line.value}: " + "; ".join(detail) + ". Reconciliation on the Rent Roll "
                f"Rebuild tab: " + reconciliation_text,
                sheet=ctx.sheet,
                cell=coord,
                evidence=evidence,
                on_dy_path=True,
            )
        )
    return findings
