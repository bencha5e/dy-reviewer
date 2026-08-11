"""MEDIUM checks - linkage and robustness (spec section 4)."""

from __future__ import annotations

import re
from collections import Counter, defaultdict

from .. import formula as F
from ..context import LoanContext
from ..model import Finding, Severity, Status
from ..osar import Line

#: Period tokens that distinguish one quarter's copy of a tab from another's.
#: Stripping them groups "1Q RR" with "1Q26 RR" but keeps "Actuals (T12)"
#: separate from "T12" - those are different reports, not stale duplicates.
_PERIOD_TOKEN = re.compile(r"\b\d{0,2}Q\d{0,4}\b|\b(?:19|20)\d{2}\b|\((?:new|old)\)", re.I)

_ROLE_PATTERNS = {
    "rent roll": re.compile(r"rent\s*roll|(^|\s)rr($|\s)", re.I),
    "operating statement / T12": re.compile(r"t-?12|ttm|trailing|operating statement", re.I),
}


def _base_name(sheet: str) -> str:
    return re.sub(r"\s+", " ", _PERIOD_TOKEN.sub("", sheet)).strip().lower()


# ---------------------------------------------------------------------------
# CHK_EXTERNAL_REFS
# ---------------------------------------------------------------------------


def check_external_refs(ctx: LoanContext) -> list[Finding]:
    """Links into another workbook leave only cached values behind.

    Detected from `[n]` prefixes in **cell formulas**, never from the workbook's
    external-link parts. Campus carries 130 such parts and 12,024 defined names
    as legacy residue while no formula references any of them; reporting those
    would bury a clean model in noise.
    """
    wb, tab = ctx.wb, ctx.tab
    hits: list[tuple[str, str, str]] = []
    for sheet, coord, formula, _value in wb.iter_all_cells():
        if formula and F.has_external_refs(formula):
            hits.append((sheet, coord, F.normalize(formula)))

    if not hits:
        return [
            Finding(
                "CHK_EXTERNAL_REFS",
                Severity.MEDIUM,
                Status.PASS,
                "No formula links to another workbook.",
                sheet=tab.sheet,
            )
        ]

    on_path = [h for h in hits if h[0] == tab.sheet]
    sample = "; ".join(f"{s}!{c} = {f[:44]}" for s, c, f in hits[:3])
    return [
        Finding(
            "CHK_EXTERNAL_REFS",
            Severity.MEDIUM,
            Status.FLAG,
            f"{len(hits)} formula(s) link to an external workbook, "
            f"{len(on_path)} of them on the OSAR output tab. External links keep only cached "
            f"values once the other file is unavailable, so these will silently go stale.",
            sheet=hits[0][0],
            cell=hits[0][1],
            evidence=sample,
            on_dy_path=bool(on_path),
        )
    ]


# ---------------------------------------------------------------------------
# CHK_LINK_TARGETS
# ---------------------------------------------------------------------------

#: An unconditional aggregation over an entire column. Criteria and lookup forms
#: (SUMIF, INDEX/MATCH, XLOOKUP) also span whole columns but select a row, so
#: they are not fragile in the same way and must not be flagged.
_BARE_AGGREGATE = re.compile(r"\b(SUM|SUMPRODUCT|AVERAGE|COUNT)\s*\(", re.I)
_SELECTIVE = re.compile(r"\b(SUMIF|SUMIFS|COUNTIF|COUNTIFS|INDEX|MATCH|XLOOKUP|VLOOKUP|LOOKUP)\b", re.I)


def check_link_targets(ctx: LoanContext) -> list[Finding]:
    """Fragile links: whole-column aggregates, single-cell anchors, stale notes."""
    wb, tab = ctx.wb, ctx.tab
    findings: list[Finding] = []

    income_rows = [
        row
        for line, row in tab.rows.items()
        if line
        in (
            Line.GPR,
            Line.VACANCY,
            Line.REIMBURSEMENT,
            Line.PARKING,
            Line.OTHER_INCOME,
            Line.TAX,
            Line.INSURANCE,
            Line.MGMT_FEE,
            Line.CAPEX,
        )
    ]

    for row in sorted(income_rows):
        coord = f"{tab.dy_column}{row}"
        res = F.resolve_defining_formula(wb, tab.sheet, coord)
        if res.formula is None:
            continue
        body = F.normalize(res.formula)
        label = tab.raw_labels.get(row, "").strip()

        whole_columns = [r for r in F.iter_refs(body) if r.is_whole_column]
        if whole_columns and _BARE_AGGREGATE.search(body) and not _SELECTIVE.search(body):
            findings.append(
                Finding(
                    "CHK_LINK_TARGETS",
                    Severity.LOW,
                    Status.FLAG,
                    f"{label!r} aggregates an entire column "
                    f"({', '.join(str(r) for r in whole_columns)}). A whole-column sum picks up "
                    f"any stray value added to that column later; anchor it to the tenant rows.",
                    sheet=tab.sheet,
                    cell=coord,
                    evidence=f"{res.ref} = {body}",
                    on_dy_path=True,
                )
            )

        # A revenue line anchored to one cell silently ignores tenants added below it.
        if tab.rows.get(Line.GPR) == row:
            refs = [r for r in F.iter_refs(body) if not r.is_range]
            if len(refs) == 1 and not F.iter_refs(body)[0].is_range and res.hopped:
                findings.append(
                    Finding(
                        "CHK_LINK_TARGETS",
                        Severity.LOW,
                        Status.FLAG,
                        f"Gross potential rent resolves to the single cell {res.ref} rather than a "
                        f"tenant range. Any tenant added to the rent roll outside that cell will "
                        f"not reach the OSAR.",
                        sheet=tab.sheet,
                        cell=coord,
                        evidence=f"{tab.sheet}!{coord} -> {' -> '.join(res.path)}",
                        on_dy_path=True,
                    )
                )

    findings.extend(_check_note_matches_formula(ctx))
    if not findings:
        findings.append(
            Finding(
                "CHK_LINK_TARGETS",
                Severity.MEDIUM,
                Status.PASS,
                "Column links resolve to specific ranges and on-tab notes match their formulas.",
                sheet=tab.sheet,
            )
        )
    return findings


#: An on-tab note naming the management-fee base.
_NOTE_BASE = re.compile(r"3(?:\.0+)?\s*%\s*(?:of\s+)?(GPR|gross potential rent|EGI|gross revenue)", re.I)


def _check_note_matches_formula(ctx: LoanContext) -> list[Finding]:
    """A note describing a different calculation than the formula performs.

    Hialeah's management fee is correctly 3% x EGI while the note beside it reads
    "3% of GPR". The calculation is right; the note invites someone to "correct"
    it to a figure that is wrong.
    """
    wb, tab = ctx.wb, ctx.tab
    row = tab.rows.get(Line.MGMT_FEE)
    egi_row = tab.rows.get(Line.EGI)
    if row is None or egi_row is None:
        return []
    coord = f"{tab.dy_column}{row}"
    body = F.normalize(F.resolve_defining_formula(wb, tab.sheet, coord).formula)
    if not body:
        return []
    uses_egi = f"{tab.dy_column}{egi_row}" in body.upper().replace("$", "")

    for column in ("J", "K", "L", "M", "N"):
        note = wb.text(tab.sheet, f"{column}{row}")
        if not note:
            continue
        m = _NOTE_BASE.search(note)
        if not m:
            continue
        stated = m.group(1).upper()
        if uses_egi and stated not in ("EGI",):
            return [
                Finding(
                    "CHK_LINK_TARGETS",
                    Severity.LOW,
                    Status.FLAG,
                    f"The note beside the management fee says {note.strip()!r}, but the formula "
                    f"correctly applies 3% to EGI. The calculation is right - the note is "
                    f"misleading and invites a wrong correction.",
                    sheet=tab.sheet,
                    cell=f"{column}{row}",
                    evidence=f"{tab.sheet}!{coord} = {body}",
                    on_dy_path=False,
                )
            ]
    return []


# ---------------------------------------------------------------------------
# CHK_COLUMN_H_SOURCE
# ---------------------------------------------------------------------------


def check_column_h_source(ctx: LoanContext) -> list[Finding]:
    """The reference (T12) column should pull every line from one actuals tab."""
    wb, tab = ctx.wb, ctx.tab
    start = tab.rows.get(Line.GPR)
    end = tab.rows.get(Line.TOTAL_OPEX)
    if start is None or end is None:
        return []

    sources: dict[int, str] = {}
    for row in range(start, end + 1):
        formula = wb.formula(tab.sheet, f"H{row}")
        if not formula:
            continue
        refs = [r for r in F.iter_refs(formula) if r.sheet]
        if refs:
            sources[row] = refs[0].sheet

    if len(set(sources.values())) <= 1:
        return [
            Finding(
                "CHK_COLUMN_H_SOURCE",
                Severity.MEDIUM,
                Status.PASS,
                f"The reference column links consistently to "
                f"{next(iter(set(sources.values())), 'a single tab')!r}.",
                sheet=tab.sheet,
            )
        ]

    dominant, _count = Counter(sources.values()).most_common(1)[0]
    odd = {row: sheet for row, sheet in sources.items() if sheet != dominant}
    detail = "; ".join(
        f"H{row} -> {sheet!r} ({tab.raw_labels.get(row, '').strip()})" for row, sheet in odd.items()
    )
    return [
        Finding(
            "CHK_COLUMN_H_SOURCE",
            Severity.MEDIUM,
            Status.FLAG,
            f"{len(odd)} row(s) in the reference column pull from a different tab than the rest, "
            f"which link to {dominant!r}. The value may agree today, but the odd link will not "
            f"follow when the actuals tab is refreshed.",
            sheet=tab.sheet,
            cell=f"H{min(odd)}",
            evidence=detail,
            on_dy_path=False,
        )
    ]


# ---------------------------------------------------------------------------
# CHK_UNIT_SF_TIE
# ---------------------------------------------------------------------------


def check_unit_sf_tie(ctx: LoanContext) -> list[Finding]:
    """The unit or square-foot count must tie between rent roll and OSAR.

    The rent roll's own total is taken from the denominator of the occupancy
    formula, which is the figure the model itself treats as the property total.
    """
    wb, tab = ctx.wb, ctx.tab
    count_row = tab.rows.get(Line.NRSF)
    occ_row = tab.rows.get(Line.OCCUPANCY)
    if count_row is None or occ_row is None:
        return []

    osar_count = None
    count_cell = None
    for column in ("E", "D", tab.dy_column):
        value = wb.number(tab.sheet, f"{column}{count_row}")
        if value:
            osar_count, count_cell = value, f"{column}{count_row}"
            break
    if osar_count is None:
        return []

    res = F.resolve_defining_formula(wb, tab.sheet, f"{tab.dy_column}{occ_row}")
    if res.formula is None:
        for column in ("H",):
            candidate = F.resolve_defining_formula(wb, tab.sheet, f"{column}{occ_row}")
            if candidate.formula:
                res = candidate
                break
    body = F.normalize(res.formula)
    if "/" not in body:
        return []

    denominator = body.rsplit("/", 1)[1]
    refs = [r for r in F.iter_refs(denominator) if not r.is_range]
    if not refs:
        return []
    ref = refs[0]
    sheet = ref.resolved_sheet(res.sheet)
    rr_total = wb.number(sheet, ref.coord)
    if rr_total is None:
        return []

    evidence = (
        f"{tab.sheet}!{count_cell} = {osar_count:,.0f}; rent-roll total {sheet}!{ref.coord} = "
        f"{rr_total:,.0f} (occupancy denominator in {res.ref} = {body})"
    )
    if abs(rr_total - osar_count) < 0.5:
        return [
            Finding(
                "CHK_UNIT_SF_TIE",
                Severity.MEDIUM,
                Status.PASS,
                f"Unit/SF count ties between the rent roll and the OSAR ({osar_count:,.0f}).",
                sheet=tab.sheet,
                cell=count_cell,
                evidence=evidence,
                on_dy_path=True,
            )
        ]

    difference = rr_total - osar_count
    reserve = wb.number(tab.sheet, f"{tab.dy_column}{tab.rows.get(Line.CAPEX, 0)}")
    impact = ""
    if reserve and osar_count:
        impact = f" The reserve is sized on the OSAR figure, so it is off by about {abs(difference) * reserve / osar_count:,.0f}."
    return [
        Finding(
            "CHK_UNIT_SF_TIE",
            Severity.MEDIUM,
            Status.FLAG,
            f"The rent roll totals {rr_total:,.0f} but the OSAR carries {osar_count:,.0f} "
            f"({difference:+,.0f}). Occupancy is computed on the rent-roll figure while capital "
            f"items are sized on the OSAR figure, so the two must be reconciled.{impact}",
            sheet=tab.sheet,
            cell=count_cell,
            evidence=evidence,
            on_dy_path=True,
        )
    ]


# ---------------------------------------------------------------------------
# CHK_DUPLICATE_TABS
# ---------------------------------------------------------------------------


def check_duplicate_tabs(ctx: LoanContext) -> list[Finding]:
    """Several copies of the rent roll or period tab invite referencing a stale one."""
    wb, tab = ctx.wb, ctx.tab
    groups: dict[tuple[str, str], list[str]] = defaultdict(list)
    for sheet in wb.sheet_names:
        for role, pattern in _ROLE_PATTERNS.items():
            if pattern.search(sheet):
                groups[(role, _base_name(sheet))].append(sheet)
                break

    referenced = _referenced_sheets(ctx)
    findings: list[Finding] = []
    for (role, _base), sheets in sorted(groups.items()):
        if len(sheets) < 2:
            continue
        live = [s for s in sheets if s in referenced]
        stale = [s for s in sheets if s not in referenced]
        hidden = [s for s in sheets if not wb.is_visible(s)]
        errors = [
            s for s in stale if any(wb.is_error(s, c) for c, _f, _v in wb.iter_cells(s, max_row=200))
        ]
        findings.append(
            Finding(
                "CHK_DUPLICATE_TABS",
                Severity.MEDIUM,
                Status.FLAG,
                f"{len(sheets)} {role} tabs are present ({', '.join(repr(s) for s in sheets)}). "
                f"The output references {', '.join(repr(s) for s in live) or 'none of them'}"
                f"{'; ' + ', '.join(repr(s) for s in errors) + ' still contain error cells' if errors else ''}. "
                f"Confirm the live tab again each quarter - the correct one changes as periods roll.",
                sheet=live[0] if live else sheets[0],
                evidence=f"hidden: {hidden or 'none'}; unreferenced: {stale or 'none'}",
                on_dy_path=True,
            )
        )

    if not findings:
        return [
            Finding(
                "CHK_DUPLICATE_TABS",
                Severity.MEDIUM,
                Status.PASS,
                "No duplicate rent-roll or period tabs.",
                sheet=tab.sheet,
            )
        ]
    return findings


def _referenced_sheets(ctx: LoanContext) -> set[str]:
    """Sheets the OSAR output tab links to, directly or one hop away."""
    wb, tab = ctx.wb, ctx.tab
    seen: set[str] = set()
    for coord, formula, _value in wb.iter_cells(tab.sheet, max_row=120, max_col=20):
        if not formula:
            continue
        for ref in F.iter_refs(formula):
            sheet = ref.resolved_sheet(tab.sheet)
            if sheet != tab.sheet and wb.has_sheet(sheet):
                seen.add(sheet)
                for inner_coord, inner_formula, _v in wb.iter_cells(sheet, max_row=80, max_col=20):
                    if inner_formula:
                        for inner_ref in F.iter_refs(inner_formula):
                            inner_sheet = inner_ref.resolved_sheet(sheet)
                            if wb.has_sheet(inner_sheet):
                                seen.add(inner_sheet)
                    del inner_coord
    return seen


MEDIUM_CHECKS = (
    check_external_refs,
    check_link_targets,
    check_column_h_source,
    check_unit_sf_tie,
    check_duplicate_tabs,
)


def run_medium(ctx: LoanContext) -> list[Finding]:
    findings: list[Finding] = []
    for check in MEDIUM_CHECKS:
        findings.extend(check(ctx))
    return findings
