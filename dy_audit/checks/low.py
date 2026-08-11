"""LOW checks - whole-file hygiene (spec section 4)."""

from __future__ import annotations

import re
from collections import defaultdict

from .. import formula as F
from ..context import LoanContext
from ..model import Finding, Severity, Status
from ..osar import Line
from ..workbook import ERROR_VALUES

#: Constants that carry no judgment: unit conversions, percentages already
#: checked elsewhere, and the small integers used for indexing and rounding.
_INNOCUOUS = {0.0, 1.0, 2.0, 3.0, 4.0, 6.0, 12.0, 100.0, 365.0, 360.0, 1000.0, 1000000.0}

#: A number added into or subtracted from a formula. No lookahead: the plug the
#: spec names as its example, `(F54-F41-F7+34000)/F54`, closes a parenthesis
#: immediately after the constant, so excluding numbers followed by `)` would
#: skip exactly the case this check exists to catch.
_PLUG = re.compile(r"[+\-]\s*(\d{3,}(?:\.\d+)?)")

#: `SUM(...)` divided or multiplied outside the IF that guards it.
_ANNUALIZER = re.compile(
    r"IF\s*\(\s*SUM\([^)]*\)\s*=\s*0\s*,\s*\"\"\s*,\s*SUM\([^)]*\)\s*\)\s*[/*]", re.I
)
#: A horizontal SUM combined with an annualisation factor.
_SUM_RANGE = re.compile(r"SUM\(\s*\$?([A-Z]{1,3})\$?(\d+)\s*:\s*\$?([A-Z]{1,3})\$?(\d+)\s*\)", re.I)
_MULTIPLIER = re.compile(r"\)\s*\*\s*(\d+(?:\.\d+)?)")
_DIVISOR = re.compile(r"/\s*(\$?[A-Z]{1,3}\$?\d+|\d+(?:\.\d+)?)\s*\*\s*12")


# ---------------------------------------------------------------------------
# CHK_ERROR_CELLS
# ---------------------------------------------------------------------------


def check_error_cells(ctx: LoanContext) -> list[Finding]:
    """Report every cached error value, separating the DY path from the rest.

    Hidden sheets are swept too. The hidden-tab rule governs which OSAR tab is
    audited, not whole-file hygiene - and spec section 6 reports errors sitting
    on Strada's hidden tabs.
    """
    wb, tab = ctx.wb, ctx.tab
    on_path_sheets = {tab.sheet} | set(
        (ctx.facts.get("source_tabs") or {}).get(k) or "" for k in ("rent_roll", "t12", "ar")
    )

    groups: dict[tuple[str, str], list[str]] = defaultdict(list)
    for sheet, coord, _formula, value in wb.iter_all_cells():
        if isinstance(value, str) and value.strip() in ERROR_VALUES:
            groups[(sheet, value.strip())].append(coord)

    if not groups:
        return [
            Finding(
                "CHK_ERROR_CELLS",
                Severity.LOW,
                Status.PASS,
                "No cached error values anywhere in the workbook.",
                sheet=tab.sheet,
            )
        ]

    findings: list[Finding] = []
    for (sheet, kind), coords in sorted(groups.items()):
        on_path = sheet in on_path_sheets
        visibility = "visible" if wb.is_visible(sheet) else "hidden"
        shown = ", ".join(coords[:6]) + (f", ... (+{len(coords) - 6})" if len(coords) > 6 else "")
        findings.append(
            Finding(
                "CHK_ERROR_CELLS",
                Severity.LOW,
                Status.FLAG,
                f"{len(coords)} {kind} cell(s) on the {visibility} tab {sheet!r}"
                + (
                    " - this tab feeds the debt-yield calculation, so confirm none of them "
                    "reaches the output."
                    if on_path
                    else " - off the debt-yield path."
                ),
                sheet=sheet,
                cell=coords[0],
                evidence=shown,
                on_dy_path=on_path,
            )
        )
    return findings


# ---------------------------------------------------------------------------
# CHK_HARDCODE_IN_FORMULA
# ---------------------------------------------------------------------------


def check_hardcode_in_formula(ctx: LoanContext) -> list[Finding]:
    """Numeric plugs added inside otherwise-linked formulas.

    The search follows the OSAR lines into their driver cells on other tabs, so a
    plug sitting in the rent roll's occupancy formula is found and marked as
    being on the debt-yield path.
    """
    wb, tab = ctx.wb, ctx.tab
    findings: list[Finding] = []
    seen: set[str] = set()

    key_lines = [
        Line.GPR,
        Line.VACANCY,
        Line.OCCUPANCY,
        Line.REIMBURSEMENT,
        Line.OTHER_INCOME,
        Line.PARKING,
        Line.EGI,
        Line.TAX,
        Line.INSURANCE,
        Line.MGMT_FEE,
        Line.CAPEX,
        Line.NOI,
        Line.NCF,
    ]
    targets: dict[str, tuple[str, str]] = {}
    for line in key_lines:
        row = tab.rows.get(line)
        if row is None:
            continue
        for column in (tab.dy_column, "H"):
            coord = f"{column}{row}"
            if wb.formula(tab.sheet, coord):
                key = f"{tab.sheet}!{coord}"
                targets[key] = (tab.sheet, coord)
                targets.update(F.dependencies(wb, tab.sheet, coord, max_depth=4))
                break

    for sheet, coord in targets.values():
        key = f"{sheet}!{coord}"
        if key in seen:
            continue
        seen.add(key)
        formula = wb.formula(sheet, coord)
        if not formula:
            continue
        body = F.strip_strings(F.normalize(formula))
        constants = [
            float(m.group(1)) for m in _PLUG.finditer(body) if float(m.group(1)) not in _INNOCUOUS
        ]
        if not constants:
            continue
        links = [r for r in F.iter_refs(body)]

        if links:
            # A plug: a manual adjustment mixed into an otherwise-linked formula.
            findings.append(
                Finding(
                    "CHK_HARDCODE_IN_FORMULA",
                    Severity.MEDIUM if sheet != tab.sheet else Severity.LOW,
                    Status.FLAG,
                    f"A constant of {constants[0]:,.0f} is added inside an otherwise-linked "
                    f"formula. An undocumented manual adjustment on the debt-yield path should "
                    f"be sourced or moved to its own labelled input cell.",
                    sheet=sheet,
                    cell=coord,
                    evidence=f"{sheet}!{coord} = {body}",
                    on_dy_path=True,
                )
            )
        else:
            # A cell built purely from typed-in numbers. Not a plug - the
            # calculation around it may be perfectly correct - but the figure
            # itself has no source, so it cannot be tied to a document.
            findings.append(
                Finding(
                    "CHK_HARDCODE_IN_FORMULA",
                    Severity.LOW,
                    Status.FLAG,
                    f"{sheet}!{coord} is built entirely from typed-in numbers ({body}) and feeds "
                    f"the debt-yield calculation. The arithmetic may be right, but the figure "
                    f"cannot be tied back to a source document - confirm it against the bill or "
                    f"invoice.",
                    sheet=sheet,
                    cell=coord,
                    evidence=f"{sheet}!{coord} = {body}",
                    on_dy_path=True,
                )
            )

    if not findings:
        return [
            Finding(
                "CHK_HARDCODE_IN_FORMULA",
                Severity.LOW,
                Status.PASS,
                "No numeric plugs found inside the formulas feeding the debt yield.",
                sheet=tab.sheet,
            )
        ]
    return findings


# ---------------------------------------------------------------------------
# CHK_ANNUALIZATION_FORMULA
# ---------------------------------------------------------------------------


def check_annualization_formula(ctx: LoanContext) -> list[Finding]:
    """`IF(SUM(...)=0,"",SUM(...))/N*12` divides the empty string it just returned."""
    wb, tab = ctx.wb, ctx.tab
    groups: dict[str, list[str]] = defaultdict(list)
    for sheet, coord, formula, _value in wb.iter_all_cells():
        if formula and _ANNUALIZER.search(F.normalize(formula)):
            groups[sheet].append(coord)

    if not groups:
        return [
            Finding(
                "CHK_ANNUALIZATION_FORMULA",
                Severity.LOW,
                Status.PASS,
                "Short-history annualisers keep their division inside the guarding IF.",
                sheet=tab.sheet,
            )
        ]

    findings: list[Finding] = []
    for sheet, coords in sorted(groups.items()):
        errors = [c for c in coords if wb.is_error(sheet, c)]
        example = wb.formula(sheet, coords[0])
        findings.append(
            Finding(
                "CHK_ANNUALIZATION_FORMULA",
                Severity.LOW,
                Status.FLAG,
                f"{len(coords)} annualiser(s) on {sheet!r} divide outside the IF that guards "
                f"them, so an empty row returns \"\" and is then divided, yielding #VALUE! "
                f"({len(errors)} currently in error). Move the division inside: "
                f"=IF(SUM(...)=0,\"\",SUM(...)/$N$3*12).",
                sheet=sheet,
                cell=coords[0],
                evidence=f"{sheet}!{coords[0]} = {F.normalize(example)}",
                on_dy_path=False,
            )
        )
    return findings


# ---------------------------------------------------------------------------
# CHK_MONTH_COUNT
# ---------------------------------------------------------------------------


def check_month_count(ctx: LoanContext) -> list[Finding]:
    """Surface the month count behind every annualised line, and flag omissions.

    A fixed window (a T3 line multiplied by 4) legitimately sums fewer months
    than are available - the window shrinks on short history but never grows.
    What is wrong is an annualiser whose divisor disagrees with the months it
    actually sums.
    """
    wb, tab = ctx.wb, ctx.tab
    tabs = ctx.facts.get("source_tabs") or {}
    # Scan every trailing-period tab, not only the one column H links to: Ares
    # keeps its monthly detail on "T12" while column H reads the summary on
    # "Actuals (T12)", and the annualisers live on the detail tab.
    t12_like = re.compile(r"t-?12|ttm|trailing|operating statement", re.I)
    sheets: list[str] = []
    for candidate in (tabs.get("t12"), tabs.get("rent_roll"), *wb.sheet_names):
        if candidate and wb.has_sheet(candidate) and candidate not in sheets:
            if candidate in (tabs.get("t12"), tabs.get("rent_roll")) or t12_like.search(candidate):
                sheets.append(candidate)
    findings: list[Finding] = []
    reported: list[str] = []

    for sheet in sheets:
        seen_shapes: set[str] = set()
        for coord, formula, _value in wb.iter_cells(sheet, max_row=200, max_col=40):
            if not formula:
                continue
            body = F.normalize(formula)
            span = _SUM_RANGE.search(body)
            if not span:
                continue
            multiplier = _MULTIPLIER.search(body)
            divisor = _DIVISOR.search(body)
            if not multiplier and not divisor:
                continue

            start_col, row, end_col, _end_row = span.group(1), int(span.group(2)), span.group(3), span.group(4)
            shape = f"{start_col}:{end_col}|{bool(divisor)}"
            if shape in seen_shapes:
                continue
            seen_shapes.add(shape)

            from openpyxl.utils import column_index_from_string, get_column_letter

            c1, c2 = column_index_from_string(start_col), column_index_from_string(end_col)
            width = abs(c2 - c1) + 1
            with_data = sum(
                1
                for c in range(min(c1, c2), max(c1, c2) + 1)
                if wb.number(sheet, f"{get_column_letter(c)}{row}") not in (None, 0)
            )

            if divisor:
                token = divisor.group(1)
                implied = (
                    wb.number(sheet, token.replace("$", ""))
                    if re.match(r"\$?[A-Z]{1,3}\$?\d+$", token)
                    else float(token)
                )
                label = f"divides by {token}"
            else:
                implied = 12.0 / float(multiplier.group(1)) if float(multiplier.group(1)) else None
                label = f"multiplies by {multiplier.group(1)}"

            reported.append(
                f"{sheet}!{coord}: sums {width} column(s) {start_col}:{end_col}, {with_data} with "
                f"data, {label} (implying {implied:g} month(s))"
                if implied
                else f"{sheet}!{coord}: sums {width} column(s)"
            )

            if divisor and implied and abs(with_data - implied) > 0.5:
                findings.append(
                    Finding(
                        "CHK_MONTH_COUNT",
                        Severity.HIGH,
                        Status.FLAG,
                        f"An annualised line on {sheet!r} sums {with_data} month(s) of data but "
                        f"annualises as though there were {implied:g}. Confirm which months "
                        f"belong in the window - the tool does not reclassify them.",
                        sheet=sheet,
                        cell=coord,
                        evidence=f"{sheet}!{coord} = {body[:110]}",
                        on_dy_path=True,
                    )
                )

    if reported:
        findings.append(
            Finding(
                "CHK_MONTH_COUNT",
                Severity.INFO,
                Status.PASS,
                "Month counts used by the annualised lines: " + "; ".join(reported[:6]) + ".",
                sheet=sheets[0] if sheets else tab.sheet,
                on_dy_path=True,
            )
        )
    elif not findings:
        findings.append(
            Finding(
                "CHK_MONTH_COUNT",
                Severity.INFO,
                Status.MANUAL_REVIEW,
                "No annualisation formulas were found on the supporting tabs, so the month "
                "counts behind the trailing figures could not be surfaced.",
                sheet=tab.sheet,
            )
        )
    return findings


LOW_CHECKS = (
    check_error_cells,
    check_hardcode_in_formula,
    check_annualization_formula,
    check_month_count,
)


def run_low(ctx: LoanContext) -> list[Finding]:
    findings: list[Finding] = []
    for check in LOW_CHECKS:
        findings.extend(check(ctx))
    return findings
