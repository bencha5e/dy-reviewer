"""Standing flags - emitted on every run regardless of what the model says."""

from __future__ import annotations

from .. import formula as F
from ..context import LoanContext
from ..model import Finding, Severity, Status
from ..osar import Line


def check_upb_confirm(ctx: LoanContext) -> list[Finding]:
    """Always ask for the UPB to be confirmed against internal records.

    There is no source file for unpaid principal balance, so the tool cannot
    verify it and must not imply that it has. It echoes the figure the model
    divided by, together with the cell it came from - which differs per loan:
    three of the four use E6, Ares uses D6, itself a link to its debt-service tab.
    """
    wb, tab = ctx.wb, ctx.tab
    ref = ctx.facts.get("upb_ref")

    if ref is None:
        dy_coord = tab.cell(Line.DEBT_YIELD)
        body = F.normalize(wb.formula(tab.sheet, dy_coord)) if dy_coord else ""
        refs = [r for r in F.iter_refs(body) if not r.is_range]
        if len(refs) >= 2:
            ref = (refs[1].resolved_sheet(tab.sheet), refs[1].coord)

    if ref is None:
        return [
            Finding(
                "CHK_UPB_CONFIRM",
                Severity.STANDING,
                Status.MANUAL_REVIEW,
                "Confirm the unpaid principal balance against internal records. The tool could "
                "not identify which cell the debt yield divided by, so no figure is echoed here.",
                sheet=tab.sheet,
                on_dy_path=True,
            )
        ]

    sheet, coord = ref
    value = wb.number(sheet, coord)
    source = F.resolve_defining_formula(wb, sheet, coord)
    trail = ""
    if source.hopped:
        trail = f" (the cell links through {' -> '.join(source.path[1:])})"

    amount = f"{value:,.2f}" if value is not None else "not readable"
    return [
        Finding(
            "CHK_UPB_CONFIRM",
            Severity.STANDING,
            Status.MANUAL_REVIEW,
            f"Confirm the unpaid principal balance against internal records. The debt yield was "
            f"calculated on {amount} taken from {sheet}!{coord}{trail}. There is no source file "
            f"for UPB, so this figure is echoed rather than verified.",
            sheet=sheet,
            cell=coord,
            evidence=f"{sheet}!{coord} = {amount}",
            on_dy_path=True,
        )
    ]


STANDING_CHECKS = (check_upb_confirm,)


def run_standing(ctx: LoanContext) -> list[Finding]:
    findings: list[Finding] = []
    for check in STANDING_CHECKS:
        findings.extend(check(ctx))
    return findings
