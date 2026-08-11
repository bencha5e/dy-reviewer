"""Per-loan orchestration: run every check and collect the findings."""

from __future__ import annotations

import traceback

from .checks.blockers import run_blockers
from .checks.high import run_high
from .checks.low import run_low
from .checks.medium import run_medium
from .checks.standing import run_standing
from .context import LoanContext
from .definitions import parse_definitions
from .model import LoanFiles, LoanResult
from .osar import Line, select_osar
from .rebuild import run_rebuild
from .recompute import run_recompute
from .workbook import Workbook


def collect_facts(ctx: LoanContext) -> dict:
    """Headline figures for the summary sheet."""
    wb, tab = ctx.wb, ctx.tab
    dy = wb.number(tab.sheet, tab.cell(Line.DEBT_YIELD)) if tab.has(Line.DEBT_YIELD) else None
    upb_ref = ctx.facts.get("upb_ref")
    facts = {
        "loan_name": ctx.loan_name,
        "workbook": ctx.files.xlsx_path.name,
        "definitions": ctx.files.defs_path.name,
        "osar_tab": tab.sheet,
        "dy_column": tab.dy_column,
        "period_end": tab.period_end,
        "debt_yield": dy,
        "noi": wb.number(tab.sheet, tab.cell(Line.NOI)) if tab.has(Line.NOI) else None,
        "ncf": wb.number(tab.sheet, tab.cell(Line.NCF)) if tab.has(Line.NCF) else None,
        "egi": wb.number(tab.sheet, tab.cell(Line.EGI)) if tab.has(Line.EGI) else None,
        "upb": wb.number(*upb_ref) if upb_ref else None,
        "upb_cell": f"{upb_ref[0]}!{upb_ref[1]}" if upb_ref else None,
        "occupancy": None,
        "hidden_osar_tabs": tab.hidden_osar_tabs,
        "source_tabs": ctx.facts.get("source_tabs", {}),
    }
    occ_row = tab.rows.get(Line.OCCUPANCY)
    if occ_row:
        for column in (tab.dy_column, "H"):
            value = wb.number(tab.sheet, f"{column}{occ_row}")
            if value is not None:
                facts["occupancy"] = value
                break
    return facts


def audit_loan(files: LoanFiles) -> LoanResult:
    """Run the full check suite for one loan.

    Wrapped so a workbook that cannot be read is recorded as this loan's failure
    rather than ending the run for every other loan.
    """
    result = LoanResult(loan_name=files.loan_name, files=files)
    wb = None
    try:
        wb = Workbook(files.xlsx_path)
        ctx = LoanContext(
            files=files,
            wb=wb,
            tab=select_osar(wb),
            params=parse_definitions(files.defs_path, files.loan_name),
        )
        # Order matters: blockers record the UPB reference and the period check
        # records the source tabs, both of which later checks reuse.
        for finding in run_blockers(ctx):
            result.add(finding)
        for finding in run_high(ctx):
            result.add(finding)
        for finding in run_recompute(ctx):
            result.add(finding)
        # The rebuild reuses the recompute's parse as its fallback, so it runs after.
        for finding in run_rebuild(ctx):
            result.add(finding)
        for finding in run_medium(ctx):
            result.add(finding)
        for finding in run_low(ctx):
            result.add(finding)
        for finding in run_standing(ctx):
            result.add(finding)

        result.facts = collect_facts(ctx)
        result.facts["params"] = ctx.params
        result.facts["rebuilds"] = ctx.facts.get("rebuilds") or []
    except Exception as exc:  # noqa: BLE001 - one loan's failure must not stop the run
        result.error = f"{type(exc).__name__}: {exc}"
        result.facts["traceback"] = traceback.format_exc()
    finally:
        if wb is not None:
            wb.close()
    return result


__all__ = ["audit_loan", "collect_facts"]
