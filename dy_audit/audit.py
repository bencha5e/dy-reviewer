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
from .llm import REVENUE_CHECK_IDS, review_revenue_with_llm
from .model import Finding, LoanFiles, LoanResult
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


def _keep(finding: Finding, use_llm: bool) -> bool:
    """Whether a deterministic finding survives to the report.

    With the LLM review on, the revenue checks still run - they build the
    rent-roll rebuild the report renders and the source tabs later checks read -
    but their verdicts give way to the model's, so a loan is never reported
    twice for the same defect under the same check ID.
    """
    return not (use_llm and finding.check_id in REVENUE_CHECK_IDS)


def audit_loan(
    files: LoanFiles,
    *,
    use_llm: bool = False,
    llm_provider: str = "anthropic",
    llm_model: str | None = None,
    llm_effort: str = "xhigh",
) -> LoanResult:
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
            use_llm=use_llm,
            llm_provider=llm_provider,
            llm_model=llm_model,
            llm_effort=llm_effort,
        )
        # Order matters: blockers record the UPB reference and the period check
        # records the source tabs, both of which later checks reuse.
        stages = (run_blockers, run_high, run_recompute, run_rebuild, run_medium,
                  run_low, run_standing)
        # The rebuild reuses the recompute's parse as its fallback, so it runs after.
        for stage in stages:
            for finding in stage(ctx):
                if _keep(finding, use_llm):
                    result.add(finding)

        # Revenue last: it reads the rebuild and source tabs the stages above
        # produced, and reports against them.
        if use_llm:
            for finding in review_revenue_with_llm(
                ctx,
                provider=llm_provider,
                model=llm_model,
                effort=llm_effort,
            ):
                result.add(finding)

        result.facts = collect_facts(ctx)
        result.facts["params"] = ctx.params
        result.facts["rebuilds"] = ctx.facts.get("rebuilds") or []
        result.facts["llm_usage"] = ctx.facts.get("llm_usage")
        result.facts["revenue_review"] = ctx.facts.get("revenue_review")
    except Exception as exc:  # noqa: BLE001 - one loan's failure must not stop the run
        result.error = f"{type(exc).__name__}: {exc}"
        result.facts["traceback"] = traceback.format_exc()
    finally:
        if wb is not None:
            wb.close()
    return result


__all__ = ["audit_loan", "collect_facts"]
