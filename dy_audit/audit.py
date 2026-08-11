"""Per-loan orchestration: run every check and collect the findings."""

from __future__ import annotations

import re
import traceback

from . import formula as F
from .checks.blockers import run_blockers
from .checks.high import run_high
from .checks.low import run_low
from .checks.medium import run_medium
from .checks.standing import run_standing
from .context import LoanContext
from .definitions import parse_definitions
from .model import Finding, LoanFiles, LoanResult, Severity, Status
from .osar import Line, select_osar
from .recompute import run_recompute
from .workbook import Workbook

#: Label vocabulary that marks a covenant level rather than a computed yield.
_THRESHOLD_LABEL = re.compile(
    r"\b(event|threshold|covenant|minimum|min|required|hurdle|trigger|target)\b", re.I
)
_DY_LABEL = re.compile(r"debt yield|(^|\b)dy\b", re.I)


def find_covenant_threshold(ctx: LoanContext) -> tuple[float, str] | None:
    """Locate the debt-yield covenant level, if the workbook states one.

    Only a labelled literal on a tab other than the OSAR counts. The OSAR's own
    columns E-G hold prior and at-contribution debt yields, which look like
    thresholds but are history - reading one as a covenant would produce a
    confident and wrong pass/fail.
    """
    from openpyxl.utils import column_index_from_string, get_column_letter

    wb, tab = ctx.wb, ctx.tab
    for sheet in wb.sheet_names:
        if sheet == tab.sheet or not wb.is_visible(sheet):
            continue
        for coord, _formula, value in wb.iter_cells(sheet, max_row=120, max_col=25):
            if not isinstance(value, str):
                continue
            label = value.strip()
            if not _DY_LABEL.search(label) or not _THRESHOLD_LABEL.search(label):
                continue
            m = re.match(r"([A-Z]+)(\d+)", coord)
            if not m:
                continue
            index = column_index_from_string(m.group(1))
            for step in range(1, 6):
                near = f"{get_column_letter(index + step)}{m.group(2)}"
                number = wb.number(sheet, near)
                if number is not None and 0 < number < 1.0 and wb.formula(sheet, near) is None:
                    return number, f"{sheet}!{near} ({label})"
    return None


def _covenant_findings(ctx: LoanContext, dy: float | None) -> list[Finding]:
    """Compare the reported debt yield to the covenant level."""
    threshold = find_covenant_threshold(ctx)
    if dy is None:
        return []
    if threshold is None:
        return [
            Finding(
                "CHK_COVENANT",
                Severity.INFO,
                Status.MANUAL_REVIEW,
                f"Reported debt yield is {dy:.4%}. No covenant threshold is stated anywhere in "
                f"this workbook, so pass/fail could not be determined - confirm the level and "
                f"its effective date from the loan documents. (The OSAR's own columns E-G hold "
                f"prior and at-contribution debt yields, not a covenant level.)",
                sheet=ctx.tab.sheet,
                cell=ctx.tab.cell(Line.DEBT_YIELD),
                on_dy_path=True,
            )
        ]

    level, where = threshold
    passing = dy >= level
    return [
        Finding(
            "CHK_COVENANT",
            Severity.INFO if passing else Severity.HIGH,
            Status.PASS if passing else Status.FLAG,
            f"Reported debt yield {dy:.4%} "
            f"{'clears' if passing else 'is below'} the {level:.4%} covenant level by "
            f"{abs(dy - level) * 100:.2f} basis points x100. Confirm the level's effective date - "
            f"a threshold not yet in force does not bind.",
            sheet=ctx.tab.sheet,
            cell=ctx.tab.cell(Line.DEBT_YIELD),
            evidence=f"threshold read from {where}",
            on_dy_path=True,
        )
    ]


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
    threshold = find_covenant_threshold(ctx)
    facts["covenant"] = threshold[0] if threshold else None
    facts["covenant_source"] = threshold[1] if threshold else None
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
        for finding in run_medium(ctx):
            result.add(finding)
        for finding in run_low(ctx):
            result.add(finding)
        for finding in run_standing(ctx):
            result.add(finding)

        result.facts = collect_facts(ctx)
        for finding in _covenant_findings(ctx, result.facts.get("debt_yield")):
            result.add(finding)
        result.facts["params"] = ctx.params
    except Exception as exc:  # noqa: BLE001 - one loan's failure must not stop the run
        result.error = f"{type(exc).__name__}: {exc}"
        result.facts["traceback"] = traceback.format_exc()
    finally:
        if wb is not None:
            wb.close()
    return result


__all__ = ["audit_loan", "collect_facts", "find_covenant_threshold", "F"]
