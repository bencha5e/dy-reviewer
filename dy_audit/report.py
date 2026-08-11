"""Per-loan findings workbook (spec section 0).

One `.xlsx` per loan, matching the source material so it can be read alongside
the model. A summary sheet up top, then every check with its severity, status,
cell reference and evidence, then the parameters read out of the loan agreement.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

from openpyxl import Workbook as XlsxWorkbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from .model import SEVERITY_ORDER, LoanResult, Severity, Status

_HEADER_FILL = PatternFill("solid", fgColor="1F3864")
_HEADER_FONT = Font(bold=True, color="FFFFFF", size=11)
_TITLE_FONT = Font(bold=True, size=14)
_LABEL_FONT = Font(bold=True)
_WRAP = Alignment(vertical="top", wrap_text=True)
_TOP = Alignment(vertical="top")

#: Severity colours, loudest first.
_SEVERITY_FILL = {
    Severity.BLOCKER: PatternFill("solid", fgColor="C00000"),
    Severity.HIGH: PatternFill("solid", fgColor="E8A33D"),
    Severity.MEDIUM: PatternFill("solid", fgColor="FFD966"),
    Severity.LOW: PatternFill("solid", fgColor="D9E2F3"),
    Severity.STANDING: PatternFill("solid", fgColor="C6E0B4"),
    Severity.INFO: PatternFill("solid", fgColor="F2F2F2"),
}
_SEVERITY_FONT = {
    Severity.BLOCKER: Font(bold=True, color="FFFFFF"),
    Severity.HIGH: Font(bold=True),
}

#: Status colours. MANUAL_REVIEW must never look like a pass.
_STATUS_FILL = {
    Status.FLAG: PatternFill("solid", fgColor="F4B6B6"),
    Status.PASS: PatternFill("solid", fgColor="D7E8D4"),
    Status.MANUAL_REVIEW: PatternFill("solid", fgColor="FFE699"),
    Status.UNVERIFIABLE: PatternFill("solid", fgColor="D0CECE"),
}

_THIN = Side(style="thin", color="BFBFBF")
_BORDER = Border(left=_THIN, right=_THIN, top=_THIN, bottom=_THIN)

_FINDING_COLUMNS = [
    ("Severity", 11),
    ("Status", 16),
    ("Check", 28),
    ("Sheet", 22),
    ("Cell", 9),
    ("On DY path", 11),
    ("Finding", 82),
    ("Evidence", 74),
]


def _sort_key(finding):
    # Flags first within each severity, so the reader meets the actions first.
    status_rank = {Status.FLAG: 0, Status.MANUAL_REVIEW: 1, Status.UNVERIFIABLE: 2, Status.PASS: 3}
    return (
        SEVERITY_ORDER.get(finding.severity, 9),
        status_rank.get(finding.status, 9),
        finding.check_id,
    )


def _write_row(sheet, row: int, label: str, value, number_format: str | None = None) -> int:
    sheet.cell(row=row, column=1, value=label).font = _LABEL_FONT
    cell = sheet.cell(row=row, column=2, value=value)
    if number_format:
        cell.number_format = number_format
    cell.alignment = _TOP
    return row + 1


def _summary_sheet(sheet, result: LoanResult, run_date: dt.date) -> None:
    facts = result.facts
    sheet.cell(row=1, column=1, value=f"Debt Yield Audit - {result.loan_name}").font = _TITLE_FONT

    row = 3
    row = _write_row(sheet, row, "Run date", run_date.strftime("%Y-%m-%d"))
    row = _write_row(sheet, row, "Workbook", facts.get("workbook"))
    row = _write_row(sheet, row, "Definitions", facts.get("definitions"))
    row = _write_row(sheet, row, "OSAR tab audited", facts.get("osar_tab"))
    row = _write_row(sheet, row, "DY-test column", facts.get("dy_column"))
    period = facts.get("period_end")
    row = _write_row(sheet, row, "Statement ending", period.strftime("%m/%d/%Y") if period else "not read")
    hidden = facts.get("hidden_osar_tabs") or []
    row = _write_row(
        sheet, row, "Hidden OSAR tabs (ignored)", ", ".join(hidden) if hidden else "none"
    )

    row += 1
    sheet.cell(row=row, column=1, value="Reported figures").font = _HEADER_FONT
    sheet.cell(row=row, column=1).fill = _HEADER_FILL
    sheet.cell(row=row, column=2).fill = _HEADER_FILL
    row += 1
    row = _write_row(sheet, row, "Debt yield (NCF / UPB)", facts.get("debt_yield"), "0.0000%")
    covenant = facts.get("covenant")
    if covenant is None:
        row = _write_row(
            sheet, row, "Covenant threshold", "not stated in this workbook - confirm from loan docs"
        )
    else:
        row = _write_row(sheet, row, "Covenant threshold", covenant, "0.0000%")
        dy = facts.get("debt_yield")
        verdict = "n/a" if dy is None else ("PASS" if dy >= covenant else "FAIL")
        row = _write_row(sheet, row, "Covenant result", verdict)
        row = _write_row(sheet, row, "Threshold source", facts.get("covenant_source"))
    row = _write_row(sheet, row, "Net cash flow (NCF)", facts.get("ncf"), "#,##0.00")
    row = _write_row(sheet, row, "Net operating income", facts.get("noi"), "#,##0.00")
    row = _write_row(sheet, row, "Effective gross income", facts.get("egi"), "#,##0.00")
    row = _write_row(sheet, row, "Occupancy", facts.get("occupancy"), "0.00%")
    row = _write_row(sheet, row, "UPB used", facts.get("upb"), "#,##0.00")
    row = _write_row(sheet, row, "UPB cell", facts.get("upb_cell"))
    sheet.cell(row=row, column=1, value="UPB is echoed, not verified").font = _LABEL_FONT
    sheet.cell(
        row=row,
        column=2,
        value="There is no source file for unpaid principal balance - confirm against internal records.",
    ).alignment = _TOP
    row += 2

    sheet.cell(row=row, column=1, value="Findings by severity").font = _HEADER_FONT
    sheet.cell(row=row, column=1).fill = _HEADER_FILL
    sheet.cell(row=row, column=2).fill = _HEADER_FILL
    row += 1
    counts = result.count_by_severity()
    for severity in sorted(Severity, key=lambda s: SEVERITY_ORDER.get(s, 9)):
        count = counts.get(severity, 0)
        sheet.cell(row=row, column=1, value=severity.value).font = _LABEL_FONT
        if count:
            sheet.cell(row=row, column=1).fill = _SEVERITY_FILL[severity]
            font = _SEVERITY_FONT.get(severity)
            if font:
                sheet.cell(row=row, column=1).font = font
        sheet.cell(row=row, column=2, value=count)
        row += 1

    row += 1
    sources = facts.get("source_tabs") or {}
    sheet.cell(row=row, column=1, value="Supporting tabs used").font = _HEADER_FONT
    sheet.cell(row=row, column=1).fill = _HEADER_FILL
    sheet.cell(row=row, column=2).fill = _HEADER_FILL
    row += 1
    for role, label in (("rent_roll", "Rent roll"), ("t12", "T12 / operating statement"), ("ar", "AR aging")):
        row = _write_row(sheet, row, label, sources.get(role) or "not present in this workbook")

    if result.error:
        row += 1
        sheet.cell(row=row, column=1, value="PROCESSING ERROR").font = Font(bold=True, color="C00000")
        sheet.cell(row=row, column=2, value=result.error).alignment = _WRAP

    sheet.column_dimensions["A"].width = 30
    sheet.column_dimensions["B"].width = 82
    sheet.freeze_panes = "A3"


def _findings_sheet(sheet, result: LoanResult) -> None:
    for index, (title, width) in enumerate(_FINDING_COLUMNS, start=1):
        cell = sheet.cell(row=1, column=index, value=title)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        cell.alignment = _TOP
        sheet.column_dimensions[get_column_letter(index)].width = width

    for row_index, finding in enumerate(sorted(result.findings, key=_sort_key), start=2):
        values = [
            finding.severity.value,
            finding.status.value,
            finding.check_id,
            finding.sheet or "",
            finding.cell or "",
            "" if finding.on_dy_path is None else ("yes" if finding.on_dy_path else "no"),
            finding.message,
            finding.evidence or "",
        ]
        for column_index, value in enumerate(values, start=1):
            cell = sheet.cell(row=row_index, column=column_index, value=value)
            cell.alignment = _WRAP if column_index >= 7 else _TOP
            cell.border = _BORDER
        severity_cell = sheet.cell(row=row_index, column=1)
        severity_cell.fill = _SEVERITY_FILL[finding.severity]
        font = _SEVERITY_FONT.get(finding.severity)
        if font:
            severity_cell.font = font
        sheet.cell(row=row_index, column=2).fill = _STATUS_FILL[finding.status]

    sheet.freeze_panes = "A2"
    if result.findings:
        sheet.auto_filter.ref = f"A1:{get_column_letter(len(_FINDING_COLUMNS))}{len(result.findings) + 1}"


def _parameters_sheet(sheet, result: LoanResult) -> None:
    params = result.facts.get("params")
    headers = ["Parameter", "Value", "Quoted from the loan agreement"]
    widths = [30, 26, 108]
    for index, (title, width) in enumerate(zip(headers, widths), start=1):
        cell = sheet.cell(row=1, column=index, value=title)
        cell.font = _HEADER_FONT
        cell.fill = _HEADER_FILL
        sheet.column_dimensions[get_column_letter(index)].width = width

    if params is None:
        sheet.cell(row=2, column=1, value="Definitions file was not parsed for this loan.")
        return

    rows = [
        ("Vacancy floor", params.vacancy_floor),
        ("Occupancy threshold", None),
        ("Investment-grade carve-out", params.ig_vacancy_carveout),
        ("Reserve rate", params.reserve_rate),
        ("Reserve basis", params.reserve_basis),
        ("Management fee %", params.mgmt_fee_pct),
        ("Base stated in agreement", params.mgmt_stated_base),
        ("Base actually enforced", None),
        ("Delinquency window", params.delinquency),
        ("New-lease occupancy window", params.new_lease_days),
        ("New-lease window, investment grade", params.new_lease_ig_days),
        ("Rent-step window", params.rent_step_window),
        ("Known-vacate window", params.known_vacate_days),
    ]
    row_index = 2
    for label, parsed in rows:
        sheet.cell(row=row_index, column=1, value=label).font = _LABEL_FONT
        if label == "Occupancy threshold":
            threshold = params.occupancy_threshold
            sheet.cell(row=row_index, column=2, value=f"{threshold:.2%}" if threshold else "n/a")
            sheet.cell(row=row_index, column=3, value="derived as 1 - vacancy floor").alignment = _WRAP
        elif label == "Base actually enforced":
            sheet.cell(row=row_index, column=2, value=params.mgmt_base)
            sheet.cell(
                row=row_index,
                column=3,
                value="House rule: the 3% base is always EGI regardless of the agreement's wording.",
            ).alignment = _WRAP
        else:
            sheet.cell(row=row_index, column=2, value=str(parsed))
            sheet.cell(row=row_index, column=3, value=parsed.quote or "").alignment = _WRAP
        row_index += 1

    if params.conflicts:
        row_index += 1
        sheet.cell(row=row_index, column=1, value="Conflicts with the build spec").font = _HEADER_FONT
        sheet.cell(row=row_index, column=1).fill = _HEADER_FILL
        row_index += 1
        for conflict in params.conflicts:
            sheet.cell(row=row_index, column=1, value="Loan agreement governs").font = _LABEL_FONT
            sheet.cell(row=row_index, column=2, value=conflict).alignment = _WRAP
            row_index += 1
    sheet.freeze_panes = "A2"


def write_findings_workbook(result: LoanResult, destination: Path, run_date: dt.date | None = None) -> Path:
    """Write one loan's findings workbook and return the path written."""
    run_date = run_date or dt.date.today()
    workbook = XlsxWorkbook()

    summary = workbook.active
    summary.title = "Summary"
    _summary_sheet(summary, result, run_date)
    _findings_sheet(workbook.create_sheet("Findings"), result)
    _parameters_sheet(workbook.create_sheet("Loan Parameters"), result)

    destination = Path(destination)
    destination.parent.mkdir(parents=True, exist_ok=True)
    workbook.save(destination)
    return destination


def findings_filename(loan_name: str, run_date: dt.date) -> str:
    safe = "".join(c if c.isalnum() or c in " -_" else "_" for c in loan_name).strip()
    return f"{safe} - DY Audit Findings - {run_date:%Y_%m_%d}.xlsx"
