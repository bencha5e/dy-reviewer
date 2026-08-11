"""Stages 5-7: findings workbook, run log, folder move, and the CLI end to end.

Everything here runs against copies in tmp_path. The move logic deletes from the
input folder, so no test may point at the repository's own files.

Every CLI invocation passes --no-llm. These tests assert on the deterministic
output, and they must run with no API key and no network; the revenue review is
covered separately in test_llm.py against a stubbed client.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import openpyxl
import pytest

from dy_audit.audit import audit_loan
from dy_audit.cli import run
from dy_audit.context import LoanContext
from dy_audit.definitions import parse_definitions
from dy_audit.discovery import discover
from dy_audit.filemove import finalise_loan, next_version_folder, safe_folder_name
from dy_audit.model import Severity, Status
from dy_audit.osar import select_osar
from dy_audit.report import findings_filename, write_findings_workbook
from dy_audit.runlog import build_summary, write_summary
from dy_audit.workbook import Workbook

RUN_DATE = dt.date(2026, 8, 11)
LOANS = {"Strada", "Campus at Villa La Jolla", "Hialeah", "Ares55thAve", "Lydian"}


@pytest.fixture(scope="module")
def audited():
    """Audit all four loans once, reading in place (nothing is written or moved)."""
    from tests.conftest import REPO_ROOT

    pairs, problems = discover(REPO_ROOT)
    assert problems == []
    return {p.loan_name: audit_loan(p) for p in pairs}


# -- audit orchestration ------------------------------------------------------


def test_every_loan_audits_without_error(audited):
    assert set(audited) == LOANS
    for name, result in audited.items():
        assert result.succeeded, f"{name} failed: {result.error}"
        assert result.findings


def test_headline_facts_are_collected(audited):
    expected_dy = {
        "Strada": 0.064657506908786189,
        "Campus at Villa La Jolla": 0.060705295896825412,
        "Hialeah": 0.021506099536749368,
        "Ares55thAve": 0.07088346216872761,
        "Lydian": 0.06370338683206106,
    }
    for name, dy in expected_dy.items():
        facts = audited[name].facts
        assert facts["debt_yield"] == pytest.approx(dy)
        assert facts["ncf"] is not None and facts["upb"] is not None
        # DY must reconcile to the NCF and UPB reported beside it.
        assert facts["ncf"] / facts["upb"] == pytest.approx(dy, rel=1e-9)


def test_a_broken_workbook_fails_only_its_own_loan(input_copy):
    (input_copy / "3. Hialeah Industrial Park 1Q26 DY Test_vF vBCS.xlsx").write_text("not an xlsx")
    pairs, _problems = discover(input_copy)
    results = {p.loan_name: audit_loan(p) for p in pairs}
    assert not results["Hialeah"].succeeded
    assert "traceback" in results["Hialeah"].facts
    for other in ("Strada", "Campus at Villa La Jolla", "Ares55thAve"):
        assert results[other].succeeded


# -- findings workbook --------------------------------------------------------


@pytest.fixture
def workbook_for(audited, tmp_path):
    def build(loan: str):
        path = tmp_path / findings_filename(loan, RUN_DATE)
        write_findings_workbook(audited[loan], path, RUN_DATE)
        return openpyxl.load_workbook(path), path

    return build


def test_findings_workbook_has_the_five_sheets(workbook_for):
    wb, path = workbook_for("Hialeah")
    assert wb.sheetnames == [
        "Summary",
        "Findings",
        "Rent Roll Rebuild",
        "Revenue Review",
        "Loan Parameters",
    ]
    assert path.exists() and path.stat().st_size > 0


def test_revenue_review_sheet_says_so_when_no_model_review_ran(workbook_for):
    # These fixtures are audited with --no-llm, so the sheet exists but has no
    # payload behind it. It must say which of the two reasons applies rather
    # than sitting blank, which would read as a clean revenue section.
    wb, _ = workbook_for("Hialeah")
    text = wb["Revenue Review"].cell(row=1, column=1).value
    assert "--no-llm" in text and "CHK_REVENUE_LLM" in text


def test_findings_sheet_lists_every_finding_blockers_first(audited, workbook_for):
    wb, _ = workbook_for("Strada")
    sheet = wb["Findings"]
    assert sheet.max_row - 1 == len(audited["Strada"].findings)
    assert sheet.auto_filter.ref
    severities = [sheet.cell(row=r, column=1).value for r in range(2, sheet.max_row + 1)]
    assert severities[0] == "BLOCKER"
    # Flags come before passes inside a severity band.
    statuses = [sheet.cell(row=r, column=2).value for r in range(2, sheet.max_row + 1)]
    assert statuses[0] == "FLAG"


def test_manual_review_is_visually_distinct_from_pass(workbook_for):
    # A check the tool could not complete must never look like a clean pass.
    wb, _ = workbook_for("Ares55thAve")
    sheet = wb["Findings"]
    fills = {}
    for row in range(2, sheet.max_row + 1):
        status = sheet.cell(row=row, column=2).value
        fills.setdefault(status, sheet.cell(row=row, column=2).fill.fgColor.rgb)
    assert fills["MANUAL_REVIEW"] != fills["PASS"]
    assert fills["FLAG"] != fills["PASS"]


def test_summary_reports_the_dy_and_upb(workbook_for):
    wb, _ = workbook_for("Ares55thAve")
    labels = {}
    sheet = wb["Summary"]
    for row in range(1, sheet.max_row + 1):
        key = sheet.cell(row=row, column=1).value
        if key:
            labels[key] = sheet.cell(row=row, column=2).value
    assert labels["Debt yield (NCF / UPB)"] == pytest.approx(0.07088346216872761)
    # No covenant comparison: whether the loan clears its threshold is decided in
    # a separate workflow. This tool only establishes that the DY is correct.
    assert "Covenant threshold" not in labels
    assert "Covenant result" not in labels
    assert labels["UPB used"] == pytest.approx(10_409_000)
    assert labels["UPB cell"] == "Comm OSAR!D6"
    assert "confirm against internal records" in labels["UPB is echoed, not verified"]
    assert labels["AR aging"] == "not present in this workbook"


def test_summary_names_the_hidden_osar_tab_that_was_ignored(workbook_for):
    wb, _ = workbook_for("Campus at Villa La Jolla")
    sheet = wb["Summary"]
    values = {
        sheet.cell(row=r, column=1).value: sheet.cell(row=r, column=2).value
        for r in range(1, sheet.max_row + 1)
    }
    assert values["Hidden OSAR tabs (ignored)"] == "OSAR"
    assert values["OSAR tab audited"] == "Comm OSAR"


def test_parameters_sheet_quotes_the_loan_agreement(workbook_for):
    wb, _ = workbook_for("Strada")
    sheet = wb["Loan Parameters"]
    rows = {
        sheet.cell(row=r, column=1).value: (
            sheet.cell(row=r, column=2).value,
            sheet.cell(row=r, column=3).value,
        )
        for r in range(2, sheet.max_row + 1)
    }
    assert rows["Vacancy floor"][0] == "0.05"
    assert "vacancy" in (rows["Vacancy floor"][1] or "").lower()
    assert rows["Base actually enforced"][0] == "EGI"
    assert "always EGI" in rows["Base actually enforced"][1]
    assert rows["Delinquency window"][0] == "CURRENT"


# -- run log ------------------------------------------------------------------


def test_run_log_records_totals_blockers_and_failures(audited, tmp_path):
    results = list(audited.values())
    results[0].error = "BadZipFile: File is not a zip file"
    text = build_summary(
        results, ["Unpaired model: stray.xlsx"], dt.datetime(2026, 8, 11, 9, 0), tmp_path, tmp_path, True
    )
    assert "# Debt Yield Audit" in text
    assert "4 succeeded, 1 failed" in text
    assert "left in the input queue" in text
    assert "BadZipFile" in text
    assert "Unpaired model: stray.xlsx" in text
    assert "CHK_INS_MAX" in text  # blocker detail is spelled out
    results[0].error = None


def test_a_deterministic_run_reports_no_review_cost(audited, tmp_path):
    # No review ran, so there is nothing to price. An empty cost table would
    # read as a review that was free rather than one that never happened.
    text = build_summary(
        list(audited.values()), [], dt.datetime(2026, 8, 11, 9, 0), tmp_path, tmp_path, False
    )
    assert "Revenue review cost" not in text


def test_review_cost_is_reported_per_loan_and_totalled(audited, tmp_path):
    results = list(audited.values())
    # The shape of a real run: the first loan writes the rules prefix and reads
    # nothing back, every loan after it reads that prefix from cache.
    results[0].facts["llm_usage"] = {
        "input_tokens": 4_000,
        "output_tokens": 9_000,
        "cache_read_input_tokens": 0,
        "cache_creation_input_tokens": 6_000,
    }
    results[1].facts["llm_usage"] = {
        "input_tokens": 4_000,
        "output_tokens": 11_000,
        "cache_read_input_tokens": 6_000,
        "cache_creation_input_tokens": 0,
    }
    text = build_summary(
        results, [], dt.datetime(2026, 8, 11, 9, 0), tmp_path, tmp_path, False
    )
    assert "Revenue review cost" in text
    # The cache writer reads nothing, and says so rather than reporting a rate
    # off cache traffic alone, which would have been 0/0.
    assert "| 0.0% |" in text
    assert "| 60.0% |" in text          # 6,000 read of 10,000 prompt tokens
    assert "| **Run total** |" in text
    assert "| 20,000 |" in text          # output, summed across both loans
    assert "| 30.0% |" in text           # 6,000 read of 20,000 prompt tokens
    # Loans the review never reached are simply absent, not zero-filled: the
    # other three loans carry no usage and must not appear as rows of nothing.
    section = text.split("## Revenue review cost")[1].split("##")[0]
    assert sum(1 for line in section.splitlines() if line.startswith("| ")) == 4
    assert results[2].loan_name not in section
    for result in results:
        result.facts.pop("llm_usage", None)


def test_write_summary_creates_the_log_at_the_output_root(audited, tmp_path):
    path = write_summary(list(audited.values()), [], tmp_path, dt.datetime(2026, 8, 11, 9, 0))
    assert path.parent == tmp_path and path.suffix == ".md"
    assert "Loans reviewed" in path.read_text(encoding="utf-8")


# -- folder naming and the move ----------------------------------------------


def test_version_folder_increments_and_never_overwrites(tmp_path):
    first = next_version_folder(tmp_path, "Strada", RUN_DATE)
    assert first.name == "Strada - 2026_08_11 - v1"
    first.mkdir()
    (first / "marker.txt").write_text("keep me")
    second = next_version_folder(tmp_path, "Strada", RUN_DATE)
    assert second.name == "Strada - 2026_08_11 - v2"
    assert (first / "marker.txt").read_text() == "keep me"


def test_folder_name_strips_characters_windows_rejects():
    assert safe_folder_name('A/B:C*D?"') == "A_B_C_D__"
    assert safe_folder_name("   ") == "Unnamed Loan"


def test_finalise_moves_sources_only_when_asked(audited, input_copy, tmp_path):
    pairs, _ = discover(input_copy)
    pair = next(p for p in pairs if p.loan_name == "Strada")
    result = audit_loan(pair)
    findings = tmp_path / "out" / findings_filename("Strada", RUN_DATE)
    write_findings_workbook(result, findings, RUN_DATE)

    folder = finalise_loan(result, tmp_path / "out", findings, RUN_DATE, move_sources=False)
    assert (folder / findings.name).exists()
    assert pair.xlsx_path.exists(), "sources must stay put unless move is requested"

    folder2 = finalise_loan(result, tmp_path / "out", folder / findings.name, RUN_DATE, move_sources=True)
    assert not pair.xlsx_path.exists() and not pair.defs_path.exists()
    assert (folder2 / pair.xlsx_path.name).exists()
    assert (folder2 / pair.defs_path.name).exists()


def test_failed_loan_keeps_its_sources_in_the_queue(input_copy, tmp_path):
    (input_copy / "3. Hialeah Industrial Park 1Q26 DY Test_vF vBCS.xlsx").write_text("broken")
    pairs, _ = discover(input_copy)
    pair = next(p for p in pairs if p.loan_name == "Hialeah")
    result = audit_loan(pair)
    assert not result.succeeded

    findings = tmp_path / "out" / findings_filename("Hialeah", RUN_DATE)
    write_findings_workbook(result, findings, RUN_DATE)
    finalise_loan(result, tmp_path / "out", findings, RUN_DATE, move_sources=result.succeeded)
    assert pair.xlsx_path.exists() and pair.defs_path.exists()


# -- CLI end to end -----------------------------------------------------------


def test_cli_moves_the_queue_by_default(input_copy):
    # Moving is the default: the input folder is a queue of files not yet
    # reviewed, so a clean run should leave nothing behind but the output.
    assert run(["--input-dir", str(input_copy), "--no-llm", "--quiet"]) == 0
    leftover = [p.name for p in input_copy.iterdir() if p.name != "DY Review Output"]
    assert leftover == []
    for folder in (input_copy / "DY Review Output").iterdir():
        if folder.is_dir():
            assert len(list(folder.glob("*.xlsx"))) == 2  # model plus findings
            assert len(list(folder.glob("*DYDefinitions.md"))) == 1


def test_cli_no_move_leaves_the_input_folder_alone(input_copy):
    before = sorted(p.name for p in input_copy.iterdir())
    assert run(["--input-dir", str(input_copy), "--no-move", "--no-llm", "--quiet"]) == 0
    after = sorted(p.name for p in input_copy.iterdir() if p.name != "DY Review Output")
    assert after == before

    output = input_copy / "DY Review Output"
    folders = sorted(p.name for p in output.iterdir() if p.is_dir())
    assert len(folders) == 5
    assert all("v1" in name for name in folders)
    assert len(list(output.glob("*.md"))) == 1
    for folder in output.iterdir():
        if folder.is_dir():
            assert list(folder.glob("*DY Audit Findings*.xlsx"))


def test_cli_move_drains_the_queue_and_keeps_failures(input_copy, tmp_path):
    (input_copy / "3. Hialeah Industrial Park 1Q26 DY Test_vF vBCS.xlsx").write_text("broken")
    output = tmp_path / "Review Output"
    exit_code = run(
        ["--input-dir", str(input_copy), "--output-dir", str(output), "--move", "--no-llm", "--quiet"]
    )
    assert exit_code == 1, "a failed loan must be reported through the exit code"

    remaining = sorted(p.name for p in input_copy.iterdir())
    assert remaining == [
        "3. Hialeah Industrial Park 1Q26 DY Test_vF vBCS.xlsx",
        "3. Hialeah_DYDefinitions.md",
    ]
    for loan in ("Strada", "Campus at Villa La Jolla", "Ares55thAve"):
        folder = next(output.glob(f"{loan} - *- v1"))
        assert len(list(folder.glob("*.xlsx"))) == 2  # model plus findings
        assert len(list(folder.glob("*DYDefinitions.md"))) == 1
    log = next(output.glob("*.md")).read_text(encoding="utf-8")
    assert "Hialeah" in log and "left in the input queue" in log
    assert "moved out of the input queue: yes" in log


def test_cli_second_run_same_day_creates_v2(input_copy, tmp_path):
    output = tmp_path / "Review Output"
    for _ in range(2):
        run([
            "--input-dir", str(input_copy),
            "--output-dir", str(output),
            "--no-move", "--no-llm", "--quiet", "--loan", "Strada",
        ])
    folders = sorted(p.name for p in output.iterdir() if p.is_dir())
    assert len(folders) == 2
    assert folders[0].endswith("v1") and folders[1].endswith("v2")


def test_cli_reports_when_the_folder_holds_nothing(tmp_path):
    empty = tmp_path / "empty"
    empty.mkdir()
    assert run(["--input-dir", str(empty), "--no-llm", "--quiet"]) == 1
    assert list((empty / "DY Review Output").glob("*.md"))


def test_cli_rejects_a_missing_input_folder(tmp_path):
    assert run(["--input-dir", str(tmp_path / "nope"), "--no-llm", "--quiet"]) == 2
