"""Stage 4 acceptance: MEDIUM and LOW checks, month count, standing UPB flag.

The sharpest test here is a negative one. Campus carries 130 external-link parts
and 12,024 defined names while no formula references any of them, and spec
section 6 calls it the cleanest of the four - so CHK_EXTERNAL_REFS must stay
silent on it.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from dy_audit.checks.blockers import run_blockers
from dy_audit.checks.high import run_high
from dy_audit.checks.low import run_low
from dy_audit.checks.medium import run_medium
from dy_audit.checks.standing import run_standing
from dy_audit.context import LoanContext
from dy_audit.definitions import parse_definitions
from dy_audit.discovery import discover
from dy_audit.model import Severity, Status
from dy_audit.recompute import run_recompute
from dy_audit.osar import select_osar
from dy_audit.workbook import Workbook

INPUT_DIR = Path(__file__).resolve().parent.parent
FLAG, PASS, MANUAL = Status.FLAG, Status.PASS, Status.MANUAL_REVIEW
LOANS = ["Strada", "Campus at Villa La Jolla", "Hialeah", "Ares55thAve"]


@pytest.fixture(scope="module")
def contexts():
    pairs, problems = discover(INPUT_DIR)
    assert problems == []
    opened = {}
    for pair in pairs:
        wb = Workbook(pair.xlsx_path)
        ctx = LoanContext(
            files=pair,
            wb=wb,
            tab=select_osar(wb),
            params=parse_definitions(pair.defs_path, pair.loan_name),
        )
        run_blockers(ctx)
        run_high(ctx)
        run_recompute(ctx)
        opened[pair.loan_name] = ctx
    yield opened
    for ctx in opened.values():
        ctx.wb.close()


@pytest.fixture(scope="module")
def results(contexts):
    return {n: run_medium(c) + run_low(c) + run_standing(c) for n, c in contexts.items()}


def _findings(results, loan, check_id, status=None):
    out = [f for f in results[loan] if f.check_id == check_id]
    return [f for f in out if f.status is status] if status else out


# -- CHK_EXTERNAL_REFS: the false-positive trap ------------------------------


def test_campus_external_link_parts_do_not_trigger_a_finding(contexts, results):
    # The parts exist in the file; no cell formula uses them.
    import zipfile

    with zipfile.ZipFile(contexts["Campus at Villa La Jolla"].files.xlsx_path) as z:
        parts = [n for n in z.namelist() if "externalLink" in n and n.endswith(".xml")]
    assert len(parts) > 100, "expected Campus to carry many external-link parts"
    assert _findings(results, "Campus at Villa La Jolla", "CHK_EXTERNAL_REFS", FLAG) == []


@pytest.mark.parametrize("loan", ["Strada", "Campus at Villa La Jolla", "Ares55thAve"])
def test_external_refs_silent_where_no_formula_links_out(results, loan):
    assert _findings(results, loan, "CHK_EXTERNAL_REFS")[0].status is PASS


def test_hialeah_external_links_are_reported(results):
    finding = _findings(results, "Hialeah", "CHK_EXTERNAL_REFS", FLAG)
    assert len(finding) == 1
    assert "(Old) Comm OSAR" in finding[0].evidence


# -- CHK_UNIT_SF_TIE ---------------------------------------------------------


def test_strada_unit_count_mismatch(results):
    finding = _findings(results, "Strada", "CHK_UNIT_SF_TIE", FLAG)
    assert len(finding) == 1
    assert "497" in finding[0].message and "495" in finding[0].message


def test_campus_square_foot_mismatch(results):
    finding = _findings(results, "Campus at Villa La Jolla", "CHK_UNIT_SF_TIE", FLAG)
    assert len(finding) == 1
    assert "191,544" in finding[0].message and "191,454" in finding[0].message


@pytest.mark.parametrize("loan", ["Hialeah", "Ares55thAve"])
def test_unit_count_ties_where_the_osar_links_to_the_rent_roll(results, loan):
    assert _findings(results, loan, "CHK_UNIT_SF_TIE")[0].status is PASS


# -- CHK_COLUMN_H_SOURCE -----------------------------------------------------


def test_ares_reference_column_has_one_odd_link(results):
    finding = _findings(results, "Ares55thAve", "CHK_COLUMN_H_SOURCE", FLAG)
    assert len(finding) == 1
    assert finding[0].cell == "H47"
    assert "Actuals (T12)" in finding[0].message


@pytest.mark.parametrize("loan", ["Strada", "Campus at Villa La Jolla", "Hialeah"])
def test_reference_column_consistent_elsewhere(results, loan):
    assert _findings(results, loan, "CHK_COLUMN_H_SOURCE")[0].status is PASS


# -- CHK_DUPLICATE_TABS ------------------------------------------------------


def test_hialeah_duplicate_period_and_rent_roll_tabs(results):
    flagged = _findings(results, "Hialeah", "CHK_DUPLICATE_TABS", FLAG)
    assert len(flagged) == 2
    messages = " ".join(f.message for f in flagged)
    assert "'1Q26 T12'" in messages and "'1Q RR'" in messages
    # The output does reference the live tabs, so this is a confirm-each-quarter
    # note rather than an assertion that the wrong tab is wired up.
    assert "references '1Q26 RR'" in messages and "references '1Q26 T12'" in messages


@pytest.mark.parametrize("loan", ["Strada", "Campus at Villa La Jolla", "Ares55thAve"])
def test_duplicate_tabs_silent_elsewhere(results, loan):
    # Ares has both "Actuals (T12)" and "T12" - different reports, not stale
    # copies of one tab, so grouping must not treat them as duplicates. Strada
    # and Campus each have a second OSAR tab, which the visible-tab rule already
    # handles.
    assert _findings(results, loan, "CHK_DUPLICATE_TABS")[0].status is PASS


# -- CHK_LINK_TARGETS --------------------------------------------------------


def test_hialeah_whole_column_sums_are_flagged(results):
    flagged = _findings(results, "Hialeah", "CHK_LINK_TARGETS", FLAG)
    cells = {f.cell for f in flagged}
    assert "I25" in cells and "I29" in cells


def test_hialeah_management_note_contradicts_its_formula(results):
    notes = [
        f for f in _findings(results, "Hialeah", "CHK_LINK_TARGETS", FLAG)
        if "note beside the management fee" in f.message
    ]
    assert len(notes) == 1
    assert "3% of GPR" in notes[0].message
    assert "calculation is right" in notes[0].message


@pytest.mark.parametrize("loan", ["Strada", "Campus at Villa La Jolla"])
def test_selective_whole_column_lookups_are_not_flagged(results, loan):
    # Strada uses SUMIF and INDEX/MATCH over whole columns and Campus uses
    # XLOOKUP; those select a row rather than aggregating everything, so they
    # are not fragile in the way a bare SUM(V:V) is.
    assert _findings(results, loan, "CHK_LINK_TARGETS", FLAG) == []


# -- CHK_ERROR_CELLS ---------------------------------------------------------


def test_strada_reports_errors_on_hidden_and_visible_tabs(results):
    # S-12. Hidden tabs are swept: the hidden-tab rule governs which OSAR tab is
    # audited, not whole-file hygiene.
    flagged = _findings(results, "Strada", "CHK_ERROR_CELLS", FLAG)
    sheets = {f.sheet for f in flagged}
    assert {"Comm OSAR", "Paydown Scenario", "Renovation Schedule", "Variance Analysis"} <= sheets
    assert all(f.on_dy_path is False for f in flagged), "none of these feed the debt yield"


def test_campus_single_cosmetic_error(results):
    flagged = _findings(results, "Campus at Villa La Jolla", "CHK_ERROR_CELLS", FLAG)
    assert len(flagged) == 1
    assert flagged[0].cell == "K26"


def test_ares_error_cells_include_the_annualiser_fallout(results):
    flagged = _findings(results, "Ares55thAve", "CHK_ERROR_CELLS", FLAG)
    by_sheet = {f.sheet: f for f in flagged}
    assert "99 #VALUE!" in by_sheet["T12"].message  # A-10
    assert by_sheet["Debt Service"].cell == "C32"  # A-11


# -- CHK_HARDCODE_IN_FORMULA -------------------------------------------------


def test_hialeah_occupancy_plug_is_found_on_the_rent_roll(results):
    # H-5. The plug sits two hops from the OSAR, inside the rent roll's occupancy
    # formula, and drives the vacancy branch.
    flagged = _findings(results, "Hialeah", "CHK_HARDCODE_IN_FORMULA", FLAG)
    assert len(flagged) == 1
    assert flagged[0].sheet == "1Q26 RR" and flagged[0].cell == "F62"
    assert "34,000" in flagged[0].message
    assert flagged[0].on_dy_path is True
    assert flagged[0].severity is Severity.MEDIUM


def test_ares_tax_invoice_is_typed_in_rather_than_sourced(results):
    # Not in spec section 6. 'RE Taxes'!T5 = 13748.26+89345.08 feeds the tax MAX
    # invoice term, so the MAX is structurally correct but its input is unsourced.
    flagged = _findings(results, "Ares55thAve", "CHK_HARDCODE_IN_FORMULA", FLAG)
    assert len(flagged) == 1
    assert flagged[0].cell == "T5"
    assert "typed-in numbers" in flagged[0].message
    assert flagged[0].severity is Severity.LOW


@pytest.mark.parametrize("loan", ["Strada", "Campus at Villa La Jolla"])
def test_no_plugs_on_the_clean_models(results, loan):
    assert _findings(results, loan, "CHK_HARDCODE_IN_FORMULA")[0].status is PASS


# -- CHK_ANNUALIZATION_FORMULA -----------------------------------------------


def test_ares_annualiser_divides_outside_its_guard(results):
    flagged = _findings(results, "Ares55thAve", "CHK_ANNUALIZATION_FORMULA", FLAG)
    assert len(flagged) == 1 and flagged[0].sheet == "T12"
    assert "outside the IF" in flagged[0].message


@pytest.mark.parametrize("loan", ["Strada", "Campus at Villa La Jolla", "Hialeah"])
def test_annualiser_check_silent_elsewhere(results, loan):
    assert _findings(results, loan, "CHK_ANNUALIZATION_FORMULA")[0].status is PASS


# -- CHK_MONTH_COUNT ---------------------------------------------------------


def test_ares_seven_month_annualisation_is_surfaced_not_flagged(results):
    # Spec section 6 records Ares as correctly annualising 7 months x 12/7. The
    # window is short because the history is short, which is not an omission.
    findings = _findings(results, "Ares55thAve", "CHK_MONTH_COUNT")
    assert not [f for f in findings if f.status is FLAG]
    assert any("7 with data" in f.message and "divides by $N$3" in f.message for f in findings)


def test_strada_fixed_windows_are_surfaced_not_flagged(results):
    # T6 (x2) and T3 (x4) legitimately sum fewer months than are available: a
    # window shrinks on short history but never grows past its definition.
    findings = _findings(results, "Strada", "CHK_MONTH_COUNT")
    assert not [f for f in findings if f.status is FLAG]
    message = " ".join(f.message for f in findings)
    assert "multiplies by 2" in message and "multiplies by 4" in message


@pytest.mark.parametrize("loan", LOANS)
def test_month_count_never_flags_on_these_four(results, loan):
    assert _findings(results, loan, "CHK_MONTH_COUNT", FLAG) == []


# -- CHK_UPB_CONFIRM ---------------------------------------------------------


@pytest.mark.parametrize(
    "loan,cell,amount",
    [
        ("Strada", "E6", "92,071,459.13"),
        ("Campus at Villa La Jolla", "E6", "63,000,000.00"),
        ("Hialeah", "E6", "42,768,425.00"),
        ("Ares55thAve", "D6", "10,409,000.00"),
    ],
)
def test_upb_flag_is_always_emitted_and_echoes_the_model_figure(results, loan, cell, amount):
    findings = _findings(results, loan, "CHK_UPB_CONFIRM")
    assert len(findings) == 1
    finding = findings[0]
    # Never a pass: there is no source to verify against, so it stays an action.
    assert finding.status is MANUAL
    assert finding.severity is Severity.STANDING
    assert finding.cell == cell
    assert amount in finding.message


def test_ares_upb_flag_records_the_indirection(results):
    # Ares divides by D6, which is itself a link to its debt-service tab.
    finding = _findings(results, "Ares55thAve", "CHK_UPB_CONFIRM")[0]
    assert "Debt Service!C4" in finding.message
