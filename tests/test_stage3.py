"""Stage 3 acceptance: loan parameters, HIGH checks, and the two recomputes.

Four expectations here deliberately diverge from spec section 6, each one a
ruling made during the build. They are marked DIVERGENCE and explained inline.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from dy_audit.checks.blockers import run_blockers
from dy_audit.checks.high import nearby_label, run_high, source_tabs
from dy_audit.context import LoanContext
from dy_audit.definitions import parse_definitions
from dy_audit.discovery import discover
from dy_audit.model import Severity, Status
from dy_audit.osar import Line, select_osar
from dy_audit.params import CURRENT
from dy_audit.recompute import run_recompute
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
        run_blockers(ctx)  # records the UPB reference the HIGH checks reuse
        opened[pair.loan_name] = ctx
    yield opened
    for ctx in opened.values():
        ctx.wb.close()


@pytest.fixture(scope="module")
def results(contexts):
    return {name: run_high(ctx) + run_recompute(ctx) for name, ctx in contexts.items()}


def _findings(results, loan, check_id):
    return [f for f in results[loan] if f.check_id == check_id]


def _status(results, loan, check_id) -> Status:
    found = _findings(results, loan, check_id)
    assert found, f"{loan}: {check_id} produced no finding"
    if any(f.status is FLAG for f in found):
        return FLAG
    if any(f.status is MANUAL for f in found):
        return MANUAL
    return found[0].status


# -- loan parameters come from the agreement, not from a default -------------


@pytest.mark.parametrize(
    "loan,floor,rate,basis,delinquency",
    [
        ("Strada", 0.05, 250.0, "unit", CURRENT),
        ("Campus at Villa La Jolla", 0.05, 0.25, "sf", 45),
        ("Hialeah", 0.05, 0.25, "sf", 60),
        ("Ares55thAve", 0.05, 0.10, "sf", 60),
    ],
)
def test_parameters_are_read_from_the_loan_agreement(contexts, loan, floor, rate, basis, delinquency):
    params = contexts[loan].params
    assert params.vacancy_floor.value == pytest.approx(floor)
    assert params.reserve_rate.value == pytest.approx(rate)
    assert params.reserve_basis.value == basis
    assert params.delinquency.value == delinquency
    assert params.mgmt_fee_pct.value == pytest.approx(0.03)
    assert params.missing() == []
    # Every parsed value must quote the text it came from.
    assert params.vacancy_floor.quote


def test_occupancy_threshold_derives_from_the_parsed_floor(contexts):
    for loan in LOANS:
        assert contexts[loan].params.occupancy_threshold == pytest.approx(0.95)


def test_only_campus_carves_investment_grade_out_of_the_vacancy_factor(contexts):
    # The carve-out sits inside a clause containing "(5.0%)", so a pattern that
    # stops at any period never reaches it.
    assert contexts["Campus at Villa La Jolla"].params.ig_vacancy_carveout.value is True
    for loan in ("Strada", "Hialeah", "Ares55thAve"):
        assert contexts[loan].params.ig_vacancy_carveout.value is False


def test_management_fee_base_is_egi_despite_the_agreements_wording(contexts):
    for loan in LOANS:
        params = contexts[loan].params
        assert params.mgmt_base == "EGI"
        assert "gross" in params.mgmt_stated_base.value.lower()


def test_hialeah_rent_step_window_follows_the_agreement_not_the_spec(contexts):
    # DIVERGENCE: the build spec generalises rent steps to "the next 12 months",
    # which is Campus's language. Hialeah's agreement says 90 days.
    assert contexts["Hialeah"].params.rent_step_window.value == "90 days"
    assert contexts["Campus at Villa La Jolla"].params.rent_step_window.value == "12 months"


def test_strada_delinquency_conflict_is_reported_not_resolved(results):
    conflicts = _findings(results, "Strada", "CHK_DEFINITIONS_CONFLICT")
    assert len(conflicts) == 1
    assert "current on their rental obligations" in conflicts[0].message
    assert "loan agreement governs" in conflicts[0].message
    for loan in ("Campus at Villa La Jolla", "Hialeah", "Ares55thAve"):
        assert _findings(results, loan, "CHK_DEFINITIONS_CONFLICT") == []


# -- HIGH checks --------------------------------------------------------------


@pytest.mark.parametrize("loan", LOANS)
def test_vacancy_floor_in_the_model_matches_the_agreement(results, loan):
    assert _status(results, loan, "CHK_VACANCY_FLOOR") is PASS


@pytest.mark.parametrize("loan", LOANS)
def test_reserve_rate_matches_the_agreement(results, loan):
    assert _status(results, loan, "CHK_RESERVE_RATE") is PASS


def test_ares_reserve_rate_passes_within_rounding_tolerance(contexts, results):
    # Ares stores 0.09996 for a $0.10/sf reserve. An exact comparison would
    # report a false positive on a rate the spec marks correct.
    assert contexts["Ares55thAve"].wb.number("Comm OSAR", "D14") == pytest.approx(0.09996)
    finding = _findings(results, "Ares55thAve", "CHK_RESERVE_RATE")[0]
    assert finding.status is PASS
    assert "rounding tolerance" in finding.message


@pytest.mark.parametrize("loan", LOANS)
def test_management_fee_is_tested_against_egi(results, loan):
    assert _status(results, loan, "CHK_MGMT_BASE") is PASS


def test_nearby_label_prefers_the_adjacent_column(contexts):
    # Strada's operating statement labels R54 from Q54 ("Delinquent Rents") while
    # B54 holds an unrelated "UTILITIES" heading on the same row.
    ctx = contexts["Strada"]
    assert nearby_label(ctx, "Operating Statement", "R54") == "Delinquent Rents"


# -- period ------------------------------------------------------------------


def test_period_checks_pass_on_every_available_source_tab(results):
    for loan in LOANS:
        for finding in _findings(results, loan, "CHK_PERIOD"):
            # Only the two workbooks without an aging tab may fall short of PASS.
            assert finding.status in (PASS, MANUAL)
            if finding.status is MANUAL:
                assert "No AR or aging tab" in finding.message


def test_strada_aging_report_run_date_is_not_read_as_a_stale_period(results):
    # The tab is stamped "FISCAL PERIOD 032026 AS OF 05/15/2026" - March data
    # exported in May. Reading the export date as the reporting period would be
    # a false positive on a loan the spec marks correct.
    ar = [f for f in _findings(results, "Strada", "CHK_PERIOD") if f.sheet == "AR"]
    assert len(ar) == 1 and ar[0].status is PASS


def test_campus_aging_db_caption_is_not_read_as_a_period(contexts, results):
    # The aging report leads with "DB Caption: Live 09/11/2023".
    assert "Live 09/11/2023" in contexts["Campus at Villa La Jolla"].wb.text("Aging Report", "A2")
    aging = [
        f for f in _findings(results, "Campus at Villa La Jolla", "CHK_PERIOD")
        if f.sheet == "Aging Report"
    ]
    assert len(aging) == 1 and aging[0].status is PASS


def test_rent_roll_tab_is_found_even_when_gpr_is_a_single_cell_sum(contexts):
    # Ares's GPR is SUM(RR!J8); the pass-through walker follows it to a literal,
    # so the tab has to come from the resolution path rather than a formula.
    assert source_tabs(contexts["Ares55thAve"])["rent_roll"] == "RR"


# -- exclusions ---------------------------------------------------------------


@pytest.mark.parametrize("loan", ["Hialeah", "Ares55thAve"])
def test_missing_aging_tab_cannot_pass(contexts, results, loan):
    # DIVERGENCE: neither workbook has an AR or aging tab, so the source of truth
    # for delinquency is absent. That must never read as a clean pass.
    wb = contexts[loan].wb
    assert not any(
        s.lower() == "ar" or "aging" in s.lower() for s in wb.sheet_names
    ), f"{loan} unexpectedly has an aging tab"
    statuses = {f.status for f in _findings(results, loan, "CHK_EXCLUSIONS")}
    assert MANUAL in statuses
    assert PASS not in statuses


@pytest.mark.parametrize("loan", ["Strada", "Campus at Villa La Jolla"])
def test_present_aging_tab_supports_the_delinquency_test(results, loan):
    assert any(
        f.status is PASS and "aging tab" in f.message
        for f in _findings(results, loan, "CHK_EXCLUSIONS")
    )


def test_strada_adds_delinquent_rent_to_gpr_instead_of_deducting_it(results):
    # DIVERGENCE (not in spec section 6): 'Operating Statement'!R54 is added to
    # GPR. It is 0 today, so the reported DY is right; a balance would inflate it.
    flagged = [f for f in _findings(results, "Strada", "CHK_EXCLUSIONS") if f.status is FLAG]
    assert len(flagged) == 1
    assert "adds a delinquency term" in flagged[0].message
    assert flagged[0].severity is Severity.HIGH
    for loan in ("Campus at Villa La Jolla", "Hialeah"):
        assert not [
            f for f in _findings(results, loan, "CHK_EXCLUSIONS")
            if f.status is FLAG and "delinquency term" in f.message
        ]


def test_ares_recoveries_escape_the_vacancy_factor(results):
    # A-7, reported at MEDIUM to match the spec's severity for this item.
    flagged = [
        f for f in _findings(results, "Ares55thAve", "CHK_EXCLUSIONS")
        if f.status is FLAG and "recoveries" in f.message
    ]
    assert len(flagged) == 1
    assert flagged[0].severity is Severity.MEDIUM
    assert "187,398" in flagged[0].message
    for loan in ("Strada", "Campus at Villa La Jolla", "Hialeah"):
        assert not [
            f for f in _findings(results, loan, "CHK_EXCLUSIONS")
            if f.status is FLAG and "recoveries" in f.message
        ]


# -- the Hialeah double count -------------------------------------------------


def test_hialeah_applies_the_vacancy_floor_twice(contexts, results):
    # DIVERGENCE (not in spec section 6): '1Q26 RR'!V7 = (Q7*12)*0.95 applies the
    # 5% floor inside GPR while actual vacancy is already 55%.
    ctx = contexts["Hialeah"]
    assert "*0.95" in ctx.wb.formula("1Q26 RR", "V7").replace(" ", "")
    flagged = _findings(results, "Hialeah", "CHK_VACANCY_DOUBLE_COUNT")
    assert len(flagged) == 1 and flagged[0].status is FLAG
    assert flagged[0].severity is Severity.HIGH
    assert "74,4" in flagged[0].message  # ~$74.4K understated


@pytest.mark.parametrize("loan", ["Strada", "Campus at Villa La Jolla", "Ares55thAve"])
def test_double_count_check_is_silent_elsewhere(results, loan):
    assert _findings(results, loan, "CHK_VACANCY_DOUBLE_COUNT") == []


# -- recomputes ---------------------------------------------------------------


@pytest.mark.parametrize("loan", LOANS)
def test_gpr_rebuild_ties_to_the_model(results, loan):
    assert _status(results, loan, "CHK_GPR_RECOMPUTE") is PASS


def test_strada_gpr_rebuild_uses_occupied_tenant_rows(contexts, results):
    # Strada reaches its rents through a SUMIF summary block, so both the tenant
    # range (T7:T517) and the summary range (M531:M539) are reachable and sum to
    # the same total. Only the tenant range aligns with the statuses in H.
    parse = contexts["Strada"].facts["rent_roll_parse"]
    assert parse.rent_column == "T" and parse.status_column == "H"
    assert len(parse.included_rows()) == 457  # occupied units per the rent roll
    assert parse.annualized() == pytest.approx(7_693_836.0, abs=1.0)


def test_recompute_flags_a_derived_rent_column_as_a_weaker_tie(results):
    # Campus and Hialeah foot to a column the model calculates, so the finding
    # must not claim the per-tenant adjustments were independently verified.
    for loan in ("Campus at Villa La Jolla", "Hialeah"):
        finding = _findings(results, loan, "CHK_GPR_RECOMPUTE")[0]
        assert "calculated by the model" in finding.message
    for loan in ("Strada", "Ares55thAve"):
        assert "calculated by the model" not in _findings(results, loan, "CHK_GPR_RECOMPUTE")[0].message


@pytest.mark.parametrize("loan", LOANS)
def test_vacancy_recompute_ties_to_the_model(results, loan):
    assert _status(results, loan, "CHK_VACANCY_RECOMPUTE") is PASS


def test_ares_vacancy_recompute_reproduces_the_floor_deduction(results):
    finding = _findings(results, "Ares55thAve", "CHK_VACANCY_RECOMPUTE")[0]
    assert "-39,820" in finding.evidence


@pytest.mark.parametrize("loan", ["Strada", "Campus at Villa La Jolla", "Hialeah"])
def test_vacancy_is_zero_where_actual_vacancy_exceeds_the_floor(results, loan):
    finding = _findings(results, loan, "CHK_VACANCY_RECOMPUTE")[0]
    assert finding.status is PASS
    assert "correctly 0" in finding.message
