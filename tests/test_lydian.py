"""Acceptance tests for the top-line revenue checks, built from the Lydian miss.

Lydian's rent roll sums rent for all seven unit statuses into GPR - including
14 Applicant rows (people who have not moved in) and 13 Pending-renewal rows
(duplicates of units already counted as Occupied) - overstating GPR by
$572,580/yr (+14%). The workbook's own unit count excludes those statuses, its
own `RR!K261` shows rent per occupied unit jumping +11.9% in a quarter, and the
first release of this tool passed it: the GPR trace stopped at the summary
block `SUM(G251:G257)` and re-summed the model's own subtotals.

Strada uses the identical rent-roll report and excludes those statuses
correctly (`M542 = M540-M532-M537-M536`), so every check here is asserted in
both directions: Lydian must flag, Strada and the rest must not.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from dy_audit.audit import audit_loan
from dy_audit.discovery import discover
from dy_audit.model import Severity, Status

INPUT_DIR = Path(__file__).resolve().parent.parent

FLAG, PASS = Status.FLAG, Status.PASS
OTHER_LOANS = ("Strada", "Campus at Villa La Jolla", "Hialeah", "Ares55thAve")


@pytest.fixture(scope="module")
def results():
    pairs, problems = discover(INPUT_DIR)
    assert problems == []
    return {p.loan_name: audit_loan(p) for p in pairs}


def _findings(results, loan, check_id):
    return [f for f in results[loan].findings if f.check_id == check_id]


def _one(results, loan, check_id):
    found = _findings(results, loan, check_id)
    assert len(found) == 1, f"{loan} {check_id}: {[f.status for f in found]}"
    return found[0]


# -- the deep trace: the recompute must reach tenant rows, not summary blocks --


def test_lydian_gpr_recompute_reaches_the_tenant_rows_and_flags(results):
    finding = _one(results, "Lydian", "CHK_GPR_RECOMPUTE")
    assert finding.status is FLAG
    # The independent figure is the occupied-status rent, not the model's total.
    assert "4,081,248" in finding.evidence
    assert "-572,580" in finding.message
    # 192 occupied rows via the status column, not the 7-cell summary block.
    assert "192 included tenant row(s)" in finding.evidence
    assert "status column G" in finding.evidence


def test_the_other_rent_rolls_still_tie(results):
    for loan in OTHER_LOANS:
        assert _one(results, loan, "CHK_GPR_RECOMPUTE").status is PASS


def test_strada_recompute_still_uses_its_tenant_range(results):
    # The deep trace adds candidate ranges on Strada too (its summary block is
    # now expanded as well); the parse must still land on the tenant rows.
    finding = _one(results, "Strada", "CHK_GPR_RECOMPUTE")
    assert "457 included tenant row(s)" in finding.evidence
    assert "status column H" in finding.evidence


# -- CHK_RENT_STATUS: vacant / applicant / pending rent -----------------------


def test_lydian_flags_applicant_and_pending_renewal_rent(results):
    finding = _one(results, "Lydian", "CHK_RENT_STATUS")
    assert finding.status is FLAG
    assert finding.severity is Severity.HIGH
    assert "'Applicant': 14 row(s), 287,760 annualized" in finding.message
    assert "'Pending renewal': 13 row(s), 284,820 annualized" in finding.message
    # The pending-renewal rows duplicate units already counted as occupied.
    assert "share a unit with a row already counted as occupied" in finding.message
    assert finding.on_dy_path


def test_strada_excludes_the_same_statuses_and_passes(results):
    finding = _one(results, "Strada", "CHK_RENT_STATUS")
    assert finding.status is PASS
    assert "correctly excluded" in finding.message
    assert "'Applicant': 10 row(s)" in finding.message


def test_loans_without_a_status_column_emit_no_status_finding(results):
    # Campus marks vacancy by tenant name, Hialeah by section captions, Ares
    # not at all; none has a per-row status column, so the check must stay
    # quiet rather than inventing one (their GPR recompute carries the caveat).
    for loan in ("Campus at Villa La Jolla", "Hialeah", "Ares55thAve"):
        assert _findings(results, loan, "CHK_RENT_STATUS") == []


# -- CHK_GPR_TREND: sanity against the prior DY test in column G ---------------


def test_lydian_rent_per_unit_jump_flags(results):
    finding = _one(results, "Lydian", "CHK_GPR_TREND")
    assert finding.status is FLAG
    assert "+11.9%" in finding.message
    # The per-unit dollars the reviewer would recognise from the model itself.
    assert "2,019.89" in finding.message and "1,804.42" in finding.message


@pytest.mark.parametrize("loan", OTHER_LOANS)
def test_prior_quarter_trend_passes_on_the_q1_set(results, loan):
    # Actual margins: Strada -0.1%, Campus -1.6%, Hialeah +3.4%, Ares +0.0% -
    # all inside the 5% band, so a false positive here means the band is wrong.
    assert _one(results, loan, "CHK_GPR_TREND").status is PASS


# -- CHK_REVENUE_DOUBLE_COUNT --------------------------------------------------


def test_no_loan_double_counts_a_revenue_source(results):
    for loan in ("Lydian", *OTHER_LOANS):
        finding = _one(results, loan, "CHK_REVENUE_DOUBLE_COUNT")
        assert finding.status is PASS, f"{loan}: {finding.message}"


def test_lydian_netted_parking_is_not_a_false_positive(results):
    # Other Income = T3 other income + parking rows - parking - concessions;
    # the parking row enters Parking (+) and Other Income (-), which is netting,
    # not double-counting. The check must see both lines and stay a PASS.
    finding = _one(results, "Lydian", "CHK_REVENUE_DOUBLE_COUNT")
    assert "OTHER_INCOME" in finding.evidence and "PARKING" in finding.evidence


# -- the DY column must not drift onto Lydian's variance helpers ---------------


def test_lydian_headline_figures(results):
    facts = results["Lydian"].facts
    assert facts["dy_column"] == "I"
    assert facts["debt_yield"] == pytest.approx(0.06370338683206106)
    assert facts["ncf"] == pytest.approx(3338057.47)
    assert facts["upb"] == pytest.approx(52400000.0)


def test_lydian_inherits_stradas_inverted_vacancy_branch(results):
    # Same Knightvest template, same latent blocker: `M25 = 5%-(1-I16)` turns
    # the vacancy line positive above the floor.
    finding = _one(results, "Lydian", "CHK_VACANCY_SIGN")
    assert finding.status is FLAG
    assert finding.severity is Severity.BLOCKER
