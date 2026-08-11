"""Acceptance tests for the BLOCKER checks against the Q1 2026 findings.

Asserted in both directions. A check that fires on an item spec section 6 marks
correct is as much a failure as a missed flag, so every loan pins the full status
of all five checks rather than only the ones expected to flag.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from dy_audit import formula as F
from dy_audit.checks.blockers import (
    check_dy_basis,
    check_dy_consistency,
    check_ins_max,
    check_tax_max,
    check_vacancy_sign,
    run_blockers,
)
from dy_audit.context import LoanContext
from dy_audit.definitions import parse_definitions
from dy_audit.discovery import discover
from dy_audit.model import Status
from dy_audit.osar import Line, select_osar
from dy_audit.workbook import Workbook

INPUT_DIR = Path(__file__).resolve().parent.parent

FLAG, PASS = Status.FLAG, Status.PASS

#: loan -> {check_id: expected status}, straight from spec section 6.
EXPECTED = {
    # S-1 insurance MAX missing; S-3 vacancy sign inverted. Tax MAX is correct.
    "Strada": {
        "CHK_TAX_MAX": PASS,
        "CHK_INS_MAX": FLAG,
        "CHK_VACANCY_SIGN": FLAG,
        "CHK_DY_BASIS": PASS,
        "CHK_DY_CONSISTENCY": PASS,
    },
    # "Cleanest of the four" - no blocker may fire.
    "Campus at Villa La Jolla": {
        "CHK_TAX_MAX": PASS,
        "CHK_INS_MAX": PASS,
        "CHK_VACANCY_SIGN": PASS,
        "CHK_DY_BASIS": PASS,
        "CHK_DY_CONSISTENCY": PASS,
    },
    # H-4 two conflicting DYs; everything else correct.
    "Hialeah": {
        "CHK_TAX_MAX": PASS,
        "CHK_INS_MAX": PASS,
        "CHK_VACANCY_SIGN": PASS,
        "CHK_DY_BASIS": PASS,
        "CHK_DY_CONSISTENCY": FLAG,
    },
    # A-2 insurance invoice term broken. Vacancy flags at the boundary: the
    # branch uses a strict `<95%`, so at exactly 95% occupancy it deducts where
    # the line must be 0. Spec section 6 marks this correct; the deliberate
    # divergence was ruled on during the build.
    "Ares55thAve": {
        "CHK_TAX_MAX": PASS,
        "CHK_INS_MAX": FLAG,
        "CHK_VACANCY_SIGN": FLAG,
        "CHK_DY_BASIS": PASS,
        "CHK_DY_CONSISTENCY": PASS,
    },
}


@pytest.fixture(scope="module")
def contexts():
    """Open each model once; loading a workbook twice per loan is not cheap."""
    pairs, problems = discover(INPUT_DIR)
    assert problems == []
    opened = {}
    for pair in pairs:
        wb = Workbook(pair.xlsx_path)
        opened[pair.loan_name] = LoanContext(
            files=pair,
            wb=wb,
            tab=select_osar(wb),
            params=parse_definitions(pair.defs_path, pair.loan_name),
        )
    yield opened
    for ctx in opened.values():
        ctx.wb.close()


@pytest.fixture(scope="module")
def results(contexts):
    return {name: run_blockers(ctx) for name, ctx in contexts.items()}


def _status(findings, check_id) -> Status:
    matches = [f for f in findings if f.check_id == check_id]
    assert matches, f"{check_id} produced no finding"
    # A check that flags more than once (multiple rival DYs) is still a flag.
    return FLAG if any(f.status is FLAG for f in matches) else matches[0].status


@pytest.mark.parametrize("loan", sorted(EXPECTED))
def test_blocker_statuses_match_q1_2026_findings(results, loan):
    actual = {cid: _status(results[loan], cid) for cid in EXPECTED[loan]}
    assert actual == EXPECTED[loan]


def test_exactly_five_blocker_flags_across_the_portfolio(results):
    # S-1, S-3, H-4, A-2, plus Ares's boundary defect - no more, no fewer.
    flagged = {
        (loan, f.check_id)
        for loan, findings in results.items()
        for f in findings
        if f.status is FLAG
    }
    assert flagged == {
        ("Strada", "CHK_INS_MAX"),
        ("Strada", "CHK_VACANCY_SIGN"),
        ("Hialeah", "CHK_DY_CONSISTENCY"),
        ("Ares55thAve", "CHK_INS_MAX"),
        ("Ares55thAve", "CHK_VACANCY_SIGN"),
    }


def test_no_blocker_is_left_unverifiable(results):
    stuck = [
        (loan, f.check_id, f.status.value)
        for loan, findings in results.items()
        for f in findings
        if f.status in (Status.UNVERIFIABLE, Status.MANUAL_REVIEW)
    ]
    assert stuck == []


# -- S-1 / A-2: the two insurance failure modes ------------------------------


def test_strada_insurance_has_no_max_at_all(contexts):
    # S-1: I37 is an INDEX/MATCH into the operating statement and never looks at
    # the Insurance tab, so a rising premium would go unnoticed.
    finding = check_ins_max(contexts["Strada"])[0]
    assert finding.status is FLAG
    assert "only one side" in finding.message
    assert "INDEX" in finding.evidence and "MAX" not in finding.evidence.split("=", 1)[1]


def test_ares_insurance_max_is_present_but_can_never_win(contexts):
    # A-2: the invoice term is a $-in-millions display divided by 1,000,000.
    finding = check_ins_max(contexts["Ares55thAve"])[0]
    assert finding.status is FLAG
    assert "can never win" in finding.message
    assert "Insurance!C3/Insurance!C4" in finding.evidence


def test_ares_insurance_resolves_through_the_lender_calc_rollup(contexts):
    # The spec's appendix cites Lender Calc!E21 directly, but the OSAR points at
    # E20, which rolls up E21. Resolution has to walk both hops or the MAX check
    # sees only a bare reference and passes.
    ctx = contexts["Ares55thAve"]
    res = F.resolve_defining_formula(ctx.wb, ctx.tab.sheet, ctx.tab.cell(Line.INSURANCE))
    assert res.path == ["Comm OSAR!I39", "Lender Calc!E20", "Lender Calc!E21"]
    assert res.formula.upper().startswith("=MAX(")


@pytest.mark.parametrize("loan", ["Campus at Villa La Jolla", "Hialeah"])
def test_working_insurance_max_is_not_flagged(results, loan):
    assert _status(results[loan], "CHK_INS_MAX") is PASS


# -- S-3: the vacancy sign, invisible at current occupancy -------------------


def _vacancy_trace(ctx) -> dict[float, float]:
    """Re-run the vacancy formula across the occupancy scenarios."""
    res = F.resolve_defining_formula(ctx.wb, ctx.tab.sheet, ctx.tab.cell(Line.VACANCY))
    deps = F.dependencies(ctx.wb, res.sheet, res.coord)
    occ_row = ctx.tab.rows[Line.OCCUPANCY]
    occ_key = next(
        k for k, (s, c) in deps.items() if s == ctx.tab.sheet and int(c[1:]) == occ_row
    )
    dependents = {occ_key} | {
        k for k, (s, c) in deps.items() if occ_key in F.dependencies(ctx.wb, s, c)
    }
    out = {}
    for occ in (1.00, 0.97, 0.92):
        resolver = F.make_resolver(ctx.wb, overrides={occ_key: occ}, recompute=dependents)
        out[occ] = F.to_number(F.evaluate(res.formula, res.sheet, resolver))
    return out


def test_strada_vacancy_adds_income_above_95_percent(contexts):
    # At the reported 91.95% occupancy the line is 0 and looks healthy; the
    # inversion only shows above the floor.
    trace = _vacancy_trace(contexts["Strada"])
    assert trace[1.00] > 0
    assert trace[0.97] > 0
    assert trace[0.92] == pytest.approx(0.0, abs=0.01)

    current = contexts["Strada"].wb.number("NEW OSAR", "I25")
    assert current == pytest.approx(0.0, abs=0.01), "the defect must be invisible today"


@pytest.mark.parametrize("loan", ["Campus at Villa La Jolla", "Hialeah", "Ares55thAve"])
def test_vacancy_formulas_deduct_above_the_floor_and_zero_below(contexts, loan):
    trace = _vacancy_trace(contexts[loan])
    assert trace[1.00] < 0, "above the floor the line must reduce EGI"
    assert trace[0.97] < 0
    assert trace[0.92] == pytest.approx(0.0, abs=0.01), "below the floor it must be 0"


def test_ares_vacancy_floor_bites_at_full_occupancy(contexts):
    # Spec section 6's worked example: 100% occupied gives -GPR x 5%.
    assert _vacancy_trace(contexts["Ares55thAve"])[1.00] == pytest.approx(-39820.188, abs=0.01)


# -- H-4: conflicting debt yields --------------------------------------------


def test_hialeah_reports_both_conflicting_debt_yields(contexts):
    findings = check_dy_consistency(contexts["Hialeah"])
    assert len(findings) == 1
    finding = findings[0]
    assert finding.status is FLAG
    assert finding.sheet == "Debt Service" and finding.cell == "C16"
    assert "2.3253%" in finding.message and "2.1506%" in finding.message


def test_campus_hidden_osar_debt_yield_is_not_a_conflict(contexts):
    # Campus's hidden OSAR tab carries a 9.20% debt yield. Hidden tabs are out of
    # scope for DY auditing, so this must not surface as a blocker.
    ctx = contexts["Campus at Villa La Jolla"]
    assert "OSAR" in ctx.wb.sheet_names and not ctx.wb.is_visible("OSAR")
    assert check_dy_consistency(ctx)[0].status is PASS


def test_ares_covenant_threshold_is_not_read_as_a_rival_calculation(contexts):
    # Debt Service!C14 "DY Event" = 6.50% is the covenant level, stored as a
    # literal. A hardcoded number is an input, not a competing calculation.
    ctx = contexts["Ares55thAve"]
    assert ctx.wb.number("Debt Service", "C14") == pytest.approx(0.065)
    assert ctx.wb.formula("Debt Service", "C14") is None
    assert check_dy_consistency(ctx)[0].status is PASS


# -- DY basis ----------------------------------------------------------------


@pytest.mark.parametrize("loan", sorted(EXPECTED))
def test_debt_yield_is_driven_off_ncf_not_noi(contexts, loan):
    ctx = contexts[loan]
    finding = check_dy_basis(ctx)[0]
    assert finding.status is PASS
    numerator_row = ctx.tab.rows[Line.NCF]
    assert f"I{numerator_row}" in F.normalize(
        ctx.wb.formula(ctx.tab.sheet, ctx.tab.cell(Line.DEBT_YIELD))
    )


def test_dy_basis_records_the_upb_reference_for_the_standing_flag(contexts):
    # UPB has no source file, so the standing flag must echo the model's own
    # value - which means knowing which cell the DY divided by. Ares uses D6,
    # the other three use E6.
    expected = {
        "Strada": ("NEW OSAR", "E6"),
        "Campus at Villa La Jolla": ("Comm OSAR", "E6"),
        "Hialeah": ("(New) Comm OSAR", "E6"),
        "Ares55thAve": ("Comm OSAR", "D6"),
    }
    for loan, ref in expected.items():
        ctx = contexts[loan]
        check_dy_basis(ctx)
        assert ctx.facts["upb_ref"] == ref


def test_tax_max_passes_on_every_loan(results):
    # Universal hard rule, and all four models honour it - so a false positive
    # here would be immediately visible.
    for loan in EXPECTED:
        assert _status(results[loan], "CHK_TAX_MAX") is PASS
        assert check_tax_max is not None


def test_vacancy_sign_check_is_conclusive_everywhere(results):
    for loan in EXPECTED:
        assert _status(results[loan], "CHK_VACANCY_SIGN") in (FLAG, PASS)
        assert check_vacancy_sign is not None
