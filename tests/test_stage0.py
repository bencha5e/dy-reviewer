"""Stage 0 acceptance: discovery, pairing, tab selection, period identification.

These assert against the four Q1 2026 models in the repository root. They are
read-only - no test may move or rewrite an input file.
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import pytest

from dy_audit.discovery import discover, loan_name_from_definitions
from dy_audit.osar import Line, is_quarter_end, normalize_label, select_osar
from dy_audit.workbook import Workbook

INPUT_DIR = Path(__file__).resolve().parent.parent

#: loan name -> (osar tab, dy column, label column, hidden OSAR tabs, DY)
EXPECTED = {
    "Strada": ("NEW OSAR", "I", "C", ["Comm OSAR"], 0.064657506908786189),
    "Campus at Villa La Jolla": ("Comm OSAR", "I", "C", ["OSAR"], 0.060705295896825412),
    "Hialeah": ("(New) Comm OSAR", "I", "C", [], 0.021506099536749368),
    "Ares55thAve": ("Comm OSAR", "I", "B", [], 0.07088346216872761),
}

QUARTER_END = dt.date(2026, 3, 31)


@pytest.fixture(scope="module")
def pairs():
    found, problems = discover(INPUT_DIR)
    assert problems == [], f"discovery reported problems: {problems}"
    return {p.loan_name: p for p in found}


def test_discovers_every_loan_pair(pairs):
    assert set(pairs) == set(EXPECTED)


def test_pairs_the_right_definitions_file(pairs):
    # Pairing is by parsed loan name, so a model must never borrow another
    # loan's definitions file just because the folder order lines up.
    assert pairs["Strada"].defs_path.name.startswith("1.")
    assert pairs["Campus at Villa La Jolla"].defs_path.name.startswith("2.")
    assert pairs["Hialeah"].defs_path.name.startswith("3.")
    assert pairs["Ares55thAve"].defs_path.name.startswith("4.")
    for name, pair in pairs.items():
        assert pair.xlsx_path.suffix == ".xlsx"
        assert "definition" in pair.defs_path.stem.lower()


def test_loan_name_comes_from_the_definitions_header(pairs):
    # The authoritative name is the `*****<Name>*****` header in the loan
    # agreement extract, not anything guessed from a filename.
    assert loan_name_from_definitions(pairs["Strada"].defs_path) == "Strada"
    assert (
        loan_name_from_definitions(pairs["Campus at Villa La Jolla"].defs_path)
        == "Campus at Villa La Jolla"
    )


@pytest.mark.parametrize("loan", sorted(EXPECTED))
def test_selects_the_visible_osar_tab(pairs, loan):
    sheet, dy_col, label_col, hidden, _ = EXPECTED[loan]
    with Workbook(pairs[loan].xlsx_path) as wb:
        tab = select_osar(wb)
        assert tab.sheet == sheet
        assert tab.dy_column == dy_col
        assert tab.label_column == label_col
        assert wb.is_visible(tab.sheet)
        # Hidden OSAR tabs are recorded but never selected or audited.
        assert sorted(tab.hidden_osar_tabs) == sorted(hidden)
        assert tab.sheet not in tab.hidden_osar_tabs


@pytest.mark.parametrize("loan", sorted(EXPECTED))
def test_period_is_quarter_end(pairs, loan):
    with Workbook(pairs[loan].xlsx_path) as wb:
        tab = select_osar(wb)
        assert tab.period_end == QUARTER_END
        assert is_quarter_end(tab.period_end)


@pytest.mark.parametrize("loan", sorted(EXPECTED))
def test_resolves_the_lines_the_checks_depend_on(pairs, loan):
    required = (
        Line.GPR,
        Line.VACANCY,
        Line.EGI,
        Line.TAX,
        Line.INSURANCE,
        Line.MGMT_FEE,
        Line.NOI,
        Line.CAPEX,
        Line.NCF,
        Line.DEBT_YIELD,
    )
    with Workbook(pairs[loan].xlsx_path) as wb:
        tab = select_osar(wb)
        missing = [line.value for line in required if not tab.has(line)]
        assert not missing, f"{loan}: unresolved OSAR lines {missing}"


@pytest.mark.parametrize("loan", sorted(EXPECTED))
def test_reported_debt_yield_matches_the_model(pairs, loan):
    expected_dy = EXPECTED[loan][4]
    with Workbook(pairs[loan].xlsx_path) as wb:
        tab = select_osar(wb)
        dy = wb.number(tab.sheet, tab.cell(Line.DEBT_YIELD))
        assert dy == pytest.approx(expected_dy, abs=1e-9)


def test_normalize_label_handles_template_decoration():
    assert normalize_label("  *Effective Gross Income") == "effective gross income"
    assert normalize_label("      Gross Potential Rent (3)") == "gross potential rent"
    assert normalize_label("         Less: Vacancy Loss") == "less: vacancy loss"
    assert normalize_label("* Debt Yield: (NCF/Debt Balance)") == "debt yield: (ncf/debt balance)"
    # "Net Cash Flow after Debt Service" must stay distinguishable from NCF.
    assert normalize_label(" *Net Cash Flow") == "net cash flow"
    assert normalize_label(" *Net Cash Flow after Debt Service") != "net cash flow"


def test_ncf_row_is_not_the_after_debt_service_row(pairs):
    # A row-offset here would silently move the DY numerator onto the wrong line.
    with Workbook(pairs["Campus at Villa La Jolla"].xlsx_path) as wb:
        tab = select_osar(wb)
        assert tab.rows[Line.NCF] == 62
        assert tab.raw_labels[62].strip().startswith("*Net Cash Flow")
