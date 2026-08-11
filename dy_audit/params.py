"""Loan parameters extracted from the loan agreement, plus the house rules.

The definitions file is authoritative. Where the build spec and a definitions
file disagree, the definitions file wins and the tool records the conflict
rather than resolving it silently.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any

#: Management-fee percentage base, fixed by house rule regardless of the wording
#: in any loan agreement. Every definitions file says "Gross Revenues" or "gross
#: operating income"; spec section 5 rule 3 resolves all of them to EGI. The
#: override is recorded as an INFO finding so it stays visible.
MGMT_FEE_BASE = "EGI"

#: Sentinel for a loan whose delinquency test is "current on rental obligations"
#: with no day count - stricter than any numeric window.
CURRENT = "CURRENT"


@dataclass
class Parsed:
    """A parameter value together with the loan-agreement text it came from."""

    value: Any = None
    quote: str | None = None

    @property
    def found(self) -> bool:
        return self.value is not None

    def __str__(self) -> str:
        return "not found" if not self.found else str(self.value)


@dataclass
class LoanParams:
    """Calculation rules for one loan, read out of its definitions file."""

    loan_name: str

    #: Vacancy factor floor, e.g. 0.05. The occupancy threshold is 1 - floor.
    vacancy_floor: Parsed = field(default_factory=Parsed)
    #: True when investment-grade tenants are carved out of the vacancy factor.
    ig_vacancy_carveout: Parsed = field(default_factory=Parsed)

    #: Replacement reserve rate and whether it is per unit or per square foot.
    reserve_rate: Parsed = field(default_factory=Parsed)
    reserve_basis: Parsed = field(default_factory=Parsed)

    #: Management fee percentage and the base the loan agreement names.
    mgmt_fee_pct: Parsed = field(default_factory=Parsed)
    mgmt_stated_base: Parsed = field(default_factory=Parsed)

    #: Days delinquent that exclude a tenant, or CURRENT when the agreement
    #: requires tenants simply be current.
    delinquency: Parsed = field(default_factory=Parsed)

    #: Trailing window for other income, in months, and for concessions where
    #: the agreement puts them on a different window (Strada: T3 and T6).
    other_income_months: Parsed = field(default_factory=Parsed)
    concession_months: Parsed = field(default_factory=Parsed)

    #: Informational windows; rent steps and free rent stay manual review.
    new_lease_days: Parsed = field(default_factory=Parsed)
    new_lease_ig_days: Parsed = field(default_factory=Parsed)
    rent_step_window: Parsed = field(default_factory=Parsed)
    known_vacate_days: Parsed = field(default_factory=Parsed)

    #: Conflicts between the build spec's parameter table and the loan text.
    conflicts: list[str] = field(default_factory=list)

    @property
    def occupancy_threshold(self) -> float | None:
        """Occupancy at or below which the vacancy line must be zero."""
        if not self.vacancy_floor.found:
            return None
        return 1.0 - float(self.vacancy_floor.value)

    @property
    def mgmt_base(self) -> str:
        """The base the tool enforces - always EGI, by house rule."""
        return MGMT_FEE_BASE

    def missing(self) -> list[str]:
        """Parameters the checks need but the parser could not find."""
        required = {
            "vacancy_floor": self.vacancy_floor,
            "reserve_rate": self.reserve_rate,
            "mgmt_fee_pct": self.mgmt_fee_pct,
        }
        return [name for name, parsed in required.items() if not parsed.found]
