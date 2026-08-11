"""Per-loan context handed to every check."""

from __future__ import annotations

from dataclasses import dataclass, field

from .model import LoanFiles
from .osar import OsarTab
from .params import LoanParams
from .workbook import Workbook


@dataclass
class LoanContext:
    """Everything a check needs about one loan, assembled once per run."""

    files: LoanFiles
    wb: Workbook
    tab: OsarTab
    #: Calculation rules read from the loan agreement. Checks that need a
    #: loan-specific number (the vacancy floor, the reserve rate) must read it
    #: from here rather than assuming the value common to today's roster.
    params: LoanParams | None = None
    #: Scratch space for facts one stage learns and a later stage reuses
    #: (the UPB cell found by CHK_DY_BASIS, the occupancy driver, ...).
    facts: dict = field(default_factory=dict)
    #: When set, the revenue line items are reviewed by an LLM and the
    #: deterministic revenue findings are dropped in favour of its verdict.
    #: The deterministic revenue checks still *run* either way: they produce the
    #: rent-roll rebuild the report renders and the source tabs later stages
    #: read, so only their findings are suppressed, never their side effects.
    use_llm: bool = False
    llm_provider: str = "anthropic"
    llm_model: str | None = None
    llm_effort: str = "xhigh"

    @property
    def loan_name(self) -> str:
        return self.files.loan_name

    @property
    def sheet(self) -> str:
        return self.tab.sheet
