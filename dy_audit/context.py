"""Per-loan context handed to every check."""

from __future__ import annotations

from dataclasses import dataclass, field

from .model import LoanFiles
from .osar import OsarTab
from .workbook import Workbook


@dataclass
class LoanContext:
    """Everything a check needs about one loan, assembled once per run."""

    files: LoanFiles
    wb: Workbook
    tab: OsarTab
    #: Scratch space for facts one stage learns and a later stage reuses
    #: (the UPB cell found by CHK_DY_BASIS, the occupancy driver, ...).
    facts: dict = field(default_factory=dict)

    @property
    def loan_name(self) -> str:
        return self.files.loan_name

    @property
    def sheet(self) -> str:
        return self.tab.sheet
