"""Core data types shared by every stage of the audit.

Findings are the single currency of this tool: every check returns them, the
report layer renders them, and the acceptance tests assert on them.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from enum import Enum
from pathlib import Path


class Severity(str, Enum):
    """Severity ladder from spec section 4, plus INFO for recorded judgment calls."""

    BLOCKER = "BLOCKER"
    HIGH = "HIGH"
    MEDIUM = "MEDIUM"
    LOW = "LOW"
    INFO = "INFO"
    STANDING = "STANDING"


#: Sort order for reports: loudest first.
SEVERITY_ORDER = {
    Severity.BLOCKER: 0,
    Severity.HIGH: 1,
    Severity.MEDIUM: 2,
    Severity.LOW: 3,
    Severity.STANDING: 4,
    Severity.INFO: 5,
}


class Status(str, Enum):
    """Outcome of a check.

    MANUAL_REVIEW and UNVERIFIABLE exist so that a check the tool could not
    complete never reads as a clean pass. See the rent-roll fallback rule.
    """

    FLAG = "FLAG"
    PASS = "PASS"
    MANUAL_REVIEW = "MANUAL_REVIEW"
    UNVERIFIABLE = "UNVERIFIABLE"


@dataclass
class Finding:
    check_id: str
    severity: Severity
    status: Status
    message: str
    sheet: str | None = None
    cell: str | None = None
    evidence: str | None = None
    #: True/False when the check knows whether the cell feeds the debt-yield
    #: calculation; None when the distinction does not apply.
    on_dy_path: bool | None = None

    @property
    def ref(self) -> str:
        """Human-readable `Sheet!Cell` locator, or an empty string."""
        if self.sheet and self.cell:
            return f"{self.sheet}!{self.cell}"
        return self.sheet or self.cell or ""

    def is_flag(self) -> bool:
        return self.status is Status.FLAG


@dataclass
class LoanFiles:
    """One loan's input pair, as discovered in the input folder."""

    loan_name: str
    xlsx_path: Path
    defs_path: Path
    #: Numeric filename prefix ("1", "2", ...) when present; used only as a
    #: pairing signal, never as an ordering assumption.
    prefix: str | None = None

    def __str__(self) -> str:
        return f"{self.loan_name} ({self.xlsx_path.name} + {self.defs_path.name})"


@dataclass
class LoanResult:
    """Everything one loan's processing produced, success or failure."""

    loan_name: str
    files: LoanFiles
    findings: list[Finding] = field(default_factory=list)
    facts: dict = field(default_factory=dict)
    error: str | None = None

    @property
    def succeeded(self) -> bool:
        return self.error is None

    def add(self, finding: Finding) -> Finding:
        self.findings.append(finding)
        return finding

    def flags(self) -> list[Finding]:
        return [f for f in self.findings if f.is_flag()]

    def count_by_severity(self) -> dict[Severity, int]:
        """Findings that need a reviewer's attention, by severity.

        UNVERIFIABLE counts alongside FLAG and MANUAL_REVIEW. A check the tool
        could not complete must never read as a clean pass, and leaving it out
        of the headline counts is exactly that - a loan whose whole revenue
        review failed would otherwise summarise as quietly as a clean one.
        """
        counts: dict[Severity, int] = {}
        for f in self.findings:
            if f.status in (
                Status.FLAG,
                Status.MANUAL_REVIEW,
                Status.UNVERIFIABLE,
            ) or f.severity in (Severity.STANDING, Severity.INFO):
                counts[f.severity] = counts.get(f.severity, 0) + 1
        return counts
