"""Run-level summary log written to the output root (spec section 0)."""

from __future__ import annotations

import datetime as dt
from pathlib import Path

from .model import SEVERITY_ORDER, LoanResult, Severity, Status


def _severity_totals(results: list[LoanResult]) -> dict[Severity, int]:
    totals: dict[Severity, int] = {}
    for result in results:
        for severity, count in result.count_by_severity().items():
            totals[severity] = totals.get(severity, 0) + count
    return totals


def build_summary(
    results: list[LoanResult],
    problems: list[str],
    run_at: dt.datetime,
    input_dir: Path,
    output_dir: Path,
    moved: bool,
) -> str:
    """Render the run summary as markdown."""
    succeeded = [r for r in results if r.succeeded]
    failed = [r for r in results if not r.succeeded]

    lines = [
        f"# Debt Yield Audit - run {run_at:%Y-%m-%d %H:%M}",
        "",
        f"- Input folder: `{input_dir}`",
        f"- Output folder: `{output_dir}`",
        f"- Loans processed: {len(results)} ({len(succeeded)} succeeded, {len(failed)} failed)",
        f"- Source files moved out of the input queue: {'yes' if moved else 'no (--move not set)'}",
        "",
    ]

    totals = _severity_totals(results)
    lines += ["## Findings by severity", ""]
    if totals:
        lines += ["| Severity | Count |", "|---|---:|"]
        for severity in sorted(Severity, key=lambda s: SEVERITY_ORDER.get(s, 9)):
            if totals.get(severity):
                lines.append(f"| {severity.value} | {totals[severity]} |")
    else:
        lines.append("No findings recorded.")
    lines.append("")

    if succeeded:
        lines += ["## Loans reviewed", "", "| Loan | Debt yield | Covenant | Blockers | High | Needs review |", "|---|---:|---|---:|---:|---:|"]
        for result in sorted(succeeded, key=lambda r: r.loan_name):
            facts = result.facts
            dy = facts.get("debt_yield")
            covenant = facts.get("covenant")
            if covenant is None:
                verdict = "not stated"
            elif dy is None:
                verdict = "n/a"
            else:
                verdict = "PASS" if dy >= covenant else "FAIL"
            counts = result.count_by_severity()
            review = sum(1 for f in result.findings if f.status is Status.MANUAL_REVIEW)
            lines.append(
                f"| {result.loan_name} | {dy:.4%} | {verdict} | "
                f"{counts.get(Severity.BLOCKER, 0)} | {counts.get(Severity.HIGH, 0)} | {review} |"
                if dy is not None
                else f"| {result.loan_name} | not read | {verdict} | "
                f"{counts.get(Severity.BLOCKER, 0)} | {counts.get(Severity.HIGH, 0)} | {review} |"
            )
        lines.append("")

        blockers = [
            (r.loan_name, f)
            for r in succeeded
            for f in r.findings
            if f.severity is Severity.BLOCKER and f.status is Status.FLAG
        ]
        if blockers:
            lines += ["## Blocker findings", ""]
            for loan, finding in blockers:
                lines.append(f"- **{loan}** `{finding.check_id}` at `{finding.ref}` - {finding.message}")
            lines.append("")

    if failed:
        lines += [
            "## Loans that failed and were left in the input queue",
            "",
            "These were not moved; they will be retried on the next run.",
            "",
        ]
        for result in sorted(failed, key=lambda r: r.loan_name):
            lines.append(f"- **{result.loan_name}** ({result.files.xlsx_path.name}): {result.error}")
        lines.append("")

    if problems:
        lines += ["## Files that could not be paired", ""]
        lines += [f"- {problem}" for problem in problems]
        lines.append("")

    return "\n".join(lines)


def write_summary(
    results: list[LoanResult],
    problems: list[str],
    output_dir: Path,
    run_at: dt.datetime | None = None,
    input_dir: Path | None = None,
    moved: bool = False,
) -> Path:
    run_at = run_at or dt.datetime.now()
    output_dir = Path(output_dir)
    output_dir.mkdir(parents=True, exist_ok=True)
    path = output_dir / f"DY Audit Run Log - {run_at:%Y_%m_%d_%H%M}.md"
    path.write_text(
        build_summary(results, problems, run_at, input_dir or output_dir, output_dir, moved),
        encoding="utf-8",
    )
    return path
