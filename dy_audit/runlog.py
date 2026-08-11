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


#: The usage keys the review reports, in the order they are rendered.
_USAGE_KEYS = (
    ("input_tokens", "Input"),
    ("cache_creation_input_tokens", "Cache write"),
    ("cache_read_input_tokens", "Cache read"),
    ("output_tokens", "Output"),
)


def _cache_rate(usage: dict) -> float | None:
    """Cached reads as a share of every prompt token the run was billed for.

    Denominated on the whole prompt side rather than on cache traffic alone, so
    the number answers the question actually being asked - how much of what we
    sent came back from cache - and cannot be flattered by a run that cached
    little but re-read it often.
    """
    prompt = sum(
        usage.get(key, 0) or 0
        for key in ("input_tokens", "cache_read_input_tokens", "cache_creation_input_tokens")
    )
    if not prompt:
        return None
    return (usage.get("cache_read_input_tokens", 0) or 0) / prompt


def _cost_section(succeeded: list[LoanResult]) -> list[str]:
    """What the revenue review cost, per loan and for the run.

    Only rendered where a review actually ran. A deterministic run has no usage
    to report, and an empty table would read as a review that cost nothing
    rather than one that never happened.
    """
    priced = [r for r in succeeded if isinstance(r.facts.get("llm_usage"), dict)]
    if not priced:
        return []

    lines = [
        "## Revenue review cost",
        "",
        "| Loan | " + " | ".join(label for _, label in _USAGE_KEYS) + " | Cache read |",
        "|---|" + "---:|" * (len(_USAGE_KEYS) + 1),
    ]

    totals = {key: 0 for key, _ in _USAGE_KEYS}
    for result in sorted(priced, key=lambda r: r.loan_name):
        usage = result.facts["llm_usage"]
        cells = []
        for key, _ in _USAGE_KEYS:
            value = usage.get(key, 0) or 0
            totals[key] += value
            cells.append(f"{value:,}")
        rate = _cache_rate(usage)
        cells.append("n/a" if rate is None else f"{rate:.1%}")
        lines.append(f"| {result.loan_name} | " + " | ".join(cells) + " |")

    rate = _cache_rate(totals)
    total_cells = [f"{totals[key]:,}" for key, _ in _USAGE_KEYS]
    total_cells.append("n/a" if rate is None else f"{rate:.1%}")
    lines.append("| **Run total** | " + " | ".join(total_cells) + " |")
    lines.append("")
    return lines


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
        f"- Source files moved out of the input queue: {'yes' if moved else 'no (--no-move)'}",
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
        lines += [
            "## Loans reviewed",
            "",
            "| Loan | Debt yield | Blockers | High | Needs review |",
            "|---|---:|---:|---:|---:|",
        ]
        for result in sorted(succeeded, key=lambda r: r.loan_name):
            dy = result.facts.get("debt_yield")
            counts = result.count_by_severity()
            # UNVERIFIABLE counts too: a check that could not run needs a human
            # just as much as one that returned an ambiguous answer.
            review = sum(
                1
                for f in result.findings
                if f.status in (Status.MANUAL_REVIEW, Status.UNVERIFIABLE)
            )
            yield_text = f"{dy:.4%}" if dy is not None else "not read"
            lines.append(
                f"| {result.loan_name} | {yield_text} | {counts.get(Severity.BLOCKER, 0)} | "
                f"{counts.get(Severity.HIGH, 0)} | {review} |"
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

    lines += _cost_section(succeeded)

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
