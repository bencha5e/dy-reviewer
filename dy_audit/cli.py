"""Command-line entry point: `python -m dy_audit` or `python audit_dy.py`.

One run reviews every loan pair in the input folder, writes each loan's findings
workbook into its own output folder, and drops a run-level log at the output root.
Each loan is processed independently so one broken workbook cannot stop the rest.
"""

from __future__ import annotations

import argparse
import datetime as dt
import sys
from pathlib import Path

from .audit import audit_loan
from .discovery import discover
from .filemove import finalise_loan
from .model import LoanResult, Severity, Status
from .report import findings_filename, write_findings_workbook
from .runlog import write_summary

#: The folder the spec names. Overridable so the tool runs anywhere.
DEFAULT_INPUT = Path(
    r"C:\Users\bstamp\OneDrive - Bellwether\Desktop\Projects\DebtYieldFiles\Input DY Tests"
)
DEFAULT_OUTPUT_SUBFOLDER = "DY Review Output"


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="audit_dy",
        description="Review quarterly debt-yield test models against their loan definitions.",
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=DEFAULT_INPUT,
        help="Folder holding this quarter's .xlsx models and their definitions files.",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=None,
        help=f"Output root. Defaults to '<input-dir>/{DEFAULT_OUTPUT_SUBFOLDER}'.",
    )
    parser.add_argument(
        "--move",
        action="store_true",
        help=(
            "Move each successfully reviewed loan's source files out of the input queue into "
            "its output folder. Off by default, so a run never disturbs the input folder "
            "unless asked."
        ),
    )
    parser.add_argument(
        "--loan",
        action="append",
        default=None,
        metavar="NAME",
        help="Review only loans whose name contains NAME. Repeatable.",
    )
    parser.add_argument("--quiet", action="store_true", help="Suppress per-loan console output.")
    return parser


def _echo(message: str, quiet: bool = False) -> None:
    if not quiet:
        print(message)


def _summarise(result: LoanResult) -> str:
    if not result.succeeded:
        return f"FAILED - {result.error}"
    counts = result.count_by_severity()
    dy = result.facts.get("debt_yield")
    review = sum(1 for f in result.findings if f.status is Status.MANUAL_REVIEW)
    parts = [f"DY {dy:.4%}" if dy is not None else "DY not read"]
    for severity in (Severity.BLOCKER, Severity.HIGH, Severity.MEDIUM, Severity.LOW):
        if counts.get(severity):
            parts.append(f"{counts[severity]} {severity.value.lower()}")
    if review:
        parts.append(f"{review} to review")
    return ", ".join(parts)


def run(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir) if args.output_dir else input_dir / DEFAULT_OUTPUT_SUBFOLDER

    if not input_dir.is_dir():
        print(f"Input folder does not exist: {input_dir}", file=sys.stderr)
        return 2

    run_at = dt.datetime.now()
    run_date = run_at.date()

    pairs, problems = discover(input_dir)
    if args.loan:
        wanted = [name.lower() for name in args.loan]
        pairs = [p for p in pairs if any(w in p.loan_name.lower() for w in wanted)]

    if not pairs:
        _echo(f"No loan pairs found in {input_dir}", args.quiet)
        for problem in problems:
            _echo(f"  {problem}", args.quiet)
        write_summary([], problems, output_dir, run_at, input_dir, args.move)
        return 1 if problems else 0

    _echo(f"Reviewing {len(pairs)} loan(s) from {input_dir}", args.quiet)
    results: list[LoanResult] = []

    for pair in pairs:
        # Each loan is isolated: a workbook that cannot be read is recorded as
        # this loan's failure and leaves its source files in the queue.
        result = audit_loan(pair)
        results.append(result)
        try:
            findings_path = output_dir / findings_filename(result.loan_name, run_date)
            write_findings_workbook(result, findings_path, run_date)
            finalise_loan(
                result,
                output_dir,
                findings_path,
                run_date,
                move_sources=args.move and result.succeeded,
            )
        except Exception as exc:  # noqa: BLE001 - reporting must not lose the audit
            result.error = result.error or f"could not write output: {type(exc).__name__}: {exc}"
        _echo(f"  {result.loan_name}: {_summarise(result)}", args.quiet)

    log_path = write_summary(results, problems, output_dir, run_at, input_dir, args.move)
    _echo(f"Run log: {log_path}", args.quiet)

    failed = [r for r in results if not r.succeeded]
    if failed:
        _echo(
            f"{len(failed)} loan(s) failed and were left in the input folder for the next run.",
            args.quiet,
        )
        return 1
    return 0


def main() -> None:
    raise SystemExit(run())


if __name__ == "__main__":
    main()
