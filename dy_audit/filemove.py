"""Output folder creation and the move out of the input queue (spec section 0).

Two rules govern everything here:

- **Never overwrite.** A second run on the same day for the same loan makes v2,
  a third v3. An existing folder is never written into.
- **Only move on success.** A loan whose processing failed keeps its source files
  in the input folder so the next run retries it.

The input folder is a queue of "not yet reviewed" files, so a successful run must
leave it empty of what it just processed.
"""

from __future__ import annotations

import datetime as dt
import re
import shutil
from pathlib import Path

from .model import LoanResult

#: Characters Windows will not accept in a folder name.
_UNSAFE = re.compile(r'[<>:"/\\|?*]')


def safe_folder_name(loan_name: str) -> str:
    cleaned = _UNSAFE.sub("_", loan_name).strip().rstrip(".")
    return cleaned or "Unnamed Loan"


def next_version_folder(output_root: Path, loan_name: str, run_date: dt.date) -> Path:
    """`<Loan Name> - <YYYY_MM_DD> - v<N>`, incrementing past anything present."""
    output_root = Path(output_root)
    base = f"{safe_folder_name(loan_name)} - {run_date:%Y_%m_%d}"
    version = 1
    while (output_root / f"{base} - v{version}").exists():
        version += 1
    return output_root / f"{base} - v{version}"


def _move(source: Path, destination_dir: Path) -> Path:
    """Move one file, refusing to clobber an existing name."""
    target = destination_dir / source.name
    if target.exists():
        stem, suffix = target.stem, target.suffix
        counter = 2
        while target.exists():
            target = destination_dir / f"{stem} ({counter}){suffix}"
            counter += 1
    shutil.move(str(source), str(target))
    return target


def finalise_loan(
    result: LoanResult,
    output_root: Path,
    findings_path: Path,
    run_date: dt.date | None = None,
    move_sources: bool = False,
) -> Path:
    """Create this loan's output folder and place its three files in it.

    The findings workbook always lands in the new folder. The source workbook and
    definitions file are moved out of the input queue only when `move_sources` is
    set and the loan actually succeeded.
    """
    run_date = run_date or dt.date.today()
    folder = next_version_folder(output_root, result.loan_name, run_date)
    folder.mkdir(parents=True, exist_ok=False)

    findings_path = Path(findings_path)
    if findings_path.exists() and findings_path.parent != folder:
        findings_path = _move(findings_path, folder)

    if move_sources and result.succeeded:
        for source in (result.files.xlsx_path, result.files.defs_path):
            if Path(source).exists():
                _move(Path(source), folder)

    result.facts["output_folder"] = str(folder)
    return folder
