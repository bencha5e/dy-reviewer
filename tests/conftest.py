"""Shared fixtures.

Anything that moves or writes files works on a copy in `tmp_path`. No test may
touch the repository's own copies of the four models - the move logic is
destructive by design.
"""

from __future__ import annotations

import shutil
from pathlib import Path

import pytest

REPO_ROOT = Path(__file__).resolve().parent.parent


@pytest.fixture
def input_copy(tmp_path: Path) -> Path:
    """A throwaway input folder holding copies of all four loan pairs."""
    folder = tmp_path / "Input DY Tests"
    folder.mkdir()
    for pattern in ("*.xlsx", "*DYDefinitions.md"):
        for source in REPO_ROOT.glob(pattern):
            shutil.copy2(source, folder / source.name)
    return folder
