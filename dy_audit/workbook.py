"""Workbook access layer.

Every model is opened **twice**: once for formulas and once for the values Excel
cached at its last save. The tool never recalculates. Campus's column H is built
on `XLOOKUP`, which neither openpyxl nor LibreOffice evaluates, so a recalc pass
would silently blank out real figures. Cached values are the source of truth for
numbers; formulas are the source of truth for structure.
"""

from __future__ import annotations

import warnings
from dataclasses import dataclass
from pathlib import Path
from typing import Iterator

import openpyxl
from openpyxl.utils import get_column_letter

#: Excel error literals that can be cached in a cell.
ERROR_VALUES = frozenset(
    {"#REF!", "#DIV/0!", "#VALUE!", "#NAME?", "#N/A", "#NULL!", "#NUM!", "#GETTING_DATA"}
)

#: Guard against pathological max_row/max_column on sheets with stray formatting.
_MAX_SCAN_ROWS = 5000
_MAX_SCAN_COLS = 120


@dataclass(frozen=True)
class CellRef:
    sheet: str
    coord: str

    def __str__(self) -> str:
        return f"{self.sheet}!{self.coord}"


class Workbook:
    """Dual-view access to one .xlsx model."""

    def __init__(self, path: Path):
        self.path = Path(path)
        with warnings.catch_warnings():
            # These models carry legacy data-validation and defined-name residue
            # (Campus has 12,024 defined names); openpyxl warns but reads fine.
            warnings.simplefilter("ignore")
            self._wb_f = openpyxl.load_workbook(self.path, data_only=False)
            self._wb_v = openpyxl.load_workbook(self.path, data_only=True)

    def close(self) -> None:
        self._wb_f.close()
        self._wb_v.close()

    def __enter__(self) -> "Workbook":
        return self

    def __exit__(self, *exc) -> None:
        self.close()

    # -- sheet inventory ---------------------------------------------------

    @property
    def sheet_names(self) -> list[str]:
        return list(self._wb_f.sheetnames)

    def state(self, sheet: str) -> str:
        """'visible', 'hidden', or 'veryHidden'."""
        return self._wb_f[sheet].sheet_state

    def is_visible(self, sheet: str) -> bool:
        return self.state(sheet) == "visible"

    def visible_sheets(self) -> list[str]:
        return [s for s in self.sheet_names if self.is_visible(s)]

    def has_sheet(self, sheet: str) -> bool:
        return sheet in self._wb_f.sheetnames

    # -- cell access -------------------------------------------------------

    def formula(self, sheet: str, coord: str) -> str | None:
        """Formula text (leading '=' included) or None when the cell is a literal.

        openpyxl translates shared formulas on read, so Campus's shared column-I
        cells return their own translated text rather than an empty string.
        """
        if not self.has_sheet(sheet):
            return None
        val = self._wb_f[sheet][coord].value
        if isinstance(val, str) and val.startswith("="):
            return val
        return None

    def literal(self, sheet: str, coord: str):
        """Raw contents from the formula view: a formula string or a literal."""
        if not self.has_sheet(sheet):
            return None
        return self._wb_f[sheet][coord].value

    def value(self, sheet: str, coord: str):
        """Cached value Excel stored at last save."""
        if not self.has_sheet(sheet):
            return None
        return self._wb_v[sheet][coord].value

    def number(self, sheet: str, coord: str) -> float | None:
        """Cached value as a float, or None if it is not numeric."""
        v = self.value(sheet, coord)
        if isinstance(v, bool):
            return None
        if isinstance(v, (int, float)):
            return float(v)
        return None

    def text(self, sheet: str, coord: str) -> str | None:
        """Cached value as trimmed text, preferring the literal for label cells."""
        v = self.value(sheet, coord)
        if v is None:
            v = self.literal(sheet, coord)
            if isinstance(v, str) and v.startswith("="):
                return None
        if v is None:
            return None
        return str(v).strip()

    def is_error(self, sheet: str, coord: str) -> bool:
        v = self.value(sheet, coord)
        return isinstance(v, str) and v.strip() in ERROR_VALUES

    # -- iteration ---------------------------------------------------------

    def bounds(self, sheet: str) -> tuple[int, int]:
        ws = self._wb_f[sheet]
        return (
            min(ws.max_row or 1, _MAX_SCAN_ROWS),
            min(ws.max_column or 1, _MAX_SCAN_COLS),
        )

    def iter_cells(
        self, sheet: str, max_row: int | None = None, max_col: int | None = None
    ) -> Iterator[tuple[str, str | None, object]]:
        """Yield (coord, formula_or_None, cached_value) for populated cells."""
        if not self.has_sheet(sheet):
            return
        rows, cols = self.bounds(sheet)
        rows = min(rows, max_row) if max_row else rows
        cols = min(cols, max_col) if max_col else cols
        ws_f = self._wb_f[sheet]
        ws_v = self._wb_v[sheet]
        for r in range(1, rows + 1):
            for c in range(1, cols + 1):
                raw = ws_f.cell(row=r, column=c).value
                cached = ws_v.cell(row=r, column=c).value
                if raw is None and cached is None:
                    continue
                coord = f"{get_column_letter(c)}{r}"
                formula = raw if isinstance(raw, str) and raw.startswith("=") else None
                yield coord, formula, cached

    def iter_all_cells(self) -> Iterator[tuple[str, str, str | None, object]]:
        """Yield (sheet, coord, formula, cached_value) across every sheet.

        Hidden sheets are included: the hidden-tab rule governs OSAR selection and
        DY auditing only, while whole-file hygiene sweeps (error cells) still
        cover them and report their findings as off-path.
        """
        for sheet in self.sheet_names:
            for coord, formula, cached in self.iter_cells(sheet):
                yield sheet, coord, formula, cached
