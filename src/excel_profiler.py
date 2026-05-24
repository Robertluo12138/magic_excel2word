"""Read an Excel workbook into traceable per-cell records.

Each numeric cell becomes an :class:`ExcelCell` carrying the sheet, address,
raw and numeric values, and the surrounding row/column text that we will
later use as context tokens for matching. The profiler is intentionally
liberal — it returns *every* numeric cell rather than guessing which ones
are "interesting", so the matcher can search the full space.
"""
from __future__ import annotations

from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import openpyxl


@dataclass
class ExcelCell:
    sheet: str
    address: str
    row: int
    column: int
    raw_value: object
    numeric_value: float
    row_context: List[str] = field(default_factory=list)
    column_context: List[str] = field(default_factory=list)


def profile_workbook(path: Path) -> List[ExcelCell]:
    """Return every numeric cell in the workbook with surrounding text context."""
    wb = openpyxl.load_workbook(str(path), data_only=True)
    cells: List[ExcelCell] = []
    for sheet in wb.worksheets:
        # Pre-scan to map row→[(col, text)] and col→[(row, text)] for cheap
        # context lookup per numeric cell.
        row_text: Dict[int, List[Tuple[int, str]]] = defaultdict(list)
        col_text: Dict[int, List[Tuple[int, str]]] = defaultdict(list)
        numeric_cells: List[Tuple[int, int, str, object, float]] = []

        for row in sheet.iter_rows():
            for cell in row:
                v = cell.value
                if v is None:
                    continue
                if isinstance(v, str):
                    s = v.strip()
                    if s:
                        row_text[cell.row].append((cell.column, s))
                        col_text[cell.column].append((cell.row, s))
                    continue
                num = _to_float(v)
                if num is None:
                    continue
                numeric_cells.append((cell.row, cell.column, cell.coordinate, v, num))

        for r, c, addr, raw, num in numeric_cells:
            row_ctx = [t for col, t in sorted(row_text.get(r, [])) if col != c]
            col_ctx = [t for row_idx, t in sorted(col_text.get(c, [])) if row_idx < r]
            cells.append(ExcelCell(
                sheet=sheet.title,
                address=addr,
                row=r,
                column=c,
                raw_value=raw,
                numeric_value=num,
                row_context=row_ctx,
                column_context=col_ctx,
            ))
    wb.close()
    return cells


def _to_float(v) -> Optional[float]:
    # bool is a subclass of int — exclude it so TRUE/FALSE cells don't pollute
    # the numeric pool.
    if isinstance(v, bool):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    return None
