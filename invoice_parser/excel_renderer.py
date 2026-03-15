"""
excel_renderer.py  -  Visual-fidelity Excel generation (v2)

Creates .xlsx that visually matches the original PDF:
  - Grid lines REMOVED (clean document look)
  - Precise column widths from PDF coordinate distances
  - Cell merging for text spanning multiple columns
  - Background colours from PDF filled rectangles
  - White text on dark backgrounds
  - Borders around table regions
  - Proportional row heights
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, List, Set, Tuple

from .config import Config, DEFAULT_CONFIG
from .grid_builder import ExcelGrid, GridCell

logger = logging.getLogger(__name__)

try:
    from openpyxl import Workbook
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side,
    )
    from openpyxl.utils import get_column_letter
    from openpyxl.worksheet.views import SheetView
    _OK = True
except ImportError:
    _OK = False
    logger.error("openpyxl not installed.  pip install openpyxl")

_PT_TO_CHAR = 1.0 / 5.25

# Minimum Excel column width in character units
_MIN_COL_CHAR = 1.5


class ExcelRenderer:
    def __init__(self, config: Config = DEFAULT_CONFIG) -> None:
        self.cfg = config

    def render(self, grids: List[ExcelGrid], output_path: str | Path) -> Path:
        if not _OK:
            raise RuntimeError("openpyxl required. pip install openpyxl")

        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)

        wb = Workbook()
        wb.remove(wb.active)

        for grid in grids:
            self._render_page(wb, grid)

        if not wb.worksheets:
            wb.create_sheet("Empty")

        wb.save(str(out))
        logger.info("Saved Excel file: %s", out)
        return out

    def _render_page(self, wb: "Workbook", grid: ExcelGrid) -> None:
        cfg = self.cfg
        ws = wb.create_sheet(title=f"Page {grid.page_num}")

        # ══════════════════════════════════════════════════════════════
        # 1. HIDE GRID LINES - multiple methods for maximum compatibility
        # ══════════════════════════════════════════════════════════════
        ws.sheet_view.showGridLines = False
        # Also set via the views collection for older Excel versions
        try:
            for sv in ws.views.sheetView:
                sv.showGridLines = False
        except Exception:
            pass

        # ══════════════════════════════════════════════════════════════
        # 2. COLUMN WIDTHS - proportional to PDF points
        # ══════════════════════════════════════════════════════════════
        for i, w_pt in enumerate(grid.col_widths_pt):
            col_letter = get_column_letter(i + 1)
            char_w = w_pt * _PT_TO_CHAR
            # Ensure minimum readable width
            char_w = max(_MIN_COL_CHAR, char_w)
            char_w = min(cfg.max_col_width, char_w)
            ws.column_dimensions[col_letter].width = char_w

        # ══════════════════════════════════════════════════════════════
        # 3. ROW HEIGHTS
        # ══════════════════════════════════════════════════════════════
        for i, h_pt in enumerate(grid.row_heights_pt):
            ws.row_dimensions[i + 1].height = max(5.0, h_pt * 0.85)

        # ══════════════════════════════════════════════════════════════
        # 4. RESOLVE CELL COLLISIONS (same row overlapping columns)
        # ══════════════════════════════════════════════════════════════
        occupied: Dict[Tuple[int, int], GridCell] = {}
        final_cells: List[GridCell] = []

        for gc in grid.cells:
            conflict = False
            for c in range(gc.start_col, gc.end_col + 1):
                if (gc.excel_row, c) in occupied:
                    occupied[(gc.excel_row, c)].text += " " + gc.text
                    conflict = True
                    break
            if not conflict:
                for c in range(gc.start_col, gc.end_col + 1):
                    occupied[(gc.excel_row, c)] = gc
                final_cells.append(gc)

        # ══════════════════════════════════════════════════════════════
        # 5. MERGE CELLS first (openpyxl requires merge before write)
        # ══════════════════════════════════════════════════════════════
        for gc in final_cells:
            if gc.end_col > gc.start_col:
                try:
                    ws.merge_cells(
                        start_row=gc.excel_row, start_column=gc.start_col,
                        end_row=gc.excel_row, end_column=gc.end_col,
                    )
                except Exception:
                    pass

        # ══════════════════════════════════════════════════════════════
        # 6. WRITE CELL CONTENT + FORMATTING
        # ══════════════════════════════════════════════════════════════
        header_rows: Set[int] = {gc.excel_row for gc in final_cells if gc.is_header}
        total_rows: Set[int] = {gc.excel_row for gc in final_cells if gc.is_total}

        # Styles cache
        _thin = Side(border_style="thin", color="FF888888")
        thin_border = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)

        for gc in final_cells:
            ws_cell = ws.cell(
                row=gc.excel_row, column=gc.start_col,
                value=gc.text,
            )

            # Font
            bold = gc.bold or gc.excel_row in header_rows or gc.excel_row in total_rows
            fs = gc.font_size if gc.font_size and gc.font_size > 0 else 10
            fc = gc.text_color or "FF000000"
            ws_cell.font = Font(bold=bold, size=fs, color=fc)

            # Alignment
            ws_cell.alignment = Alignment(
                horizontal=gc.align_h,
                vertical="center",
                wrap_text=True,
            )

            # Background fill
            if gc.bg_color:
                ws_cell.fill = PatternFill(fill_type="solid", fgColor=gc.bg_color)
            elif gc.excel_row in total_rows:
                ws_cell.fill = PatternFill(fill_type="solid", fgColor="FFF2F2F2")

        # ══════════════════════════════════════════════════════════════
        # 7. TABLE BORDERS
        # ══════════════════════════════════════════════════════════════
        for (fr, lr, fc, lc) in grid.table_row_ranges:
            for r in range(fr, lr + 1):
                for c in range(fc, lc + 1):
                    cell = ws.cell(row=r, column=c)
                    cell.border = thin_border

        # ══════════════════════════════════════════════════════════════
        # 8. PAGE SETUP for printing
        # ══════════════════════════════════════════════════════════════
        ws.page_setup.fitToWidth = 1
        ws.page_setup.fitToHeight = 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True

        logger.debug(
            "Page %d: %d cells, %d merges, %d cols x %d rows",
            grid.page_num, len(final_cells),
            sum(1 for gc in final_cells if gc.end_col > gc.start_col),
            grid.total_cols, grid.total_rows,
        )
