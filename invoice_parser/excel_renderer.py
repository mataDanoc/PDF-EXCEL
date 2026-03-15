"""
excel_renderer.py - Stage 5: Excel Generation

Takes ExcelGrid objects (one per PDF page) produced by grid_builder.py
and writes them into a single .xlsx file using openpyxl.

Layout decisions
----------------
- Each PDF page becomes its own Excel worksheet.
- Column widths come from ColWidthHint objects; unspecified columns receive
  the default width from config.
- Row height is uniform (config.row_height) except where cells have a
  significantly larger font.
- Bold text is preserved where detected.
- Table header rows (is_header=True) receive bold + light background.
- Totals rows (is_total=True) receive bold.
- Text wrapping is enabled for cells that contain newlines or are very wide.
- Collision handling: if two cells map to the same (row, col) they are
  concatenated with a space so no data is lost.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from .config import Config, DEFAULT_CONFIG
from .grid_builder import ExcelGrid, GridCell, ColWidthHint

logger = logging.getLogger(__name__)

try:
    import openpyxl
    from openpyxl import Workbook
    from openpyxl.styles import (
        Font, PatternFill, Alignment, Border, Side, numbers
    )
    from openpyxl.utils import get_column_letter
    _OPENPYXL_OK = True
except ImportError:
    _OPENPYXL_OK = False
    logger.error("openpyxl not installed.  Run: pip install openpyxl")


# Colour constants (ARGB)
_HEADER_BG   = "FFD9E1F2"   # light blue
_TOTAL_BG    = "FFFFF2CC"   # light yellow
_BORDER_CLR  = "FFB8CCE4"   # light border

_THIN_BORDER = None          # populated lazily when openpyxl is available


def _make_thin_border() -> "Border":
    side = Side(border_style="thin", color=_BORDER_CLR)
    return Border(left=side, right=side, top=side, bottom=side)


# ─────────────────────────────────────────────────────────────────────────────
# Renderer
# ─────────────────────────────────────────────────────────────────────────────


class ExcelRenderer:
    """Renders one or more ExcelGrid objects into an .xlsx file."""

    def __init__(self, config: Config = DEFAULT_CONFIG) -> None:
        self.cfg = config

    # ------------------------------------------------------------------ #
    # Public API
    # ------------------------------------------------------------------ #

    def render(self, grids: List[ExcelGrid], output_path: str | Path) -> Path:
        """
        Write *grids* to *output_path*.

        Parameters
        ----------
        grids       : one ExcelGrid per PDF page (may be empty list)
        output_path : destination .xlsx file path (parent must exist)

        Returns
        -------
        The resolved Path of the written file.
        """
        if not _OPENPYXL_OK:
            raise RuntimeError(
                "openpyxl is required.  Install it with: pip install openpyxl"
            )

        out = Path(output_path)
        out.parent.mkdir(parents=True, exist_ok=True)

        wb = Workbook()
        # Remove the default empty sheet
        wb.remove(wb.active)

        for grid in grids:
            self._render_page(wb, grid)

        if not wb.worksheets:
            # Safety net: create at least one sheet
            wb.create_sheet("Empty")

        wb.save(str(out))
        logger.info("Saved Excel file: %s", out)
        return out

    # ------------------------------------------------------------------ #
    # Per-page rendering
    # ------------------------------------------------------------------ #

    def _render_page(self, wb: "Workbook", grid: ExcelGrid) -> None:
        """Create one worksheet for *grid*."""
        sheet_name = f"Page {grid.page_num}"
        ws = wb.create_sheet(title=sheet_name)

        # ── Resolve collisions: merge cells mapped to the same (row, col) ──
        cell_map: Dict[Tuple[int, int], GridCell] = {}
        for gc in grid.cells:
            key = (gc.excel_row, gc.excel_col)
            if key in cell_map:
                # Concatenate text; keep formatting of first occurrence
                cell_map[key].text += " " + gc.text
            else:
                cell_map[key] = gc

        # ── Identify header and totals rows ────────────────────────────────
        header_rows: set = set()
        total_rows: set = set()
        for gc in cell_map.values():
            if gc.is_header:
                header_rows.add(gc.excel_row)
            if gc.is_total:
                total_rows.add(gc.excel_row)

        # ── Write cells ─────────────────────────────────────────────────────
        global _THIN_BORDER
        if _THIN_BORDER is None:
            _THIN_BORDER = _make_thin_border()

        for (er, ec), gc in sorted(cell_map.items()):
            ws_cell = ws.cell(row=er, column=ec, value=gc.text)

            # Font
            bold = gc.bold or er in header_rows or er in total_rows
            font_size = gc.font_size if gc.font_size and gc.font_size > 0 else 10
            ws_cell.font = Font(bold=bold, size=font_size)

            # Alignment
            ws_cell.alignment = Alignment(
                wrap_text=True,
                vertical="top",
            )

            # Background fills
            if er in header_rows:
                ws_cell.fill = PatternFill(
                    fill_type="solid", fgColor=_HEADER_BG
                )
            elif er in total_rows:
                ws_cell.fill = PatternFill(
                    fill_type="solid", fgColor=_TOTAL_BG
                )

        # ── Column widths ───────────────────────────────────────────────────
        cfg = self.cfg
        width_by_col: Dict[int, float] = {}
        for hint in grid.col_width_hints:
            width_by_col[hint.excel_col] = hint.width

        # Auto-widen based on actual text lengths for unhinged columns
        for (er, ec), gc in cell_map.items():
            if ec not in width_by_col:
                # Approximate: text length + small margin
                approx = min(cfg.max_col_width, len(gc.text) * 1.1 + 2)
                width_by_col[ec] = max(
                    width_by_col.get(ec, cfg.min_col_width), approx
                )

        for ec, w in width_by_col.items():
            col_letter = get_column_letter(ec)
            w_clamped = max(cfg.min_col_width, min(cfg.max_col_width, w))
            ws.column_dimensions[col_letter].width = w_clamped

        # Default width for unset columns
        ws.sheet_format.defaultColWidth = cfg.default_col_width

        # ── Row heights ─────────────────────────────────────────────────────
        if grid.total_rows > 0:
            for r in range(1, grid.total_rows + 1):
                ws.row_dimensions[r].height = cfg.row_height

        logger.debug(
            "Rendered sheet '%s': %d cells, %d rows, %d cols",
            sheet_name, len(cell_map), grid.total_rows, grid.total_cols,
        )
