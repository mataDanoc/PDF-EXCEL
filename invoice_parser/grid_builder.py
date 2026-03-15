"""
grid_builder.py - Stage 4: Grid Coordinate Mapping

Converts the spatial layout analysis (PageLayout) into an abstract Excel
grid representation: a two-dimensional sparse map of

    (excel_row, excel_col) → GridCell

The mapping preserves relative positioning by:

- Assigning each distinct PDF text-row its own Excel row.
- Inserting blank spacer rows between PDF rows where the vertical gap
  is large (gap > spacer_row_threshold inserts extra blank rows).
- Mapping PDF x0 coordinates to Excel column indices using a linear
  scale tied to the page width.
- Optionally using detected column boundaries (for table regions) to
  produce exact column indices rather than scaled approximations.

The resulting ExcelGrid object is consumed by excel_renderer.py.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .config import Config, DEFAULT_CONFIG
from .layout_detector import (
    Cell, ColumnBoundary, PageLayout, TableRegion, TextRow, RegionKind,
    DocumentRegion,
)

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Output data classes
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class GridCell:
    """A positioned cell in the Excel grid."""

    text: str
    excel_row: int       # 1-based
    excel_col: int       # 1-based
    bold: bool = False
    font_size: Optional[float] = None
    is_header: bool = False      # first row of a detected table
    is_total: bool = False       # heuristically detected totals row
    col_span: int = 1            # merge across this many columns


@dataclass
class ColWidthHint:
    """Suggested Excel column width derived from PDF coordinate spacing."""

    excel_col: int
    width: float     # Excel column width units (≈ characters wide)


@dataclass
class ExcelGrid:
    """The complete grid representation for one PDF page."""

    page_num: int
    cells: List[GridCell] = field(default_factory=list)
    col_width_hints: List[ColWidthHint] = field(default_factory=list)
    total_rows: int = 0
    total_cols: int = 0


# ─────────────────────────────────────────────────────────────────────────────
# Grid Builder
# ─────────────────────────────────────────────────────────────────────────────


class GridBuilder:
    """
    Converts a PageLayout into an ExcelGrid.

    Design decisions
    ----------------
    * Every PDF text-row maps to exactly one Excel row.
    * Blank spacer rows are inserted when the vertical gap between
      consecutive PDF rows exceeds `spacer_row_threshold`.
    * For table regions the column mapping uses the detected column
      boundaries (stable, aligned).
    * For non-table regions the column mapping uses linear scaling based
      on the PDF page width and `excel_grid_cols`.
    """

    # Fraction of page width that a row's single cell must span to be
    # treated as a full-width header (and optionally merged).
    FULL_WIDTH_FRACTION = 0.70

    def __init__(self, config: Config = DEFAULT_CONFIG) -> None:
        self.cfg = config

    # ------------------------------------------------------------------ #
    # Public API
    # ------------------------------------------------------------------ #

    def build(self, layout: PageLayout) -> ExcelGrid:
        """Convert *layout* to an ExcelGrid ready for rendering."""

        grid = ExcelGrid(page_num=layout.page_num)

        # Assign an excel_row to every TextRow in reading order
        row_map: Dict[int, int] = {}   # id(TextRow) → excel_row (1-based)
        excel_row = 1

        # We need the PDF rows in top-to-bottom order
        all_rows_sorted: List[TextRow] = sorted(
            layout.text_rows, key=lambda r: r.y0
        )

        prev_row: Optional[TextRow] = None
        for text_row in all_rows_sorted:
            if prev_row is not None:
                gap = text_row.y0 - prev_row.y1
                spacers = self._spacer_count(gap, prev_row.height)
                excel_row += spacers

            row_map[id(text_row)] = excel_row
            excel_row += 1
            prev_row = text_row

        grid.total_rows = excel_row - 1

        # Scale factor: PDF x → Excel column (for non-table text)
        scale = self.cfg.excel_grid_cols / max(layout.page_width, 1.0)

        # Build a lookup: id(TextRow) → TableRegion (if any)
        row_to_table: Dict[int, TableRegion] = {
            id(r): t
            for t in layout.table_regions
            for r in t.rows
        }

        # Track which column boundaries each table uses so we can compute
        # a global column-boundary → excel_col mapping.
        # For simplicity we assign table columns sequentially starting at
        # the first scaled column of the table's leftmost boundary.
        table_col_maps: Dict[int, Dict[int, int]] = {}  # id(table) → {bound_idx → excel_col}

        for table in layout.table_regions:
            if not table.col_boundaries:
                continue
            # Use the x position of the first boundary to anchor the table
            anchor_col = max(1, round(table.col_boundaries[0].x * scale))
            col_map: Dict[int, int] = {}
            for i, boundary in enumerate(table.col_boundaries):
                col_map[i] = anchor_col + i
            table_col_maps[id(table)] = col_map

        # ---------------------------------------------------------------- #
        # Emit GridCells
        # ---------------------------------------------------------------- #

        all_excel_cols: List[int] = []
        table_first_rows: Dict[int, int] = {
            id(t): id(t.rows[0]) for t in layout.table_regions if t.rows
        }

        for text_row in all_rows_sorted:
            er = row_map[id(text_row)]
            owning_table = row_to_table.get(id(text_row))

            is_table_first_row = (
                owning_table is not None
                and id(text_row) == table_first_rows.get(id(owning_table))
            )

            for cell_idx, cell in enumerate(text_row.cells):

                if owning_table is not None:
                    # Table region: use column-boundary-based mapping
                    col_map = table_col_maps.get(id(owning_table), {})
                    ec = self._map_cell_to_table_col(
                        cell, owning_table.col_boundaries, col_map, scale
                    )
                else:
                    # Free text: linear scale
                    ec = max(1, round(cell.x0 * scale) + 1)

                is_total = self._is_totals_cell(cell, text_row, layout.page_width)

                grid_cell = GridCell(
                    text=cell.text,
                    excel_row=er,
                    excel_col=ec,
                    bold=cell.bold,
                    font_size=cell.font_size,
                    is_header=is_table_first_row,
                    is_total=is_total,
                )
                grid.cells.append(grid_cell)
                all_excel_cols.append(ec)

        grid.total_cols = max(all_excel_cols, default=1)

        # ---------------------------------------------------------------- #
        # Column width hints
        # ---------------------------------------------------------------- #
        grid.col_width_hints = self._compute_col_widths(
            layout, scale, table_col_maps
        )

        return grid

    # ------------------------------------------------------------------ #
    # Private helpers
    # ------------------------------------------------------------------ #

    def _spacer_count(self, gap_pts: float, avg_line_height: float) -> int:
        """
        Return the number of blank Excel rows to insert for a given
        vertical gap in PDF points.
        """
        threshold = max(self.cfg.spacer_row_threshold, avg_line_height * 0.5)
        if gap_pts <= threshold:
            return 0
        # One spacer per threshold worth of gap, capped at 4
        return min(4, int(gap_pts / threshold))

    def _map_cell_to_table_col(
        self,
        cell: Cell,
        boundaries: List[ColumnBoundary],
        col_map: Dict[int, int],
        scale: float,
    ) -> int:
        """Return the Excel column index for *cell* within a table region."""
        if not boundaries:
            return max(1, round(cell.x0 * scale) + 1)

        tolerance = self.cfg.column_tolerance * 2
        best_idx = min(
            range(len(boundaries)),
            key=lambda k: abs(boundaries[k].x - cell.x0),
        )

        if abs(boundaries[best_idx].x - cell.x0) <= tolerance:
            return col_map.get(best_idx, max(1, round(cell.x0 * scale) + 1))

        # Fallback: linear scale
        return max(1, round(cell.x0 * scale) + 1)

    def _is_totals_cell(
        self, cell: Cell, row: TextRow, page_width: float
    ) -> bool:
        """
        Heuristic: mark a cell as a "totals" cell when the row's content
        is right-aligned (cells clustered near the right margin) AND the
        text contains a currency pattern.
        """
        import re
        cfg = self.cfg
        # Must contain a digit (amount) and be right of 50% of page
        if not re.search(r"\d", cell.text):
            return False
        if cell.x0 < page_width * 0.50:
            return False
        if any(sym in cell.text for sym in cfg.currency_symbols):
            return True
        # Numeric-only cell far right
        if re.match(r"^[\d\s.,\-\+]+$", cell.text.strip()):
            return True
        return False

    def _compute_col_widths(
        self,
        layout: PageLayout,
        scale: float,
        table_col_maps: Dict[int, Dict[int, int]],
    ) -> List[ColWidthHint]:
        """
        Derive suggested Excel column widths from PDF coordinate spacing.

        For table columns: width is proportional to the space between
        consecutive column boundaries.
        For other columns: use the default width.
        """
        cfg = self.cfg
        hints: Dict[int, float] = {}

        for table in layout.table_regions:
            col_map = table_col_maps.get(id(table), {})
            boundaries = table.col_boundaries
            for i, boundary in enumerate(boundaries):
                ec = col_map.get(i, round(boundary.x * scale) + 1)

                if i + 1 < len(boundaries):
                    pdf_width = boundaries[i + 1].x - boundary.x
                else:
                    pdf_width = layout.page_width - boundary.x

                # Convert PDF points to Excel character width
                # Excel "character width" ≈ 7 pixels ≈ 5.25 pt
                char_width = max(
                    cfg.min_col_width,
                    min(cfg.max_col_width, pdf_width / 7.0),
                )
                hints[ec] = max(hints.get(ec, 0.0), char_width)

        return [ColWidthHint(excel_col=ec, width=w) for ec, w in sorted(hints.items())]
