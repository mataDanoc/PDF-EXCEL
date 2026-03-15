"""
grid_builder.py  -  Visual-fidelity grid mapping

Algorithm: Dynamic Column Boundaries + Cell Merge Spans
---------------------------------------------------------
Instead of using a fixed N-column grid, this module collects EVERY unique
x-position (x0 and x1 of every text cell) and uses them as column
boundaries.  Each adjacent pair of boundaries becomes one narrow Excel
column whose width is proportional to the PDF-point distance.

A text cell starting at boundary i and ending at boundary j is rendered as
a MERGED Excel cell spanning columns i+1 .. j.  This reproduces the
exact horizontal layout of the original PDF.

Vertical layout uses per-row heights derived from the actual PDF text
heights, with proportional spacer rows for vertical gaps.
"""

from __future__ import annotations

import bisect
import logging
import re
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .config import Config, DEFAULT_CONFIG
from .layout_detector import (
    Cell, PageLayout, TableRegion, TextRow, FilledRect, DetectedLine,
)

logger = logging.getLogger(__name__)

# PDF points -> Excel "character width" conversion factor.
# 1 Excel character width ~ 7 pixels ~ 5.25 pt at 96 DPI.
_PT_TO_CHAR = 1.0 / 5.25


# ── Data classes ──────────────────────────────────────────────────────────────

@dataclass
class GridCell:
    """A cell mapped into the Excel grid with merge-span information."""

    text: str
    excel_row: int        # 1-based
    start_col: int        # 1-based  (leftmost Excel column)
    end_col: int          # 1-based  (rightmost Excel column, inclusive)
    bold: bool = False
    font_size: float = 10.0
    bg_color: Optional[str] = None       # ARGB hex e.g. "FF1F4E79"
    text_color: Optional[str] = None     # ARGB hex for font colour
    is_header: bool = False
    is_total: bool = False
    align_h: str = "left"                # left | center | right


@dataclass
class ExcelGrid:
    """Complete grid ready for rendering by excel_renderer."""

    page_num: int
    cells: List[GridCell] = field(default_factory=list)
    col_widths_pt: List[float] = field(default_factory=list)   # width in PDF-pt per column
    row_heights_pt: List[float] = field(default_factory=list)  # height in PDF-pt per row
    total_rows: int = 0
    total_cols: int = 0
    # Rows that belong to a table (for border application)
    table_row_ranges: List[Tuple[int, int, int, int]] = field(default_factory=list)
    # (first_row, last_row, first_col, last_col) for each table


# ── Grid Builder ──────────────────────────────────────────────────────────────

class GridBuilder:
    # Boundaries closer than this (PDF-pt) are merged into one.
    MIN_COL_GAP = 3.0

    def __init__(self, config: Config = DEFAULT_CONFIG) -> None:
        self.cfg = config

    def build(self, layout: PageLayout) -> ExcelGrid:
        all_rows_sorted = sorted(layout.text_rows, key=lambda r: r.y0)
        if not all_rows_sorted:
            return ExcelGrid(page_num=layout.page_num)

        # ── Step 1: Build dynamic x-boundaries ───────────────────────────
        x_set: set[float] = {0.0, layout.page_width}
        for row in all_rows_sorted:
            for cell in row.cells:
                x_set.add(cell.x0)
                x_set.add(cell.x1)

        boundaries = self._merge_close(sorted(x_set), self.MIN_COL_GAP)
        num_cols = len(boundaries) - 1
        if num_cols < 1:
            return ExcelGrid(page_num=layout.page_num)

        # Column widths in PDF points
        col_widths = [boundaries[i + 1] - boundaries[i] for i in range(num_cols)]

        # ── Step 2: Assign Excel rows (with proportional spacers) ────────
        row_map: Dict[int, int] = {}       # id(TextRow) -> excel_row
        row_heights: List[float] = []      # excel_row (0-indexed) -> height_pt

        excel_row = 1
        prev_row: Optional[TextRow] = None

        for text_row in all_rows_sorted:
            if prev_row is not None:
                gap = text_row.y0 - prev_row.y1
                spacers, spacer_h = self._spacer_info(gap, prev_row.height)
                for _ in range(spacers):
                    row_heights.append(spacer_h)
                    excel_row += 1

            row_map[id(text_row)] = excel_row
            row_h = max(text_row.height, 10.0) * 1.15   # slight padding
            row_heights.append(row_h)
            excel_row += 1
            prev_row = text_row

        total_rows = excel_row - 1

        # ── Step 3: Build table membership sets ──────────────────────────
        table_row_ids: set = set()
        table_first_rows: Dict[int, int] = {}
        for t in layout.table_regions:
            if t.rows:
                table_first_rows[id(t)] = id(t.rows[0])
            for r in t.rows:
                table_row_ids.add(id(r))

        row_to_table: Dict[int, TableRegion] = {
            id(r): t for t in layout.table_regions for r in t.rows
        }

        # ── Step 4: Map cells to grid ────────────────────────────────────
        grid_cells: List[GridCell] = []

        for text_row in all_rows_sorted:
            er = row_map[id(text_row)]
            owning_table = row_to_table.get(id(text_row))
            is_first_table_row = (
                owning_table is not None
                and id(text_row) == table_first_rows.get(id(owning_table))
            )

            for cell in text_row.cells:
                sc = self._x_to_col(cell.x0, boundaries)
                ec = self._x_to_col_end(cell.x1, boundaries)
                ec = max(ec, sc)  # ensure at least 1 col

                # Background colour from filled rectangles
                bg_color = self._find_bg_color(
                    cell.x0, cell.y0, cell.x1, cell.y1,
                    layout.filled_rects,
                )
                text_color = None
                if bg_color:
                    # If background is dark, use white text
                    fr = self._find_rect(cell.x0, cell.y0, cell.x1, cell.y1,
                                         layout.filled_rects)
                    if fr and fr.is_dark:
                        text_color = "FFFFFFFF"

                # Alignment heuristic
                align_h = self._detect_alignment(cell, boundaries, sc, ec)

                # Totals detection
                is_total = self._is_total(cell, layout.page_width)

                grid_cells.append(GridCell(
                    text=cell.text,
                    excel_row=er,
                    start_col=sc,
                    end_col=ec,
                    bold=cell.bold or is_first_table_row,
                    font_size=cell.font_size or 10.0,
                    bg_color=bg_color,
                    text_color=text_color,
                    is_header=is_first_table_row,
                    is_total=is_total,
                    align_h=align_h,
                ))

        # ── Step 5: Table row ranges for border drawing ──────────────────
        table_ranges: List[Tuple[int, int, int, int]] = []
        for table in layout.table_regions:
            ers = [row_map.get(id(r)) for r in table.rows if id(r) in row_map]
            if not ers:
                continue
            min_row, max_row = min(ers), max(ers)
            # Find column extent of the table
            t_cols = []
            for r in table.rows:
                for c in r.cells:
                    t_cols.append(self._x_to_col(c.x0, boundaries))
                    t_cols.append(self._x_to_col_end(c.x1, boundaries))
            if t_cols:
                table_ranges.append((min_row, max_row, min(t_cols), max(t_cols)))

        return ExcelGrid(
            page_num=layout.page_num,
            cells=grid_cells,
            col_widths_pt=col_widths,
            row_heights_pt=row_heights,
            total_rows=total_rows,
            total_cols=num_cols,
            table_row_ranges=table_ranges,
        )

    # ── Helpers ───────────────────────────────────────────────────────────

    @staticmethod
    def _merge_close(sorted_vals: List[float], min_gap: float) -> List[float]:
        """Merge values closer than min_gap into one boundary."""
        if not sorted_vals:
            return []
        merged = [sorted_vals[0]]
        for v in sorted_vals[1:]:
            if v - merged[-1] >= min_gap:
                merged.append(v)
        return merged

    @staticmethod
    def _x_to_col(x: float, boundaries: List[float]) -> int:
        """Map x-coordinate to 1-based Excel column (floor boundary)."""
        idx = bisect.bisect_right(boundaries, x + 0.5) - 1
        idx = max(0, min(idx, len(boundaries) - 2))
        return idx + 1   # 1-based

    @staticmethod
    def _x_to_col_end(x: float, boundaries: List[float]) -> int:
        """Map x-coordinate to 1-based Excel column (ceiling boundary)."""
        idx = bisect.bisect_left(boundaries, x - 0.5)
        idx = max(1, min(idx, len(boundaries) - 1))
        return idx  # 1-based

    def _spacer_info(
        self, gap_pts: float, line_height: float
    ) -> Tuple[int, float]:
        """Return (count, height_pt) for spacer rows."""
        threshold = max(self.cfg.spacer_row_threshold, line_height * 0.8)
        if gap_pts <= threshold * 0.5:
            return 0, 0.0
        n = min(3, max(1, round(gap_pts / threshold)))
        h = gap_pts / n
        return n, max(h, 4.0)

    def _find_bg_color(
        self, x0: float, y0: float, x1: float, y1: float,
        rects: List[FilledRect],
    ) -> Optional[str]:
        """Find the background colour for the cell region."""
        mid_x = (x0 + x1) / 2
        mid_y = (y0 + y1) / 2
        for rect in rects:
            if rect.contains_point(mid_x, mid_y):
                return rect.hex_color
        return None

    def _find_rect(
        self, x0: float, y0: float, x1: float, y1: float,
        rects: List[FilledRect],
    ) -> Optional[FilledRect]:
        mid_x = (x0 + x1) / 2
        mid_y = (y0 + y1) / 2
        for rect in rects:
            if rect.contains_point(mid_x, mid_y):
                return rect
        return None

    def _detect_alignment(
        self, cell: Cell, boundaries: List[float],
        start_col: int, end_col: int,
    ) -> str:
        """Heuristic alignment detection."""
        # Numbers → right-aligned
        text = cell.text.strip()
        if re.match(r'^[\d\s.,\-\+%]+$', text):
            return "right"
        # Short centred labels (column headers)
        span_start = boundaries[start_col - 1] if start_col - 1 < len(boundaries) else 0
        span_end = boundaries[end_col] if end_col < len(boundaries) else boundaries[-1]
        avail_width = span_end - span_start
        text_width = cell.x1 - cell.x0
        if avail_width > 0 and text_width < avail_width * 0.7:
            left_gap = cell.x0 - span_start
            right_gap = span_end - cell.x1
            if left_gap > 5 and right_gap > 5 and abs(left_gap - right_gap) < avail_width * 0.2:
                return "center"
        return "left"

    def _is_total(self, cell: Cell, page_width: float) -> bool:
        text = cell.text.lower().strip()
        if cell.x0 < page_width * 0.40:
            return False
        triggers = {"total", "subtotal", "nentotal", "tvsh", "vat", "tax", "grand"}
        if any(t in text for t in triggers):
            return True
        return False
