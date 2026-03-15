"""
grid_builder.py  -  Visual-fidelity grid mapping  (v2)

Algorithm: Clustered Column Boundaries + Cell Merge Spans
----------------------------------------------------------
1. Collect all x0 positions from every text cell across the page.
2. CLUSTER nearby x0 values (within ~15 pt) to find the NATURAL column
   structure of the document (usually 8-15 columns for an invoice).
3. Add the right edge of the page as a final boundary.
4. Each cluster centroid becomes one Excel column boundary.
5. A text cell is placed at the column whose boundary is nearest to its
   x0, and MERGED rightward until the column whose boundary is nearest
   to its x1.
6. Column widths are proportional to the PDF-point distance between
   consecutive boundaries.

This produces a COMPACT column set (10-20 columns) that mirrors the
natural visual structure of the original document.
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

_PT_TO_CHAR = 1.0 / 5.25


# ── Data classes ──────────────────────────────────────────────────────────────

@dataclass
class GridCell:
    text: str
    excel_row: int
    start_col: int        # 1-based
    end_col: int          # 1-based (inclusive)
    bold: bool = False
    font_size: float = 10.0
    bg_color: Optional[str] = None
    text_color: Optional[str] = None
    is_header: bool = False
    is_total: bool = False
    align_h: str = "left"


@dataclass
class ExcelGrid:
    page_num: int
    cells: List[GridCell] = field(default_factory=list)
    col_widths_pt: List[float] = field(default_factory=list)
    row_heights_pt: List[float] = field(default_factory=list)
    total_rows: int = 0
    total_cols: int = 0
    table_row_ranges: List[Tuple[int, int, int, int]] = field(default_factory=list)


# ── Grid Builder ──────────────────────────────────────────────────────────────

class GridBuilder:
    # Minimum gap between column boundaries in PDF points.
    # x0 values closer than this are merged into one column.
    COL_CLUSTER_GAP = 15.0

    # Maximum desired number of columns. If exceeded, increase gap.
    MAX_COLS = 30

    def __init__(self, config: Config = DEFAULT_CONFIG) -> None:
        self.cfg = config

    def build(self, layout: PageLayout) -> ExcelGrid:
        all_rows = sorted(layout.text_rows, key=lambda r: r.y0)
        if not all_rows:
            return ExcelGrid(page_num=layout.page_num)

        # ── Step 1: Cluster x0 positions into column boundaries ──────
        all_x0: List[float] = []
        all_x1: List[float] = []
        for row in all_rows:
            for cell in row.cells:
                all_x0.append(cell.x0)
                all_x1.append(cell.x1)

        boundaries = self._build_boundaries(
            all_x0, all_x1, layout.page_width
        )
        num_cols = len(boundaries) - 1
        if num_cols < 1:
            return ExcelGrid(page_num=layout.page_num)

        col_widths = [boundaries[i + 1] - boundaries[i] for i in range(num_cols)]

        logger.debug("Built %d column boundaries for page %d", num_cols, layout.page_num)

        # ── Step 2: Assign Excel rows with spacers ───────────────────
        row_map: Dict[int, int] = {}
        row_heights: List[float] = []
        excel_row = 1
        prev_row: Optional[TextRow] = None

        for text_row in all_rows:
            if prev_row is not None:
                gap = text_row.y0 - prev_row.y1
                spacers, spacer_h = self._spacer_info(gap, prev_row.height)
                for _ in range(spacers):
                    row_heights.append(spacer_h)
                    excel_row += 1

            row_map[id(text_row)] = excel_row
            row_h = max(text_row.height, 10.0) * 1.15
            row_heights.append(row_h)
            excel_row += 1
            prev_row = text_row

        total_rows = excel_row - 1

        # ── Step 3: Table membership lookup ──────────────────────────
        table_first_rows: Dict[int, int] = {}
        for t in layout.table_regions:
            if t.rows:
                table_first_rows[id(t)] = id(t.rows[0])

        row_to_table: Dict[int, TableRegion] = {
            id(r): t for t in layout.table_regions for r in t.rows
        }

        # ── Step 4: Map cells to grid ────────────────────────────────
        # Process each row: compute start_col for ALL cells first, then
        # derive end_col so that no two cells on the same row overlap.
        grid_cells: List[GridCell] = []

        for text_row in all_rows:
            er = row_map[id(text_row)]
            owning_table = row_to_table.get(id(text_row))
            is_first = (
                owning_table is not None
                and id(text_row) == table_first_rows.get(id(owning_table))
            )

            # First pass: compute start_col for every cell in this row
            row_positioned: List[Tuple[Cell, int]] = []
            for cell in text_row.cells:
                sc = self._snap_start(cell.x0, boundaries)
                row_positioned.append((cell, sc))

            # Sort by start_col (left to right)
            row_positioned.sort(key=lambda x: x[1])

            # Second pass: compute end_col, capped before the next cell
            for i, (cell, sc) in enumerate(row_positioned):
                # Natural end column from the text's right edge
                ec = self._snap_end(cell.x1, boundaries, sc)

                # CAP: do not extend into the next cell's start column
                if i + 1 < len(row_positioned):
                    next_sc = row_positioned[i + 1][1]
                    ec = min(ec, next_sc - 1)

                ec = max(ec, sc)  # at least 1 column wide

                bg_color = self._find_bg(cell, layout.filled_rects)
                text_color = None
                if bg_color:
                    fr = self._find_rect(cell, layout.filled_rects)
                    if fr and fr.is_dark:
                        text_color = "FFFFFFFF"

                align_h = self._detect_align(cell, boundaries, sc, ec)
                is_total = self._is_total(cell, layout.page_width)

                grid_cells.append(GridCell(
                    text=cell.text,
                    excel_row=er,
                    start_col=sc,
                    end_col=ec,
                    bold=cell.bold or is_first,
                    font_size=cell.font_size or 10.0,
                    bg_color=bg_color,
                    text_color=text_color,
                    is_header=is_first,
                    is_total=is_total,
                    align_h=align_h,
                ))

        # ── Step 5: Table ranges for borders ─────────────────────────
        table_ranges: List[Tuple[int, int, int, int]] = []
        for table in layout.table_regions:
            ers = [row_map[id(r)] for r in table.rows if id(r) in row_map]
            if not ers:
                continue
            min_r, max_r = min(ers), max(ers)
            t_cols = []
            for r in table.rows:
                for c in r.cells:
                    t_cols.append(self._snap_start(c.x0, boundaries))
                    t_cols.append(self._snap_end(c.x1, boundaries, 1))
            if t_cols:
                table_ranges.append((min_r, max_r, min(t_cols), max(t_cols)))

        return ExcelGrid(
            page_num=layout.page_num,
            cells=grid_cells,
            col_widths_pt=col_widths,
            row_heights_pt=row_heights,
            total_rows=total_rows,
            total_cols=num_cols,
            table_row_ranges=table_ranges,
        )

    # ── Column boundary construction ─────────────────────────────────

    def _build_boundaries(
        self,
        all_x0: List[float],
        all_x1: List[float],
        page_width: float,
    ) -> List[float]:
        """
        Build column boundaries by clustering all x0 positions.

        Strategy:
        1. Cluster x0 values with greedy merge (gap < COL_CLUSTER_GAP).
        2. Take the MINIMUM of each cluster as the boundary (left edge).
        3. Ensure the page's right edge is included.
        4. If too many columns, increase the gap and retry.
        """
        gap = self.COL_CLUSTER_GAP

        # Adaptive: increase gap if too many columns
        for _ in range(5):
            raw = sorted(set(all_x0))
            clusters: List[List[float]] = [[raw[0]]] if raw else []
            for x in raw[1:]:
                if x - clusters[-1][-1] <= gap:
                    clusters[-1].append(x)
                else:
                    clusters.append([x])

            # Boundary = minimum (left edge) of each cluster
            bounds = [min(c) for c in clusters]

            # Add right edge: max of all x1 values or page width
            right_edge = max(max(all_x1, default=page_width), page_width)
            if not bounds or right_edge - bounds[-1] > gap:
                bounds.append(right_edge)
            else:
                bounds[-1] = max(bounds[-1], right_edge)

            # Ensure left margin is included
            if bounds[0] > 5:
                bounds.insert(0, 0.0)

            if len(bounds) - 1 <= self.MAX_COLS:
                break
            gap *= 1.4  # increase gap and retry

        logger.debug(
            "Column boundaries (%d): gap=%.1f, bounds=%s",
            len(bounds), gap, [round(b, 1) for b in bounds],
        )
        return bounds

    # ── Snap helpers ─────────────────────────────────────────────────

    @staticmethod
    def _snap_start(x: float, boundaries: List[float]) -> int:
        """Find the 1-based column index for a cell's left edge."""
        idx = bisect.bisect_right(boundaries, x + 1.0) - 1
        idx = max(0, min(idx, len(boundaries) - 2))
        return idx + 1

    @staticmethod
    def _snap_end(x: float, boundaries: List[float], min_col: int) -> int:
        """Find the 1-based column index for a cell's right edge."""
        idx = bisect.bisect_left(boundaries, x - 1.0)
        idx = max(1, min(idx, len(boundaries) - 1))
        col = idx  # 1-based
        return max(col, min_col)

    # ── Spacer logic ─────────────────────────────────────────────────

    def _spacer_info(self, gap_pts: float, line_h: float) -> Tuple[int, float]:
        threshold = max(self.cfg.spacer_row_threshold, line_h * 0.8)
        if gap_pts <= threshold * 0.5:
            return 0, 0.0
        n = min(3, max(1, round(gap_pts / threshold)))
        return n, max(gap_pts / n, 4.0)

    # ── Background / colour helpers ──────────────────────────────────

    @staticmethod
    def _find_bg(cell: Cell, rects: List[FilledRect]) -> Optional[str]:
        mx, my = (cell.x0 + cell.x1) / 2, (cell.y0 + cell.y1) / 2
        for r in rects:
            if r.contains_point(mx, my):
                return r.hex_color
        return None

    @staticmethod
    def _find_rect(cell: Cell, rects: List[FilledRect]) -> Optional[FilledRect]:
        mx, my = (cell.x0 + cell.x1) / 2, (cell.y0 + cell.y1) / 2
        for r in rects:
            if r.contains_point(mx, my):
                return r
        return None

    # ── Alignment / totals heuristics ────────────────────────────────

    def _detect_align(
        self, cell: Cell, boundaries: List[float],
        sc: int, ec: int,
    ) -> str:
        text = cell.text.strip()
        if re.match(r'^[\d\s.,\-\+%]+$', text):
            return "right"
        if re.match(r'^[\d\s.,\-\+%]+\s*(ALL|EUR|USD|LEK)$', text, re.IGNORECASE):
            return "right"
        return "left"

    def _is_total(self, cell: Cell, page_width: float) -> bool:
        text = cell.text.lower().strip()
        if cell.x0 < page_width * 0.35:
            return False
        triggers = {"total", "subtotal", "nentotal", "tvsh", "vat", "tax",
                     "grand", "paguar", "nentotal"}
        return any(t in text for t in triggers)
