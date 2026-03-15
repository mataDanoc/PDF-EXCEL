"""
layout_detector.py - Stage 3: Spatial Layout Analysis

This is the algorithmic core of the pipeline.  It takes a list of Word
objects (from pdf_loader / ocr_engine) and transforms them into a
structured layout representation:

    PageLayout
    ├── List[TextRow]          – words grouped into horizontal text rows
    ├── List[ColumnBoundary]   – detected vertical column positions
    ├── List[TableRegion]      – detected structured table regions
    └── List[DocumentRegion]   – semantic sections (header / body / footer)

Algorithm overview
------------------
A. Row Detection
   Words are sorted by their vertical midpoint.  A greedy scan groups
   words whose midpoints lie within `row_tolerance` of the current row's
   running-average midpoint.  This tolerates slight baseline variations
   common in justified text.

B. Cell Merging
   Within each row, adjacent words are merged into a single "cell" when
   the horizontal gap between them is smaller than `cell_merge_gap`.
   This correctly handles e.g. "Invoice Number:" as one cell rather than
   two separate words.

C. Column Detection
   For every row we record each cell's x0 position.  All x0 values are
   clustered by proximity (within `column_tolerance`) to produce a list
   of column boundary positions that represent recurring vertical
   alignment patterns in the document.

D. Table Detection
   A contiguous block of TextRows is classified as a table when a minimum
   fraction of those rows use the same column boundaries, i.e. when they
   are "column-consistent".  Additionally, if the PDF contains vector
   lines that form a grid, those regions are unconditionally treated as
   tables.

E. Region Segmentation
   The page is split into header (top `header_fraction`), footer (bottom
   `footer_fraction`), and body based on Y coordinates.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from typing import List, Optional, Tuple, Dict

from .config import Config, DEFAULT_CONFIG
from .pdf_loader import Word, DetectedLine, FilledRect, PageData

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Output data classes
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class Cell:
    """One logical cell within a text row (may span multiple PDF words)."""

    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    bold: bool = False
    font_size: Optional[float] = None

    @property
    def width(self) -> float:
        return self.x1 - self.x0

    @property
    def mid_x(self) -> float:
        return (self.x0 + self.x1) / 2.0


@dataclass
class TextRow:
    """A horizontal row of Cell objects sharing approximately the same Y."""

    cells: List[Cell] = field(default_factory=list)
    y0: float = 0.0    # top of row (minimum y0 of all cells)
    y1: float = 0.0    # bottom of row

    @property
    def mid_y(self) -> float:
        return (self.y0 + self.y1) / 2.0

    @property
    def height(self) -> float:
        return self.y1 - self.y0

    def col_count(self) -> int:
        return len(self.cells)


@dataclass
class ColumnBoundary:
    """A detected vertical column position (x coordinate in PDF points)."""

    x: float            # representative x position
    members: List[float] = field(default_factory=list)   # contributing x0 values

    def index_of(self, x: float, tolerance: float) -> bool:
        return abs(self.x - x) <= tolerance


@dataclass
class TableRegion:
    """A contiguous block of TextRows identified as a structured table."""

    rows: List[TextRow] = field(default_factory=list)
    col_boundaries: List[ColumnBoundary] = field(default_factory=list)
    has_line_borders: bool = False   # True if vector lines delimit cells

    @property
    def row_count(self) -> int:
        return len(self.rows)

    @property
    def col_count(self) -> int:
        return len(self.col_boundaries)


class RegionKind:
    HEADER = "header"
    BODY = "body"
    FOOTER = "footer"
    TABLE = "table"


@dataclass
class DocumentRegion:
    """A semantic section of the page."""

    kind: str
    rows: List[TextRow] = field(default_factory=list)
    table: Optional[TableRegion] = None


@dataclass
class PageLayout:
    """Complete layout analysis result for one page."""

    page_num: int
    page_width: float
    page_height: float
    text_rows: List[TextRow] = field(default_factory=list)
    col_boundaries: List[ColumnBoundary] = field(default_factory=list)
    table_regions: List[TableRegion] = field(default_factory=list)
    regions: List[DocumentRegion] = field(default_factory=list)
    filled_rects: List[FilledRect] = field(default_factory=list)
    lines: List[DetectedLine] = field(default_factory=list)


# ─────────────────────────────────────────────────────────────────────────────
# Layout Detector
# ─────────────────────────────────────────────────────────────────────────────


class LayoutDetector:
    """Analyses a PageData and returns a structured PageLayout."""

    def __init__(self, config: Config = DEFAULT_CONFIG) -> None:
        self.cfg = config

    # ------------------------------------------------------------------ #
    # Public API
    # ------------------------------------------------------------------ #

    def analyse(self, page: PageData) -> PageLayout:
        """Full layout analysis for a single page."""

        # A. Cluster words into rows
        text_rows = self._detect_rows(page.words)

        # B. Detect global column boundaries
        col_boundaries = self._detect_columns(text_rows)

        # C. Detect table regions (line-based first, then heuristic)
        table_regions = self._detect_tables(
            text_rows, col_boundaries, page.lines, page.width
        )

        # D. Segment page into semantic regions
        regions = self._segment_regions(
            text_rows, table_regions, page.height
        )

        return PageLayout(
            page_num=page.page_num,
            page_width=page.width,
            page_height=page.height,
            text_rows=text_rows,
            col_boundaries=col_boundaries,
            table_regions=table_regions,
            regions=regions,
            filled_rects=page.filled_rects,
            lines=page.lines,
        )

    # ------------------------------------------------------------------ #
    # A. Row Detection
    # ------------------------------------------------------------------ #

    def _detect_rows(self, words: List[Word]) -> List[TextRow]:
        """
        Group words into horizontal rows using Y-midpoint clustering.

        Words are sorted by (mid_y, x0).  A greedy scan accumulates words
        into the current row as long as their mid_y is within
        `row_tolerance` of the row's running-average mid_y.
        """
        if not words:
            return []

        sorted_words = sorted(words, key=lambda w: (w.mid_y, w.x0))
        tolerance = self.cfg.row_tolerance

        rows_raw: List[List[Word]] = []
        current: List[Word] = [sorted_words[0]]
        current_avg_y = sorted_words[0].mid_y

        for w in sorted_words[1:]:
            if abs(w.mid_y - current_avg_y) <= tolerance:
                current.append(w)
                # Update running average
                current_avg_y = sum(cw.mid_y for cw in current) / len(current)
            else:
                rows_raw.append(current)
                current = [w]
                current_avg_y = w.mid_y

        rows_raw.append(current)

        # Convert raw word lists to TextRow objects with cells
        text_rows: List[TextRow] = []
        for raw_row in rows_raw:
            row_sorted = sorted(raw_row, key=lambda w: w.x0)
            cells = self._merge_words_into_cells(row_sorted)
            if not cells:
                continue
            row_obj = TextRow(
                cells=cells,
                y0=min(w.y0 for w in raw_row),
                y1=max(w.y1 for w in raw_row),
            )
            text_rows.append(row_obj)

        # Sort rows top-to-bottom
        text_rows.sort(key=lambda r: r.y0)
        logger.debug("Detected %d text rows", len(text_rows))
        return text_rows

    def _merge_words_into_cells(self, sorted_words: List[Word]) -> List[Cell]:
        """
        Merge consecutive words within a row into Cell objects.

        Words are merged when the horizontal gap between them is smaller
        than `cell_merge_gap`.  This handles phrases like "Invoice Number:"
        or "Unit Price" as single cells.
        """
        if not sorted_words:
            return []

        gap = self.cfg.cell_merge_gap
        cells: List[Cell] = []

        # Start first cell
        group: List[Word] = [sorted_words[0]]

        for w in sorted_words[1:]:
            prev = group[-1]
            if w.x0 - prev.x1 < gap:
                group.append(w)
            else:
                cells.append(self._words_to_cell(group))
                group = [w]

        cells.append(self._words_to_cell(group))
        return cells

    @staticmethod
    def _words_to_cell(words: List[Word]) -> Cell:
        text = " ".join(w.text for w in words)
        bold = any(w.bold for w in words)
        sizes = [w.font_size for w in words if w.font_size]
        font_size = max(sizes) if sizes else None
        return Cell(
            text=text,
            x0=words[0].x0,
            y0=min(w.y0 for w in words),
            x1=words[-1].x1,
            y1=max(w.y1 for w in words),
            bold=bold,
            font_size=font_size,
        )

    # ------------------------------------------------------------------ #
    # B. Column Detection
    # ------------------------------------------------------------------ #

    def _detect_columns(
        self, text_rows: List[TextRow]
    ) -> List[ColumnBoundary]:
        """
        Find recurring vertical column positions across all rows.

        The algorithm collects every cell's x0 value, sorts them, and
        greedily clusters values within `column_tolerance` of each other.
        The resulting clusters represent stable column boundaries.
        """
        tolerance = self.cfg.column_tolerance

        all_x0: List[float] = [
            cell.x0 for row in text_rows for cell in row.cells
        ]
        if not all_x0:
            return []

        sorted_x0 = sorted(all_x0)
        clusters: List[List[float]] = [[sorted_x0[0]]]

        for x in sorted_x0[1:]:
            if x - clusters[-1][-1] <= tolerance:
                clusters[-1].append(x)
            else:
                clusters.append([x])

        boundaries: List[ColumnBoundary] = []
        for cluster in clusters:
            centroid = sum(cluster) / len(cluster)
            boundaries.append(ColumnBoundary(x=centroid, members=cluster))

        logger.debug("Detected %d column boundaries", len(boundaries))
        return boundaries

    # ------------------------------------------------------------------ #
    # C. Table Detection
    # ------------------------------------------------------------------ #

    def _detect_tables(
        self,
        text_rows: List[TextRow],
        global_cols: List[ColumnBoundary],
        lines: List[DetectedLine],
        page_width: float,
    ) -> List[TableRegion]:
        """
        Identify table regions using two methods:

        1. Line-based: if the PDF has horizontal + vertical vector lines
           that form a grid, the enclosed text rows are a table.
        2. Consistency-based: a contiguous block of rows that share the
           same column alignment pattern is classified as a table.
        """
        tables: List[TableRegion] = []

        # Method 1 – line-based table detection
        line_tables = self._detect_line_tables(text_rows, lines, page_width)
        tables.extend(line_tables)

        # Method 2 – column-consistency heuristic for rows not already in a
        # line-detected table
        covered_rows = {id(r) for t in tables for r in t.rows}
        remaining = [r for r in text_rows if id(r) not in covered_rows]
        heuristic_tables = self._detect_heuristic_tables(remaining)
        tables.extend(heuristic_tables)

        logger.debug("Detected %d table region(s)", len(tables))
        return tables

    def _detect_line_tables(
        self,
        text_rows: List[TextRow],
        lines: List[DetectedLine],
        page_width: float,
    ) -> List[TableRegion]:
        """
        Detect tables formed by vector lines.

        A region bounded by at least 2 horizontal lines and spanning most
        of the page width is treated as a table.
        """
        cfg = self.cfg
        h_lines = [l for l in lines if l.is_horizontal and l.length > page_width * 0.2]
        if len(h_lines) < 2:
            return []

        # Sort horizontal lines by Y
        h_lines.sort(key=lambda l: l.y0)

        # Find bands between consecutive horizontal lines
        tables: List[TableRegion] = []
        for i in range(len(h_lines) - 1):
            top_y = h_lines[i].y0
            bot_y = h_lines[i + 1].y0

            # Collect rows in this band
            band_rows = [
                r for r in text_rows
                if r.y0 >= top_y - cfg.row_tolerance
                and r.y1 <= bot_y + cfg.row_tolerance
            ]
            if not band_rows:
                continue

            # Merge with subsequent bands (greedy: build one table per
            # contiguous set of banded rows)
            # (simplified: each band-pair becomes a table candidate, merge later)

        # Simpler approach: find contiguous sets of rows bounded by h-lines
        if len(h_lines) >= 2:
            top_y = h_lines[0].y0
            bot_y = h_lines[-1].y0

            band_rows = [
                r for r in text_rows
                if r.y0 >= top_y - cfg.row_tolerance
                and r.y1 <= bot_y + cfg.row_tolerance
            ]
            if len(band_rows) >= cfg.table_min_rows:
                col_bounds = self._detect_columns(band_rows)
                if len(col_bounds) >= cfg.table_min_cols:
                    tables.append(TableRegion(
                        rows=band_rows,
                        col_boundaries=col_bounds,
                        has_line_borders=True,
                    ))

        return tables

    def _detect_heuristic_tables(
        self, text_rows: List[TextRow]
    ) -> List[TableRegion]:
        """
        Detect tables by column alignment consistency.

        Scans rows sequentially.  When a contiguous block of rows shows
        consistent column alignment (i.e. their cells' x0 values map to the
        same column boundaries), the block is declared a table.
        """
        cfg = self.cfg
        if not text_rows:
            return []

        tables: List[TableRegion] = []
        i = 0

        while i < len(text_rows):
            # Start a candidate table block
            candidate: List[TextRow] = [text_rows[i]]
            candidate_cols = self._detect_columns(candidate)

            j = i + 1
            while j < len(text_rows):
                # Check vertical gap
                gap = text_rows[j].y0 - candidate[-1].y1
                if gap > cfg.table_max_row_gap:
                    break

                # Check column consistency
                next_row_cols = self._row_col_indices(
                    text_rows[j], candidate_cols
                )
                consistency = self._col_consistency(
                    text_rows[j], candidate_cols
                )
                if consistency >= cfg.table_consistency_threshold:
                    candidate.append(text_rows[j])
                    j += 1
                else:
                    break

            if (
                len(candidate) >= cfg.table_min_rows
                and len(candidate_cols) >= cfg.table_min_cols
            ):
                tables.append(TableRegion(
                    rows=candidate,
                    col_boundaries=candidate_cols,
                    has_line_borders=False,
                ))
                i = j
            else:
                i += 1

        return tables

    def _row_col_indices(
        self, row: TextRow, boundaries: List[ColumnBoundary]
    ) -> List[int]:
        """Assign each cell in *row* to the nearest column boundary index."""
        tolerance = self.cfg.column_tolerance
        indices: List[int] = []
        for cell in row.cells:
            nearest = min(
                range(len(boundaries)),
                key=lambda k: abs(boundaries[k].x - cell.x0),
            )
            if abs(boundaries[nearest].x - cell.x0) <= tolerance * 2:
                indices.append(nearest)
        return indices

    def _col_consistency(
        self, row: TextRow, boundaries: List[ColumnBoundary]
    ) -> float:
        """
        Return the fraction of cells in *row* that align with any known
        column boundary within `column_tolerance * 2`.
        """
        if not row.cells or not boundaries:
            return 0.0
        tolerance = self.cfg.column_tolerance * 2
        aligned = sum(
            1 for cell in row.cells
            if any(abs(b.x - cell.x0) <= tolerance for b in boundaries)
        )
        return aligned / len(row.cells)

    # ------------------------------------------------------------------ #
    # D. Region Segmentation
    # ------------------------------------------------------------------ #

    def _segment_regions(
        self,
        text_rows: List[TextRow],
        table_regions: List[TableRegion],
        page_height: float,
    ) -> List[DocumentRegion]:
        """
        Split the page into header / body / footer based on Y position.

        Rows that belong to a detected table are wrapped in a TABLE region.
        """
        cfg = self.cfg
        header_y = page_height * cfg.header_fraction
        footer_y = page_height * (1.0 - cfg.footer_fraction)

        table_row_ids = {id(r) for t in table_regions for r in t.rows}
        table_start_ys = {min(r.y0 for r in t.rows): t for t in table_regions}

        regions: List[DocumentRegion] = []
        pending_rows: List[TextRow] = []
        current_kind: Optional[str] = None

        def flush(kind: str) -> None:
            if pending_rows:
                regions.append(DocumentRegion(kind=kind, rows=list(pending_rows)))
                pending_rows.clear()

        emitted_tables: set = set()

        for row in text_rows:
            # Determine semantic kind
            if row.y0 < header_y:
                kind = RegionKind.HEADER
            elif row.y0 > footer_y:
                kind = RegionKind.FOOTER
            else:
                kind = RegionKind.BODY

            if id(row) in table_row_ids:
                # Find the table this row belongs to
                owning_table = next(
                    (t for t in table_regions if id(row) in {id(r) for r in t.rows}),
                    None,
                )
                if owning_table and id(owning_table) not in emitted_tables:
                    flush(current_kind or kind)
                    regions.append(DocumentRegion(
                        kind=RegionKind.TABLE,
                        rows=owning_table.rows,
                        table=owning_table,
                    ))
                    emitted_tables.add(id(owning_table))
                continue

            if kind != current_kind:
                flush(current_kind or kind)
                current_kind = kind

            pending_rows.append(row)

        flush(current_kind or RegionKind.BODY)
        return regions
