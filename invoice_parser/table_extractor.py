"""
table_extractor.py - Table Structure Reconstruction

Refines raw TableRegion objects produced by layout_detector into
well-structured tables with named columns.

The extractor:
1. Identifies column headers (first row of a table, or the row whose cells
   are all bold / contain known invoice column keywords).
2. Assigns a column index to every cell in every row, handling cells that
   span multiple detected column positions.
3. Detects and marks totals / subtotals rows.
4. Returns a list of StructuredTable objects ready for rendering.

Column keyword detection supports common invoice column names in several
languages (English, German, French, Albanian/Shqip, Italian, Spanish).
"""

from __future__ import annotations

import re
import logging
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Tuple

from .config import Config, DEFAULT_CONFIG
from .layout_detector import TableRegion, TextRow, Cell, ColumnBoundary

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Known column keywords (multilingual)
# ─────────────────────────────────────────────────────────────────────────────

_HEADER_KEYWORDS: Dict[str, List[str]] = {
    "item":        ["item", "nr", "no", "pos", "nr.", "#", "artikull", "position"],
    "description": [
        "description", "desc", "descri", "details", "detail",
        "bezeichnung", "artikel", "produkt", "description", "désignation",
        "përshkrim", "denominazione", "articulo",
    ],
    "quantity":    [
        "qty", "quantity", "qnt", "amount", "menge", "quantité",
        "sasia", "quantità", "cantidad",
    ],
    "unit":        ["unit", "uom", "u/m", "einheit", "unité", "njësi", "unità"],
    "unit_price":  [
        "unit price", "price", "rate", "preis", "unitaire", "çmimi",
        "prezzo", "precio", "p/unit",
    ],
    "tax":         ["tax", "vat", "mwst", "tva", "tvsh", "iva"],
    "total":       [
        "total", "amount", "subtotal", "sum", "gesamt", "montant",
        "totali", "totale", "importe",
    ],
}

_TOTAL_TRIGGER_WORDS = {
    "total", "subtotal", "grand total", "sum", "vat", "tax", "tvsh",
    "gesamt", "nettobetrag", "bruttobetrag", "totale", "totali",
}


# ─────────────────────────────────────────────────────────────────────────────
# Output data classes
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class StructuredCell:
    """A positioned cell within a StructuredTable."""

    text: str
    col_index: int       # 0-based column index within the table
    bold: bool = False
    font_size: Optional[float] = None


@dataclass
class StructuredRow:
    """A row within a StructuredTable."""

    cells: List[StructuredCell] = field(default_factory=list)
    is_header: bool = False
    is_total: bool = False
    original_row: Optional[TextRow] = None


@dataclass
class StructuredTable:
    """
    A fully resolved table with named columns and structured rows.

    Attributes
    ----------
    col_names   : list of inferred column names (may be empty strings for
                  columns without detectable headers)
    rows        : structured rows including the header row
    source      : the originating TableRegion
    """

    col_names: List[str] = field(default_factory=list)
    rows: List[StructuredRow] = field(default_factory=list)
    source: Optional[TableRegion] = None

    @property
    def num_cols(self) -> int:
        return len(self.col_names)

    @property
    def data_rows(self) -> List[StructuredRow]:
        """Non-header rows."""
        return [r for r in self.rows if not r.is_header]


# ─────────────────────────────────────────────────────────────────────────────
# Table Extractor
# ─────────────────────────────────────────────────────────────────────────────


class TableExtractor:
    """Converts raw TableRegion objects into StructuredTable objects."""

    def __init__(self, config: Config = DEFAULT_CONFIG) -> None:
        self.cfg = config

    # ------------------------------------------------------------------ #
    # Public API
    # ------------------------------------------------------------------ #

    def extract(self, table: TableRegion) -> StructuredTable:
        """
        Analyse *table* and return a StructuredTable.

        The first row is treated as a header when:
        - all its cells are bold, OR
        - it contains known column keywords.
        """
        if not table.rows:
            return StructuredTable(source=table)

        boundaries = table.col_boundaries
        num_cols = max(len(boundaries), 1)

        # Assign column indices to all cells
        structured_rows: List[StructuredRow] = []
        for raw_row in table.rows:
            s_row = self._structure_row(raw_row, boundaries, num_cols)
            structured_rows.append(s_row)

        # Detect header row
        header_idx = self._find_header_row(structured_rows)
        if header_idx is not None:
            structured_rows[header_idx].is_header = True

        # Detect totals rows
        for row in structured_rows:
            if self._is_totals_row(row):
                row.is_total = True

        # Extract column names from header row
        col_names = self._extract_col_names(structured_rows, num_cols, header_idx)

        return StructuredTable(
            col_names=col_names,
            rows=structured_rows,
            source=table,
        )

    def extract_all(self, tables: List[TableRegion]) -> List[StructuredTable]:
        """Extract all tables in the list."""
        return [self.extract(t) for t in tables]

    # ------------------------------------------------------------------ #
    # Private helpers
    # ------------------------------------------------------------------ #

    def _structure_row(
        self,
        raw_row: TextRow,
        boundaries: List[ColumnBoundary],
        num_cols: int,
    ) -> StructuredRow:
        """Assign a column index to every cell in *raw_row*."""
        structured_cells: List[StructuredCell] = []
        tolerance = self.cfg.column_tolerance * 2

        for cell in raw_row.cells:
            col_idx = self._nearest_col(cell.x0, boundaries, tolerance)
            structured_cells.append(
                StructuredCell(
                    text=cell.text,
                    col_index=col_idx,
                    bold=cell.bold,
                    font_size=cell.font_size,
                )
            )

        return StructuredRow(
            cells=structured_cells,
            original_row=raw_row,
        )

    def _nearest_col(
        self,
        x: float,
        boundaries: List[ColumnBoundary],
        tolerance: float,
    ) -> int:
        """Return the 0-based column index nearest to *x*."""
        if not boundaries:
            return 0
        idx = min(range(len(boundaries)), key=lambda i: abs(boundaries[i].x - x))
        return idx

    def _find_header_row(
        self, rows: List[StructuredRow]
    ) -> Optional[int]:
        """
        Return the index of the header row, or None if not found.

        Preference order:
        1. First row where every non-empty cell is bold.
        2. First row that contains known column keywords.
        """
        for i, row in enumerate(rows[:3]):   # check first 3 rows only
            if not row.cells:
                continue
            non_empty = [c for c in row.cells if c.text.strip()]
            if not non_empty:
                continue

            # Bold heuristic
            if all(c.bold for c in non_empty):
                return i

            # Keyword heuristic
            texts_lower = [c.text.lower().strip() for c in non_empty]
            keyword_matches = sum(
                1 for t in texts_lower
                if any(kw in t for kws in _HEADER_KEYWORDS.values() for kw in kws)
            )
            if keyword_matches >= min(2, len(non_empty)):
                return i

        return None

    def _is_totals_row(self, row: StructuredRow) -> bool:
        """Heuristic totals-row detection."""
        for cell in row.cells:
            lower = cell.text.lower().strip()
            if any(trigger in lower for trigger in _TOTAL_TRIGGER_WORDS):
                return True
            # Bold row with a number in the rightmost cell(s)
            if cell.bold and re.search(r"\d[\d.,\s]*$", cell.text):
                return True
        return False

    def _extract_col_names(
        self,
        rows: List[StructuredRow],
        num_cols: int,
        header_idx: Optional[int],
    ) -> List[str]:
        """Build a list of column name strings from the header row."""
        names = [""] * num_cols
        if header_idx is None or header_idx >= len(rows):
            return names

        header_row = rows[header_idx]
        for cell in header_row.cells:
            idx = cell.col_index
            if 0 <= idx < num_cols:
                names[idx] = cell.text.strip()

        return names
