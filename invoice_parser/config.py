"""
config.py - Centralized configuration for the PDF-to-Excel pipeline.

All tunable parameters live here so every other module imports them from
a single source of truth.  Edit this file to adapt the system to different
PDF styles without touching algorithm code.
"""

from __future__ import annotations
from dataclasses import dataclass, field
from typing import List, Optional


@dataclass
class Config:
    # ------------------------------------------------------------------ #
    # ROW DETECTION                                                        #
    # ------------------------------------------------------------------ #
    # Two words are on the same row when their vertical midpoints differ
    # by at most this many PDF points (1 pt ≈ 0.353 mm).
    row_tolerance: float = 4.0

    # ------------------------------------------------------------------ #
    # COLUMN DETECTION                                                     #
    # ------------------------------------------------------------------ #
    # Two x0 coordinates are considered the same column boundary when they
    # differ by at most this many PDF points.
    column_tolerance: float = 8.0

    # Minimum horizontal gap (points) that separates two distinct cells
    # within the same row.  Smaller gaps → words are merged into one cell.
    cell_merge_gap: float = 6.0

    # ------------------------------------------------------------------ #
    # TABLE DETECTION                                                      #
    # ------------------------------------------------------------------ #
    # A block of rows is classified as a table when at least this fraction
    # of rows share the same detected column boundaries.
    table_consistency_threshold: float = 0.60

    # Minimum number of rows required to declare a table.
    table_min_rows: int = 2

    # Minimum number of distinct columns required to declare a table.
    table_min_cols: int = 2

    # Maximum vertical gap (points) between rows still considered part of
    # the same table.
    table_max_row_gap: float = 20.0

    # ------------------------------------------------------------------ #
    # REGION SEGMENTATION                                                  #
    # ------------------------------------------------------------------ #
    # Fraction of page height used as the divider between header and body.
    header_fraction: float = 0.25

    # Fraction of page height measured from the bottom for footer detection.
    footer_fraction: float = 0.10

    # ------------------------------------------------------------------ #
    # OCR                                                                  #
    # ------------------------------------------------------------------ #
    # Tesseract language string (e.g. "eng", "deu", "fra", "sqi+eng").
    ocr_language: str = "eng"

    # Resolution for rasterising a page before OCR.
    ocr_dpi: int = 300

    # If a page contains fewer characters than this threshold it is treated
    # as scanned and sent to OCR.
    ocr_text_threshold: int = 30

    # ------------------------------------------------------------------ #
    # EXCEL OUTPUT                                                         #
    # ------------------------------------------------------------------ #
    # The PDF page width is divided into this many Excel columns to form
    # the base coordinate grid.  Higher value → finer horizontal resolution.
    excel_grid_cols: int = 60

    # Default Excel column width in character units when no better measure
    # is available.
    default_col_width: float = 8.5

    # Absolute limits for computed column widths.
    min_col_width: float = 2.0
    max_col_width: float = 60.0

    # Excel row height in points.
    row_height: float = 14.0

    # Extra blank Excel rows inserted between PDF text rows when the
    # vertical gap exceeds this many PDF points.
    spacer_row_threshold: float = 14.0

    # Bold text detection: font names containing any of these substrings
    # (case-insensitive) are treated as bold.
    bold_font_substrings: List[str] = field(
        default_factory=lambda: ["bold", "heavy", "black", "demi", "semibold"]
    )

    # ------------------------------------------------------------------ #
    # PERFORMANCE                                                          #
    # ------------------------------------------------------------------ #
    # Hard limit on the number of pages processed per PDF.
    max_pages: int = 500

    # ------------------------------------------------------------------ #
    # AI HEURISTICS (optional enhancements)                               #
    # ------------------------------------------------------------------ #
    # Patterns for recognising invoice metadata fields.
    invoice_number_patterns: List[str] = field(
        default_factory=lambda: [
            r"invoice\s*#?\s*:?\s*(\S+)",
            r"inv\s*\.?\s*no\.?\s*:?\s*(\S+)",
            r"faktur[ë\s]+nr\.?\s*:?\s*(\S+)",
        ]
    )

    currency_symbols: List[str] = field(
        default_factory=lambda: ["$", "€", "£", "¥", "ALL", "LEK"]
    )


# Singleton – import this instance everywhere else.
DEFAULT_CONFIG = Config()
