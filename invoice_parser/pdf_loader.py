"""
pdf_loader.py - Stage 1: Document Ingestion

Loads a PDF and extracts all text elements with their bounding boxes using
pdfplumber (word-level precision) and PyMuPDF (vector line / drawing
detection for table-border recognition).

The module returns a list of PageData objects – one per page – that every
downstream stage operates on.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional

import pdfplumber
import fitz  # PyMuPDF

from .config import Config, DEFAULT_CONFIG

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Data classes
# ─────────────────────────────────────────────────────────────────────────────


@dataclass
class Word:
    """A single word extracted from the PDF with its spatial metadata."""

    text: str
    x0: float          # left edge  (PDF points, origin = top-left)
    y0: float          # top edge
    x1: float          # right edge
    y1: float          # bottom edge
    page_num: int      # 1-based

    font_size: Optional[float] = None
    font_name: Optional[str] = None
    bold: bool = False

    @property
    def mid_y(self) -> float:
        """Vertical midpoint – used for row clustering."""
        return (self.y0 + self.y1) / 2.0

    @property
    def mid_x(self) -> float:
        return (self.x0 + self.x1) / 2.0

    @property
    def height(self) -> float:
        return self.y1 - self.y0

    @property
    def width(self) -> float:
        return self.x1 - self.x0


@dataclass
class DetectedLine:
    """A vector line found in the PDF (potential table border)."""

    x0: float
    y0: float
    x1: float
    y1: float
    page_num: int

    @property
    def is_horizontal(self) -> bool:
        return abs(self.y1 - self.y0) < 2.0

    @property
    def is_vertical(self) -> bool:
        return abs(self.x1 - self.x0) < 2.0

    @property
    def length(self) -> float:
        return ((self.x1 - self.x0) ** 2 + (self.y1 - self.y0) ** 2) ** 0.5


@dataclass
class PageData:
    """All extracted data for a single PDF page."""

    page_num: int
    width: float          # PDF points
    height: float
    words: List[Word] = field(default_factory=list)
    lines: List[DetectedLine] = field(default_factory=list)
    is_scanned: bool = False

    @property
    def char_count(self) -> int:
        return sum(len(w.text) for w in self.words)


# ─────────────────────────────────────────────────────────────────────────────
# Loader
# ─────────────────────────────────────────────────────────────────────────────


class PDFLoader:
    """
    Loads a PDF file and returns a list of PageData objects.

    Strategy
    --------
    - pdfplumber handles word extraction with precise bounding boxes and
      font metadata.
    - PyMuPDF handles vector drawing / line detection (table borders,
      rectangles) without duplicating word extraction.
    """

    def __init__(self, config: Config = DEFAULT_CONFIG) -> None:
        self.config = config

    # ------------------------------------------------------------------ #

    def load(self, pdf_path: str | Path) -> List[PageData]:
        """
        Load *pdf_path* and return one PageData per page.

        Raises
        ------
        FileNotFoundError  – if the file does not exist.
        ValueError         – if the file is not a valid PDF.
        """
        path = Path(pdf_path)
        if not path.exists():
            raise FileNotFoundError(f"PDF not found: {path}")

        logger.info("Loading %s", path.name)
        pages: List[PageData] = []

        try:
            with pdfplumber.open(str(path)) as plumber_doc:
                fitz_doc = fitz.open(str(path))
                try:
                    n_pages = min(len(plumber_doc.pages), self.config.max_pages)
                    for page_num in range(n_pages):
                        plumber_page = plumber_doc.pages[page_num]
                        fitz_page = fitz_doc[page_num]

                        words = self._extract_words(plumber_page, page_num + 1)
                        lines = self._extract_lines(fitz_page, page_num + 1)

                        total_chars = sum(len(w.text) for w in words)
                        is_scanned = total_chars < self.config.ocr_text_threshold

                        pages.append(
                            PageData(
                                page_num=page_num + 1,
                                width=float(plumber_page.width),
                                height=float(plumber_page.height),
                                words=words,
                                lines=lines,
                                is_scanned=is_scanned,
                            )
                        )
                        logger.debug(
                            "Page %d: %d words, %d lines, scanned=%s",
                            page_num + 1, len(words), len(lines), is_scanned,
                        )
                finally:
                    fitz_doc.close()
        except Exception as exc:
            raise ValueError(f"Failed to read PDF '{path}': {exc}") from exc

        logger.info("Loaded %d page(s) from %s", len(pages), path.name)
        return pages

    # ------------------------------------------------------------------ #
    # Private helpers
    # ------------------------------------------------------------------ #

    def _extract_words(self, page, page_num: int) -> List[Word]:
        """Extract words with bounding boxes and font info via pdfplumber."""
        words: List[Word] = []

        try:
            raw = page.extract_words(
                x_tolerance=3,
                y_tolerance=3,
                keep_blank_chars=False,
                use_text_flow=False,
                extra_attrs=["fontname", "size"],
            )
        except Exception as exc:
            logger.warning("Word extraction failed on page %d: %s", page_num, exc)
            return words

        cfg = self.config
        for w in raw:
            text = (w.get("text") or "").strip()
            if not text:
                continue

            font_name = w.get("fontname") or ""
            bold = any(
                sub in font_name.lower()
                for sub in cfg.bold_font_substrings
            )

            words.append(
                Word(
                    text=text,
                    x0=float(w["x0"]),
                    y0=float(w["top"]),
                    x1=float(w["x1"]),
                    y1=float(w["bottom"]),
                    page_num=page_num,
                    font_size=w.get("size"),
                    font_name=font_name or None,
                    bold=bold,
                )
            )

        return words

    def _extract_lines(self, fitz_page, page_num: int) -> List[DetectedLine]:
        """
        Extract vector lines from PyMuPDF drawings.

        Rectangles are decomposed into their four border lines so that
        downstream table detection sees individual horizontal / vertical
        segments.
        """
        lines: List[DetectedLine] = []

        try:
            drawings = fitz_page.get_drawings()
        except Exception as exc:
            logger.warning("Line extraction failed on page %d: %s", page_num, exc)
            return lines

        for drawing in drawings:
            for item in drawing.get("items", []):
                kind = item[0]

                if kind == "l":          # simple line segment
                    p1, p2 = item[1], item[2]
                    lines.append(
                        DetectedLine(p1.x, p1.y, p2.x, p2.y, page_num)
                    )

                elif kind == "re":       # rectangle → four border lines
                    r = item[1]
                    # Normalise coordinates (fitz can return inverted rects)
                    x0, x1 = min(r.x0, r.x1), max(r.x0, r.x1)
                    y0, y1 = min(r.y0, r.y1), max(r.y0, r.y1)
                    if x1 - x0 < 2 and y1 - y0 < 2:
                        continue  # degenerate rectangle
                    lines.extend([
                        DetectedLine(x0, y0, x1, y0, page_num),  # top
                        DetectedLine(x0, y1, x1, y1, page_num),  # bottom
                        DetectedLine(x0, y0, x0, y1, page_num),  # left
                        DetectedLine(x1, y0, x1, y1, page_num),  # right
                    ])

                elif kind == "qu":       # quadrilateral (less common)
                    pts = item[1]
                    if len(pts) >= 2:
                        for i in range(len(pts)):
                            p1 = pts[i]
                            p2 = pts[(i + 1) % len(pts)]
                            lines.append(
                                DetectedLine(p1.x, p1.y, p2.x, p2.y, page_num)
                            )

        return lines
