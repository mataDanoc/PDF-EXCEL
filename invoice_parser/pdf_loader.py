"""
pdf_loader.py - Stage 1: Document Ingestion

Extracts every text element and visual element from the PDF:
  - Words with bounding boxes, font info, bold detection
  - Vector lines (table borders)
  - Filled rectangles with their RGB colours (backgrounds)

Two libraries work in tandem:
  - pdfplumber  -> word-level extraction with font metadata
  - PyMuPDF     -> vector drawings, lines, fills, rectangles
"""

from __future__ import annotations

import logging
from dataclasses import dataclass, field
from pathlib import Path
from typing import List, Optional, Tuple

import pdfplumber
import fitz  # PyMuPDF

from .config import Config, DEFAULT_CONFIG

logger = logging.getLogger(__name__)


# ── Data classes ──────────────────────────────────────────────────────────────

@dataclass
class Word:
    text: str
    x0: float
    y0: float
    x1: float
    y1: float
    page_num: int
    font_size: Optional[float] = None
    font_name: Optional[str] = None
    bold: bool = False

    @property
    def mid_y(self) -> float:
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
class FilledRect:
    """A filled rectangle from the PDF (table header background, row stripe, etc.)."""

    x0: float
    y0: float
    x1: float
    y1: float
    page_num: int
    color_rgb: Tuple[int, int, int]   # (R, G, B) 0-255

    @property
    def hex_color(self) -> str:
        """ARGB hex string for openpyxl."""
        r, g, b = self.color_rgb
        return f"FF{r:02X}{g:02X}{b:02X}"

    @property
    def is_dark(self) -> bool:
        """Luminance < 128 → dark background → needs white text."""
        r, g, b = self.color_rgb
        return (r * 0.299 + g * 0.587 + b * 0.114) < 128

    @property
    def is_white_or_near(self) -> bool:
        """Ignore near-white fills (they are page background, not styling)."""
        r, g, b = self.color_rgb
        return r > 240 and g > 240 and b > 240

    def contains_point(self, x: float, y: float, margin: float = 2.0) -> bool:
        return (self.x0 - margin <= x <= self.x1 + margin and
                self.y0 - margin <= y <= self.y1 + margin)


@dataclass
class PageData:
    page_num: int
    width: float
    height: float
    words: List[Word] = field(default_factory=list)
    lines: List[DetectedLine] = field(default_factory=list)
    filled_rects: List[FilledRect] = field(default_factory=list)
    is_scanned: bool = False

    @property
    def char_count(self) -> int:
        return sum(len(w.text) for w in self.words)


# ── Loader ────────────────────────────────────────────────────────────────────

class PDFLoader:
    def __init__(self, config: Config = DEFAULT_CONFIG) -> None:
        self.config = config

    def load(self, pdf_path: str | Path) -> List[PageData]:
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
                        lines, filled_rects = self._extract_visuals(fitz_page, page_num + 1)

                        total_chars = sum(len(w.text) for w in words)
                        is_scanned = total_chars < self.config.ocr_text_threshold

                        pages.append(PageData(
                            page_num=page_num + 1,
                            width=float(plumber_page.width),
                            height=float(plumber_page.height),
                            words=words,
                            lines=lines,
                            filled_rects=filled_rects,
                            is_scanned=is_scanned,
                        ))
                        logger.debug(
                            "Page %d: %d words, %d lines, %d fills, scanned=%s",
                            page_num + 1, len(words), len(lines),
                            len(filled_rects), is_scanned,
                        )
                finally:
                    fitz_doc.close()
        except Exception as exc:
            raise ValueError(f"Failed to read PDF '{path}': {exc}") from exc

        logger.info("Loaded %d page(s) from %s", len(pages), path.name)
        return pages

    # ── Word extraction (pdfplumber) ──────────────────────────────────────

    def _extract_words(self, page, page_num: int) -> List[Word]:
        words: List[Word] = []
        try:
            raw = page.extract_words(
                x_tolerance=3, y_tolerance=3,
                keep_blank_chars=False, use_text_flow=False,
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
            bold = any(sub in font_name.lower() for sub in cfg.bold_font_substrings)
            words.append(Word(
                text=text,
                x0=float(w["x0"]), y0=float(w["top"]),
                x1=float(w["x1"]), y1=float(w["bottom"]),
                page_num=page_num,
                font_size=w.get("size"),
                font_name=font_name or None,
                bold=bold,
            ))
        return words

    # ── Visual extraction (PyMuPDF) ──────────────────────────────────────

    def _extract_visuals(
        self, fitz_page, page_num: int
    ) -> Tuple[List[DetectedLine], List[FilledRect]]:
        """Extract vector lines AND filled rectangles from PyMuPDF drawings."""
        lines: List[DetectedLine] = []
        fills: List[FilledRect] = []

        try:
            drawings = fitz_page.get_drawings()
        except Exception as exc:
            logger.warning("Drawing extraction failed on page %d: %s", page_num, exc)
            return lines, fills

        for drawing in drawings:
            fill_color = drawing.get("fill")       # (R,G,B) 0..1 or None
            stroke_color = drawing.get("color")     # stroke colour

            for item in drawing.get("items", []):
                kind = item[0]

                if kind == "l":
                    p1, p2 = item[1], item[2]
                    lines.append(DetectedLine(p1.x, p1.y, p2.x, p2.y, page_num))

                elif kind == "re":
                    r = item[1]
                    x0, x1 = min(r.x0, r.x1), max(r.x0, r.x1)
                    y0, y1 = min(r.y0, r.y1), max(r.y0, r.y1)
                    if x1 - x0 < 2 and y1 - y0 < 2:
                        continue

                    # Border lines from rectangle
                    lines.extend([
                        DetectedLine(x0, y0, x1, y0, page_num),
                        DetectedLine(x0, y1, x1, y1, page_num),
                        DetectedLine(x0, y0, x0, y1, page_num),
                        DetectedLine(x1, y0, x1, y1, page_num),
                    ])

                    # Filled rectangle → background colour
                    if fill_color and (x1 - x0) > 10 and (y1 - y0) > 3:
                        rgb = (
                            int(fill_color[0] * 255),
                            int(fill_color[1] * 255),
                            int(fill_color[2] * 255),
                        )
                        fr = FilledRect(x0, y0, x1, y1, page_num, rgb)
                        if not fr.is_white_or_near:
                            fills.append(fr)

                elif kind == "qu":
                    pts = item[1]
                    if len(pts) >= 2:
                        for i in range(len(pts)):
                            p1 = pts[i]
                            p2 = pts[(i + 1) % len(pts)]
                            lines.append(DetectedLine(p1.x, p1.y, p2.x, p2.y, page_num))

        return lines, fills
