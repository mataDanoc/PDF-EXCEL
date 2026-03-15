"""
ocr_engine.py - Stage 2: OCR Fallback

When a PDF page contains fewer machine-readable characters than the
configured threshold (indicating a scanned document), this module
rasterises the page and runs Tesseract OCR to recover word bounding boxes.

The recovered words are returned as the same Word dataclass used by
pdf_loader so downstream stages need no special-casing.
"""

from __future__ import annotations

import logging
from pathlib import Path
from typing import List, Optional

from .config import Config, DEFAULT_CONFIG
from .pdf_loader import PageData, Word

logger = logging.getLogger(__name__)

# Optional imports – gracefully degrade if not installed
try:
    import pytesseract
    from PIL import Image
    import numpy as np
    import fitz  # PyMuPDF – used for rasterisation
    _OCR_AVAILABLE = True
except ImportError as _e:
    _OCR_AVAILABLE = False
    logger.warning("OCR dependencies not available (%s). Scanned pages will be skipped.", _e)

try:
    import cv2
    _CV2_AVAILABLE = True
except ImportError:
    _CV2_AVAILABLE = False
    logger.debug("OpenCV not available – skipping image pre-processing for OCR.")


# ─────────────────────────────────────────────────────────────────────────────
# OCR Engine
# ─────────────────────────────────────────────────────────────────────────────


class OCREngine:
    """
    Converts a scanned PDF page to a list of Word objects via Tesseract.

    Processing pipeline
    -------------------
    1. Rasterise the page at `config.ocr_dpi` using PyMuPDF.
    2. (Optional) Apply OpenCV pre-processing: grayscale, threshold,
       deskew – to improve OCR accuracy.
    3. Run pytesseract with image_to_data() to obtain per-word bounding
       boxes and confidence scores.
    4. Filter low-confidence results and convert coordinates back to PDF
       point space.
    """

    def __init__(self, config: Config = DEFAULT_CONFIG) -> None:
        self.config = config

    # ------------------------------------------------------------------ #

    def process_page(self, page_data: PageData, pdf_path: str | Path) -> List[Word]:
        """
        Run OCR on *page_data* and return a list of Word objects.

        Returns an empty list (with a warning) when OCR dependencies are
        unavailable.
        """
        if not _OCR_AVAILABLE:
            logger.warning(
                "Page %d is scanned but OCR libraries are not installed. "
                "Install: pip install pytesseract pillow opencv-python-headless",
                page_data.page_num,
            )
            return []

        logger.info("Running OCR on page %d …", page_data.page_num)

        pil_image, scale_x, scale_y = self._rasterise(
            pdf_path, page_data.page_num - 1,  # fitz is 0-indexed
            page_data.width, page_data.height,
        )
        if pil_image is None:
            return []

        if _CV2_AVAILABLE:
            pil_image = self._preprocess(pil_image)

        words = self._run_tesseract(pil_image, page_data.page_num, scale_x, scale_y)
        logger.info("OCR recovered %d words on page %d", len(words), page_data.page_num)
        return words

    # ------------------------------------------------------------------ #
    # Private helpers
    # ------------------------------------------------------------------ #

    def _rasterise(
        self,
        pdf_path: str | Path,
        page_index: int,
        page_width_pt: float,
        page_height_pt: float,
    ):
        """
        Render the PDF page to a PIL Image at the configured DPI.

        Returns
        -------
        (pil_image, scale_x, scale_y) – scale factors convert pixel coords
        back to PDF points.  Returns (None, 1, 1) on failure.
        """
        try:
            doc = fitz.open(str(pdf_path))
            page = doc[page_index]
            dpi = self.config.ocr_dpi
            mat = fitz.Matrix(dpi / 72.0, dpi / 72.0)
            pix = page.get_pixmap(matrix=mat, alpha=False)
            doc.close()

            img_array = np.frombuffer(pix.samples, dtype=np.uint8).reshape(
                pix.h, pix.w, pix.n
            )
            pil_image = Image.fromarray(img_array, "RGB")

            scale_x = page_width_pt / pix.w
            scale_y = page_height_pt / pix.h
            return pil_image, scale_x, scale_y

        except Exception as exc:
            logger.error("Rasterisation failed for page %d: %s", page_index + 1, exc)
            return None, 1.0, 1.0

    def _preprocess(self, pil_image: "Image.Image") -> "Image.Image":
        """
        Apply OpenCV pre-processing to improve OCR accuracy:
        - Convert to grayscale
        - Apply adaptive thresholding
        - Mild deskew
        """
        try:
            img = np.array(pil_image)
            gray = cv2.cvtColor(img, cv2.COLOR_RGB2GRAY)

            # Adaptive threshold – handles uneven lighting / shadows
            binary = cv2.adaptiveThreshold(
                gray, 255,
                cv2.ADAPTIVE_THRESH_GAUSSIAN_C,
                cv2.THRESH_BINARY, 31, 10
            )

            # Deskew using moments
            coords = np.column_stack(np.where(binary < 128))
            if len(coords) > 100:
                angle = cv2.minAreaRect(coords)[-1]
                if angle < -45:
                    angle = 90 + angle
                if abs(angle) > 0.5:
                    h, w = binary.shape
                    center = (w // 2, h // 2)
                    M = cv2.getRotationMatrix2D(center, angle, 1.0)
                    binary = cv2.warpAffine(
                        binary, M, (w, h),
                        flags=cv2.INTER_CUBIC,
                        borderMode=cv2.BORDER_REPLICATE,
                    )

            return Image.fromarray(binary)
        except Exception as exc:
            logger.debug("Pre-processing skipped: %s", exc)
            return pil_image

    def _run_tesseract(
        self,
        pil_image: "Image.Image",
        page_num: int,
        scale_x: float,
        scale_y: float,
        confidence_threshold: int = 30,
    ) -> List[Word]:
        """Run pytesseract and convert pixel bounding boxes to PDF points."""
        try:
            data = pytesseract.image_to_data(
                pil_image,
                lang=self.config.ocr_language,
                output_type=pytesseract.Output.DICT,
            )
        except Exception as exc:
            logger.error("Tesseract failed on page %d: %s", page_num, exc)
            return []

        words: List[Word] = []
        n = len(data["text"])

        for i in range(n):
            text = (data["text"][i] or "").strip()
            if not text:
                continue
            conf = int(data["conf"][i]) if data["conf"][i] != "-1" else 0
            if conf < confidence_threshold:
                continue

            x_px = data["left"][i]
            y_px = data["top"][i]
            w_px = data["width"][i]
            h_px = data["height"][i]

            x0 = x_px * scale_x
            y0 = y_px * scale_y
            x1 = (x_px + w_px) * scale_x
            y1 = (y_px + h_px) * scale_y

            words.append(
                Word(
                    text=text,
                    x0=x0, y0=y0, x1=x1, y1=y1,
                    page_num=page_num,
                    font_size=None,  # not available from OCR
                    font_name=None,
                    bold=False,
                )
            )

        return words
