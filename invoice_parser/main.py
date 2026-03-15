"""
main.py - Pipeline Orchestrator & Batch Processor

Wires together all pipeline stages:

    PDFLoader → (OCREngine) → LayoutDetector → GridBuilder → ExcelRenderer

Public functions
----------------
convert(pdf_path, output_path, config)
    Convert a single PDF to Excel.

batch_convert(input_dir, output_dir, config)
    Convert all PDF files in input_dir, writing results to output_dir.

CLI usage
---------
    python -m invoice_parser.main input/invoice.pdf output/invoice.xlsx
    python -m invoice_parser.main input/ output/
"""

from __future__ import annotations

import logging
import sys
import time
from pathlib import Path
from typing import List, Optional

from .config import Config, DEFAULT_CONFIG
from .pdf_loader import PDFLoader, PageData
from .ocr_engine import OCREngine
from .layout_detector import LayoutDetector
from .grid_builder import GridBuilder, ExcelGrid
from .excel_renderer import ExcelRenderer

logger = logging.getLogger(__name__)


# ─────────────────────────────────────────────────────────────────────────────
# Pipeline
# ─────────────────────────────────────────────────────────────────────────────


class ConversionPipeline:
    """
    Full PDF → Excel conversion pipeline for a single file.

    Stages
    ------
    1. PDFLoader      – load pages, extract words and vector lines
    2. OCREngine      – OCR fallback for scanned pages
    3. LayoutDetector – detect rows, columns, tables, regions
    4. GridBuilder    – map layout to Excel grid coordinates
    5. ExcelRenderer  – write .xlsx file
    """

    def __init__(self, config: Config = DEFAULT_CONFIG) -> None:
        self.cfg = config
        self.loader    = PDFLoader(config)
        self.ocr       = OCREngine(config)
        self.detector  = LayoutDetector(config)
        self.builder   = GridBuilder(config)
        self.renderer  = ExcelRenderer(config)

    # ------------------------------------------------------------------ #

    def run(self, pdf_path: str | Path, output_path: str | Path) -> Path:
        """
        Convert *pdf_path* to an Excel file at *output_path*.

        Returns
        -------
        Path of the written .xlsx file.

        Raises
        ------
        FileNotFoundError – if *pdf_path* does not exist.
        ValueError        – if the file is not a valid PDF.
        RuntimeError      – if a required library is missing.
        """
        pdf_path = Path(pdf_path)
        output_path = Path(output_path)

        t_start = time.perf_counter()
        logger.info("-" * 60)
        logger.info("Converting: %s", pdf_path.name)

        # Stage 1: Load
        pages: List[PageData] = self.loader.load(str(pdf_path))

        # Stage 2: OCR fallback for scanned pages
        for page in pages:
            if page.is_scanned:
                logger.info("Page %d is scanned – running OCR …", page.page_num)
                ocr_words = self.ocr.process_page(page, str(pdf_path))
                if ocr_words:
                    page.words = ocr_words
                    page.is_scanned = False

        # Stages 3–4: Layout analysis + grid building
        grids: List[ExcelGrid] = []
        for page in pages:
            if not page.words:
                logger.warning("Page %d has no extractable text – skipping.", page.page_num)
                continue

            layout  = self.detector.analyse(page)
            grid    = self.builder.build(layout)
            grids.append(grid)

            logger.debug(
                "Page %d → %d grid cells, %d rows, %d cols",
                page.page_num, len(grid.cells), grid.total_rows, grid.total_cols,
            )

        if not grids:
            raise ValueError(f"No extractable text found in '{pdf_path.name}'.")

        # Stage 5: Render Excel
        result = self.renderer.render(grids, output_path)

        elapsed = time.perf_counter() - t_start
        logger.info(
            "Done: %s -> %s  (%.2f s, %d page(s))",
            pdf_path.name, result.name, elapsed, len(grids),
        )
        return result


# ─────────────────────────────────────────────────────────────────────────────
# Public convenience functions
# ─────────────────────────────────────────────────────────────────────────────


def convert(
    pdf_path: str | Path,
    output_path: str | Path,
    config: Config = DEFAULT_CONFIG,
) -> Path:
    """
    Convert a single PDF to an Excel file.

    Parameters
    ----------
    pdf_path    : path to the source .pdf file
    output_path : desired path for the .xlsx output
    config      : optional Config object to override defaults

    Returns
    -------
    Resolved Path of the written .xlsx file.

    Example
    -------
    >>> from invoice_parser import convert
    >>> convert("invoices/inv_001.pdf", "output/inv_001.xlsx")
    """
    pipeline = ConversionPipeline(config)
    return pipeline.run(pdf_path, output_path)


def batch_convert(
    input_dir: str | Path,
    output_dir: str | Path,
    config: Config = DEFAULT_CONFIG,
    glob: str = "*.pdf",
    fail_fast: bool = False,
) -> List[Path]:
    """
    Convert all PDF files in *input_dir* to Excel files in *output_dir*.

    Parameters
    ----------
    input_dir   : directory containing .pdf files (searched recursively
                  when glob contains '**')
    output_dir  : directory for .xlsx output files (created if missing)
    config      : optional Config object
    glob        : file-name pattern (default "*.pdf"; use "**/*.pdf" for
                  recursive search)
    fail_fast   : if True, stop on first error; otherwise log and continue

    Returns
    -------
    List of paths of successfully written .xlsx files.

    Example
    -------
    >>> from invoice_parser import batch_convert
    >>> batch_convert("input/", "output/")
    """
    in_dir  = Path(input_dir)
    out_dir = Path(output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)

    pdf_files = sorted(in_dir.glob(glob))
    if not pdf_files:
        logger.warning("No PDF files found in %s (pattern=%s)", in_dir, glob)
        return []

    logger.info("Batch converting %d PDF(s) from %s -> %s", len(pdf_files), in_dir, out_dir)

    pipeline  = ConversionPipeline(config)
    succeeded: List[Path] = []
    failed:    List[Path] = []

    for pdf_file in pdf_files:
        xlsx_name   = pdf_file.stem + ".xlsx"
        output_path = out_dir / xlsx_name
        try:
            result = pipeline.run(pdf_file, output_path)
            succeeded.append(result)
        except Exception as exc:
            logger.error("FAILED %s: %s", pdf_file.name, exc)
            failed.append(pdf_file)
            if fail_fast:
                raise

    logger.info(
        "Batch complete: %d succeeded, %d failed",
        len(succeeded), len(failed),
    )
    if failed:
        logger.warning(
            "Failed files: %s",
            ", ".join(f.name for f in failed),
        )
    return succeeded


# ─────────────────────────────────────────────────────────────────────────────
# CLI entry point
# ─────────────────────────────────────────────────────────────────────────────


def _setup_logging(verbose: bool = False) -> None:
    level = logging.DEBUG if verbose else logging.INFO
    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
    )


def main(argv: Optional[List[str]] = None) -> int:
    """CLI entry point. Returns exit code."""
    import argparse

    parser = argparse.ArgumentParser(
        prog="invoice_parser",
        description="Convert PDF invoices to layout-preserving Excel files.",
    )
    parser.add_argument("input",  help="PDF file or directory of PDFs")
    parser.add_argument("output", help="Excel file (.xlsx) or output directory")
    parser.add_argument("-v", "--verbose", action="store_true")
    parser.add_argument(
        "--row-tolerance",  type=float, default=None,
        help="Row clustering tolerance in PDF points (default: 4.0)",
    )
    parser.add_argument(
        "--col-tolerance",  type=float, default=None,
        help="Column clustering tolerance in PDF points (default: 8.0)",
    )
    parser.add_argument(
        "--ocr-lang",       default=None,
        help="Tesseract language string, e.g. 'eng', 'sqi+eng' (default: eng)",
    )
    parser.add_argument(
        "--grid-cols",      type=int, default=None,
        help="Number of Excel grid columns (default: 60)",
    )
    parser.add_argument(
        "--fail-fast",      action="store_true",
        help="Stop batch processing on first error",
    )

    args = parser.parse_args(argv)
    _setup_logging(args.verbose)

    # Build config from CLI overrides
    cfg = Config()
    if args.row_tolerance is not None:
        cfg.row_tolerance = args.row_tolerance
    if args.col_tolerance is not None:
        cfg.column_tolerance = args.col_tolerance
    if args.ocr_lang is not None:
        cfg.ocr_language = args.ocr_lang
    if args.grid_cols is not None:
        cfg.excel_grid_cols = args.grid_cols

    in_path  = Path(args.input)
    out_path = Path(args.output)

    try:
        if in_path.is_dir():
            results = batch_convert(
                in_path, out_path, config=cfg, fail_fast=args.fail_fast
            )
            print(f"Converted {len(results)} file(s) → {out_path}")
        elif in_path.is_file():
            # Single file mode: if output is a directory, derive filename
            if out_path.is_dir() or not out_path.suffix:
                out_path = out_path / (in_path.stem + ".xlsx")
            result = convert(in_path, out_path, config=cfg)
            print(f"Saved: {result}")
        else:
            print(f"Error: '{in_path}' does not exist.", file=sys.stderr)
            return 1
    except Exception as exc:
        print(f"Error: {exc}", file=sys.stderr)
        if args.verbose:
            import traceback
            traceback.print_exc()
        return 1

    return 0


if __name__ == "__main__":
    sys.exit(main())
