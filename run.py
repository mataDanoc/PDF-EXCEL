"""
run.py - Quick-start runner

Drop your PDF files into the input/ folder, run this script,
and find the Excel files in output/.

Usage:
    python run.py                  # converts all PDFs in input/
    python run.py invoice.pdf      # converts a single file from input/
    python run.py path/to/any.pdf  # converts any absolute or relative path
"""

from __future__ import annotations
import sys
import logging
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from invoice_parser import convert, batch_convert, Config

# ── Logging setup ─────────────────────────────────────────────────────────────
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)

# ── Paths ─────────────────────────────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
INPUT_DIR  = BASE_DIR / "input"
OUTPUT_DIR = BASE_DIR / "output"

# ── Configuration – edit here to tune behaviour ───────────────────────────────
cfg = Config(
    row_tolerance          = 4.0,   # pt – increase for loose line spacing
    column_tolerance       = 8.0,   # pt – increase for wide column gaps
    cell_merge_gap         = 6.0,   # pt – max gap to merge adjacent words
    excel_grid_cols        = 60,    # Excel columns for non-table text
    ocr_language           = "eng", # or "sqi+eng" for Albanian+English
    ocr_dpi                = 300,
    row_height             = 14.0,
    spacer_row_threshold   = 14.0,
    table_min_rows         = 2,
    table_min_cols         = 2,
    table_consistency_threshold = 0.60,
)

# ── Main ──────────────────────────────────────────────────────────────────────

def main() -> None:
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    if len(sys.argv) > 1:
        # Single file mode
        arg = sys.argv[1]
        pdf_path = Path(arg)
        if not pdf_path.is_absolute():
            # Try relative to input/ first
            candidate = INPUT_DIR / arg
            if candidate.exists():
                pdf_path = candidate

        if not pdf_path.exists():
            print(f"File not found: {pdf_path}")
            sys.exit(1)

        out = OUTPUT_DIR / (pdf_path.stem + ".xlsx")
        result = convert(pdf_path, out, config=cfg)
        print(f"Saved: {result}")

    else:
        # Batch mode: all PDFs in input/
        INPUT_DIR.mkdir(exist_ok=True)
        pdf_files = list(INPUT_DIR.glob("*.pdf"))
        if not pdf_files:
            print(f"No PDF files found in {INPUT_DIR}")
            print("Put your PDF files in the 'input/' folder and run again.")
            sys.exit(0)

        print(f"Found {len(pdf_files)} PDF(s) in {INPUT_DIR}")
        results = batch_convert(INPUT_DIR, OUTPUT_DIR, config=cfg)
        print(f"\nDone: {len(results)} file(s) converted.")
        print(f"Output folder: {OUTPUT_DIR}")

if __name__ == "__main__":
    main()
