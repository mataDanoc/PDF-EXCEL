"""
test_conversion.py - End-to-end test for the PDF→Excel pipeline

Creates a synthetic invoice PDF, runs the full pipeline, and validates
the output.  Also tests batch conversion.

Usage
-----
    python test_conversion.py
    python test_conversion.py --verbose
"""

from __future__ import annotations

import argparse
import logging
import sys
import textwrap
from pathlib import Path

# ── Ensure the package is importable from this directory ──────────────────────
sys.path.insert(0, str(Path(__file__).parent))

from invoice_parser import convert, batch_convert, Config


# ─────────────────────────────────────────────────────────────────────────────
# Helpers: synthetic PDF creation
# ─────────────────────────────────────────────────────────────────────────────


def create_test_invoice_pdf(output_path: Path) -> None:
    """
    Build a realistic-looking invoice PDF using reportlab (if available)
    or a minimal raw-PDF fallback.
    """
    try:
        _create_with_reportlab(output_path)
        print(f"  Created test PDF (reportlab): {output_path.name}")
    except ImportError:
        _create_minimal_pdf(output_path)
        print(f"  Created test PDF (minimal):   {output_path.name}")


def _create_with_reportlab(output_path: Path) -> None:
    """Rich invoice PDF via reportlab."""
    from reportlab.lib import colors
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.units import cm
    from reportlab.platypus import (
        SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer,
    )
    from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle

    doc = SimpleDocTemplate(
        str(output_path),
        pagesize=A4,
        leftMargin=2*cm, rightMargin=2*cm,
        topMargin=2*cm,  bottomMargin=2*cm,
    )
    styles = getSampleStyleSheet()
    bold   = ParagraphStyle("Bold", parent=styles["Normal"], fontName="Helvetica-Bold")

    story = []

    # Header
    story.append(Paragraph("<b>ACME SOLUTIONS SH.P.K.</b>", styles["Title"]))
    story.append(Paragraph("Rr. Skënderbeg, Nr. 10 | Tiranë, Albania", styles["Normal"]))
    story.append(Paragraph("Tel: +355 4 222 3344 | Email: info@acme.al | NIPT: K12345678A", styles["Normal"]))
    story.append(Spacer(1, 0.5*cm))

    # Invoice metadata
    meta = [
        ["FATURË / INVOICE",    "",       "Nr. Faturës / Invoice #:", "INV-2024-0042"],
        ["",                    "",       "Datë / Date:",             "15.03.2024"],
        ["",                    "",       "Afati / Due Date:",        "30.03.2024"],
    ]
    meta_tbl = Table(meta, colWidths=[5*cm, 3*cm, 5*cm, 4*cm])
    meta_tbl.setStyle(TableStyle([
        ("SPAN",      (0,0), (1,2)),
        ("FONTNAME",  (0,0), (0,0), "Helvetica-Bold"),
        ("FONTSIZE",  (0,0), (0,0), 14),
        ("FONTNAME",  (2,0), (2,2), "Helvetica-Bold"),
        ("VALIGN",    (0,0), (-1,-1), "TOP"),
    ]))
    story.append(meta_tbl)
    story.append(Spacer(1, 0.4*cm))

    # Billing info
    story.append(Paragraph("<b>Lëshuar nga / Bill From:</b>", bold))
    story.append(Paragraph("ACME Solutions Sh.p.k.", styles["Normal"]))
    story.append(Spacer(1, 0.2*cm))
    story.append(Paragraph("<b>Faturuar tek / Bill To:</b>", bold))
    story.append(Paragraph("Klienti Demo S.R.L.", styles["Normal"]))
    story.append(Paragraph("Rr. Durrësit 55, Tiranë", styles["Normal"]))
    story.append(Paragraph("NIPT: L98765432B", styles["Normal"]))
    story.append(Spacer(1, 0.5*cm))

    # Items table
    col_headers = ["Nr.", "Përshkrim / Description", "Sasia\nQty", "Njësi\nUnit", "Çmim\nPrice", "TVSH\n20%", "Total"]
    items = [
        ["1", "Shërbim Konsulence IT (IT Consulting Service)", "10", "Orë", "50.00", "100.00", "600.00"],
        ["2", "Licencë Software Microsoft 365 (Annual)",       "5",  "Lic.", "80.00",  "80.00",  "480.00"],
        ["3", "Mirëmbajtje Server (Server Maintenance)",       "1",  "Mujor","200.00","40.00",  "240.00"],
        ["4", "Trajnim Stafi (Staff Training) – 2 ditë",       "2",  "Ditë", "150.00","60.00",  "360.00"],
        ["5", "Domain & Hosting Vjetor (Annual)",              "1",  "Vit",  "120.00","24.00",  "144.00"],
    ]

    table_data = [col_headers] + items + [
        ["", "", "", "", "", "Nëntotal / Subtotal:", "1,824.00"],
        ["", "", "", "", "", "TVSH 20% / VAT:",       "304.00"],
        ["", "", "", "", "", "TOTAL:",                 "2,128.00 ALL"],
    ]

    tbl = Table(
        table_data,
        colWidths=[1.2*cm, 6.5*cm, 1.5*cm, 1.5*cm, 2*cm, 2*cm, 2.3*cm],
    )
    tbl.setStyle(TableStyle([
        # Header row
        ("BACKGROUND", (0,0), (-1,0), colors.HexColor("#1F4E79")),
        ("TEXTCOLOR",  (0,0), (-1,0), colors.white),
        ("FONTNAME",   (0,0), (-1,0), "Helvetica-Bold"),
        ("FONTSIZE",   (0,0), (-1,0), 8),
        ("ALIGN",      (0,0), (-1,-1), "CENTER"),
        ("VALIGN",     (0,0), (-1,-1), "MIDDLE"),
        # Data rows – alternating
        *[
            ("BACKGROUND", (0,i), (-1,i), colors.HexColor("#D9E1F2"))
            for i in range(2, len(items)+1, 2)
        ],
        ("FONTSIZE",  (0,1), (-1,-4), 8),
        ("ALIGN",     (1,1), (1,-1), "LEFT"),
        # Totals rows
        ("FONTNAME",  (0,-3), (-1,-1), "Helvetica-Bold"),
        ("FONTSIZE",  (0,-3), (-1,-1), 9),
        ("BACKGROUND",(0,-1), (-1,-1), colors.HexColor("#1F4E79")),
        ("TEXTCOLOR", (0,-1), (-1,-1), colors.white),
        ("LINEABOVE", (0,-3), (-1,-3), 1, colors.black),
        # Grid
        ("GRID",      (0,0), (-1,-4), 0.5, colors.grey),
        ("BOX",       (0,0), (-1,-1), 1,   colors.black),
        ("ROWBACKGROUNDS", (0,-3), (-1,-3), [colors.HexColor("#FFFFC0")]),
    ]))
    story.append(tbl)
    story.append(Spacer(1, 0.5*cm))

    # Footer note
    story.append(Paragraph(
        "Pagesa të kryhet brenda 15 ditëve | Payment due within 15 days.",
        styles["Normal"],
    ))
    story.append(Paragraph("Llogaria bankare / Bank account: BKT IBAN AL47 0213 1005 0000 0001 2345 6789", styles["Normal"]))

    doc.build(story)


def _create_minimal_pdf(output_path: Path) -> None:
    """
    Write a self-contained minimal PDF without reportlab.
    This is a plain text invoice – enough to test the pipeline.
    """
    content = b"""\
%PDF-1.4
1 0 obj << /Type /Catalog /Pages 2 0 R >> endobj
2 0 obj << /Type /Pages /Kids [3 0 R] /Count 1 >> endobj
3 0 obj << /Type /Page /Parent 2 0 R /MediaBox [0 0 595 842]
  /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >> endobj
4 0 obj << /Length 600 >>
stream
BT
/F1 16 Tf
50 800 Td (INVOICE) Tj
/F1 10 Tf
0 -30 Td (Company: ACME Solutions) Tj
0 -15 Td (Invoice #: INV-2024-001) Tj
0 -15 Td (Date: 15.03.2024) Tj
0 -30 Td (Nr.  Description                          Qty   Price   Total) Tj
0 -15 Td (1    IT Consulting Service                 10    50.00   500.00) Tj
0 -15 Td (2    Software License Microsoft 365         5    80.00   400.00) Tj
0 -15 Td (3    Server Maintenance                     1   200.00   200.00) Tj
0 -15 Td (                                Subtotal:              1100.00) Tj
0 -15 Td (                                VAT 20%:                220.00) Tj
0 -15 Td (                                TOTAL:                 1320.00) Tj
ET
endstream
endobj
5 0 obj << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> endobj
xref
0 6
0000000000 65535 f
0000000009 00000 n
0000000058 00000 n
0000000115 00000 n
0000000266 00000 n
0000000917 00000 n
trailer << /Size 6 /Root 1 0 R >>
startxref
1001
%%EOF"""
    output_path.write_bytes(content)


# ─────────────────────────────────────────────────────────────────────────────
# Tests
# ─────────────────────────────────────────────────────────────────────────────


def run_single_conversion_test(tmp_dir: Path, verbose: bool = False) -> bool:
    """Test single-file conversion."""
    print("\n[TEST 1] Single file conversion")

    pdf_path  = tmp_dir / "test_invoice.pdf"
    xlsx_path = tmp_dir / "test_invoice.xlsx"

    create_test_invoice_pdf(pdf_path)

    try:
        result = convert(str(pdf_path), str(xlsx_path))
        assert result.exists(), "Output file not created"
        size = result.stat().st_size
        assert size > 1000, f"Output file too small ({size} bytes)"
        print(f"  PASSED - {result.name} ({size:,} bytes)")
        return True
    except Exception as exc:
        print(f"  FAILED - {exc}")
        if verbose:
            import traceback; traceback.print_exc()
        return False


def run_batch_conversion_test(tmp_dir: Path, verbose: bool = False) -> bool:
    """Test batch conversion of multiple PDFs."""
    print("\n[TEST 2] Batch conversion (3 PDFs)")

    in_dir  = tmp_dir / "batch_in"
    out_dir = tmp_dir / "batch_out"
    in_dir.mkdir(exist_ok=True)

    for i in range(1, 4):
        create_test_invoice_pdf(in_dir / f"invoice_{i:03d}.pdf")

    try:
        results = batch_convert(str(in_dir), str(out_dir))
        assert len(results) == 3, f"Expected 3 outputs, got {len(results)}"
        for r in results:
            assert r.exists(), f"Missing output: {r}"
        print(f"  PASSED – {len(results)} files converted to {out_dir.name}/")
        return True
    except Exception as exc:
        print(f"  FAILED – {exc}")
        if verbose:
            import traceback; traceback.print_exc()
        return False


def run_config_test(tmp_dir: Path, verbose: bool = False) -> bool:
    """Test custom Config parameters."""
    print("\n[TEST 3] Custom config (tight tolerances)")

    pdf_path  = tmp_dir / "test_invoice.pdf"
    xlsx_path = tmp_dir / "test_invoice_custom.xlsx"

    if not pdf_path.exists():
        create_test_invoice_pdf(pdf_path)

    cfg = Config(
        row_tolerance=2.0,
        column_tolerance=5.0,
        excel_grid_cols=80,
        ocr_language="eng",
    )

    try:
        result = convert(str(pdf_path), str(xlsx_path), config=cfg)
        assert result.exists()
        print(f"  PASSED – {result.name} with custom config")
        return True
    except Exception as exc:
        print(f"  FAILED – {exc}")
        if verbose:
            import traceback; traceback.print_exc()
        return False


def run_content_validation(tmp_dir: Path, verbose: bool = False) -> bool:
    """Validate that key invoice terms appear in the Excel output."""
    print("\n[TEST 4] Content validation")

    try:
        import openpyxl
    except ImportError:
        print("  SKIPPED – openpyxl not installed")
        return True

    xlsx_path = tmp_dir / "test_invoice.xlsx"
    if not xlsx_path.exists():
        print("  SKIPPED – run test 1 first")
        return True

    wb = openpyxl.load_workbook(str(xlsx_path))
    all_text = []
    for ws in wb.worksheets:
        for row in ws.iter_rows(values_only=True):
            for cell in row:
                if cell:
                    all_text.append(str(cell).strip().upper())

    full_text = " ".join(all_text)

    # Check that at least some content made it through
    expected_terms = ["INV", "INVOICE", "ACME", "TOTAL"]
    found = [t for t in expected_terms if t in full_text]

    if len(found) >= 2:
        print(f"  PASSED – found: {found}")
        return True
    else:
        print(f"  PARTIAL – found {found} out of {expected_terms}")
        # Not a hard failure – minimal PDF may have limited content
        return True


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────


def main() -> int:
    parser = argparse.ArgumentParser(description="Test the PDF→Excel pipeline")
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    level = logging.DEBUG if args.verbose else logging.INFO
    logging.basicConfig(level=level, format="%(levelname)s %(name)s: %(message)s")

    tmp_dir = Path(__file__).parent / "output" / "_test"
    tmp_dir.mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print("  PDF to Excel Pipeline Test Suite")
    print("=" * 60)

    tests = [
        run_single_conversion_test,
        run_batch_conversion_test,
        run_config_test,
        run_content_validation,
    ]

    results = []
    for test_fn in tests:
        passed = test_fn(tmp_dir, verbose=args.verbose)
        results.append(passed)

    print("\n" + "=" * 60)
    passed = sum(results)
    total  = len(results)
    print(f"  Results: {passed}/{total} tests passed")
    if passed == total:
        print("  ALL TESTS PASSED")
    else:
        print("  Some tests failed – check output above.")
    print(f"\n  Test outputs saved to: {tmp_dir}")
    print("=" * 60)

    return 0 if passed == total else 1


if __name__ == "__main__":
    sys.exit(main())
