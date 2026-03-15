"""
invoice_parser – PDF-to-Excel layout-preserving conversion engine.

Public API
----------
    from invoice_parser import convert, batch_convert, Config

    convert("invoice.pdf", "invoice.xlsx")
    batch_convert("input/", "output/")
"""

from .config import Config, DEFAULT_CONFIG
from .main import convert, batch_convert

__all__ = ["Config", "DEFAULT_CONFIG", "convert", "batch_convert"]
__version__ = "1.0.0"
