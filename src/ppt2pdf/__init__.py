"""Top-level package for ppt2pdf.

This package exposes a simple API for converting Microsoft PowerPoint
presentations into PDF documents by delegating to LibreOffice's
``soffice`` command line interface.  The :func:`convert` helper is the
main entry point for library consumers.
"""

from .converter import (
    ConversionError,
    ConversionOptions,
    ConversionResult,
    LibreOfficeNotFoundError,
    convert,
)

__all__ = [
    "ConversionError",
    "ConversionOptions",
    "ConversionResult",
    "LibreOfficeNotFoundError",
    "convert",
]

__version__ = "0.1.0"
