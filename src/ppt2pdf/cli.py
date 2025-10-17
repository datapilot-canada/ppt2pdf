"""Command line interface for the ppt2pdf package."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import __version__
from .converter import ConversionError, LibreOfficeNotFoundError, convert


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="ppt2pdf",
        description="Convert Microsoft PowerPoint presentations into PDF documents.",
    )
    parser.add_argument("input", type=Path, help="Path to the .ppt or .pptx file to convert")
    parser.add_argument(
        "output",
        nargs="?",
        type=Path,
        help="Optional path to the generated PDF file (defaults to input stem).",
    )
    parser.add_argument(
        "--outdir",
        type=Path,
        help="Directory where the PDF should be created when --output is omitted.",
    )
    parser.add_argument(
        "--soffice",
        type=Path,
        help="Path to the LibreOffice 'soffice' executable (falls back to PATH lookup).",
    )
    parser.add_argument(
        "--timeout",
        type=float,
        help="Maximum number of seconds to wait for LibreOffice to finish.",
    )
    parser.add_argument(
        "--extra-arg",
        dest="extra_args",
        action="append",
        help="Additional option forwarded to the soffice command. Repeat for multiple arguments.",
    )
    parser.add_argument(
        "--version",
        action="version",
        version=f"ppt2pdf {__version__}",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    parser = _build_parser()
    args = parser.parse_args(argv)

    try:
        result = convert(
            input_file=args.input,
            output_file=args.output,
            output_dir=args.outdir,
            soffice_path=args.soffice,
            timeout=args.timeout,
            extra_args=args.extra_args,
        )
    except (FileNotFoundError, LibreOfficeNotFoundError, ConversionError) as exc:
        parser.error(str(exc))
    except Exception as exc:  # pragma: no cover - defensive safety net
        parser.error(str(exc))

    print(result.output_path)
    return 0


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    sys.exit(main())
