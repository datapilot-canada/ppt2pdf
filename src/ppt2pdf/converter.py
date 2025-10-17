"""Conversion helpers for turning PowerPoint presentations into PDFs."""

from __future__ import annotations

import shutil
import subprocess
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional


class ConversionError(RuntimeError):
    """Raised when LibreOffice fails to convert the input file."""


class LibreOfficeNotFoundError(FileNotFoundError):
    """Raised when the ``soffice`` binary cannot be located."""


@dataclass(frozen=True)
class ConversionOptions:
    """Configuration controlling how the ``soffice`` command is invoked."""

    soffice_path: Path
    timeout: Optional[float] = None
    extra_args: Optional[Iterable[str]] = None


@dataclass(frozen=True)
class ConversionResult:
    """Information about a successful conversion."""

    input_path: Path
    output_path: Path
    command: List[str]
    stdout: str
    stderr: str


def find_soffice(candidate: Optional[str] = None) -> Path:
    """Return the path to the LibreOffice ``soffice`` executable.

    Parameters
    ----------
    candidate:
        Optional path provided by the caller.  When ``None`` the system
        ``PATH`` is searched.
    """

    if candidate:
        path = Path(candidate).expanduser().resolve()
        if path.is_file():
            return path
    resolved = shutil.which("soffice")
    if resolved:
        return Path(resolved)
    raise LibreOfficeNotFoundError(
        "LibreOffice 'soffice' executable not found. Please install LibreOffice "
        "and ensure 'soffice' is available on your PATH."
    )


def _build_command(options: ConversionOptions, input_path: Path, output_dir: Path) -> List[str]:
    command = [
        str(options.soffice_path),
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--convert-to",
        "pdf",
        "--outdir",
        str(output_dir),
        str(input_path),
    ]
    if options.extra_args:
        command[1:1] = list(options.extra_args)
    return command


def convert(
    input_file: Path | str,
    output_file: Path | str | None = None,
    *,
    output_dir: Path | str | None = None,
    soffice_path: Path | str | None = None,
    timeout: Optional[float] = None,
    extra_args: Optional[Iterable[str]] = None,
) -> ConversionResult:
    """Convert ``input_file`` into a PDF document.

    Parameters
    ----------
    input_file:
        The PowerPoint presentation to convert.  ``.ppt`` and ``.pptx``
        files are supported by LibreOffice.
    output_file:
        Optional path to the output PDF.  When omitted the file is stored
        in ``output_dir`` with the same stem as the input.
    output_dir:
        Directory in which the PDF should be written.  Ignored when
        ``output_file`` is provided.
    soffice_path:
        Optional override for the ``soffice`` executable location.
    timeout:
        Maximum number of seconds to wait for LibreOffice to finish.
    extra_args:
        Additional command line arguments forwarded to ``soffice``.
    """

    input_path = Path(input_file).expanduser().resolve()
    if not input_path.exists():
        raise FileNotFoundError(f"Input file does not exist: {input_path}")

    if output_file and output_dir:
        raise ValueError("Specify either output_file or output_dir, not both.")

    soffice = find_soffice(str(soffice_path) if soffice_path else None)

    if output_file:
        output_path = Path(output_file).expanduser().resolve()
        output_path.parent.mkdir(parents=True, exist_ok=True)
        target_dir = output_path.parent
    else:
        target_dir = Path(output_dir).expanduser().resolve() if output_dir else input_path.parent
        target_dir.mkdir(parents=True, exist_ok=True)
        output_path = target_dir / (input_path.stem + ".pdf")

    options = ConversionOptions(
        soffice_path=soffice,
        timeout=timeout,
        extra_args=extra_args,
    )

    command = _build_command(options, input_path, target_dir)

    try:
        completed = subprocess.run(
            command,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            check=True,
            timeout=timeout,
            text=True,
        )
    except subprocess.TimeoutExpired as exc:  # pragma: no cover - passthrough
        raise ConversionError("LibreOffice timed out while converting the presentation.") from exc
    except subprocess.CalledProcessError as exc:
        stdout = getattr(exc, "stdout", getattr(exc, "output", ""))
        stderr = getattr(exc, "stderr", "")
        raise ConversionError(
            "LibreOffice failed to convert the presentation.\n"
            f"Command: {' '.join(command)}\n"
            f"Stdout: {stdout}\n"
            f"Stderr: {stderr}"
        ) from exc

    if not output_path.exists():
        # LibreOffice may save using the original name when converting from templates.
        candidate = target_dir / (input_path.stem + ".pdf")
        if candidate.exists():
            output_path = candidate
        else:
            raise ConversionError(
                "LibreOffice reported success but the expected PDF was not created."
            )

    if output_file and output_path != Path(output_file).resolve():
        output_path.replace(Path(output_file).resolve())
        output_path = Path(output_file).resolve()

    return ConversionResult(
        input_path=input_path,
        output_path=output_path,
        command=command,
        stdout=completed.stdout,
        stderr=completed.stderr,
    )


__all__ = [
    "ConversionError",
    "LibreOfficeNotFoundError",
    "ConversionOptions",
    "ConversionResult",
    "convert",
]
