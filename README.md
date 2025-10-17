# ppt2pdf

Convert Microsoft PowerPoint presentations into high-quality PDF documents using a
simple command line interface or a reusable Python API.

## Features

- ✅ Minimal dependency footprint – delegates to the battle-tested LibreOffice `soffice` CLI
- ✅ Works with both `.ppt` and `.pptx` files
- ✅ Usable as a CLI application *and* as an importable Python module
- ✅ Includes helpful error messages and type-annotated API for easy integration

## Requirements

- Python 3.9 or newer
- LibreOffice (provides the `soffice` binary used for conversion)

Ensure LibreOffice is installed and `soffice` is available on your system `PATH`. On
Linux this typically means installing the `libreoffice` package. On macOS the binary
is usually located at `/Applications/LibreOffice.app/Contents/MacOS/soffice`.

## Installation

```bash
pip install ppt2pdf
```

To work on the project locally, install the optional development dependencies:

```bash
pip install -e .[dev]
```

## Command Line Usage

Once installed, the `ppt2pdf` command becomes available:

```bash
ppt2pdf path/to/presentation.pptx
```

The resulting PDF is written next to the source file. Key options include:

| Option | Description |
| ------ | ----------- |
| `ppt2pdf INPUT OUTPUT` | Convert `INPUT` and save the PDF to `OUTPUT` |
| `--outdir DIR` | Place the output PDF inside `DIR` (ignored when `OUTPUT` provided) |
| `--soffice PATH` | Explicit path to the LibreOffice executable |
| `--timeout SECONDS` | Abort the conversion if LibreOffice runs longer than the limit |
| `--extra-arg VALUE` | Forward an additional argument to the `soffice` command. Repeat as needed. |

### Examples

Convert `deck.pptx` to `deck.pdf` in the same directory:

```bash
ppt2pdf deck.pptx
```

Save the PDF to a specific directory:

```bash
ppt2pdf deck.pptx --outdir build/pdf
```

Provide a custom output filename:

```bash
ppt2pdf deck.pptx reports/quarterly.pdf
```

Specify a custom LibreOffice binary:

```bash
ppt2pdf deck.pptx --soffice /Applications/LibreOffice.app/Contents/MacOS/soffice
```

## Python API Usage

```python
from pathlib import Path

from ppt2pdf import convert

result = convert("deck.pptx", output_dir="exports")
print(f"PDF created at {result.output_path}")
```

The returned [`ConversionResult`](src/ppt2pdf/converter.py) contains:

- `input_path`: `Path` pointing to the original presentation
- `output_path`: `Path` to the generated PDF file
- `command`: the exact command executed
- `stdout` / `stderr`: captured text output from LibreOffice

### Advanced Options

The `convert` helper supports additional keyword arguments:

```python
from ppt2pdf import convert

convert(
    "deck.pptx",
    output_file="reports/2024-q1.pdf",
    soffice_path="/usr/bin/soffice",
    timeout=120,
    extra_args=["--invisible"],
)
```

## Error Handling

- `LibreOfficeNotFoundError`: raised when the `soffice` executable cannot be located.
- `ConversionError`: raised when LibreOffice fails to produce a PDF or times out.
- `FileNotFoundError`: raised when the input presentation is missing.

Each exception includes a helpful message ready to display to end users.

## Development Workflow

1. Clone the repository and install the project in editable mode with development
   dependencies.
2. Run the test suite via `pytest`.
3. Format / lint as needed (type checking with `mypy` is recommended).

```bash
pip install -e .[dev]
pytest
```

Before publishing to PyPI, build the distribution artifacts:

```bash
python -m build
```

Then upload them using `twine`:

```bash
twine upload dist/*
```

## User Guide

### Quick Start

1. Install LibreOffice and ensure the `soffice` command is on your `PATH`.
2. Install `ppt2pdf` using `pip install ppt2pdf`.
3. Convert a presentation:
   - CLI: `ppt2pdf slides.pptx`
   - Python: `from ppt2pdf import convert; convert("slides.pptx")`

### Batch Conversion

Use a shell loop to convert multiple files:

```bash
for file in presentations/*.pptx; do
  ppt2pdf "$file" --outdir exports
done
```

Or via Python:

```python
from pathlib import Path
from ppt2pdf import convert

presentations = Path("presentations")
for ppt in presentations.glob("*.pptx"):
    convert(ppt, output_dir="exports")
```

### Troubleshooting

- **`soffice` not found** – Install LibreOffice or set `--soffice` to the binary path.
- **Conversion hangs** – Use `--timeout` to abort long-running jobs and inspect `stdout`/`stderr`.
- **PDF missing** – Ensure the destination directory is writable. The CLI prints the absolute
  path to the generated PDF for confirmation.

## License

This project is licensed under the terms of the MIT license. See [LICENSE](LICENSE) for details.
