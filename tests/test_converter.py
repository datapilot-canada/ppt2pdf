from __future__ import annotations

import subprocess
from pathlib import Path
from typing import List

import pytest

from ppt2pdf import converter
from ppt2pdf.converter import ConversionError, ConversionResult, convert


class DummyCompletedProcess(subprocess.CompletedProcess[str]):
    def __init__(self, args: List[str]):
        super().__init__(args=args, returncode=0, stdout="ok", stderr="")


def test_convert_writes_to_output_dir(tmp_path, monkeypatch):
    soffice = tmp_path / "soffice"
    soffice.write_text("")

    input_file = tmp_path / "slides.pptx"
    input_file.write_text("dummy")

    output_dir = tmp_path / "out"

    def fake_find_soffice(candidate: str | None = None) -> Path:
        return soffice

    def fake_run(command, **kwargs):
        pdf_path = output_dir / "slides.pdf"
        pdf_path.write_text("pdf")
        return DummyCompletedProcess(command)

    monkeypatch.setattr(converter, "find_soffice", fake_find_soffice)
    monkeypatch.setattr(converter.subprocess, "run", fake_run)

    result = convert(input_file, output_dir=output_dir)

    assert isinstance(result, ConversionResult)
    assert result.output_path == output_dir / "slides.pdf"
    assert "--headless" in result.command


def test_convert_renames_when_output_file_specified(tmp_path, monkeypatch):
    soffice = tmp_path / "soffice"
    soffice.write_text("")

    input_file = tmp_path / "slides.pptx"
    input_file.write_text("dummy")

    target_file = tmp_path / "custom" / "presentation.pdf"

    def fake_find_soffice(candidate: str | None = None) -> Path:
        return soffice

    def fake_run(command, **kwargs):
        default_pdf = tmp_path / "custom" / "slides.pdf"
        default_pdf.parent.mkdir(parents=True, exist_ok=True)
        default_pdf.write_text("pdf")
        return DummyCompletedProcess(command)

    monkeypatch.setattr(converter, "find_soffice", fake_find_soffice)
    monkeypatch.setattr(converter.subprocess, "run", fake_run)

    result = convert(input_file, output_file=target_file)

    assert result.output_path == target_file
    assert target_file.read_text() == "pdf"


def test_convert_rejects_conflicting_output_arguments(tmp_path):
    input_file = tmp_path / "slides.pptx"
    input_file.write_text("dummy")

    with pytest.raises(ValueError):
        convert(input_file, output_file=tmp_path / "slides.pdf", output_dir=tmp_path)


def test_convert_raises_when_command_fails(tmp_path, monkeypatch):
    soffice = tmp_path / "soffice"
    soffice.write_text("")

    input_file = tmp_path / "slides.ppt"
    input_file.write_text("dummy")

    def fake_find_soffice(candidate: str | None = None) -> Path:
        return soffice

    def fake_run(command, **kwargs):
        raise subprocess.CalledProcessError(1, command, output="", stderr="boom")

    monkeypatch.setattr(converter, "find_soffice", fake_find_soffice)
    monkeypatch.setattr(converter.subprocess, "run", fake_run)

    with pytest.raises(ConversionError):
        convert(input_file)
