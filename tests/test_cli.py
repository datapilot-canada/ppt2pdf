from __future__ import annotations

from pathlib import Path

import pytest

from ppt2pdf import cli


class DummyResult:
    def __init__(self, output_path: Path):
        self.output_path = output_path


def test_main_invokes_converter(monkeypatch, tmp_path, capsys):
    output = tmp_path / "slides.pdf"

    def fake_convert(**kwargs):
        assert kwargs["input_file"] == tmp_path / "slides.pptx"
        return DummyResult(output)

    monkeypatch.setattr(cli, "convert", fake_convert)

    exit_code = cli.main([str(tmp_path / "slides.pptx")])

    captured = capsys.readouterr()
    assert exit_code == 0
    assert captured.out.strip() == str(output)


def test_main_returns_error_code_on_failure(monkeypatch, tmp_path):
    class DummyError(Exception):
        pass

    def fake_convert(**kwargs):
        raise DummyError("boom")

    monkeypatch.setattr(cli, "convert", fake_convert)

    with pytest.raises(SystemExit) as excinfo:
        cli.main([str(tmp_path / "slides.pptx")])

    assert excinfo.value.code != 0
