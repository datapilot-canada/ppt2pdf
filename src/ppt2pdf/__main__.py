"""Allow ``python -m ppt2pdf`` to behave like the ``ppt2pdf`` CLI."""

from .cli import main


if __name__ == "__main__":  # pragma: no cover - CLI entry point
    raise SystemExit(main())
