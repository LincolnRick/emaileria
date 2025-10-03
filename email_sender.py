"""Compatibility module that exposes the Emaileria CLI entry point."""

from __future__ import annotations

from emaileria.cli import main

__all__ = ["main"]


if __name__ == "__main__":
    main()
