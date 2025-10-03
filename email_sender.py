"""Compatibility module that exposes the Emaileria CLI entry point."""

from __future__ import annotations

import os
from typing import List, Optional

from emaileria.cli import main as _cli_main

RATE_LIMIT_PER_MINUTE = int(os.getenv("RATE_LIMIT_PER_MINUTE", "80"))
"""Limite padrão de envios por minuto usado pelo token bucket."""


def main(argv: Optional[List[str]] = None) -> None:
    """Entrypoint that delegates to :mod:`emaileria.cli` after setting defaults."""

    os.environ.setdefault("RATE_LIMIT_PER_MINUTE", str(RATE_LIMIT_PER_MINUTE))
    _cli_main(argv)


__all__ = ["main", "RATE_LIMIT_PER_MINUTE"]


if __name__ == "__main__":  # pragma: no cover - manual execution entry point
    main()
