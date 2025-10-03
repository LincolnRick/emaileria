"""Utility helpers to summarize sending results."""

from __future__ import annotations

from collections import Counter
from typing import Iterable

from .providers.base import ResultadoEnvio


def summarize_results(results: Iterable[ResultadoEnvio]) -> dict[str, int]:
    """Return a simple summary counting successful and failed sends."""
    counter = Counter()
    for result in results:
        if result.sucesso:
            counter["success"] += 1
        else:
            counter["failure"] += 1
    return dict(counter)
