"""Data source helpers for loading contact spreadsheets."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd

REQUIRED_COLUMNS = {"email", "tratamento", "nome"}


def _normalize_required_columns(columns: Iterable[str]) -> dict[str, str]:
    lower_map = {column.lower(): column for column in columns}
    missing = REQUIRED_COLUMNS - lower_map.keys()
    if missing:
        raise ValueError(
            "Missing required columns: " + ", ".join(sorted(missing))
        )
    return {lower_map[column]: column for column in REQUIRED_COLUMNS}


def load_contacts(path: Path, sheet: str | None = None) -> pd.DataFrame:
    """Load contacts from an XLSX or CSV file, normalizing required headers."""
    path = Path(path)

    if path.suffix.lower() == ".csv":
        data = pd.read_csv(path)
    else:
        data = pd.read_excel(path, sheet_name=sheet)

    rename_map = _normalize_required_columns(data.columns)
    return data.rename(columns=rename_map)
