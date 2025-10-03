"""Data source helpers for loading contact spreadsheets."""

from __future__ import annotations

from pathlib import Path
from typing import Iterable

import pandas as pd

REQUIRED_COLUMNS = {"email", "tratamento", "nome"}


def _clean_column_name(column: str) -> str:
    return str(column).strip()


def _normalize_required_columns(columns: Iterable[str]) -> dict[str, str]:
    cleaned = [_clean_column_name(column) for column in columns]
    lower_map = {column.lower(): column for column in cleaned}
    missing = REQUIRED_COLUMNS - lower_map.keys()
    if missing:
        required_list = ", ".join(sorted(REQUIRED_COLUMNS))
        missing_list = ", ".join(sorted(missing))
        raise ValueError(
            "Planilha inválida: colunas obrigatórias ausentes: "
            f"{missing_list}. Certifique-se de incluir pelo menos: {required_list}."
        )
    return {lower_map[column]: column for column in REQUIRED_COLUMNS}


def load_contacts(path: Path, sheet: str | None = None) -> pd.DataFrame:
    """Load contacts from an XLSX or CSV file, normalizing required headers."""
    path = Path(path)

    if path.suffix.lower() == ".csv":
        data = pd.read_csv(path)
    else:
        data = pd.read_excel(path, sheet_name=sheet)

    cleaned_columns = {column: _clean_column_name(column) for column in data.columns}
    data = data.rename(columns=cleaned_columns)

    rename_map = _normalize_required_columns(data.columns)
    return data.rename(columns=rename_map)
