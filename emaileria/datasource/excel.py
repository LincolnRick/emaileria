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


def load_contacts(path: str | Path, sheet: str | None = None) -> pd.DataFrame:
    """Load contacts from an XLSX or CSV file, normalizing required headers."""
    file_path = Path(path)

    if not file_path.exists():
        raise FileNotFoundError(f"Arquivo não encontrado: {file_path}")

    if file_path.suffix.lower() == ".csv":
        data = pd.read_csv(file_path, dtype=str).fillna("")
    else:
        if sheet and sheet.strip():
            data = pd.read_excel(file_path, sheet_name=sheet.strip(), dtype=str).fillna("")
        else:
            with pd.ExcelFile(file_path) as workbook:
                if not workbook.sheet_names:
                    raise ValueError("Nenhuma aba encontrada no arquivo Excel.")
                first_sheet = workbook.sheet_names[0]
                data = pd.read_excel(workbook, sheet_name=first_sheet, dtype=str).fillna("")

    cleaned_columns = {column: _clean_column_name(column) for column in data.columns}
    data = data.rename(columns=cleaned_columns)

    rename_map = _normalize_required_columns(data.columns)
    return data.rename(columns=rename_map)
