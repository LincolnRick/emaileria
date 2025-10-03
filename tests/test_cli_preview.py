"""Tests for the preview command in the CLI."""

from __future__ import annotations

from pathlib import Path
import sys

import pandas as pd
import pytest

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import emaileria.cli as cli_module


def _fake_contacts() -> pd.DataFrame:
    return pd.DataFrame(
        [
            {
                "email": "joao@example.com",
                "tratamento": "Sr.",
                "nome": "João",
                "departamento": "Financeiro",
            },
            {
                "email": "maria@example.com",
                "tratamento": "Sra.",
                "nome": "Maria",
                "departamento": "Comercial",
            },
        ]
    )


def test_preview_command_generates_gallery(tmp_path, monkeypatch):
    excel_path = tmp_path / "contacts.xlsx"
    excel_path.write_text("placeholder", encoding="utf-8")

    monkeypatch.chdir(tmp_path)

    contacts_df = _fake_contacts()

    def fake_load_contacts(path: Path, sheet: str | None = None):
        assert path == excel_path
        assert sheet is None
        return contacts_df.copy()

    monkeypatch.setattr(cli_module, "load_contacts", fake_load_contacts)

    subject_template = "Assunto para {{ nome }}"
    body_template = "<p>Olá {{ tratamento }} {{ nome }} do {{ departamento }}</p>"

    cli_module.main(
        [
            "preview",
            str(excel_path),
            "--subject-template",
            subject_template,
            "--body-template",
            body_template,
            "--limit",
            "1",
        ]
    )

    preview_dirs = list((tmp_path / "previews").iterdir())
    assert len(preview_dirs) == 1
    html_file = preview_dirs[0] / "index.html"
    assert html_file.exists()

    html_content = html_file.read_text(encoding="utf-8")
    assert "João" in html_content
    assert "Maria" not in html_content
    assert "Olá Sr. João" in html_content


def test_preview_limit_must_be_positive(monkeypatch):
    contacts = _fake_contacts().to_dict(orient="records")

    with pytest.raises(ValueError):
        cli_module._render_preview(
            contacts=contacts,
            subject_template="{{ nome }}",
            body_template="{{ nome }}",
            limit=0,
        )
