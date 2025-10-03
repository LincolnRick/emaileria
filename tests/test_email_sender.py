from pathlib import Path
import sys

import pandas as pd
from jinja2 import Template

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import email_sender


def test_load_contacts_preserves_optional_case(tmp_path, monkeypatch):
    contacts = pd.DataFrame(
        [
            {
                "Email": "user@example.com",
                "Tratamento": "Sr.",
                "NOME": "João",
                "DEPARTAMENTO": "Financeiro",
                "InfoExtra": "VIP",
            }
        ]
    )
    excel_path = tmp_path / "contacts.xlsx"
    def fake_read_excel(path, sheet_name=None):
        assert path == excel_path
        assert sheet_name is None
        return contacts.copy()

    monkeypatch.setattr(email_sender.pd, "read_excel", fake_read_excel)

    loaded = email_sender.load_contacts(excel_path)

    assert set(loaded.columns) == {
        "email",
        "tratamento",
        "nome",
        "DEPARTAMENTO",
        "InfoExtra",
    }

    contexts = []

    def render_with_capture(template: str, context):
        contexts.append(dict(context))
        return Template(template).render(**context)

    monkeypatch.setattr(email_sender, "render_template", render_with_capture)

    subject_template = "{{ tratamento }} {{ nome }} - {{ DEPARTAMENTO }}"
    body_template = "Olá {{ tratamento }} {{ nome }}, setor: {{ DEPARTAMENTO }} ({{ InfoExtra }})"

    email_sender.send_messages(
        smtp=None,
        sender="sender@example.com",
        contacts=loaded.to_dict(orient="records"),
        subject_template=subject_template,
        body_template=body_template,
        dry_run=True,
    )

    assert contexts, "Expected at least one rendered context"
    context = contexts[0]
    assert context["email"] == "user@example.com"
    assert context["tratamento"] == "Sr."
    assert context["nome"] == "João"
    # Optional columns should keep their original case.
    assert context["DEPARTAMENTO"] == "Financeiro"
    assert context["InfoExtra"] == "VIP"

    rendered_body = Template(body_template).render(**context)
    assert "Financeiro" in rendered_body
    assert "VIP" in rendered_body
