from pathlib import Path
import logging
import sys

import pandas as pd
from jinja2 import Template

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from emaileria.datasource.excel import load_contacts
import emaileria.cli as cli_module
import emaileria.sender as sender_module


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

    monkeypatch.setattr(pd, "read_excel", fake_read_excel)

    loaded = load_contacts(excel_path)

    assert set(loaded.columns) == {
        "email",
        "tratamento",
        "nome",
        "DEPARTAMENTO",
        "InfoExtra",
    }

    contexts = []

    def render_with_capture(subject_template: str, body_template: str, context):
        contexts.append(dict(context))
        return Template(subject_template).render(**context), Template(body_template).render(
            **context
        )

    monkeypatch.setattr(sender_module, "render", render_with_capture)

    subject_template = "{{ tratamento }} {{ nome }} - {{ DEPARTAMENTO }}"
    body_template = "Olá {{ tratamento }} {{ nome }}, setor: {{ DEPARTAMENTO }} ({{ InfoExtra }})"

    sender_module.send_messages(
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


def test_send_messages_retries_temporary_errors(monkeypatch):
    attempts: list[int] = []

    class DummyProvider:
        def __init__(self) -> None:
            self.calls = 0

        def send(self, message):
            self.calls += 1
            attempts.append(self.calls)
            if self.calls < 4:
                return sender_module.ResultadoEnvio(
                    destinatario=message["To"],
                    sucesso=False,
                    erro="451 4.3.0 Temporary local problem",
                )
            return sender_module.ResultadoEnvio(
                destinatario=message["To"], sucesso=True
            )

    class DummyLimiter:
        def __init__(self) -> None:
            self.calls = 0

        def acquire(self) -> None:
            self.calls += 1

    limiter = DummyLimiter()

    monkeypatch.setattr(sender_module, "_get_rate_limiter", lambda: limiter)

    sleeps: list[float] = []
    monkeypatch.setattr(sender_module.time, "sleep", lambda seconds: sleeps.append(seconds))

    results = sender_module.send_messages(
        sender="sender@example.com",
        contacts=[{"email": "user@example.com", "tratamento": "Sr.", "nome": "João"}],
        subject_template="Olá {{ nome }}",
        body_template="Olá {{ nome }}",
        provider=DummyProvider(),
        dry_run=False,
    )

    assert len(results) == 1
    assert results[0].sucesso is True
    assert attempts == [1, 2, 3, 4]
    assert sleeps == [1, 2, 4]
    assert limiter.calls == 4


def test_send_messages_does_not_retry_non_temporary_error(monkeypatch):
    class DummyProvider:
        def __init__(self) -> None:
            self.calls = 0

        def send(self, message):
            self.calls += 1
            return sender_module.ResultadoEnvio(
                destinatario=message["To"],
                sucesso=False,
                erro="550 Permanent failure",
            )

    limiter_calls: list[int] = []

    class DummyLimiter:
        def acquire(self) -> None:
            limiter_calls.append(1)

    monkeypatch.setattr(sender_module, "_get_rate_limiter", lambda: DummyLimiter())
    monkeypatch.setattr(sender_module.time, "sleep", lambda seconds: (_ for _ in ()).throw(AssertionError("sleep called")))

    results = sender_module.send_messages(
        sender="sender@example.com",
        contacts=[{"email": "user@example.com", "tratamento": "Sr.", "nome": "João"}],
        subject_template="Olá {{ nome }}",
        body_template="Olá {{ nome }}",
        provider=DummyProvider(),
        dry_run=False,
    )

    assert len(results) == 1
    assert results[0].sucesso is False
    assert results[0].erro == "550 Permanent failure"
    assert limiter_calls == [1]


def test_cli_dry_run_logs_only_summary(tmp_path, monkeypatch, caplog):
    caplog.set_level(logging.DEBUG)

    excel_path = tmp_path / "contacts.xlsx"
    excel_path.write_text("placeholder", encoding="utf-8")

    contacts_df = pd.DataFrame(
        [
            {
                "email": "joao@example.com",
                "nome": "João",
            }
        ]
    )

    monkeypatch.setattr(
        cli_module,
        "_load_contacts",
        lambda path, sheet=None: contacts_df.copy(),
    )

    monkeypatch.setattr(
        cli_module,
        "send_messages",
        lambda **kwargs: [
            sender_module.ResultadoEnvio(
                destinatario="joao@example.com",
                sucesso=True,
                assunto="Sensitive subject João",
            )
        ],
    )

    cli_module.main(
        [
            str(excel_path),
            "--sender",
            "sender@example.com",
            "--smtp-password",
            "hunter2",
            "--subject-template",
            "Sensitive subject {{ nome }}",
            "--body-template",
            "<p>Olá {{ nome }}</p>",
            "--dry-run",
        ]
    )

    debug_messages = [record.getMessage() for record in caplog.records if record.levelno == logging.DEBUG]
    assert any("Resumo do dry-run" in message for message in debug_messages)
    for message in debug_messages:
        assert "Sensitive subject" not in message
        assert "hunter2" not in message

    summary_message = next(message for message in debug_messages if "Resumo do dry-run" in message)
    assert "total=1" in summary_message
    assert "sucesso=1" in summary_message
