"""Orchestration logic for sending templated emails."""

from __future__ import annotations

import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Dict, Iterable, List

import pandas as pd

from .providers.base import EmailProvider, ResultadoEnvio
from .templating import render

logger = logging.getLogger(__name__)

REQUIRED_KEYS = {"email", "tratamento", "nome"}


def _prepare_context(row: Dict[str, object]) -> Dict[str, str]:
    context: Dict[str, str] = {}
    for key, value in row.items():
        normalized_value = "" if pd.isna(value) else str(value)
        lowercase_key = key.lower()
        if lowercase_key in REQUIRED_KEYS:
            context[lowercase_key] = normalized_value
        else:
            context[key] = normalized_value
    missing_keys = REQUIRED_KEYS - context.keys()
    if missing_keys:
        raise KeyError(
            "Missing required contact data: " + ", ".join(sorted(missing_keys))
        )
    return context


def _create_message(sender: str, recipient: str, subject: str, body_html: str) -> MIMEMultipart:
    message = MIMEMultipart("alternative")
    message["To"] = recipient
    message["From"] = sender
    message["Subject"] = subject
    message.attach(MIMEText(body_html, "html", "utf-8"))
    return message


def send_messages(
    *,
    sender: str,
    contacts: Iterable[Dict[str, object]],
    subject_template: str,
    body_template: str,
    provider: EmailProvider | None = None,
    dry_run: bool = False,
) -> List[ResultadoEnvio]:
    """Render and optionally send messages for every contact."""
    if not dry_run and provider is None:
        raise ValueError("provider is required when dry_run is False")

    results: List[ResultadoEnvio] = []

    for row in contacts:
        context = _prepare_context(row)
        subject, body = render(subject_template, body_template, context)
        logger.info("Prepared email to %s with subject '%s'", context["email"], subject)

        if dry_run or provider is None:
            results.append(
                ResultadoEnvio(destinatario=context["email"], sucesso=True)
            )
            continue

        message = _create_message(sender, context["email"], subject, body)
        result = provider.send(message)
        results.append(result)

    return results
