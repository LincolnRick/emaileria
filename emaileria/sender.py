"""Orchestration logic for sending templated emails."""

from __future__ import annotations

import logging
import os
import re
import threading
import time
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from typing import Dict, Iterable, List, Optional, Sequence

import pandas as pd

from .providers.base import EmailProvider, ResultadoEnvio
from .templating import TemplateRenderingError, render

logger = logging.getLogger(__name__)

REQUIRED_KEYS = {"email", "tratamento", "nome"}

_RETRY_DELAYS_SECONDS = [0, 1, 2, 4]
_MAX_ATTEMPTS = len(_RETRY_DELAYS_SECONDS)
_TEMPORARY_SMTP_CODES = {"421", "450", "451", "452"}
_DEFAULT_RATE_LIMIT_PER_MINUTE = 80


class _TokenBucket:
    """Simple token bucket implementation for rate limiting."""

    def __init__(self, capacity: int, refill_interval: float = 60.0) -> None:
        self.capacity = capacity
        self.refill_interval = refill_interval
        self.tokens = float(capacity)
        self.refill_rate = capacity / refill_interval if refill_interval else 0.0
        self._lock = threading.Lock()
        self._last_refill = time.monotonic()

    def _refill(self) -> None:
        now = time.monotonic()
        elapsed = now - self._last_refill
        if elapsed <= 0:
            return
        tokens_to_add = elapsed * self.refill_rate
        if tokens_to_add <= 0:
            return
        self.tokens = min(self.capacity, self.tokens + tokens_to_add)
        self._last_refill = now

    def acquire(self, tokens: int = 1) -> None:
        if self.capacity <= 0 or self.refill_rate <= 0:
            return

        while True:
            with self._lock:
                self._refill()
                if self.tokens >= tokens:
                    self.tokens -= tokens
                    return
                missing_tokens = tokens - self.tokens
                wait_time = missing_tokens / self.refill_rate if self.refill_rate else 0.0
            if wait_time > 0:
                logger.debug(
                    "Rate limit reached, sleeping for %.2fs before sending next email",
                    wait_time,
                )
                time.sleep(wait_time)


def _get_rate_limiter() -> Optional[_TokenBucket]:
    env_value = os.getenv("RATE_LIMIT_PER_MINUTE")

    if env_value is None or env_value == "":
        rate_limit = _DEFAULT_RATE_LIMIT_PER_MINUTE
    else:
        try:
            rate_limit = int(env_value)
        except ValueError:
            logger.warning(
                "Invalid RATE_LIMIT_PER_MINUTE value '%s'. Falling back to default (%s).",
                env_value,
                _DEFAULT_RATE_LIMIT_PER_MINUTE,
            )
            rate_limit = _DEFAULT_RATE_LIMIT_PER_MINUTE

    if rate_limit <= 0:
        logger.info("RATE_LIMIT_PER_MINUTE <= 0, disabling rate limiting.")
        return None

    return _TokenBucket(rate_limit)


def _is_temporary_error(error: Optional[str]) -> bool:
    if not error:
        return False

    lower_error = error.lower()
    if "timeout" in lower_error or "timed out" in lower_error:
        return True

    if "4xx" in lower_error:
        return True

    if re.search(r"\b4\.\d\.\d\b", error):
        return True

    for match in re.findall(r"\b4\d{2}\b", error):
        if match in _TEMPORARY_SMTP_CODES:
            return True

    return False


def _safe_send(provider: EmailProvider, message: MIMEMultipart) -> ResultadoEnvio:
    try:
        return provider.send(message)
    except Exception as exc:  # pragma: no cover - provider specific failures
        to_address = message["To"]
        logger.exception("Unexpected exception while sending to %s", to_address)
        return ResultadoEnvio(destinatario=to_address, sucesso=False, erro=str(exc))


def _send_with_retries(
    provider: EmailProvider,
    message: MIMEMultipart,
    rate_limiter: Optional[_TokenBucket],
) -> ResultadoEnvio:
    to_address = message["To"]

    for attempt in range(1, _MAX_ATTEMPTS + 1):
        if rate_limiter is not None:
            rate_limiter.acquire()

        logger.info("Sending email to %s (attempt %s/%s)", to_address, attempt, _MAX_ATTEMPTS)
        result = _safe_send(provider, message)
        result.tentativas = attempt

        if result.sucesso:
            return result

        error_message = result.erro or "Unknown error"
        logger.error(
            "Attempt %s to send email to %s failed: %s", attempt, to_address, error_message
        )

        if attempt == _MAX_ATTEMPTS:
            return result

        if not _is_temporary_error(result.erro):
            return result

        sleep_time = _RETRY_DELAYS_SECONDS[attempt]
        if sleep_time > 0:
            logger.info(
                "Retrying email to %s in %ss due to temporary error.",
                to_address,
                sleep_time,
            )
            time.sleep(sleep_time)

    return ResultadoEnvio(
        destinatario=to_address,
        sucesso=False,
        erro="Unknown error",
        tentativas=_MAX_ATTEMPTS,
    )


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


def _create_message(
    sender: str,
    recipient: str,
    subject: str,
    body_html: str,
    *,
    cc: Sequence[str] | None = None,
    bcc: Sequence[str] | None = None,
    reply_to: str | None = None,
) -> MIMEMultipart:
    message = MIMEMultipart("alternative")
    message["To"] = recipient
    message["From"] = sender
    message["Subject"] = subject
    if cc:
        message["Cc"] = ", ".join(cc)
    if bcc:
        message["Bcc"] = ", ".join(bcc)
    if reply_to:
        message["Reply-To"] = reply_to
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
    allow_missing_fields: bool = False,
    interval_seconds: float = 0.0,
    cc: Sequence[str] | None = None,
    bcc: Sequence[str] | None = None,
    reply_to: str | None = None,
    cancel_event: threading.Event | None = None,
) -> List[ResultadoEnvio]:
    """Render and optionally send messages for every contact."""
    if not dry_run and provider is None:
        raise ValueError("provider is required when dry_run is False")

    rate_limiter: Optional[_TokenBucket] = None
    if not dry_run and provider is not None:
        rate_limiter = _get_rate_limiter()

    results: List[ResultadoEnvio] = []

    normalized_cc = [addr for addr in (cc or []) if addr]
    normalized_bcc = [addr for addr in (bcc or []) if addr]

    for index, row in enumerate(contacts, start=1):
        if cancel_event is not None and cancel_event.is_set():
            logger.info("Envio interrompido pelo usuário após %s registros.", index - 1)
            break
        row_position = index
        if isinstance(row, dict):
            if "__row_position__" in row:
                row_position = row.get("__row_position__", index)
                row_data = {k: v for k, v in row.items() if k != "__row_position__"}
            else:
                row_data = dict(row)
        else:
            row_data = dict(row)

        context = _prepare_context(row_data)

        missing_for_row: list[str] = []

        def _handle_missing(placeholder: str) -> None:
            missing_for_row.append(placeholder)

        render_kwargs = {}
        if allow_missing_fields:
            render_kwargs = {
                "allow_missing": True,
                "on_missing": _handle_missing,
            }
        try:
            subject, body = render(
                subject_template,
                body_template,
                context,
                **render_kwargs,
            )
        except TemplateRenderingError as exc:
            logger.error(
                "Falha ao renderizar %s na linha %s: placeholder '%s' não encontrado.",
                exc.template_type,
                row_position,
                exc.placeholder,
            )
            raise

        if allow_missing_fields and missing_for_row:
            for placeholder in sorted(set(missing_for_row)):
                logger.warning(
                    "Linha %s: placeholder '%s' ausente. Valor vazio utilizado.",
                    row_position,
                    placeholder,
                )

        logger.info("Prepared email to %s with subject '%s'", context["email"], subject)

        if dry_run or provider is None:
            result = ResultadoEnvio(
                destinatario=context["email"], sucesso=True, assunto=subject
            )
            results.append(result)
        else:
            message = _create_message(
                sender,
                context["email"],
                subject,
                body,
                cc=normalized_cc,
                bcc=normalized_bcc,
                reply_to=reply_to,
            )
            result = _send_with_retries(provider, message, rate_limiter)
            result.assunto = subject
            results.append(result)

        if interval_seconds > 0 and not (cancel_event is not None and cancel_event.is_set()):
            time.sleep(interval_seconds)

    return results
