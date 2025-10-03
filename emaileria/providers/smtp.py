"""SMTP implementation of the email provider."""

from __future__ import annotations

import logging
import smtplib
from email.message import Message
from typing import Optional

from .base import EmailProvider, ResultadoEnvio

logger = logging.getLogger(__name__)


class SMTPProvider(EmailProvider):
    """Email provider that sends messages using SMTP over SSL."""

    def __init__(
        self,
        host: str,
        port: int,
        username: str,
        password: str,
        *,
        timeout: Optional[int] = None,
    ) -> None:
        self.host = host
        self.port = port
        self.username = username
        self.password = password
        self.timeout = timeout
        self._smtp: smtplib.SMTP_SSL | None = None

    def __enter__(self) -> "SMTPProvider":
        self._ensure_connection()
        return self

    def __exit__(self, exc_type, exc, tb) -> None:
        self.close()

    def _ensure_connection(self) -> smtplib.SMTP_SSL:
        if self._smtp is None:
            logger.info("Connecting to %s:%s as %s", self.host, self.port, self.username)
            self._smtp = smtplib.SMTP_SSL(self.host, self.port, timeout=self.timeout)
            self._smtp.login(self.username, self.password)
        return self._smtp

    def send(self, message: Message) -> ResultadoEnvio:
        smtp = self._ensure_connection()
        to_address = message["To"]
        try:
            refused = smtp.sendmail(message["From"], [to_address], message.as_string())
        except Exception as exc:  # pragma: no cover - network failure
            logger.exception("Failed to send message to %s", to_address)
            return ResultadoEnvio(destinatario=to_address, sucesso=False, erro=str(exc))

        if refused:
            error_msg = str(refused.get(to_address))
            logger.error("SMTP server refused recipient %s: %s", to_address, error_msg)
            return ResultadoEnvio(destinatario=to_address, sucesso=False, erro=error_msg)

        logger.info("Sent message to %s", to_address)
        return ResultadoEnvio(destinatario=to_address, sucesso=True)

    def close(self) -> None:
        if self._smtp is not None:
            try:
                self._smtp.quit()
            finally:
                self._smtp = None
