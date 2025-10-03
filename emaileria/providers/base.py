"""Base interfaces for email sending providers."""

from __future__ import annotations

from dataclasses import dataclass
from typing import Protocol

from email.message import Message


@dataclass
class ResultadoEnvio:
    """Represents the outcome of sending a single email."""

    destinatario: str
    sucesso: bool
    erro: str | None = None
    tentativas: int = 1
    assunto: str | None = None


class EmailProvider(Protocol):
    """Protocol for email providers used by the sender orchestrator."""

    def send(self, message: Message) -> ResultadoEnvio:
        """Send a MIME message and return the delivery result."""

    def close(self) -> None:  # pragma: no cover - optional protocol method
        """Close any underlying resources."""
