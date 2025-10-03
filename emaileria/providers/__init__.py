"""Email sending providers."""

from .base import EmailProvider, ResultadoEnvio
from .smtp import SMTPProvider

__all__ = ["EmailProvider", "ResultadoEnvio", "SMTPProvider"]
