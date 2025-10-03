"""Utilities for rendering email templates."""

from __future__ import annotations

import re
from typing import Dict, Tuple

from jinja2 import Environment, StrictUndefined, UndefinedError

_PLACEHOLDER_PATTERN = r"'(.+?)' is undefined"


class TemplateRenderingError(RuntimeError):
    """Exception raised when a template cannot be rendered due to missing data."""

    def __init__(self, template_type: str, placeholder: str, original: Exception) -> None:
        self.template_type = template_type
        self.placeholder = placeholder
        self.original = original
        message = (
            f"Placeholder '{placeholder}' ausente ao renderizar template de {template_type}."
        )
        super().__init__(message)


_env = Environment(autoescape=False, undefined=StrictUndefined)


def _extract_placeholder_name(error: UndefinedError) -> str:
    match = re.search(_PLACEHOLDER_PATTERN, str(error))
    if match:
        return match.group(1)
    return str(error)


def _render_template(template: str, context: Dict[str, str], template_type: str) -> str:
    try:
        return _env.from_string(template).render(**context)
    except UndefinedError as exc:  # pragma: no cover - defensive parsing
        placeholder = _extract_placeholder_name(exc)
        raise TemplateRenderingError(template_type, placeholder, exc) from exc


def render(subject_template: str, body_template: str, context: Dict[str, str]) -> Tuple[str, str]:
    """Render subject and body templates with the provided context."""
    subject = _render_template(subject_template, context, "assunto")
    body = _render_template(body_template, context, "corpo")
    return subject, body
