"""Utilities for rendering email templates."""

from __future__ import annotations

import re
from datetime import date, datetime
from typing import Dict, Mapping, Tuple

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


def extract_placeholders(text: str) -> set[str]:
    """Extract placeholder names from template-like text."""

    return set(re.findall(r"{{\s*([a-zA-Z0-9_]+)\s*}}", text or ""))


def _datefmt(value: object, fmt: str = "%Y-%m-%d") -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime(fmt)
    return str(value)


_env.filters["datefmt"] = _datefmt


def _global_context() -> Dict[str, object]:
    now = datetime.now()
    today = date.today()
    return {
        "now": now,
        "hoje": today,
        "data_envio": today.strftime("%Y-%m-%d"),
        "hora_envio": now.strftime("%H:%M"),
    }


def _extract_placeholder_name(error: UndefinedError) -> str:
    match = re.search(_PLACEHOLDER_PATTERN, str(error))
    if match:
        return match.group(1)
    return str(error)


def _render_template(
    template: str, context: Mapping[str, object], template_type: str
) -> str:
    try:
        return _env.from_string(template).render(**context)
    except UndefinedError as exc:  # pragma: no cover - defensive parsing
        placeholder = _extract_placeholder_name(exc)
        raise TemplateRenderingError(template_type, placeholder, exc) from exc


def render(
    subject_template: str, body_template: str, context: Mapping[str, object]
) -> Tuple[str, str]:
    """Render subject and body templates with the provided context."""
    merged_context: Dict[str, object] = {**_global_context(), **dict(context)}
    subject = _render_template(subject_template, merged_context, "assunto")
    body = _render_template(body_template, merged_context, "corpo")
    return subject, body
