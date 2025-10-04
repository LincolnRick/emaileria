"""Utilities for rendering email templates."""

from __future__ import annotations

import re
from datetime import date, datetime
from typing import Dict, Mapping, Tuple

from typing import Callable

from jinja2 import Environment, StrictUndefined, Undefined, UndefinedError

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


class SoftUndefined(Undefined):
    def _fail_with_undefined_error(self, *args, **kwargs):  # type: ignore[override]
        return ""


_STRICT_ENV = Environment(
    autoescape=False, undefined=StrictUndefined, keep_trailing_newline=True
)
_SOFT_ENV = Environment(autoescape=False, undefined=SoftUndefined, keep_trailing_newline=True)


def extract_placeholders(text: str) -> set[str]:
    """Extract placeholder names from template-like text."""

    return set(re.findall(r"{{\s*([a-zA-Z0-9_]+)\s*}}", text or ""))


def _datefmt(value: object, fmt: str = "%Y-%m-%d") -> str:
    if value is None:
        return ""
    if isinstance(value, (datetime, date)):
        return value.strftime(fmt)
    return str(value)


for _environment in (_STRICT_ENV, _SOFT_ENV):
    _environment.filters["datefmt"] = _datefmt


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
    template: str,
    context: Mapping[str, object],
    template_type: str,
    *,
    environment: Environment,
) -> str:
    try:
        return environment.from_string(template).render(**context)
    except UndefinedError as exc:  # pragma: no cover - defensive parsing
        placeholder = _extract_placeholder_name(exc)
        raise TemplateRenderingError(template_type, placeholder, exc) from exc


def render(
    subject_template: str,
    body_template: str,
    context: Mapping[str, object],
    *,
    allow_missing: bool = False,
    on_missing: Callable[[str], None] | None = None,
) -> Tuple[str, str]:
    """Render subject and body templates with the provided context."""
    merged_context: Dict[str, object] = {**_global_context(), **dict(context)}
    environment = _SOFT_ENV if allow_missing else _STRICT_ENV

    missing_placeholders: set[str] = set()
    if allow_missing:
        normalized_keys = {str(key).strip().lower() for key in merged_context}
        used_placeholders = extract_placeholders(subject_template) | extract_placeholders(
            body_template
        )
        missing_placeholders = {
            placeholder
            for placeholder in used_placeholders
            if placeholder.lower() not in normalized_keys
        }

    subject = _render_template(
        subject_template, merged_context, "assunto", environment=environment
    )
    body = _render_template(
        body_template, merged_context, "corpo", environment=environment
    )

    if allow_missing and on_missing is not None:
        for placeholder in sorted(missing_placeholders):
            on_missing(placeholder)

    return subject, body
