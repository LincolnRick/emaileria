"""Utilities for rendering email templates."""

from __future__ import annotations

from typing import Dict, Tuple

from jinja2 import Environment

_env = Environment(autoescape=False)


def render(subject_template: str, body_template: str, context: Dict[str, str]) -> Tuple[str, str]:
    """Render subject and body templates with the provided context."""
    subject = _env.from_string(subject_template).render(**context)
    body = _env.from_string(body_template).render(**context)
    return subject, body
