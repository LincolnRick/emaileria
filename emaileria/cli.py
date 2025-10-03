"""Command line interface for the Emaileria package."""

from __future__ import annotations

import argparse
import getpass
import logging
from pathlib import Path

from . import config
from .datasource.excel import load_contacts
from .providers.smtp import SMTPProvider
from .sender import send_messages


def _read_template(template: str | None, template_file: Path | None) -> str:
    if template_file is not None:
        return template_file.read_text(encoding="utf-8")
    if template is None:
        raise ValueError("Template must be provided when no template file is specified")
    return template


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Send Gmail messages from an Excel contact list."
    )
    parser.add_argument("excel", type=Path, help="Path to the Excel/CSV file containing contacts.")
    parser.add_argument("--sheet", help="Excel sheet name to read.")
    parser.add_argument("--sender", required=True, help="Email address of the sender.")
    parser.add_argument(
        "--smtp-user",
        help="SMTP username. Defaults to the sender address if omitted.",
    )
    parser.add_argument(
        "--smtp-password",
        help="SMTP password or app password. If omitted, it will be requested via prompt.",
    )
    parser.add_argument(
        "--subject-template",
        required=True,
        help="Template for the email subject. Jinja2 placeholders are allowed.",
    )
    parser.add_argument(
        "--subject-template-file",
        type=Path,
        help="Path to a file containing the subject template. Overrides --subject-template when provided.",
    )
    parser.add_argument(
        "--body-template",
        required=True,
        help="Template for the HTML body. Jinja2 placeholders are allowed.",
    )
    parser.add_argument(
        "--body-template-file",
        type=Path,
        help="Path to a file containing the body template. Overrides --body-template when provided.",
    )
    parser.add_argument("--dry-run", action="store_true", help="Render messages without sending them.")
    parser.add_argument(
        "--log-level", default="INFO", help="Logging level (DEBUG, INFO, WARNING, ERROR)."
    )
    return parser


def _load_contacts(path: Path, sheet: str | None) -> list[dict[str, object]]:
    dataframe = load_contacts(path, sheet)
    return dataframe.to_dict(orient="records")


def main(argv: list[str] | None = None) -> None:
    parser = build_parser()
    args = parser.parse_args(argv)

    logging.basicConfig(level=args.log_level.upper(), format="%(levelname)s: %(message)s")

    subject_template = _read_template(args.subject_template, args.subject_template_file)
    body_template = _read_template(args.body_template, args.body_template_file)

    contacts = _load_contacts(args.excel, args.sheet)

    smtp_user = args.smtp_user or args.sender

    if args.dry_run:
        send_messages(
            sender=args.sender,
            contacts=contacts,
            subject_template=subject_template,
            body_template=body_template,
            dry_run=True,
        )
        return

    smtp_password = args.smtp_password or getpass.getpass(
        prompt="SMTP password (app password recommended): "
    )

    with SMTPProvider(
        config.SMTP_HOST,
        config.SMTP_PORT,
        smtp_user,
        smtp_password,
        timeout=config.SMTP_TIMEOUT,
    ) as provider:
        send_messages(
            sender=args.sender,
            contacts=contacts,
            subject_template=subject_template,
            body_template=body_template,
            provider=provider,
            dry_run=False,
        )
