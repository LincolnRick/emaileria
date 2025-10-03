"""Command line utility to send personalized Gmail messages from an Excel sheet.

The script expects a spreadsheet containing at least the columns:
- email: destination address
- tratamento: pronoun or greeting (ex: "Sr.")
- nome: recipient name

Extra columns are available as template variables when rendering the
message subject and body.

Authentication is performed using Gmail's SMTP server over SSL. The
user must provide their Gmail username and password (or app password).
"""
from __future__ import annotations

import argparse
import getpass
import logging
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Dict, Iterable, List, Optional

import pandas as pd
from jinja2 import Template

logger = logging.getLogger(__name__)


def load_contacts(excel_path: Path, sheet: str | None = None) -> pd.DataFrame:
    """Load the contacts spreadsheet."""
    logger.info("Loading contacts from %s", excel_path)
    data = pd.read_excel(excel_path, sheet_name=sheet)
    required_columns = {"email", "tratamento", "nome"}
    missing = required_columns - set(map(str.lower, data.columns))
    if missing:
        raise ValueError(
            "Missing required columns: " + ", ".join(sorted(missing))
        )
    # Normalize required column names to lower case for downstream usage
    data = data.rename(columns={col: col.lower() for col in data.columns})
    return data


def render_template(template: str, context: Dict[str, str]) -> str:
    return Template(template).render(**context)


def create_message(sender: str, to: str, subject: str, body_html: str) -> MIMEMultipart:
    """Compose a MIME email message."""
    message = MIMEMultipart("alternative")
    message["To"] = to
    message["From"] = sender
    message["Subject"] = subject

    part_html = MIMEText(body_html, "html", "utf-8")
    message.attach(part_html)
    return message


def send_messages(
    smtp: Optional[smtplib.SMTP],
    sender: str,
    contacts: Iterable[Dict[str, str]],
    subject_template: str,
    body_template: str,
    dry_run: bool = False,
) -> List[str]:
    """Send rendered messages to all contacts."""
    if not dry_run and smtp is None:
        raise ValueError("SMTP connection is required when not running in dry-run mode.")

    sent_to: List[str] = []
    for row in contacts:
        context = {k: str(v) if pd.notna(v) else "" for k, v in row.items()}
        subject = render_template(subject_template, context)
        body = render_template(body_template, context)
        message = create_message(
            sender=sender,
            to=context["email"],
            subject=subject,
            body_html=body,
        )

        logger.info("Prepared email to %s with subject '%s'", context["email"], subject)
        if dry_run:
            logger.debug("Dry run enabled, skipping send for %s", context["email"])
            continue

        assert smtp is not None  # for type checkers
        smtp.sendmail(sender, [context["email"]], message.as_string())
        sent_to.append(context["email"])
        logger.info("Sent message to %s", context["email"])
    return sent_to


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Send Gmail messages from an Excel contact list.")
    parser.add_argument("excel", type=Path, help="Path to the Excel file containing contacts.")
    parser.add_argument("--sheet", help="Excel sheet name to read.")
    parser.add_argument("--sender", required=True, help="Email address of the sender (must match Gmail account).")
    parser.add_argument(
        "--smtp-user",
        help="SMTP username. Defaults to the sender address if omitted.",
    )
    parser.add_argument(
        "--smtp-password",
        help="SMTP password or app password. If omitted, it will be requested via prompt.",
    )
    parser.add_argument("--subject-template", required=True, help="Template for the email subject. Jinja2 placeholders are allowed.")
    parser.add_argument("--body-template", required=True, help="Template for the HTML body. Jinja2 placeholders are allowed.")
    parser.add_argument("--dry-run", action="store_true", help="Render messages without sending them.")
    parser.add_argument("--log-level", default="INFO", help="Logging level (DEBUG, INFO, WARNING, ERROR).")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    logging.basicConfig(level=args.log_level.upper(), format="%(levelname)s: %(message)s")

    contacts_df = load_contacts(args.excel, args.sheet)
    contacts_records = contacts_df.to_dict(orient="records")

    smtp_user = args.smtp_user or args.sender

    if args.dry_run:
        send_messages(
            smtp=None,
            sender=args.sender,
            contacts=contacts_records,
            subject_template=args.subject_template,
            body_template=args.body_template,
            dry_run=True,
        )
        return

    smtp_password = args.smtp_password or getpass.getpass(
        prompt="SMTP password (app password recommended): ",
    )

    logger.info("Connecting to smtp.gmail.com:465 as %s", smtp_user)
    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
        smtp.login(smtp_user, smtp_password)
        send_messages(
            smtp=smtp,
            sender=args.sender,
            contacts=contacts_records,
            subject_template=args.subject_template,
            body_template=args.body_template,
            dry_run=False,
        )


if __name__ == "__main__":
    main()
