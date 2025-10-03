"""Command line utility to send personalized Gmail messages from an Excel sheet.

The script expects a spreadsheet containing at least the columns:
- email: destination address
- tratamento: pronoun or greeting (ex: "Sr.")
- nome: recipient name

Extra columns are available as template variables when rendering the
message subject and body.

Authentication is performed through the Gmail API using OAuth2. The
first run will prompt for authentication in the browser and persist
the resulting token in ``token.json``.
"""
from __future__ import annotations

import argparse
import base64
import logging
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from pathlib import Path
from typing import Dict, Iterable, List

import pandas as pd
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from jinja2 import Template

SCOPES = ["https://www.googleapis.com/auth/gmail.send"]

logger = logging.getLogger(__name__)


def authenticate(credentials_path: Path, token_path: Path) -> Credentials:
    """Authenticate against Gmail, creating or refreshing a token."""
    creds = None
    if token_path.exists():
        creds = Credentials.from_authorized_user_file(str(token_path), SCOPES)

    if creds and creds.valid:
        return creds

    if creds and creds.expired and creds.refresh_token:
        logger.info("Refreshing Gmail access token...")
        creds.refresh(Request())
    else:
        logger.info("Starting OAuth flow. Follow the browser instructions.")
        flow = InstalledAppFlow.from_client_secrets_file(str(credentials_path), SCOPES)
        creds = flow.run_local_server(port=0)

    token_path.write_text(creds.to_json())
    logger.info("Saved refreshed token to %s", token_path)
    return creds


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


def create_message(sender: str, to: str, subject: str, body_html: str) -> Dict[str, str]:
    """Compose a MIME email ready for Gmail API."""
    message = MIMEMultipart("alternative")
    message["To"] = to
    message["From"] = sender
    message["Subject"] = subject

    part_html = MIMEText(body_html, "html", "utf-8")
    message.attach(part_html)

    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode("utf-8")
    return {"raw": raw_message}


def send_messages(
    service,
    sender: str,
    contacts: Iterable[Dict[str, str]],
    subject_template: str,
    body_template: str,
    dry_run: bool = False,
) -> List[str]:
    """Send rendered messages to all contacts."""
    sent_ids: List[str] = []
    for row in contacts:
        context = {k: str(v) if pd.notna(v) else "" for k, v in row.items()}
        subject = render_template(subject_template, context)
        body = render_template(body_template, context)
        message = create_message(sender=sender, to=context["email"], subject=subject, body_html=body)

        logger.info("Prepared email to %s with subject '%s'", context["email"], subject)
        if dry_run:
            logger.debug("Dry run enabled, skipping send for %s", context["email"])
            continue

        sent = (
            service.users()
            .messages()
            .send(userId="me", body=message)
            .execute()
        )
        sent_ids.append(sent.get("id", ""))
        logger.info("Sent message to %s (id=%s)", context["email"], sent.get("id"))
    return sent_ids


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Send Gmail messages from an Excel contact list.")
    parser.add_argument("excel", type=Path, help="Path to the Excel file containing contacts.")
    parser.add_argument("--sheet", help="Excel sheet name to read.")
    parser.add_argument("--credentials", type=Path, default=Path("credentials.json"), help="Path to the OAuth client credentials file.")
    parser.add_argument("--token", type=Path, default=Path("token.json"), help="Path to the saved OAuth token file.")
    parser.add_argument("--sender", required=True, help="Email address of the sender (must match Gmail account).")
    parser.add_argument("--subject-template", required=True, help="Template for the email subject. Jinja2 placeholders are allowed.")
    parser.add_argument("--body-template", required=True, help="Template for the HTML body. Jinja2 placeholders are allowed.")
    parser.add_argument("--dry-run", action="store_true", help="Render messages without sending them.")
    parser.add_argument("--log-level", default="INFO", help="Logging level (DEBUG, INFO, WARNING, ERROR).")
    return parser.parse_args()


def main() -> None:
    args = parse_args()
    logging.basicConfig(level=args.log_level.upper(), format="%(levelname)s: %(message)s")

    if not args.credentials.exists():
        raise FileNotFoundError(
            f"Credentials file '{args.credentials}' not found. Download it from the Google Cloud console."
        )

    creds = authenticate(args.credentials, args.token)
    service = build("gmail", "v1", credentials=creds)

    contacts_df = load_contacts(args.excel, args.sheet)
    contacts_records = contacts_df.to_dict(orient="records")

    send_messages(
        service=service,
        sender=args.sender,
        contacts=contacts_records,
        subject_template=args.subject_template,
        body_template=args.body_template,
        dry_run=args.dry_run,
    )


if __name__ == "__main__":
    main()
