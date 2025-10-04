"""Render example messages using Emaileria without sending them."""

from __future__ import annotations

import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from emaileria.datasource.excel import load_contacts as load_contacts_dataframe
from emaileria.sender import send_messages


def _load_contacts(excel_path: Path) -> list[dict[str, str]]:
    """Load contacts from an Excel file into a list of dictionaries."""

    dataframe = load_contacts_dataframe(excel_path)
    contacts = dataframe.to_dict(orient="records")
    for contact in contacts:
        contact.setdefault("data_envio", "")
    return contacts


def main() -> None:
    examples_dir = Path(__file__).resolve().parent
    contacts_path = examples_dir / "leads_exemplo.xlsx"
    subject_template_path = examples_dir / "assunto_exemplo.txt"
    body_template_path = examples_dir / "corpo_exemplo.html"

    contacts = _load_contacts(contacts_path)
    subject_template = subject_template_path.read_text(encoding="utf-8")
    body_template = body_template_path.read_text(encoding="utf-8")

    results = send_messages(
        sender="contato@example.com",
        contacts=contacts,
        subject_template=subject_template,
        body_template=body_template,
        dry_run=True,
    )

    for result in results:
        print(f"{result.destinatario}: {result.assunto}")


if __name__ == "__main__":
    main()
