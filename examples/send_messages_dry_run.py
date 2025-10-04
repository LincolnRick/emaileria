"""Render example messages using Emaileria without sending them."""

from __future__ import annotations

import csv
import sys
from pathlib import Path

ROOT_DIR = Path(__file__).resolve().parents[1]
if str(ROOT_DIR) not in sys.path:
    sys.path.insert(0, str(ROOT_DIR))

from emaileria.sender import send_messages


def _load_contacts(csv_path: Path) -> list[dict[str, str]]:
    """Load contacts from a CSV file into a list of dictionaries."""
    with csv_path.open("r", encoding="utf-8", newline="") as csv_file:
        reader = csv.DictReader(csv_file)
        return [dict(row) for row in reader]


def main() -> None:
    examples_dir = Path(__file__).resolve().parent
    data_dir = examples_dir / "readme"

    contacts_path = data_dir / "leads.csv"
    subject_template_path = data_dir / "template_assunto.txt"
    body_template_path = data_dir / "template_corpo.html"

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
