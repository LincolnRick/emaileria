"""Command line interface for the Emaileria package."""

from __future__ import annotations

import argparse
import datetime as _dt
import getpass
import html
import logging
from pathlib import Path

from . import config
from .datasource.excel import load_contacts
from .providers.smtp import SMTPProvider
from .sender import _prepare_context, send_messages
from .templating import TemplateRenderingError, render


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
        "--offset",
        type=int,
        default=0,
        help="Número de linhas iniciais a ignorar antes do envio.",
    )
    parser.add_argument(
        "--limit",
        type=int,
        help="Quantidade máxima de contatos a serem processados após o offset.",
    )
    parser.add_argument(
        "--log-level", default="INFO", help="Logging level (DEBUG, INFO, WARNING, ERROR)."
    )
    return parser


def _build_preview_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Generate an HTML gallery preview of rendered emails."
    )
    parser.add_argument(
        "excel",
        type=Path,
        help="Path to the Excel/CSV file containing contacts.",
    )
    parser.add_argument("--sheet", help="Excel sheet name to read.")
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
    parser.add_argument(
        "--limit",
        type=int,
        default=5,
        help="Maximum number of contacts to render in the preview gallery.",
    )
    parser.add_argument(
        "--log-level",
        default="INFO",
        help="Logging level (DEBUG, INFO, WARNING, ERROR).",
    )
    return parser


def _load_contacts(path: Path, sheet: str | None):
    return load_contacts(path, sheet)


def _render_preview(
    *,
    contacts: list[dict[str, object]],
    subject_template: str,
    body_template: str,
    limit: int,
) -> list[tuple[int, str, str, str]]:
    if limit <= 0:
        raise ValueError("limit must be a positive integer")

    preview_entries: list[tuple[int, str, str, str]] = []
    for index, row in enumerate(contacts, start=1):
        if len(preview_entries) >= limit:
            break
        context = _prepare_context(row)
        try:
            subject, body = render(subject_template, body_template, context)
        except TemplateRenderingError as exc:
            raise SystemExit(
                "Falha ao renderizar "
                f"{exc.template_type} na linha {index}: placeholder '{exc.placeholder}' não encontrado."
            ) from exc
        preview_entries.append((index, context["email"], subject, body))
    return preview_entries


def _build_preview_html(
    *,
    generated_at: _dt.datetime,
    entries: list[tuple[int, str, str, str]],
) -> str:
    formatted_timestamp = generated_at.strftime("%Y-%m-%d %H:%M:%S")
    navigation_links = "\n        ".join(
        f"<a href=\"#message-{pos}\">#{pos} - {html.escape(subject or '(sem assunto)')}</a>"
        for pos, _, subject, _ in entries
    )
    navigation_markup = (
        navigation_links if navigation_links else "<span class=\"empty\">Nenhuma mensagem gerada.</span>"
    )

    gallery_items = []
    for pos, recipient, subject, body in entries:
        gallery_items.append(
            """
            <article class="message-card" id="message-{pos}">
                <header>
                    <h2>{subject}</h2>
                    <p class="meta">Destinatário: {recipient}</p>
                </header>
                <section class="message-body">{body}</section>
            </article>
            """.format(
                pos=pos,
                subject=html.escape(subject or "(sem assunto)"),
                recipient=html.escape(recipient),
                body=body,
            )
        )

    gallery_markup = "\n".join(gallery_items) if gallery_items else "<p>Não há mensagens para pré-visualizar.</p>"

    return f"""<!DOCTYPE html>
<html lang=\"pt-BR\">
<head>
    <meta charset=\"utf-8\" />
    <title>Pré-visualização de emails</title>
    <style>
        body {{
            font-family: Arial, Helvetica, sans-serif;
            margin: 2rem;
            background-color: #f7f7f9;
            color: #1f1f1f;
        }}
        h1 {{
            font-size: 1.8rem;
            margin-bottom: 0.5rem;
        }}
        .timestamp {{
            color: #555;
            margin-bottom: 1.5rem;
        }}
        nav {{
            display: flex;
            flex-wrap: wrap;
            gap: 0.75rem;
            margin-bottom: 2rem;
        }}
        nav a {{
            text-decoration: none;
            color: #1a73e8;
            background: #e7f0fe;
            padding: 0.35rem 0.75rem;
            border-radius: 999px;
            font-size: 0.9rem;
        }}
        nav a:hover {{
            background: #d2e3fc;
        }}
        nav .empty {{
            color: #6b6b6b;
            font-style: italic;
        }}
        .gallery {{
            display: grid;
            gap: 1.5rem;
            grid-template-columns: repeat(auto-fit, minmax(280px, 1fr));
        }}
        .message-card {{
            background: #ffffff;
            border-radius: 12px;
            box-shadow: 0 1px 3px rgba(0, 0, 0, 0.1);
            padding: 1.25rem;
            display: flex;
            flex-direction: column;
            gap: 0.75rem;
        }}
        .message-card h2 {{
            font-size: 1.2rem;
            margin: 0;
        }}
        .meta {{
            color: #555;
            font-size: 0.9rem;
        }}
        .message-body {{
            border-top: 1px solid #e0e0e0;
            padding-top: 0.75rem;
        }}
        .message-body * {{
            max-width: 100%;
        }}
    </style>
</head>
<body>
    <h1>Pré-visualização de emails</h1>
    <p class=\"timestamp\">Gerado em {formatted_timestamp}</p>
    <nav>{navigation_markup}</nav>
    <section class=\"gallery\">{gallery_markup}</section>
</body>
</html>
"""


def _write_preview(entries: list[tuple[int, str, str, str]]) -> Path:
    generated_at = _dt.datetime.now()
    timestamp = generated_at.strftime("%Y%m%d-%H%M%S")
    preview_dir = Path("previews") / timestamp
    preview_dir.mkdir(parents=True, exist_ok=True)

    html_content = _build_preview_html(
        generated_at=generated_at,
        entries=entries,
    )
    output_file = preview_dir / "index.html"
    output_file.write_text(html_content, encoding="utf-8")
    return output_file


def main(argv: list[str] | None = None) -> None:
    if argv and len(argv) > 0 and argv[0] == "preview":
        preview_parser = _build_preview_parser()
        preview_args = preview_parser.parse_args(argv[1:])

        logging.basicConfig(
            level=preview_args.log_level.upper(),
            format="%(levelname)s: %(message)s",
        )

        subject_template = _read_template(
            preview_args.subject_template, preview_args.subject_template_file
        )
        body_template = _read_template(
            preview_args.body_template, preview_args.body_template_file
        )

        contacts_df = _load_contacts(preview_args.excel, preview_args.sheet)
        entries = _render_preview(
            contacts=contacts_df.to_dict(orient="records"),
            subject_template=subject_template,
            body_template=body_template,
            limit=preview_args.limit,
        )
        output_path = _write_preview(entries)
        logging.info("Pré-visualização gerada em %s", output_path)
        return

    parser = build_parser()
    args = parser.parse_args(argv)

    logging.basicConfig(level=args.log_level.upper(), format="%(levelname)s: %(message)s")

    subject_template = _read_template(args.subject_template, args.subject_template_file)
    body_template = _read_template(args.body_template, args.body_template_file)

    try:
        contacts_df = _load_contacts(args.excel, args.sheet)
    except ValueError as exc:
        logging.error(str(exc))
        raise SystemExit(1) from exc

    total_contacts = len(contacts_df)
    offset = args.offset or 0
    if offset < 0:
        logging.error("offset deve ser um inteiro maior ou igual a zero.")
        raise SystemExit(1)

    limit = args.limit
    if limit is not None and limit <= 0:
        logging.error("limit deve ser um inteiro positivo quando informado.")
        raise SystemExit(1)

    filtered_df = contacts_df.iloc[offset:]
    if limit is not None:
        filtered_df = filtered_df.iloc[:limit]

    sampled_records: list[dict[str, object]] = []
    for position, (_, row) in enumerate(filtered_df.iterrows()):
        record = row.to_dict()
        record["__row_position__"] = offset + position + 1
        sampled_records.append(record)

    processed_count = len(sampled_records)

    smtp_user = args.smtp_user or args.sender

    if args.dry_run:
        send_messages(
            sender=args.sender,
            contacts=sampled_records,
            subject_template=subject_template,
            body_template=body_template,
            dry_run=True,
        )
        logging.info(
            "Pré-visualizados %s de %s registros (offset=%s, limit=%s)",
            processed_count,
            total_contacts,
            offset,
            limit if limit is not None else "None",
        )
        return

    if args.smtp_password:
        logging.warning(
            "Por segurança, evite informar --smtp-password diretamente. "
            "Considere usar o prompt interativo ou a variável de ambiente SMTP_PASSWORD."
        )

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
            contacts=sampled_records,
            subject_template=subject_template,
            body_template=body_template,
            provider=provider,
            dry_run=False,
        )
