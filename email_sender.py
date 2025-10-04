"""Email sender CLI and programmatic API for Emaileria."""

from __future__ import annotations

import argparse
import csv
import datetime as _dt
import getpass
import logging
import os
import sqlite3
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Sequence

from emaileria import config
from emaileria.datasource.excel import load_contacts
from emaileria.providers.smtp import SMTPProvider
from emaileria.sender import ResultadoEnvio, send_messages

RATE_LIMIT_PER_MINUTE = int(os.getenv("RATE_LIMIT_PER_MINUTE", "80"))
"""Limite padrão de envios por minuto usado pelo token bucket."""

_BASE_DIR = Path(__file__).resolve().parent
_LOG_DIR = _BASE_DIR / "logs"
_CSV_LOG_PATH = _LOG_DIR / "envios.csv"
_SQLITE_LOG_PATH = _LOG_DIR / "emaileria.db"
_CSV_HEADERS = ["timestamp", "email", "assunto", "status", "tentativas", "erro"]


@dataclass
class RunParams:
    """Parâmetros para execução do envio de emails."""

    input_path: Path
    sender: str
    subject_template: str
    body_html: str
    sheet: str | None = None
    smtp_user: str | None = None
    smtp_password: str | None = None
    dry_run: bool = False
    limit: int | None = None
    offset: int | None = None
    log_level: str | None = None


def _ensure_logs_dir() -> None:
    _LOG_DIR.mkdir(parents=True, exist_ok=True)


def _append_to_csv(rows: Iterable[dict[str, object]]) -> None:
    _ensure_logs_dir()
    file_exists = _CSV_LOG_PATH.exists()
    with _CSV_LOG_PATH.open("a", encoding="utf-8", newline="") as csv_file:
        writer = csv.DictWriter(csv_file, fieldnames=_CSV_HEADERS)
        if not file_exists:
            writer.writeheader()
        writer.writerows(rows)


def _append_to_sqlite(rows: Iterable[dict[str, object]]) -> None:
    if not rows:
        return

    _ensure_logs_dir()

    try:
        connection = sqlite3.connect(_SQLITE_LOG_PATH)
    except sqlite3.Error as exc:  # pragma: no cover - file permission issues
        logging.warning("Não foi possível abrir o banco de dados de logs: %s", exc)
        return

    with connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS envios (
                timestamp TEXT NOT NULL,
                email TEXT NOT NULL,
                assunto TEXT,
                status TEXT NOT NULL,
                tentativas INTEGER NOT NULL,
                erro TEXT
            )
            """
        )
        connection.executemany(
            """
            INSERT INTO envios (timestamp, email, assunto, status, tentativas, erro)
            VALUES (?, ?, ?, ?, ?, ?)
            """,
            [
                (
                    row["timestamp"],
                    row["email"],
                    row.get("assunto"),
                    row["status"],
                    int(row["tentativas"]),
                    row.get("erro"),
                )
                for row in rows
            ],
        )


def _persist_results(results: Iterable[ResultadoEnvio]) -> None:
    rows: list[dict[str, object]] = []
    for result in results:
        timestamp = _dt.datetime.now(_dt.timezone.utc).isoformat()
        rows.append(
            {
                "timestamp": timestamp,
                "email": result.destinatario,
                "assunto": result.assunto or "",
                "status": "sent" if result.sucesso else "failed",
                "tentativas": result.tentativas,
                "erro": result.erro or "",
            }
        )

    if not rows:
        return

    _append_to_csv(rows)
    _append_to_sqlite(rows)


def _configure_logging(log_level: str | None) -> None:
    if log_level is None:
        return

    root_logger = logging.getLogger()
    level_name = str(log_level).upper()

    if root_logger.handlers:
        root_logger.setLevel(level_name)
        return

    logging.basicConfig(level=level_name, format="%(levelname)s: %(message)s")


def _summarize_results(results: List[ResultadoEnvio]) -> None:
    total_results = len(results)
    successful = sum(1 for result in results if result.sucesso)
    failed = total_results - successful
    logging.debug(
        "Resumo do dry-run: total=%s sucesso=%s falha=%s",
        total_results,
        successful,
        failed,
    )


def run_program(params: RunParams) -> int:
    """Executa o envio de emails a partir dos parâmetros informados."""

    os.environ.setdefault("RATE_LIMIT_PER_MINUTE", str(RATE_LIMIT_PER_MINUTE))
    _configure_logging(params.log_level or "INFO")

    input_path = Path(params.input_path)

    offset = params.offset or 0
    if offset < 0:
        logging.error("offset deve ser um inteiro maior ou igual a zero.")
        return 1

    limit = params.limit
    if limit is not None and limit <= 0:
        logging.error("limit deve ser um inteiro positivo quando informado.")
        return 1

    try:
        contacts_df = load_contacts(input_path, params.sheet)
    except ValueError as exc:
        logging.error(str(exc))
        return 1

    total_contacts = len(contacts_df)
    filtered_df = contacts_df.iloc[offset:]
    if limit is not None:
        filtered_df = filtered_df.iloc[:limit]

    sampled_records: list[dict[str, object]] = []
    for position, (_, row) in enumerate(filtered_df.iterrows()):
        record = row.to_dict()
        record["__row_position__"] = offset + position + 1
        sampled_records.append(record)

    processed_count = len(sampled_records)

    smtp_user = params.smtp_user or params.sender

    if params.dry_run:
        try:
            results = send_messages(
                sender=params.sender,
                contacts=sampled_records,
                subject_template=params.subject_template,
                body_template=params.body_html,
                dry_run=True,
            )
        except SystemExit as exc:
            return int(exc.code or 1)

        _summarize_results(results)
        logging.info(
            "Pré-visualizados %s de %s registros (offset=%s, limit=%s)",
            processed_count,
            total_contacts,
            offset,
            limit if limit is not None else "None",
        )
        return 0

    if params.smtp_password:
        logging.warning(
            "Por segurança, evite informar --smtp-password diretamente. "
            "Considere usar o prompt interativo ou a variável de ambiente SMTP_PASSWORD."
        )

    smtp_password = params.smtp_password or os.getenv("SMTP_PASSWORD")
    if smtp_password is None:
        smtp_password = getpass.getpass(
            prompt="SMTP password (app password recommended): "
        )

    try:
        provider: Optional[SMTPProvider]
        with SMTPProvider(
            config.SMTP_HOST,
            config.SMTP_PORT,
            smtp_user,
            smtp_password,
            timeout=config.SMTP_TIMEOUT,
        ) as provider:
            try:
                results = send_messages(
                    sender=params.sender,
                    contacts=sampled_records,
                    subject_template=params.subject_template,
                    body_template=params.body_html,
                    provider=provider,
                    dry_run=False,
                )
            except SystemExit as exc:
                return int(exc.code or 1)
    except Exception as exc:  # pragma: no cover - network/authentication issues
        logging.error("Falha ao estabelecer conexão SMTP: %s", exc)
        return 1

    _persist_results(results)
    return 0


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Send Gmail messages from an Excel contact list."
    )
    parser.add_argument(
        "excel",
        type=Path,
        help="Path to the Excel/CSV file containing contacts.",
    )
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
        help="Template for the email subject. Jinja2 placeholders are allowed.",
    )
    parser.add_argument(
        "--subject-template-file",
        type=Path,
        help="Path to a file containing the subject template. Overrides --subject-template when provided.",
    )
    parser.add_argument(
        "--body-template",
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


def _resolve_template(
    value: str | None,
    file_path: Path | None,
    *,
    inline_arg: str,
    file_arg: str,
) -> str:
    if file_path is not None:
        if not file_path.exists():
            raise FileNotFoundError(
                f"Arquivo especificado em --{file_arg} não encontrado: {file_path}"
            )
        return file_path.read_text(encoding="utf-8")

    if value is None:
        raise ValueError(
            f"Informe --{inline_arg} ou --{file_arg}."
        )

    return value


def _should_delegate(argv: Sequence[str]) -> bool:
    if not argv:
        return False
    if argv[0] == "preview":
        return True
    if any(arg.startswith("--report-") for arg in argv):
        return True
    return False


def main(argv: Optional[Sequence[str]] = None) -> None:
    if argv is None:
        import sys

        argv = list(sys.argv[1:])
    else:
        argv = list(argv)

    if _should_delegate(argv):
        from emaileria.cli import main as _cli_main

        _cli_main(list(argv))
        return

    parser = _build_parser()
    args = parser.parse_args(list(argv))

    try:
        subject_template = _resolve_template(
            args.subject_template,
            args.subject_template_file,
            inline_arg="subject-template",
            file_arg="subject-template-file",
        )
        body_template = _resolve_template(
            args.body_template,
            args.body_template_file,
            inline_arg="body-template",
            file_arg="body-template-file",
        )
    except FileNotFoundError as exc:
        logging.basicConfig(level="ERROR", format="%(levelname)s: %(message)s")
        logging.error(str(exc))
        raise SystemExit(1) from exc
    except ValueError as exc:
        logging.basicConfig(level="ERROR", format="%(levelname)s: %(message)s")
        logging.error(str(exc))
        raise SystemExit(1) from exc

    params = RunParams(
        input_path=args.excel,
        sender=args.sender,
        subject_template=subject_template,
        body_html=body_template,
        sheet=args.sheet,
        smtp_user=args.smtp_user,
        smtp_password=args.smtp_password,
        dry_run=args.dry_run,
        limit=args.limit,
        offset=args.offset,
        log_level=args.log_level,
    )

    exit_code = run_program(params)
    if exit_code != 0:
        raise SystemExit(exit_code)


__all__ = ["RunParams", "run_program", "main", "RATE_LIMIT_PER_MINUTE"]


if __name__ == "__main__":  # pragma: no cover - manual execution entry point
    import sys

    main(sys.argv[1:])
