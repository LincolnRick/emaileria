"""Email sender CLI and programmatic API for Emaileria."""

from __future__ import annotations

import argparse
import csv
import datetime as _dt
import logging
import os
import sqlite3
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Iterable, List, Optional, Sequence

import pandas as pd

from emaileria import config
from emaileria.providers.smtp import SMTPProvider
from emaileria.sender import ResultadoEnvio, send_messages
from emaileria.templating import TemplateRenderingError

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

    input_path: str
    sheet: str | None
    sender: str
    smtp_user: str
    smtp_password: str
    subject_template: str
    body_html: str
    dry_run: bool = True
    limit: int | None = None
    offset: int | None = None
    log_level: str | None = "INFO"
    cc: list[str] | None = None
    bcc: list[str] | None = None
    reply_to: str | None = None
    interval_seconds: float = 0.75


_CANCEL_EVENT = threading.Event()
"""Flag global utilizada para solicitar cancelamento seguro dos envios."""


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


def load_contacts(path: str | Path, sheet: str | None) -> pd.DataFrame:
    """Carrega a planilha de contatos respeitando a aba informada."""

    file_path = Path(path)
    if not file_path.exists():
        raise ValueError(f"Arquivo não encontrado: {file_path}")

    suffix = file_path.suffix.lower()
    try:
        if suffix == ".csv":
            df = pd.read_csv(file_path)
        elif suffix in {".xls", ".xlsx"}:
            sheet_name: str | int | None
            if sheet:
                sheet_name = sheet
            else:
                sheet_name = 0
            df = pd.read_excel(file_path, sheet_name=sheet_name)
        else:
            raise ValueError(
                "Formato de arquivo não suportado. Utilize CSV ou XLSX."
            )
    except Exception as exc:  # pylint: disable=broad-except
        raise ValueError(f"Erro ao carregar planilha: {exc}") from exc

    if df is None:
        raise ValueError("Planilha vazia ou não carregada corretamente.")

    return df


def _normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    copy = df.copy()
    copy.columns = [str(column).strip().lower() for column in copy.columns]
    return copy


def _iter_contacts(df: pd.DataFrame, offset: int, limit: int | None) -> list[dict[str, object]]:
    filtered = df.iloc[offset:]
    if limit is not None:
        filtered = filtered.iloc[:limit]

    records: list[dict[str, object]] = []
    for position, (_, row) in enumerate(filtered.iterrows()):
        data = row.to_dict()
        data["__row_position__"] = offset + position + 1
        records.append(data)
    return records


def _mask_password(value: str | None) -> str:
    if not value:
        return "(não informado)"
    return "*" * 8


def request_cancel() -> None:
    """Sinaliza para que o processamento seja cancelado."""

    _CANCEL_EVENT.set()


def reset_cancel_flag() -> None:
    """Limpa a sinalização de cancelamento."""

    _CANCEL_EVENT.clear()


def _log_configuration(params: RunParams, total: int) -> None:
    masked_password = _mask_password(params.smtp_password or os.getenv("SMTP_PASSWORD"))
    logging.info("Planilha: %s", params.input_path)
    logging.info("Sheet: %s", params.sheet or "(padrão)")
    logging.info("Remetente: %s", params.sender or "(não informado)")
    logging.info("SMTP User: %s", params.smtp_user or params.sender or "(não informado)")
    logging.info("CC: %s", ", ".join(params.cc or []) or "(nenhum)")
    logging.info("BCC: %s", ", ".join(params.bcc or []) or "(nenhum)")
    logging.info("Reply-To: %s", params.reply_to or "(nenhum)")
    logging.info("Modo: %s", "Dry-run" if params.dry_run else "Envio real")
    logging.info("Intervalo entre envios: %.2fs", max(params.interval_seconds, 0.0))
    logging.info("Senha SMTP: %s", masked_password)
    logging.info("Total de contatos selecionados: %s", total)


def _render_only(params: RunParams, records: list[dict[str, object]]) -> int:
    try:
        results = send_messages(
            sender=params.sender,
            contacts=records,
            subject_template=params.subject_template,
            body_template=params.body_html,
            dry_run=True,
            allow_missing_fields=False,
            interval_seconds=0.0,
            cc=params.cc,
            bcc=params.bcc,
            reply_to=params.reply_to,
            cancel_event=_CANCEL_EVENT,
        )
    except TemplateRenderingError as exc:
        logging.error(str(exc))
        return 1
    except SystemExit as exc:  # pragma: no cover - compatibilidade antiga
        return int(exc.code or 1)

    _summarize_results(results)
    return 0


def _send_real(params: RunParams, records: list[dict[str, object]]) -> int:
    smtp_password = params.smtp_password.strip()
    if not smtp_password:
        env_password = os.getenv("SMTP_PASSWORD", "")
        smtp_password = env_password.strip()

    if not smtp_password:
        logging.error(
            "Senha SMTP não informada. Configure SMTP_PASSWORD ou forneça no parâmetro."
        )
        return 1

    smtp_user = params.smtp_user.strip() or params.sender

    try:
        with SMTPProvider(
            config.SMTP_HOST,
            config.SMTP_PORT,
            smtp_user,
            smtp_password,
            timeout=config.SMTP_TIMEOUT,
        ) as provider:
            results = send_messages(
                sender=params.sender,
                contacts=records,
                subject_template=params.subject_template,
                body_template=params.body_html,
                provider=provider,
                dry_run=False,
                allow_missing_fields=False,
                interval_seconds=max(params.interval_seconds, 0.0),
                cc=params.cc,
                bcc=params.bcc,
                reply_to=params.reply_to,
                cancel_event=_CANCEL_EVENT,
            )
    except TemplateRenderingError as exc:
        logging.error(str(exc))
        return 1
    except SystemExit as exc:  # pragma: no cover - compatibilidade antiga
        return int(exc.code or 1)
    except Exception as exc:  # pragma: no cover - network/authentication issues
        logging.error("Falha ao estabelecer conexão SMTP: %s", exc)
        return 1

    _persist_results(results)
    return 0


def run_program(params: RunParams) -> int:
    """Executa o envio de emails a partir dos parâmetros informados."""

    reset_cancel_flag()
    os.environ.setdefault("RATE_LIMIT_PER_MINUTE", str(RATE_LIMIT_PER_MINUTE))
    _configure_logging(params.log_level or "INFO")

    offset = params.offset or 0
    if offset < 0:
        logging.error("offset deve ser um inteiro maior ou igual a zero.")
        return 1

    limit = params.limit
    if limit is not None and limit <= 0:
        logging.error("limit deve ser um inteiro positivo quando informado.")
        return 1

    try:
        contacts_df = load_contacts(params.input_path, params.sheet)
    except ValueError as exc:
        logging.error(str(exc))
        return 1

    contacts_df = _normalize_headers(contacts_df)
    required_columns = {"email", "tratamento", "nome"}
    missing_columns = sorted(required_columns - set(contacts_df.columns))
    if missing_columns:
        logging.error(
            "Planilha inválida. Colunas obrigatórias ausentes: %s",
            ", ".join(missing_columns),
        )
        return 1

    records = _iter_contacts(contacts_df, offset, limit)
    processed_count = len(records)

    _log_configuration(params, processed_count)

    if processed_count == 0:
        logging.warning("Nenhum contato selecionado para processamento.")
        return 0

    logging.info(
        "Processando %s contatos (total na planilha: %s)",
        processed_count,
        len(contacts_df),
    )

    start_time = time.time()
    exit_code = _render_only(params, records) if params.dry_run else _send_real(params, records)

    if _CANCEL_EVENT.is_set():
        logging.warning("Processamento cancelado pelo usuário.")
        return 130

    elapsed = time.time() - start_time
    logging.info("Execução concluída em %.2f segundos", elapsed)
    return exit_code


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
        input_path=str(args.excel),
        sheet=args.sheet,
        sender=args.sender,
        smtp_user=args.smtp_user or args.sender,
        smtp_password=args.smtp_password or "",
        subject_template=subject_template,
        body_html=body_template,
        dry_run=bool(args.dry_run),
        limit=args.limit,
        offset=args.offset,
        log_level=args.log_level,
    )

    exit_code = run_program(params)
    if exit_code != 0:
        raise SystemExit(exit_code)


__all__ = [
    "RunParams",
    "run_program",
    "main",
    "RATE_LIMIT_PER_MINUTE",
    "load_contacts",
    "request_cancel",
    "reset_cancel_flag",
]


if __name__ == "__main__":  # pragma: no cover - manual execution entry point
    import sys

    main(sys.argv[1:])
