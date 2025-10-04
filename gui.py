import logging
import os
import queue
import re
import threading
from pathlib import Path
from typing import Optional

import pandas as pd
import PySimpleGUI as sg

try:
    from email_sender import RunParams, run_program
except ImportError as exc:  # pragma: no cover - integração com UI
    IMPORT_ERROR: Optional[Exception] = exc
    RunParams = None  # type: ignore[assignment]
    run_program = None  # type: ignore[assignment]
else:
    IMPORT_ERROR = None

from emaileria.datasource.excel import load_contacts
from emaileria.templating import TemplateRenderingError, render


REQUIRED_COLUMNS = {"email", "tratamento", "nome"}
PLACEHOLDER_PATTERN = re.compile(r"{{\s*([a-zA-Z0-9_]+)\s*}}")


def _read_html_template(path: str) -> str | None:
    if not path:
        sg.popup_error("Selecione o arquivo de template HTML.")
        return None
    try:
        return Path(path).read_text(encoding="utf-8")
    except FileNotFoundError:
        sg.popup_error("Arquivo de template HTML não encontrado.")
    except OSError as exc:
        sg.popup_error(f"Erro ao ler o template HTML: {exc}")
    return None


def _normalize_headers(df: pd.DataFrame) -> pd.DataFrame:
    copy = df.copy()
    copy.columns = [str(column).strip().lower() for column in copy.columns]
    return copy


def _extract_placeholders(*templates: str) -> set[str]:
    placeholders: set[str] = set()
    for template in templates:
        if template:
            placeholders.update(PLACEHOLDER_PATTERN.findall(template))
    return placeholders


def _update_sheet_combo(window: sg.Window, file_path: str) -> None:
    sheet_element = window["-SHEET-"]
    sheet_element.update(values=[], value="", disabled=True)
    if not file_path:
        return

    suffix = Path(file_path).suffix.lower()
    if suffix != ".xlsx":
        return

    try:
        with pd.ExcelFile(file_path) as workbook:
            sheet_names = workbook.sheet_names
    except Exception as exc:  # pylint: disable=broad-except
        sg.popup_error(f"Não foi possível ler as abas do arquivo: {exc}")
        return

    if not sheet_names:
        sg.popup_error("Nenhuma aba encontrada no arquivo Excel.")
        return

    sheet_element.update(values=sheet_names, value=sheet_names[0], disabled=False)


class QueueLogHandler(logging.Handler):
    """Handler de logging que envia mensagens para a fila da interface."""

    def __init__(self, message_queue: "queue.Queue[tuple[str, str]]") -> None:
        super().__init__()
        self._queue = message_queue
        self.setFormatter(logging.Formatter("%(levelname)s: %(message)s"))

    def emit(self, record: logging.LogRecord) -> None:  # pragma: no cover - integração com UI
        try:
            message = self.format(record)
        except Exception:  # pylint: disable=broad-except
            message = record.getMessage()
        self._queue.put(("LOG", message + "\n"))


def append_log(window: sg.Window, text: str, *, tag: str = "OUT") -> None:
    text_color = "red" if tag == "ERR" else None
    window["-LOG-"].print(text, end="", text_color=text_color)


# -------- UI --------
sg.theme("SystemDefault")

layout = [
    [
        sg.Text("Planilha (XLSX/CSV)"),
        sg.Input(key="-EXCEL-", enable_events=True),
        sg.FileBrowse(file_types=(("Excel/CSV", "*.xlsx;*.xls;*.csv"),)),
    ],
    [
        sg.Text("Aba (sheet)"),
        sg.Combo(
            values=[],
            key="-SHEET-",
            size=(25, 1),
            readonly=True,
            disabled=True,
        ),
    ],
    [sg.HorizontalSeparator()],
    [sg.Text("Remetente (From)"), sg.Input(key="-SENDER-", size=(40, 1))],
    [sg.Text("SMTP User"), sg.Input(key="-SMTPUSER-", size=(40, 1))],
    [sg.Text("SMTP Password"), sg.Input(key="-SMTPPASS-", password_char="*", size=(40, 1))],
    [sg.HorizontalSeparator()],
    [sg.Text("Assunto (Jinja2)"), sg.Input(key="-SUBJECT-", size=(60, 1))],
    [
        sg.Text("Template HTML"),
        sg.Input(key="-HTML-"),
        sg.FileBrowse(file_types=(("HTML", "*.html;*.htm;*.j2"),)),
    ],
    [
        sg.Checkbox(
            "Dry-run (não enviar, apenas pré-visualizar)",
            key="-DRYRUN-",
            default=True,
        )
    ],
    [
        sg.Text("Log level"),
        sg.Combo(
            values=["INFO", "DEBUG", "WARNING", "ERROR"],
            default_value="INFO",
            key="-LOGLEVEL-",
            readonly=True,
            size=(15, 1),
        ),
    ],
    [sg.HorizontalSeparator()],
    [
        sg.Button("Validar & Prévia", key="-PREVIEW-"),
        sg.Button("Enviar", key="-RUN-", bind_return_key=True),
        sg.Button("Sair", key="-EXIT-"),
    ],
    [
        sg.Multiline(
            key="-LOG-",
            size=(100, 20),
            autoscroll=True,
            write_only=True,
            font=("Consolas", 10),
        )
    ],
]

window = sg.Window("Emaileria — Envio de E-mails", layout, finalize=True)

log_queue: "queue.Queue[tuple[str, str]]" = queue.Queue()
queue_handler = QueueLogHandler(log_queue)
root_logger = logging.getLogger()
if not any(isinstance(handler, QueueLogHandler) for handler in root_logger.handlers):
    root_logger.addHandler(queue_handler)
root_logger.setLevel(logging.INFO)

worker_thread: Optional[threading.Thread] = None

if IMPORT_ERROR is not None:
    append_log(
        window,
        "[ERR] Não foi possível importar email_sender. Execute este programa a partir da raiz do projeto.\n",
        tag="ERR",
    )

while True:
    event, values = window.read(timeout=100)
    if event in (sg.WIN_CLOSED, "-EXIT-"):
        break

    if event == "-EXCEL-":
        excel_path = values["-EXCEL-"].strip()
        _update_sheet_combo(window, excel_path)

    if event == "-PREVIEW-":
        excel = values["-EXCEL-"].strip()
        if not excel:
            sg.popup_error("Selecione a planilha (XLSX/CSV).")
            continue

        subject_template = values["-SUBJECT-"].strip()
        if not subject_template:
            sg.popup_error("Informe o assunto (template).")
            continue

        body_html_path = values["-HTML-"].strip()
        body_html = _read_html_template(body_html_path)
        if body_html is None:
            continue

        sheet_value = (values.get("-SHEET-") or "").strip() or None

        try:
            dataframe = load_contacts(Path(excel), sheet=sheet_value)
        except Exception as exc:  # pylint: disable=broad-except
            sg.popup_error(f"Erro ao carregar planilha: {exc}")
            continue

        dataframe = _normalize_headers(dataframe)

        missing_required = sorted(REQUIRED_COLUMNS - set(dataframe.columns))
        if missing_required:
            formatted = ", ".join(missing_required)
            sg.popup_error(
                "Planilha inválida. Colunas obrigatórias ausentes: " f"{formatted}."
            )
            continue

        placeholders = _extract_placeholders(subject_template, body_html)
        missing_placeholders = sorted(placeholders - set(dataframe.columns))
        if missing_placeholders:
            formatted = ", ".join(missing_placeholders)
            sg.popup_error(
                "Planilha inválida. Colunas ausentes para placeholders: "
                f"{formatted}."
            )
            continue

        if dataframe.empty:
            sg.popup(
                "Planilha carregada, mas não há registros para pré-visualizar.",
                title="Prévia",
            )
            continue

        previews: list[str] = []
        success = True
        records = dataframe.to_dict(orient="records")
        for index, row in enumerate(records[:3], start=1):
            try:
                rendered_subject, rendered_body = render(
                    subject_template, body_html, row
                )
            except TemplateRenderingError as exc:
                sg.popup_error(str(exc))
                success = False
                break

            plain_body = re.sub(r"<[^>]+>", "", rendered_body)
            plain_body = re.sub(r"\s+", " ", plain_body).strip()
            plain_body = plain_body[:200]

            previews.append(
                "Amostra {index}\nAssunto: {subject}\nCorpo: {body}".format(
                    index=index, subject=rendered_subject, body=plain_body
                )
            )

        if not success:
            continue

        preview_text = "\n\n".join(previews)
        sg.popup_scrolled(preview_text, title="Prévia", non_blocking=False)

    if event == "-RUN-":
        if IMPORT_ERROR is not None:
            sg.popup_error(
                "Não foi possível importar o módulo email_sender. "
                "Execute este programa a partir da raiz do projeto.\n"
                f"Detalhes: {IMPORT_ERROR}"
            )
            continue

        if worker_thread is not None and worker_thread.is_alive():
            sg.popup_error("Já existe um envio em andamento. Aguarde a finalização.")
            continue

        excel = values["-EXCEL-"].strip()
        if not excel:
            sg.popup_error("Selecione a planilha (XLSX/CSV).")
            continue

        sender = values["-SENDER-"].strip()
        if not sender:
            sg.popup_error("Informe o remetente (From).")
            continue

        subject_template = values["-SUBJECT-"].strip()
        if not subject_template:
            sg.popup_error("Informe o assunto (template).")
            continue

        body_html_path = values["-HTML-"].strip()
        body_html = _read_html_template(body_html_path)
        if body_html is None:
            continue

        password_value = values["-SMTPPASS-"].strip()
        env_password = os.getenv("SMTP_PASSWORD")

        params = RunParams(
            input_path=Path(excel),
            sender=sender,
            subject_template=subject_template,
            body_html=body_html,
            sheet=(values.get("-SHEET-") or "").strip() or None,
            smtp_user=values["-SMTPUSER-"].strip() or None,
            smtp_password=password_value or env_password,
            dry_run=values["-DRYRUN-"],
            log_level=values["-LOGLEVEL-"] or None,
        )

        append_log(window, "\n[INFO] Iniciando envio\n")
        append_log(
            window,
            "[INFO] Planilha: {} | Sheet: {} | Remetente: {} | Assunto: {}\n".format(
                excel,
                (values.get("-SHEET-") or "").strip() or "(padrão)",
                sender,
                subject_template,
            ),
        )

        def _run() -> None:
            try:
                result_code = run_program(params)  # type: ignore[misc]
            except Exception as exc:  # pylint: disable=broad-except
                log_queue.put(("ERROR", f"Falha durante a execução: {exc}\n"))
            else:
                log_queue.put(("RESULT", str(result_code)))

        worker_thread = threading.Thread(target=_run, daemon=True)
        worker_thread.start()

    try:
        while True:
            tag, payload = log_queue.get_nowait()
            if tag == "LOG":
                append_log(window, payload)
            elif tag == "ERROR":
                append_log(window, payload, tag="ERR")
                worker_thread = None
            elif tag == "RESULT":
                append_log(window, f"\n[INFO] Finalizado com código {payload}\n")
                worker_thread = None
    except queue.Empty:
        pass

window.close()
