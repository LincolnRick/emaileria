import json
import logging
import os
import queue
import re
import threading
from pathlib import Path
from typing import Mapping, Optional

import pandas as pd
import PySimpleGUI as sg

try:
    from email_sender import RunParams, load_contacts, run_program
except ImportError as exc:  # pragma: no cover - integração com UI
    IMPORT_ERROR: Optional[Exception] = exc
    RunParams = None  # type: ignore[assignment]
    load_contacts = None  # type: ignore[assignment]
    run_program = None  # type: ignore[assignment]
else:
    IMPORT_ERROR = None

from emaileria.templating import TemplateRenderingError, extract_placeholders, render


REQUIRED_COLUMNS = {"email", "tratamento", "nome"}
SETTINGS_PATH = Path.home() / ".emaileria_gui.json"
PROCESSING_PATTERN = re.compile(
    r"Processando\s+(?P<processed>\d+)\s+contatos.*total[^0-9]*(?P<total>\d+)",
    re.IGNORECASE,
)

GLOBAL_PLACEHOLDERS = {"now", "hoje", "data_envio", "hora_envio"}

current_sheets: list[str] = []


def _load_settings() -> dict[str, str]:
    if not SETTINGS_PATH.exists():
        return {}
    try:
        data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}
    if not isinstance(data, dict):
        return {}
    return {str(key): str(value) for key, value in data.items() if isinstance(value, str)}


def _save_settings(values: Mapping[str, object] | None) -> None:
    if values is None:
        return

    relevant = {
        "excel": str(values.get("-EXCEL-", "") or ""),
        "html": str(values.get("-HTML-", "") or ""),
        "sheet": str(values.get("-SHEET-", "") or ""),
        "sender": str(values.get("-SENDER-", "") or ""),
    }
    try:
        SETTINGS_PATH.write_text(json.dumps(relevant, ensure_ascii=False, indent=2), encoding="utf-8")
    except OSError:
        logging.getLogger(__name__).debug("Não foi possível salvar as preferências da GUI.")


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


INTERACTIVE_KEYS = [
    "-EXCEL-",
    "-EXCEL-BROWSE-",
    "-SHEET-",
    "-SENDER-",
    "-SMTPUSER-",
    "-SMTPPASS-",
    "-SUBJECT-",
    "-SUBJECTFILE-",
    "-SUBJECTFILE-BROWSE-",
    "-HTML-",
    "-HTML-BROWSE-",
    "-HTML-PREVIEW-",
    "-DRYRUN-",
    "-ALLOWMISSING-",
    "-LOGLEVEL-",
    "-PREVIEW-",
    "-RUN-",
    "-EXIT-",
]

progress_state: dict[str, int] = {"total": 0, "sent": 0}
validation_passed = False

_VALIDATION_RESET_EVENTS = {
    "-EXCEL-",
    "-SHEET-",
    "-SENDER-",
    "-SMTPUSER-",
    "-SUBJECT-",
    "-SUBJECTFILE-",
    "-HTML-",
    "-DRYRUN-",
    "-ALLOWMISSING-",
}


def _set_controls_enabled(window: sg.Window, *, enabled: bool) -> None:
    for key in INTERACTIVE_KEYS:
        element = window.AllKeysDict.get(key)
        if element is None:
            continue
        try:
            element.update(disabled=not enabled)
        except TypeError:
            try:
                element.update(state="disabled" if not enabled else "normal")
            except Exception:  # pylint: disable=broad-except
                pass
    if enabled:
        _update_run_button_state(window)


def _update_counter_display(window: sg.Window, *, unknown_total: bool = False) -> None:
    total = progress_state.get("total", 0)
    sent = progress_state.get("sent", 0)
    if total == 0 and unknown_total:
        total_display = "?"
    else:
        total_display = str(total)
    window["-COUNTER-"].update(f"{sent}/{total_display}")


def _update_run_button_state(window: sg.Window) -> None:
    run_button = window.AllKeysDict.get("-RUN-")
    if run_button is None:
        return
    should_enable = validation_passed
    try:
        run_button.update(disabled=not should_enable)
    except TypeError:
        try:
            run_button.update(state="disabled" if not should_enable else "normal")
        except Exception:  # pylint: disable=broad-except
            pass


def _set_validation_state(window: sg.Window, *, passed: bool) -> None:
    global validation_passed  # pylint: disable=global-statement
    validation_passed = passed
    _update_run_button_state(window)


def _set_running_state(window: sg.Window, *, running: bool) -> None:
    progress_bar = window["-PROGRESS-"]
    if running:
        progress_state["total"] = 0
        progress_state["sent"] = 0
        _set_controls_enabled(window, enabled=False)
        window["-STATUS-"].update("Processando...")
        progress_bar.update(current_count=0, visible=True)
        _update_counter_display(window, unknown_total=True)
        try:
            progress_bar.Widget.start(10)
        except Exception:  # pylint: disable=broad-except
            pass
        return

    try:
        progress_bar.Widget.stop()
    except Exception:  # pylint: disable=broad-except
        pass
    progress_bar.update(visible=False)
    window["-STATUS-"].update("Pronto")
    _set_controls_enabled(window, enabled=True)
    if progress_state["total"] == 0:
        window["-COUNTER-"].update("0/0")
    else:
        _update_counter_display(window)


def _apply_saved_settings(window: sg.Window) -> None:
    settings = _load_settings()
    if not settings:
        return

    excel_path = settings.get("excel", "")
    if excel_path:
        window["-EXCEL-"].update(excel_path)
        if Path(excel_path).exists():
            _update_sheet_combo(window, excel_path)

    sheet_value = settings.get("sheet", "")
    if sheet_value and sheet_value in current_sheets:
        window["-SHEET-"].update(value=sheet_value)

    html_path = settings.get("html", "")
    if html_path:
        window["-HTML-"].update(html_path)

    sender_value = settings.get("sender", "")
    if sender_value:
        window["-SENDER-"].update(sender_value)


def _handle_progress_from_log(window: sg.Window, message: str) -> None:
    match = PROCESSING_PATTERN.search(message)
    if match:
        try:
            progress_state["total"] = int(match.group("processed"))
        except (TypeError, ValueError):
            progress_state["total"] = 0
        progress_state["sent"] = 0
        _update_counter_display(window, unknown_total=progress_state["total"] == 0)
        return

    if "Prepared email to" in message:
        progress_state["sent"] = progress_state.get("sent", 0) + 1
        if (
            progress_state.get("total")
            and progress_state["sent"] > progress_state.get("total", 0)
        ):
            progress_state["total"] = progress_state["sent"]
        _update_counter_display(
            window,
            unknown_total=progress_state.get("total", 0) == 0,
        )

def _update_sheet_combo(window: sg.Window, file_path: str) -> None:
    sheet_element = window["-SHEET-"]
    sheet_element.update(values=[], value="", disabled=True)
    global current_sheets
    current_sheets = []
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
    current_sheets = sheet_names


class QueueLogHandler(logging.Handler):
    """Handler de logging que envia mensagens para a fila da interface."""

    def __init__(self, message_queue: "queue.Queue[tuple[str, str]]") -> None:
        super().__init__()
        self._queue = message_queue
        self.setLevel(logging.NOTSET)
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


def _prepare_run_params(
    values: dict[str, object], *, dry_run_override: bool | None = None
) -> Optional[RunParams]:
    if RunParams is None:
        return None

    excel_path = str(values.get("-EXCEL-", "")).strip()
    if not excel_path:
        sg.popup_error("Selecione a planilha (XLSX/CSV).")
        return None
    if not Path(excel_path).exists():
        sg.popup_error("Arquivo de planilha não encontrado. Verifique o caminho informado.")
        return None

    subject_template = str(values.get("-SUBJECT-", "")).strip()
    subject_file_path = str(values.get("-SUBJECTFILE-", "")).strip()
    if subject_file_path:
        try:
            subject_template = Path(subject_file_path).read_text(encoding="utf-8").strip()
        except FileNotFoundError:
            sg.popup_error("Arquivo de assunto informado não foi encontrado.")
            return None
        except OSError as exc:
            sg.popup_error(f"Erro ao ler o arquivo de assunto: {exc}")
            return None

    if not subject_template:
        sg.popup_error("Informe o assunto (template).")
        return None

    body_html_path = str(values.get("-HTML-", "")).strip()
    body_html = _read_html_template(body_html_path)
    if body_html is None:
        return None
    if not body_html.strip():
        sg.popup_error("O arquivo de template HTML está vazio.")
        return None

    sheet_value_raw = str(values.get("-SHEET-", "") or "").strip()
    sheet_value = sheet_value_raw or None
    if sheet_value and current_sheets and sheet_value not in current_sheets:
        sg.popup_error("A aba selecionada não foi encontrada. Recarregue a planilha e tente novamente.")
        return None

    sender_value = str(values.get("-SENDER-", "")).strip()
    dry_run_value = (
        dry_run_override
        if dry_run_override is not None
        else bool(values.get("-DRYRUN-", True))
    )

    smtp_user_input = str(values.get("-SMTPUSER-", "")).strip()
    if not dry_run_value:
        if not sender_value:
            sg.popup_error("Informe o remetente (From).")
            return None
        if not smtp_user_input:
            sg.popup_error("Informe o SMTP User.")
            return None

    smtp_user_value = smtp_user_input or sender_value

    password_input = str(values.get("-SMTPPASS-", "")).strip()
    env_password = os.getenv("SMTP_PASSWORD", "")
    password_to_use = password_input or env_password

    if not dry_run_value and not password_to_use:
        sg.popup_error(
            "Informe a senha SMTP ou defina a variável de ambiente SMTP_PASSWORD."
        )
        return None

    log_level_value = str(values.get("-LOGLEVEL-", "") or "INFO").strip() or "INFO"
    allow_missing_value = bool(values.get("-ALLOWMISSING-", False))

    return RunParams(
        input_path=excel_path,
        sheet=sheet_value,
        sender=sender_value,
        smtp_user=smtp_user_value or sender_value,
        smtp_password=password_to_use,
        subject_template=subject_template,
        body_html=body_html,
        dry_run=dry_run_value,
        log_level=log_level_value,
        allow_missing_fields=allow_missing_value,
    )


def _start_worker(window: sg.Window, params: RunParams, *, mode: str) -> None:
    if run_program is None:
        sg.popup_error("Função de execução indisponível. Verifique a instalação.")
        return

    global worker_thread  # pylint: disable=global-statement

    action_label = "prévia" if mode == "preview" else "envio"
    append_log(window, f"\n[INFO] Iniciando {action_label}\n")

    sheet_display = params.sheet or "(padrão)"
    sender_display = params.sender or "(não informado)"
    smtp_user_display = params.smtp_user or sender_display
    mode_display = "Dry-run (sem envio real)" if params.dry_run else "Envio real"
    password_line = "[INFO] Senha: ********" if params.smtp_password else "[INFO] Senha: (não informada)"

    append_log(
        window,
        "[INFO] Planilha: {} | Sheet: {} | Remetente: {} | SMTP User: {}\n".format(
            params.input_path,
            sheet_display,
            sender_display,
            smtp_user_display,
        ),
    )
    append_log(window, f"[INFO] Assunto: {params.subject_template}\n")
    append_log(window, f"[INFO] Modo: {mode_display}\n")
    append_log(window, password_line + "\n")
    append_log(window, f"[INFO] Log level: {params.log_level or 'INFO'}\n")
    if params.allow_missing_fields:
        append_log(
            window,
            "[INFO] Modo tolerante: placeholders ausentes serão preenchidos com vazio.\n",
        )

    def _run() -> None:  # pragma: no cover - integração com UI
        try:
            result_code = run_program(params)
        except Exception as exc:  # pylint: disable=broad-except
            log_queue.put(("ERROR", f"Falha durante a execução: {exc}\n"))
        else:
            log_queue.put(("RESULT", f"{mode}:{result_code}"))

    worker_thread = threading.Thread(target=_run, daemon=True)
    worker_thread.start()


def _show_preview(params: RunParams) -> bool:
    if load_contacts is None:
        sg.popup_error("Função de carregamento indisponível. Verifique a instalação.")
        return False

    try:
        dataframe = load_contacts(params.input_path, params.sheet)
    except Exception as exc:  # pylint: disable=broad-except
        sg.popup_error(f"Erro ao carregar planilha: {exc}")
        return False

    dataframe = _normalize_headers(dataframe)

    missing_required = sorted(REQUIRED_COLUMNS - set(dataframe.columns))
    if missing_required:
        formatted = ", ".join(missing_required)
        sg.popup_error(
            "Planilha inválida. Colunas obrigatórias ausentes: " f"{formatted}."
        )
        return False

    used_placeholders = extract_placeholders(
        params.subject_template
    ) | extract_placeholders(params.body_html)
    normalized_columns = {str(column).strip().lower() for column in dataframe.columns}
    missing_placeholders = sorted(
        placeholder
        for placeholder in used_placeholders
        if placeholder.lower() not in normalized_columns | GLOBAL_PLACEHOLDERS
    )
    if missing_placeholders:
        if params.allow_missing_fields:
            for name in missing_placeholders:
                logging.warning(
                    "Placeholder '%s' não encontrado na planilha. Valor vazio será usado.",
                    name,
                )
        else:
            formatted = "\n".join(f"• {name}" for name in missing_placeholders)
            sg.popup_error(
                "Variáveis ausentes no template:\n"
                f"{formatted}\n\n"
                "Adicione as colunas na planilha ou remova os placeholders do template."
            )
            return False

    if dataframe.empty:
        sg.popup(
            "Planilha carregada, mas não há registros para pré-visualizar.",
            title="Prévia",
        )
        return False

    previews: list[str] = []
    records = dataframe.to_dict(orient="records")
    for index, row in enumerate(records[:3], start=1):
        missing_for_row: list[str] = []

        def _handle_missing(placeholder: str) -> None:
            missing_for_row.append(placeholder)

        try:
            rendered_subject, rendered_body = render(
                params.subject_template,
                params.body_html,
                row,
                allow_missing=params.allow_missing_fields,
                on_missing=_handle_missing if params.allow_missing_fields else None,
            )
        except TemplateRenderingError as exc:
            sg.popup_error(str(exc))
            return False

        if params.allow_missing_fields and missing_for_row:
            for placeholder in sorted(set(missing_for_row)):
                logging.warning(
                    "Linha %s: placeholder '%s' ausente na prévia. Valor vazio utilizado.",
                    index,
                    placeholder,
                )

        plain_body = re.sub(r"<[^>]+>", "", rendered_body)
        plain_body = re.sub(r"\s+", " ", plain_body).strip()
        plain_body = plain_body[:200]

        previews.append(
            "Amostra {index}\nAssunto: {subject}\nCorpo: {body}".format(
                index=index, subject=rendered_subject, body=plain_body
            )
        )

    preview_text = "\n\n".join(previews)
    sg.popup_scrolled(preview_text, title="Prévia", non_blocking=False)
    return True


def _show_html_quick_preview(html_path: str) -> None:
    content = _read_html_template(html_path)
    if content is None:
        return
    stripped_content = content.strip()
    if not stripped_content:
        sg.popup_error("O arquivo de template HTML está vazio.")
        return

    first_line = stripped_content.splitlines()[0] if stripped_content else ""
    preview_layout = [
        [sg.Text("Primeira linha do template HTML:")],
        [
            sg.Multiline(
                first_line,
                size=(80, 5),
                disabled=True,
                autoscroll=False,
                font=("Consolas", 10),
            )
        ],
        [sg.Button("Fechar")],
    ]
    preview_window = sg.Window(
        "Pré-visualização rápida do HTML",
        preview_layout,
        modal=True,
        finalize=True,
    )
    while True:
        event, _ = preview_window.read()
        if event in (sg.WIN_CLOSED, "Fechar"):
            break
    preview_window.close()


# -------- UI --------
sg.theme("SystemDefault")

layout = [
    [
        sg.Text("Planilha (XLSX/CSV)"),
        sg.Input(key="-EXCEL-", enable_events=True),
        sg.FileBrowse(key="-EXCEL-BROWSE-", file_types=(("Excel/CSV", "*.xlsx;*.xls;*.csv"),)),
    ],
    [
        sg.Text("Aba (sheet)"),
        sg.Combo(
            values=[],
            key="-SHEET-",
            size=(25, 1),
            readonly=True,
            disabled=True,
            enable_events=True,
        ),
    ],
    [sg.HorizontalSeparator()],
    [
        sg.Text("Remetente (From)"),
        sg.Input(key="-SENDER-", size=(40, 1), enable_events=True),
    ],
    [
        sg.Text("SMTP User"),
        sg.Input(key="-SMTPUSER-", size=(40, 1), enable_events=True),
    ],
    [sg.Text("SMTP Password"), sg.Input(key="-SMTPPASS-", password_char="*", size=(40, 1))],
    [sg.HorizontalSeparator()],
    [
        sg.Text("Assunto (Jinja2)"),
        sg.Input(key="-SUBJECT-", size=(60, 1), enable_events=True),
    ],
    [
        sg.Text("Assunto por arquivo (.txt)"),
        sg.Input(key="-SUBJECTFILE-", enable_events=True, size=(60, 1)),
        sg.FileBrowse(key="-SUBJECTFILE-BROWSE-", file_types=(("Texto", "*.txt;*.jinja;*.j2"),)),
    ],
    [
        sg.Text("Template HTML"),
        sg.Input(key="-HTML-", enable_events=True),
        sg.FileBrowse(key="-HTML-BROWSE-", file_types=(("HTML", "*.html;*.htm;*.j2"),)),
        sg.Button("Preview", key="-HTML-PREVIEW-"),
    ],
    [
        sg.Checkbox(
            "Dry-run (não enviar, apenas pré-visualizar)",
            key="-DRYRUN-",
            default=True,
            enable_events=True,
        )
    ],
    [
        sg.Checkbox(
            "Permitir campos ausentes (preencher vazio)",
            key="-ALLOWMISSING-",
            default=False,
            enable_events=True,
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
        sg.ProgressBar(
            max_value=100,
            orientation="h",
            size=(40, 20),
            key="-PROGRESS-",
            visible=False,
            bar_color=("#1f77b4", "#e0e0e0"),
        ),
        sg.Text("Pronto", key="-STATUS-", size=(20, 1)),
        sg.Text("0/0", key="-COUNTER-", size=(10, 1)),
    ],
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
window["-COUNTER-"].update("0/0")
_apply_saved_settings(window)
_set_validation_state(window, passed=False)

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
        _save_settings(values)
        break

    if event in _VALIDATION_RESET_EVENTS:
        _set_validation_state(window, passed=False)

    if event == "-EXCEL-":
        excel_path = values["-EXCEL-"].strip()
        _update_sheet_combo(window, excel_path)
        _save_settings(values)

    if event == "-SHEET-":
        _save_settings(values)

    if event == "-HTML-":
        _save_settings(values)

    if event == "-SENDER-":
        _save_settings(values)

    if event == "-SUBJECTFILE-":
        subject_file_path = str(values.get("-SUBJECTFILE-", "")).strip()
        if subject_file_path and Path(subject_file_path).exists():
            try:
                subject_text = Path(subject_file_path).read_text(encoding="utf-8").strip()
            except OSError as exc:
                sg.popup_error(f"Erro ao ler o arquivo de assunto: {exc}")
                continue
            values["-SUBJECT-"] = subject_text
            window["-SUBJECT-"].update(subject_text)

    if event == "-HTML-PREVIEW-":
        html_path = str(values.get("-HTML-", "")).strip()
        if not html_path:
            sg.popup_error("Selecione o arquivo de template HTML para visualizar.")
        else:
            _show_html_quick_preview(html_path)
        continue

    if event in ("-PREVIEW-", "-RUN-"):
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

        dry_run_override = True if event == "-PREVIEW-" else None
        params = _prepare_run_params(values, dry_run_override=dry_run_override)
        if params is None:
            continue

        window["-SUBJECT-"].update(params.subject_template)
        values["-SUBJECT-"] = params.subject_template
        _save_settings(values)
        _set_running_state(window, running=True)

        if event == "-PREVIEW-" and not _show_preview(params):
            _set_running_state(window, running=False)
            _set_validation_state(window, passed=False)
            continue

        if event == "-PREVIEW-":
            _set_validation_state(window, passed=True)

        mode = "preview" if event == "-PREVIEW-" else "run"
        _start_worker(window, params, mode=mode)

    try:
        while True:
            tag, payload = log_queue.get_nowait()
            if tag == "LOG":
                append_log(window, payload)
                _handle_progress_from_log(window, payload)
            elif tag == "ERROR":
                append_log(window, payload, tag="ERR")
                worker_thread = None
                _set_running_state(window, running=False)
            elif tag == "RESULT":
                mode_label, separator, code = payload.partition(":")
                if separator:
                    description = "Prévia" if mode_label == "preview" else "Execução"
                    result_code = code
                else:
                    description = "Execução"
                    result_code = mode_label
                append_log(
                    window,
                    f"\n[INFO] {description} finalizada com código {result_code}\n",
                )
                worker_thread = None
                _set_running_state(window, running=False)
    except queue.Empty:
        pass

window.close()
