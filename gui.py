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
    import email_sender as email_sender_module
except ImportError as exc:  # pragma: no cover - integração com UI
    IMPORT_ERROR: Optional[Exception] = exc
    email_sender_module = None  # type: ignore[assignment]
    RunParams = None  # type: ignore[assignment]
    load_contacts = None  # type: ignore[assignment]
    run_program = None  # type: ignore[assignment]
else:
    IMPORT_ERROR = None

from emaileria.templating import TemplateRenderingError, extract_placeholders, render
from emaileria.preview import build_preview_page

REQUIRED_COLUMNS = {"email", "tratamento", "nome"}
SETTINGS_PATH = Path.home() / ".emaileria_gui.json"
PROCESSING_PATTERN = re.compile(
    r"Processando\s+(?P<processed>\d+)\s+contatos.*total[^0-9]*(?P<total>\d+)",
    re.IGNORECASE,
)

GLOBAL_PLACEHOLDERS = {"now", "hoje", "data_envio", "hora_envio"}
EMAIL_PATTERN = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")

current_sheets: list[str] = []


def _load_settings() -> dict[str, object]:
    if not SETTINGS_PATH.exists():
        return {}
    try:
        data = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return {}
    if not isinstance(data, dict):
        return {}
    return data


def _save_settings(values: Mapping[str, object] | None) -> None:
    if values is None:
        return

    relevant = {
        "excel": str(values.get("-EXCEL-", "") or ""),
        "html": str(values.get("-HTML-", "") or ""),
        "sheet": str(values.get("-SHEET-", "") or ""),
        "sender": str(values.get("-SENDER-", "") or ""),
        "smtp_user": str(values.get("-SMTPUSER-", "") or ""),
        "log_level": str(values.get("-LOGLEVEL-", "INFO") or "INFO"),
        "dry_run": bool(values.get("-DRYRUN-", True)),
        "interval": float(values.get("-INTERVAL-", 0.75) or 0.0),
        "cc": str(values.get("-CC-", "") or ""),
        "bcc": str(values.get("-BCC-", "") or ""),
        "reply_to": str(values.get("-REPLYTO-", "") or ""),
        "subject_file": str(values.get("-SUBJECTFILE-", "") or ""),
    }
    try:
        SETTINGS_PATH.write_text(
            json.dumps(relevant, ensure_ascii=False, indent=2), encoding="utf-8"
        )
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


def _parse_email_list(raw_value: str) -> list[str]:
    if not raw_value:
        return []
    normalized = raw_value.replace("\n", ",")
    entries = [part.strip() for part in re.split(r"[;,]", normalized) if part.strip()]
    invalid = [entry for entry in entries if not EMAIL_PATTERN.match(entry)]
    if invalid:
        raise ValueError("Endereços de e-mail inválidos: " + ", ".join(invalid))
    return entries


def _update_interval_display(window: sg.Window, value: float) -> None:
    window["-INTERVAL-LABEL-"].update(f"{value:.2f}s")


INTERACTIVE_KEYS = [
    "-EXCEL-",
    "-EXCEL-BROWSE-",
    "-SHEET-",
    "-SENDER-",
    "-SMTPUSER-",
    "-SMTPPASS-",
    "-CC-",
    "-BCC-",
    "-REPLYTO-",
    "-SUBJECT-",
    "-SUBJECTFILE-",
    "-SUBJECTFILE-BROWSE-",
    "-HTML-",
    "-HTML-BROWSE-",
    "-HTML-PREVIEW-",
    "-DRYRUN-",
    "-LOGLEVEL-",
    "-INTERVAL-",
    "-VALIDATE-",
    "-RUN-",
    "-OPEN-LAST-PREVIEW-",
    "-EXIT-",
]

progress_state: dict[str, int] = {"total": 0, "sent": 0}
validation_passed = False
last_preview_path: Path | None = None

_VALIDATION_RESET_EVENTS = {
    "-EXCEL-",
    "-SHEET-",
    "-SENDER-",
    "-SMTPUSER-",
    "-SUBJECT-",
    "-SUBJECTFILE-",
    "-HTML-",
    "-DRYRUN-",
    "-CC-",
    "-BCC-",
    "-REPLYTO-",
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


def _open_preview_page(index_path: Path, *, window: sg.Window | None = None) -> None:
    resolved = index_path.resolve()
    selection = sg.popup(
        f"Gerado em: {resolved}",
        title="Prévia gerada",
        custom_text=("Abrir no navegador", "Ver na janela"),
    )
    if selection == "Abrir no navegador":
        import webbrowser

        webbrowser.open(resolved.as_uri())
        return

    try:
        import webview
    except Exception:
        import webbrowser

        webbrowser.open(resolved.as_uri())
        return

    try:
        if window is not None:
            window.disappear()
        webview.create_window(
            "Prévia",
            url=str(resolved.as_uri()),
            width=900,
            height=720,
        )
        gui_hint = "cef" if getattr(webview, "platform", None) == "Windows" else None
        webview.start(gui=gui_hint)
    except Exception:
        import webbrowser

        webbrowser.open(resolved.as_uri())
    finally:
        if window is not None:
            try:
                window.reappear()
            except Exception:  # pylint: disable=broad-except
                pass


def _find_latest_preview() -> Path | None:
    base = Path("previews")
    if not base.exists():
        return None
    try:
        candidates = list(base.glob("*/index.html"))
    except OSError:
        return None
    if not candidates:
        return None
    latest: Path | None = None
    latest_mtime = float("-inf")
    for candidate in candidates:
        try:
            mtime = candidate.stat().st_mtime
        except OSError:
            continue
        if mtime > latest_mtime:
            latest = candidate
            latest_mtime = mtime
    return latest


def _set_validation_state(window: sg.Window, *, passed: bool) -> None:
    global validation_passed  # pylint: disable=global-statement
    validation_passed = passed
    _update_run_button_state(window)


def _set_running_state(window: sg.Window, *, running: bool) -> None:
    progress_bar = window["-PROGRESS-"]
    cancel_button = window["-CANCEL-"]
    if running:
        progress_state["total"] = 0
        progress_state["sent"] = 0
        _set_controls_enabled(window, enabled=False)
        window["-STATUS-"].update("Processando...")
        progress_bar.update(current_count=0, max=1, visible=True)
        _update_counter_display(window, unknown_total=True)
        try:
            progress_bar.Widget.start(10)
        except Exception:  # pylint: disable=broad-except
            pass
        cancel_button.update(disabled=False)
        return

    try:
        progress_bar.Widget.stop()
    except Exception:  # pylint: disable=broad-except
        pass
    progress_bar.update(visible=False)
    window["-STATUS-"].update("Pronto")
    _set_controls_enabled(window, enabled=True)
    cancel_button.update(disabled=True)
    if progress_state["total"] == 0:
        window["-COUNTER-"].update("0/0")
    else:
        _update_counter_display(window)


def _apply_saved_settings(window: sg.Window) -> None:
    settings = _load_settings()
    if not settings:
        return

    excel_path = str(settings.get("excel", "") or "")
    if excel_path:
        window["-EXCEL-"].update(excel_path)
        if Path(excel_path).exists():
            _update_sheet_combo(window, excel_path)

    sheet_value = str(settings.get("sheet", "") or "")
    if sheet_value and sheet_value in current_sheets:
        window["-SHEET-"].update(value=sheet_value)

    html_path = str(settings.get("html", "") or "")
    if html_path:
        window["-HTML-"].update(html_path)

    sender_value = str(settings.get("sender", "") or "")
    if sender_value:
        window["-SENDER-"].update(sender_value)

    smtp_user = str(settings.get("smtp_user", "") or "")
    if smtp_user:
        window["-SMTPUSER-"].update(smtp_user)

    cc_value = str(settings.get("cc", "") or "")
    bcc_value = str(settings.get("bcc", "") or "")
    reply_value = str(settings.get("reply_to", "") or "")
    if cc_value:
        window["-CC-"].update(cc_value)
    if bcc_value:
        window["-BCC-"].update(bcc_value)
    if reply_value:
        window["-REPLYTO-"].update(reply_value)

    interval_value = settings.get("interval")
    if isinstance(interval_value, (int, float)):
        window["-INTERVAL-"].update(value=float(interval_value))
        _update_interval_display(window, float(interval_value))

    log_level = str(settings.get("log_level", "INFO") or "INFO")
    window["-LOGLEVEL-"].update(log_level)

    dry_run_state = settings.get("dry_run")
    if isinstance(dry_run_state, bool):
        window["-DRYRUN-"].update(dry_run_state)

    subject_file = str(settings.get("subject_file", "") or "")
    if subject_file:
        window["-SUBJECTFILE-"].update(subject_file)


def _handle_progress_from_log(window: sg.Window, message: str) -> None:
    match = PROCESSING_PATTERN.search(message)
    if match:
        try:
            progress_state["total"] = int(match.group("processed"))
        except (TypeError, ValueError):
            progress_state["total"] = 0
        progress_state["sent"] = 0
        _update_counter_display(window, unknown_total=progress_state["total"] == 0)
        if progress_state["total"] > 0:
            try:
                progress_bar = window["-PROGRESS-"]
                progress_bar.update(current_count=0, max=progress_state["total"], visible=True)
                progress_bar.Widget.stop()
            except Exception:  # pylint: disable=broad-except
                pass
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
        if progress_state.get("total", 0) > 0:
            try:
                progress_bar = window["-PROGRESS-"]
                progress_bar.update(current_count=progress_state["sent"])
            except Exception:  # pylint: disable=broad-except
                pass


def _update_sheet_combo(window: sg.Window, file_path: str) -> None:
    sheet_element = window["-SHEET-"]
    sheet_element.update(values=[], value="", disabled=True)
    global current_sheets  # pylint: disable=global-statement
    current_sheets = []
    if not file_path:
        return

    suffix = Path(file_path).suffix.lower()
    if suffix not in {".xlsx", ".xls"}:
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

    excel_path = str(values.get("-EXCEL-", "") or "").strip()
    if not excel_path:
        sg.popup_error("Selecione a planilha (XLSX/CSV).")
        return None
    if not Path(excel_path).exists():
        sg.popup_error("Arquivo de planilha não encontrado. Verifique o caminho informado.")
        return None

    sheet_value_raw = str(values.get("-SHEET-", "") or "").strip()
    sheet_value = sheet_value_raw or None
    if sheet_value and current_sheets and sheet_value not in current_sheets:
        sg.popup_error(
            "A aba selecionada não foi encontrada. Recarregue a planilha e tente novamente."
        )
        return None

    subject_template = str(values.get("-SUBJECT-", "") or "").strip()
    subject_file_path = str(values.get("-SUBJECTFILE-", "") or "").strip()
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

    body_html_path = str(values.get("-HTML-", "") or "").strip()
    body_html = _read_html_template(body_html_path)
    if body_html is None:
        return None
    if not body_html.strip():
        sg.popup_error("O arquivo de template HTML está vazio.")
        return None

    dry_run_value = (
        dry_run_override if dry_run_override is not None else bool(values.get("-DRYRUN-", True))
    )

    sender_value = str(values.get("-SENDER-", "") or "").strip()
    smtp_user_input = str(values.get("-SMTPUSER-", "") or "").strip()
    smtp_user_value = smtp_user_input or sender_value

    if not dry_run_value:
        if not sender_value:
            sg.popup_error("Informe o remetente (From).")
            return None
        if not smtp_user_value:
            sg.popup_error("Informe o SMTP User.")
            return None

    try:
        cc_list = _parse_email_list(str(values.get("-CC-", "") or ""))
        bcc_list = _parse_email_list(str(values.get("-BCC-", "") or ""))
    except ValueError as exc:
        sg.popup_error(str(exc))
        return None

    reply_to_raw = str(values.get("-REPLYTO-", "") or "").strip()
    reply_to_value: str | None = None
    if reply_to_raw:
        if not EMAIL_PATTERN.match(reply_to_raw):
            sg.popup_error("Reply-To inválido. Informe um endereço de e-mail válido.")
            return None
        reply_to_value = reply_to_raw

    interval_raw = values.get("-INTERVAL-", 0.75)
    try:
        interval_value = float(interval_raw)
    except (TypeError, ValueError):
        interval_value = 0.75
    interval_value = max(0.0, min(2.0, interval_value))

    log_level_value = str(values.get("-LOGLEVEL-", "INFO") or "INFO").strip().upper()
    if log_level_value not in {"DEBUG", "INFO", "WARNING", "ERROR"}:
        log_level_value = "INFO"

    password_input = str(values.get("-SMTPPASS-", "") or "").strip()
    if not dry_run_value and not (password_input or os.getenv("SMTP_PASSWORD", "").strip()):
        sg.popup_error("Informe a senha SMTP ou defina a variável de ambiente SMTP_PASSWORD.")
        return None

    values["-SUBJECT-"] = subject_template

    return RunParams(
        input_path=excel_path,
        sheet=sheet_value,
        sender=sender_value,
        smtp_user=smtp_user_value,
        smtp_password=password_input,
        subject_template=subject_template,
        body_html=body_html,
        dry_run=dry_run_value,
        log_level=log_level_value,
        cc=cc_list or None,
        bcc=bcc_list or None,
        reply_to=reply_to_value,
        interval_seconds=interval_value,
    )


def _validate_and_preview(window: sg.Window, values: dict[str, object]) -> bool:
    if load_contacts is None:
        sg.popup_error("Função de carregamento indisponível. Verifique a instalação.")
        return False

    params = _prepare_run_params(values, dry_run_override=True)
    if params is None:
        return False

    window["-SUBJECT-"].update(params.subject_template)

    try:
        dataframe = load_contacts(params.input_path, params.sheet)
    except ValueError as exc:
        sg.popup_error(str(exc))
        return False

    dataframe = _normalize_headers(dataframe)

    missing_required = sorted(REQUIRED_COLUMNS - set(dataframe.columns))
    if missing_required:
        sg.popup_error(
            "Planilha inválida. Colunas obrigatórias ausentes: " + ", ".join(missing_required)
        )
        return False

    used_placeholders = extract_placeholders(params.subject_template) | extract_placeholders(
        params.body_html
    )
    available_placeholders = {column.lower() for column in dataframe.columns} | {
        name.lower() for name in GLOBAL_PLACEHOLDERS
    }
    missing_placeholders = sorted(
        placeholder
        for placeholder in used_placeholders
        if placeholder.lower() not in available_placeholders
    )
    if missing_placeholders:
        formatted = "\n".join(f"• {name}" for name in missing_placeholders)
        sg.popup_error(
            "Variáveis ausentes no template:\n"
            f"{formatted}\n\n"
            "Adicione as colunas na planilha ou ajuste o template."
        )
        return False

    if dataframe.empty:
        sg.popup_error("A planilha não possui registros para envio.")
        return False

    sample_records = dataframe.head(3).to_dict(orient="records")
    previews: list[dict[str, str]] = []
    for index, row in enumerate(sample_records, start=1):
        cleaned_row = {key: ("" if pd.isna(value) else value) for key, value in row.items()}
        try:
            rendered_subject, rendered_body = render(
                params.subject_template,
                params.body_html,
                cleaned_row,
            )
        except TemplateRenderingError as exc:
            sg.popup_error(str(exc))
            return False

        previews.append(
            {
                "idx": index,
                "subject": str(rendered_subject or ""),
                "body_html": str(rendered_body or ""),
                "email": str(cleaned_row.get("email", "") or ""),
            }
        )

    try:
        index_path = build_preview_page(previews)
    except OSError as exc:
        sg.popup_error(f"Não foi possível criar a prévia: {exc}")
        return False

    global last_preview_path  # pylint: disable=global-statement
    last_preview_path = index_path
    _open_preview_page(index_path, window=window)
    return True


def _start_worker(window: sg.Window, params: RunParams) -> None:
    if run_program is None:
        sg.popup_error("Função de execução indisponível. Verifique a instalação.")
        return

    global worker_thread  # pylint: disable=global-statement

    if email_sender_module is not None and hasattr(email_sender_module, "reset_cancel_flag"):
        try:
            email_sender_module.reset_cancel_flag()
        except Exception:  # pylint: disable=broad-except
            pass

    append_log(window, "\n[INFO] Iniciando execução\n")
    append_log(window, f"[INFO] Planilha: {params.input_path}\n")
    append_log(window, f"[INFO] Sheet: {params.sheet or '(padrão)'}\n")
    append_log(window, f"[INFO] Remetente: {params.sender or '(não informado)'}\n")
    append_log(window, f"[INFO] SMTP User: {params.smtp_user or '(não informado)'}\n")
    if params.cc:
        append_log(window, f"[INFO] CC: {', '.join(params.cc)}\n")
    if params.bcc:
        append_log(window, f"[INFO] BCC: {', '.join(params.bcc)}\n")
    if params.reply_to:
        append_log(window, f"[INFO] Reply-To: {params.reply_to}\n")
    append_log(window, f"[INFO] Intervalo entre envios: {params.interval_seconds:.2f}s\n")
    append_log(window, f"[INFO] Log level: {params.log_level}\n")
    mode_display = "Dry-run (sem envio real)" if params.dry_run else "Envio real"
    append_log(window, f"[INFO] Modo: {mode_display}\n")
    append_log(window, f"[INFO] Assunto: {params.subject_template}\n")

    def _run() -> None:  # pragma: no cover - integração com UI
        try:
            result_code = run_program(params)
        except Exception as exc:  # pylint: disable=broad-except
            log_queue.put(("ERROR", f"Falha durante a execução: {exc}\n"))
        else:
            log_queue.put(("RESULT", str(result_code)))

    worker_thread = threading.Thread(target=_run, daemon=True)
    worker_thread.start()


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
        sg.Input(key="-EXCEL-", enable_events=True, size=(45, 1)),
        sg.FileBrowse(
            key="-EXCEL-BROWSE-",
            file_types=(("Excel/CSV", "*.xlsx;*.xls;*.csv"),),
        ),
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
    [sg.Text("Remetente (From)"), sg.Input(key="-SENDER-", size=(40, 1), enable_events=True)],
    [sg.Text("SMTP User"), sg.Input(key="-SMTPUSER-", size=(40, 1), enable_events=True)],
    [sg.Text("SMTP Password"), sg.Input(key="-SMTPPASS-", password_char="*", size=(40, 1))],
    [sg.Text("CC"), sg.Input(key="-CC-", size=(60, 1), enable_events=True)],
    [sg.Text("BCC"), sg.Input(key="-BCC-", size=(60, 1), enable_events=True)],
    [sg.Text("Reply-To"), sg.Input(key="-REPLYTO-", size=(60, 1), enable_events=True)],
    [sg.HorizontalSeparator()],
    [sg.Text("Assunto (Jinja2)"), sg.Input(key="-SUBJECT-", size=(60, 1), enable_events=True)],
    [
        sg.Text("Assunto por arquivo (.txt)"),
        sg.Input(key="-SUBJECTFILE-", enable_events=True, size=(60, 1)),
        sg.FileBrowse(
            key="-SUBJECTFILE-BROWSE-",
            file_types=(("Texto", "*.txt;*.jinja;*.j2"),),
        ),
    ],
    [
        sg.Text("Template HTML"),
        sg.Input(key="-HTML-", enable_events=True, size=(60, 1)),
        sg.FileBrowse(
            key="-HTML-BROWSE-",
            file_types=(("HTML", "*.html;*.htm;*.j2"),),
        ),
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
        sg.Text("Intervalo entre envios (segundos)"),
        sg.Slider(
            range=(0.0, 2.0),
            resolution=0.05,
            default_value=0.75,
            orientation="h",
            key="-INTERVAL-",
            enable_events=True,
            size=(30, 15),
        ),
        sg.Text("0.75s", key="-INTERVAL-LABEL-", size=(10, 1)),
    ],
    [
        sg.Text("Log level"),
        sg.Combo(
            values=["INFO", "DEBUG", "WARNING", "ERROR"],
            default_value="INFO",
            key="-LOGLEVEL-",
            readonly=True,
            size=(15, 1),
            enable_events=True,
        ),
    ],
    [sg.HorizontalSeparator()],
    [
        sg.ProgressBar(
            max_value=1,
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
        sg.Button("Validar & Prévia", key="-VALIDATE-"),
        sg.Button("Abrir última prévia", key="-OPEN-LAST-PREVIEW-"),
        sg.Button("Enviar", key="-RUN-", bind_return_key=True, disabled=True),
        sg.Button("Cancelar", key="-CANCEL-", disabled=True),
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
_update_interval_display(window, float(window["-INTERVAL-"].DefaultValue))
window["-COUNTER-"].update("0/0")
_apply_saved_settings(window)
try:
    current_interval_value = float(window["-INTERVAL-"].Widget.get())
except Exception:  # pylint: disable=broad-except
    current_interval_value = float(window["-INTERVAL-"].DefaultValue)
_update_interval_display(window, current_interval_value)
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
    append_log(window, f"[ERR] Detalhes: {IMPORT_ERROR}\n", tag="ERR")

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

    if event in {"-SENDER-", "-SMTPUSER-", "-CC-", "-BCC-", "-REPLYTO-"}:
        _save_settings(values)

    if event == "-HTML-":
        _save_settings(values)

    if event == "-DRYRUN-":
        _save_settings(values)

    if event == "-LOGLEVEL-":
        _save_settings(values)

    if event == "-INTERVAL-":
        try:
            slider_value = float(values["-INTERVAL-"])
        except (TypeError, ValueError):
            slider_value = 0.75
        _update_interval_display(window, slider_value)
        _save_settings(values)

    if event == "-SUBJECTFILE-":
        subject_file_path = str(values.get("-SUBJECTFILE-", "") or "").strip()
        if subject_file_path and Path(subject_file_path).exists():
            try:
                subject_text = Path(subject_file_path).read_text(encoding="utf-8").strip()
            except OSError as exc:
                sg.popup_error(f"Erro ao ler o arquivo de assunto: {exc}")
            else:
                values["-SUBJECT-"] = subject_text
                window["-SUBJECT-"].update(subject_text)
                _save_settings(values)

    if event == "-HTML-PREVIEW-":
        html_path = str(values.get("-HTML-", "") or "").strip()
        if not html_path:
            sg.popup_error("Selecione o arquivo de template HTML para visualizar.")
        else:
            _show_html_quick_preview(html_path)
        continue

    if event == "-VALIDATE-":
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
        if _validate_and_preview(window, values):
            _set_validation_state(window, passed=True)
            _save_settings(values)
        continue

    if event == "-OPEN-LAST-PREVIEW-":
        global last_preview_path  # pylint: disable=global-statement
        candidate = last_preview_path
        if candidate is None or not candidate.exists():
            candidate = _find_latest_preview()
        if candidate is None or not candidate.exists():
            sg.popup_error("Nenhuma prévia foi gerada ainda.")
        else:
            last_preview_path = candidate
            _open_preview_page(candidate, window=window)
        continue

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
        params = _prepare_run_params(values)
        if params is None:
            continue
        window["-SUBJECT-"].update(params.subject_template)
        _save_settings(values)
        _set_running_state(window, running=True)
        _start_worker(window, params)
        continue

    if event == "-CANCEL-":
        if worker_thread is not None and worker_thread.is_alive():
            append_log(window, "[INFO] Cancelamento solicitado. Aguarde a finalização.\n")
            if email_sender_module is not None and hasattr(
                email_sender_module, "request_cancel"
            ):
                try:
                    email_sender_module.request_cancel()
                except Exception:  # pylint: disable=broad-except
                    pass
            window["-CANCEL-"].update(disabled=True)
        else:
            window["-CANCEL-"].update(disabled=True)

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
                worker_thread = None
                _set_running_state(window, running=False)
                try:
                    code = int(payload)
                except (TypeError, ValueError):
                    code = 1
                if code == 0:
                    append_log(
                        window,
                        "\n[INFO] Execução finalizada com sucesso (código 0)\n",
                    )
                elif code == 130:
                    append_log(window, "\n[INFO] Execução cancelada pelo usuário.\n")
                else:
                    append_log(
                        window,
                        f"\n[ERR] Execução finalizada com código {code}\n",
                        tag="ERR",
                    )
    except queue.Empty:
        pass

window.close()
