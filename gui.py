import os
import sys
import subprocess
import threading
import queue
from pathlib import Path
import PySimpleGUI as sg

# -------- Utilidades --------
def read_file_text(path: str) -> str:
    if not path:
        return ""
    p = Path(path)
    if not p.exists():
        return ""
    return p.read_text(encoding="utf-8")

def stream_process(cmd_list, cwd=None):
    """
    Executa um subprocesso emitindo stdout/stderr em tempo real usando fila.
    Retorna (returncode, output_text).
    """
    q = queue.Queue()
    output_lines = []

    def enqueue_output(pipe, tag):
        for line in iter(pipe.readline, b""):
            txt = line.decode(errors="ignore")
            q.put((tag, txt))
        pipe.close()

    proc = subprocess.Popen(
        cmd_list,
        cwd=cwd,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        bufsize=1
    )

    t_out = threading.Thread(target=enqueue_output, args=(proc.stdout, "OUT"), daemon=True)
    t_err = threading.Thread(target=enqueue_output, args=(proc.stderr, "ERR"), daemon=True)
    t_out.start()
    t_err.start()

    return proc, q, output_lines

def build_command(
    excel_path: str,
    sender: str,
    subject_template: str,
    body_html_path: str,
    sheet: str,
    smtp_user: str,
    smtp_password: str,
    dry_run: bool,
    log_level: str
):
    cmd = [sys.executable, "email_sender.py", excel_path]
    if sender:
        cmd += ["--sender", sender]
    if subject_template:
        cmd += ["--subject-template", subject_template]

    body_html = read_file_text(body_html_path)
    if body_html:
        cmd += ["--body-template", body_html]

    if sheet:
        cmd += ["--sheet", sheet]
    if smtp_user:
        cmd += ["--smtp-user", smtp_user]
    if smtp_password:
        cmd += ["--smtp-password", smtp_password]
    if dry_run:
        cmd += ["--dry-run"]
    if log_level:
        cmd += ["--log-level", log_level]

    return cmd

# -------- UI --------
sg.theme("SystemDefault")

layout = [
    [sg.Text("Planilha (XLSX/CSV)"), sg.Input(key="-EXCEL-"), sg.FileBrowse(file_types=(("Excel/CSV", "*.xlsx;*.xls;*.csv"),))],
    [sg.Text("Aba (sheet)"), sg.Input(key="-SHEET-", size=(25,1))],
    [sg.HorizontalSeparator()],
    [sg.Text("Remetente (From)"), sg.Input(key="-SENDER-", size=(40,1))],
    [sg.Text("SMTP User"), sg.Input(key="-SMTPUSER-", size=(40,1))],
    [sg.Text("SMTP Password"), sg.Input(key="-SMTPPASS-", password_char="*", size=(40,1))],
    [sg.HorizontalSeparator()],
    [sg.Text("Assunto (Jinja2)"), sg.Input(key="-SUBJECT-", size=(60,1))],
    [sg.Text("Template HTML"), sg.Input(key="-HTML-"), sg.FileBrowse(file_types=(("HTML", "*.html;*.htm;*.j2"),))],
    [sg.Checkbox("Dry-run (não enviar, apenas pré-visualizar)", key="-DRYRUN-", default=True)],
    [sg.Text("Log level"), sg.Combo(values=["INFO","DEBUG","WARNING","ERROR"], default_value="INFO", key="-LOGLEVEL-", readonly=True, size=(15,1))],
    [sg.HorizontalSeparator()],
    [sg.Button("Enviar", key="-RUN-", bind_return_key=True), sg.Button("Sair", key="-EXIT-")],
    [sg.Multiline(key="-LOG-", size=(100,20), autoscroll=True, write_only=True, font=("Consolas", 10))]
]

window = sg.Window("Emaileria — Envio de E-mails", layout, finalize=True)

proc = None
queue_stream = None
buffer_lines = []

def append_log(text, tag="OUT"):
    window["-LOG-"].print(text, end="")

while True:
    event, values = window.read(timeout=100)
    if event in (sg.WIN_CLOSED, "-EXIT-"):
        break

    if event == "-RUN-":
        excel = values["-EXCEL-"].strip()
        if not excel:
            sg.popup_error("Selecione a planilha (XLSX/CSV).")
            continue

        cmd = build_command(
            excel_path=excel,
            sender=values["-SENDER-"].strip(),
            subject_template=values["-SUBJECT-"].strip(),
            body_html_path=values["-HTML-"].strip(),
            sheet=values["-SHEET-"].strip(),
            smtp_user=values["-SMTPUSER-"].strip(),
            smtp_password=values["-SMTPPASS-"].strip(),
            dry_run=values["-DRYRUN-"],
            log_level=values["-LOGLEVEL-"]
        )

        append_log(f"\n$ {' '.join([('\"'+c+'\"' if ' ' in c else c) for c in cmd])}\n")
        try:
            proc, queue_stream, buffer_lines = stream_process(cmd)
        except Exception as e:
            append_log(f"[ERR] Falha ao iniciar o envio: {e}\n", tag="ERR")
            proc = None
            queue_stream = None

    # consumir fluxo
    if queue_stream is not None and proc is not None:
        try:
            while True:
                tag, line = queue_stream.get_nowait()
                append_log(line, tag=tag)
                buffer_lines.append(line)
        except queue.Empty:
            pass

        # terminou?
        if proc.poll() is not None:
            rc = proc.returncode
            append_log(f"\n[INFO] Finalizado com código {rc}\n")
            proc = None
            queue_stream = None

window.close()
