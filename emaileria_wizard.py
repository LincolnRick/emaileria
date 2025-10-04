#!/usr/bin/env python3
"""Assistente interativo para envio de e-mails personalizados."""

from __future__ import annotations

import argparse
import csv
import getpass
import re
import smtplib
import socket
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Iterable, List, Optional, Sequence, Tuple

from email.mime.text import MIMEText
from email.utils import formatdate, make_msgid, parseaddr
import pandas as pd

from emaileria.datasource.excel import load_contacts as load_contacts_dataframe
from emaileria.templating import TemplateRenderingError, extract_placeholders, render


CONTACT_PATTERNS = ("*.csv", "*.CSV", "*.xlsx", "*.XLSX")
TEMPLATE_PATTERNS = ("*.html", "*.HTML", "*.htm", "*.HTM", "*.j2", "*.J2")
REQUIRED_COLUMNS = ("email", "tratamento", "nome")
DEFAULT_INTERVAL = 0.75
MAX_ATTEMPTS = 3
BACKOFF_SECONDS = [1, 2, 4]
SMTP_TIMEOUT = 30
GLOBAL_PLACEHOLDERS = {"now", "hoje", "data_envio", "hora_envio"}

TAG_RE = re.compile(r"<[^>]+>")
WHITESPACE_RE = re.compile(r"\s+")


class PlaceholderRenderError(RuntimeError):
    """Erro personalizado para placeholders ausentes durante a renderização."""

    def __init__(self, line_number: int, field: str, placeholder: Optional[str], original: Exception) -> None:
        self.line_number = line_number
        self.field = field
        self.placeholder = placeholder
        self.original = original
        if placeholder:
            message = f"Linha {line_number}: placeholder '{placeholder}' ausente no {field}."
        else:
            message = f"Linha {line_number}: placeholder ausente no {field}."
        super().__init__(message)


def print_header() -> None:
    print("=" * 40)
    print("Emaileria — Wizard de Envio")
    print("=" * 40)
    print(
        "Este assistente enviará e-mails personalizados a partir de uma planilha e de um template HTML.\n"
    )


def parse_args(argv: Sequence[str]) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Assistente interativo para envio de e-mails personalizados.",
    )
    parser.add_argument(
        "--allow-missing",
        action="store_true",
        help="Permite placeholders ausentes, preenchendo com vazio e registrando avisos.",
    )
    return parser.parse_args(argv)


def ask_yes_no(prompt: str, default: bool = True) -> bool:
    default_text = "s" if default else "n"
    while True:
        answer = input(prompt).strip().lower()
        if not answer:
            return default
        if answer in {"s", "sim", "y", "yes"}:
            return True
        if answer in {"n", "nao", "não", "no"}:
            return False
        print(f"Digite 's' para sim ou 'n' para não (padrão: {default_text}).")


def prompt_non_empty(prompt: str) -> str:
    while True:
        value = input(prompt).strip()
        if value:
            return value
        print("Este campo é obrigatório.")


def prompt_port(prompt: str) -> int:
    while True:
        value = input(prompt).strip()
        try:
            port = int(value)
            if 0 < port < 65536:
                return port
        except ValueError:
            pass
        print("Informe um número de porta válido (1-65535).")


def prompt_interval(prompt: str, default: float) -> float:
    while True:
        value = input(prompt).strip()
        if not value:
            return default
        value = value.replace(",", ".")
        try:
            interval = float(value)
            if interval >= 0:
                return interval
        except ValueError:
            pass
        print("Informe um intervalo numérico maior ou igual a zero.")


def format_path_for_display(path: Path) -> str:
    try:
        return str(path.relative_to(Path.cwd()))
    except ValueError:
        return str(path)


def gather_candidates(patterns: Sequence[str], base_dirs: Iterable[Path]) -> List[Path]:
    results: List[Path] = []
    seen = set()
    for base in base_dirs:
        if not base.exists() or not base.is_dir():
            continue
        for pattern in patterns:
            for found in sorted(base.glob(pattern)):
                if not found.is_file():
                    continue
                resolved = found.resolve()
                if resolved not in seen:
                    seen.add(resolved)
                    results.append(resolved)
    return results


def pick_from_list_or_path(prompt: str, candidates: Sequence[Path]) -> Path:
    print(prompt)
    if candidates:
        for idx, candidate in enumerate(candidates, start=1):
            print(f"  {idx}) {format_path_for_display(candidate)}")
    else:
        print("  (Nenhum arquivo encontrado automaticamente.)")
    while True:
        choice = input("Digite o número desejado ou informe um caminho manual: ").strip()
        if not choice:
            print("Informe um número ou caminho válido.")
            continue
        if choice.isdigit() and candidates:
            index = int(choice)
            if 1 <= index <= len(candidates):
                return candidates[index - 1]
            print("Número fora da faixa apresentada.")
            continue
        manual_path = Path(choice).expanduser()
        if not manual_path.is_absolute():
            manual_path = Path.cwd() / manual_path
        if manual_path.exists() and manual_path.is_file():
            return manual_path.resolve()
        print("Arquivo não encontrado. Tente novamente.")


def choose_excel_sheet(path: Path) -> str:
    try:
        with pd.ExcelFile(path) as workbook:
            sheets = list(workbook.sheet_names)
    except Exception as exc:  # pragma: no cover - erro tratado via CLI
        raise RuntimeError(
            f"Não foi possível listar as abas de '{format_path_for_display(path)}': {exc}"
        ) from exc
    if not sheets:
        raise RuntimeError("Nenhuma aba foi encontrada no arquivo Excel.")
    if len(sheets) == 1:
        print(f"Aba detectada: {sheets[0]}")
        return sheets[0]
    print("Abas disponíveis no arquivo:")
    for idx, name in enumerate(sheets, start=1):
        print(f"  {idx}) {name}")
    while True:
        choice = input("Informe o número ou o nome da aba desejada: ").strip()
        if not choice:
            print("Selecione uma opção válida.")
            continue
        if choice.isdigit():
            index = int(choice)
            if 1 <= index <= len(sheets):
                return sheets[index - 1]
            print("Número de aba inválido.")
            continue
        if choice in sheets:
            return choice
        print("Aba não encontrada. Digite novamente.")


def load_contacts(path: Path, sheet: Optional[str] = None) -> pd.DataFrame:
    try:
        df = load_contacts_dataframe(path, sheet)
    except FileNotFoundError as exc:  # pragma: no cover - erro tratado via CLI
        raise RuntimeError(str(exc)) from exc
    except Exception as exc:  # pragma: no cover - erro tratado via CLI
        raise RuntimeError(
            f"Não foi possível carregar '{format_path_for_display(path)}': {exc}"
        ) from exc
    df = df.dropna(how="all")
    return df


def validate_schema(df: pd.DataFrame) -> None:
    if df.empty:
        raise ValueError("A planilha está vazia.")
    normalized = {col.lower(): col for col in df.columns}
    missing = [col for col in REQUIRED_COLUMNS if col not in normalized]
    if missing:
        formatted = ", ".join(missing)
        raise ValueError(
            "As colunas obrigatórias "
            f"{formatted} não foram encontradas (verificação sem diferenciar maiúsculas/minúsculas)."
        )


def build_context(row: Dict[str, Any], index: int) -> Dict[str, Any]:
    context: Dict[str, Any] = {}
    lower_aliases: Dict[str, Any] = {}
    for key, value in row.items():
        key_str = str(key).strip()
        normalized_value = "" if pd.isna(value) else value
        context[key_str] = normalized_value
        lower_aliases[key_str.lower()] = normalized_value
    for alias, value in lower_aliases.items():
        context.setdefault(alias, value)
    context["linha"] = index
    context["index"] = index
    return context


def build_records(df: pd.DataFrame) -> List[Dict[str, Any]]:
    records: List[Dict[str, Any]] = []
    skipped = 0
    for idx, row in enumerate(df.to_dict(orient="records"), start=1):
        context = build_context(row, idx)
        email_value = str(context.get("email", "")).strip()
        if not email_value:
            skipped += 1
            continue
        context["email"] = email_value
        tratamento = context.get("tratamento")
        if isinstance(tratamento, str):
            context["tratamento"] = tratamento.strip()
        nome = context.get("nome")
        if isinstance(nome, str):
            context["nome"] = nome.strip()
        records.append({"index": idx, "email": email_value, "context": context})
    if skipped:
        print(f"Aviso: {skipped} linha(s) foram ignoradas por estarem sem e-mail.")
    return records


def read_template_file(path: Path) -> str:
    try:
        return path.read_text(encoding="utf-8")
    except UnicodeDecodeError as exc:  # pragma: no cover - erro tratado via CLI
        raise RuntimeError(
            f"O template deve estar em UTF-8. Erro ao ler '{format_path_for_display(path)}': {exc}"
        ) from exc
    except OSError as exc:  # pragma: no cover - erro tratado via CLI
        raise RuntimeError(
            f"Não foi possível ler o template '{format_path_for_display(path)}': {exc}"
        ) from exc


def html_to_snippet(html: str, limit: int = 200) -> str:
    text = TAG_RE.sub(" ", html)
    text = WHITESPACE_RE.sub(" ", text).strip()
    if len(text) > limit:
        return text[:limit].rstrip() + "..."
    return text


def prepare_previews(
    records: Sequence[Dict[str, Any]],
    subject_template: str,
    body_template: str,
    *,
    allow_missing: bool,
) -> List[Tuple[int, str, str, str]]:
    previews: List[Tuple[int, str, str, str]] = []
    for position, record in enumerate(records, start=1):
        context = record["context"]
        missing_for_row: list[str] = []

        def _handle_missing(name: str) -> None:
            missing_for_row.append(name)

        try:
            subject, html_body = render(
                subject_template,
                body_template,
                context,
                allow_missing=allow_missing,
                on_missing=_handle_missing if allow_missing else None,
            )
        except TemplateRenderingError as exc:
            field = exc.template_type if exc.template_type in {"assunto", "corpo"} else "corpo"
            raise PlaceholderRenderError(
                record["index"], field, exc.placeholder, exc
            ) from exc

        if allow_missing and missing_for_row:
            for placeholder in sorted(set(missing_for_row)):
                print(
                    f"[AVISO] Linha {record['index']}: placeholder '{placeholder}' ausente na prévia. "
                    "Valor vazio utilizado."
                )

        if position <= 3:
            previews.append((record["index"], record["email"], subject, html_body))
    return previews


def is_temporary_smtp_error(exc: BaseException) -> bool:
    if isinstance(
        exc,
        (
            socket.timeout,
            TimeoutError,
            smtplib.SMTPServerDisconnected,
            smtplib.SMTPConnectError,
        ),
    ):
        return True
    if isinstance(exc, smtplib.SMTPResponseException):
        return 400 <= exc.smtp_code < 500
    if isinstance(exc, smtplib.SMTPRecipientsRefused):
        return any(400 <= code < 500 for code, _ in exc.recipients.values())
    return False


def send_all(
    host: str,
    port: int,
    user: str,
    password: str,
    from_address: str,
    records: Sequence[Dict[str, Any]],
    subject_template: str,
    body_template: str,
    interval: float,
    log_path: Path,
    *,
    allow_missing: bool,
) -> Dict[str, Any]:
    total = len(records)
    if total == 0:
        return {"sucesso": 0, "falha": 0, "total": 0, "log_path": log_path.resolve()}
    log_path.parent.mkdir(parents=True, exist_ok=True)
    log_exists = log_path.exists()
    envelope_from = parseaddr(from_address)[1] or user
    success_count = 0
    failure_count = 0
    with open(log_path, "a", newline="", encoding="utf-8") as log_file:
        writer = csv.writer(log_file)
        if not log_exists:
            writer.writerow(["timestamp", "email", "assunto", "status", "tentativas", "erro"])
        try:
            with smtplib.SMTP_SSL(host, port, timeout=SMTP_TIMEOUT) as server:
                try:
                    server.login(user, password)
                except smtplib.SMTPAuthenticationError as exc:
                    raise RuntimeError("Falha na autenticação SMTP: verifique usuário e senha.") from exc
                except smtplib.SMTPException as exc:
                    raise RuntimeError(
                        f"Erro durante a autenticação no servidor SMTP: {exc}"
                    ) from exc
                for position, record in enumerate(records, start=1):
                    context = record["context"]
                    context.setdefault("posicao_envio", position)
                    missing_for_row: list[str] = []

                    def _handle_missing(name: str) -> None:
                        missing_for_row.append(name)

                    try:
                        subject, html_body = render(
                            subject_template,
                            body_template,
                            context,
                            allow_missing=allow_missing,
                            on_missing=_handle_missing if allow_missing else None,
                        )
                    except TemplateRenderingError as exc:
                        field = (
                            exc.template_type if exc.template_type in {"assunto", "corpo"} else "corpo"
                        )
                        raise PlaceholderRenderError(
                            record["index"], field, exc.placeholder, exc
                        ) from exc

                    if allow_missing and missing_for_row:
                        for placeholder in sorted(set(missing_for_row)):
                            print(
                                f"[AVISO] Linha {record['index']}: placeholder '{placeholder}' ausente. "
                                "Valor vazio utilizado."
                            )
                    message = MIMEText(html_body, "html", "utf-8")
                    message["Subject"] = subject
                    message["From"] = from_address
                    message["To"] = record["email"]
                    message["Date"] = formatdate(localtime=True)
                    message["Message-ID"] = make_msgid()
                    attempts = 0
                    sent = False
                    last_error = ""
                    while attempts < MAX_ATTEMPTS and not sent:
                        attempts += 1
                        try:
                            server.sendmail(
                                envelope_from,
                                [record["email"]],
                                message.as_string(),
                            )
                        except (smtplib.SMTPException, OSError, socket.timeout) as exc:
                            last_error = str(exc)
                            if is_temporary_smtp_error(exc) and attempts < MAX_ATTEMPTS:
                                wait_index = min(attempts - 1, len(BACKOFF_SECONDS) - 1)
                                wait_time = BACKOFF_SECONDS[wait_index]
                                print(
                                    f"[{position}/{total}] {record['email']} — ERRO temporário (tentativa {attempts}): "
                                    f"{last_error}. Retentando em {wait_time}s."
                                )
                                time.sleep(wait_time)
                                continue
                            break
                        else:
                            sent = True
                    timestamp = datetime.now().isoformat(timespec="seconds")
                    if sent:
                        print(f"[{position}/{total}] {record['email']} — OK")
                        writer.writerow([timestamp, record["email"], subject, "sucesso", attempts, ""])
                        writer.flush()
                        success_count += 1
                        if position < total and interval > 0:
                            time.sleep(interval)
                    else:
                        error_message = last_error or "Falha desconhecida."
                        print(f"[{position}/{total}] {record['email']} — ERRO: {error_message}")
                        writer.writerow(
                            [timestamp, record["email"], subject, "erro", attempts, error_message]
                        )
                        writer.flush()
                        failure_count += 1
        except PlaceholderRenderError:
            raise
        except (smtplib.SMTPException, OSError, socket.timeout) as exc:
            raise RuntimeError(f"Falha na comunicação com o servidor SMTP: {exc}") from exc
    return {
        "sucesso": success_count,
        "falha": failure_count,
        "total": total,
        "log_path": log_path.resolve(),
    }


def main(argv: Optional[Sequence[str]] = None) -> None:
    if argv is None:
        argv = sys.argv[1:]
    else:
        argv = list(argv)

    args = parse_args(argv)
    allow_missing = bool(args.allow_missing)

    print_header()
    smtp_user = prompt_non_empty("SMTP User (e-mail do remetente): ")
    smtp_password = ""
    while not smtp_password:
        smtp_password = getpass.getpass("Senha SMTP: ").strip()
        if not smtp_password:
            print("A senha é obrigatória.")
    from_address_input = input(
        "Remetente (Nome <email@dominio.com>) [Enter para usar o SMTP User]: "
    ).strip()
    from_address = from_address_input or smtp_user
    use_gmail = ask_yes_no("Usar Gmail padrão (smtp.gmail.com:465)? (S/n): ", default=True)
    if use_gmail:
        host = "smtp.gmail.com"
        port = 465
    else:
        host = prompt_non_empty("Host SMTP: ")
        port = prompt_port("Porta SMTP: ")

    contact_dirs = [Path.cwd()]
    data_dir = Path.cwd() / "data"
    if data_dir.exists() and data_dir.is_dir():
        contact_dirs.append(data_dir)
        for child in sorted(data_dir.iterdir()):
            if child.is_dir():
                contact_dirs.append(child)

    contact_prompt = "\nSelecione o arquivo de contatos detectado ou informe o caminho manual:"
    contact_path = pick_from_list_or_path(
        contact_prompt, gather_candidates(CONTACT_PATTERNS, contact_dirs)
    )

    sheet_name = None
    if contact_path.suffix.lower() in {".xlsx", ".xlsm", ".xls"}:
        try:
            sheet_name = choose_excel_sheet(contact_path)
        except Exception as exc:
            print(f"Erro ao selecionar aba: {exc}", file=sys.stderr)
            sys.exit(1)

    try:
        contacts_df = load_contacts(contact_path, sheet_name)
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)

    try:
        validate_schema(contacts_df)
    except ValueError as exc:
        print(f"Erro de validação da planilha: {exc}", file=sys.stderr)
        sys.exit(1)

    records = build_records(contacts_df)
    if not records:
        print("Nenhum contato com e-mail válido foi encontrado na planilha.", file=sys.stderr)
        sys.exit(1)

    print(f"\n{len(records)} contato(s) serão considerados no envio.\n")

    template_dirs = [Path.cwd()]
    templates_dir = Path.cwd() / "templates"
    if templates_dir.exists() and templates_dir.is_dir():
        template_dirs.append(templates_dir)

    template_prompt = "Selecione o template de e-mail (HTML/Jinja2) ou informe o caminho manual:"
    template_path = pick_from_list_or_path(
        template_prompt, gather_candidates(TEMPLATE_PATTERNS, template_dirs)
    )

    try:
        template_content = read_template_file(template_path)
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)

    subject_text = prompt_non_empty(
        "Assunto (placeholders Jinja2, ex.: {{ tratamento }} {{ nome }}): "
    )

    used_placeholders = extract_placeholders(subject_text) | extract_placeholders(
        template_content
    )
    normalized_columns = {str(column).strip().lower() for column in contacts_df.columns}
    missing_placeholders = sorted(
        placeholder
        for placeholder in used_placeholders
        if placeholder.lower() not in normalized_columns | GLOBAL_PLACEHOLDERS
    )
    if missing_placeholders:
        if allow_missing:
            print("\nAviso: as seguintes variáveis não foram encontradas na planilha e serão preenchidas com vazio:")
            for name in missing_placeholders:
                print(f"  - {name}")
        else:
            print("\nAs seguintes variáveis estão faltando na planilha:")
            for name in missing_placeholders:
                print(f"  - {name}")
            print("Adicione as colunas na planilha ou remova os placeholders do template.")
            sys.exit(1)

    try:
        # Render com contexto vazio para validar sintaxe dos templates.
        render(subject_text, template_content, {}, allow_missing=True)
    except Exception as exc:  # pragma: no cover - erro tratado via CLI
        print(f"Erro ao preparar os templates Jinja2: {exc}", file=sys.stderr)
        sys.exit(1)

    try:
        previews = prepare_previews(
            records,
            subject_text,
            template_content,
            allow_missing=allow_missing,
        )
    except PlaceholderRenderError as exc:
        placeholder_text = f" '{exc.placeholder}'" if exc.placeholder else ""
        print(
            f"Erro de placeholder na linha {exc.line_number}{placeholder_text} do {exc.field}: {exc.original}",
            file=sys.stderr,
        )
        sys.exit(1)

    print("\nPré-visualização (até 3 primeiros contatos):")
    if not previews:
        print("  Nenhuma linha disponível para pré-visualizar.")
    else:
        for line_idx, email_addr, subject, html_body in previews:
            snippet = html_to_snippet(html_body)
            print(f"  Linha {line_idx} — {email_addr}")
            print(f"    Assunto: {subject}")
            print(f"    Corpo: {snippet}")

    confirm = input("\nConfirmar envio real? (s/N): ").strip().lower()
    if confirm not in {"s", "sim", "y", "yes"}:
        print("Envio cancelado pelo usuário. Nenhum e-mail foi enviado.")
        sys.exit(0)

    interval = prompt_interval("Intervalo entre envios (segundos) [0.75]: ", default=DEFAULT_INTERVAL)
    log_path = Path("logs") / "envios.csv"

    try:
        stats = send_all(
            host,
            port,
            smtp_user,
            smtp_password,
            from_address,
            records,
            subject_text,
            template_content,
            interval,
            log_path,
            allow_missing=allow_missing,
        )
    except PlaceholderRenderError as exc:
        placeholder_text = f" '{exc.placeholder}'" if exc.placeholder else ""
        print(
            f"Erro de placeholder durante o envio na linha {exc.line_number}{placeholder_text} do {exc.field}: {exc.original}",
            file=sys.stderr,
        )
        sys.exit(1)
    except RuntimeError as exc:
        print(str(exc), file=sys.stderr)
        sys.exit(1)

    print("\nResumo final:")
    print(f"  Enviados com sucesso: {stats['sucesso']}")
    print(f"  Falhas: {stats['falha']}")
    print(f"  Log disponível em: {format_path_for_display(stats['log_path'])}")

    exit_code = 0 if stats["falha"] == 0 else 1
    sys.exit(exit_code)


if __name__ == "__main__":
    try:
        main()
    except KeyboardInterrupt:
        print("\nOperação cancelada pelo usuário.")
        sys.exit(1)


# Instruções para empacotar o executável:
# Windows: pyinstaller --onefile --noconsole emaileria_wizard.py
# macOS: pyinstaller --onefile --windowed emaileria_wizard.py
