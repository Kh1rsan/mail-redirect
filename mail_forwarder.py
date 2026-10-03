"""Forward new Mail.ru inbox messages to another email address."""

from __future__ import annotations

import email
import email.policy
import html
from html.parser import HTMLParser
import imaplib
import json
import os
from pathlib import Path
import re
import shlex
import smtplib
import ssl
import sys
import time
from email.message import EmailMessage
from email.utils import parseaddr


def load_dotenv(path: Path | None = None) -> None:
    """Load KEY=VALUE pairs from the project .env without overriding existing variables."""
    dotenv_path = path or Path(__file__).resolve().with_name(".env")
    if not dotenv_path.exists():
        return

    for line_number, line in enumerate(dotenv_path.read_text(encoding="utf-8-sig").splitlines(), start=1):
        stripped = line.strip()
        if not stripped or stripped.startswith("#"):
            continue
        if stripped.startswith("export "):
            stripped = stripped[7:].lstrip()
        name, separator, raw_value = stripped.partition("=")
        name = name.strip()
        if not separator or not re.fullmatch(r"[A-Za-z_][A-Za-z0-9_]*", name):
            raise RuntimeError(f"Некорректная строка {line_number} в {dotenv_path}")
        try:
            parts = shlex.split(raw_value, comments=True, posix=True)
        except ValueError as exc:
            raise RuntimeError(f"Некорректное значение в строке {line_number} файла {dotenv_path}: {exc}") from exc
        os.environ.setdefault(name, " ".join(parts))


load_dotenv()

IMAP_HOST = os.getenv("MAIL_IMAP_HOST", "imap.mail.ru")
IMAP_PORT = int(os.getenv("MAIL_IMAP_PORT", "993"))
POLL_SECONDS = int(os.getenv("POLL_SECONDS", "60"))
STATE_FILE = Path(os.getenv("STATE_FILE", "mail_forwarder_state.json"))
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.mail.ru")
SMTP_PORT = int(os.getenv("SMTP_PORT", "465"))
SMTP_SECURITY = os.getenv("SMTP_SECURITY", "ssl").lower()
ALLOWED_SENDER = "noreply@rockstargames.com"
ALLOWED_SUBJECT = "Ваш проверочный код Rockstar Games"


class _HTMLText(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.parts: list[str] = []

    def handle_data(self, data: str) -> None:
        self.parts.append(data)

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        if tag in {"br", "p", "div", "li", "tr"}:
            self.parts.append("\n")


def required_env(name: str) -> str:
    value = os.getenv(name, "").strip()
    if not value:
        raise RuntimeError(f"Не задана обязательная переменная окружения {name}")
    return value


def load_state() -> dict[str, int]:
    try:
        data = json.loads(STATE_FILE.read_text(encoding="utf-8"))
        return {"uidvalidity": int(data["uidvalidity"]), "last_uid": int(data["last_uid"])}
    except FileNotFoundError:
        return {"uidvalidity": 0, "last_uid": 0}
    except (ValueError, KeyError, TypeError, json.JSONDecodeError) as exc:
        raise RuntimeError(f"Файл состояния {STATE_FILE} повреждён; сохраните его копию и удалите файл") from exc


def save_state(state: dict[str, int]) -> None:
    temporary = STATE_FILE.with_suffix(STATE_FILE.suffix + ".tmp")
    temporary.write_text(json.dumps(state), encoding="utf-8")
    temporary.replace(STATE_FILE)


def decode_header(value: str | None) -> str:
    if not value:
        return "(без темы)"
    return str(email.header.make_header(email.header.decode_header(value)))


def extract_body(message: email.message.Message) -> str:
    plain: list[str] = []
    html_parts: list[str] = []
    for part in message.walk():
        if part.is_multipart() or part.get_content_disposition() == "attachment":
            continue
        content_type = part.get_content_type()
        if content_type not in {"text/plain", "text/html"}:
            continue
        try:
            content = part.get_content()
        except (LookupError, UnicodeDecodeError, AttributeError):
            raw = part.get_payload(decode=True) or b""
            content = raw.decode(part.get_content_charset() or "utf-8", errors="replace")
        if not isinstance(content, str):
            continue
        if content_type == "text/plain":
            plain.append(content)
        else:
            html_parts.append(content)

    if plain:
        body = "\n\n".join(plain)
    elif html_parts:
        parser = _HTMLText()
        parser.feed("\n".join(html_parts))
        body = html.unescape("".join(parser.parts))
        body = re.sub(r"\n[ \t\r\n]+", "\n", body)
    else:
        body = "(В письме нет текстового содержимого.)"
    return body.strip() or "(Пустое письмо.)"


def send_email(
    smtp_user: str,
    smtp_password: str,
    sender: str,
    recipient: str,
    original_sender: str,
    subject: str,
    body: str,
) -> None:
    forwarded = EmailMessage()
    forwarded["From"] = sender
    forwarded["To"] = recipient
    forwarded["Subject"] = f"Переслано: {subject}"
    forwarded.set_content(f"От: {original_sender}\nТема: {subject}\n\n{body}")

    try:
        if SMTP_SECURITY == "starttls":
            with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as smtp:
                smtp.starttls(context=ssl.create_default_context())
                smtp.login(smtp_user, smtp_password)
                smtp.send_message(forwarded)
        elif SMTP_SECURITY == "ssl":
            with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT, timeout=30) as smtp:
                smtp.login(smtp_user, smtp_password)
                smtp.send_message(forwarded)
        else:
            raise RuntimeError("SMTP_SECURITY должен быть ssl или starttls")
    except (smtplib.SMTPException, OSError) as exc:
        raise RuntimeError(f"Не удалось отправить письмо через SMTP {SMTP_HOST}:{SMTP_PORT}: {exc}") from exc


def mailbox_uidvalidity(mailbox: imaplib.IMAP4_SSL) -> int:
    value = mailbox.response("UIDVALIDITY")[1]
    if not value or not value[0]:
        raise RuntimeError("Mail.ru не вернул UIDVALIDITY для папки INBOX")
    return int(value[0])


def highest_uid(mailbox: imaplib.IMAP4_SSL) -> int:
    status, data = mailbox.uid("search", None, "ALL")
    if status != "OK":
        raise RuntimeError("Не удалось прочитать список писем в INBOX")
    values = data[0].split() if data and data[0] else []
    return max((int(value) for value in values), default=0)


def fetch_message(mailbox: imaplib.IMAP4_SSL, uid: int) -> email.message.Message:
    status, data = mailbox.uid("fetch", str(uid), "(RFC822)")
    if status != "OK":
        raise RuntimeError(f"Не удалось скачать письмо UID {uid}")
    for item in data:
        if isinstance(item, tuple) and isinstance(item[1], bytes):
            return email.message_from_bytes(item[1], policy=email.policy.default)
    raise RuntimeError(f"Mail.ru вернул пустое письмо UID {uid}")


def forward_new_messages() -> None:
    email_address = required_env("MAIL_EMAIL")
    mail_password = required_env("MAIL_APP_PASSWORD")
    smtp_user = os.getenv("SMTP_USER", email_address).strip()
    smtp_password = os.getenv("SMTP_PASSWORD", mail_password)
    smtp_sender = os.getenv("SMTP_FROM", smtp_user).strip()
    recipient = required_env("FORWARD_TO")
    state = load_state()

    with imaplib.IMAP4_SSL(IMAP_HOST, IMAP_PORT) as mailbox:
        mailbox.login(email_address, mail_password)
        status, _ = mailbox.select("INBOX", readonly=True)
        if status != "OK":
            raise RuntimeError("Не удалось открыть папку INBOX")
        current_uidvalidity = mailbox_uidvalidity(mailbox)
        if state["uidvalidity"] != current_uidvalidity:
            # On first run (or after mailbox recreation), start with messages arriving from now on.
            state = {"uidvalidity": current_uidvalidity, "last_uid": highest_uid(mailbox)}
            save_state(state)
            print(f"Инициализация: существующие письма пропущены, последний UID {state['last_uid']}.")
            return

        status, data = mailbox.uid("search", None, "ALL")
        if status != "OK":
            raise RuntimeError("Не удалось найти новые письма")
        uids = sorted(
            uid
            for uid in (int(value) for value in (data[0].split() if data and data[0] else []))
            if uid > state["last_uid"]
        )
        for uid in uids:
            message = fetch_message(mailbox, uid)
            sender = decode_header(message.get("From"))
            subject = decode_header(message.get("Subject"))
            sender_address = parseaddr(sender)[1].strip().lower()
            if sender_address != ALLOWED_SENDER or subject != ALLOWED_SUBJECT:
                state["last_uid"] = uid
                save_state(state)
                print(f"Письмо UID {uid} пропущено: отправитель или тема не совпадают.")
                continue
            body = extract_body(message)
            send_email(smtp_user, smtp_password, smtp_sender, recipient, sender, subject, body)
            state["last_uid"] = uid
            save_state(state)
            print(f"Письмо UID {uid} переслано на {recipient} (тема: {subject}).")


def main() -> int:
    try:
        if POLL_SECONDS < 10:
            raise RuntimeError("POLL_SECONDS должен быть не меньше 10 секунд")
        if SMTP_SECURITY not in {"ssl", "starttls"}:
            raise RuntimeError("SMTP_SECURITY должен быть ssl или starttls")
        for name in ("MAIL_EMAIL", "MAIL_APP_PASSWORD", "FORWARD_TO"):
            required_env(name)
        print("Пересылка новых писем Mail.ru на электронную почту запущена. Остановить: Ctrl+C.")
        while True:
            try:
                forward_new_messages()
            except (imaplib.IMAP4.error, OSError, RuntimeError, ValueError) as exc:
                print(f"Ошибка: {exc}", file=sys.stderr)
            time.sleep(POLL_SECONDS)
    except KeyboardInterrupt:
        print("\nОстановлено.")
    except RuntimeError as exc:
        print(f"Ошибка конфигурации: {exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
