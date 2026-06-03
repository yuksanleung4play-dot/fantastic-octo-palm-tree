"""發件箱郵件來源。

支援兩種取得「已寄出郵件」的方式：

1. IMAP（線上）：連線到郵件伺服器，讀取「寄件備份 / Sent」資料夾。
2. 本地檔案（離線）：解析 ``.mbox`` 或 ``.eml`` 檔，方便沒有帳密、
   或想用匯出的郵件做測試時使用。

兩種方式都會回傳統一的 :class:`~weekly_report.models.EmailMessage` 清單。
"""

from __future__ import annotations

import email
import imaplib
import mailbox
import os
import re
from datetime import date, datetime, timezone
from email.header import decode_header, make_header
from email.message import Message
from email.utils import getaddresses, parsedate_to_datetime
from typing import Iterable, List, Optional

from .models import EmailMessage

# 常見的「寄件備份」資料夾名稱（不同郵件商命名不一）。
COMMON_SENT_FOLDERS = [
    "Sent",
    "Sent Items",
    "Sent Messages",
    "已发送",
    "已發送",
    "寄件備份",
    "已寄郵件",
    "[Gmail]/Sent Mail",
    "INBOX.Sent",
]


def _decode(value: Optional[str]) -> str:
    """安全地解碼 MIME 編碼過的標頭（中文／日文等）。"""

    if not value:
        return ""
    try:
        return str(make_header(decode_header(value))).strip()
    except Exception:
        return value.strip()


def _to_datetime(raw_date: Optional[str]) -> datetime:
    if not raw_date:
        return datetime.now()
    try:
        dt = parsedate_to_datetime(raw_date)
    except (TypeError, ValueError):
        return datetime.now()
    if dt is None:
        return datetime.now()
    # 轉成本地時間（去掉 tzinfo 方便和 date 比較）。
    if dt.tzinfo is not None:
        dt = dt.astimezone().replace(tzinfo=None)
    return dt


def _extract_snippet(msg: Message, max_len: int = 200) -> str:
    """擷取郵件純文字內文的開頭，作為摘要。"""

    text = ""
    if msg.is_multipart():
        for part in msg.walk():
            if part.get_content_type() == "text/plain" and not part.get_filename():
                text = _payload_to_text(part)
                if text:
                    break
    else:
        if msg.get_content_type() == "text/plain":
            text = _payload_to_text(msg)

    text = re.sub(r"\s+", " ", text).strip()
    if len(text) > max_len:
        text = text[:max_len].rstrip() + "…"
    return text


def _payload_to_text(part: Message) -> str:
    payload = part.get_payload(decode=True)
    if payload is None:
        return ""
    charset = part.get_content_charset() or "utf-8"
    try:
        return payload.decode(charset, errors="replace")
    except (LookupError, UnicodeDecodeError):
        return payload.decode("utf-8", errors="replace")


def _message_to_email(msg: Message) -> EmailMessage:
    subject = _decode(msg.get("Subject")) or "(無主旨)"
    sender = _decode(msg.get("From"))
    sent_at = _to_datetime(msg.get("Date"))

    to_values = msg.get_all("To", []) + msg.get_all("Cc", [])
    recipients = [
        _decode(name) and f"{_decode(name)} <{addr}>" or addr
        for name, addr in getaddresses(to_values)
        if addr
    ]

    return EmailMessage(
        subject=subject,
        sent_at=sent_at,
        recipients=recipients,
        sender=sender,
        snippet=_extract_snippet(msg),
    )


def _in_range(msg: EmailMessage, start: date, end: date) -> bool:
    return start <= msg.sent_date <= end


# --------------------------------------------------------------------------- #
# 本地檔案來源
# --------------------------------------------------------------------------- #
def load_from_file(
    path: str,
    period_start: date,
    period_end: date,
) -> List[EmailMessage]:
    """從本地的 ``.mbox`` 或 ``.eml`` 檔讀取已寄出郵件。

    若 ``path`` 是資料夾，則讀取其中所有 ``.eml`` 檔。
    """

    if not os.path.exists(path):
        raise FileNotFoundError(f"找不到郵件檔案或資料夾：{path}")

    messages: List[Message] = []

    if os.path.isdir(path):
        for name in sorted(os.listdir(path)):
            if name.lower().endswith(".eml"):
                with open(os.path.join(path, name), "rb") as fh:
                    messages.append(email.message_from_binary_file(fh))
    elif path.lower().endswith(".mbox"):
        box = mailbox.mbox(path)
        try:
            messages.extend(iter(box))
        finally:
            box.close()
    elif path.lower().endswith(".eml"):
        with open(path, "rb") as fh:
            messages.append(email.message_from_binary_file(fh))
    else:
        raise ValueError(
            f"不支援的郵件檔格式：{path}（僅支援 .mbox、.eml 或內含 .eml 的資料夾）"
        )

    emails = [_message_to_email(m) for m in messages]
    emails = [m for m in emails if _in_range(m, period_start, period_end)]
    emails.sort(key=lambda m: m.sent_at)
    return emails


# --------------------------------------------------------------------------- #
# IMAP 來源
# --------------------------------------------------------------------------- #
def load_from_imap(
    host: str,
    username: str,
    password: str,
    period_start: date,
    period_end: date,
    port: int = 993,
    use_ssl: bool = True,
    sent_folder: Optional[str] = None,
) -> List[EmailMessage]:
    """透過 IMAP 讀取「寄件備份」資料夾內某段期間的郵件。"""

    if use_ssl:
        client = imaplib.IMAP4_SSL(host, port)
    else:
        client = imaplib.IMAP4(host, port)

    try:
        client.login(username, password)
        folder = sent_folder or _detect_sent_folder(client)
        if folder is None:
            raise RuntimeError(
                "無法自動找到寄件備份資料夾，請用 --sent-folder 指定。"
            )

        status, _ = client.select(_quote_folder(folder), readonly=True)
        if status != "OK":
            raise RuntimeError(f"無法開啟資料夾：{folder}")

        # IMAP SEARCH 以日期粒度過濾，再於本地做精確範圍過濾。
        since = period_start.strftime("%d-%b-%Y")
        status, data = client.search(None, "SINCE", since)
        if status != "OK":
            return []

        ids = data[0].split()
        emails: List[EmailMessage] = []
        for msg_id in ids:
            status, msg_data = client.fetch(msg_id, "(RFC822)")
            if status != "OK" or not msg_data or not msg_data[0]:
                continue
            raw = msg_data[0][1]
            if not isinstance(raw, (bytes, bytearray)):
                continue
            parsed = email.message_from_bytes(raw)
            emails.append(_message_to_email(parsed))

        emails = [m for m in emails if _in_range(m, period_start, period_end)]
        emails.sort(key=lambda m: m.sent_at)
        return emails
    finally:
        try:
            client.logout()
        except Exception:
            pass


def _quote_folder(folder: str) -> str:
    if " " in folder or "/" in folder:
        return f'"{folder}"'
    return folder


def _detect_sent_folder(client: imaplib.IMAP4) -> Optional[str]:
    """嘗試從伺服器列出的資料夾中找到寄件備份。"""

    status, folders = client.list()
    if status != "OK" or not folders:
        # 退而求其次：直接嘗試常見名稱。
        for name in COMMON_SENT_FOLDERS:
            if client.select(_quote_folder(name), readonly=True)[0] == "OK":
                return name
        return None

    available = []
    for raw in folders:
        if not raw:
            continue
        line = raw.decode(errors="replace") if isinstance(raw, bytes) else str(raw)
        # 可能含 \Sent 這個特殊用途屬性，直接命中。
        if "\\Sent" in line:
            name = line.rsplit(" ", 1)[-1].strip().strip('"')
            return name
        available.append(line)

    for candidate in COMMON_SENT_FOLDERS:
        for line in available:
            if line.strip().strip('"').endswith(candidate):
                return candidate
    return None
