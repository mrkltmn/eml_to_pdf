"""Parse .eml and .msg files into a common, format-independent Email model."""

from __future__ import annotations

import email
import email.policy
import email.utils
import mimetypes
import re
from dataclasses import dataclass, field
from datetime import datetime, timezone
from pathlib import Path
from urllib.parse import unquote

import extract_msg


@dataclass
class Address:
    name: str = ""
    email: str = ""

    @property
    def display(self) -> str:
        """Name only (falls back to address) -- used for filenames."""
        return self.name or self.email

    @property
    def full(self) -> str:
        """'Name <email>' -- used in the PDF header."""
        if self.name and self.email:
            return f"{self.name} <{self.email}>"
        return self.name or self.email


@dataclass
class Attachment:
    filename: str
    content_type: str = "application/octet-stream"
    data: bytes | None = None  # raw bytes (also used to save the native file)
    embedded: Email | None = None  # set when the attachment is itself an e-mail
    content_id: str | None = None


@dataclass
class Email:
    subject: str = ""
    sender: Address | None = None
    to: list[Address] = field(default_factory=list)
    cc: list[Address] = field(default_factory=list)
    bcc: list[Address] = field(default_factory=list)
    date: datetime | None = None
    html: str | None = None
    text: str | None = None
    attachments: list[Attachment] = field(default_factory=list)

    def referenced_cids(self) -> set[str]:
        if not self.html:
            return set()
        return {_norm_cid(unquote(m)) for m in _CID_RE.findall(self.html)}

    def is_inline(self, att: Attachment) -> bool:
        """True for images embedded in the body (e.g. signature logos)."""
        return bool(att.content_id) and _norm_cid(att.content_id) in self.referenced_cids()


_CID_RE = re.compile(r"""cid:([^"'\s)>]+)""", re.IGNORECASE)


def _norm_cid(cid: str) -> str:
    return cid.strip().strip("<>").strip().lower()


def _clean_name(name: str | None) -> str:
    name = (name or "").strip()
    while len(name) >= 2 and name[0] == name[-1] and name[0] in "'\"":
        name = name[1:-1].strip()
    return name


def _addr(name: str | None, addr: str | None) -> Address:
    name, addr = _clean_name(name), (addr or "").strip()
    if name.lower() == addr.lower():
        name = ""
    return Address(name=name, email=addr)


def _guess_filename(index: int, content_type: str) -> str:
    ext = mimetypes.guess_extension(content_type or "") or ".bin"
    return f"attachment_{index}{ext}"


def _decode_bytes(data: bytes, charset: str | None = None) -> str:
    for enc in filter(None, (charset, "utf-8", "cp1252")):
        try:
            return data.decode(enc)
        except (UnicodeDecodeError, LookupError):
            continue
    return data.decode("utf-8", errors="replace")


# --------------------------------------------------------------------------- EML


def parse_eml_bytes(data: bytes) -> Email:
    msg = email.message_from_bytes(data, policy=email.policy.default)
    return _from_email_message(msg)


def _header(msg: email.message.Message, name: str) -> str:
    try:
        value = msg.get(name)
        return str(value) if value is not None else ""
    except Exception:
        raw = msg.get_all(name, failobj=[]) if hasattr(msg, "get_all") else []
        return " ".join(str(r) for r in raw)


def _address_list(msg: email.message.Message, name: str) -> list[Address]:
    value = _header(msg, name)
    if not value:
        return []
    return [_addr(n, a) for n, a in email.utils.getaddresses([value]) if n or a]


def _part_text(part: email.message.Message) -> str:
    try:
        return part.get_content()
    except Exception:
        payload = part.get_payload(decode=True) or b""
        return _decode_bytes(payload, part.get_content_charset())


def _from_email_message(msg: email.message.EmailMessage) -> Email:
    em = Email(subject=_header(msg, "Subject"))
    senders = _address_list(msg, "From")
    em.sender = senders[0] if senders else None
    em.to = _address_list(msg, "To")
    em.cc = _address_list(msg, "Cc")
    em.bcc = _address_list(msg, "Bcc")

    date_str = _header(msg, "Date")
    if date_str:
        try:
            em.date = email.utils.parsedate_to_datetime(date_str)
            if em.date.tzinfo is None:
                em.date = em.date.replace(tzinfo=timezone.utc)
        except (TypeError, ValueError):
            em.date = None

    html_part = text_part = None
    try:
        html_part = msg.get_body(preferencelist=("html",))
        text_part = msg.get_body(preferencelist=("plain",))
    except Exception:
        pass
    if html_part is not None:
        em.html = _part_text(html_part)
    if text_part is not None:
        em.text = _part_text(text_part)
    body_ids = {id(p) for p in (html_part, text_part) if p is not None}

    def walk(part: email.message.Message, parent_type: str) -> None:
        ctype = part.get_content_type()
        if ctype == "message/rfc822":
            inner = part.get_payload(0) if part.is_multipart() else None
            if inner is None:
                return
            embedded = _from_email_message(inner)
            try:
                raw = inner.as_bytes()
            except Exception:
                raw = None
            name = part.get_filename() or f"{embedded.subject or 'attached message'}.eml"
            em.attachments.append(Attachment(name, ctype, raw, embedded, part.get("Content-ID")))
            return
        if part.is_multipart():
            for sub in part.iter_parts():
                walk(sub, ctype)
            return
        if id(part) in body_ids:
            return
        filename = part.get_filename()
        disposition = part.get_content_disposition()
        if (
            not filename
            and disposition != "attachment"
            and ctype in ("text/plain", "text/html")
            and parent_type == "multipart/alternative"
        ):
            return  # alternate representation of the body already shown
        data = part.get_payload(decode=True) or b""
        name = filename or _guess_filename(len(em.attachments) + 1, ctype)
        em.attachments.append(Attachment(name, ctype, data, None, part.get("Content-ID")))

    if msg.is_multipart():
        walk(msg, "")
    return em


# --------------------------------------------------------------------------- MSG


def parse_msg_source(source: str | Path | bytes) -> Email:
    msg = extract_msg.openMsg(source if isinstance(source, bytes) else str(source))
    try:
        return _from_msg(msg)
    finally:
        try:
            msg.close()
        except Exception:
            pass


def _msg_html(m) -> str | None:
    try:
        raw = m.htmlBody
    except Exception:
        return None
    if not raw:
        return None
    if isinstance(raw, str):
        return raw
    match = re.search(rb"charset\s*=\s*[\"']?([\w-]+)", raw[:4096], re.IGNORECASE)
    return _decode_bytes(raw, match.group(1).decode() if match else None)


def _from_msg(m) -> Email:
    em = Email(subject=getattr(m, "subject", None) or "")
    sender = getattr(m, "sender", None)
    if sender:
        name, addr = email.utils.parseaddr(sender)
        if "@" not in addr:
            name, addr = sender, ""
        em.sender = _addr(name, addr)

    for r in getattr(m, "recipients", None) or []:
        rtype = getattr(getattr(r, "type", None), "value", getattr(r, "type", None))
        target = {1: em.to, 2: em.cc, 3: em.bcc}.get(rtype)
        if target is not None:
            target.append(_addr(getattr(r, "name", ""), getattr(r, "email", "")))

    date = getattr(m, "date", None)
    if isinstance(date, datetime):
        em.date = date if date.tzinfo else date.replace(tzinfo=timezone.utc)

    em.html = _msg_html(m)
    try:
        em.text = getattr(m, "body", None)
    except Exception:
        em.text = None

    for i, a in enumerate(getattr(m, "attachments", None) or [], start=1):
        filename = (
            getattr(a, "longFilename", None)
            or getattr(a, "shortFilename", None)
            or getattr(a, "displayName", None)
        )
        cid = getattr(a, "cid", None) or getattr(a, "contentId", None)
        try:
            data = a.data
        except Exception:
            data = None
        if isinstance(data, (bytes, bytearray)):
            ctype = getattr(a, "mimetype", None) or mimetypes.guess_type(filename or "")[0]
            ctype = ctype or "application/octet-stream"
            em.attachments.append(
                Attachment(filename or _guess_filename(i, ctype), ctype, bytes(data), None, cid)
            )
        elif data is not None and hasattr(data, "attachments"):
            embedded = _from_msg(data)
            try:
                raw = data.exportBytes()
            except Exception:
                raw = None
            name = filename or f"{embedded.subject or 'attached message'}"
            if not name.lower().endswith(".msg"):
                name += ".msg"
            em.attachments.append(Attachment(name, "application/vnd.ms-outlook", raw, embedded, cid))
        else:
            em.attachments.append(
                Attachment(filename or f"attachment_{i}", "application/octet-stream", None, None, cid)
            )
    return em


def parse_file(path: Path) -> Email:
    suffix = path.suffix.lower()
    if suffix == ".eml":
        return parse_eml_bytes(path.read_bytes())
    if suffix == ".msg":
        return parse_msg_source(path)
    raise ValueError(f"Unsupported e-mail format: {suffix}")
