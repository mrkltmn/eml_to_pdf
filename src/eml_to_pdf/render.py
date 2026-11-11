"""Build Outlook-style print HTML for an e-mail and render HTML to PDF with Chromium."""

from __future__ import annotations

import base64
import html as htmlmod
import re
import subprocess
import sys
from datetime import datetime
from urllib.parse import unquote

from .model import Address, Email, _CID_RE, _norm_cid

# Windows already reports long names; macOS/Linux report abbreviations.
_TZ_LONG = {
    "EST": "Eastern Standard Time", "EDT": "Eastern Daylight Time",
    "CST": "Central Standard Time", "CDT": "Central Daylight Time",
    "MST": "Mountain Standard Time", "MDT": "Mountain Daylight Time",
    "PST": "Pacific Standard Time", "PDT": "Pacific Daylight Time",
    "AKST": "Alaska Standard Time", "AKDT": "Alaska Daylight Time",
    "HST": "Hawaii-Aleutian Standard Time", "UTC": "Coordinated Universal Time",
    "GMT": "Greenwich Mean Time",
}


def format_long_date(dt: datetime) -> str:
    """e.g. 'Tuesday, November 10, 2026 at 10:10:41 AM Central Standard Time' (local time)."""
    local = dt.astimezone()
    hour = local.hour % 12 or 12
    ampm = "AM" if local.hour < 12 else "PM"
    tz = local.tzname() or ""
    tz = _TZ_LONG.get(tz, tz)
    return (
        f"{local:%A}, {local:%B} {local.day}, {local.year} at "
        f"{hour}:{local:%M}:{local:%S} {ampm} {tz}"
    ).strip()


_CSS = """
#e2p-printed { font-family: Aptos, Calibri, 'Segoe UI', Helvetica, Arial, sans-serif !important;
  font-size: 9pt !important; font-weight: bold !important; text-align: right !important;
  color: #000 !important; border-bottom: 2px solid #555 !important;
  padding: 0 0 3px 0 !important; margin: 0 0 10px 0 !important; }
#e2p-header { border-collapse: collapse !important; margin: 0 0 14px 0 !important;
  width: 100% !important; border-bottom: 1px solid #ccc !important; }
#e2p-header th, #e2p-header td { font-family: Aptos, Calibri, 'Segoe UI', Helvetica, Arial, sans-serif !important;
  font-size: 8.5pt !important; color: #000 !important; vertical-align: top !important;
  text-align: left !important; padding: 1px 6px 2px 0 !important; border: none !important;
  background: none !important; line-height: 1.25 !important; }
#e2p-header th { font-weight: bold !important; white-space: nowrap !important; width: 1% !important; }
#e2p-header td { font-weight: normal !important; overflow-wrap: anywhere !important; }
.e2p-plain { white-space: pre-wrap; overflow-wrap: anywhere;
  font-family: Aptos, Calibri, 'Segoe UI', Helvetica, Arial, sans-serif; font-size: 11pt; }
img { max-width: 100%; height: auto; }
/* Outlook wraps bodies in named pages (div.WordSection1 {page: WordSection1}), which
   makes Chromium force a page break between our header and the body. */
* { page: auto !important; }
.e2p-ext-img { display: inline-block !important; max-width: 100% !important;
  border: 1px dashed #888 !important; background: #f3f3f3 !important; color: #444 !important;
  font: 7.5pt Aptos, Calibri, 'Segoe UI', Helvetica, Arial, sans-serif !important;
  padding: 3px 5px !important; margin: 2px 0 !important; overflow-wrap: anywhere !important;
  text-align: left !important; text-decoration: none !important; }
"""

# Remote 1x1 images are tracking pixels (e.g. Mimecast); drop them so they are never
# requested and do not trigger "remote content not loaded" warnings.
_IMG_TAG = re.compile(r"<img\b[^>]*>", re.IGNORECASE)
_DIM = re.compile(r"""\b(width|height)\s*=\s*["']?\s*(\d+)""", re.IGNORECASE)


def _strip_tracking_pixels(html: str) -> str:
    def repl(m: re.Match) -> str:
        tag = m.group(0)
        if not re.search(r"""src\s*=\s*["']?\s*https?:""", tag, re.IGNORECASE):
            return tag
        dims = {k.lower(): int(v) for k, v in _DIM.findall(tag)}
        if dims.get("width", 99) <= 1 and dims.get("height", 99) <= 1:
            return ""
        return tag

    return _IMG_TAG.sub(repl, html)


_SRC = re.compile(r"""\bsrc\s*=\s*(?:"([^"]*)"|'([^']*)'|([^\s>]+))""", re.IGNORECASE)
_PLACEHOLDER_URL_MAX = 150


def replace_remote_images(html: str, urls: list[str] | None = None) -> str:
    """Replace externally hosted <img> tags with a visible placeholder (they are never
    downloaded). Each replaced URL is appended to `urls` for the report."""

    def repl(m: re.Match) -> str:
        tag = m.group(0)
        src = _SRC.search(tag)
        url = htmlmod.unescape(next((g for g in src.groups() if g is not None), "")).strip() if src else ""
        if not re.match(r"(?i)(https?|ftp)://", url):
            return tag
        if urls is not None:
            urls.append(url)
        shown = url if len(url) <= _PLACEHOLDER_URL_MAX else url[:_PLACEHOLDER_URL_MAX] + "…"
        return (
            '<span class="e2p-ext-img">[External image not downloaded: '
            f"{htmlmod.escape(shown)}]</span>"
        )

    return _IMG_TAG.sub(repl, _strip_tracking_pixels(html))

_SLIP_CSS = """
body { font-family: Aptos, Calibri, 'Segoe UI', Helvetica, Arial, sans-serif; margin: 0; }
.wrap { margin-top: 3.2in; text-align: center; }
.label { font-size: 13pt; letter-spacing: .08em; text-transform: uppercase; color: #333; }
.name { font-size: 18pt; font-weight: bold; margin: 14px 0.5in; overflow-wrap: anywhere; }
.note { font-size: 11pt; color: #333; margin: 18px 0.75in 0; white-space: pre-wrap; }
"""


def _addresses(addrs: list[Address]) -> str:
    return ", ".join(a.full for a in addrs)


def _header_html(em: Email, printed_on: str) -> str:
    rows: list[tuple[str, str]] = [("Subject:", em.subject)]
    rows.append(("Date:", format_long_date(em.date) if em.date else ""))
    rows.append(("From:", em.sender.full if em.sender else ""))
    rows.append(("To:", _addresses(em.to)))
    if em.cc:
        rows.append(("CC:", _addresses(em.cc)))
    if em.bcc:
        rows.append(("BCC:", _addresses(em.bcc)))
    if em.attachments:
        rows.append(("Attachments:", ", ".join(a.filename for a in em.attachments)))
    body = "".join(
        f"<tr><th>{htmlmod.escape(k)}</th><td>{htmlmod.escape(v)}</td></tr>" for k, v in rows
    )
    return (
        f'<div id="e2p-printed">{htmlmod.escape(printed_on)}</div>'
        f'<table id="e2p-header">{body}</table>'
    )


def _inline_cid_map(em: Email) -> dict[str, str]:
    cids = {}
    for a in em.attachments:
        if a.content_id and a.data and em.is_inline(a):
            b64 = base64.b64encode(a.data).decode("ascii")
            cids[_norm_cid(a.content_id)] = f"data:{a.content_type};base64,{b64}"
    return cids


def build_email_html(em: Email, printed_on: str, remote_urls: list[str] | None = None) -> str:
    header = _header_html(em, printed_on)
    style = f"<style>{_CSS}</style>"
    if em.html:
        cids = _inline_cid_map(em)
        body = _CID_RE.sub(
            lambda m: cids.get(_norm_cid(unquote(m.group(1))), m.group(0)),
            replace_remote_images(em.html, remote_urls),
        )
        head_close = re.search(r"</head\s*>", body, re.IGNORECASE)
        if head_close:
            body = body[: head_close.start()] + style + body[head_close.start():]
        else:
            body = style + body
        body_open = re.search(r"<body\b[^>]*>", body, re.IGNORECASE)
        if body_open:
            return body[: body_open.end()] + header + body[body_open.end():]
        return header + body
    text = htmlmod.escape(em.text or "")
    return (
        f'<!doctype html><html><head><meta charset="utf-8">{style}</head>'
        f'<body>{header}<div class="e2p-plain">{text}</div></body></html>'
    )


def html_attachment(html: str, remote_urls: list[str] | None = None) -> str:
    """An .htm/.html attachment, with remote images replaced by placeholders."""
    html = replace_remote_images(html, remote_urls)
    style = f"<style>{_CSS}</style>"
    head_close = re.search(r"</head\s*>", html, re.IGNORECASE)
    if head_close:
        return html[: head_close.start()] + style + html[head_close.start():]
    return style + html


def text_document_html(text: str) -> str:
    return (
        f'<!doctype html><html><head><meta charset="utf-8"><style>{_CSS}</style></head>'
        f'<body><div class="e2p-plain" style="font-family: Consolas, Menlo, monospace; '
        f'font-size: 9.5pt">{htmlmod.escape(text)}</div></body></html>'
    )


def slip_sheet_html(label: str, filename: str, note: str = "") -> str:
    note_html = f'<div class="note">{htmlmod.escape(note)}</div>' if note else ""
    return (
        f'<!doctype html><html><head><meta charset="utf-8"><style>{_SLIP_CSS}</style></head>'
        f'<body><div class="wrap"><div class="label">{htmlmod.escape(label)}</div>'
        f'<div class="name">{htmlmod.escape(filename)}</div>{note_html}</div></body></html>'
    )


class Renderer:
    """Headless Chromium wrapper. JavaScript and all network access are disabled, so
    remote content (tracking pixels, external images) is never fetched."""

    def __init__(self) -> None:
        self._pw = None
        self._browser = None
        self._context = None
        self.blocked_requests = 0

    def __enter__(self) -> Renderer:
        from playwright.sync_api import sync_playwright

        self._pw = sync_playwright().start()
        self._browser = self._launch()
        self._context = self._browser.new_context(java_script_enabled=False, offline=True)
        self._context.route(re.compile(r"^(https?|ftp|wss?|file):"), self._block)
        return self

    def _block(self, route) -> None:
        self.blocked_requests += 1
        route.abort()

    def _launch(self):
        chromium = self._pw.chromium
        errors = []
        for kwargs in ({}, {"channel": "msedge"}, {"channel": "chrome"}):
            try:
                return chromium.launch(**kwargs)
            except Exception as exc:  # browser not installed in that location
                errors.append(str(exc).splitlines()[0])
        print("Installing headless Chromium for PDF rendering (one-time download)...", flush=True)
        subprocess.run([sys.executable, "-m", "playwright", "install", "chromium"], check=True)
        return chromium.launch()

    def __exit__(self, *exc) -> None:
        for closer in (
            lambda: self._context.close(),
            lambda: self._browser.close(),
            lambda: self._pw.stop(),
        ):
            try:
                closer()
            except Exception:
                pass

    def html_to_pdf(self, html: str) -> tuple[bytes, int]:
        """Render HTML to Letter-size PDF bytes; also returns the count of blocked
        remote resources referenced by the page."""
        before = self.blocked_requests
        page = self._context.new_page()
        try:
            page.set_content(html, wait_until="load", timeout=120_000)
            pdf = page.pdf(
                format="Letter",
                print_background=True,
                margin={"top": "0.5in", "bottom": "0.6in", "left": "0.6in", "right": "0.6in"},
            )
        finally:
            page.close()
        return pdf, self.blocked_requests - before
