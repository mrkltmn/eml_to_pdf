"""Assemble one PDF per e-mail: the e-mail, then a slip sheet + rendering per attachment."""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pymupdf

from .model import Attachment, Email, _decode_bytes, parse_eml_bytes, parse_msg_source
from .office import OFFICE_EXT, OfficeConverter
from .render import Renderer, build_email_html, slip_sheet_html, text_document_html

IMAGE_EXT = {".png", ".jpg", ".jpeg", ".jpe", ".gif", ".bmp", ".tif", ".tiff", ".jxr", ".webp"}
TEXT_EXT = {".txt", ".log", ".text"}
HTML_EXT = {".htm", ".html"}
MAX_NESTING = 10

LETTER = pymupdf.paper_rect("letter")
MARGIN = 36  # points


class AttachmentError(Exception):
    pass


@dataclass
class Native:
    """An attachment that could not be rendered; saved alongside the PDF."""

    filename: str
    data: bytes


@dataclass
class BuildResult:
    warnings: list[str] = field(default_factory=list)
    natives: list[Native] = field(default_factory=list)


class Assembler:
    def __init__(self, renderer: Renderer, office: OfficeConverter, printed_on: str) -> None:
        self.renderer = renderer
        self.office = office
        self.printed_on = printed_on

    # ----------------------------------------------------------------- helpers

    def _html_doc(self, html: str, result: BuildResult, context: str) -> pymupdf.Document:
        pdf, blocked = self.renderer.html_to_pdf(html)
        if blocked:
            result.warnings.append(
                f"{context}: {blocked} remote resource(s) (e.g. externally hosted images) "
                "were not downloaded"
            )
        return pymupdf.open("pdf", pdf)

    def _slip(self, label: str, filename: str, note: str = "") -> pymupdf.Document:
        pdf, _ = self.renderer.html_to_pdf(slip_sheet_html(label, filename, note))
        return pymupdf.open("pdf", pdf)

    # ----------------------------------------------------------------- e-mail

    def build_email(
        self, em: Email, result: BuildResult, context: str = "E-mail", depth: int = 0
    ) -> pymupdf.Document:
        doc = self._html_doc(build_email_html(em, self.printed_on), result, context)
        attachments = [a for a in em.attachments if not em.is_inline(a)]
        total = len(attachments)
        for i, att in enumerate(attachments, start=1):
            label = f"Attachment {i} of {total}"
            att_context = f"{context} > {att.filename}"
            try:
                att_doc = self._convert_attachment(att, result, att_context, depth)
                slip = self._slip(label, att.filename)
            except Exception as exc:
                reason = str(exc) if isinstance(exc, AttachmentError) else f"{type(exc).__name__}: {exc}"
                note = f"This attachment could not be rendered to PDF.\nReason: {reason}"
                if att.data:
                    note += "\nThe original file has been preserved in native format."
                    result.natives.append(Native(att.filename, att.data))
                result.warnings.append(f"{att_context}: {reason}")
                att_doc = None
                slip = self._slip(label, att.filename, note)
            doc.insert_pdf(slip)
            if att_doc is not None:
                doc.insert_pdf(att_doc)
        return doc

    # ----------------------------------------------------------------- attachments

    def _convert_attachment(
        self, att: Attachment, result: BuildResult, context: str, depth: int
    ) -> pymupdf.Document:
        ext = Path(att.filename).suffix.lower()
        ctype = (att.content_type or "").lower()

        if att.embedded is not None or ext in (".eml", ".msg") or ctype == "message/rfc822":
            if depth >= MAX_NESTING:
                raise AttachmentError("attached e-mails are nested too deeply")
            em = att.embedded
            if em is None:
                if not att.data:
                    raise AttachmentError("attached e-mail is empty")
                em = parse_msg_source(att.data) if ext == ".msg" else parse_eml_bytes(att.data)
            return self.build_email(em, result, context, depth + 1)

        if not att.data:
            raise AttachmentError("attachment content is missing or could not be read")

        if ext == ".pdf" or ctype == "application/pdf":
            return self._pdf(att.data)
        if ext in IMAGE_EXT or (ctype.startswith("image/") and not ext):
            return self._image(att.data, ext or "." + ctype.split("/", 1)[1])
        if ext in OFFICE_EXT:
            return pymupdf.open("pdf", self.office.convert(att.data, att.filename))
        if ext in TEXT_EXT or (ctype == "text/plain" and not ext):
            return self._html_doc(text_document_html(_decode_bytes(att.data)), result, context)
        if ext in HTML_EXT or (ctype == "text/html" and not ext):
            return self._html_doc(_decode_bytes(att.data), result, context)
        raise AttachmentError(f"file type '{ext or ctype}' cannot be rendered to PDF")

    @staticmethod
    def _pdf(data: bytes) -> pymupdf.Document:
        try:
            doc = pymupdf.open("pdf", data)
        except Exception as exc:
            raise AttachmentError(f"PDF is corrupt or unreadable ({exc})") from exc
        if doc.needs_pass and not doc.authenticate(""):
            raise AttachmentError("PDF is password-protected")
        if doc.page_count == 0:
            raise AttachmentError("PDF has no pages")
        return doc

    @staticmethod
    def _image(data: bytes, ext: str) -> pymupdf.Document:
        try:
            img = pymupdf.open(stream=data, filetype=ext.lstrip("."))
            src = pymupdf.open("pdf", img.convert_to_pdf())
        except Exception as exc:
            raise AttachmentError(f"image could not be read ({exc})") from exc
        out = pymupdf.open()
        for page in src:
            w, h = page.rect.width, page.rect.height
            landscape = w > h
            pw, ph = (LETTER.height, LETTER.width) if landscape else (LETTER.width, LETTER.height)
            new = out.new_page(width=pw, height=ph)
            # Fit within margins, but never enlarge small images beyond natural size.
            scale = min(1.0, (pw - 2 * MARGIN) / w, (ph - 2 * MARGIN) / h)
            tw, th = w * scale, h * scale
            x0 = (pw - tw) / 2
            new.show_pdf_page(pymupdf.Rect(x0, MARGIN, x0 + tw, MARGIN + th), src, page.number)
        return out


def add_page_numbers(doc: pymupdf.Document) -> None:
    """Stamp 'i of N' at the bottom right of every page (handles rotated pages)."""
    total = doc.page_count
    for i, page in enumerate(doc, start=1):
        text = f"{i} of {total}"
        size = 9
        width = pymupdf.get_text_length(text, fontname="helv", fontsize=size)
        r = page.rect  # visible (rotated) page rectangle
        point = pymupdf.Point(r.width - 36 - width, r.height - 18) * page.derotation_matrix
        page.insert_text(
            point, text, fontname="helv", fontsize=size, rotate=page.rotation, color=(0, 0, 0)
        )
