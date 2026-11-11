"""Command-line entry point: find e-mails, convert each to a PDF, write a report."""

from __future__ import annotations

import argparse
import hashlib
import os
import re
import sys
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path

from .assemble import Assembler, BuildResult, add_page_numbers
from .model import Email, parse_file
from .office import OfficeConverter
from .render import Renderer, format_long_date

SUPPORTED_EXT = {".eml", ".msg"}
# Other mail containers we recognise but do not process -- listed in the report.
OTHER_MAIL_EXT = {".emlx", ".mbox", ".mbx", ".pst", ".ost", ".olm", ".mht", ".mhtml", ".nsf"}

_ILLEGAL = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
_WINDOWS_RESERVED = {"CON", "PRN", "AUX", "NUL", *(f"COM{i}" for i in range(1, 10)),
                     *(f"LPT{i}" for i in range(1, 10))}


@dataclass
class Outcome:
    source: Path
    sha256: str = ""
    output: Path | None = None
    warnings: list[str] = field(default_factory=list)
    natives_dir: Path | None = None
    error: str = ""


# --------------------------------------------------------------------------- discovery


def discover(root: Path, output_dir: Path) -> tuple[list[Path], list[Path]]:
    emails, other = [], []
    for dirpath, dirnames, filenames in os.walk(root):
        current = Path(dirpath)
        dirnames[:] = sorted(
            d for d in dirnames
            if not d.startswith(".") and (current / d).resolve() != output_dir
        )
        for name in sorted(filenames):
            if name.startswith(("._", "~$", ".")):
                continue  # macOS metadata, Office lock files, hidden files
            ext = Path(name).suffix.lower()
            if ext in SUPPORTED_EXT:
                emails.append(current / name)
            elif ext in OTHER_MAIL_EXT:
                other.append(current / name)
    return emails, other


# --------------------------------------------------------------------------- naming


def _safe(part: str, limit: int = 60) -> str:
    part = _ILLEGAL.sub("", part).replace("'", "").strip()
    part = re.sub(r"\s+", " ", part)[:limit].strip(" .")
    return part


def pdf_name(em: Email) -> str:
    sender = _safe(em.sender.display) if em.sender else ""
    recipient = _safe(em.to[0].display) if em.to else ""
    when = em.date.astimezone().strftime("%Y-%m-%d_%H%M") if em.date else ""
    stem = f"{sender or 'Unknown Sender'}_{recipient or 'No Recipient'}_{when or 'No Date'}"
    if stem.upper() in _WINDOWS_RESERVED:
        stem += "_"
    return stem


def unique_path(directory: Path, stem: str, suffix: str, taken: set[str]) -> Path:
    candidate, n = directory / f"{stem}{suffix}", 1
    while str(candidate).lower() in taken:
        n += 1
        candidate = directory / f"{stem}_{n}{suffix}"
    taken.add(str(candidate).lower())
    return candidate


def save_natives(result: BuildResult, pdf_path: Path) -> Path | None:
    if not result.natives:
        return None
    folder = pdf_path.with_name(f"{pdf_path.stem}_natives")
    folder.mkdir(parents=True, exist_ok=True)
    taken: set[str] = set()
    for native in result.natives:
        name = _safe(native.filename, 150) or "attachment"
        stem, suffix = Path(name).stem, Path(name).suffix
        unique_path(folder, stem, suffix, taken).write_bytes(native.data)
    return folder


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with path.open("rb") as fh:
        for chunk in iter(lambda: fh.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


# --------------------------------------------------------------------------- report


def write_report(
    report_path: Path, root: Path, output_dir: Path, started: str,
    outcomes: list[Outcome], other: list[Path],
) -> None:
    ok = [o for o in outcomes if o.output and not o.warnings]
    warn = [o for o in outcomes if o.output and o.warnings]
    failed = [o for o in outcomes if not o.output]

    def rel(p: Path, base: Path) -> str:
        try:
            return str(p.relative_to(base))
        except ValueError:
            return str(p)

    lines = [
        "E-MAIL TO PDF CONVERSION REPORT",
        "=" * 72,
        f"Run started:   {started}",
        f"Source folder: {root}",
        f"Output folder: {output_dir}",
        "",
        f"E-mails found:                {len(outcomes)}",
        f"Converted:                    {len(ok)}",
        f"Converted with warnings:      {len(warn)}",
        f"Could NOT be converted:       {len(failed)}",
        f"Unsupported mail files found: {len(other)}",
        "",
    ]

    def section(title: str, items: list[Outcome], detail) -> None:
        lines.extend([title + f" ({len(items)})", "-" * 72])
        if not items:
            lines.append("  (none)")
        for o in items:
            lines.append(f"  {rel(o.source, root)}")
            lines.extend(detail(o))
        lines.append("")

    def converted_detail(o: Outcome) -> list[str]:
        out = [f"      -> {rel(o.output, output_dir)}", f"      SHA-256 (source): {o.sha256}"]
        if o.natives_dir:
            out.append(f"      Native files saved: {rel(o.natives_dir, output_dir)}")
        out.extend(f"      WARNING: {w}" for w in o.warnings)
        return out

    section("CONVERTED", ok, converted_detail)
    section("CONVERTED WITH WARNINGS", warn, converted_detail)
    section(
        "COULD NOT BE CONVERTED", failed,
        lambda o: [f"      SHA-256 (source): {o.sha256}", f"      ERROR: {o.error}"],
    )
    lines.extend([f"UNSUPPORTED MAIL FILES - NOT PROCESSED ({len(other)})", "-" * 72])
    lines.extend(f"  {rel(p, root)}" for p in other)
    if not other:
        lines.append("  (none)")
    report_path.write_text("\n".join(lines) + "\n", encoding="utf-8")


# --------------------------------------------------------------------------- main


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(
        prog="eml-to-pdf",
        description="Convert every .eml/.msg e-mail in the current folder (and subfolders) "
        "into a PDF that includes its attachments.",
    )
    parser.add_argument(
        "output", nargs="?", default="converted_pdf",
        help="name of the output folder (default: converted_pdf)",
    )
    args = parser.parse_args(argv)

    root = Path.cwd().resolve()
    output_dir = (root / args.output).resolve()
    now = datetime.now().astimezone()
    printed_on = format_long_date(now)

    emails, other = discover(root, output_dir)
    if not emails:
        print(f"No .eml or .msg files found in {root}")
        return 1
    output_dir.mkdir(parents=True, exist_ok=True)
    print(f"Found {len(emails)} e-mail(s). Writing PDFs to {output_dir}")

    outcomes: list[Outcome] = []
    taken: set[str] = set()
    office = OfficeConverter()
    try:
        with Renderer() as renderer:
            assembler = Assembler(renderer, office, printed_on)
            for index, path in enumerate(emails, start=1):
                rel = path.relative_to(root)
                outcome = Outcome(source=path)
                outcomes.append(outcome)
                print(f"[{index}/{len(emails)}] {rel} ... ", end="", flush=True)
                try:
                    outcome.sha256 = sha256(path)
                    em = parse_file(path)
                    result = BuildResult()
                    doc = assembler.build_email(em, result)
                    add_page_numbers(doc)
                    target_dir = output_dir / rel.parent
                    target_dir.mkdir(parents=True, exist_ok=True)
                    pdf_path = unique_path(target_dir, pdf_name(em), ".pdf", taken)
                    doc.save(pdf_path, garbage=3, deflate=True)
                    doc.close()
                    outcome.output = pdf_path
                    outcome.warnings = result.warnings
                    outcome.natives_dir = save_natives(result, pdf_path)
                    print("OK" if not result.warnings else f"OK ({len(result.warnings)} warning(s))")
                except KeyboardInterrupt:
                    outcome.error = "Interrupted by user"
                    raise
                except Exception as exc:
                    outcome.error = f"{type(exc).__name__}: {exc}"
                    print(f"FAILED - {outcome.error}")
                    if os.environ.get("EML_TO_PDF_DEBUG"):
                        traceback.print_exc()
    except KeyboardInterrupt:
        print("\nInterrupted - writing partial report.")
    finally:
        office.close()
        report = output_dir / f"conversion_report_{now:%Y-%m-%d_%H%M%S}.txt"
        write_report(report, root, output_dir, printed_on, outcomes, other)
        print(f"Report written to {report}")

    return 0 if all(o.output for o in outcomes) else 2


if __name__ == "__main__":
    sys.exit(main())
