"""Convert Word / Excel / PowerPoint attachments to PDF using the installed MS Office apps.

* Windows: COM automation (pywin32), using private, invisible app instances.
* macOS:   AppleScript via osascript. Office for Mac is sandboxed, so files are staged
           inside each app's own container to avoid "Grant File Access" prompts.

A dummy password is always supplied when opening, so password-protected files fail
quickly instead of hanging on a password dialog.
"""

from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path

WORD_EXT = {".doc", ".docx", ".docm", ".dot", ".dotx", ".dotm", ".rtf", ".odt", ".wpd"}
EXCEL_EXT = {".xls", ".xlsx", ".xlsm", ".xlsb", ".xltx", ".csv", ".ods"}
POWERPOINT_EXT = {".ppt", ".pptx", ".pptm", ".pps", ".ppsx", ".pot", ".potx", ".odp"}
OFFICE_EXT = WORD_EXT | EXCEL_EXT | POWERPOINT_EXT

_DUMMY_PASSWORD = "e2p-no-password"
_TIMEOUT_SECONDS = 300


class OfficeError(Exception):
    pass


def _app_for(ext: str) -> str:
    if ext in WORD_EXT:
        return "word"
    if ext in EXCEL_EXT:
        return "excel"
    return "powerpoint"


# --------------------------------------------------------------------------- macOS

_MAC_BUNDLE = {
    "word": "com.microsoft.Word",
    "excel": "com.microsoft.Excel",
    "powerpoint": "com.microsoft.Powerpoint",
}

_MAC_SCRIPTS = {
    "word": f"""
on run argv
  tell application "Microsoft Word"
    open file name (item 1 of argv) password document "{_DUMMY_PASSWORD}" without add to recent files and confirm conversions
    set theDoc to active document
    save as theDoc file name (item 2 of argv) file format format PDF
    close theDoc saving no
  end tell
end run
""",
    "excel": f"""
on run argv
  tell application "Microsoft Excel"
    set display alerts to false
    open workbook workbook file name (item 1 of argv) update links do not update links read only true password "{_DUMMY_PASSWORD}" ignore read only recommended true add to mru false
    set wb to active workbook
    save workbook as wb filename (item 2 of argv) file format PDF file format
    close wb saving no
  end tell
end run
""",
    "powerpoint": """
on run argv
  tell application "Microsoft PowerPoint"
    open (POSIX file (item 1 of argv))
    set p to active presentation
    save p in (POSIX file (item 2 of argv)) as save as PDF
    close p
  end tell
end run
""",
}


_MAC_COLLECTION = {"word": "documents", "excel": "workbooks", "powerpoint": "presentations"}


def _mac_is_running(app: str) -> bool:
    proc = subprocess.run(
        ["osascript", "-e", f'application id "{_MAC_BUNDLE[app]}" is running'],
        capture_output=True, text=True, timeout=30,
    )
    return proc.stdout.strip() == "true"


def _mac_quit_if_idle(app: str) -> bool:
    """Quit the app unless it has documents open (e.g. the user opened one mid-run)."""
    script = f"""
tell application id "{_MAC_BUNDLE[app]}"
  if (count of {_MAC_COLLECTION[app]}) is 0 then
    quit saving no
    return "quit"
  end if
  return "left open"
end tell
"""
    try:
        proc = subprocess.run(
            ["osascript", "-"], input=script, capture_output=True, text=True, timeout=60
        )
    except subprocess.TimeoutExpired:
        return False
    return proc.stdout.strip() == "quit"


def _mac_stage_dir(app: str) -> Path:
    container = Path.home() / "Library" / "Containers" / _MAC_BUNDLE[app] / "Data"
    base = container / "tmp" / "eml_to_pdf" if container.is_dir() else Path(tempfile.gettempdir())
    d = base / uuid.uuid4().hex
    d.mkdir(parents=True, exist_ok=True)
    return d


def _convert_mac(app: str, data: bytes, ext: str) -> bytes:
    work = _mac_stage_dir(app)
    try:
        src, out = work / f"input{ext}", work / "output.pdf"
        src.write_bytes(data)
        try:
            proc = subprocess.run(
                ["osascript", "-", str(src), str(out)],
                input=_MAC_SCRIPTS[app],
                capture_output=True,
                text=True,
                timeout=_TIMEOUT_SECONDS,
            )
        except subprocess.TimeoutExpired:
            raise OfficeError(f"Microsoft {app.title()} timed out after {_TIMEOUT_SECONDS}s")
        if proc.returncode != 0 or not out.exists():
            msg = (proc.stderr or proc.stdout).strip().splitlines()
            raise OfficeError(
                f"Microsoft {app.title()} could not convert the file"
                + (f" ({msg[-1]})" if msg else "")
                + " - it may be password-protected or corrupt"
            )
        return out.read_bytes()
    finally:
        shutil.rmtree(work, ignore_errors=True)


# --------------------------------------------------------------------------- Windows


class _WindowsOffice:
    def __init__(self) -> None:
        import pythoncom  # noqa: F401  (pywin32)

        pythoncom.CoInitialize()
        self._apps: dict[str, object] = {}

    def _app(self, app: str):
        import win32com.client

        if app not in self._apps:
            progid = {"word": "Word.Application", "excel": "Excel.Application",
                      "powerpoint": "PowerPoint.Application"}[app]
            inst = win32com.client.DispatchEx(progid)
            if app == "word":
                inst.Visible = False
                inst.DisplayAlerts = 0  # wdAlertsNone
            elif app == "excel":
                inst.Visible = False
                inst.DisplayAlerts = False
                inst.AskToUpdateLinks = False
            self._apps[app] = inst
        return self._apps[app]

    def convert(self, app: str, data: bytes, ext: str) -> bytes:
        work = Path(tempfile.mkdtemp(prefix="eml_to_pdf_"))
        try:
            src, out = work / f"input{ext}", work / "output.pdf"
            src.write_bytes(data)
            inst = self._app(app)
            try:
                if app == "word":
                    doc = inst.Documents.Open(
                        str(src), ConfirmConversions=False, ReadOnly=True,
                        AddToRecentFiles=False, PasswordDocument=_DUMMY_PASSWORD,
                        Visible=False, OpenAndRepair=False, NoEncodingDialog=True,
                    )
                    try:
                        doc.ExportAsFixedFormat(str(out), 17)  # wdExportFormatPDF
                    finally:
                        doc.Close(False)
                elif app == "excel":
                    wb = inst.Workbooks.Open(
                        str(src), UpdateLinks=0, ReadOnly=True, Password=_DUMMY_PASSWORD,
                        IgnoreReadOnlyRecommended=True, Notify=False, AddToMru=False,
                    )
                    try:
                        wb.ExportAsFixedFormat(0, str(out))  # xlTypePDF
                    finally:
                        wb.Close(False)
                else:
                    try:  # "::password::" suffix makes protected decks fail instead of prompting
                        pres = inst.Presentations.Open(
                            f"{src}::{_DUMMY_PASSWORD}::", ReadOnly=True, Untitled=False, WithWindow=False
                        )
                    except Exception:
                        pres = inst.Presentations.Open(
                            str(src), ReadOnly=True, Untitled=False, WithWindow=False
                        )
                    try:
                        pres.SaveAs(str(out), 32)  # ppSaveAsPDF
                    finally:
                        pres.Close()
            except Exception as exc:
                raise OfficeError(
                    f"Microsoft {app.title()} could not convert the file ({exc}) "
                    "- it may be password-protected or corrupt"
                ) from exc
            if not out.exists():
                raise OfficeError(f"Microsoft {app.title()} produced no PDF")
            return out.read_bytes()
        finally:
            shutil.rmtree(work, ignore_errors=True)

    def close(self) -> None:
        for inst in self._apps.values():
            try:
                inst.Quit()
            except Exception:
                pass
        self._apps.clear()


# --------------------------------------------------------------------------- public


class OfficeConverter:
    """Office apps launched by this converter are quit by close(). On macOS, apps that
    were already running before first use are left alone so the user's session is not
    disturbed. On Windows, private instances are always used and always quit."""

    def __init__(self) -> None:
        self._win: _WindowsOffice | None = None
        self._mac_checked: set[str] = set()
        self._mac_launched: set[str] = set()

    def convert(self, data: bytes, filename: str) -> bytes:
        ext = Path(filename).suffix.lower()
        if ext not in OFFICE_EXT:
            raise OfficeError(f"'{ext}' is not an Office file type")
        app = _app_for(ext)
        if sys.platform == "darwin":
            if app not in self._mac_checked:
                self._mac_checked.add(app)
                try:
                    if not _mac_is_running(app):
                        self._mac_launched.add(app)
                except (OSError, subprocess.SubprocessError):
                    pass
            return _convert_mac(app, data, ext)
        if sys.platform == "win32":
            if self._win is None:
                try:
                    self._win = _WindowsOffice()
                except ImportError as exc:
                    raise OfficeError("pywin32 is not installed") from exc
            return self._win.convert(app, data, ext)
        raise OfficeError("Office conversion requires Windows or macOS with Microsoft Office")

    def close(self) -> None:
        if self._win is not None:
            self._win.close()
            self._win = None
        for app in sorted(self._mac_launched):
            if not _mac_quit_if_idle(app):
                print(f"Note: Microsoft {app.title()} was left open because it has open documents.")
        self._mac_launched.clear()
