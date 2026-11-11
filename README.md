# eml-to-pdf

Batch-convert `.eml` and `.msg` e-mails into PDFs for litigation. Each e-mail becomes one
PDF laid out like an Outlook printout, followed by each of its attachments.

## Install

Requires [uv](https://docs.astral.sh/uv/) and Microsoft Office (for Word/Excel/PowerPoint
attachments). Works on macOS and Windows.

```sh
uv tool install /path/to/eml_to_pdf        # installs `eml-to-pdf` into ~/.local/bin
uv tool install --reinstall /path/to/eml_to_pdf   # after updating the code
```

PDF rendering uses headless Chromium. If Playwright's Chromium, Microsoft Edge, or Google
Chrome is found it is used; otherwise Chromium is downloaded automatically on first run.

## Usage

Run from the folder containing the e-mails:

```sh
eml-to-pdf                 # writes to ./converted_pdf
eml-to-pdf "Production 001"
```

- Subfolders are searched; the output mirrors the source folder structure.
- PDFs are named `From_To_YYYY-MM-DD_HHMM.pdf` (sender name, first To recipient, local
  time). Duplicates get `_2`, `_3`, ...
- A `conversion_report_<timestamp>.txt` lists converted e-mails, e-mails converted with
  warnings, failures, and unsupported mail files (e.g. `.pst`, `.mbox`), with the SHA-256
  of each source file.

## What goes into each PDF

1. A "printed on" timestamp and header block (Subject, Date, From, To, CC, BCC,
   Attachments), then the message body with inline images (signature logos).
2. For each attachment, a slip sheet ("Attachment 1 of 3: filename") followed by:
   - PDFs: included as-is
   - Images: one page per image (multi-page TIFFs supported)
   - Word / Excel / PowerPoint / RTF / CSV: converted by the installed Office apps
   - Attached e-mails (`.eml`/`.msg`): rendered the same way, including their attachments
   - Text and HTML files: rendered as-is
3. "N of M" page numbers on every page.

Attachments that cannot be rendered (zip files, password-protected or corrupt files,
unknown types) get a slip sheet stating the reason. The original file is saved in a
`<pdf name>_natives` folder next to the PDF, and the report flags the e-mail as
"converted with warnings".

## Notes

- Remote content is never downloaded: JavaScript and network access are disabled during
  rendering, so tracking pixels are not triggered and the PDF reflects only what was in the
  e-mail file. Externally hosted images appear as a boxed placeholder,
  `[External image not downloaded: <URL>]`, and the report lists each full URL. 1x1
  tracking pixels are dropped silently.
- macOS: the first run will ask permission for your terminal to control Microsoft Word,
  Excel, and PowerPoint. Files are staged inside Office's sandbox folders to avoid
  per-file access prompts.
- Office apps started by the script are quit when the run finishes. On macOS, an app that
  was already open beforehand, or that has documents open at the end, is left running.
- Set `EML_TO_PDF_DEBUG=1` to print full tracebacks for failures.
