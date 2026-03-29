#!/usr/bin/env python3
"""
Local conversion server for the PDF Tools Chrome extension.
Converts PDF → DOCX using whatever office engine is already on your machine —
no new downloads required.

Engine priority (first found wins):
  1. Microsoft Word  – Windows only, via PowerShell COM automation
  2. LibreOffice     – cross-platform, via soffice --headless

Requirements:
  - Python 3.8+  (python.org — free, lightweight)
  - One of the above office engines already installed

Usage:
  python local-server.py              # listens on http://127.0.0.1:8765
  python local-server.py --port 9000

In the extension:
  PDF to Word workspace → Server URL → http://localhost:8765
"""

import argparse
import glob
import os
import subprocess
import sys
import tempfile
from http.server import BaseHTTPRequestHandler, HTTPServer

# ---------------------------------------------------------------------------
# Engine detection
# ---------------------------------------------------------------------------

def _find_word() -> str | None:
    """Return path to WINWORD.EXE if Microsoft Word is installed, else None."""
    if sys.platform != "win32":
        return None
    candidates = [
        r"C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE",
        r"C:\Program Files\Microsoft Office\root\Office15\WINWORD.EXE",
        r"C:\Program Files\Microsoft Office\root\Office14\WINWORD.EXE",
        r"C:\Program Files\Microsoft Office\Office16\WINWORD.EXE",
        r"C:\Program Files\Microsoft Office\Office15\WINWORD.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office15\WINWORD.EXE",
        r"C:\Program Files (x86)\Microsoft Office\Office16\WINWORD.EXE",
    ]
    for path in candidates:
        if os.path.isfile(path):
            return path
    # Wildcard search for any Office version
    for pattern in [
        r"C:\Program Files\Microsoft Office\root\Office*\WINWORD.EXE",
        r"C:\Program Files (x86)\Microsoft Office\root\Office*\WINWORD.EXE",
        r"C:\Program Files\Microsoft Office\Office*\WINWORD.EXE",
    ]:
        matches = glob.glob(pattern)
        if matches:
            return matches[0]
    return None


def _find_soffice() -> str | None:
    """Return path to soffice if LibreOffice is installed, else None."""
    import shutil
    candidates = [
        "soffice",
        r"C:\Program Files\LibreOffice\program\soffice.exe",
        r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
        "/Applications/LibreOffice.app/Contents/MacOS/soffice",
        "/usr/bin/soffice",
        "/usr/lib/libreoffice/program/soffice",
    ]
    for path in candidates:
        if shutil.which(path) or os.path.isfile(path):
            return path
    return None


WORD_PATH = _find_word()
SOFFICE_PATH = _find_soffice()

# ---------------------------------------------------------------------------
# Conversion engines
# ---------------------------------------------------------------------------

def _convert_with_word(pdf_path: str, docx_path: str) -> tuple[bool, str]:
    """
    Drive Microsoft Word via PowerShell COM to open the PDF and save as DOCX.
    Returns (success, error_message).
    """
    # Escape backslashes and single quotes for PowerShell string literals
    def ps_escape(path: str) -> str:
        return path.replace("'", "''")

    script = f"""
$ErrorActionPreference = 'Stop'
try {{
    $word = New-Object -ComObject Word.Application
    $word.Visible = $false
    $word.DisplayAlerts = 0
    $doc = $word.Documents.Open('{ps_escape(pdf_path)}', $false, $true)
    $doc.SaveAs2('{ps_escape(docx_path)}', 16)
    $doc.Close([ref]$false)
    $word.Quit()
    [System.Runtime.InteropServices.Marshal]::ReleaseComObject($word) | Out-Null
    Write-Output 'OK'
}} catch {{
    Write-Error $_.Exception.Message
    exit 1
}}
"""
    try:
        result = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True,
            timeout=120,
        )
    except (subprocess.TimeoutExpired, FileNotFoundError) as e:
        return False, str(e)

    if result.returncode != 0:
        err = result.stderr.decode("utf-8", errors="replace").strip()
        return False, err or "PowerShell returned non-zero exit code"

    if not os.path.exists(docx_path):
        return False, "DOCX file was not created — Word may have encountered an error"

    return True, ""


def _convert_with_soffice(pdf_path: str, docx_path: str) -> tuple[bool, str]:
    """
    Run LibreOffice headless to convert the PDF to DOCX.
    Returns (success, error_message).
    """
    outdir = os.path.dirname(docx_path)
    try:
        result = subprocess.run(
            [
                SOFFICE_PATH,
                "--headless",
                "--convert-to", "docx:MS Word 2007 XML",
                "--outdir", outdir,
                pdf_path,
            ],
            capture_output=True,
            timeout=120,
        )
    except (subprocess.TimeoutExpired, FileNotFoundError) as e:
        return False, str(e)

    if result.returncode != 0 or not os.path.exists(docx_path):
        err = result.stderr.decode("utf-8", errors="replace").strip()
        return False, err or f"soffice exited with code {result.returncode}"

    return True, ""


def convert_pdf_to_docx(pdf_path: str, docx_path: str) -> tuple[bool, str]:
    """Try Word first, then LibreOffice. Returns (success, error_message)."""
    if WORD_PATH:
        ok, err = _convert_with_word(pdf_path, docx_path)
        if ok:
            return True, ""
        print(f"  [Word failed: {err}] — trying LibreOffice…", file=sys.stderr)

    if SOFFICE_PATH:
        return _convert_with_soffice(pdf_path, docx_path)

    return False, (
        "No conversion engine found.\n"
        "Install one of:\n"
        "  • Microsoft Word (already on most Windows machines)\n"
        "  • LibreOffice — https://www.libreoffice.org/download/  (free)"
    )


# ---------------------------------------------------------------------------
# Multipart parser  (stdlib only — no pip installs needed)
# ---------------------------------------------------------------------------

def _extract_file_field(content_type: str, body: bytes, field: str = "fileInput") -> bytes | None:
    boundary = None
    for part in content_type.split(";"):
        part = part.strip()
        if part.startswith("boundary="):
            boundary = part[len("boundary="):].strip('"')
            break
    if not boundary:
        return None

    sep = ("--" + boundary).encode()
    for segment in body.split(sep)[1:]:
        if segment in (b"--\r\n", b"--"):
            break
        if b"\r\n\r\n" not in segment:
            continue
        raw_headers, _, payload = segment.partition(b"\r\n\r\n")
        headers = raw_headers.decode("utf-8", errors="replace").lower()
        if f'name="{field.lower()}"' in headers or "filename=" in headers:
            return payload.rstrip(b"\r\n")
    return None


# ---------------------------------------------------------------------------
# HTTP handler
# ---------------------------------------------------------------------------

class _Handler(BaseHTTPRequestHandler):

    def do_OPTIONS(self):
        self._cors(200)
        self.end_headers()

    def do_POST(self):
        if self.path != "/api/v1/convert/pdf/word":
            self._error(404, "Not found")
            return

        if not WORD_PATH and not SOFFICE_PATH:
            self._error(500, (
                "No conversion engine found. Install Microsoft Word or LibreOffice."
            ))
            return

        content_type = self.headers.get("Content-Type", "")
        if "multipart/form-data" not in content_type:
            self._error(400, "Expected multipart/form-data")
            return

        length = int(self.headers.get("Content-Length", 0))
        body = self.rfile.read(length)

        pdf_bytes = _extract_file_field(content_type, body, "fileInput")
        if pdf_bytes is None:
            self._error(400, "Missing 'fileInput' field")
            return

        with tempfile.TemporaryDirectory() as tmpdir:
            pdf_path = os.path.join(tmpdir, "input.pdf")
            docx_path = os.path.join(tmpdir, "input.docx")

            with open(pdf_path, "wb") as fh:
                fh.write(pdf_bytes)

            ok, err = convert_pdf_to_docx(pdf_path, docx_path)
            if not ok:
                self._error(500, err)
                return

            with open(docx_path, "rb") as fh:
                docx_bytes = fh.read()

        self._cors(200)
        self.send_header(
            "Content-Type",
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        )
        self.send_header("Content-Disposition", 'attachment; filename="converted.docx"')
        self.send_header("Content-Length", str(len(docx_bytes)))
        self.end_headers()
        self.wfile.write(docx_bytes)

    def _cors(self, code: int):
        self.send_response(code)
        self.send_header("Access-Control-Allow-Origin", "*")
        self.send_header("Access-Control-Allow-Methods", "POST, OPTIONS")
        self.send_header("Access-Control-Allow-Headers", "Content-Type")

    def _error(self, code: int, message: str):
        body = message.encode()
        self._cors(code)
        self.send_header("Content-Type", "text/plain; charset=utf-8")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def log_message(self, fmt, *args):
        print(f"  {self.address_string()} - {fmt % args}", file=sys.stderr)


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Local PDF→DOCX server for the PDF Tools Chrome extension"
    )
    parser.add_argument("--port", type=int, default=8765)
    args = parser.parse_args()

    print("Engine detection:", file=sys.stderr)
    if WORD_PATH:
        print(f"  Microsoft Word : {WORD_PATH}", file=sys.stderr)
    else:
        print("  Microsoft Word : not found", file=sys.stderr)

    if SOFFICE_PATH:
        print(f"  LibreOffice    : {SOFFICE_PATH}", file=sys.stderr)
    else:
        print("  LibreOffice    : not found", file=sys.stderr)

    if not WORD_PATH and not SOFFICE_PATH:
        print(
            "\nERROR: No conversion engine found.\n"
            "Install Microsoft Word or LibreOffice (free) to use this server.\n"
            "LibreOffice: https://www.libreoffice.org/download/",
            file=sys.stderr,
        )
        sys.exit(1)

    engine = "Microsoft Word" if WORD_PATH else "LibreOffice"
    print(f"\nActive engine  : {engine}", file=sys.stderr)
    print(f"Listening      : http://127.0.0.1:{args.port}", file=sys.stderr)
    print(f"Extension      : PDF to Word → Server URL → http://localhost:{args.port}", file=sys.stderr)
    print("Ctrl+C to stop.\n", file=sys.stderr)

    server = HTTPServer(("127.0.0.1", args.port), _Handler)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print("\nServer stopped.", file=sys.stderr)


if __name__ == "__main__":
    main()
