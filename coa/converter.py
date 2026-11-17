"""Step 4 (optional): Word -> PDF using LibreOffice, one call for the whole batch."""
import os
import shutil
import subprocess
import tempfile
from pathlib import Path

CANDIDATES = [
    "soffice", "libreoffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
    "/Applications/LibreOffice.app/Contents/MacOS/soffice",
]


def find_soffice():
    for c in CANDIDATES:
        path = shutil.which(c) or (c if os.path.isfile(c) else None)
        if path:
            return path
    return None


def pdf_available():
    return find_soffice() is not None


def docx_to_pdf(docx_paths, out_dir):
    """Convert many .docx files at once. Returns {docx_path: pdf_path or None}."""
    exe = find_soffice()
    if not exe:
        raise RuntimeError("PDF export needs LibreOffice (free) installed. Word output still works.")
    docx_paths = [str(p) for p in docx_paths]
    # A throw-away profile folder avoids 'user profile locked' errors.
    with tempfile.TemporaryDirectory() as profile:
        cmd = [exe, f"-env:UserInstallation={Path(profile).as_uri()}",
               "--headless", "--convert-to", "pdf", "--outdir", str(out_dir), *docx_paths]
        subprocess.run(cmd, check=True, capture_output=True, timeout=60 + 3 * len(docx_paths))
    result = {}
    for p in docx_paths:
        pdf = Path(out_dir) / (Path(p).stem + ".pdf")
        result[p] = str(pdf) if pdf.exists() else None
    return result
