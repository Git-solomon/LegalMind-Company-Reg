"""Drive Microsoft Word to convert files (.doc -> .docx, .docx -> .pdf).

- macOS: AppleScript. Word for Mac is sandboxed and can only reliably write
  inside its own container (~/Library/Containers/com.microsoft.Word/Data/),
  so the source is copied there, converted, and the result copied back.
- Windows: Word's COM interface through PowerShell (no extra Python packages).
- Fallback on either: LibreOffice (soffice).

Conversions always run on ASCII-named temporary copies, so Hebrew output paths
never reach AppleScript / PowerShell / soffice.
"""
import os
import shutil
import subprocess
import sys
import tempfile
import uuid
from pathlib import Path

IS_MAC = sys.platform == "darwin"
IS_WIN = sys.platform == "win32"

MAC_WORD = "/Applications/Microsoft Word.app"
MAC_BOX = Path.home() / "Library/Containers/com.microsoft.Word/Data/Documents/_cid_tmp"
MAC_FORMATS = {"docx": "format document default", "pdf": "format PDF"}
# WdSaveFormat: wdFormatDocumentDefault = 16, wdFormatPDF = 17
WIN_FORMATS = {"docx": 16, "pdf": 17}


def _soffice():
    candidates = ["soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"]
    if IS_WIN:
        for base in (os.environ.get("ProgramFiles", r"C:\Program Files"),
                     os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")):
            candidates.append(os.path.join(base, "LibreOffice", "program", "soffice.exe"))
    for p in candidates:
        if shutil.which(p) or os.path.exists(p):
            return p
    return None


def _win_word_available():
    if not IS_WIN:
        return False
    r = subprocess.run(["powershell", "-NoProfile", "-Command",
                        "if (Test-Path 'Registry::HKEY_CLASSES_ROOT\\Word.Application') { 'yes' }"],
                       capture_output=True)
    return b"yes" in r.stdout


def _mac_word(src: Path, dst: Path, fmt: str):
    MAC_BOX.mkdir(parents=True, exist_ok=True)
    tag = uuid.uuid4().hex[:8]
    box_src = MAC_BOX / f"in_{tag}{src.suffix}"
    box_dst = MAC_BOX / f"out_{tag}.{fmt}"
    shutil.copy2(src, box_src)
    script = f'''
with timeout of 180 seconds
tell application "Microsoft Word"
    set display alerts to alerts none
    open (POSIX file "{box_src}")
    delay 2
    set theDoc to document 1
    save as theDoc file name ((POSIX file "{box_dst}") as text) file format {MAC_FORMATS[fmt]}
    close every document saving no
end tell
end timeout
'''
    try:
        r = subprocess.run(["osascript", "-e", script], capture_output=True)
        if not box_dst.exists():
            raise RuntimeError("Word conversion failed: "
                               + (r.stderr or r.stdout).decode("utf-8", "replace").strip())
        shutil.move(str(box_dst), dst)
    finally:
        for p in list(MAC_BOX.glob(f"*{tag}*")) + list(MAC_BOX.glob("~$*")):
            p.unlink(missing_ok=True)


def _win_word(src: Path, dst: Path, fmt: str):
    with tempfile.TemporaryDirectory() as td:
        tmp_src = Path(td) / f"in{src.suffix}"
        tmp_dst = Path(td) / f"out.{fmt}"
        shutil.copy2(src, tmp_src)
        src_ps, dst_ps = (str(x).replace("'", "''") for x in (tmp_src, tmp_dst))
        ps = f'''
$ErrorActionPreference = 'Stop'
$word = New-Object -ComObject Word.Application
try {{
    $word.Visible = $false
    $word.DisplayAlerts = 0
    # Open(FileName, ConfirmConversions, ReadOnly, AddToRecentFiles)
    $doc = $word.Documents.Open('{src_ps}', $false, $true, $false)
    $doc.SaveAs2('{dst_ps}', {WIN_FORMATS[fmt]})
    $doc.Close(0)
}} finally {{
    $word.Quit(0)
    [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($word)
}}
'''
        r = subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass",
                            "-Command", ps], capture_output=True, timeout=300)
        if not tmp_dst.exists():
            raise RuntimeError("Word conversion failed: "
                               + (r.stderr or r.stdout).decode("utf-8", "replace").strip())
        shutil.move(str(tmp_dst), dst)


def _soffice_convert(src: Path, dst: Path, fmt: str):
    exe = _soffice()
    with tempfile.TemporaryDirectory() as td:
        tmp_src = Path(td) / f"in{src.suffix}"
        shutil.copy2(src, tmp_src)
        subprocess.run([exe, "--headless", "--convert-to", fmt, "--outdir", td, str(tmp_src)],
                       check=True, capture_output=True)
        shutil.move(str(Path(td) / f"in.{fmt}"), dst)


def backend():
    """Name of the converter that will be used, or None."""
    if IS_MAC and os.path.exists(MAC_WORD):
        return "word-mac"
    if _win_word_available():
        return "word-windows"
    if _soffice():
        return "libreoffice"
    return None


def convert(src, dst, fmt):
    """Convert src to dst in format fmt ('docx' or 'pdf')."""
    src, dst = Path(src).resolve(), Path(dst).resolve()
    dst.parent.mkdir(parents=True, exist_ok=True)
    engine = backend()
    if engine == "word-mac":
        _mac_word(src, dst, fmt)
    elif engine == "word-windows":
        _win_word(src, dst, fmt)
    elif engine == "libreoffice":
        _soffice_convert(src, dst, fmt)
    else:
        raise RuntimeError("Neither Microsoft Word nor LibreOffice is available for conversion.")
    return dst
