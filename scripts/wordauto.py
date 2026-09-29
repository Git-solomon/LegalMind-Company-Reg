"""Drive Microsoft Word for Mac (via AppleScript) to convert files.

Word for Mac is sandboxed: it can only reliably write inside its own container
(~/Library/Containers/com.microsoft.Word/Data/). So every conversion copies the
source into that container, has Word "save as" there, and copies the result back.

Falls back to LibreOffice (soffice) if Word is not installed.
"""
import os
import shutil
import subprocess
import tempfile
import uuid
from pathlib import Path

WORD_APP = "/Applications/Microsoft Word.app"
WORD_BOX = Path.home() / "Library/Containers/com.microsoft.Word/Data/Documents/_cid_tmp"

# AppleScript enum names for Word's "save as ... file format"
FORMATS = {"docx": "format document default", "pdf": "format PDF", "doc": "format document97"}


def _soffice():
    for p in ("soffice", "/Applications/LibreOffice.app/Contents/MacOS/soffice"):
        if shutil.which(p) or os.path.exists(p):
            return p
    return None


def _word_convert(src: Path, dst: Path, fmt: str):
    WORD_BOX.mkdir(parents=True, exist_ok=True)
    tag = uuid.uuid4().hex[:8]
    # ASCII-only names inside the container avoid AppleScript/HFS encoding issues
    box_src = WORD_BOX / f"in_{tag}{src.suffix}"
    box_dst = WORD_BOX / f"out_{tag}.{fmt}"
    shutil.copy2(src, box_src)
    script = f'''
with timeout of 180 seconds
tell application "Microsoft Word"
    set display alerts to alerts none
    open (POSIX file "{box_src}")
    delay 2
    set theDoc to document 1
    save as theDoc file name ((POSIX file "{box_dst}") as text) file format {FORMATS[fmt]}
    close every document saving no
end tell
end timeout
'''
    try:
        r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
        if not box_dst.exists():
            raise RuntimeError(f"Word conversion failed: {r.stderr.strip() or r.stdout.strip()}")
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(box_dst), dst)
    finally:
        for p in WORD_BOX.glob(f"*{tag}*"):
            p.unlink(missing_ok=True)
        for p in WORD_BOX.glob("~$*"):
            p.unlink(missing_ok=True)


def _soffice_convert(src: Path, dst: Path, fmt: str):
    exe = _soffice()
    with tempfile.TemporaryDirectory() as td:
        subprocess.run([exe, "--headless", "--convert-to", fmt, "--outdir", td, str(src)],
                       check=True, capture_output=True)
        out = next(Path(td).glob(f"*.{fmt}"))
        dst.parent.mkdir(parents=True, exist_ok=True)
        shutil.move(str(out), dst)


def convert(src, dst, fmt):
    """Convert src to dst in format fmt ('docx', 'pdf', 'doc')."""
    src, dst = Path(src).resolve(), Path(dst).resolve()
    if os.path.exists(WORD_APP):
        _word_convert(src, dst, fmt)
    elif _soffice():
        _soffice_convert(src, dst, fmt)
    else:
        raise RuntimeError("Neither Microsoft Word nor LibreOffice is available for conversion.")
    return dst
