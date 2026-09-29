#!/usr/bin/env python3
"""Make a Registrar-of-Companies template freely editable.

  python3 prepare.py <template.doc|.docx> <out.docx>

Follows the manual unlock procedure, automated:
  1. Word saves the file as "Word XML Document" (flat .xml, a single text file).
     This also takes care of old binary .doc templates.
  2. In that XML, the protection element's w:enforcement="1" is changed to "0".
  3. Word opens the fixed .xml and saves it as a normal .docx, now unlocked.
Legacy form fields (FORMTEXT / FORMCHECKBOX) survive both conversions.

The protection is only an enforcement flag plus a password hash, so turning the
flag off unlocks the document without knowing the password.

If only LibreOffice is available (it cannot write Word XML), the same flag is
switched off directly inside the .docx (word/settings.xml).
"""
import re
import shutil
import sys
import tempfile
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import wordauto  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")  # Hebrew output on Windows consoles

ENFORCEMENT_RE = re.compile(r'(<w:(?:documentProtection|writeProtection)\b[^>]*?w:enforcement=")(1|true|on)(")')


def disable_enforcement(xml: str):
    """Set w:enforcement="0" on the protection element(s). Returns (xml, count)."""
    return ENFORCEMENT_RE.subn(r"\g<1>0\g<3>", xml)


def unprotect(docx: Path):
    """Switch off enforcement inside a .docx directly (used without Word)."""
    tmp = docx.with_suffix(".tmp.docx")
    changed = 0
    with zipfile.ZipFile(docx) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/settings.xml":
                text, changed = disable_enforcement(data.decode("utf-8"))
                data = text.encode("utf-8")
            zout.writestr(item, data)
    tmp.replace(docx)
    return changed


def unlock_via_word_xml(src: Path, dst: Path):
    with tempfile.TemporaryDirectory() as td:
        flat = Path(td) / "doc.xml"
        wordauto.convert(src, flat, "xml")                          # 1. save as Word XML
        text, changed = disable_enforcement(flat.read_text(encoding="utf-8"))
        flat.write_text(text, encoding="utf-8")                     # 2. enforcement 1 -> 0
        wordauto.convert(flat, dst, "docx")                         # 3. reopen, save .docx
    return changed


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    src, dst = Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve()
    dst.parent.mkdir(parents=True, exist_ok=True)
    if wordauto.backend() in ("word-mac", "word-windows"):
        n = unlock_via_word_xml(src, dst)
        print(f"Word XML route: enforcement switched off on {n} element(s) -> {dst}")
    else:
        if src.suffix.lower() == ".doc":
            wordauto.convert(src, dst, "docx")
        else:
            shutil.copy2(src, dst)
        n = unprotect(dst)
        print(f"direct route: enforcement switched off on {n} element(s) -> {dst}")


if __name__ == "__main__":
    main()
