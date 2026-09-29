#!/usr/bin/env python3
"""Make a Registrar-of-Companies template freely editable.

  python3 prepare.py <template.doc|.docx> <out.docx>

1. Old binary Word (.doc) is converted to .docx through Word (or LibreOffice),
   so its XML can be edited. Legacy form fields (FORMTEXT / FORMCHECKBOX)
   survive the conversion intact.
2. The "restrict editing / filling in forms" protection is removed by deleting
   <w:documentProtection> (and <w:writeProtection>) from word/settings.xml.
   The protection is only an enforcement flag plus a password hash - removing the
   element unlocks the file without needing the password.
"""
import re
import shutil
import sys
import zipfile
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import wordauto  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")  # Hebrew output on Windows consoles

PROTECTION_RE = re.compile(rb"<w:(documentProtection|writeProtection)\b[^>]*/>")


def unprotect(docx: Path):
    tmp = docx.with_suffix(".tmp.docx")
    removed = 0
    with zipfile.ZipFile(docx) as zin, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            data = zin.read(item.filename)
            if item.filename == "word/settings.xml":
                data, removed = PROTECTION_RE.subn(b"", data)
            zout.writestr(item, data)
    tmp.replace(docx)
    return removed


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    src, dst = Path(sys.argv[1]).resolve(), Path(sys.argv[2]).resolve()
    dst.parent.mkdir(parents=True, exist_ok=True)
    if src.suffix.lower() == ".doc":
        wordauto.convert(src, dst, "docx")
        print(f"converted .doc -> .docx")
    else:
        shutil.copy2(src, dst)
    n = unprotect(dst)
    print(f"removed {n} protection element(s) -> {dst}")


if __name__ == "__main__":
    main()
