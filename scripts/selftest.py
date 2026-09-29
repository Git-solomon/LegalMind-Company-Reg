#!/usr/bin/env python3
"""Check that this machine can run the whole skill (macOS or Windows).

  python3 selftest.py        (on Windows: py selftest.py)

Runs every stage on the bundled templates with dummy data and reports which
stage fails: Python deps, Word/LibreOffice, .doc conversion, unlocking, filling,
PDF export, and the company-name lookup. Output PDFs go to a temp folder whose
path is printed, so they can be opened and eyeballed.
"""
import json
import subprocess
import sys
import tempfile
import traceback
from pathlib import Path

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")  # Hebrew output on Windows consoles

HERE = Path(__file__).parent
TEMPLATES = HERE.parent / "templates"
sys.path.insert(0, str(HERE))

results = []


def stage(name, fn):
    try:
        detail = fn()
        results.append((name, True, detail or ""))
        print(f"[OK]   {name} {detail or ''}")
        return True
    except Exception as e:  # report and keep going so every stage is covered
        results.append((name, False, str(e)))
        print(f"[FAIL] {name}: {e}")
        traceback.print_exc(limit=1)
        return False


def main():
    out = Path(tempfile.mkdtemp(prefix="cid_selftest_"))
    print(f"platform: {sys.platform}, python {sys.version.split()[0]}, output: {out}\n")

    def deps():
        import lxml  # noqa: F401
        return "(lxml)"
    if not stage("python packages", deps):
        print("\ninstall with:  python -m pip install lxml")
        return 1

    import fields as F
    import prepare
    import wordauto

    def engine():
        b = wordauto.backend()
        if not b:
            raise RuntimeError("no Microsoft Word or LibreOffice found")
        return f"({b})"
    has_engine = stage("conversion engine", engine)

    form1 = out / "form1.docx"
    if has_engine:
        stage(".doc -> .docx conversion (Form 1)",
              lambda: wordauto.convert(TEMPLATES / "form1-company-registration.doc", form1, "docx") and None)

    def unlock():
        src = TEMPLATES / "director-declaration.docx"
        dst = out / "director.docx"
        dst.write_bytes(src.read_bytes())
        n = prepare.unprotect(dst)
        if n != 1:
            raise RuntimeError(f"expected 1 protection element, switched off {n}")
        return f"({len(F.collect_fields(F.load(dst)))} fields)"
    stage("unlock .docx (direct)", unlock)

    def unlock_word_xml():
        dst = out / "form1-unlocked.docx"
        n = prepare.unlock_via_word_xml(TEMPLATES / "form1-company-registration.doc", dst)
        with __import__("zipfile").ZipFile(dst) as z:
            settings = z.read("word/settings.xml").decode("utf-8")
        if n != 1 or 'w:enforcement="1"' in settings:
            raise RuntimeError("document is still protected after the Word XML route")
        return f"({len(F.collect_fields(F.load(dst)))} fields)"
    if has_engine and wordauto.backend() != "libreoffice":
        stage("unlock via Word XML (Form 1)", unlock_word_xml)

    def fill():
        vals = out / "values.json"
        vals.write_text(json.dumps({"fields": {"0": "בדיקה בע\"מ", "1": "כהן", "2": "דנה", "3": "012345678"}},
                                   ensure_ascii=False), encoding="utf-8")
        F.cmd_fill(str(out / "director.docx"), str(vals), str(out / "director-filled.docx"))
        F.cmd_flatten(str(out / "director-filled.docx"), str(out / "director-flat.docx"))
        got = F.text_value(F.collect_fields(F.load(out / "director-filled.docx"))[1])
        if got != "כהן":
            raise RuntimeError(f"field 1 reads back {got!r}")
    stage("fill + flatten", fill)

    if has_engine:
        stage("PDF export", lambda: wordauto.convert(out / "director-flat.docx",
                                                     out / "director.pdf", "pdf") and f"-> {out / 'director.pdf'}")

    def name_check():
        r = subprocess.run([sys.executable, str(HERE / "check_name.py"), "--json", "סנסו נובה בע\"מ"],
                           capture_output=True, check=True)
        data = json.loads(r.stdout.decode("utf-8"))
        if data[0]["verdict"] != "תפוס":
            raise RuntimeError(f"expected 'תפוס' for a known company, got {data[0]['verdict']!r}")
    stage("company name lookup (data.gov.il)", name_check)

    failed = [r for r in results if not r[1]]
    print(f"\n{len(results) - len(failed)}/{len(results)} stages passed. Files: {out}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())
