#!/usr/bin/env python3
"""Build the whole incorporation package from one plan file.

  python3 build.py <plan.json> <output_dir>

plan.json:
{
  "documents": [
    {
      "template": "director-declaration",          # name in templates/ (no extension)
      "output": "02 - הצהרת דירקטור - ישראל ישראלי",   # output base name
      "fields": {"0": "...", "1": "..."},             # same format as fields.py fill
      "replace": [], "cells": []                      # optional
    }
  ]
}

For each document: prepare (unlock / convert .doc) -> fill -> flatten -> PDF.
Writes <output_dir>/<output>.pdf and keeps the filled .docx (fields still live,
unlocked) in <output_dir>/docx/ for later touch-ups. Prints every text field left empty, so you can confirm
each one is intentionally blank (signature/registrar-only/not applicable).
"""
import json
import re
import subprocess
import sys
import tempfile
from pathlib import Path

HERE = Path(__file__).parent
TEMPLATES = HERE.parent / "templates"
sys.path.insert(0, str(HERE))
import fields as F  # noqa: E402
import wordauto  # noqa: E402

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")  # Hebrew output on Windows consoles


def safe_filename(name):
    """Windows forbids <>:"/\\|?* in file names; בע"מ becomes בע״מ (Hebrew gershayim)."""
    name = name.replace('"', "\u05f4").replace("'", "\u05f3")
    return re.sub(r'[<>:/\\|?*]', "-", name).strip(" .")


def find_template(name):
    for ext in (".docx", ".doc"):
        p = TEMPLATES / f"{name}{ext}"
        if p.exists():
            return p
    sys.exit(f"template not found: {name}")


def main():
    if len(sys.argv) != 3:
        sys.exit(__doc__)
    plan = json.loads(Path(sys.argv[1]).read_text(encoding="utf-8"))
    out = Path(sys.argv[2]).resolve()
    (out / "docx").mkdir(parents=True, exist_ok=True)
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        for i, doc in enumerate(plan["documents"]):
            src = find_template(doc["template"])
            prepared = td / f"prep_{i}.docx"
            subprocess.run([sys.executable, str(HERE / "prepare.py"), str(src), str(prepared)], check=True)
            values = td / f"values_{i}.json"
            values.write_text(json.dumps({k: doc.get(k, []) if k != "fields" else doc.get(k, {})
                                          for k in ("fields", "replace", "cells")}, ensure_ascii=False),
                              encoding="utf-8")
            base = safe_filename(doc["output"])
            filled = out / "docx" / f"{base}.docx"
            F.cmd_fill(str(prepared), str(values), str(filled))
            flat = td / f"flat_{i}.docx"
            F.cmd_flatten(str(filled), str(flat))
            pdf = out / f"{base}.pdf"
            wordauto.convert(flat, pdf, "pdf")
            print(f"PDF -> {pdf}")
            root = F.load(filled)
            empty = [n for n, f in enumerate(F.collect_fields(root))
                     if F.ftype(f["ffData"]) == "text" and not F.text_value(f).strip("\u2002 ")]
            if empty:
                print(f"   empty text fields in {doc['output']}: {empty}")


if __name__ == "__main__":
    main()
