#!/usr/bin/env python3
"""List and fill legacy Word form fields (FORMTEXT / FORMCHECKBOX) in a .docx.

  python3 fields.py list <doc.docx> [--json out.json]
      Prints every form field with a stable index, its type, current value and
      the surrounding label text so you can tell what each field is for.

  python3 fields.py fill <doc.docx> <values.json> <out.docx>
      values.json:
        {
          "fields": {"0": "חברה בע\"מ", "7": true, "12": "שורה 1\\nשורה 2"},
          "replace": [{"find": "טקסט קיים", "with": "טקסט חדש"}],  # optional
          "cells": [{"table_caption": "DHR00", "row": 0, "col": 0, "text": "..."}]  # optional
        }
      Text fields take a string (\\n = line break). Checkboxes take true/false.

  python3 fields.py flatten <doc.docx> <out.docx>
      Turns every form field into plain text / a box glyph (used before PDF export).

  python3 fields.py probe <doc.docx> <out.docx>
      Writes each text field's own index (e.g. "#12") into it and ticks every
      checkbox, so a PDF render shows visually which index sits where.
      "replace" swaps literal text inside a single run - for text that is not a
      form field. "cells" writes text into an empty table cell that has no form
      field (e.g. the company-name box of the shareholder declaration); the
      table is chosen by its caption (tblCaption) or by "table" = body order index.
      Fields not listed are left as they are (blank in a template).

Run `prepare.py` first so the document is a .docx without edit protection.
"""
import copy
import json
import sys
import zipfile
from pathlib import Path

from lxml import etree

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8")  # Hebrew output on Windows consoles

W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W}
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


def q(tag):
    return f"{{{W}}}{tag}"


def load(path):
    with zipfile.ZipFile(path) as z:
        return etree.fromstring(z.read("word/document.xml"))


def save(src, dst, root):
    data = etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
    with zipfile.ZipFile(src) as zin, zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for item in zin.infolist():
            zout.writestr(item, data if item.filename == "word/document.xml" else zin.read(item.filename))


def collect_fields(root):
    """Walk runs in document order and return form fields with their result runs."""
    fields, stack = [], []
    for run in root.iter(q("r")):
        fc = run.find("w:fldChar", NS)
        if fc is not None:
            kind = fc.get(q("fldCharType"))
            if kind == "begin":
                ff = fc.find("w:ffData", NS)
                entry = None
                if ff is not None:
                    entry = {"ffData": ff, "begin": run, "separate": None, "results": [], "end": None}
                    fields.append(entry)
                stack.append({"entry": entry, "in_result": False})
            elif kind == "separate" and stack:
                stack[-1]["in_result"] = True
                if stack[-1]["entry"]:
                    stack[-1]["entry"]["separate"] = run
            elif kind == "end" and stack:
                top = stack.pop()
                if top["entry"]:
                    top["entry"]["end"] = run
            continue
        if stack and stack[-1]["entry"] and stack[-1]["in_result"] and run.find("w:t", NS) is not None:
            stack[-1]["entry"]["results"].append(run)
    return fields


def ftype(ff):
    if ff.find("w:checkBox", NS) is not None:
        return "checkbox"
    if ff.find("w:ddList", NS) is not None:
        return "dropdown"
    return "text"


def text_value(f):
    return "".join(t.text or "" for r in f["results"] for t in r.iter(q("t")))


def checkbox_value(ff):
    cb = ff.find("w:checkBox", NS)
    el = cb.find("w:checked", NS)
    if el is None:
        el = cb.find("w:default", NS)
    return el is not None and el.get(q("val"), "1") not in ("0", "false")


def plain_text(el, skip_runs):
    """Text of an element, skipping field result runs (so labels read cleanly)."""
    return "".join(t.text or "" for r in el.iter(q("r")) if r not in skip_runs for t in r.iter(q("t"))).strip()


def grid_col(tc):
    col = 0
    for prev in tc.itersiblings(preceding=True):
        if prev.tag == q("tc"):
            span = prev.find("w:tcPr/w:gridSpan", NS)
            col += int(span.get(q("val"))) if span is not None else 1
    return col


def cell_at(tr, col):
    c = 0
    for tc in tr.findall("w:tc", NS):
        span = tc.find("w:tcPr/w:gridSpan", NS)
        width = int(span.get(q("val"))) if span is not None else 1
        if c <= col < c + width:
            return tc
        c += width
    return None


def context(f, skip_runs):
    ctx = {}
    p = f["begin"].getparent()
    while p is not None and p.tag != q("p"):
        p = p.getparent()
    # text before the field in the same paragraph
    before = []
    for r in p.iter(q("r")):
        if r is f["begin"]:
            break
        if r not in skip_runs:
            before.extend(t.text or "" for t in r.iter(q("t")))
    ctx["before"] = "".join(before).strip()[-60:]
    if not ctx["before"]:
        prev = p.getprevious()
        while prev is not None and prev.tag == q("p") and not plain_text(prev, skip_runs):
            prev = prev.getprevious()
        if prev is not None and prev.tag == q("p"):
            ctx["prev_paragraph"] = plain_text(prev, skip_runs)[-60:]
    tc = p.getparent() if p.getparent().tag == q("tc") else None
    if tc is not None:
        own = plain_text(tc, skip_runs)
        if own:
            ctx["cell"] = own[:60]
        tr = tc.getparent()
        col = grid_col(tc)
        # nearest row above that has a label in the same grid column
        for prev_tr in tr.itersiblings(q("tr"), preceding=True):
            cell = cell_at(prev_tr, col)
            if cell is not None and cell.find(".//w:ffData", NS) is None:
                label = plain_text(cell, skip_runs)
                if label:
                    ctx["column_label"] = label[:60]
                    break
        # some forms put the label under the field instead
        nxt = tr.getnext()
        if nxt is not None and nxt.tag == q("tr"):
            cell = cell_at(nxt, col)
            if cell is not None and cell.find(".//w:ffData", NS) is None:
                label = plain_text(cell, skip_runs)
                if label:
                    ctx["label_below"] = label[:60]
    return ctx


def cmd_list(path, json_out=None):
    root = load(path)
    fields = collect_fields(root)
    skip = {r for f in fields for r in f["results"]}
    out = []
    for i, f in enumerate(fields):
        ff = f["ffData"]
        name = ff.find("w:name", NS)
        t = ftype(ff)
        item = {"index": i, "type": t, "name": name.get(q("val")) if name is not None else ""}
        item["value"] = checkbox_value(ff) if t == "checkbox" else text_value(f).strip("  ")
        ml = ff.find("w:textInput/w:maxLength", NS)
        if ml is not None:
            item["maxLength"] = int(ml.get(q("val")))
        item.update(context(f, skip))
        out.append(item)
    if json_out:
        Path(json_out).write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    for it in out:
        extras = " | ".join(f"{k}: {it[k]}" for k in ("column_label", "label_below", "cell", "before", "prev_paragraph") if it.get(k))
        print(f"[{it['index']:>3}] {it['type']:<8} {it['name']:<12} val={it['value']!r:<8} {extras}")


def set_text(f, value):
    lines = str(value).split("\n")
    if f["results"]:
        target, others = f["results"][0], f["results"][1:]
    else:
        # empty result: create a run right after "separate", borrowing its formatting
        target = copy.deepcopy(f["separate"])
        for child in list(target):
            if child.tag != q("rPr"):
                target.remove(child)
        f["separate"].addnext(target)
        others = []
    for r in others:
        r.getparent().remove(r)
    for child in list(target):
        if child.tag != q("rPr"):
            target.remove(child)
    for i, line in enumerate(lines):
        if i:
            etree.SubElement(target, q("br"))
        t = etree.SubElement(target, q("t"))
        t.text = line
        t.set(XML_SPACE, "preserve")


def set_checkbox(ff, checked):
    cb = ff.find("w:checkBox", NS)
    for el in cb.findall("w:checked", NS):
        cb.remove(el)
    el = etree.SubElement(cb, q("checked"))
    el.set(q("val"), "1" if checked else "0")


def fill_cell(root, spec):
    tables = list(root.iter(q("tbl")))
    if "table_caption" in spec:
        tables = [t for t in tables
                  if (c := t.find("w:tblPr/w:tblCaption", NS)) is not None and c.get(q("val")) == spec["table_caption"]]
        if not tables:
            sys.exit(f"no table with caption {spec['table_caption']!r}")
        tbl = tables[0]
    else:
        tbl = tables[spec["table"]]
    tr = tbl.findall("w:tr", NS)[spec.get("row", 0)]
    tc = tr.findall("w:tc", NS)[spec.get("col", 0)]
    p = tc.find("w:p", NS)
    for r in p.findall("w:r", NS):
        p.remove(r)
    r = etree.SubElement(p, q("r"))
    ppr_rpr = p.find("w:pPr/w:rPr", NS)
    rpr = copy.deepcopy(ppr_rpr) if ppr_rpr is not None else etree.Element(q("rPr"))
    rpr.tag = q("rPr")
    if rpr.find("w:rtl", NS) is None:
        etree.SubElement(rpr, q("rtl"))
    r.append(rpr)
    for i, line in enumerate(str(spec["text"]).split("\n")):
        if i:
            etree.SubElement(r, q("br"))
        t = etree.SubElement(r, q("t"))
        t.text = line
        t.set(XML_SPACE, "preserve")


def cmd_fill(path, values_path, out_path):
    root = load(path)
    fields = collect_fields(root)
    spec = json.loads(Path(values_path).read_text(encoding="utf-8"))
    for key, value in spec.get("fields", {}).items():
        i = int(key)
        if i >= len(fields):
            sys.exit(f"field index {i} does not exist (document has {len(fields)} fields)")
        f = fields[i]
        t = ftype(f["ffData"])
        if t == "checkbox":
            set_checkbox(f["ffData"], bool(value))
        else:
            ml = f["ffData"].find("w:textInput/w:maxLength", NS)
            if ml is not None and len(str(value)) > int(ml.get(q("val"))):
                # Word ignores maxLength for programmatic text, but flag it for review
                print(f"note: field {i} value longer than maxLength {ml.get(q('val'))}")
            set_text(f, value)
    for rep in spec.get("replace", []):
        hits = 0
        for t in root.iter(q("t")):
            if t.text and rep["find"] in t.text:
                t.text = t.text.replace(rep["find"], rep["with"])
                t.set(XML_SPACE, "preserve")
                hits += 1
        if not hits:
            print(f"warning: replace text not found in a single run: {rep['find']!r}")
    for cell in spec.get("cells", []):
        fill_cell(root, cell)
    save(path, out_path, root)
    print(f"filled {len(spec.get('fields', {}))} field(s) -> {out_path}")


def cmd_probe(path, out_path):
    root = load(path)
    fields = collect_fields(root)
    for i, f in enumerate(fields):
        if ftype(f["ffData"]) == "checkbox":
            set_checkbox(f["ffData"], True)
        else:
            set_text(f, f"#{i}")
    save(path, out_path, root)
    print(f"probed {len(fields)} field(s) -> {out_path}")


def flatten(root):
    """Replace every form field with its plain result, for the final PDF.

    Word draws grey brackets/shading around legacy form fields in PDF export.
    Text fields keep only their result runs (the 5-space placeholder of an empty
    field is removed so blanks stay clean). Checkboxes become a Wingdings box
    glyph (checked / empty) in the same formatting.
    """
    fields = collect_fields(root)
    for f in fields:
        begin, end = f["begin"], f["end"]
        if end is None or begin.getparent() is not end.getparent():
            continue  # unusual field spanning paragraphs - leave it as is
        ff = f["ffData"]
        parent = begin.getparent()
        kids = list(parent)
        span = kids[kids.index(begin): kids.index(end) + 1]
        if ftype(ff) == "checkbox":
            box = copy.deepcopy(begin)
            for child in list(box):
                if child.tag != q("rPr"):
                    box.remove(child)
            sym = etree.SubElement(box, q("sym"))
            sym.set(q("font"), "Wingdings")
            sym.set(q("char"), "F0FD" if checkbox_value(ff) else "F0A8")
            begin.addprevious(box)
            keep = set()
        else:
            if not text_value(f).strip("\u2002 "):
                for r in f["results"]:
                    for t in r.findall("w:t", NS):
                        t.text = ""
            keep = set(f["results"])
        for el in span:
            if el not in keep and el.getparent() is parent:
                parent.remove(el)
    # the fields' bookmarks would otherwise still draw brackets
    for tag in ("bookmarkStart", "bookmarkEnd", "permStart", "permEnd"):
        for el in list(root.iter(q(tag))):
            el.getparent().remove(el)
    return len(fields)


def cmd_flatten(path, out_path):
    root = load(path)
    n = flatten(root)
    save(path, out_path, root)
    print(f"flattened {n} field(s) -> {out_path}")


def main():
    a = sys.argv[1:]
    if len(a) >= 2 and a[0] == "list":
        cmd_list(a[1], a[3] if len(a) >= 4 and a[2] == "--json" else None)
    elif len(a) == 3 and a[0] == "flatten":
        cmd_flatten(*a[1:])
    elif len(a) == 3 and a[0] == "probe":
        cmd_probe(*a[1:])
    elif len(a) == 4 and a[0] == "fill":
        cmd_fill(*a[1:])
    else:
        sys.exit(__doc__)


if __name__ == "__main__":
    main()
