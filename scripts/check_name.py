#!/usr/bin/env python3
"""Check whether proposed company names are already taken in the Israeli
Registrar of Companies.

  python3 check_name.py "אלפא טכנולוגיות בע\"מ" "אלפא טק בע\"מ" [--json]

Source: the Registrar's official company list published on data.gov.il
(resource f004176c-..., refreshed daily). The Registrar's own site
("תאגידים ברשת") is an interactive app without an open API, so this is the
reliable programmatic source.

Limitations to tell the user: the list does not include names reserved or
applications still being processed, and the Registrar can also reject a name
that is merely similar or misleading. Treat "פנוי" as "no conflict found",
not as a guarantee; the final check happens when filing online.
"""
import json
import re
import subprocess
import sys
import urllib.parse
from difflib import SequenceMatcher

API = "https://data.gov.il/api/3/action/datastore_search"
RESOURCE = "f004176c-b85f-4542-8901-7b3176f9a054"
NAME, EN_NAME, NUMBER, STATUS = "שם חברה", "שם באנגלית", "מספר חברה", "סטטוס חברה"

# suffixes and noise that do not distinguish one name from another
SUFFIXES = [r"בע\W?מ", r"בעמ", r"\bltd\b", r"\blimited\b", r"\binc\b", r"חל\W?צ", r"\(\s*\d{4}\s*\)"]
FINALS = str.maketrans("ךםןףץ", "כמנפצ")
# generic business words: two names differing only in these are still "similar"
GENERIC = {normalize_word for normalize_word in """
טכנולוגיות טכנולוגיה השקעות שירותים ייעוץ יעוץ נכסים אחזקות החזקות יזמות סחר בניה בנייה
פתרונות מערכות תוכנה תכנה ניהול הנדסה פיתוח שיווק יבוא יצוא ויצוא ייבוא קבוצת קבוצה
ישראל גרופ חברה חברת מדיה דיגיטל אנרגיה נדלן נדל"ן פרויקטים הפקות תעשיות
technologies technology tech investments investment holdings services solutions systems
software group israel consulting management media digital energy capital ventures
""".translate(str.maketrans("ךםןףץ", "כמנפצ")).split()}


def normalize(name):
    s = name.lower().replace("~", '"')
    for pat in SUFFIXES:
        s = re.sub(pat, " ", s)
    s = re.sub(r"[^\w\s]", " ", s).translate(FINALS)
    return " ".join(s.split())


def distinctive_words(normalized):
    return [w for w in normalized.split() if w not in GENERIC] or normalized.split()


def distinctive(normalized):
    return "".join(distinctive_words(normalized))


def query(q, limit=1000):
    params = urllib.parse.urlencode({"resource_id": RESOURCE, "q": q, "limit": limit,
                                     "fields": ",".join([NUMBER, NAME, EN_NAME, STATUS])})
    # curl rather than urllib: python.org builds on macOS often lack CA certificates
    out = subprocess.run(["curl", "-sS", "-m", "60", f"{API}?{params}"],
                         capture_output=True, text=True, check=True).stdout
    data = json.loads(out)
    if not data.get("success"):
        raise RuntimeError(f"data.gov.il error: {data.get('error')}")
    return data["result"]["records"]


def check(name):
    core = normalize(name)
    if not core:
        return {"name": name, "verdict": "שם ריק אחרי ניקוי", "exact": [], "similar": []}
    dwords = distinctive_words(core)
    records = {}
    # full name (AND), the distinctive words only (AND), and the longest distinctive word
    for q in {core, " ".join(dwords), max(dwords, key=len)}:
        for rec in query(q):
            records[rec[NUMBER]] = rec
    exact, similar = [], []
    compact, core_d = core.replace(" ", ""), distinctive(core)
    for rec in records.values():
        for field in (NAME, EN_NAME):
            other = normalize(rec.get(field) or "")
            if not other:
                continue
            oc = other.replace(" ", "")
            item = {"number": rec[NUMBER], "name": (rec[NAME] or "").replace("~", '"'),
                    "english": rec.get(EN_NAME) or "", "status": rec[STATUS]}
            if oc == compact:
                exact.append(item)
                break
            other_d = distinctive(other)
            ratio = SequenceMatcher(None, core_d, other_d).ratio()
            if other_d == core_d:
                item["why"], item["similarity"] = "אותו חלק ייחודי", 1.0
            elif ratio >= 0.85:
                item["why"], item["similarity"] = "כתיב דומה", round(ratio, 2)
            elif (set(dwords) <= set(distinctive_words(other))
                  or set(distinctive_words(other)) <= set(dwords)):
                item["why"], item["similarity"] = "מכיל", round(ratio, 2)
            else:
                continue
            similar.append(item)
            break
    similar.sort(key=lambda x: (x["status"] != "פעילה", -x["similarity"]))
    active_exact = [e for e in exact if e["status"] == "פעילה"]
    if active_exact:
        verdict = "תפוס"
    elif exact:
        verdict = "קיים שם זהה של חברה לא פעילה — לבדוק מול הרשם"
    elif similar:
        verdict = "לא נמצא שם זהה, יש שמות דומים"
    else:
        verdict = "לא נמצא שם זהה או דומה"
    return {"name": name, "normalized": core, "verdict": verdict,
            "exact": exact, "similar": similar[:15], "similar_total": len(similar)}


def main():
    args = [a for a in sys.argv[1:] if a != "--json"]
    if not args:
        sys.exit(__doc__)
    try:
        results = [check(n) for n in args]
    except (subprocess.CalledProcessError, json.JSONDecodeError, RuntimeError) as e:
        sys.exit(f"הבדיקה נכשלה — אין תשובה תקינה מ־data.gov.il ({e}). אין לראות בשם פנוי.")
    if "--json" in sys.argv:
        print(json.dumps(results, ensure_ascii=False, indent=1))
        return
    for r in results:
        print(f"\n=== {r['name']}  ->  {r['verdict']}")
        for e in r["exact"]:
            print(f"  זהה:  {e['number']}  {e['name']}  ({e['status']})")
        for s in r["similar"]:
            print(f"  דומה: {s['number']}  {s['name']}  [{s['english']}]  ({s['status']}; {s['why']})")
        if r.get("similar_total", 0) > len(r["similar"]):
            print(f"  ...ועוד {r['similar_total'] - len(r['similar'])} שמות דומים")


if __name__ == "__main__":
    main()
