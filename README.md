# LegalMind-Company-Reg
סקיל לכתיבת מסמכים להקמת חברה בישראל על בסיס מסמכי הרשם. מיועד לשימוש של עורכי דין בלבד ואינו מהווה תחליף לשיקול דעתם המקצועית.

## מה הסקיל עושה
מקבל מייל של לקוח (‎.eml, טקסט, PDF או Gmail) ומפיק חבילת PDF של טפסי רשם החברות, מלאה ומוכנה לחתימה:
1. מסווג את סוג החברה (חברת יחיד / חברה רגילה / לתועלת הציבור) ובוחר את הטפסים הרלוונטיים.
2. שולף מהמייל את הפרטים (בעלי מניות, דירקטורים, הון, מען וכו').
3. בודק אם שם החברה תפוס ברשימת החברות של רשם החברות (data.gov.il).
4. שואל שאלה מרוכזת אחת על כל פרט חסר — לא ממציא ערכים.
5. ממלא את הטפסים ומפיק PDF.

## התקנה
macOS:
```bash
git clone git@github.com:Git-solomon/LegalMind-Company-Reg.git ~/.claude/skills/company-incorporation-docs
python3 -m pip install lxml
python3 ~/.claude/skills/company-incorporation-docs/scripts/selftest.py
```

Windows (PowerShell):
```powershell
git clone git@github.com:Git-solomon/LegalMind-Company-Reg.git "$HOME\.claude\skills\company-incorporation-docs"
py -m pip install lxml
py "$HOME\.claude\skills\company-incorporation-docs\scripts\selftest.py"
```
`selftest.py` מריץ את כל השלבים עם נתוני דמה ומדווח אם משהו חסר.

## דרישות
- Microsoft Word (macOS או Windows), או LibreOffice כחלופה — להמרת ‎.doc ולהפקת PDF.
- Python 3 עם `lxml`.
- `curl` (לבדיקת שם החברה) — מובנה ב־macOS וב־Windows 10 ומעלה.

## מבנה
- `SKILL.md` — ההנחיות לסקיל.
- `templates/` — טפסי רשם החברות, שמורים פתוחים לעריכה (בלי הנעילה של הרשם).
- `references/` — איזה טופס לאיזה סוג חברה, ומפת השדות של כל טופס.
- `scripts/` — מילוי, הפקת PDF, בדיקת שם, ופתיחה של טופס נעול חדש מהרשם.

פרטי עורך הדין המאמת נשמרים מקומית ב־`lawyer-profile.json`, שאינו נשמר בגיט.
