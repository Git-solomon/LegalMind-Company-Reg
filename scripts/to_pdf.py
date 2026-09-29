#!/usr/bin/env python3
"""Render a .docx to PDF with Microsoft Word (fallback: LibreOffice).

  python3 to_pdf.py <in.docx> <out.pdf>

Word gives the most faithful output for these Hebrew RTL government forms
(fonts, table layout, form-field results), so it is preferred.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))
import wordauto  # noqa: E402

if len(sys.argv) != 3:
    sys.exit(__doc__)
print(wordauto.convert(sys.argv[1], sys.argv[2], "pdf"))
