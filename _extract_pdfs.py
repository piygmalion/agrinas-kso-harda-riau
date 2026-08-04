# -*- coding: utf-8 -*-
import os
import pdfplumber

base = r"C:\Users\Patron\Downloads\sawit lagi"
out = os.path.join(base, "_extract")
os.makedirs(out, exist_ok=True)

pdfs = [f for f in os.listdir(base) if f.lower().endswith(".pdf")]
# skip duplicate
pdfs = [f for f in pdfs if "(1)" not in f]

for f in pdfs:
    path = os.path.join(base, f)
    try:
        pages = []
        with pdfplumber.open(path) as pdf:
            for i, page in enumerate(pdf.pages):
                t = page.extract_text() or ""
                pages.append(f"--- PAGE {i+1} ---\n{t}")
                tables = page.extract_tables() or []
                for ti, table in enumerate(tables):
                    pages.append(f"=== TABLE p{i+1}-{ti+1} ===")
                    for row in table:
                        cells = [(c or "").replace("\n", " ").strip() for c in row]
                        pages.append(" | ".join(cells))
        text = "\n".join(pages)
        with open(os.path.join(out, f + ".txt"), "w", encoding="utf-8") as fh:
            fh.write(text)
        print(f"OK {f}: chars={len(text)}")
    except Exception as e:
        print(f"FAIL {f}: {e}")
