import pymupdf
f = r"C:\Users\Muhammad Fahad\Desktop\Research paper\PhD DETAILED CONCEPT NOTE 2026_CHECK_261111_113441.pdf"
doc = pymupdf.open(f)
for i,p in enumerate(doc, 1):
    lines=[l.strip() for l in p.get_text("text").splitlines() if l.strip()]
    head=" | ".join(lines[:3])[:150]
    print(f"{i:3d} [{len(p.get_text('text')):5d}] {head}")
