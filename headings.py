import pymupdf, re
f = r"C:\Users\Muhammad Fahad\Desktop\Research paper\PhD DETAILED CONCEPT NOTE 2026_CHECK_261111_113441.pdf"
doc = pymupdf.open(f)
pat = re.compile(r"^\s*((\d+(\.\d+)*\.?)|(Chapter\s+\d+))\s+\S")
for i,p in enumerate(doc, 1):
    for l in p.get_text("text").splitlines():
        s=l.strip()
        if 3 < len(s) < 95 and pat.match(s) and not s.endswith('.') is False:
            pass
        if 3 < len(s) < 95 and pat.match(s):
            print(f"p{i:3d}  {s}")
