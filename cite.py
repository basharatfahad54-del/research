import pymupdf, re, collections
f = r"C:\Users\Muhammad Fahad\Desktop\Research paper\PhD DETAILED CONCEPT NOTE 2026_CHECK_261111_113441.pdf"
doc = pymupdf.open(f)
raw = re.sub(r"\s+"," "," ".join(p.get_text("text") for p in doc)).replace("\u2019","'")
pat = re.compile(r"([A-Z][A-Za-z'\-]+(?:\s+(?:and|&|et al\.?|,)\s*[A-Za-z'\-]*){0,4})[,\s]*\(?((?:19|20)\d{2})[a-z]?\)?")
c = collections.Counter()
for m in pat.finditer(raw):
    c[f"{m.group(1).strip()} ({m.group(2)})"] += 1
with open(r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\cites_raw.txt","w",encoding="utf-8") as fh:
    fh.write(f"TOTAL distinct patterns: {len(c)}\n")
    for k,v in c.most_common(320):
        fh.write(f"{v:3d}  {k}\n")
print("words total:", len(raw.split()))
print("distinct:", len(c))
