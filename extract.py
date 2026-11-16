import pymupdf, os, json, glob
src = r"C:\Users\Muhammad Fahad\Desktop\Research paper"
out = os.path.join(src, "research_work", "00_extracted_text")
report = []
for f in sorted(glob.glob(os.path.join(src, "*.pdf"))):
    doc = pymupdf.open(f)
    txt = []
    chars = 0
    for p in doc:
        t = p.get_text("text")
        chars += len(t.strip())
        txt.append(t)
    name = os.path.splitext(os.path.basename(f))[0]
    safe = "".join(c if c.isalnum() or c in " -_()" else "_" for c in name).strip()
    with open(os.path.join(out, safe + ".txt"), "w", encoding="utf-8") as fh:
        fh.write("\n\n=== PAGE BREAK ===\n\n".join(txt))
    report.append({"file": os.path.basename(f), "pages": len(doc), "chars": chars, "out": safe + ".txt"})
    doc.close()
print(json.dumps(report, indent=2))
