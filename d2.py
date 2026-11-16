import pymupdf
f = r"C:\Users\Muhammad Fahad\Desktop\Research paper\PhD DETAILED CONCEPT NOTE 2026_CHECK_261111_113441.pdf"
doc = pymupdf.open(f)
with open(r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\front_dump.txt","w",encoding="utf-8") as fh:
    for n in [3,4,5]:
        fh.write(f"\n##### PDF PAGE {n} #####\n"+doc[n-1].get_text("text"))
print("ok")
