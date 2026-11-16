import pymupdf
f = r"C:\Users\Muhammad Fahad\Desktop\Research paper\PhD DETAILED CONCEPT NOTE 2026_CHECK_261111_113441.pdf"
doc = pymupdf.open(f)
ranges = {"Front matter (TOC, lists, abstract)":(1,6),"Ch1 Introduction":(7,26),"Ch2 Literature review":(27,122),"Conceptual framework ('Chapter 1')":(123,146),"Methodology ('Chapter 2')":(147,186),"Proposal tail (contributions, plan, ethics, TOC)":(187,199)}
tot=0
for k,(a,b) in ranges.items():
    w=sum(len(doc[i-1].get_text("text").split()) for i in range(a,b+1))
    tot+=w
    print(f"{k:52s} pp {a:3d}-{b:3d}  {b-a+1:3d} pp  {w:6d} words")
print(f"{'TOTAL':52s} pp   1-199  {199:3d} pp  {tot:6d} words")
# tables/figures counts
import re
t=re.sub(r"\s+"," "," ".join(p.get_text("text") for p in doc))
figs = re.findall(r"Figure\s+\d+\.\d+", t)
tabs = re.findall(r"Table\s+\d+\.\d+", t)
print("Figure mentions:", len(figs), "unique:", len(set(figs)))
print("Table mentions:", len(tabs), "unique:", len(set(tabs)))
