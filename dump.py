import pymupdf
f = r"C:\Users\Muhammad Fahad\Desktop\Research paper\PhD DETAILED CONCEPT NOTE 2026_CHECK_261111_113441.pdf"
doc = pymupdf.open(f)
def dump(a,b,out):
    with open(out,"w",encoding="utf-8") as fh:
        for n in range(a,b+1):
            fh.write(f"\n########## PDF PAGE {n} ##########\n")
            fh.write(doc[n-1].get_text("text"))
dump(6,26,r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\ch1_dump.txt")
dump(27,60,r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\ch2a_dump.txt")
dump(61,122,r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\ch2b_dump.txt")
dump(123,146,r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\ch3_dump.txt")
dump(147,199,r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\ch4_dump.txt")
print("done")
