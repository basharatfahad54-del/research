import re
p = r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\00_extracted_text\PhD DETAILED CONCEPT NOTE 2026_CHECK_261111_113441.txt"
t = open(p, encoding="utf-8").read()
for term in ["posthuman","Posthuman","just transition","Just Transition","energy justice","Energy justice","TAM","Technology Acceptance","energy democracy","capability approach","socio-technical","Socio-technical","intersectional"]:
    print(f"{term:25s} {t.count(term)}")
print("gender mentions:", t.lower().count("gender"))
print("women mentions:", t.lower().count("women"))
