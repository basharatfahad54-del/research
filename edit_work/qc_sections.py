# -*- coding: utf-8 -*-
"""QC on an edited docx: per-section word counts vs original, structural checks."""
import sys, re
import docx
from docx.oxml.ns import qn

PATH = sys.argv[1] if len(sys.argv) > 1 else r"Edited_Proposal_Output\Research_Proposal_Refined.docx"
d = docx.Document(PATH)

# --- per-section counts -------------------------------------------------
SEC_RE = re.compile(r"^(\d+)\.\s")
NAME_BY_NUM = {3: "3. Introduction", 4: "4. Literature Review", 5: "5. Theoretical Framework",
               6: "6. Conceptual Framework", 7: "7. Research Design and Methodology",
               8: "8. Context and Case Study", 9: "9. Expected Contributions",
               10: "10. Research Plan and Timeline", 11: "11. Preliminary Chapter Outline"}
ORIG = {"2. Abstract": 308, "3. Introduction": 3300, "4. Literature Review": 3377,
        "5. Theoretical Framework": 8969, "6. Conceptual Framework": 4090,
        "7. Research Design and Methodology": 3702, "8. Context and Case Study": 1689,
        "9. Expected Contributions": 1284, "10. Research Plan and Timeline": 1181,
        "11. Preliminary Chapter Outline": 349, "12. References": 1596, "Appendices": 1474}

sections = {}
cur = "Front matter"
for para in d.paragraphs:
    t = para.text.strip()
    if not t:
        continue
    style = para.style.name if para.style else ""
    if t.startswith("2. Abstract"):
        cur = "2. Abstract"
    elif t.startswith("12. References"):
        cur = "12. References"
    elif t.startswith("Appendices") or ("Appendix" in t and style.startswith("Heading")):
        cur = "Appendices"
    elif style.startswith("Heading"):
        m = SEC_RE.match(t)
        if m and int(m.group(1)) in NAME_BY_NUM:
            cur = NAME_BY_NUM[int(m.group(1))]
    sections[cur] = sections.get(cur, 0) + len(t.split())

total = 0
for k, v in sections.items():
    total += v
    o = ORIG.get(k)
    if o:
        print(f"{k:45s} {v:6d}  (orig {o:6d}, {100 * (o - v) / o:+.0f}%)")
    else:
        print(f"{k:45s} {v:6d}")
print(f"{'TOTAL':45s} {total:6d}  (orig 31391, {100 * (31391 - total) / 31391:+.1f}%)")

print("paragraphs:", len(d.paragraphs), "| tables:", len(d.tables),
      "| images:", sum(1 for p in d.paragraphs for _ in p._p.findall(".//" + qn("w:drawing"))),
      "| headings:", sum(1 for p in d.paragraphs if p.style and p.style.name.startswith("Heading")))
