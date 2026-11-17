# -*- coding: utf-8 -*-
"""Verify pass-2 edits and typography in Research_Proposal_Final.docx."""
import docx

fin = docx.Document(r"research-main\Edited_Proposal_Output\Research_Proposal_Final.docx")
ref = docx.Document(r"research-main\Edited_Proposal_Output\Research_Proposal_Refined.docx")

fp, rp = list(fin.paragraphs), list(ref.paragraphs)
print("paragraph counts equal:", len(fp) == len(rp))

# 1. six REVS present
checks = ["is a Microsoft Word field",
          "draft instruments and workshop protocol",
          "impose an additional burden",
          "the paragraphs that follow show what it means for each question",
          "Ethics is addressed both through procedural clearance",
          "to conduct qualitative research in a context"]
body_fin = "\n".join(p.text for p in fp)
for c in checks:
    print("REV present:", c[:52], "->", c in body_fin)

# 2. Table 3.6 cell
print("Table 3.6 RQ5 cell:", fin.tables[11].rows[4].cells[2].text.strip())

# 3. glossary order
g = fin.tables[30]
terms = [g.rows[i].cells[0].text.strip() for i in range(1, len(g.rows))]
print("glossary terms:", terms)
print("glossary sorted:", terms == sorted(terms))

# 4. typography counts, refined vs final
body_ref = "\n".join(p.text for p in rp)


def cnt(text):
    return text.count("'"), text.count('"')


print("whole-doc straight ' refined -> final:", cnt(body_ref)[0], "->", cnt(body_fin)[0])
print("whole-doc straight \" refined -> final:", cnt(body_ref)[1], "->", cnt(body_fin)[1])

# remaining straight apostrophes: where?
for i, p in enumerate(fp):
    if "'" in p.text:
        print(f"  remaining ' at P{i}: {p.text[:80]}")

# remaining straight double quotes: where?
for i, p in enumerate(fp):
    if '"' in p.text:
        print(f"  remaining \" at P{i}: {p.text[:80]}")

# straight chars inside table cells (excluding reference-note box 27)
print("--- table cells with straight quotes/apostrophes ---")
for ti, t in enumerate(fin.tables):
    if ti == 27:
        continue
    for r in t.rows:
        for c in r.cells:
            if "'" in c.text or '"' in c.text:
                print(f"  T{ti}: {c.text[:80]}")

# 5. references unchanged: compare P533-P605
same = all(fp[i].text == rp[i].text for i in range(533, 606))
print("reference paragraphs identical:", same)

# 6. heading count + placeholders sanity
heads = [p for p in fp if p.style and p.style.name.startswith("Heading")]
print("headings:", len(heads))
