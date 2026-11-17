# -*- coding: utf-8 -*-
"""Supplemental pass on Final: plural possessives (residents' -> residents\u2019),
then re-verify. Same exclusions: reference paragraphs P533-P605, tables[27],
and cells containing search syntax (' OR ')."""
import re
import docx
from docx.oxml.ns import qn

PATH = r"research-main\Edited_Proposal_Output\Research_Proposal_Final.docx"
d = docx.Document(PATH)
paras = list(d.paragraphs)
REF_RANGE = set(range(533, 606))

# 1) plural possessive sweep: apostrophe directly after a word character
rule = re.compile(r"(?<=\w)'")
changed = 0
for idx, p in enumerate(paras):
    if idx in REF_RANGE:
        continue
    for node in p._p.iter(qn("w:t")):
        if node.text and "'" in node.text:
            new = rule.sub("\u2019", node.text)
            if new != node.text:
                node.text = new
                changed += 1

note_tbl_el = d.tables[27]._tbl
for t in d.tables:
    if t._tbl is note_tbl_el:
        continue
    for row in t.rows:
        for c in row.cells:
            full = c.text
            if " OR " in full:
                continue
            for node in c._tc.iter(qn("w:t")):
                if node.text and "'" in node.text:
                    new = rule.sub("\u2019", node.text)
                    if new != node.text:
                        node.text = new
                        changed += 1
print("plural-possessive fixes:", changed)
d.save(PATH)

# 2) re-verify
d2 = docx.Document(PATH)
fp = list(d2.paragraphs)
body = "\n".join(p.text for p in fp)
print("straight apos in body (excl refs):",
      sum(p.text.count("'") for i, p in enumerate(fp) if i not in REF_RANGE))
print("straight quotes in body (excl refs):",
      sum(p.text.count('"') for i, p in enumerate(fp) if i not in REF_RANGE))
for ti, t in enumerate(d2.tables):
    if ti == 27:
        continue
    for r in t.rows:
        for c in r.cells:
            if "' OR " in c.text or " OR '" in c.text:
                continue
            if "'" in c.text and " OR " not in c.text:
                print(f"  T{ti} straight ': {c.text[:70]}")
            if '"' in c.text and " OR " not in c.text:
                print(f"  T{ti} straight \": {c.text[:70]}")

# TOC field still present?
xml = d2.element.body.xml
print("TOC field present:", 'TOC \\o' in xml and 'fldChar' in xml)
# spot-check paired quotes rendered correctly
for t in (d2.tables[9], d2.tables[10]):
    pass
cell = d2.tables[10].rows[4].cells[2]
print("T10 RQ3 cell:", cell.text.strip())
print("P311 caption:", fp[311].text[:90])
print("P47 old check:", fp[47].text[:110])
words = sum(len(p.text.split()) for p in fp if p.text.strip())
print("body words:", words)
