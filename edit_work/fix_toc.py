# -*- coding: utf-8 -*-
"""Restore the TOC field in Research_Proposal_Final.docx.

The Refined document's TOC is a real Word field whose *cached result* is the
instruction note; the pass-2 P22 rewrite cleared the field characters along
with the old text. This restores the field, with the polished note as the
cached result (Word replaces it with the generated contents on F9)."""
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

PATH = r"research-main\Edited_Proposal_Output\Research_Proposal_Final.docx"
d = docx.Document(PATH)
paras = list(d.paragraphs)

NOTE = ("The table of contents is a Microsoft Word field: select the whole document "
        "(Ctrl+A) and press F9, or right-click on the field and choose "
        "\u201cUpdate Field\u201d, to generate or refresh the list of contents.")

p = paras[22]
assert p.text.startswith("The table of contents is a Microsoft Word field"), p.text[:60]
assert "instrText" not in p._p.xml
assert all("instrText" not in q._p.xml for q in paras), "unexpected field elsewhere"

for r in list(p.runs):
    r._r.getparent().remove(r._r)


def mk(tag, **attrs):
    el = OxmlElement(tag)
    for k, v in attrs.items():
        el.set(qn(k), v)
    return el


r1 = mk("w:r"); r1.append(mk("w:fldChar", **{"w:fldCharType": "begin"}))
r2 = mk("w:r")
instr = mk("w:instrText", **{"xml:space": "preserve"})
instr.text = ' TOC \\o "1-3" \\h \\z \\u '
r2.append(instr)
r3 = mk("w:r"); r3.append(mk("w:fldChar", **{"w:fldCharType": "separate"}))
r4 = mk("w:r")
t = OxmlElement("w:t")
t.text = NOTE
r4.append(t)
r5 = mk("w:r"); r5.append(mk("w:fldChar", **{"w:fldCharType": "end"}))
for el in (r1, r2, r3, r4, r5):
    p._p.append(el)

d.save(PATH)

# verify
d2 = docx.Document(PATH)
p2 = list(d2.paragraphs)[22]
xml = p2._p.xml
print("field restored:", "instrText" in xml and xml.count("fldChar") == 3)
print("cached note:", p2.text[:60])
print("words unchanged check:", sum(len(q.text.split()) for q in d2.paragraphs if q.text.strip()))
