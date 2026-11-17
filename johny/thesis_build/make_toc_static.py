# -*- coding: utf-8 -*-
"""Builds the static Table_of_Contents.docx from toc_entries.txt (from dump_toc.ps1)."""
import io, os, re
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_TAB_ALIGNMENT, WD_TAB_LEADER
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = r"C:\Users\Muhammad Fahad\Desktop\research-main\research-main"
TB = os.path.join(ROOT, "johny", "thesis_build")
OUT = os.path.join(ROOT, "johny", "Final_Thesis", "Table_of_Contents.docx")
TNR = "Times New Roman"

with io.open(os.path.join(TB, "toc_entries.txt"), "r", encoding="utf-8") as f:
    raw = [ln for ln in f.read().splitlines() if ln.strip()]

entries = []
for ln in raw:
    parts = ln.split("<<>>")
    if len(parts) < 3:
        continue
    style, label, page = parts[0].strip(), parts[1].strip(), parts[2].strip()
    level = 1
    if "2" in style:
        level = 2
    elif "3" in style:
        level = 3
    entries.append((level, label, page))

doc = Document()
st = doc.styles["Normal"]
st.font.name = TNR
st.font.size = Pt(12)
st.element.rPr.rFonts.set(qn("w:eastAsia"), TNR)
st.paragraph_format.line_spacing = 1.5
st.paragraph_format.space_after = Pt(4)
sec = doc.sections[0]
sec.page_width, sec.page_height = Cm(21.0), Cm(29.7)
for m in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
    setattr(sec, m, Cm(2.54))
footer = sec.footer
fp = footer.paragraphs[0]
fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = fp.add_run()
fb, it, fe = OxmlElement("w:fldChar"), OxmlElement("w:instrText"), OxmlElement("w:fldChar")
fb.set(qn("w:fldCharType"), "begin")
it.set(qn("xml:space"), "preserve")
it.text = " PAGE "
fe.set(qn("w:fldCharType"), "end")
r._r.append(fb); r._r.append(it); r._r.append(fe)
r.font.name = TNR
r.font.size = Pt(11)

h = doc.add_paragraph()
hr = h.add_run("Table of Contents")
hr.font.name = TNR
hr.font.size = Pt(16)
hr.bold = True
note = doc.add_paragraph()
nr = note.add_run("Page numbers refer to Complete_Thesis.docx/.pdf. The compiled thesis itself "
                  "contains an automatic, updatable table of contents field.")
nr.font.name = TNR
nr.font.size = Pt(10)
nr.italic = True
nr.font.color.rgb = RGBColor(0x50, 0x50, 0x50)

skip_first = False
for idx, (level, label, page) in enumerate(entries):
    par = doc.add_paragraph()
    pf = par.paragraph_format
    pf.line_spacing = 1.15
    pf.space_after = Pt(2)
    indent = Cm(0.0 if level == 1 else (0.6 if level == 2 else 1.2))
    pf.left_indent = indent
    pf.tab_stops.add_tab_stop(Cm(15.9 - (0.0 if level == 1 else (0.6 if level == 2 else 1.2))),
                              WD_TAB_ALIGNMENT.RIGHT, WD_TAB_LEADER.DOTS)
    run = par.add_run(label)
    run.font.name = TNR
    run.font.size = Pt(12)
    if level == 1:
        run.bold = True
    try:
        pageno = str(int(page))
    except ValueError:
        pageno = page
    r2 = par.add_run("\t" + pageno)
    r2.font.name = TNR
    r2.font.size = Pt(12)
    if level == 1:
        r2.bold = True

doc.save(OUT)
print("WROTE", OUT, len(entries), "entries")
