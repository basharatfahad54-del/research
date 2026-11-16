# -*- coding: utf-8 -*-
"""
Builds the refined SETILL PhD proposal from markdown-like content files.
Internal build script: not shipped to the client.
"""
import os, re, sys, glob
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = r"C:\Users\Muhammad Fahad\Desktop\Research paper"
BASE = os.path.join(ROOT, "research_work", "proposal")
ASSETS = os.path.join(BASE, "assets")
OUTDIR = os.path.join(ROOT, "docs")
os.makedirs(OUTDIR, exist_ok=True)

TNR = "Times New Roman"
NAVY = RGBColor(0x1F, 0x38, 0x64)
GREY = RGBColor(0x40, 0x40, 0x40)


# --------------------------------------------------------------------------
# low level
# --------------------------------------------------------------------------
def _rfonts(rpr, name=TNR):
    rf = rpr.find(qn('w:rFonts'))
    if rf is None:
        rf = OxmlElement('w:rFonts')
        rpr.append(rf)
    for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
        rf.set(qn(a), name)


def style_font(style, size=12, bold=None, italic=None, color=(0, 0, 0)):
    style.font.name = TNR
    style.font.size = Pt(size)
    if bold is not None:
        style.font.bold = bold
    if italic is not None:
        style.font.italic = italic
    style.font.color.rgb = RGBColor(*color)
    _rfonts(style.element.get_or_add_rPr())


def base_document():
    doc = Document()
    for s in doc.sections:
        s.page_width = Cm(21.0)
        s.page_height = Cm(29.7)
        s.top_margin = Cm(2.54)
        s.bottom_margin = Cm(2.54)
        s.left_margin = Cm(3.0)
        s.right_margin = Cm(2.54)
        s.header_distance = Cm(1.25)
        s.footer_distance = Cm(1.25)

    n = doc.styles['Normal']
    style_font(n, 12)
    pf = n.paragraph_format
    pf.line_spacing = 1.5
    pf.space_after = Pt(8)
    pf.space_before = Pt(0)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    pf.widow_control = True

    specs = (('Heading 1', 14, True), ('Heading 2', 12.5, True), ('Heading 3', 12, True))
    for name, size, bold in specs:
        st = doc.styles[name]
        style_font(st, size, bold=bold, italic=False, color=(0x1F, 0x38, 0x64))
        st.paragraph_format.space_before = Pt(18 if name == 'Heading 1' else 12)
        st.paragraph_format.space_after = Pt(8 if name == 'Heading 1' else 6)
        st.paragraph_format.keep_with_next = True
        st.paragraph_format.line_spacing = 1.3
        st.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return doc


def field(paragraph, instr, placeholder="1", size=9):
    run = paragraph.add_run()
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = instr
    f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t'); t.text = placeholder
    f3 = OxmlElement('w:fldChar'); f3.set(qn('w:fldCharType'), 'end')
    r = run._r
    r.append(f1); r.append(it); r.append(f2); r.append(t); r.append(f3)


def set_pgnum(section, fmt=None, start=None):
    sectPr = section._sectPr
    for old in sectPr.findall(qn('w:pgNumType')):
        sectPr.remove(old)
    pg = OxmlElement('w:pgNumType')
    if fmt:
        pg.set(qn('w:fmt'), fmt)
    if start is not None:
        pg.set(qn('w:start'), str(start))
    sectPr.append(pg)


def add_header_footer(section, header_text, roman=False):
    section.different_first_page_header_footer = True
    hp = section.header.paragraphs[0]
    hp.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    r = hp.add_run(header_text)
    r.font.name = TNR; r.font.size = Pt(8.5); r.font.color.rgb = RGBColor(0x59, 0x59, 0x59)
    pPr = hp._p.get_or_add_pPr()
    bd = OxmlElement('w:pBdr'); bo = OxmlElement('w:bottom')
    bo.set(qn('w:val'), 'single'); bo.set(qn('w:sz'), '4'); bo.set(qn('w:space'), '1'); bo.set(qn('w:color'), 'BFBFBF')
    bd.append(bo); pPr.append(bd)

    fp = section.footer.paragraphs[0]
    fp.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = fp.add_run("Page ")
    r.font.name = TNR; r.font.size = Pt(9); r.font.color.rgb = GREY
    field(fp, ' PAGE ')
    r = fp.add_run(" of ")
    r.font.name = TNR; r.font.size = Pt(9); r.font.color.rgb = GREY
    field(fp, ' SECTIONPAGES ')
    for run in fp.runs:
        run.font.name = TNR; run.font.size = Pt(9); run.font.color.rgb = GREY


def add_toc(doc):
    head = doc.add_paragraph()
    head.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = head.add_run("Table of Contents")
    r.font.name = TNR; r.font.size = Pt(14); r.font.bold = True; r.font.color.rgb = NAVY
    head.paragraph_format.space_after = Pt(12)
    p = doc.add_paragraph()
    run = p.add_run()
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve')
    it.text = 'TOC \\o "1-2" \\h \\z \\u'
    f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t')
    t.text = ("[The table of contents is generated automatically. In Word, press Ctrl+A and then F9, "
              "or right-click here and choose Update Field.]")
    f3 = OxmlElement('w:fldChar'); f3.set(qn('w:fldCharType'), 'end')
    r = run._r
    r.append(f1); r.append(it); r.append(f2); r.append(t); r.append(f3)


def enable_update_fields(doc):
    settings = doc.settings.element
    uf = OxmlElement('w:updateFields'); uf.set(qn('w:val'), 'true')
    settings.append(uf)


# --------------------------------------------------------------------------
# blocks
# --------------------------------------------------------------------------
def h1(doc, text, first=False):
    p = doc.add_paragraph(text, style='Heading 1')
    if not first:
        p.paragraph_format.page_break_before = True
    p.paragraph_format.space_before = Pt(0)
    return p


def h2(doc, text):
    return doc.add_paragraph(text, style='Heading 2')


def h3(doc, text):
    return doc.add_paragraph(text, style='Heading 3')


def para(doc, text, italic=False, size=12, align=WD_ALIGN_PARAGRAPH.JUSTIFY,
         space_after=8, first_indent=0.0, left_indent=0.0):
    p = doc.add_paragraph()
    _runs_from_markup(p, text, size=size, italic=italic)
    p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    if first_indent:
        p.paragraph_format.first_line_indent = Cm(first_indent)
    if left_indent:
        p.paragraph_format.left_indent = Cm(left_indent)
    return p


def _runs_from_markup(p, text, size=12, italic=False):
    """Supports **bold** and *italic* inline."""
    tokens = re.split(r'(\*\*[^*]+\*\*|\*[^*]+\*)', text)
    for tok in tokens:
        if not tok:
            continue
        if tok.startswith('**') and tok.endswith('**'):
            r = p.add_run(tok[2:-2]); r.bold = True
        elif tok.startswith('*') and tok.endswith('*') and len(tok) > 2:
            r = p.add_run(tok[1:-1]); r.italic = True
        else:
            r = p.add_run(tok)
        r.font.name = TNR; r.font.size = Pt(size)
        if italic:
            r.italic = True


def bullet(doc, text, style='List Bullet', size=12):
    p = doc.add_paragraph(style=style)
    _runs_from_markup(p, text, size=size)
    p.paragraph_format.space_after = Pt(3)
    p.paragraph_format.line_spacing = 1.5
    p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.first_line_indent = Cm(-0.5)
    return p


def shade(cell, hexfill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hexfill)
    tcPr.append(shd)


def cell_margins(table, top=40, bottom=40, left=80, right=80):
    tblPr = table._tbl.tblPr
    mar = OxmlElement('w:tblCellMar')
    for side, val in (('top', top), ('bottom', bottom), ('left', left), ('right', right)):
        el = OxmlElement('w:' + side)
        el.set(qn('w:w'), str(val)); el.set(qn('w:type'), 'dxa')
        mar.append(el)
    tblPr.append(mar)


def repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    h = OxmlElement('w:tblHeader'); h.set(qn('w:val'), 'true')
    trPr.append(h)


def no_split(row):
    trPr = row._tr.get_or_add_trPr()
    h = OxmlElement('w:cantSplit'); h.set(qn('w:val'), 'true')
    trPr.append(h)


def add_table(doc, caption, headers, rows, widths=None, fsize=9, header_fill='1F3864'):
    if caption:
        p = doc.add_paragraph()
        r = p.add_run(caption)
        r.font.name = TNR; r.font.size = Pt(9.5); r.font.bold = True
        p.paragraph_format.space_before = Pt(10)
        p.paragraph_format.space_after = Pt(4)
        p.paragraph_format.keep_with_next = True
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    t = doc.add_table(rows=1, cols=len(headers))
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    cell_margins(t)
    if not widths:
        widths = [15.4 / len(headers)] * len(headers)
    scale = 15.4 / sum(widths)
    widths = [w * scale for w in widths]
    hdr = t.rows[0]
    repeat_header(hdr)
    no_split(hdr)
    for i, htxt in enumerate(headers):
        c = hdr.cells[i]; c.width = Cm(widths[i]); c.text = ''
        pp = c.paragraphs[0]
        pp.alignment = WD_ALIGN_PARAGRAPH.LEFT
        pp.paragraph_format.space_after = Pt(2); pp.paragraph_format.line_spacing = 1.0
        rr = pp.add_run(htxt)
        rr.font.name = TNR; rr.font.size = Pt(fsize); rr.bold = True
        rr.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
        shade(c, header_fill)
    for ri, row in enumerate(rows):
        cells = t.add_row().cells
        no_split(t.rows[-1])
        for i, val in enumerate(row):
            if i >= len(cells):
                continue
            c = cells[i]; c.width = Cm(widths[i]); c.text = ''
            parts = str(val).split('\n')
            for j, part in enumerate(parts):
                pp = c.paragraphs[0] if j == 0 else c.add_paragraph()
                pp.alignment = WD_ALIGN_PARAGRAPH.LEFT
                pp.paragraph_format.space_after = Pt(1); pp.paragraph_format.line_spacing = 1.05
                _runs_from_markup(pp, part, size=fsize)
            if ri % 2 == 1:
                shade(c, 'F2F5FA')
    doc.add_paragraph().paragraph_format.space_after = Pt(2)
    return t


def add_figure(doc, path, width_cm=15.2, caption=None):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.space_before = Pt(8)
    p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    if os.path.exists(path):
        p.add_run().add_picture(path, width=Cm(width_cm))
    else:
        p.add_run("[figure missing: %s]" % os.path.basename(path))
    if caption:
        c = doc.add_paragraph()
        r = c.add_run(caption)
        r.font.name = TNR; r.font.size = Pt(9.5); r.italic = True
        c.alignment = WD_ALIGN_PARAGRAPH.LEFT
        c.paragraph_format.space_after = Pt(12)


def note_box(doc, text):
    t = doc.add_table(rows=1, cols=1)
    t.style = 'Table Grid'
    t.alignment = WD_TABLE_ALIGNMENT.CENTER
    t.autofit = False
    c = t.rows[0].cells[0]
    c.width = Cm(15.4)
    shade(c, 'EDF2F9')
    c.text = ''
    p = c.paragraphs[0]
    p.paragraph_format.space_before = Pt(5)
    p.paragraph_format.space_after = Pt(5)
    p.paragraph_format.line_spacing = 1.25
    p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    _runs_from_markup(p, text, size=10)
    doc.add_paragraph().paragraph_format.space_after = Pt(4)


def reference_entry(doc, text):
    p = doc.add_paragraph()
    _runs_from_markup(p, text, size=11)
    p.paragraph_format.left_indent = Cm(1.0)
    p.paragraph_format.first_line_indent = Cm(-1.0)
    p.paragraph_format.space_after = Pt(8)
    p.paragraph_format.line_spacing = 1.15
    p.alignment = WD_ALIGN_PARAGRAPH.LEFT
    return p


# --------------------------------------------------------------------------
# cover + front matter
# --------------------------------------------------------------------------
def cover(doc):
    for _ in range(2):
        doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("UNIVERSITY OF JOHANNESBURG")
    r.font.name = TNR; r.font.size = Pt(12); r.bold = True; r.font.color.rgb = NAVY
    p.paragraph_format.space_after = Pt(2)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Faculty of Engineering and the Built Environment")
    r.font.name = TNR; r.font.size = Pt(10.5); r.font.color.rgb = GREY
    p.paragraph_format.space_after = Pt(0)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Postgraduate School of Engineering Management")
    r.font.name = TNR; r.font.size = Pt(10.5); r.font.color.rgb = GREY
    p.paragraph_format.space_after = Pt(40)

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = p._p.get_or_add_pPr()
    bd = OxmlElement('w:pBdr'); bo = OxmlElement('w:bottom')
    bo.set(qn('w:val'), 'single'); bo.set(qn('w:sz'), '8'); bo.set(qn('w:space'), '1'); bo.set(qn('w:color'), '1F3864')
    bd.append(bo); pPr.append(bd)
    p.paragraph_format.space_after = Pt(28)

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Developing an Innovative Systemic Living Lab\nfor Just Energy Transition in a\nSouth African Informal Settlement")
    r.font.name = TNR; r.font.size = Pt(21); r.bold = True; r.font.color.rgb = NAVY
    p.paragraph_format.line_spacing = 1.25
    p.paragraph_format.space_after = Pt(18)

    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("A research proposal submitted in fulfilment of the requirements\nfor the degree of Doctor of Philosophy in Engineering Management")
    r.font.name = TNR; r.font.size = Pt(12); r.italic = True
    p.paragraph_format.space_after = Pt(34)

    meta = [("Candidate", "[Candidate Name]"),
            ("Student number", "[Student Number]"),
            ("Supervisor", "Dr K. Aboalez"),
            ("Department", "Postgraduate School of Engineering Management"),
            ("Date", "November 2026")]
    for k, v in meta:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.1
        r = p.add_run(k + ":  "); r.font.name = TNR; r.font.size = Pt(11); r.bold = True
        r2 = p.add_run(v); r2.font.name = TNR; r2.font.size = Pt(11)
    for _ in range(3):
        doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("Proposal version 1.0  |  November 2026")
    r.font.name = TNR; r.font.size = Pt(9); r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    doc.add_page_break()


def fm_head(doc, text):
    p = doc.add_paragraph(text, style='Heading 1')
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(12)
    for r in p.runs:
        r.font.name = TNR; r.font.size = Pt(14); r.bold = True; r.font.color.rgb = NAVY
    return p


def front_matter(doc, abstract, keywords, figs, tables, abbreviations):
    # declaration (page 2)
    p = fm_head(doc, "Declaration")
    r = p.runs[0]
    r.bold = True
    p.paragraph_format.space_after = Pt(14)
    for t in [
        "I declare that this research proposal is my own work. Where the work of others has been used, "
        "it has been acknowledged in the text and in the reference list. I have not submitted this proposal, "
        "in whole or in part, for any other degree or examination.",
        "I confirm that no fieldwork has been conducted for this study and that no empirical data exist at the "
        "time of submission. The methods described in Section 7 are proposed methods. Fieldwork will begin only "
        "after written ethical clearance has been granted by the Faculty of Engineering and the Built Environment "
        "Research Ethics Committee and after permission has been obtained from the relevant municipality and "
        "community structures.",
        "I understand that the University of Johannesburg's policy on research integrity applies to this work, "
        "including its rules on the use of text-generation tools and on the citation of sources.",
    ]:
        para(doc, t)
    doc.add_paragraph()
    doc.add_paragraph()
    p = doc.add_paragraph()
    p.paragraph_format.space_after = Pt(2)
    r = p.add_run("Signature: ______________________________        Date: ____________________")
    r.font.name = TNR; r.font.size = Pt(11)
    doc.add_page_break()

    # abstract
    p = fm_head(doc, "Abstract")
    p.paragraph_format.space_after = Pt(12)
    para(doc, abstract)
    p = doc.add_paragraph()
    r = p.add_run("Keywords: "); r.bold = True; r.font.name = TNR; r.font.size = Pt(11)
    r2 = p.add_run(keywords); r2.font.name = TNR; r2.font.size = Pt(11)
    p.paragraph_format.space_before = Pt(10)
    doc.add_page_break()

    # toc
    add_toc(doc)
    doc.add_page_break()

    # list of figures
    p = fm_head(doc, "List of Figures")
    p.paragraph_format.space_after = Pt(12)
    for label, cap in figs:
        q = doc.add_paragraph()
        q.paragraph_format.space_after = Pt(6)
        q.paragraph_format.left_indent = Cm(1.1)
        q.paragraph_format.first_line_indent = Cm(-1.1)
        rr = q.add_run(label + "\t"); rr.font.name = TNR; rr.font.size = Pt(11); rr.bold = True
        rr2 = q.add_run(cap); rr2.font.name = TNR; rr2.font.size = Pt(11)

    doc.add_page_break()
    p = fm_head(doc, "List of Tables")
    p.paragraph_format.space_after = Pt(12)
    for label, cap in tables:
        q = doc.add_paragraph()
        q.paragraph_format.space_after = Pt(6)
        q.paragraph_format.left_indent = Cm(1.1)
        q.paragraph_format.first_line_indent = Cm(-1.1)
        rr = q.add_run(label + "\t"); rr.font.name = TNR; rr.font.size = Pt(11); rr.bold = True
        rr2 = q.add_run(cap); rr2.font.name = TNR; rr2.font.size = Pt(11)

    doc.add_page_break()
    p = fm_head(doc, "List of Abbreviations")
    p.paragraph_format.space_after = Pt(12)
    add_table(doc, None, ["Abbreviation", "Meaning"], abbreviations, [3.2, 12.2], fsize=9.5)


# --------------------------------------------------------------------------
# parser
# --------------------------------------------------------------------------
def build_from_parts(parts):
    doc = base_document()
    sec1 = doc.sections[0]
    set_pgnum(sec1, fmt='lowerRoman', start=1)
    add_header_footer(sec1, "SETILL PhD Research Proposal")

    cover(doc)
    front_matter(doc, ABSTRACT, KEYWORDS, FIGURES, TABLES, ABBREVIATIONS)

    # body section, arabic page numbers restart at 1
    body = doc.add_section(WD_SECTION.NEW_PAGE)
    set_pgnum(body, fmt='decimal', start=1)

    first_h1 = True
    for part in parts:
        path = os.path.join(BASE, part)
        with open(path, encoding='utf-8') as f:
            lines = f.read().splitlines()

        i = 0
        in_table = False
        tbl_rows = []
        tbl_caption = None
        tbl_widths = None

        def flush_table():
            nonlocal in_table, tbl_rows, tbl_caption, tbl_widths
            if in_table and tbl_rows:
                headers = tbl_rows[0]
                rows = []
                for r in tbl_rows[1:]:
                    if r and all(re.fullmatch(r':?-{2,}:?', c.strip()) for c in r):
                        continue
                    rows.append(r)
                add_table(doc, tbl_caption, headers, rows, tbl_widths)
            in_table = False
            tbl_rows = []
            tbl_caption = None
            tbl_widths = None

        for raw in lines:
            line = raw.rstrip()
            if not line.strip():
                continue
            if line.startswith('|'):
                if not in_table:
                    in_table = True
                    tbl_rows = []
                cells = [c.strip() for c in line.strip().strip('|').split('|')]
                tbl_rows.append(cells)
                continue
            else:
                if in_table:
                    flush_table()

            if line.startswith('### '):
                h3(doc, line[4:].strip())
            elif line.startswith('## '):
                h2(doc, line[3:].strip())
            elif line.startswith('# '):
                h1(doc, line[2:].strip(), first=first_h1)
                first_h1 = False
            elif line.startswith('[TABLE]'):
                payload = line[len('[TABLE]'):].strip()
                m = re.match(r'^(.*?)\s*\{([\d\.,\s]+)\}$', payload)
                if m:
                    tbl_caption = m.group(1).strip()
                    tbl_widths = [float(x) for x in m.group(2).split(',')]
                else:
                    tbl_caption = payload
                    tbl_widths = None
                in_table = True
                tbl_rows = []
            elif line.startswith('[FIG]'):
                payload = line[len('[FIG]'):].strip()
                bits = [b.strip() for b in payload.split('|')]
                add_figure(doc, os.path.join(ASSETS, bits[0]), float(bits[1]) if len(bits) > 1 else 15.2)
            elif line.startswith('[CAPTION]'):
                add_figure_caption(doc, line[len('[CAPTION]'):].strip())
            elif line.startswith('[NOTE]'):
                note_box(doc, line[len('[NOTE]'):].strip())
            elif line.startswith('[QUOTE]'):
                p = doc.add_paragraph()
                _runs_from_markup(p, line[len('[QUOTE]'):].strip(), size=11.5, italic=True)
                p.paragraph_format.left_indent = Cm(1.0)
                p.paragraph_format.right_indent = Cm(0.6)
                p.paragraph_format.space_before = Pt(6)
                p.paragraph_format.space_after = Pt(10)
            elif line.startswith('[REF]'):
                reference_entry(doc, line[len('[REF]'):].strip())
            elif line.startswith('[PAGEBREAK]'):
                doc.add_page_break()
            elif re.match(r'^\d+\.\s', line):
                bullet(doc, line, style='List Number')
            elif line.startswith('- '):
                bullet(doc, line[2:].strip())
            elif line.startswith('> '):
                p = doc.add_paragraph()
                _runs_from_markup(p, line[2:].strip(), size=11)
                p.paragraph_format.left_indent = Cm(1.0)
                p.paragraph_format.space_after = Pt(8)
            else:
                para(doc, line)
        if in_table:
            flush_table()

    enable_update_fields(doc)
    return doc


def add_figure_caption(doc, text):
    c = doc.add_paragraph()
    r = c.add_run(text)
    r.font.name = TNR; r.font.size = Pt(9.5); r.italic = True
    c.alignment = WD_ALIGN_PARAGRAPH.LEFT
    c.paragraph_format.space_after = Pt(12)


# --------------------------------------------------------------------------
# content declarations
# --------------------------------------------------------------------------
ABSTRACT = (
    "South Africa's electricity system is unreliable, carbon-intensive and increasingly expensive, and the "
    "households least able to absorb those costs are often those living in informal settlements. National policy "
    "now frames decarbonisation as a just transition, yet most energy interventions still arrive as finished "
    "technologies to be accepted or rejected. This proposal asks what changes when residents are treated as "
    "co-designers of the socio-material arrangement around a technology rather than as recipients of it. The study "
    "develops and tests a Systemic Living Lab for Just Energy Transition (SETILL) in an informal settlement in "
    "Tzaneen, Limpopo Province. Critical posthumanism provides the theoretical base: agency is treated as distributed "
    "across people, artefacts, infrastructures, fuels and institutions, and acceptance is understood as something "
    "produced through those relations rather than as an attitude held by an individual. The framework is "
    "operationalised in a single embedded case study using the four stages of Design Research Methodology. Data will "
    "be collected through semi-structured interviews, focus groups, transect walks, household energy diaries, "
    "photovoice and co-design workshops, and analysed through reflexive thematic analysis, situational mapping and "
    "causal loop diagramming. The study is expected to contribute a tested framework for organising energy innovation "
    "in informal settlements, a justice-oriented set of evaluation criteria, and practical guidance for municipalities "
    "and technology developers. No fieldwork has been conducted; the proposal sets out an intended design, subject to "
    "ethical approval."
)
KEYWORDS = ("energy justice; just transition; living lab; critical posthumanism; informal settlements; "
            "biogas; engineering management; South Africa")

FIGURES = [
    ("Figure 1.1", "Research design: four stages of Design Research Methodology within an embedded case study"),
    ("Figure 2.1", "Energy pathways and conversion losses in a typical South African informal settlement household"),
    ("Figure 3.1", "The SETILL conceptual framework"),
    ("Figure 3.2", "Justice-oriented evaluation space for the SETILL Living Lab"),
    ("Figure 5.1", "Research timeline, 2027–2029"),
]
TABLES = [
    ("Table 1.1", "Research objectives, questions, methods and chapters"),
    ("Table 1.2", "Key terms used in this proposal"),
    ("Table 2.1", "Search strategy and inclusion criteria"),
    ("Table 2.2", "Summary of the principal literature streams"),
    ("Table 2.3", "Knowledge gaps and how the study addresses them"),
    ("Table 3.1", "Comparison of the four posthumanist thinkers used in this study"),
    ("Table 3.2", "What posthumanism is not"),
    ("Table 3.3", "Supporting perspectives: contribution, compatibility and verdict"),
    ("Table 3.4", "Ontological assumptions of TAM/UTAUT and of this study"),
    ("Table 3.5", "How the theoretical framework informs each research question"),
    ("Table 3.6", "Theoretical propositions and their observable implications"),
    ("Table 4.1", "Derivation of the SETILL components from the theoretical propositions"),
    ("Table 4.2", "The ten components of the SETILL framework"),
    ("Table 4.3", "Research questions, framework components and evidence"),
    ("Table 4.4", "Justice-oriented evaluation criteria"),
    ("Table 5.1", "Site selection criteria for the Tzaneen case"),
    ("Table 5.2", "Participant groups and sampling logic"),
    ("Table 5.3", "Data collection methods, participants and research questions"),
    ("Table 5.4", "Data analysis procedures"),
    ("Table 5.5", "Trustworthiness criteria and strategies"),
    ("Table 5.6", "Ethical risks and mitigation"),
    ("Table 5.7", "Risk register"),
    ("Table 6.1", "Tzaneen at a glance"),
    ("Table 7.1", "Expected contributions and how they will be evidenced"),
    ("Table 8.1", "Research phases, activities and milestones"),
    ("Table 9.1", "Outline of the proposed thesis"),
    ("Table A.1", "Reference verification summary"),
]
ABBREVIATIONS = [
    ["CBO", "Community-based organisation"],
    ["CLD", "Causal loop diagram"],
    ["DRM", "Design Research Methodology"],
    ["FGD", "Focus group discussion"],
    ["GHG", "Greenhouse gas"],
    ["IDP", "Integrated Development Plan"],
    ["IAP", "Indoor air pollution"],
    ["IRP", "Integrated Resource Plan"],
    ["JET", "Just Energy Transition"],
    ["JETP", "Just Energy Transition Partnership"],
    ["LPG", "Liquefied petroleum gas"],
    ["NGO", "Non-governmental organisation"],
    ["O&M", "Operation and maintenance"],
    ["PCC", "Presidential Climate Commission"],
    ["POPIA", "Protection of Personal Information Act"],
    ["REIPPPP", "Renewable Energy Independent Power Producer Procurement Programme"],
    ["RTA", "Reflexive thematic analysis"],
    ["SDG", "Sustainable Development Goal"],
    ["SETILL", "Systemic Living Lab for Just Energy Transition"],
    ["Stats SA", "Statistics South Africa"],
    ["TAM", "Technology Acceptance Model"],
    ["UJ", "University of Johannesburg"],
    ["UTAUT", "Unified Theory of Acceptance and Use of Technology"],
]


if __name__ == "__main__":
    parts = sorted(glob.glob(os.path.join(BASE, "part*.md")))
    parts = [os.path.basename(p) for p in parts]
    print("PARTS:", parts)
    doc = build_from_parts(parts)
    out = os.path.join(OUTDIR, "SETILL_PhD_Research_Proposal_UJ_REFINED.docx")
    doc.save(out)
    print("SAVED", out)
