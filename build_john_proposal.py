# -*- coding: utf-8 -*-
"""
Builds D-13: the complete SETILL research proposal document (school sections
1-12 + appendices) for John, from the long-form draft (proposal/part1-part6).

DOCX first (editable master), then PDF via the pack's PyMuPDF renderer.

Integrity: this is an assembly of the candidate's existing draft material.
Nothing is invented: every paragraph, table, figure and reference comes from
the workspace sources. The document carries the AI-assisted working-draft
notice and the outstanding-decision record.
"""
import os, re

import build_fd_john as h
import build_fd_john_pack as pack   # sets h.OUT to the deliverables folder
from build_fd_john_pack import (md_inline, strip_md, para_md, add_inline,
                                table_widths, esc)
from docx.shared import Pt, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH

ROOT = pack.ROOT
DELIV = pack.DELIV
ASSETS = os.path.join(ROOT, 'proposal', 'assets')
PARTS = ['part1_intro.md', 'part2_lit.md', 'part3_theory.md',
         'part4_concept.md', 'part5_method.md', 'part6_end.md']

# ==========================================================================
# parser for the long-form markdown parts
# ==========================================================================
def parse_part(path):
    with open(path, encoding='utf-8-sig') as f:
        lines = f.read().replace('\r\n', '\n').split('\n')
    blocks, i, n = [], 0, len(lines)
    while i < n:
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        if s.startswith('### '):
            blocks.append(('h3', s[4:].strip())); i += 1; continue
        if s.startswith('## '):
            blocks.append(('h2', s[3:].strip())); i += 1; continue
        if s.startswith('# '):
            blocks.append(('h1', s[2:].strip())); i += 1; continue
        m = re.match(r'^\[TABLE\]\s*(.*)$', s)
        if m:
            cap = m.group(1).strip()
            widths = None
            mw = re.match(r'^(.*?)\s*\{([\d.,\s]+)\}$', cap)
            if mw:
                cap = mw.group(1).strip()
                widths = [float(x) for x in mw.group(2).split(',')]
            i += 1
            rows = []
            while i < n and lines[i].strip().startswith('|'):
                cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
                if not all(re.fullmatch(r':?-+:?', c) for c in cells):
                    rows.append(cells)
                i += 1
            width = max(len(r) for r in rows)
            rows = [r + [''] * (width - len(r)) for r in rows]
            blocks.append(('table', cap, rows, widths))
            continue
        m = re.match(r'^\[FIG\]\s*(\S+)\s*\|\s*([\d.]+)', s)
        if m:
            fn, w = m.group(1), float(m.group(2))
            i += 1
            caption = None
            if i < n and lines[i].strip().startswith('[CAPTION]'):
                caption = lines[i].strip()[len('[CAPTION]'):].strip()
                i += 1
            blocks.append(('fig', fn, caption, w))
            continue
        if s.startswith('[NOTE]'):
            buf = [s[len('[NOTE]'):].strip()]
            i += 1
            while i < n and lines[i].strip() and not lines[i].strip().startswith('['):
                buf.append(lines[i].strip())
                i += 1
            blocks.append(('note', ' '.join(buf)))
            continue
        if s.startswith('[REF]'):
            refs = []
            while i < n:
                s2 = lines[i].strip()
                if s2.startswith('[REF]'):
                    refs.append(s2[len('[REF]'):].strip())
                    i += 1
                elif not s2:
                    i += 1
                else:
                    break
            blocks.append(('refs', refs))
            continue
        if re.match(r'^\d+\.\s+\S', s):
            items = []
            while i < n:
                m2 = re.match(r'^(\d+)\.\s+(.*)$', lines[i].strip())
                if m2:
                    num, text = m2.group(1), m2.group(2)
                    i += 1
                    while i < n and lines[i].strip() and not re.match(
                            r'^(\d+\.\s|\[|#|\||- |\* )', lines[i].strip()):
                        text += ' ' + lines[i].strip()
                        i += 1
                    items.append((num, text))
                else:
                    break
            blocks.append(('numbered', items))
            continue
        if s.startswith('- ') or s.startswith('* '):
            items = []
            while i < n and (lines[i].strip().startswith('- ') or
                             lines[i].strip().startswith('* ')):
                items.append(lines[i].strip()[2:])
                i += 1
            blocks.append(('bullets', items))
            continue
        # bracketed meta note (word count etc.)
        if s.startswith('['):
            buf = [s]
            i += 1
            while i < n and ']' not in lines[i]:
                buf.append(lines[i].strip())
                i += 1
            if i < n:
                buf.append(lines[i].strip())
                i += 1
            blocks.append(('smallnote', ' '.join(buf)))
            continue
        buf = [s]
        i += 1
        while i < n and lines[i].strip() and not re.match(
                r'^(#|\[|- |\* |\d+\.\s)', lines[i].strip()):
            buf.append(lines[i].strip())
            i += 1
        blocks.append(('p', ' '.join(buf)))
    return blocks

# ==========================================================================
# DOCX emitter
# ==========================================================================
def emit(doc, blocks):
    for b in blocks:
        kind = b[0]
        if kind == 'h1':
            h.h1(doc, b[1])
        elif kind == 'h2':
            h.h2(doc, b[1])
        elif kind == 'h3':
            doc.add_paragraph(b[1], style='Heading 3')
        elif kind == 'p':
            para_md(doc, b[1])
        elif kind == 'bullets':
            h.bullets(doc, [strip_md(x) for x in b[1]])
        elif kind == 'numbered':
            for num, text in b[1]:
                add_inline(doc, md_inline(text), size=12, space_after=6)
                # keep the draft's own "1." numbering text
        elif kind == 'table':
            cap, rows, widths = b[1], b[2], b[3]
            ncol = len(rows[0])
            fsize = 8 if ncol >= 5 else 9
            w = widths if widths and len(widths) == ncol else table_widths(rows[0], rows[1:])
            h.table(doc, cap, rows[0], [[strip_md(str(c)) for c in r] for r in rows[1:]],
                    w, fsize=fsize)
        elif kind == 'fig':
            fn, caption, w = b[1], b[2], b[3]
            path = os.path.join(ASSETS, fn)
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            p.add_run().add_picture(path, width=Cm(min(w, 15.9)))
            if caption:
                cp = doc.add_paragraph()
                r = cp.add_run(strip_md(caption))
                r.font.name = h.TNR; r.font.size = Pt(9.5); r.font.bold = True
                cp.alignment = WD_ALIGN_PARAGRAPH.CENTER
                cp.paragraph_format.space_after = Pt(10)
        elif kind == 'note':
            h.note_box(doc, strip_md(b[1]))
        elif kind == 'smallnote':
            p = doc.add_paragraph()
            r = p.add_run(b[1])
            r.font.name = h.TNR; r.font.size = Pt(10); r.italic = True
            p.paragraph_format.space_after = Pt(8)
        elif kind == 'refs':
            for ref in b[1]:
                p = doc.add_paragraph()
                p.paragraph_format.left_indent = Cm(0.75)
                p.paragraph_format.first_line_indent = Cm(-0.75)
                p.paragraph_format.space_after = Pt(6)
                p.paragraph_format.line_spacing = 1.5
                for text, bold, italic in md_inline(ref):
                    r = p.add_run(text)
                    r.font.name = h.TNR; r.font.size = Pt(11)
                    r.bold = bold; r.italic = italic

# ==========================================================================
# abstract section (section 2) from the verified abstract file
# ==========================================================================
def abstract_blocks():
    with open(os.path.join(pack.FINAL, '03_ABSTRACT.txt'), encoding='utf-8-sig') as f:
        lines = [l for l in f.read().replace('\r\n', '\n').split('\n')]
    paras, buf = [], []
    note = None
    for s in lines:
        s = s.strip()
        if not s:
            if buf:
                paras.append(' '.join(buf)); buf = []
            continue
        if s == 'ABSTRACT':
            continue
        if s.startswith('['):
            note = s.lstrip('[').rstrip(']')
            continue
        buf.append(s)
    if buf:
        paras.append(' '.join(buf))
    return paras, note

# ==========================================================================
# build
# ==========================================================================
def build():
    doc = h.base_document()

    # --- Section 1: cover page -------------------------------------------
    h.cover(doc,
            "PhD Research Proposal",
            "A Systemic Living Lab for Just Energy Transition (SETILL)",
            "Co-designing responsive energy technology innovations with "
            "end-users in a South African informal settlement",
            [("Degree", "Doctor of Philosophy in Engineering Management"),
             ("Institution", "University of Johannesburg"),
             ("Candidate", "John"),
             ("Date", h.DATE)])
    h.add_footer(doc, "SETILL Research Proposal")
    doc.add_page_break()

    # --- status notice ----------------------------------------------------
    h.h1(doc, "Status of this document")
    h.note_box(doc,
               "AI-assisted working draft. The school guideline prohibits AI tools "
               "for generating or rephrasing text, and decision D10 (written "
               "supervisor/faculty position on AI assistance) is outstanding. This "
               "assembly was produced at the candidate's instruction of "
               "16 November 2026 from the candidate's existing long-form draft "
               "(sections 3-12), the verified abstract and the reference record. "
               "It is not submission-ready until the candidate has revised it, "
               "taken authorship of it, and the open decisions D9 (referencing "
               "style), D10 (AI-use position) and D11 (Chapter 3 length) are "
               "closed. No fieldwork has been conducted; the proposal stage only.")
    doc.add_page_break()

    # --- table of contents -------------------------------------------------
    h.add_toc(doc)
    doc.add_page_break()

    # --- Section 2: abstract ------------------------------------------------
    h.h1(doc, "2. Abstract")
    paras, note = abstract_blocks()
    for p_ in paras:
        h.para(doc, p_)
    if note:
        p = doc.add_paragraph()
        r = p.add_run(note)
        r.font.name = h.TNR; r.font.size = Pt(10); r.italic = True
        p.paragraph_format.space_before = Pt(8)
    doc.add_page_break()

    # --- Sections 3-12 + appendices ----------------------------------------
    for fn in PARTS:
        blocks = parse_part(os.path.join(ROOT, 'proposal', fn))
        emit(doc, blocks)

    h.save(doc, "D13_Complete_Proposal.docx",
           "SETILL Research Proposal - Complete Draft",
           "Complete proposal, school sections 1-12 with appendices")
    print("WROTE D13_Complete_Proposal.docx")

def to_pdf():
    docx = os.path.join(DELIV, "D13_Complete_Proposal.docx")
    pdf = os.path.join(DELIV, "D13_Complete_Proposal.pdf")
    pack.docx_to_pdf(docx, pdf, "D13_Complete_Proposal")

if __name__ == '__main__':
    build()
    to_pdf()
