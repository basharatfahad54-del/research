# -*- coding: utf-8 -*-
"""
Builds the FD JOHN DELIVERABLES pack, aligned to the Client Requirements
Specification (Table 4, deliverables D-01 to D-13).

Every deliverable is produced in DOCX (editable master) and PDF (read-only
rendering generated from the same DOCX with PyMuPDF), per requirement in
Section 4 of the specification.  Word COM is not used.

Outputs to:  fd john\\deliverables\\
"""
import os, re, sys, html, shutil, tempfile
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

import build_fd_john as h

ROOT = os.path.dirname(os.path.abspath(__file__))
DELIV = os.path.join(ROOT, "fd john", "deliverables")
FD = os.path.join(ROOT, "fd john")
FINAL = os.path.join(ROOT, "final_deliverables")
TMP = os.path.join(DELIV, "_tmp")
os.makedirs(TMP, exist_ok=True)

h.OUT = DELIV  # redirect helper saves into the deliverables folder

# ==========================================================================
# inline markdown -> runs
# ==========================================================================
TOKEN_RE = re.compile(r'(\*\*.+?\*\*|\*[^*\n]+?\*|`[^`]+`)')

def md_inline(s):
    """Split a markdown string into (text, bold, italic) segments."""
    out = []
    for part in TOKEN_RE.split(s):
        if not part:
            continue
        if part.startswith('**') and part.endswith('**') and len(part) > 4:
            out.append((part[2:-2], True, False))
        elif part.startswith('*') and part.endswith('*') and len(part) > 2:
            out.append((part[1:-1], False, True))
        elif part.startswith('`') and part.endswith('`') and len(part) > 2:
            out.append((part[1:-1], False, False))
        else:
            out.append((part, False, False))
    return out or [(s, False, False)]


def strip_md(s):
    return ''.join(t for t, b, i in md_inline(s))


def esc(s):
    return html.escape(s, quote=False)

# ==========================================================================
# DOCX inline paragraph with bold/italic runs
# ==========================================================================
def add_inline(doc, segments, size=12, align=None, space_after=6, style=None):
    p = doc.add_paragraph(style=style)
    for text, bold, italic in segments:
        r = p.add_run(text)
        r.font.name = h.TNR
        r.font.size = Pt(size)
        r.bold = bold
        r.italic = italic
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    p.paragraph_format.line_spacing = 1.5
    return p


def para_md(doc, text, size=12):
    return add_inline(doc, md_inline(text), size=size)

# ==========================================================================
# markdown parser -> block list
# ==========================================================================
def parse_md(path):
    with open(path, encoding='utf-8-sig') as f:
        lines = f.read().replace('\r\n', '\n').split('\n')
    blocks, i = [], 0
    n = len(lines)
    while i < n:
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        if s == '---':
            i += 1
            continue
        if s.startswith('#'):
            level = len(s) - len(s.lstrip('#'))
            blocks.append(('h%d' % min(level, 3), s.lstrip('#').strip()))
            i += 1
            continue
        if s.startswith('|'):
            rows = []
            while i < n and lines[i].strip().startswith('|'):
                cells = [c.strip() for c in lines[i].strip().strip('|').split('|')]
                if not all(re.fullmatch(r':?-+:?', c) for c in cells if c is not None):
                    rows.append(cells)
                i += 1
            width = max(len(r) for r in rows)
            rows = [r + [''] * (width - len(r)) for r in rows]
            blocks.append(('table', rows[0], rows[1:]))
            continue
        if s.startswith('- ') or s.startswith('* '):
            items = []
            while i < n and (lines[i].strip().startswith('- ') or lines[i].strip().startswith('* ')):
                items.append(lines[i].strip()[2:])
                i += 1
            blocks.append(('bullets', items))
            continue
        m = re.match(r'^(\d+)\.\s+(.*)$', s)
        if m:
            items = []
            while i < n:
                m2 = re.match(r'^(\d+)\.\s+(.*)$', lines[i].strip())
                if not m2:
                    break
                items.append((m2.group(1), m2.group(2)))
                i += 1
            blocks.append(('numbered', items))
            continue
        buf = [s]
        i += 1
        while i < n and lines[i].strip() and not re.match(r'^(#|\||- |\* |\d+\. )', lines[i].strip()):
            buf.append(lines[i].strip())
            i += 1
        blocks.append(('p', ' '.join(buf)))
    return blocks

# ==========================================================================
# plain-text parser (chapters, abstract, front matter, coverage report)
# ==========================================================================
CONSTRUCT_NAMES = {
    1: 'Configuration', 2: 'End-user agency', 3: 'Participation', 4: 'Acceptance',
    5: 'Energy justice', 6: 'Energy security', 7: 'Energy poverty',
    8: 'Living Lab apparatus', 9: 'Reflexivity', 10: 'Just transition outcome',
}

SVG_MAP = {
    '4.1': 'fig4_1_setill_framework.svg',
    '4.2': 'fig4_2_setill_loops.svg',
}


def render_svgs():
    """Render the two editable SVGs to PNG for embedding."""
    import pymupdf
    out = {}
    for k, fn in SVG_MAP.items():
        src = os.path.join(FINAL, 'assets', fn)
        d = pymupdf.open(src)
        pix = d[0].get_pixmap(matrix=pymupdf.Matrix(2, 2), alpha=False)
        png = os.path.join(TMP, fn.replace('.svg', '.png'))
        pix.save(png)
        d.close()
        out[k] = png
    return out


def fixed_width_table(header_line, body_lines):
    """Parse a fixed-width table using the header's column offsets."""
    cols = [m.start() for m in re.finditer(r'\S', header_line)]
    # merge offsets that are too close (e.g. 'Main contribution' wraps)
    offs = []
    for c in cols:
        if not offs or c - offs[-1] > 3:
            offs.append(c)
    rows = []
    cur = None
    for ln in body_lines:
        if not ln.strip():
            continue
        pieces = []
        for j in range(len(offs)):
            start = offs[j]
            end = offs[j + 1] if j + 1 < len(offs) else len(ln)
            pieces.append(ln[start:end].strip() if start < len(ln) else '')
        if pieces[0]:                      # new row
            if cur:
                rows.append(cur)
            cur = pieces
        elif cur:                          # continuation
            for j in range(1, len(cur)):
                if pieces[j]:
                    cur[j] = (cur[j] + ' ' + pieces[j]).strip()
            if pieces[0]:
                pass
    if cur:
        rows.append(cur)
    return rows


def parse_table_41(block_lines):
    """Parse TABLE 4.1 (constructs C1-C10)."""
    rows = []
    text = ' '.join(l.strip() for l in block_lines if l.strip())
    text = re.sub(r'\s+', ' ', text)
    for m in re.finditer(r'C(\d+)\s', text):
        pass
    parts = re.split(r'(?=C\d+\s)', text)
    for part in parts:
        m = re.match(r'C(\d+)\s+(.*)', part.strip())
        if not m:
            continue
        code = int(m.group(1))
        rest = m.group(2)
        status = 'Proposed' if re.search(r'(?<![\w-])Proposed(?![\w-])', rest) else 'Theory-derived'
        defn = re.sub(r'\s*Proposed\s*', ' ', rest) if status == 'Proposed' else re.sub(r'\s*Theory-derived\s*', ' ', rest)
        defn = re.sub(r'\s+', ' ', defn).strip()
        name = CONSTRUCT_NAMES.get(code, '')
        defn = re.sub(r'^%s\s*' % re.escape(name), '', defn).strip()
        rows.append(['C%d' % code, name, defn, status])
    return rows


def parse_table_32(block_lines):
    """Parse TABLE 3.2 (RQ alignment matrix, label-per-line structure)."""
    headers = ['Research question', 'Relevant theory', 'Theoretical concepts',
               'How the theory informs the investigation', 'Expected analytical contribution']
    rows = []
    cur = {}
    for ln in block_lines:
        s = ln.strip()
        if not s:
            if cur:
                rows.append([cur.get(k, '') for k in range(5)])
                cur = {}
            continue
        for k, label in enumerate(['RQ:', 'Relevant theory:', 'Theoretical concepts:',
                                   'How the theory informs the investigation:',
                                   'Expected analytical contribution:']):
            if s.startswith(label):
                cur[k] = s[len(label):].strip()
                break
        else:
            if cur and 0 in cur:
                cur[0] += ' ' + s
    if cur:
        rows.append([cur.get(k, '') for k in range(5)])
    return headers, rows


def parse_coverage(lines):
    """Parse the coverage report sections A and B."""
    rows_a, rows_b = [], []
    cur = None
    for ln in lines:
        s = ln.rstrip()
        if re.match(r'^L\d+ \|', s.strip()):
            parts = [p.strip() for p in s.strip().split('|')]
            cur = parts + ['']
            rows_a.append(cur)
        elif s.strip().startswith('10.') and '|' in s:
            parts = [p.strip() for p in s.strip().split('|')]
            rows_b.append(parts)
        elif cur is not None and s.startswith('     ') and s.strip():
            rows_a[-1][-1] = s.strip()
        elif s.strip() and not s.strip().startswith(('A.', 'B.', 'REFERENCE', '===', 'IMPORTANT', 'Generated', 'DOI-verified', 'Matrix rows', '- linked', '- no verified', 'Verified records')):
            if rows_a and cur is not None and s.startswith(' '):
                rows_a[-1][-1] = (rows_a[-1][-1] + ' ' + s.strip()).strip()
    return rows_a, rows_b


def parse_txt(path, svg_pngs=None):
    with open(path, encoding='utf-8-sig') as f:
        lines = f.read().replace('\r\n', '\n').split('\n')
    blocks, i = [], 0
    n = len(lines)
    while i < n:
        s = lines[i].strip()
        if not s:
            i += 1
            continue
        # chapter title
        if re.match(r'^CHAPTER \d+', s):
            blocks.append(('h1', s))
            i += 1
            continue
        # table captions
        m = re.match(r'^TABLE (\d+)\.(\d+):\s*(.*)$', s)
        if m:
            tid = '%s.%s' % (m.group(1), m.group(2))
            blocks.append(('tcaption', 'Table %s: %s' % (tid, m.group(3).strip())))
            i += 1
            while i < n and not lines[i].strip():
                i += 1
            blk = []
            if tid == '3.2':
                while i < n:
                    s2 = lines[i].strip()
                    if s2 and re.match(r'^(\d+\.\d+\s+[A-Z]|TABLE \d|FIGURE \d|CHAPTER \d)', s2):
                        break
                    blk.append(lines[i])
                    i += 1
            else:
                while i < n and lines[i].strip():
                    blk.append(lines[i])
                    i += 1
            if tid == '4.1':
                blocks.append(('table', ['Code', 'Construct', 'Definition', 'Status'],
                               parse_table_41(blk)))
            elif tid == '3.1':
                hdr = blk[0]
                body = blk[1:]
                blocks.append(('table', [c.strip() for c in
                                ['Theory', 'Foundation and key ideas', 'Main contribution',
                                 'Relevance here', 'Main limitation', 'Role']],
                               fixed_width_table(hdr, body)))
            elif tid == '3.2':
                hdrs, rows = parse_table_32(blk)
                blocks.append(('table', hdrs, rows))
            else:
                blocks.append(('pre', '\n'.join(blk)))
            continue
        # figure captions -> embed the rendered SVG
        m = re.match(r'^FIGURE (\d+)\.(\d+):\s*(.*)$', s)
        if m and svg_pngs:
            fid = '%s.%s' % (m.group(1), m.group(2))
            blocks.append(('fcaption', 'Figure %s: %s' % (fid, m.group(3).strip())))
            i += 1
            # skip the 'provided as editable SVG file' lines
            while i < n and lines[i].strip() and not re.match(r'^(Description|Notes)', lines[i].strip()):
                i += 1
            if fid in svg_pngs:
                blocks.append(('image', svg_pngs[fid],
                               'Editable vector source supplied in D07_Required_Diagrams (%s).' % SVG_MAP[fid]))
            continue
        # headings: numbered uppercase
        m = re.match(r'^(\d+(?:\.\d+)+)\s+([A-Z0-9].+)$', s)
        if m and s == s.upper() and len(s) < 90:
            level = 3 if m.group(1).count('.') >= 2 else 2
            blocks.append(('h%d' % level, '%s  %s' % (m.group(1), m.group(2))))
            i += 1
            continue
        m = re.match(r'^(\d+\.\d+)\s+([A-Z0-9].+)$', s)
        if m and len(s) < 90:
            blocks.append(('h2', '%s  %s' % (m.group(1), m.group(2))))
            i += 1
            continue
        # standalone uppercase headings (LIST OF FIGURES, NUMBERING NOTE, A./B.)
        if s == s.upper() and len(s) < 80 and not s.endswith('.'):
            blocks.append(('h2', s))
            i += 1
            continue
        # bracketed notes
        if s.startswith('['):
            buf = [s]
            while i < n and ']' not in lines[i]:
                i += 1
                buf.append(lines[i].strip())
            blocks.append(('note', ' '.join(buf)))
            i += 1
            continue
        # coverage-report rows
        if re.match(r'^L\d+ \|', s) or s.startswith('10.') and '|' in s:
            buf = []
            while i < n and (re.match(r'^L\d+ \|', lines[i].strip()) or
                             (lines[i].strip().startswith('10.') and '|' in lines[i]) or
                             (lines[i].startswith('     ') and lines[i].strip())):
                buf.append(lines[i].rstrip())
                i += 1
            rows_a, rows_b = parse_coverage(buf)
            if rows_a:
                blocks.append(('table', ['ID', 'Author', 'Year', 'Identifier', 'Title'], rows_a))
            if rows_b:
                blocks.append(('table', ['DOI', 'Author (year)', 'Title'], rows_b))
            continue
        # front-matter dotted list lines (Figure 4.1 ... [page])
        if re.match(r'^(Figure|Table) \d', s) and '[page]' in s:
            items = []
            while i < n and '[page]' in lines[i]:
                items.append(re.sub(r'\.{3,}\s*\[page\]', '', lines[i].strip()).strip())
                i += 1
            blocks.append(('bullets', items))
            continue
        # acronyms list lines (CBO        Community-based organisation)
        m = re.match(r'^([A-Z][A-Za-z0-9\-/]{1,10})\s{2,}(.+)$', s)
        if m and re.match(r'^[A-Z]{2,}', m.group(1)):
            items = []
            while i < n:
                m2 = re.match(r'^([A-Z][A-Za-z0-9\-/]{1,10})\s{2,}(.+)$', lines[i].strip())
                if not m2:
                    break
                items.append('%s \u2014 %s' % (m2.group(1), m2.group(2)))
                i += 1
            blocks.append(('bullets', items))
            continue
        # paragraph: join until blank
        buf = [s]
        i += 1
        while i < n and lines[i].strip():
            nxt = lines[i].strip()
            if (re.match(r'^(TABLE|FIGURE) \d+\.\d+:', nxt) or
                    re.match(r'^CHAPTER \d+', nxt) or
                    re.match(r'^\d+(\.\d+)*\s+[A-Z]', nxt) and nxt == nxt.upper() and len(nxt) < 90 or
                    nxt == nxt.upper() and len(nxt) < 80 and not nxt.endswith('.') or
                    re.match(r'^L\d+ \|', nxt) or nxt.startswith('[')):
                break
            buf.append(nxt)
            i += 1
        blocks.append(('p', ' '.join(buf)))
    return blocks

# ==========================================================================
# block emitter (docx)
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
                add_inline(doc, [('%s.  ' % num, True, False)] + md_inline(text),
                           space_after=4)
        elif kind == 'table':
            headers, rows = b[1], b[2]
            ncol = len(headers)
            norm = []
            for r in rows:
                r = [str(c) for c in r][:ncol]
                r += [''] * (ncol - len(r))
                norm.append(r)
            fsize = 8 if ncol >= 5 else 9.5
            h.table(doc, None, headers, [[strip_md(c) for c in r] for r in norm],
                    table_widths(headers, norm), fsize=fsize)
        elif kind == 'note':
            h.note_box(doc, b[1])
        elif kind == 'tcaption':
            p = doc.add_paragraph()
            r = p.add_run(b[1])
            r.font.name = h.TNR; r.font.size = Pt(9.5); r.font.bold = True
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
            p.paragraph_format.keep_with_next = True
        elif kind == 'fcaption':
            p = doc.add_paragraph()
            r = p.add_run(b[1])
            r.font.name = h.TNR; r.font.size = Pt(9.5); r.font.bold = True
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_before = Pt(10)
            p.paragraph_format.space_after = Pt(4)
        elif kind == 'image':
            p = doc.add_paragraph()
            p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            p.add_run().add_picture(b[1], width=Cm(15.4))
            note_p = doc.add_paragraph()
            r = note_p.add_run(b[2])
            r.font.name = h.TNR; r.font.size = Pt(9); r.italic = True
            note_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            note_p.paragraph_format.space_after = Pt(10)
        elif kind == 'pre':
            for ln in b[1].split('\n'):
                p = doc.add_paragraph()
                r = p.add_run(ln)
                r.font.name = 'Courier New'; r.font.size = Pt(8)
                p.paragraph_format.space_after = Pt(0)
                p.paragraph_format.line_spacing = 1.0
        elif kind == 'pagebreak':
            doc.add_page_break()


def table_widths(headers, rows):
    ncol = len(headers)
    lens = [len(str(headers[j])) for j in range(ncol)]
    for r in rows:
        for j in range(ncol):
            if j < len(r):
                lens[j] += min(len(str(r[j])), 90)
    total = sum(lens) or 1
    return [max(1.4, 15.9 * l / total) for l in lens]

# ==========================================================================
# deliverable cover helper
# ==========================================================================
def deliverable_cover(doc, did, title, subtitle, spec_status, extra=None):
    meta = [("Deliverable ID", did),
            ("Project", "PhD in Engineering Management"),
            ("Research topic", "Systemic Living Lab for Just Energy Transition (SETILL)"),
            ("Institution", "University of Johannesburg"),
            ("Date", h.DATE),
            ("Spec status", spec_status)]
    if extra:
        meta = extra + meta
    h.cover(doc, "SETILL Research Deliverable", title, subtitle, meta)


def spec_doc(did, title, subtitle, spec_status, footer, blocks=None, toc=True):
    doc = h.base_document()
    deliverable_cover(doc, did, title, subtitle, spec_status)
    h.add_footer(doc, "%s  |  %s" % (did, footer))
    if toc:
        h.add_toc(doc)
    doc.add_page_break()
    if blocks:
        emit(doc, blocks)
    return doc

# ==========================================================================
# D-02 status report
# ==========================================================================
def build_d02():
    blocks = parse_md(os.path.join(FINAL, '00_PROJECT_STATUS_REPORT.md'))
    doc = spec_doc("D-02", "Project Status Report",
                   "Overall status, workstreams, decisions and next steps",
                   "Delivered (per D-02)", "Project Status Report", blocks)
    h.save(doc, "D02_Project_Status_Report.docx",
           "Project Status Report", "SETILL status, workstreams A-L, decisions D1-D11")

# ==========================================================================
# D-03 chapter 3
# ==========================================================================
def build_d03():
    blocks = parse_txt(os.path.join(ROOT, 'proposal_v3', '03_chapter3_theoretical_framework.txt'))
    doc = spec_doc("D-03", "Chapter 3 \u2014 Theoretical Framework",
                   "Critical posthumanism and supporting perspectives, justified against the five RQs",
                   "Drafted \u2014 length decision D11 open", "Chapter 3 Theoretical Framework", blocks)
    h.save(doc, "D03_Chapter3_Theoretical_Framework.docx",
           "Chapter 3 Theoretical Framework", "Theoretical framework draft")

# ==========================================================================
# D-04 chapter 4 with embedded figures
# ==========================================================================
def build_d04(svg_pngs):
    blocks = parse_txt(os.path.join(FINAL, '08_CHAPTER4_CONCEPTUAL_FRAMEWORK.txt'), svg_pngs)
    doc = spec_doc("D-04", "Chapter 4 \u2014 Conceptual Framework: the SETILL Model",
                   "Constructs, proposed relationships, components and figures",
                   "Drafted (approx. 2,100 words)", "Chapter 4 Conceptual Framework", blocks)
    h.save(doc, "D04_Chapter4_Conceptual_Framework.docx",
           "Chapter 4 Conceptual Framework", "SETILL model draft with Figures 4.1 and 4.2")

# ==========================================================================
# D-05 abstract
# ==========================================================================
def build_d05():
    blocks = parse_txt(os.path.join(FINAL, '03_ABSTRACT.txt'))
    doc = spec_doc("D-05", "Abstract",
                   "250\u2013300 words, four-element structure required by the school",
                   "Drafted (267 words)", "Abstract", blocks, toc=False)
    h.save(doc, "D05_Abstract.docx", "Abstract", "Draft abstract, 267 words")

# ==========================================================================
# D-06 front matter
# ==========================================================================
def build_d06():
    blocks = parse_txt(os.path.join(FINAL, '04_FRONT_MATTER.txt'))
    # insert page breaks between the three lists
    out = []
    seen_lists = 0
    for b in blocks:
        if b[0] == 'h2' and b[1].upper().startswith(('LIST OF', 'NUMBERING')):
            if seen_lists:
                out.append(('pagebreak',))
            seen_lists += 1
        out.append(b)
    doc = spec_doc("D-06", "Front Matter",
                   "List of figures, list of tables, acronyms and numbering note",
                   "Delivered (working lists)", "Front Matter", out)
    h.save(doc, "D06_Front_Matter.docx", "Front Matter", "Working front matter lists")

# ==========================================================================
# D-07 figures index
# ==========================================================================
FIGURES = [
    ("fig1_design.png", "Figure 1. Research design",
     "The four-stage DRM cycle within the theoretical and conceptual lenses; embedded single case in Greater Tzaneen. Required by D-01 (methodology)."),
    ("fig2_energy.png", "Figure 2. Household energy chain",
     "Sources, conversion devices and services, with the unpaid burdens the chain does not show. Required by D-01 (problem statement and Chapter 2)."),
    ("fig3_setill.png", "Figure 3. The SETILL framework",
     "Context, four Living Lab steps, technology artefact, acceptance and outcomes under justice criteria. Required by D-04 and D-07."),
    ("fig4_justice.png", "Figure 4. Justice assessment",
     "Four dimensions at once, reported as a profile rather than a single score. Required by D-01 (evaluation) and RQ5."),
    ("fig5_timeline.png", "Figure 5. Project timeline",
     "Thirty-six-month schedule with milestones; fieldwork and analysis overlap deliberately. Required by D-01 (plan of work)."),
]


def build_d07(svg_pngs):
    doc = spec_doc("D-07", "Required Diagrams \u2014 Figures Index",
                   "The five required diagrams plus the two editable framework figures",
                   "Delivered (7 files)", "Required Diagrams Index", toc=False)
    h.para(doc, "This index shows every figure at delivery quality. The separate image and vector files are "
                "supplied alongside this document. Standards: editable source, live text, numbered caption, and "
                "notes reconciled with the narrative in the same edit (specification Section 5.6).")
    for fn, title, note in FIGURES:
        h.h2(doc, title)
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run().add_picture(os.path.join(ROOT, 'proposal', 'assets', fn), width=Cm(15.4))
        np_ = doc.add_paragraph()
        r = np_.add_run(note)
        r.font.name = h.TNR; r.font.size = Pt(9.5); r.italic = True
        np_.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        np_.paragraph_format.space_after = Pt(12)
    h.h2(doc, "Figure 4.1 \u2014 The SETILL conceptual framework (editable SVG)")
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(svg_pngs['4.1'], width=Cm(15.4))
    para_md(doc, "Reads left to right: context constructs (C1, C7, C6), intervention constructs (C8, C3, C2), "
                 "outcome constructs (C4, C5, C10), with reflexivity (C9) as a cross-cutting practice and arrows "
                 "showing the proposed relationships R1\u2013R8. The dashed boundary marks the boundary of the case "
                 "and is itself an object of critique (LL1). Supplied as fig4_1_setill_framework.svg.", size=9.5)
    h.h2(doc, "Figure 4.2 \u2014 Feedback loops in the SETILL model (editable SVG)")
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.add_run().add_picture(svg_pngs['4.2'], width=Cm(15.4))
    para_md(doc, "Shows the learning loop (LL9 \u2192 LL4 \u2192 LL3) and the justice loop (LL8 \u2192 LL6/LL5 \u2192 LL9), "
                 "with reinforcing and balancing loops labelled and reflexivity (C9) connecting the two. "
                 "Supplied as fig4_2_setill_loops.svg.", size=9.5)
    h.h2(doc, "File list")
    h.table(doc, "Table 1: Diagram files supplied with this deliverable",
            ["File", "Format", "Role"],
            [["fig1_design.png", "PNG", "Research design (DRM cycle)"],
             ["fig2_energy.png", "PNG", "Household energy chain"],
             ["fig3_setill.png", "PNG", "SETILL framework"],
             ["fig4_justice.png", "PNG", "Justice dimensions"],
             ["fig5_timeline.png", "PNG", "Project timeline"],
             ["fig4_1_setill_framework.svg", "SVG (editable)", "Conceptual framework, live text"],
             ["fig4_2_setill_loops.svg", "SVG (editable)", "Feedback loops, live text"]],
            [6.0, 3.4, 6.5], fsize=9)
    h.save(doc, os.path.join("D07_Required_Diagrams", "D07_Figures_Index.docx"),
           "Required Diagrams Index", "Index of the required diagrams")

# ==========================================================================
# D-08 references
# ==========================================================================
def parse_ris(path):
    with open(path, encoding='utf-8-sig') as f:
        text = f.read()
    records = []
    for chunk in text.split('ER  -'):
        if not chunk.strip():
            continue
        rec = {}
        for line in chunk.splitlines():
            m = re.match(r'^([A-Z][A-Z0-9])  - (.*)$', line.strip('\ufeff'))
            if m:
                tag, val = m.group(1), m.group(2).strip()
                if tag == 'AU':
                    rec.setdefault('AU', []).append(val)
                else:
                    rec[tag] = val
        if rec.get('TI'):
            records.append(rec)
    records.sort(key=lambda r: r['AU'][0].lower() if r.get('AU') else r['TI'].lower())
    return records


def ris_entry(r):
    au = ', '.join(r.get('AU', [])) or 'n.n.'
    year = r.get('PY', 'n.d.')
    ti = r.get('TI', '').rstrip('.') or 'Untitled'
    jo = r.get('JO') or r.get('T2') or ''
    bits = ['%s (%s) \u2018%s\u2019' % (au, year, ti)]
    if jo:
        bits.append(jo)
    loc = []
    if r.get('VL'):
        loc.append(r['VL'])
    if r.get('IS'):
        loc.append('(%s)' % r['IS'])
    if r.get('SP'):
        loc.append(str(r['SP']))
        if r.get('EP'):
            loc.append('\u2013%s' % r['EP'])
    if loc:
        bits.append(', ' + ''.join(loc))
    if r.get('DO'):
        bits.append('. doi: %s' % r['DO'])
    bits.append(' [%s]' % r.get('AN', 'unmapped'))
    return ''.join(bits)


def build_d08():
    doc = spec_doc("D-08", "Verified Reference List and Reference Files",
                   "34 DOI-verified records, importable into EndNote / Zotero / Mendeley",
                   "Delivered (34 verified records)", "Verified Reference List", toc=False)
    h.note_box(doc, "This list contains only records whose metadata were verified against Crossref. "
                    "It is NOT a complete bibliography: 40 of the 49 matrix rows still lack a verified "
                    "identifier (see D-09 Reference Coverage Report). No entry has been reconstructed "
                    "from memory. Each entry carries its matrix link code in square brackets.")
    h.h1(doc, "Verification evidence")
    h.para(doc, "Each record below was checked field by field: DOI resolves, author order, year, title, outlet "
                "and locator confirmed. Two DOI lookups failed (10.1086/345321 HTTP 429; "
                "10.1177/1350508410373519 HTTP 404) and are therefore absent until resolved.")
    h.h1(doc, "Verified references (alphabetical)")
    records = parse_ris(os.path.join(FINAL, '09_references.ris'))
    for r in records:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.75)
        p.paragraph_format.first_line_indent = Cm(-0.75)
        p.paragraph_format.space_after = Pt(8)
        p.paragraph_format.line_spacing = 1.5
        r = p.add_run(ris_entry(r))
        r.font.name = h.TNR; r.font.size = Pt(10.5)
    h.h1(doc, "Record count")
    h.table(doc, None, ["File", "Format", "Records", "Purpose"],
            [["EndNote_References.ris", "RIS (UTF-8 with BOM)", "34", "EndNote / Zotero / Mendeley import"],
             ["references.bib", "BibTeX", "34", "LaTeX documents"],
             ["references_endnote.xml", "EndNote XML", "34", "EndNote XML import"],
             ["Article_PDFs", "PDF", "2", "Lawfully held source articles"]],
            [5.4, 4.0, 2.0, 4.5], fsize=9)
    h.save(doc, os.path.join("D08_References", "D08_Verified_Reference_List.docx"),
           "Verified Reference List", "34 verified bibliographic records")

    # copy the data files
    for fn, dst in [('09_references.ris', 'EndNote_References.ris'),
                    ('10_references.bib', 'references.bib'),
                    ('11_references_endnote.xml', 'references_endnote.xml')]:
        shutil.copy2(os.path.join(FINAL, fn), os.path.join(DELIV, 'D08_References', dst))
    art = os.path.join(DELIV, 'D08_References', 'Article_PDFs')
    os.makedirs(art, exist_ok=True)
    pairs = [('Aboalez, K (2025)_261102_070901.pdf',
              'Aboalez_2025_Posthumanist_social_approach_to_energy_transition.pdf'),
             ('Molepo, P M, Aboalez, K, & Mathaba, T N D (2_261102_070808.pdf',
              'Molepo_Aboalez_Mathaba_2025_Barriers_to_South_Africas_energy_transition.pdf')]
    for s, d in pairs:
        sp = os.path.join(ROOT, s)
        if os.path.exists(sp):
            shutil.copy2(sp, os.path.join(art, d))
    print("COPIED reference data files and article PDFs")

# ==========================================================================
# D-09 verification and coverage
# ==========================================================================
def build_d09():
    blocks = parse_md(os.path.join(FINAL, '02_SOURCE_VERIFICATION_REPORT.md'))
    doc = spec_doc("D-09", "Source Verification Report",
                   "Verification state of every source in the 49-row matrix",
                   "Delivered", "Source Verification Report", blocks)
    h.save(doc, "D09_Source_Verification_Report.docx",
           "Source Verification Report", "V1/V2/V3 levels, 9 verified rows, 2 failed lookups")

    cov_lines = []
    with open(os.path.join(FINAL, '12_reference_coverage_report.txt'), encoding='utf-8-sig') as f:
        cov_lines = f.read().split('\n')
    # simple structure: heading block, summary lines, section A table, section B table
    blocks = [('h1', 'Reference File Coverage Report')]
    summary = []
    secA, secB = [], []
    mode = None
    for ln in cov_lines:
        s = ln.strip()
        if s.startswith('===') or not s:
            continue
        if s.startswith(('A.', 'B.')):
            mode = 'A' if s.startswith('A.') else 'B'
            blocks.append(('h2', s))
            continue
        if s.startswith(('REFERENCE FILE', 'Generated')):
            continue
        if mode is None:
            summary.append(s)
        elif mode == 'A':
            secA.append(ln)
        else:
            secB.append(ln)
    blocks.append(('p', ' '.join(summary)))
    rows_a, rows_b = parse_coverage(secA), parse_coverage(secB)
    blocks.append(('table', ['ID', 'Author', 'Year', 'Identifier', 'Title'], rows_a))
    blocks.append(('table', ['DOI', 'Author (year)', 'Title'], rows_b))
    doc = spec_doc("D-09", "Reference Coverage Report",
                   "Which matrix rows are and are not covered by the verified reference files",
                   "Delivered", "Reference Coverage Report", blocks)
    h.save(doc, "D09_Reference_Coverage_Report.docx",
           "Reference Coverage Report", "34 verified records vs 49 matrix rows")

# ==========================================================================
# D-10 QA and compliance
# ==========================================================================
def build_d10():
    for src, name, title, sub, foot in [
            ('01_CHAPTER3_COMPLIANCE_REVIEW.md', 'D10_Chapter3_Compliance_Review.docx',
             'Chapter 3 Compliance Review', "Chapter 3 checked against the school's four requirements",
             'Chapter 3 Compliance Review'),
            ('06_SCHOOL_COMPLIANCE_MATRIX.md', 'D10_School_Compliance_Matrix.docx',
             'School Compliance Matrix', "The school's 12 required sections mapped to the package",
             'School Compliance Matrix'),
            ('07_QA_AND_OUTSTANDING_ISSUES.md', 'D10_QA_and_Outstanding_Issues.docx',
             'QA and Outstanding Issues', 'QA results, defects fixed and the open-issues register',
             'QA and Outstanding Issues')]:
        blocks = parse_md(os.path.join(FINAL, src))
        doc = spec_doc("D-10", title, sub, "Delivered", foot, blocks)
        h.save(doc, name, title, sub)

# ==========================================================================
# D-11 terminology change log
# ==========================================================================
def build_d11():
    blocks = parse_md(os.path.join(FINAL, '05_TERMINOLOGY_CHANGE_LOG.md'))
    doc = spec_doc("D-11", "Terminology Change Log",
                   "Old concept-note terms mapped to the proposal's controlled vocabulary",
                   "Delivered", "Terminology Change Log", blocks)
    h.save(doc, "D11_Terminology_Change_Log.docx",
           "Terminology Change Log", "Terminology control and rationale")

# ==========================================================================
# D-12 / D-13 outstanding register
# ==========================================================================
def build_d12_d13():
    doc = spec_doc("D-12 / D-13", "Deliverables Status Register",
                   "Formal status of the two deliverables the specification lists as outstanding",
                   "Updated 16 November 2026 \u2014 D-13 working draft delivered; submission version outstanding",
                   "Deliverables Status Register")
    h.h1(doc, "1. Why this register exists")
    h.para(doc, "The Client Requirements Specification lists two deliverables as outstanding: the Chapters 1 and 2 "
                "revisions (D-12) and the complete proposal document (D-13). On 16 November 2026 the candidate "
                "instructed that the complete proposal be assembled now as a working draft. That instruction has "
                "been carried out: the assembly is delivered as D13_Complete_Proposal (DOCX and PDF, 87 pages), "
                "built strictly from the candidate's existing long-form draft (sections 3\u201312), the verified "
                "abstract and the 73-entry reference record. No new content was written for it and no research "
                "finding, reference or claim was invented. The submission-ready version of D-13 remains "
                "outstanding, for the reasons recorded below.")
    h.h1(doc, "2. Status of each deliverable")
    h.table(doc, "Table 1: Status and release conditions",
            ["ID", "Deliverable", "Status", "Release condition", "What exists today"],
            [["D-12", "Chapters 1 and 2 revisions", "Outstanding",
              "The written AI-use position (D10) and the Chapter 3 length decision (D11)",
              "Long-form sections 3\u20134 material (proposal/part1, part2) carried into the D-13 assembly; revision plan; gap analysis; five-RQ coverage map"],
             ["D-13", "Complete proposal document (sections 1\u201312)", "Working draft delivered 16 November 2026; submission version outstanding",
              "D-12 complete; D9 referencing style; D10 AI-use position; D11 length decision; reference list reconciliation",
              "D13_Complete_Proposal.docx and .pdf (87 pp): cover, abstract, sections 3\u201312, appendices, 5 figures, 27 tables, 73-entry reference list with verification record"]],
            [1.4, 3.6, 3.0, 4.0, 4.4], fsize=8)
    h.h1(doc, "3. What the D-13 working draft is, and is not")
    h.para(doc, "It is a faithful assembly of the candidate's existing draft material in the school's twelve-section "
                "order, with the AI-assisted working-draft notice, the outstanding-decision record and the reference "
                "verification appendix carried through unchanged. It is not a submission document: the school "
                "guideline prohibits AI-generated or AI-rephrased text, the candidate has not yet taken authorship "
                "of every sentence, the referencing style (D9) is undecided, and the reference list must be "
                "reconciled with the 49-row literature matrix, of which 40 rows still lack a verified identifier.")
    h.h1(doc, "4. The blocking decisions")
    h.table(doc, "Table 2: Open decisions (specification Section 8)",
            ["#", "Decision", "Owner", "Blocks"],
            [["D9", "Referencing style: Harvard or APA 7", "Faculty", "Reference list finalisation"],
             ["D10", "Written position on AI assistance", "Supervisor and faculty", "All drafting \u2014 the school guideline prohibits AI-generated or AI-rephrased text"],
             ["D11", "Chapter 3 length", "Supervisor", "Finalisation of Chapter 3 and downstream chapters"],
             ["\u2014", "40 of 49 matrix rows unverified", "Candidate", "Complete bibliography"]],
            [1.2, 6.0, 3.4, 5.3], fsize=9)
    h.h1(doc, "5. Statement of integrity")
    h.para(doc, "No research finding, reference, DOI, page number or institutional requirement has been invented "
                "anywhere in this pack. Every sentence of the D-13 working draft comes from the candidate's own "
                "workspace material. Where a document does not exist it is not included; its absence is recorded "
                "here as a matter of design, not oversight. No Turnitin or AI-detection claim is made anywhere "
                "in the pack.")
    h.save(doc, "D12_D13_Deliverables_Status_Register.docx",
           "Deliverables Status Register", "Status of D-12 and D-13")

# ==========================================================================
# D-01 copy
# ==========================================================================
def copy_d01():
    src = os.path.join(FD, 'Client_Requirements_Specification.docx')
    dst = os.path.join(DELIV, 'D01_Client_Requirements_Specification.docx')
    shutil.copy2(src, dst)
    print("COPIED D01_Client_Requirements_Specification.docx")

# ==========================================================================
# figures folder
# ==========================================================================
def copy_d07_files(svg_pngs):
    figdir = os.path.join(DELIV, 'D07_Required_Diagrams')
    os.makedirs(figdir, exist_ok=True)
    for f in FIGURES:
        shutil.copy2(os.path.join(ROOT, 'proposal', 'assets', f[0]), os.path.join(figdir, f[0]))
    for k, fn in SVG_MAP.items():
        shutil.copy2(os.path.join(FINAL, 'assets', fn), os.path.join(figdir, fn))
        shutil.copy2(svg_pngs[k], os.path.join(figdir, fn.replace('.svg', '.png')))
    print("COPIED D07 figure files")

# ==========================================================================
# PDF renderer: DOCX -> PDF with PyMuPDF (Story)
# ==========================================================================
PDIR = os.path.join(TMP, 'pdfassets')


def write_scaled_image(blob, name):
    """Store an embedded image at a sane resolution and compressed format."""
    import pymupdf
    try:
        src = pymupdf.open(stream=blob, filetype='png')
    except Exception:
        src = pymupdf.open(stream=blob)
    page = src[0]
    max_w = 1400
    if page.rect.width > max_w:
        pix = page.get_pixmap(matrix=pymupdf.Matrix(max_w / page.rect.width,
                                                    max_w / page.rect.width), alpha=False)
    else:
        pix = page.get_pixmap(alpha=False)
    path = os.path.join(PDIR, name)
    try:
        pix.save(path, jpg_quality=85)
    except TypeError:
        with open(path, 'wb') as g:
            g.write(pix.tobytes('jpeg', jpg_quality=85))
    src.close()


def recompress_pdf(doc2):
    """Downsample, JPEG-compress and deduplicate embedded images."""
    try:
        doc2.rewrite_images(dpi_threshold=200, dpi_target=150, quality=80,
                            lossy=True, lossless=False, bitonal=True,
                            color=True, gray=True)
        return True
    except Exception as e:
        print('  rewrite_images unavailable (%s); trying manual replace' % e)
        try:
            for page in doc2:
                for img in page.get_images(full=True):
                    xref = img[0]
                    pix = pymupdf.Pixmap(doc2, xref)
                    if pix.colorspace is None or pix.width < 50:
                        continue
                    if pix.colorspace.n > 3:
                        pix = pymupdf.Pixmap(pymupdf.csRGB, pix)
                    if pix.alpha:
                        pix = pymupdf.Pixmap(pix, 0)
                    page.replace_image(xref, pixmap=pix)
            return True
        except Exception as e2:
            print('  manual image replace failed:', e2)
            return False


def prep_pdf_assets():
    import pymupdf
    os.makedirs(PDIR, exist_ok=True)
    for f in ('times.ttf', 'timesbd.ttf', 'timesi.ttf', 'timesbi.ttf'):
        dst = os.path.join(PDIR, f)
        if not os.path.exists(dst) or os.path.getsize(dst) < 10000:
            make_subset_font(os.path.join(r'C:\Windows\Fonts', f), dst)
    return pymupdf.Archive(PDIR)


UNICODE_RANGES = [(0x20, 0xFF), (0x100, 0x24F), (0x2010, 0x205E), (0x20A0, 0x20BF),
                  (0x2100, 0x214F), (0x2190, 0x21FF), (0x2200, 0x22FF),
                  (0x25A0, 0x25CF), (0x2B00, 0x2BFF), (0x3010, 0x3011)]


def make_subset_font(src, dst):
    """Subset a Windows font to the ranges the pack actually uses."""
    from fontTools import ttLib
    from fontTools import subset as ftsubset
    font = ttLib.TTFont(src)
    unicodes = []
    for a, b in UNICODE_RANGES:
        unicodes.extend(range(a, b + 1))
    opts = ftsubset.Options()
    opts.name_IDs = [1, 2, 3, 4, 6]
    opts.notdef_outline = True
    ss = ftsubset.Subsetter(options=opts)
    ss.populate(unicodes=unicodes)
    ss.subset(font)
    font.save(dst)
    print('SUBSET', os.path.basename(src), '->', os.path.getsize(dst) // 1024, 'KB')


def collect_images(doc, used_names):
    """Write every embedded image of the docx into PDIR; return element walk."""
    from docx.text.paragraph import Paragraph
    from docx.table import Table
    items = []
    body = doc.element.body
    for child in body.iterchildren():
        if child.tag == qn('w:p'):
            p = Paragraph(child, doc)
            if p.text.strip() == 'Table of Contents':
                items.append(('skip_toc_head', None))
                continue
            imgs = child.findall('.//' + qn('w:drawing'))
            img_paths = []
            for dr in imgs:
                blips = dr.findall('.//' + qn('a:blip'))
                for b in blips:
                    rid = b.get(qn('r:embed'))
                    if rid and rid in doc.part.related_parts:
                        blob = doc.part.related_parts[rid].blob
                        name = 'img%03d.jpg' % used_names[0]
                        used_names[0] += 1
                        write_scaled_image(blob, name)
                        img_paths.append(name)
            items.append(('p', p, img_paths))
        elif child.tag == qn('w:tbl'):
            items.append(('tbl', Table(child, doc)))
    return items


def para_html(p, img_paths):
    txt = p.text.strip()
    if not txt and not img_paths:
        return None
    style = p.style.name if p.style is not None else 'Normal'
    align = p.alignment
    centered = (align is not None and int(align) == 1)
    cls = []
    if centered:
        cls.append('c')
    # cover detection via run sizes
    sizes = [r.font.size.pt for r in p.runs if r.font.size is not None]
    if centered and sizes and max(sizes) >= 20:
        return '<div class="ctitle">%s</div>' % esc(txt)
    if centered and any(r.italic for r in p.runs) and not any(r.bold for r in p.runs) and sizes and max(sizes) <= 13:
        return '<div class="csub">%s</div>' % esc(txt)
    if style == 'Heading 1':
        return '<h1 id="hd">%s</h1>' % esc(txt)
    if style == 'Heading 2':
        return '<h2 id="hd">%s</h2>' % esc(txt)
    if style == 'Heading 3':
        return '<h3>%s</h3>' % esc(txt)
    if style == 'List Bullet':
        return '<li>%s</li>' % inline_html(p)
    if style in ('List Number',):
        return '<li>%s</li>' % inline_html(p)
    if img_paths:
        imgs = ''.join('<br><img src="%s" width="450">' % n for n in img_paths)
        return '<p class="%s">%s%s</p>' % (' '.join(cls), inline_html(p), imgs)
    if not cls:
        return '<p>%s</p>' % inline_html(p)
    return '<p class="%s">%s</p>' % (' '.join(cls), inline_html(p))


def inline_html(p):
    out = []
    for r in p.runs:
        t = esc(r.text)
        if not t:
            continue
        if r.bold:
            t = '<b>%s</b>' % t
        if r.italic:
            t = '<i>%s</i>' % t
        out.append(t)
    # keep the TOC placeholder out of the PDF
    joined = ''.join(out)
    if joined.startswith('[Table of contents'):
        return ''
    return joined


def tbl_html(t):
    ncols = len(t.columns)
    body = []
    for i, row in enumerate(t.rows):
        cells = []
        for c in row.cells:
            txt = esc(' '.join(pp.text for pp in c.paragraphs).strip())
            cells.append('<%s>%s</%s>' % ('th' if i == 0 else 'td', txt, 'th' if i == 0 else 'td'))
        body.append('<tr>%s</tr>' % ''.join(cells))
    return '<table>%s</table>' % ''.join(body)


CSS = """
@font-face { font-family: times; src: url(times.ttf); }
@font-face { font-family: times; src: url(timesbd.ttf); font-weight: bold; }
@font-face { font-family: times; src: url(timesi.ttf); font-style: italic; }
@font-face { font-family: times; src: url(timesbi.ttf); font-weight: bold; font-style: italic; }
body { font-family: times; font-size: 11.5pt; line-height: 1.45; text-align: justify; }
h1 { font-size: 14pt; font-weight: bold; margin: 16pt 0 8pt 0; }
h2 { font-size: 12.5pt; font-weight: bold; margin: 12pt 0 6pt 0; }
h3 { font-size: 11.5pt; font-weight: bold; margin: 10pt 0 5pt 0; }
p { margin: 0 0 6pt 0; }
li { margin: 0 0 3pt 0; }
table { border-collapse: collapse; width: 100%; margin: 6pt 0; }
td, th { border: 0.5pt solid #333333; padding: 2pt 3pt; font-size: 8.5pt; line-height: 1.25; text-align: left; }
th { background-color: #E8E8E8; font-weight: bold; }
.c { text-align: center; }
.ctitle { font-size: 22pt; font-weight: bold; text-align: center; margin: 14pt 0 10pt 0; }
.csub { font-size: 12pt; font-style: italic; text-align: center; margin: 0 0 18pt 0; }
"""


def fix_tounicode(doc2):
    """Story encodes the hyphen glyph as U+00AD (soft hyphen) in the ToUnicode
    CMaps, which breaks text search and copy-paste. Patch 00AD -> 002D."""
    fixed = 0
    for x in range(1, doc2.xref_length()):
        if not doc2.xref_is_stream(x):
            continue
        try:
            raw = doc2.xref_stream(x)
        except Exception:
            continue
        try:
            text = raw.decode('latin-1')
        except Exception:
            continue
        if 'beginbfrange' not in text:
            continue
        new = re.sub(r'(<[0-9A-Fa-f]{4}>\s*<[0-9A-Fa-f]{4}>\s*)<00AD>',
                     lambda m: m.group(1) + '<002D>', text, flags=re.I)
        if new != text:
            doc2.update_stream(x, new.encode('latin-1'))
            fixed += 1
    return fixed


def docx_to_pdf(docx_path, pdf_path, label):
    import pymupdf
    used = [0]
    arch = prep_pdf_assets()
    doc = Document(docx_path)
    items = collect_images(doc, used)
    # split into page sections at explicit page breaks
    sections, cur = [], []
    for it in items:
        if it[0] == 'skip_toc_head':
            cur.append('SKIPNEXT')
            continue
        if it[0] == 'p' and it[1]._p.findall('.//' + qn('w:br')):
            if any(br.get(qn('w:type')) == 'page' for br in it[1]._p.findall('.//' + qn('w:br'))):
                sections.append(cur)
                cur = []
                continue
        cur.append(it)
    if cur:
        sections.append(cur)

    # render to a temp story file first: the destination must never be opened
    # by this process, or its Windows file handle blocks the final replace
    story_pdf = os.path.join(TMP, 'story_render.pdf')
    out = pymupdf.DocumentWriter(story_pdf)
    mediabox = pymupdf.paper_rect("a4")
    where = mediabox + (72, 72, -72, -72)
    headings = []           # (level, text, id)
    positions = []          # (id, page)
    page_no = [0]

    def rec(elpos):
        elid = getattr(elpos, 'id', None)
        if elid:
            positions.append((elid, page_no[0]))

    skipflag = [False]
    for sec in sections:
        # handle SKIPNEXT marker
        html_parts = []
        pending_skip = False
        for it in sec:
            if it == 'SKIPNEXT':
                pending_skip = True
                continue
            if it[0] == 'p':
                # skip the TOC placeholder paragraph in the PDF
                if pending_skip:
                    pending_skip = False
                    if 'Table of contents' in it[1].text or it[1].text.strip().startswith('['):
                        continue
                htm = para_html(it[1], it[2])
            elif it[0] == 'tbl':
                htm = tbl_html(it[1])
            else:
                htm = None
            if htm:
                html_parts.append(htm)
        if not html_parts:
            html_parts = ['<p></p>']
        html_doc = '<body>%s</body>' % ''.join(html_parts)
        # assign unique heading ids sequentially
        new_html, hcount = [], 0
        for part in re.split(r'(<h[12] id="hd">.*?</h[12]>)', html_doc):
            m = re.match(r'<(h[12]) id="hd">(.*?)</h[12]>', part, re.S)
            if m:
                hcount += 1
                hid = 'hd%d' % hcount
                headings.append((1 if m.group(1) == 'h1' else 2,
                                 re.sub(r'<[^>]+>', '', m.group(2)), hid))
                new_html.append('<%s id="%s">%s</%s>' % (m.group(1), hid, m.group(2), m.group(1)))
            else:
                new_html.append(part)
        html_doc = ''.join(new_html)

        story = pymupdf.Story(html=html_doc, archive=arch, user_css=CSS)
        more = 1
        while more:
            page_no[0] += 1
            dev = out.begin_page(mediabox)
            more, _ = story.place(where)
            story.element_positions(rec, {})
            story.draw(dev)
            out.end_page()
        story = None
    out.close()
    del out
    import gc
    gc.collect()

    # post-process: footers + outline
    doc2 = pymupdf.open(story_pdf)
    total = doc2.page_count
    for i, page in enumerate(doc2):
        if i == 0:
            continue
        txt = '%s   |   Page %d of %d' % (label, i + 1, total)
        page.insert_text(pymupdf.Point(297.6 - len(txt) * 2.2, 806), txt,
                         fontname='helv', fontsize=8, color=(0.27, 0.27, 0.27))
    seen = set()
    toc = []
    first_page = {}
    for elid, pg in positions:
        if elid not in first_page:
            first_page[elid] = pg
    for lvl, text, hid in headings:
        pg = first_page.get(hid)
        if pg is not None and 1 <= pg <= total:
            toc.append([lvl, text, pg])
    if toc:
        while toc and toc[0][0] > 1:
            toc.pop(0)
        fixed, prev = [], 0
        for lvl, text, pg in toc:
            lvl = min(lvl, prev + 1)
            fixed.append([lvl, text, pg])
            prev = lvl
        toc = fixed
    doc2.set_metadata({'title': os.path.basename(pdf_path), 'author': 'PhD Research Project'})
    nfix = fix_tounicode(doc2)
    recompress_pdf(doc2)
    if toc:
        doc2.set_toc(toc)
    tmp_out = pdf_path + '.tmp'
    doc2.save(tmp_out, garbage=4, deflate=True, deflate_images=True, deflate_fonts=True)
    doc2.close()
    # release MuPDF references (the loop's `page` object keeps the source file
    # handle open on Windows, which blocks os.replace below)
    page = None
    del doc2, page
    import gc
    gc.collect()
    for attempt in range(12):
        try:
            os.replace(tmp_out, pdf_path)
            break
        except OSError:
            import time
            time.sleep(2)
    else:
        raise RuntimeError('could not replace %s (locked); outline copy left at %s'
                           % (pdf_path, tmp_out))
    print("PDF  ", os.path.basename(pdf_path), 'pages=%d' % total, 'outline=%d' % len(toc),
          '%.1f MB' % (os.path.getsize(pdf_path) / 1048576))


def convert_all():
    jobs = []
    for dirpath, _, files in os.walk(DELIV):
        for f in sorted(files):
            if f.endswith('.docx') and not f.startswith('~'):
                rel = os.path.relpath(dirpath, DELIV)
                docx = os.path.join(dirpath, f)
                pdf = os.path.join(dirpath, f[:-5] + '.pdf')
                label = f[:-5]
                jobs.append((docx, pdf, label))
    # import pymupdf once
    global pymupdf
    import pymupdf
    for docx, pdf, label in jobs:
        docx_to_pdf(docx, pdf, label)

# ==========================================================================
# manifest and readme
# ==========================================================================
def build_manifest():
    rows = []
    for dirpath, _, files in os.walk(DELIV):
        for f in sorted(files):
            if f.startswith('~') or '_tmp' in dirpath:
                continue
            rel = os.path.relpath(os.path.join(dirpath, f), DELIV)
            rows.append('| %s | %s KB |' % (rel, os.path.getsize(os.path.join(dirpath, f)) // 1024))
    man = """# MANIFEST - FD JOHN DELIVERABLES PACK

**Project:** SETILL - Systemic Living Lab for Just Energy Transition
**Institution:** University of Johannesburg, Engineering Management
**Date:** %s
**Control:** Client Requirements Specification (D-01). Every deliverable below is
produced in DOCX (editable master) and PDF (read-only rendering generated from
the same DOCX), per Section 4 of the specification.

## File index

| File | Size |
|---|---|
%s

## Notes

- PDFs are generated from the DOCX masters with PyMuPDF; content is identical
  at delivery. PDFs carry clickable outline bookmarks in place of the Word
  auto-TOC field, which is present in the DOCX files.
- D-12 and D-13 are outstanding by design (see the Outstanding Deliverables
  Register). Nothing in this pack simulates work that does not exist.
- Reference files contain 34 verified records only. 40 matrix rows remain
  unverified (see D-09 coverage report).
- Figures are supplied as PNG plus editable SVG with live text (D-07).
""" % (h.DATE, '\n'.join(rows))
    with open(os.path.join(DELIV, 'MANIFEST.md'), 'w', encoding='utf-8', newline='\r\n') as f:
        f.write(man)
    print("WROTE MANIFEST.md")

# ==========================================================================
# main
# ==========================================================================
if __name__ == '__main__':
    global pymupdf
    import pymupdf  # noqa: F811
    svg_pngs = render_svgs()
    copy_d01()
    build_d02()
    build_d03()
    build_d04(svg_pngs)
    build_d05()
    build_d06()
    build_d07(svg_pngs)
    build_d08()
    build_d09()
    build_d10()
    build_d11()
    build_d12_d13()
    copy_d07_files(svg_pngs)
    build_manifest()
    convert_all()
    print('\nDONE')
