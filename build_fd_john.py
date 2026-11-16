# -*- coding: utf-8 -*-
"""
Builds the FD JOHN client-requirements pack:
  fd john\\Client_Requirements_Specification.docx
  fd john\\figures\\            (the five required diagrams + two editable SVGs)

Academic document generator: A4, Times New Roman 12, 1.5 spacing, justified,
professional headings, page fields, automatic TOC, numbered tables, figures.
"""
import os, shutil
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(ROOT, "fd john")
ASSETS = os.path.join(ROOT, "proposal", "assets")
SVGDIR = os.path.join(ROOT, "final_deliverables", "assets")
FIGDIR = os.path.join(OUT, "figures")

TNR = "Times New Roman"
DATE = "16 November 2026"
TITLE = ("Developing an Innovative Systemic Living Lab for Just Energy "
         "Transition in a South African Informal Settlement (SETILL)")

# --------------------------------------------------------------------------
# helpers (same conventions as build_deliverables.py)
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
        s.left_margin = Cm(2.54)
        s.right_margin = Cm(2.54)
    n = doc.styles['Normal']
    style_font(n, 12)
    pf = n.paragraph_format
    pf.line_spacing = 1.5
    pf.space_after = Pt(6)
    pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    for name, size, bold in (('Heading 1', 14, True), ('Heading 2', 13, True), ('Heading 3', 12, True)):
        st = doc.styles[name]
        style_font(st, size, bold=bold, italic=False)
        st.paragraph_format.space_before = Pt(18 if name == 'Heading 1' else 12)
        st.paragraph_format.space_after = Pt(8 if name == 'Heading 1' else 6)
        st.paragraph_format.keep_with_next = True
        st.paragraph_format.line_spacing = 1.5
    return doc


def _field(paragraph, instr):
    run = paragraph.add_run()
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve'); it.text = instr
    f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t'); t.text = "1"
    f3 = OxmlElement('w:fldChar'); f3.set(qn('w:fldCharType'), 'end')
    r = run._r
    r.append(f1); r.append(it); r.append(f2); r.append(t); r.append(f3)


def add_footer(doc, label):
    sec = doc.sections[0]
    sec.different_first_page_header_footer = True
    p = sec.footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(label + "   |   Page ")
    _field(p, ' PAGE ')
    r2 = p.add_run(" of ")
    _field(p, ' NUMPAGES ')
    for run in p.runs:
        run.font.name = TNR; run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x44, 0x44, 0x44)


def add_toc(doc):
    h = doc.add_paragraph()
    r = h.add_run("Table of Contents")
    r.font.name = TNR; r.font.size = Pt(14); r.font.bold = True
    h.paragraph_format.space_after = Pt(12)
    p = doc.add_paragraph()
    run = p.add_run()
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve')
    it.text = 'TOC \\o "1-2" \\h \\z \\u'
    f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t'); t.text = ("[Table of contents is generated automatically when the "
                                      "document is opened. Press Ctrl+A then F9 to refresh.]")
    f3 = OxmlElement('w:fldChar'); f3.set(qn('w:fldCharType'), 'end')
    r = run._r
    r.append(f1); r.append(it); r.append(f2); r.append(t); r.append(f3)


def cover(doc, kicker, title, subtitle, meta_pairs):
    for _ in range(3):
        doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(kicker.upper()); r.font.name = TNR; r.font.size = Pt(10.5); r.font.bold = True
    r.font.color.rgb = RGBColor(0x55, 0x55, 0x55)
    p.paragraph_format.space_after = Pt(18)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(title); r.font.name = TNR; r.font.size = Pt(22); r.font.bold = True
    p.paragraph_format.space_after = Pt(10)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run(subtitle); r.font.name = TNR; r.font.size = Pt(12.5); r.font.italic = True
    r.font.color.rgb = RGBColor(0x33, 0x33, 0x33)
    p.paragraph_format.space_after = Pt(30)
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    pPr = p._p.get_or_add_pPr()
    bd = OxmlElement('w:pBdr'); bo = OxmlElement('w:bottom')
    bo.set(qn('w:val'), 'single'); bo.set(qn('w:sz'), '6'); bo.set(qn('w:space'), '1'); bo.set(qn('w:color'), '999999')
    bd.append(bo); pPr.append(bd)
    p.paragraph_format.space_after = Pt(24)
    doc.add_paragraph()
    for k, v in meta_pairs:
        p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.space_after = Pt(2)
        p.paragraph_format.line_spacing = 1.15
        r = p.add_run(k + ":  "); r.font.name = TNR; r.font.size = Pt(10.5); r.font.bold = True
        r2 = p.add_run(v); r2.font.name = TNR; r2.font.size = Pt(10.5)
    for _ in range(4):
        doc.add_paragraph()
    p = doc.add_paragraph(); p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run("University of Johannesburg  |  Engineering Management")
    r.font.name = TNR; r.font.size = Pt(9.5); r.font.color.rgb = RGBColor(0x66, 0x66, 0x66)
    doc.add_page_break()


def h1(doc, text):
    return doc.add_paragraph(text, style='Heading 1')


def h2(doc, text):
    return doc.add_paragraph(text, style='Heading 2')


def para(doc, text, italic=False, bold=False, size=12, space_after=6):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.name = TNR; r.font.size = Pt(size); r.italic = italic; r.bold = bold
    p.paragraph_format.space_after = Pt(space_after)
    return p


def bullets(doc, items):
    for it in items:
        p = doc.add_paragraph(style='List Bullet')
        r = p.add_run(it)
        r.font.name = TNR; r.font.size = Pt(12)
        p.paragraph_format.space_after = Pt(3)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY


def _shade(cell, hexfill):
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear'); shd.set(qn('w:color'), 'auto'); shd.set(qn('w:fill'), hexfill)
    tcPr.append(shd)


def _repeat_header(row):
    trPr = row._tr.get_or_add_trPr()
    h = OxmlElement('w:tblHeader'); h.set(qn('w:val'), 'true')
    trPr.append(h)


def table(doc, caption, headers, rows, widths, fsize=9.5):
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
    hdr = t.rows[0]
    _repeat_header(hdr)
    for i, htxt in enumerate(headers):
        c = hdr.cells[i]
        c.width = Cm(widths[i])
        c.text = ''
        pp = c.paragraphs[0]
        pp.paragraph_format.space_after = Pt(2)
        pp.paragraph_format.line_spacing = 1.0
        rr = pp.add_run(htxt)
        rr.font.name = TNR; rr.font.size = Pt(fsize); rr.bold = True
        _shade(c, 'E8E8E8')
    for row in rows:
        cells = t.add_row().cells
        for i, val in enumerate(row):
            c = cells[i]
            c.width = Cm(widths[i])
            c.text = ''
            for j, part in enumerate(str(val).split('\n')):
                pp = c.paragraphs[0] if j == 0 else c.add_paragraph()
                pp.paragraph_format.space_after = Pt(1)
                pp.paragraph_format.line_spacing = 1.0
                rr = pp.add_run(part)
                rr.font.name = TNR; rr.font.size = Pt(fsize)
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def note_box(doc, text):
    t = doc.add_table(rows=1, cols=1)
    t.style = 'Table Grid'
    t.autofit = False
    c = t.rows[0].cells[0]
    c.width = Cm(15.9)
    _shade(c, 'F5F5F5')
    c.text = ''
    p = c.paragraphs[0]
    p.paragraph_format.space_before = Pt(4)
    p.paragraph_format.space_after = Pt(4)
    p.paragraph_format.line_spacing = 1.3
    r = p.add_run(text)
    r.font.name = TNR; r.font.size = Pt(10); r.italic = True
    doc.add_paragraph().paragraph_format.space_after = Pt(2)


def figure(doc, png, caption, note, width_cm=15.9):
    path = os.path.join(ASSETS, png)
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.paragraph_format.keep_with_next = True
    p.add_run().add_picture(path, width=Cm(width_cm))
    c = doc.add_paragraph()
    c.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = c.add_run(caption)
    r.font.name = TNR; r.font.size = Pt(9.5); r.font.bold = True
    c.paragraph_format.space_after = Pt(4)
    n = doc.add_paragraph()
    n.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    r2 = n.add_run(note)
    r2.font.name = TNR; r2.font.size = Pt(9.5); r2.italic = True
    n.paragraph_format.space_after = Pt(12)


def save(doc, relpath, title, subject):
    path = os.path.join(OUT, relpath)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    cp = doc.core_properties
    cp.title = title
    cp.subject = subject
    cp.author = "PhD Research Project"
    doc.save(path)
    print("WROTE", relpath)


# ==========================================================================
# DOC - CLIENT REQUIREMENTS SPECIFICATION
# ==========================================================================
def build_requirements():
    doc = base_document()
    cover(doc, "Client Requirements", "Client Requirements Specification",
          "Deliverables, diagrams and acceptance criteria for the SETILL research project",
          [("Project", "PhD in Engineering Management"),
            ("Research topic", "Systemic Living Lab for Just Energy Transition (SETILL)"),
            ("Institution", "University of Johannesburg"),
            ("Supervisor", "Dr K. Aboalez"),
            ("Date", DATE),
            ("Document status", "Issued for client confirmation")])
    add_footer(doc, "Client Requirements Specification")
    add_toc(doc)
    doc.add_page_break()

    # ---- 1. Purpose -----------------------------------------------------
    h1(doc, "1. Purpose and scope")
    para(doc, "This specification records, in one place, what the client requires of the SETILL research project: "
              "the scholarly requirements set by the research brief, the deliverables to be produced and in which "
              "formats, the diagrams each deliverable must carry, the integrity rules that apply to every document, "
              "and the criteria against which each deliverable will be accepted. It is a control document: the "
              "deliverables are checked against it, and changes to it are recorded.")
    para(doc, "The scope is the full project: the proposal document (current stage) and, in due course, the "
              "examination-ready thesis. No fieldwork has been conducted; every method in every deliverable is "
              "written as a proposed method.")
    note_box(doc, "Status of the project: proposal only. Nothing in this specification, and nothing in any "
                  "deliverable governed by it, may claim that data were collected, that participants were "
                  "recruited, or that a prototype was built, until those events have actually occurred.")

    # ---- 2. Project overview -------------------------------------------
    h1(doc, "2. Project overview")
    table(doc, "Table 1: Binding project parameters (decisions D1\u2013D8)",
          ["Parameter", "Requirement"],
          [["Institution", "University of Johannesburg, Engineering Management"],
           ["Supervisor", "Dr K. Aboalez"],
           ["Research title", "Developing an Innovative Systemic Living Lab for Just Energy Transition in a South African Informal Settlement"],
           ["Acronym", "SETILL \u2014 Systemic Living Lab for Just Energy Transition (brief's expansion; one footnote records the older expansion)"],
           ["Case", "Tzaneen, Limpopo, South Africa only. The former Nigerian case material is out of scope"],
           ["Technology focus", "Community-scale biogas. Wind and solar IPP material is out of scope"],
           ["Research paradigm", "Posthumanist-informed, with a constructivist-interpretivist epistemology to be reconciled explicitly in the methodology"],
           ["Primary theory", "Critical posthumanism (Braidotti, Barad, Haraway, Hayles), including the supervisor's potestas/potentia distinction and posthumanist cartography"],
           ["Supporting perspectives", "Systems thinking, energy justice, socio-technical transitions, Living Lab theory, complexity theory"],
           ["Rejected as core", "Technology Acceptance Model and Lean Thinking (used as contrasts only, and represented fairly)"],
           ["Document status", "Proposal. The 199-page concept note is a draft to be revised, not a completed study"],
           ["Referencing", "EndNote temporary citations {Author, Year #ID}; final style pending decision D9"]],
          [4.6, 11.3], fsize=9)

    # ---- 3. Research requirements --------------------------------------
    h1(doc, "3. Research requirements")
    h2(doc, "3.1 The five briefed research questions")
    para(doc, "The brief specifies five end-user-centred research questions. Every deliverable must be traceable to "
              "at least one of them; content that answers none of them is out of scope.")
    table(doc, "Table 2: Research questions and where they must be answered",
          ["#", "Research question (as briefed)", "Primary chapter"],
          [["RQ1", "How Systems Living Lab concepts frame the relationship between inequality, energy security, energy poverty, energy technology innovation and informal settlements", "Chapters 2 and 3"],
           ["RQ2", "Factors determining end-user involvement in decision-making, co-design, co-creation, co-research, prototyping and testing", "Chapters 3, 4 and 5"],
           ["RQ3", "Cultural elements, norms, practices and technological mechanisms influencing acceptance", "Chapters 3 and 4"],
           ["RQ4", "Mechanisms to improve acceptance and end-user contribution to energy security", "Chapters 4 and 6"],
           ["RQ5", "Whether the proposed Systems Living Lab can engage end-users and stakeholders in real-world settings", "Chapters 5 and 7"]],
          [1.2, 10.6, 4.1], fsize=9)

    h2(doc, "3.2 Thesis architecture (Option A)")
    para(doc, "The seven-chapter architecture in Table 3 is binding. It separates the theoretical framework from "
              "the conceptual framework, which the brief requires and which keeps the mandated theory visible to "
              "an examiner.")
    table(doc, "Table 3: Chapter architecture and indicative word budget",
          ["Chapter", "Content", "Target words"],
          [["1", "Introduction: problem, aim, objectives, five RQs, justification, scope", "8,000"],
           ["2", "Literature review organised around RQ1\u2013RQ5, ending in an explicit gap statement", "19,000"],
           ["3", "Theoretical framework: critical posthumanism and supporting perspectives (new chapter)", "18,000\u201322,000 (D11 open)"],
           ["4", "Conceptual framework: the SETILL model re-derived from Chapter 3", "9,000"],
           ["5", "Methodology: proposed design, single case, instruments aligned to the five RQs", "12,000"],
           ["6", "Results or proposed analysis strategy (requires supervisor decision D12)", "6,000"],
           ["7", "Discussion and conclusion with separate theoretical, methodological and practical contributions", "7,000"]],
          [1.6, 10.7, 3.6], fsize=9)

    h2(doc, "3.3 School structure compliance")
    para(doc, "The proposal must follow the school's proposal-structure note: twelve numbered sections from "
              "Introduction (3) to References (12), with an abstract of 250\u2013300 words covering problem, "
              "objectives, methodology and contributions. The compliance matrix in the final-deliverables package "
              "maps each school section to the work that satisfies it, and is kept current.")

    # ---- 4. Deliverable requirements ------------------------------------
    h1(doc, "4. Deliverable requirements")
    para(doc, "Table 4 is the deliverable schedule. Every substantial document is supplied twice: a Word file as "
              "the editable master and a PDF as a read-only rendering generated from the same source. The two are "
              "identical at delivery; a PDF is regenerated after any edit to its master. Where a deliverable is "
              "marked outstanding, it is absent by design and its dependency is stated.")
    table(doc, "Table 4: Deliverable schedule and acceptance criteria",
          ["ID", "Deliverable", "Formats", "Acceptance criteria", "Status"],
          [["D-01", "Client Requirements Specification (this document)", "DOCX, PDF", "All requirements stated and traceable; no invented facts", "Delivered"],
           ["D-02", "Project status report", "DOCX, PDF", "States workstreams, decisions D1\u2013D11 and open items accurately", "Delivered"],
           ["D-03", "Chapter 3 \u2014 Theoretical Framework", "DOCX, PDF", "Covers the school's four requirements; justified against all five RQs; built from primary theory texts", "Drafted; length decision D11 open"],
           ["D-04", "Chapter 4 \u2014 Conceptual Framework (SETILL model)", "DOCX, PDF", "Every component derived from Chapter 3; construct codes do not collide with model components", "Drafted (approx. 2,100 words)"],
           ["D-05", "Abstract", "DOCX, PDF", "250\u2013300 words; four-element school structure; no findings claimed", "Drafted (267 words)"],
           ["D-06", "Front matter", "DOCX, PDF", "List of figures, list of tables and list of acronyms on separate pages, numbered to match the chapters", "Delivered (working lists)"],
           ["D-07", "Required diagrams (Section 5)", "PNG + editable SVG", "Vector source, live text, numbered captions, reconciled with the narrative", "Delivered (7 files)"],
           ["D-08", "Reference files", "RIS, BibTeX, EndNote XML", "Only verified records (34); every DOI resolves; no fabricated entries", "Delivered"],
           ["D-09", "Reference verification and coverage reports", "DOCX, PDF", "Each record checked field by field; unverified rows listed, not guessed", "Delivered"],
           ["D-10", "Quality-assurance and compliance reports", "DOCX, PDF", "Compliance matrix maps the school's 12 sections; QA issues recorded, not hidden", "Delivered"],
           ["D-11", "Terminology change log", "DOCX, PDF", "Every old-to-new term change recorded with its justification", "Delivered"],
           ["D-12", "Chapters 1 and 2 revisions", "DOCX, PDF", "Aligned to the five RQs and the word budget; gap statement explicit", "Outstanding \u2014 depends on D-03 finalisation"],
           ["D-13", "Complete proposal document (sections 3\u201312)", "DOCX, PDF", "Single numbering system; full front matter; complete verified reference list", "Outstanding \u2014 depends on D9, D10, D-12"]],
          [1.3, 4.2, 2.3, 5.5, 2.6], fsize=8)

    # ---- 5. Required diagrams -------------------------------------------
    h1(doc, "5. Required diagrams")
    para(doc, "Five diagrams are required, one for each load-bearing part of the project. All five are reproduced "
              "below at delivery quality and are supplied as separate files in the figures folder. The standard "
              "for every diagram is: an editable source (vector or regeneration script), live text rather than "
              "embedded bitmaps, a numbered caption, and notes reconciled with the narrative in the same edit, so "
              "that figure and text do not drift apart.")

    h2(doc, "5.1 Research design")
    figure(doc, "fig1_design.png",
           "Figure 1. Research design: the four-stage DRM cycle within the theoretical and conceptual lenses.",
           "Required by D-01 (methodology). Shows how the clarifying, descriptive and prescriptive stages "
           "generate and revise the SETILL model, and locates the embedded single case in Greater Tzaneen with "
           "its four units of analysis. The orange return path expresses the iterative revision requirement: "
           "findings feed back into the model and the next co-design cycle.")

    h2(doc, "5.2 The household energy problem")
    figure(doc, "fig2_energy.png",
           "Figure 2. Household energy chain: sources, conversion devices and services, with the unpaid burdens the chain does not show.",
           "Required by D-01 (problem statement and Chapter 2). Establishes fuel stacking as normal practice and "
           "states, in the figure itself, why a purely technical efficiency account is insufficient \u2014 the "
           "justification for a justice-centred rather than device-centred study.")

    h2(doc, "5.3 The SETILL framework")
    figure(doc, "fig3_setill.png",
           "Figure 3. The SETILL framework: context, four Living Lab steps, technology artefact, acceptance and outcomes under justice criteria.",
           "Required by D-04 and D-07. The central deliverable diagram. It must show the stakeholder set, the "
           "four co-diagnose\u2013co-design\u2013prototype\u2013evaluate steps, acceptance as a relational "
           "achievement, the R1 learning loop, the barrier pathway, and the justice criteria and governance band "
           "that cut across all of it.")

    h2(doc, "5.4 The justice assessment")
    figure(doc, "fig4_justice.png",
           "Figure 4. Justice assessed on four dimensions at once, reported as a profile rather than a single score.",
           "Required by D-01 (evaluation) and RQ5. Defines the four dimensions \u2014 distributive, procedural, "
           "recognition and restorative \u2014 with the indicators used for each, and commits the write-up to "
           "reporting a profile, not a single score.")

    h2(doc, "5.5 The project timeline")
    figure(doc, "fig5_timeline.png",
           "Figure 5. Thirty-six-month schedule with milestones, showing the fieldwork\u2013analysis overlap.",
           "Required by D-01 (plan of work). Six phases from foundation and ethics to finalisation; diamonds mark "
           "milestones; the overlap between fieldwork and analysis is deliberate, because transcription and "
           "coding begin while collection continues.")

    h2(doc, "5.6 Diagram standards")
    bullets(doc, [
        "Every figure carries a sequential number, a descriptive caption and a source note.",
        "Figures are supplied both as placed images (PNG) and as editable sources (SVG with live text), so the "
        "client can rebrand or revise them without redrawing.",
        "If figure content changes, the corresponding notes in the chapter text are updated in the same pass.",
        "Low-resolution raster images embedded in the old concept note are defects; they are being redrawn, not "
        "exported.",
    ])

    # ---- 6. Integrity and compliance ------------------------------------
    h1(doc, "6. Integrity and compliance requirements")
    para(doc, "These rules apply to every deliverable without exception. They are requirements, not aspirations.")
    table(doc, "Table 5: Integrity rules",
          ["#", "Rule", "Check applied"],
          [["I-1", "Originals are never modified. All work products are new files.", "File-hash comparison of copies against originals"],
           ["I-2", "No invented references. Uncertain fields are marked [VERIFY], never guessed.", "Field-by-field check against publisher metadata (Crossref/OpenAlex)"],
           ["I-3", "No invented findings. All methods are written as proposed methods.", "Tense audit against the claim register; no past-tense empirical claims"],
           ["I-4", "Quotations are short, marked and page-referenced; paraphrase by default.", "Quotation register with page numbers"],
           ["I-5", "EndNote temporary citations {Author, Year #ID}; IDs renumbered against the real library.", "RIS import test; ID coverage check"],
           ["I-6", "No promises about Turnitin or AI-detection scores.", "No similarity score appears anywhere in the deliverables"],
           ["I-7", "AI-use position disclosed as required (decision D10).", "Disclosure statement present in every deliverable pack"],
           ["I-8", "Gender language: 'end-users' or 'household members' in the project's own text.", "Terminology change log; exceptions only for cited titles, quotations and sources that genuinely argue gender"],
           ["I-9", "Encoding: UTF-8 throughout; RIS carries a BOM deliberately for EndNote on Windows.", "QC script checks encoding and record counts"]],
          [1.2, 8.6, 6.1], fsize=8.5)

    # ---- 7. Terminology --------------------------------------------------
    h1(doc, "7. Terminology requirements")
    para(doc, "One term, one meaning. The control list below binds the wording of every deliverable; the "
              "terminology change log records each change made against the old concept note.")
    table(doc, "Table 6: Terminology control list",
          ["Term", "Adopted form", "Rule"],
          [["Model name", "SETILL", "One form only, defined at first use; footnote records the superseded expansion"],
           ["End-user", "end-user (hyphenated)", "Preferred throughout; replaces consumer, user, beneficiary, household member where the referent is the same"],
           ["Acceptance", "acceptance", "Defined once; adoption, uptake, take-up and acceptability are related terms, not synonyms"],
           ["Living Lab", "Living Lab", "Capitalised as a proper noun"],
           ["Just transition", "just energy transition", "Used in full at first use in each chapter"],
           ["Informal settlement", "informal settlement", "Consistent with the statistical definition adopted in the study"]],
          [3.4, 4.4, 8.1], fsize=9)

    # ---- 8. Approvals ----------------------------------------------------
    h1(doc, "8. Approvals and open decisions")
    para(doc, "The following decisions block parts of the schedule in Table 4. They are requirements placed on "
              "the client and the faculty, not on the analysis.")
    table(doc, "Table 7: Open decisions",
          ["#", "Decision", "Owner", "Blocks"],
          [["D9", "Referencing style: Harvard or APA 7", "Faculty", "Final formatting of the reference list and EndNote output"],
           ["D10", "Written position on AI assistance", "Supervisor and faculty", "All drafting; the school guideline prohibits AI-generated or AI-rephrased text"],
           ["D11", "Chapter 3 length (draft is ~5,835 words against a 3,500\u20134,500 working budget)", "Supervisor", "Finalisation of Chapter 3 and downstream chapters"],
           ["D12", "Whether a proposal requires a results chapter", "Supervisor", "Existence and content of Chapter 6"],
           ["D13", "Access to the four primary theory texts (Braidotti, Haraway, Barad, Hayles)", "Client and library", "Scholarly standard of Chapter 3; it cannot be written from secondary sources"],
           ["D14", "University template and word limits", "Faculty", "Front matter, heading hierarchy and the word budget"]],
          [1.2, 7.4, 3.0, 4.3], fsize=8.5)

    # ---- 9. Acceptance ---------------------------------------------------
    h1(doc, "9. Acceptance and sign-off")
    para(doc, "A deliverable is accepted when all of the following hold. Rejection is recorded with the failing "
              "criterion, and the corrected deliverable is reissued against the same criterion.")
    bullets(doc, [
        "Every requirement in Sections 3 to 8 that applies to it is satisfied and can be shown to be satisfied.",
        "Its Word master and PDF are generated from the same source and match at delivery.",
        "Every figure it uses meets the diagram standards in Section 5.6 and matches its caption and notes.",
        "The integrity checks in Table 5 pass for the deliverable, including the tense audit and the "
        "no-invented-references check.",
        "Any unresolved item inside the deliverable is flagged rather than hidden, and listed in the QA register.",
    ])
    para(doc, "The client signs off each deliverable by confirming the acceptance criteria against the delivered "
              "files. This specification is itself deliverable D-01 and is subject to the same test.")

    # ---- Appendix --------------------------------------------------------
    h1(doc, "Appendix A. Source documents")
    table(doc, "Table 8: Core source documents",
          ["Document", "Role"],
          [["PhD Detailed Concept Note 2026 (199 pp, 68,011 words)", "Old proposal draft; source of the revision programme"],
           ["Aboalez (2025), Procedia CIRP 135, 8\u201313", "Supervisor's theoretical position: a posthumanist social approach to energy transition"],
           ["Molepo, Aboalez & Mathaba (2025), Energy for Sustainable Development 88, 101777", "Barriers to South Africa's energy transition; 15 expert interviews, AKTESP framework"],
           ["Structure of a Proposal in Engineering Management", "School's proposal-structure note; fixes the abstract at 250\u2013300 words"],
           ["literature_matrix.xlsx (49 sources, 5 sheets)", "Source of truth for the reading list"]],
          [8.4, 7.5], fsize=9)
    para(doc, "The two supervisory articles are held lawfully and are supplied in the figures-of-record folder of "
              "the wider pack. No source has been taken from an unauthorised repository.")

    save(doc, "Client_Requirements_Specification.docx",
         "Client Requirements Specification", "SETILL deliverables, diagrams and acceptance criteria")


def copy_figures():
    os.makedirs(FIGDIR, exist_ok=True)
    for f in ("fig1_design.png", "fig2_energy.png", "fig3_setill.png",
              "fig4_justice.png", "fig5_timeline.png"):
        shutil.copy2(os.path.join(ASSETS, f), os.path.join(FIGDIR, f))
        print("COPIED", f)
    for f in ("fig4_1_setill_framework.svg", "fig4_2_setill_loops.svg"):
        shutil.copy2(os.path.join(SVGDIR, f), os.path.join(FIGDIR, f))
        print("COPIED", f)


if __name__ == "__main__":
    os.makedirs(OUT, exist_ok=True)
    copy_figures()
    build_requirements()
    print("\nDONE")
