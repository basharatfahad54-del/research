# -*- coding: utf-8 -*-
"""
Builds the CLIENT_DELIVERABLES pack.
Academic document generator: A4, Times New Roman 12, 1.5 spacing, justified,
professional headings, page fields, automatic TOC, numbered tables.
Internal build script - never shipped to the client.
"""
import os, shutil
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.section import WD_SECTION
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = r"C:\Users\Muhammad Fahad\Desktop\Research paper"
OUT = os.environ.get("DELIV_OUT", os.path.join(ROOT, "CLIENT_DELIVERABLES"))

FOLDER_MAP = {}


def _map(relpath):
    for k, v in FOLDER_MAP.items():
        if relpath.startswith(k):
            return v + relpath[len(k):]
    return relpath

TNR = "Times New Roman"

# --------------------------------------------------------------------------
# low-level helpers
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
    # page setup: A4
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
    pf.space_before = Pt(0)
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
    r.font.name = TNR; r.font.size = Pt(9); r.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
    _field(p, ' PAGE ')
    r2 = p.add_run(" of ")
    r2.font.name = TNR; r2.font.size = Pt(9); r2.font.color.rgb = RGBColor(0x44, 0x44, 0x44)
    _field(p, ' NUMPAGES ')
    for run in p.runs:
        run.font.name = TNR; run.font.size = Pt(9)
        run.font.color.rgb = RGBColor(0x44, 0x44, 0x44)


def add_toc(doc):
    h = doc.add_paragraph()
    h.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = h.add_run("Table of Contents")
    r.font.name = TNR; r.font.size = Pt(14); r.font.bold = True
    h.paragraph_format.space_after = Pt(12)
    p = doc.add_paragraph()
    run = p.add_run()
    f1 = OxmlElement('w:fldChar'); f1.set(qn('w:fldCharType'), 'begin')
    it = OxmlElement('w:instrText'); it.set(qn('xml:space'), 'preserve')
    it.text = 'TOC \\o "1-2" \\h \\z \\u'
    f2 = OxmlElement('w:fldChar'); f2.set(qn('w:fldCharType'), 'separate')
    t = OxmlElement('w:t'); t.text = "[Table of contents is generated automatically when the document is opened. Press Ctrl+A then F9 to refresh.]"
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

    # rule
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
    p = doc.add_paragraph(text, style='Heading 1')
    return p


def h2(doc, text):
    p = doc.add_paragraph(text, style='Heading 2')
    return p


def para(doc, text, italic=False, bold=False, size=12, align=None, space_after=6):
    p = doc.add_paragraph()
    r = p.add_run(text)
    r.font.name = TNR; r.font.size = Pt(size); r.italic = italic; r.bold = bold
    if align is not None:
        p.alignment = align
    p.paragraph_format.space_after = Pt(space_after)
    return p


def bullets(doc, items, numbered=False):
    for it in items:
        p = doc.add_paragraph(style='List Number' if numbered else 'List Bullet')
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
            parts = str(val).split('\n')
            for j, part in enumerate(parts):
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


def save(doc, relpath, title, subject):
    path = os.path.join(OUT, _map(relpath))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    cp = doc.core_properties
    cp.title = title
    cp.subject = subject
    cp.author = "PhD Research Project"
    cp.comments = ""
    doc.save(path)
    print("WROTE", relpath)


DATE = "12 November 2026"
PROJ = "Developing an Innovative Systemic Living Lab for Just Energy Transition in a South African Informal Settlement"

# ==========================================================================
# DOC 1 - REFERENCE VERIFICATION REPORT
# ==========================================================================
def build_reference_verification():
    doc = base_document()
    cover(doc, "Quality Assurance Report", "Reference Verification Report",
          "Verification status of the bibliographic base",
          [("Project", "PhD in Engineering Management"),
           ("Research topic", "Systemic Living Lab for Just Energy Transition"),
           ("Institution", "University of Johannesburg"),
           ("Date", DATE),
           ("Report status", "Final for this phase")])
    add_footer(doc, "Reference Verification Report")
    add_toc(doc)
    doc.add_page_break()

    h1(doc, "1. Purpose and scope")
    para(doc, "This report records the verification status of the bibliographic base underlying the research "
              "document. It answers a single question: which references can be shown to exist, in the form in "
              "which they are cited, against an authoritative metadata source?")
    para(doc, "The scope is the reference base of the current 199-page research document (68,011 words) and the "
              "two supervisory papers supplied with the project. Verification began from zero, because the "
              "research document contains no reference list: the section headed \u201c8. References\u201d is an "
              "empty heading. Approximately 250 to 300 distinct in-text citations therefore had to be treated as "
              "unverified at the outset.")
    note_box(doc, "Integrity statement: no reference in this report has been reconstructed by guesswork. Where a "
                  "source could not be located, it is recorded as unresolved and left unresolved. No DOI, page "
                  "range, volume or year has been supplied from memory or inference.")

    h1(doc, "2. Verification method")
    para(doc, "Every candidate source is subjected to the same eight-step protocol. A record is described as "
              "verified only when it has passed steps 1 to 4 against an authoritative metadata source.")
    table(doc, "Table 1: Verification protocol",
          ["Step", "Action", "Evidence required"],
          [["1", "Extract the in-text citation exactly as it appears", "Raw string and page number"],
           ["2", "Search by DOI, then by bibliographic query", "DOI resolved, or documented as not found"],
           ["3", "Confirm author order, year, title, outlet, volume, issue, pages", "All fields match"],
           ["4", "Record the DOI", "Resolves at doi.org"],
           ["5", "Confirm peer review or accreditation status", "Journal indexing or publisher record"],
           ["6", "Obtain full text lawfully", "File held, or a logged access note"],
           ["7", "Read the relevant section before quoting or paraphrasing closely", "Page numbers recorded"],
           ["8", "Enter into the reference manager and export", "RIS record"]],
          [1.4, 7.2, 7.3])

    h1(doc, "3. Summary of results")
    table(doc, "Table 2: Verification position at the date of this report",
          ["Measure", "Result"],
          [["Distinct in-text citations in the document", "approximately 250\u2013300"],
           ["Records verified against Crossref", "9"],
           ["Records with a confirmed DOI", "9"],
           ["Records whose full text is held", "2"],
           ["Records that could not be located", "2 major sources (37 in-text uses)"],
           ["Reference list present in the document", "None \u2014 heading is empty"],
           ["Proportion of the base verified", "approximately 3\u20134 per cent"]],
          [9.0, 6.9])
    para(doc, "The low proportion is a statement about the state of the document, not about the quality of its "
              "scholarship. The nine verified records form a coherent and identifiable research programme, which "
              "is a sound foundation for rebuilding the reference base.")

    h1(doc, "4. Verified records")
    para(doc, "All nine records below were checked field by field. Author order, year, title, outlet, volume and "
              "DOI were confirmed. Where a field could not be confirmed, this is stated in the final column "
              "rather than filled in.")
    rows = [
        ["V1", "Aboalez (2025)", "Procedia CIRP, 135, 8\u201313", "10.1016/j.procir.2025.01.042", "Complete"],
        ["V2", "Molepo, Aboalez & Mathaba (2025)", "Energy for Sustainable Development, 88, 101777", "10.1016/j.esd.2025.101777", "Complete"],
        ["V3", "Musango et al. (2020)", "Energy Research & Social Science, 70, 101715", "10.1016/j.erss.2020.101715", "Complete"],
        ["V4", "Mukama et al. (2022)", "Technology in Society, 68, 101850", "10.1016/j.techsoc.2021.101850", "Complete"],
        ["V5", "Anditi et al. (2022)", "Energy Research & Social Science, 88, 102476", "10.1016/j.erss.2021.102476", "Complete"],
        ["V6", "Petrulaityte et al. (2022)", "Sustainability, 14, 6291", "10.3390/su14106291", "Complete; open access"],
        ["V7", "Batinge et al. (2022)", "Systems Research and Behavioral Science, 39(2), 287\u2013304", "10.1002/sres.2793", "Volume year 2022; first published online 2021"],
        ["V8", "Oosthuizen, de Kock & Musango (2020)", "South African Journal of Industrial Engineering, 31(3)", "10.7166/31-3-2421", "Page range not returned by the publisher record"],
        ["V9", "Musango (2022)", "Energy Research & Social Science, 88, 102525", "10.1016/j.erss.2022.102525", "Complete"],
    ]
    table(doc, "Table 3: Records verified against Crossref",
          ["ID", "Citation", "Outlet and locator", "DOI", "Notes"],
          rows, [1.2, 4.6, 4.6, 3.4, 2.1], fsize=8.5)

    h2(doc, "4.1 Corrections arising from verification")
    para(doc, "Verification is not a formality. Three corrections follow directly from checking the records "
              "against the publisher metadata:")
    bullets(doc, [
        "Oosthuizen, de Kock and Musango (2020): the first author's initial is L. (Lelani), not R. The citation "
        "must be corrected wherever it appears.",
        "Batinge et al.: the article was published online in 2021 and assigned to volume 39, issue 2 in 2022. "
        "The thesis must cite one year consistently. Volume year 2022 is recommended.",
        "Oosthuizen et al. (2020): the publisher record does not return a page range. The article must not be "
        "cited with invented page numbers; an article number or the DOI alone should be used until the printed "
        "pagination can be confirmed.",
    ])
    h2(doc, "4.2 Significance of the verified cluster")
    para(doc, "Records V3 to V9 are not a random selection. They constitute the published output of a single "
              "research programme concerned with gendered energy innovation in poor urban environments in Africa. "
              "That programme is the immediate intellectual lineage of the present study. Two consequences follow.")
    para(doc, "First, the literature review can be rebuilt around an identifiable body of work rather than "
              "scattered citations. That is a substantial strength of the project, and it is currently invisible "
              "because the document cites these authors without a bibliography.")
    para(doc, "Second, the novelty claim must be sharpened. The programme has already produced a gender analysis "
              "framework, living labs for energy technology innovation, a multi-stakeholder codesign toolkit and "
              "systems modelling of energy services. The present study cannot claim to be first in any of those "
              "areas. Its defensible novelty lies in the combination of a posthumanist theoretical foundation, a "
              "just-transition framing, energy acceptance as an organising construct, and a South African "
              "informal settlement focus.")

    h1(doc, "5. Citations that could not be resolved")
    para(doc, "The following citations appear in the document but could not be matched to a verifiable record. "
              "They are recorded here with the search performed, and they remain unresolved.")
    table(doc, "Table 4: Unresolved and defective citations",
          ["Citation as used", "Uses", "Issue", "Action required"],
          [["Okoye (2020)", "28", "No matching record located. Page references suggest a thesis or monograph, which is poorly indexed.", "Search institutional repositories and ProQuest; identify degree, institution and title; if unobtainable, reduce reliance"],
           ["Kovacic et al. (2019)", "9", "No matching record located.", "Search library databases; likely a chapter or report"],
           ["\u201cSphokazi Phelokazi (2022)\u201d", "9", "Appears to be a given name and surname used as an author name.", "Identify the actual author; likely a citation error"],
           ["\u201cMosango et al. (2020)\u201d", "1", "Misspelling of Musango (V3).", "Correct"],
           ["\u201cMamphweli (2016)\u201d", "6", "Fragment of a three-author citation.", "Restore the full author string"],
           ["\u201cLee and Wareham (2012)\u201d", "9", "Fragment of a three-author citation.", "Restore the full author string"],
           ["\u201cLalk and Thopil (2023)\u201d", "7", "Fragment of a three-author citation.", "Restore the full author string"],
           ["\u201cNations (2014)\u201d", "5", "Fragment of \u201cUnited Nations\u201d.", "Correct"],
           ["\u201cAgenda (2030)\u201d", "3", "Fragment of \u201cUN Agenda 2030\u201d.", "Correct"],
           ["\u201cFilho et al. (2022)\u201d", "5", "Fragment of \u201cLeal Filho et al.\u201d.", "Correct"],
           ["\u201cAgri (2023)\u201d", "4", "A commercial web source used as a figure source.", "Replace with a scholarly source"],
           ["Gates (2019)", "4", "Quoted with a page number that cannot be traced to any located work.", "Trace to the original page or remove the quotation"]],
          [3.9, 1.1, 5.4, 5.5], fsize=8.5)
    para(doc, "The pattern is important. The high proportion of fragments is an artefact of citation practice "
              "rather than evidence of invented sources, but it means the reference list cannot be generated "
              "mechanically from the text. Each citation must be resolved against a real record before it is "
              "retained.")

    h1(doc, "6. Full-text access")
    table(doc, "Table 5: Access status by source group",
          ["Group", "Route", "Status"],
          [["V1, V2 (supervisory papers)", "Supplied with the project; open access", "Held"],
           ["V3\u2013V9 (verified cluster)", "Publisher platforms; V6 open access", "Obtainable through the university library"],
           ["Posthumanist theory monographs", "Library holdings and publisher e-books", "Library access only"],
           ["Energy justice literature", "Publisher platforms; some author manuscripts", "Obtainable"],
           ["South African policy and statistics", "Public departmental and statistical releases", "Obtainable"],
           ["Okoye (2020); Kovacic et al. (2019)", "Unknown until identified", "Uncertain"]],
          [5.2, 7.2, 3.5], fsize=9)
    para(doc, "No paywall circumvention, shadow library or unauthorised repository has been used or will be used. "
              "Where full text cannot be lawfully obtained, only the DOI and the official link are recorded, and "
              "the source is not quoted.")

    h1(doc, "7. Limitations of this report")
    bullets(doc, [
        "Nine records out of an estimated 250 to 300 have been verified. The remaining citations are unverified, "
        "not disproven.",
        "Verification confirms that a record exists and that its metadata match the citation. It does not confirm "
        "that the cited work supports the proposition for which it is cited. That is a separate reading task and "
        "remains outstanding for almost every citation.",
        "No quotation in the document has yet been checked against its original page, with the exception of the "
        "unverifiable Gates quotation identified above.",
        "This report makes no claim about the originality of the document. It is a bibliographic audit only.",
    ])

    h1(doc, "8. Actions required")
    table(doc, "Table 6: Actions arising",
          ["#", "Action", "Owner", "Priority"],
          [["1", "Resolve or remove Okoye (2020) and Kovacic et al. (2019)", "Author / library", "High"],
           ["2", "Correct the four author-name fragments and the Oosthuizen initial", "Author", "High"],
           ["3", "Trace or delete the Gates (2019) quotation", "Author", "High"],
           ["4", "Verify the remaining citations in batches of fifty", "Project", "Medium"],
           ["5", "Confirm the printed pagination for Oosthuizen et al. (2020)", "Library", "Low"],
           ["6", "Replace the commercial figure source with a scholarly equivalent", "Author", "Low"]],
          [1.0, 8.6, 3.2, 2.3], fsize=9.5)
    save(doc, r"08_Quality_Assurance\Reference_Verification_Report.docx",
         "Reference Verification Report", "Bibliographic verification audit")


# ==========================================================================
# DOC 2 - ACADEMIC REVISION REPORT
# ==========================================================================
def build_academic_revision():
    doc = base_document()
    cover(doc, "Quality Assurance Report", "Academic Revision Report",
          "Structural and scholarly revision plan for the research document",
          [("Project", "PhD in Engineering Management"),
           ("Research topic", "Systemic Living Lab for Just Energy Transition"),
           ("Institution", "University of Johannesburg"),
           ("Date", DATE),
           ("Report status", "Final for this phase")])
    add_footer(doc, "Academic Revision Report")
    add_toc(doc)
    doc.add_page_break()

    h1(doc, "1. Purpose")
    para(doc, "This report sets out what must change in the research document, why each change is required, and in "
              "what order the changes should be made. It is an editorial and scholarly plan, not a rewrite. Its "
              "purpose is to give the client and the supervisor a single, checkable record of the revision logic "
              "before any chapter is redrafted.")
    para(doc, "The document under review runs to 199 pages and 68,011 words. It contains a substantial literature "
              "review, a conceptual model and a methodology, but it contains no theoretical framework, no results "
              "chapter, and no reference list.")

    h1(doc, "2. The structural decision")
    para(doc, "Three architectures are possible. The choice determines where posthumanism sits, and therefore "
              "determines whether the mandated theory is visible to an examiner or buried inside a literature "
              "review.")
    table(doc, "Table 1: Architectural options",
          ["Option", "Structure", "Assessment"],
          [["A", "Ch1 Introduction; Ch2 Literature review; Ch3 Theoretical framework; Ch4 Conceptual framework; Ch5 Methodology; Ch6 Results; Ch7 Discussion and conclusion",
            "Recommended. Separates theory from model, which the brief requires; gives posthumanism its own chapter; matches the faculty proposal guideline."],
           ["B", "Keep the conceptual framework as Chapter 3 and add theory as an extended Chapter 2 subsection",
            "Rejected. It subordinates the mandated theory to the literature review and makes the justification for posthumanism impossible to argue at the required depth."],
           ["C", "Merge theory and conceptual framework into a single chapter",
            "Rejected. It collapses the distinction the brief explicitly asks for, and produces a chapter of unmanageable length."]],
          [1.6, 6.4, 7.9], fsize=9)
    para(doc, "Option A is adopted. The existing document already contains a conceptual framework and a "
              "methodology; both are retained but renumbered to Chapters 4 and 5. Chapter 3 is new. This is the "
              "least destructive route to a compliant structure, because it preserves completed work while "
              "inserting the missing chapter.")

    h1(doc, "3. Chapter-by-chapter position")
    table(doc, "Table 2: Current state and required revision",
          ["Chapter", "Current state", "Principal issues", "Required revision", "Priority"],
          [["1. Introduction", "pp. 1\u201320; 7,506 words", "Aims, objectives and research questions are built on a superseded question set; case sites include a second country that is outside scope; the theoretical framing omits posthumanism entirely.",
            "Rebuild the problem statement, aims, objectives and questions; single case site; introduce posthumanism and just transition; state the contribution precisely.", "High"],
           ["2. Literature review", "pp. 21\u2013116; 35,163 words (52% of the document)", "Functions as a catalogue rather than an argument; roughly half of it answers questions that are no longer being asked; no critical synthesis; no gap statement.",
            "Reduce by approximately 45 per cent; reorganise around the five research questions; end with an explicit, defensible gap statement.", "High"],
           ["3. Theoretical framework", "Does not exist", "The mandated theory is absent from the entire document.",
            "New chapter of approximately 28,000 words in thirteen sections.", "Critical"],
           ["4. Conceptual framework", "pp. 117\u2013140; 8,219 words, mislabelled \u201cChapter 1\u201d", "Model is not derived from any theory; components are asserted rather than justified; figures are low-resolution images.",
            "Re-derive every component from Chapter 3; rebuild figures as editable vector graphics; add a component-by-component justification table.", "High"],
           ["5. Methodology", "pp. 141\u2013186; 12,957 words, mislabelled \u201cChapter 2\u201d", "Written against a constructivist paradigm that is incompatible with posthumanism; empirical passages describe fieldwork that has not taken place; design assumes two case sites.",
            "Rewrite rather than edit. Adopt a posthumanist-informed paradigm; convert all empirical claims to intended method; single site; align instruments to the five questions.", "Critical"],
           ["6. Results", "Does not exist", "Not applicable at proposal stage.",
            "Either a proposed analysis strategy, or omitted with the analysis plan absorbed into Chapter 5. Requires supervisor decision.", "Medium"],
           ["7. Discussion and conclusion", "Partly present in the closing block, pp. 187\u2013193", "Framed around a model that will change; no contribution statement tied to the gap.",
            "Rewrite once Chapters 3 to 5 are stable; state theoretical, methodological and practical contributions separately.", "Medium"]],
          [2.6, 3.3, 4.6, 4.4, 1.5], fsize=8)

    h1(doc, "4. Cross-cutting revisions")
    h2(doc, "4.1 Terminology")
    para(doc, "The document uses several terms interchangeably, which undermines the claim to a single coherent "
              "model. Each must be resolved to one form and used consistently thereafter. The most consequential "
              "is the model name, which appears in at least three variants.")
    bullets(doc, [
        "Model name: one form only, defined at first use.",
        "Acceptance, adoption, uptake and acceptability: one construct, defined once, with the others treated as "
        "distinct and related terms rather than synonyms.",
        "End-user: the preferred term throughout, in place of consumer, user, beneficiary and household member "
        "where the referent is the same.",
        "Informal settlement: consistent with the statistical definition adopted in the study, without drifting "
        "between rural and urban framings.",
    ])
    h2(doc, "4.2 Theory")
    para(doc, "Posthumanism must become the organising theory, not an added section. Three specific commitments "
              "follow. First, the ontology of the end-user must be reconstructed relationally, which changes how "
              "participation is described in the methodology. Second, technology acceptance models cannot serve as "
              "an explanatory framework, because their account of the individual rational adopter is incompatible "
              "with a relational ontology; they may be used only as a contrast, and must be represented fairly "
              "when they are. Third, the methodology's paradigm statement must be rewritten to match.")
    h2(doc, "4.3 Empirical integrity")
    para(doc, "The document describes completed fieldwork, including named participant counts, interview "
              "durations, a workshop design and a saturation claim. No data, instrument or ethics clearance "
              "underpins any of it. These passages must be deleted or converted to intended method. The full "
              "passage-level register is held in the Research Consistency Report. This is the single most "
              "important correction in the revision.")
    h2(doc, "4.4 References")
    para(doc, "There is no reference list. Approximately 250 to 300 citations require verification, and nine have "
              "so far been verified. Two heavily used sources cannot be located at all. The reference list cannot "
              "be produced until the citation practice is repaired.")

    h1(doc, "5. Volume estimate")
    table(doc, "Table 3: Indicative word budget after revision",
          ["Chapter", "Current", "Target", "Change"],
          [["1. Introduction", "7,506", "8,000", "+494"],
           ["2. Literature review", "35,163", "19,000", "\u221216,163"],
           ["3. Theoretical framework", "0", "28,400", "+28,400"],
           ["4. Conceptual framework", "8,219", "9,000", "+781"],
           ["5. Methodology", "12,957", "12,000", "\u2212957"],
           ["6. Results or analysis strategy", "0", "6,000", "+6,000"],
           ["7. Discussion and conclusion", "4,000", "7,000", "+3,000"],
           ["Front and end matter", "166", "2,600", "+2,434"],
           ["Total", "68,011", "92,000", "+24,000"]],
          [5.4, 3.2, 3.2, 3.2], fsize=9.5)
    para(doc, "The net growth is dominated by the new theoretical chapter. Most of the increase is offset by "
              "compressing the literature review, which is currently more than half the document and cannot remain "
              "so in a seven-chapter thesis.")

    h1(doc, "6. Sequencing")
    para(doc, "The order of work matters because several chapters depend on decisions made in others. The "
              "following sequence avoids predictable rework.")
    table(doc, "Table 4: Recommended sequence",
          ["Stage", "Work", "Depends on"],
          [["1", "Settle the decisions on institution, document status, question set and terminology", "Client and supervisor"],
           ["2", "Acquire and read the primary theory texts", "Library access"],
           ["3", "Draft Chapter 3 in four reviewed batches", "Stage 2"],
           ["4", "Re-derive Chapter 4 from Chapter 3", "Stage 3"],
           ["5", "Rewrite Chapter 5 against the new paradigm", "Stages 3 and 4"],
           ["6", "Re-scope Chapter 2 to the five questions", "Stage 3"],
           ["7", "Rebuild Chapter 1 last, once the argument is stable", "Stages 3 to 6"],
           ["8", "Write Chapter 7 and the abstract", "Stages 3 to 7"],
           ["9", "Assemble the reference list and verify every entry", "Throughout"]],
          [1.4, 9.2, 5.3], fsize=9.5)
    para(doc, "Chapter 1 is deliberately last. Its function is to state what the thesis argues, and that cannot be "
              "stated until the thesis has been argued.")

    h1(doc, "7. Dependencies and approvals")
    table(doc, "Table 5: Items requiring decision before drafting",
          ["#", "Item", "Blocks"],
          [["1", "Written supervisor position on the use of text-generation tools", "All drafting"],
           ["2", "Confirmation of the referencing style", "Reference list and reference manager configuration"],
           ["3", "University thesis or proposal template", "Front matter and formatting"],
           ["4", "Whether a proposal requires a results chapter", "Existence of Chapter 6"],
           ["5", "Access to the four primary theory texts", "Chapter 3 quality"]],
          [1.0, 9.4, 5.5], fsize=9.5)
    save(doc, r"08_Quality_Assurance\Academic_Revision_Report.docx",
         "Academic Revision Report", "Revision plan and volume estimate")


# ==========================================================================
# DOC 3 - RESEARCH CONSISTENCY REPORT
# ==========================================================================
def build_consistency():
    doc = base_document()
    cover(doc, "Quality Assurance Report", "Research Consistency Report",
          "Internal coherence of theory, terminology, questions and empirical claims",
          [("Project", "PhD in Engineering Management"),
           ("Research topic", "Systemic Living Lab for Just Energy Transition"),
           ("Institution", "University of Johannesburg"),
           ("Date", DATE),
           ("Report status", "Final for this phase")])
    add_footer(doc, "Research Consistency Report")
    add_toc(doc)
    doc.add_page_break()

    h1(doc, "1. Purpose and method")
    para(doc, "This report tests the research document for internal consistency across four dimensions: theory, "
              "terminology, alignment between questions and content, and the integrity of empirical claims. Each "
              "finding is recorded with its location so that it can be checked independently.")
    para(doc, "The method was a full-text extraction and systematic search of the document, followed by a "
              "page-by-page reading of the chapters most likely to carry inconsistency. All 199 pages were "
              "extracted as machine-readable text; no optical character recognition was required, and no page was "
              "unreadable.")

    h1(doc, "2. Theoretical consistency")
    h2(doc, "2.1 The mandated theory is absent")
    para(doc, "The term posthumanism does not appear anywhere in 68,011 words. Neither do the terms just "
              "transition, energy justice, or technology acceptance model. A thesis whose title contains "
              "\u201cjust energy transition\u201d and whose supervisor mandates posthumanism therefore currently "
              "contains neither. This is the most serious scholarly gap in the document and it cannot be repaired "
              "by insertion; it requires a new chapter and consequential revision of the chapters that follow.")
    h2(doc, "2.2 Theory and method are incompatible")
    para(doc, "The methodology adopts a constructivist-interpretivist paradigm, in which knowledge is treated as "
              "constructed by individual meaning-makers. Posthumanism treats agency as distributed across human "
              "and non-human actors. These positions are not merely different in emphasis; they generate "
              "different research designs, different units of analysis and different validity criteria. Unless "
              "the paradigm statement is changed, Chapter 3 and Chapter 5 will contradict each other on the first "
              "reading.")
    h2(doc, "2.3 The acceptance construct is undefined")
    para(doc, "Acceptance, adoption, uptake, take-up and acceptability are used interchangeably. At least five "
              "distinct definitions are implied. A research question about acceptance levels cannot be answered "
              "until one construct is defined and the others are positioned relative to it.")

    h1(doc, "3. Terminology control")
    para(doc, "The following table records each term that drifts, the variants observed, and the form to be "
              "adopted. Once fixed, the replacement can be executed mechanically and checked.")
    table(doc, "Table 1: Terminology control list",
          ["Term", "Variants observed", "Adopted form", "Rule"],
          [["Model name", "SETILL, SETI, GETI, GETILL", "To be fixed by the author", "One form only; defined at first use"],
           ["End-user", "end-user, end user, consumer, user, beneficiary", "end-user", "Preferred throughout; hyphenated"],
           ["Acceptance", "acceptance, adoption, uptake, take-up, acceptability", "acceptance", "Defined once; other terms related to it, not substituted for it"],
           ["Informal settlement", "informal settlement, informal dwelling, slum, shack settlement", "informal settlement", "Consistent with the statistical definition adopted"],
           ["Living Lab", "Living Lab, living lab, Living laboratory", "Living Lab", "Capitalised as a proper noun"],
           ["Just transition", "just transition, just energy transition", "just energy transition", "Used in full at first use in each chapter"]],
          [3.0, 6.2, 3.2, 3.5], fsize=8.5)

    h1(doc, "4. Document architecture")
    para(doc, "The document is a composite of several separately written blocks, and the seams are visible. Two "
              "chapters are mislabelled, the numbering restarts mid-document, and the front matter does not "
              "function.")
    table(doc, "Table 2: Structural defects",
          ["Location", "Defect", "Consequence"],
          [["pp. 1\u20132", "Table of contents is built on broken bookmarks; every entry reads \u201cError! Bookmark not defined\u201d", "Document cannot be navigated"],
           ["p. iii", "List of tables omits most tables in the document", "Front matter is materially incomplete"],
           ["p. iv", "List of figures omits most figures and carries a broken page reference", "Front matter is materially incomplete"],
           ["p. v", "List of acronyms contains errors and omissions", "Definitions conflict with the body text"],
           ["p. vi", "Abstract heading with no text", "Document has no abstract"],
           ["p. 117", "Conceptual framework is numbered \u201cChapter 1\u201d", "Duplicate chapter numbering"],
           ["p. 141", "Methodology is numbered \u201cChapter 2\u201d", "Duplicate chapter numbering"],
           ["p. 193", "\u201c8. References\u201d heading with no entries", "No bibliography exists"],
           ["Throughout", "Figure and table numbering repeats and runs out of sequence", "Cross-references in the text point to the wrong objects"]],
          [3.4, 8.4, 4.1], fsize=8.5)
    para(doc, "Figure numbering is worse than incomplete: at least three different figures are all numbered 2.1, "
              "and the sequence skips and repeats. A list of figures cannot be produced until the figures are "
              "renumbered during revision, which is why that list is not included in this delivery.")

    h1(doc, "5. Empirical claim integrity")
    para(doc, "The document reports fieldwork as completed. It states that ethnographic fieldwork was conducted in "
              "fifty households across two countries, that semi-structured interviews lasted twenty to thirty "
              "minutes each, that a codesign workshop divided participants into six groups of seven, that a biogas "
              "digester was developed and tested, and that data saturation was reached.")
    para(doc, "The client has confirmed that no fieldwork has been conducted and that the document is a proposal. "
              "Every passage in this category must therefore be deleted or converted to intended method. Softening "
              "the tense is not sufficient where specific counts are asserted, because a specific count asserts "
              "that an event occurred.")
    table(doc, "Table 3: Summary of passages requiring correction",
          ["Category", "Description", "Count", "Action"],
          [["A", "Passages asserting completed collection of data", "14", "Delete or rewrite as intended method"],
           ["B", "Passages asserting a second case site outside the approved scope", "3", "Delete"],
           ["C", "Correctly conditional planning language", "5", "Retain"],
           ["\u2014", "Project schedule dated 2027\u20132029", "1", "Retain; consistent with proposal status"]],
          [2.2, 8.0, 1.6, 4.1], fsize=9)
    para(doc, "The passage-level register, with page numbers and exact wording, is held separately as an internal "
              "working document and is available to the supervisor on request. It is not reproduced here because "
              "it consists of text that will be removed.")

    h1(doc, "6. Alignment between questions and content")
    para(doc, "The research brief specifies five questions centred on end-users and the Systems Living Lab. The "
              "document is organised around a different set of four questions centred on gender. The table below "
              "records where each briefed question is currently addressed, if at all.")
    table(doc, "Table 4: Coverage of the briefed research questions",
          ["Question", "Focus", "Current coverage", "Status"],
          [["RQ1", "How Systems Living Lab concepts frame the relationship between inequality, energy security, energy poverty, energy technology innovation and informal settlements",
            "Partly addressed in the Living Lab and systems sections of Chapter 2", "Requires theoretical grounding in Chapter 3"],
           ["RQ2", "Factors determining end-user involvement in decision-making, co-design, co-creation, co-research, prototyping and testing",
            "Addressed descriptively; no theoretical account of participation", "Requires reconstruction in Chapters 3 and 4"],
           ["RQ3", "Cultural elements, norms, practices and technological mechanisms influencing acceptance",
            "Not addressed; the acceptance construct is undefined", "Requires definition and a dedicated section"],
           ["RQ4", "Mechanisms to improve acceptance and end-user contribution to energy security",
            "Not addressed; no mechanism is proposed or evaluated", "Requires a design chapter"],
           ["RQ5", "Whether the proposed Systems Living Lab can engage end-users and stakeholders in real-world settings",
            "Assumed rather than interrogated", "Requires an evaluation framework"]],
          [1.5, 7.0, 4.4, 3.0], fsize=8.5)
    para(doc, "No question is fully answered, and two are not addressed at all. This is expected at proposal stage, "
              "but it confirms that the literature review cannot simply be trimmed: it must be reorganised so that "
              "each question has a corresponding body of evidence.")

    h1(doc, "7. Reference consistency")
    para(doc, "There is no reference list, so no in-text citation can currently be checked against a bibliography. "
              "Approximately 250 to 300 citations are used. Nine have been verified. Two heavily used sources "
              "cannot be located. Several citations are fragments of longer author strings, and at least one "
              "quotation carries a page number that cannot be traced. The full audit is given in the Reference "
              "Verification Report.")

    h1(doc, "8. Conclusions")
    bullets(doc, [
        "The document is structurally coherent as a proposal narrative but not as a thesis: it lacks the theory "
        "chapter, the results chapter and the bibliography that a thesis requires.",
        "The mandated theory is entirely absent, and its introduction will require changes in every subsequent "
        "chapter, not an insertion.",
        "Theory and method currently contradict one another and must be reconciled deliberately.",
        "The empirical claims cannot stand and must be removed; this is a question of research integrity, not of "
        "editing.",
        "The question set in the document is obsolete and the content must be realigned to the five briefed "
        "questions.",
        "The verified reference cluster is a genuine strength and should become the backbone of the rebuilt "
        "literature review.",
    ])
    save(doc, r"08_Quality_Assurance\Research_Consistency_Report.docx",
         "Research Consistency Report", "Internal coherence audit")


# ==========================================================================
# DOC 4 - REFERENCE LIST (VERIFIED SUBSET)
# ==========================================================================
def build_reference_list():
    doc = base_document()
    cover(doc, "Reference Management", "Verified Reference List",
          "Alphabetical list of records confirmed against publisher metadata",
          [("Project", "PhD in Engineering Management"),
           ("Research topic", "Systemic Living Lab for Just Energy Transition"),
           ("Institution", "University of Johannesburg"),
           ("Date", DATE),
           ("Status", "Partial \u2014 verified records only")])
    add_footer(doc, "Verified Reference List")
    doc.add_page_break()

    h1(doc, "Status of this list")
    note_box(doc, "This is not a complete bibliography. It contains the nine records that have been verified "
                  "field by field against publisher metadata. The research document currently contains "
                  "approximately 250 to 300 citations and no reference list of its own. A complete list will be "
                  "issued once the remaining citations have been resolved.")
    para(doc, "The list is presented in author\u2013date form, alphabetically by first author, which is compatible "
              "with the Harvard and APA conventions in general use. The final punctuation and capitalisation will "
              "be adjusted to the university's confirmed style once that style is known.")
    para(doc, "Each entry below can be resolved by its DOI. Where a field could not be confirmed from the "
              "publisher record, the entry does not supply it.")

    h1(doc, "Verified references")
    refs = [
        ("Aboalez, K. (2025) 'A posthumanist social approach to energy transition', ",
         "Procedia CIRP",
         ", 135, pp. 8\u201313. doi: 10.1016/j.procir.2025.01.042."),
        ("Anditi, C., Musango, J.K., Smit, S. and Ceschin, F. (2022) 'Addressing gender dimensions in energy innovations: a gender analysis framework for informal urban settlements in Africa', ",
         "Energy Research & Social Science",
         ", 88, 102476. doi: 10.1016/j.erss.2021.102476."),
        ("Batinge, B., Musango, J.K., Ceschin, F., Ambole, A., Smit, S. and Petrulaityte, A. (2022) 'Modelling gendered innovation for the security of energy services in poor urban environments', ",
         "Systems Research and Behavioral Science",
         ", 39(2), pp. 287\u2013304. doi: 10.1002/sres.2793."),
        ("Molepo, P.M., Aboalez, K. and Mathaba, T.N.D. (2025) 'Analysis of barriers to South Africa's energy transition: perspectives from industry experts', ",
         "Energy for Sustainable Development",
         ", 88, 101777. doi: 10.1016/j.esd.2025.101777."),
        ("Mukama, M., Musango, J.K., Smit, S., Ceschin, F. and Petrulaityte, A. (2022) 'Development of living labs to support gendered energy technology innovation in poor urban environments', ",
         "Technology in Society",
         ", 68, 101850. doi: 10.1016/j.techsoc.2021.101850."),
        ("Musango, J.K. (2022) 'Assessing gender and energy in urban household energy transitions in South Africa: a quantitative storytelling from Groenheuwel informal settlement', ",
         "Energy Research & Social Science",
         ", 88, 102525. doi: 10.1016/j.erss.2022.102525."),
        ("Musango, J.K., Smit, S., Ceschin, F., Ambole, A., Batinge, B., Anditi, C., Petrulaityte, A. and Mukama, M. (2020) 'Mainstreaming gender to achieve security of energy services in poor urban environments', ",
         "Energy Research & Social Science",
         ", 70, 101715. doi: 10.1016/j.erss.2020.101715."),
        ("Oosthuizen, L., de Kock, I. and Musango, J.K. (2020) 'Exploring factors that influence the mainstreaming of gendered energy interventions in poor urban environments: a structured literature review', ",
         "South African Journal of Industrial Engineering",
         ", 31(3). doi: 10.7166/31-3-2421."),
        ("Petrulaityte, A., Ceschin, F., Musango, J.K., Mwiti, E., Anditi, C. and Njoroge, P. (2022) 'Supporting the development of gendered energy innovations for informal urban settlements: GENS codesign toolkit for multistakeholder collaboration', ",
         "Sustainability",
         ", 14, 6291. doi: 10.3390/su14106291."),
    ]
    for pre, journal, post in refs:
        p = doc.add_paragraph()
        p.paragraph_format.left_indent = Cm(0.75)
        p.paragraph_format.first_line_indent = Cm(-0.75)
        p.paragraph_format.space_after = Pt(10)
        p.paragraph_format.line_spacing = 1.5
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        r1 = p.add_run(pre)
        r1.font.name = TNR; r1.font.size = Pt(11.5)
        r2 = p.add_run(journal)
        r2.font.name = TNR; r2.font.size = Pt(11.5); r2.italic = True
        r3 = p.add_run(post)
        r3.font.name = TNR; r3.font.size = Pt(11.5)

    h1(doc, "Verification evidence")
    table(doc, "Table 1: Field-level verification status",
          ["Record", "DOI resolves", "Authors", "Year", "Outlet", "Volume and pages"],
          [["Aboalez (2025)", "Yes", "Confirmed", "Confirmed", "Confirmed", "Confirmed"],
           ["Anditi et al. (2022)", "Yes", "Confirmed", "Confirmed", "Confirmed", "Confirmed"],
           ["Batinge et al. (2022)", "Yes", "Confirmed", "Confirmed", "Confirmed", "Confirmed"],
           ["Molepo et al. (2025)", "Yes", "Confirmed", "Confirmed", "Confirmed", "Confirmed"],
           ["Mukama et al. (2022)", "Yes", "Confirmed", "Confirmed", "Confirmed", "Confirmed"],
           ["Musango (2022)", "Yes", "Confirmed", "Confirmed", "Confirmed", "Confirmed"],
           ["Musango et al. (2020)", "Yes", "Confirmed", "Confirmed", "Confirmed", "Confirmed"],
           ["Oosthuizen et al. (2020)", "Yes", "Confirmed", "Confirmed", "Confirmed", "Pages not returned by publisher"],
           ["Petrulaityte et al. (2022)", "Yes", "Confirmed", "Confirmed", "Confirmed", "Confirmed"]],
          [4.2, 2.3, 2.2, 1.6, 2.4, 3.3], fsize=9)

    h1(doc, "Notes on style and use")
    bullets(doc, [
        "The nine records are supplied in RIS format in the same folder and can be imported directly into "
        "EndNote.",
        "Two heavily cited sources in the research document could not be located and are deliberately absent "
        "from this list: they must not be cited until identified.",
        "This list omits the primary theory texts, because those have not yet been supplied or consulted. They "
        "will be added once acquired.",
        "No entry in this list has been reconstructed from memory. Each was checked against the publisher record.",
    ])
    save(doc, r"06_References_EndNote\Reference_List.docx",
         "Verified Reference List", "Verified bibliographic records")


# ==========================================================================
# DOC 5 - LIST OF ACRONYMS
# ==========================================================================
def build_acronyms():
    doc = base_document()
    cover(doc, "Front Matter", "List of Acronyms",
          "Working list derived from the current document and research brief",
          [("Project", "PhD in Engineering Management"),
           ("Research topic", "Systemic Living Lab for Just Energy Transition"),
           ("Institution", "University of Johannesburg"),
           ("Date", DATE),
           ("Status", "Provisional \u2014 pending terminology lock")])
    add_footer(doc, "List of Acronyms")
    doc.add_page_break()

    h1(doc, "Status of this list")
    note_box(doc, "This list is provisional. It has been rebuilt from the body text of the research document "
                  "rather than copied from its existing acronym page, because that page contains errors and "
                  "conflicts with the body. Entries marked for confirmation must be settled before the list is "
                  "finalised, since several of them are also model-naming decisions.")
    para(doc, "Acronyms are listed alphabetically. The expansion given is the one used in the document, or the "
              "standard expansion where the document does not define the term. Where the document uses two "
              "different expansions for the same acronym, the conflict is recorded rather than silently resolved.")

    h1(doc, "Acronyms")
    rows = [
        ["AGECC", "Advisory Group on Energy and Climate Change", "Defined in document"],
        ["AKTESP", "Africa Knowledge and Technology Exchange Stakeholder Platform", "Confirm full expansion"],
        ["CSIR", "Council for Scientific and Industrial Research", "Defined in document"],
        ["DMRE", "Department of Mineral Resources and Energy", "Standard usage"],
        ["DRM", "Design Research Methodology", "Defined in document"],
        ["EC", "Energy Cultures", "Defined in document"],
        ["EIC", "Energy Innovation Canvas", "Defined in document"],
        ["GENS", "Gendered Energy Innovation for informal urban Settlements", "Research programme; confirm preferred expansion"],
        ["GES", "Gendered Energy Security", "Defined in document"],
        ["GETI", "Gendered Energy Technology Innovation", "Conflicts with model naming; confirm"],
        ["GSTM", "Graduate School of Technology Management", "Defined in document"],
        ["ICT", "Information and Communication Technology", "Standard usage"],
        ["IEA", "International Energy Agency", "Standard usage"],
        ["LPG", "Liquefied Petroleum Gas", "Standard usage"],
        ["MEPI", "Multidimensional Energy Poverty Index", "Defined in document"],
        ["MuSIASEM", "Multi-scale Integrated Analysis of Societal and Ecosystem Metabolism", "Defined in document"],
        ["PED", "Problem Exploration Diagram", "Defined in document"],
        ["PhD", "Doctor of Philosophy", "Standard usage"],
        ["PV", "Photovoltaic", "Standard usage"],
        ["RLabs", "Reconstructed Living Lab", "Defined in document"],
        ["SDG", "Sustainable Development Goal", "Standard usage"],
        ["SEA", "Sustainable Energy Access", "Defined in document"],
        ["SETI", "Systems Energy Technology Innovation", "Conflicts with model naming; confirm"],
        ["SETILL", "Systems Energy Technology Innovation Living Lab", "Principal model name; confirm as the single form"],
        ["SLL", "The document gives two expansions: 'Siyakhula Living Lab' and 'Sustainable Living Lab'", "Conflict \u2014 must be resolved"],
        ["SRM", "Systems Reinforcing Model", "Defined in document"],
        ["Stats SA", "Statistics South Africa", "Standard usage"],
        ["TAM", "Technology Acceptance Model", "Referenced as a contrast only"],
        ["UCD", "User-Centred Design", "Defined in document"],
        ["UJ", "University of Johannesburg", "Institutional"],
        ["UNDP", "United Nations Development Programme", "Standard usage"],
        ["WHO", "World Health Organization", "Standard usage"],
    ]
    table(doc, "Table 1: Acronyms and expansions",
          ["Acronym", "Expansion", "Note"],
          rows, [2.6, 8.6, 4.7], fsize=9)

    h1(doc, "Conflicts requiring resolution")
    para(doc, "Three entries above are not simple definitions but decisions. They are recorded here because an "
              "acronym list cannot be finalised while they remain open.")
    table(doc, "Table 2: Open naming decisions",
          ["Item", "Problem", "Decision required"],
          [["SLL", "Used in the document for both a named Living Lab and a category of Living Lab", "Adopt one meaning; introduce a distinct abbreviation if both are needed"],
           ["SETI / SETILL / GETI", "Three abbreviations circulate for what appears to be one model and one lineage of it", "Fix the model name; define once; use consistently"],
           ["GENS", "Expansion is not given in the document", "Confirm the expansion used by the originating programme"]],
          [3.0, 7.4, 5.5], fsize=9)
    para(doc, "A list of figures and a list of tables are not included in this delivery. The figure and table "
              "numbering in the current document repeats and runs out of sequence, so those lists cannot be "
              "generated until the objects are renumbered during revision.")
    save(doc, r"07_Figures_Tables_Acronyms\List_of_Acronyms.docx",
         "List of Acronyms", "Front matter: acronyms")


# ==========================================================================
# DOC 6 - CLIENT HANDOVER SUMMARY
# ==========================================================================
def build_handover():
    doc = base_document()
    cover(doc, "Client Handover", "Handover Summary",
          "Contents, status and outstanding items in the deliverables pack",
          [("Project", "PhD in Engineering Management"),
           ("Research topic", "Systemic Living Lab for Just Energy Transition"),
           ("Institution", "University of Johannesburg"),
           ("Date", DATE),
           ("Pack status", "Partial delivery \u2014 see Section 4")])
    add_footer(doc, "Handover Summary")
    add_toc(doc)
    doc.add_page_break()

    h1(doc, "1. What this pack contains")
    para(doc, "This pack contains the quality-assurance and reference-management deliverables produced from the "
              "analysis of the research document, together with the front matter that could be completed and "
              "verified. It does not contain the chapter drafts, because those have not yet been written.")
    para(doc, "The distinction matters for how the pack should be read. Everything included here is finished, "
              "checked and safe to circulate. Nothing included here is a draft. Where work is incomplete, the work "
              "is absent rather than present in an unfinished form, and it is listed in Section 4.")

    h1(doc, "2. Included documents")
    table(doc, "Table 1: Delivered files",
          ["Folder", "Document", "Format", "What it is"],
          [["06", "Verified Reference List", "Word and PDF", "Nine records confirmed against publisher metadata"],
           ["06", "EndNote References", "RIS", "The same nine records, importable into EndNote"],
           ["07", "List of Acronyms", "Word and PDF", "Rebuilt working list, with conflicts recorded"],
           ["08", "Reference Verification Report", "Word and PDF", "Bibliographic audit and protocol"],
           ["08", "Academic Revision Report", "Word and PDF", "Structural revision plan and word budget"],
           ["08", "Research Consistency Report", "Word and PDF", "Audit of theory, terminology, questions and empirical claims"],
           ["08", "Literature Matrix", "Excel", "49 sources mapped to themes, questions and evidence strength"],
           ["09", "Handover Summary", "Word and PDF", "This document"],
           ["09", "Delivery README", "Text", "Plain-language guide to the pack"]],
          [1.6, 5.0, 2.6, 6.7], fsize=8.5)

    h1(doc, "3. How to use the pack")
    h2(doc, "3.1 Editable and read-only versions")
    para(doc, "Every substantial document is supplied twice. The Word file is the editable master and is the "
              "version to work from. The PDF is a read-only rendering of the same content, for circulation and "
              "review. The two are generated from a single source, so their content is identical; the PDF will not "
              "reflect edits made to the Word file until it is regenerated.")
    h2(doc, "3.2 Importing the references into EndNote")
    para(doc, "Open EndNote, choose File, then Import, then File. Select the RIS file from the reference folder. "
              "Set the import option to Reference Manager (RIS) and the text translation to Unicode (UTF-8). "
              "Import all records into a new group so that the verified set stays distinguishable from anything "
              "added later.")
    para(doc, "The RIS file contains nine records. It is not a complete library for the thesis, and it should not "
              "be presented as one. A complete library can only be built once the remaining citations have been "
              "resolved.")
    h2(doc, "3.3 Reference PDFs")
    para(doc, "The reference PDF folder is empty. Two of the nine verified sources are already held in the project "
              "workspace and can be copied in on request; the remaining seven require library access. No source "
              "has been obtained from an unauthorised repository, and none will be.")

    h1(doc, "4. Outstanding work")
    para(doc, "The following items are not included in this pack. Each is listed with the reason, so that the "
              "absence is a matter of record rather than an omission.")
    table(doc, "Table 2: Items not delivered",
          ["Item", "Reason not delivered"],
          [["Chapter 3: Theoretical Framework", "Not yet written. Requires the four primary theory texts and a written supervisor position on permissible drafting support. Writing it without the primary texts would require inventing interpretations of named theorists."],
           ["Chapter 4: Conceptual Framework", "Depends on Chapter 3. The existing model is not derived from any theory and cannot be presented as final."],
           ["Chapter 1 and Chapter 2 revisions", "Depend on the question set, terminology lock and Chapter 3."],
           ["Abstract", "The document's abstract page is empty and no abstract can be written before the thesis argument is settled."],
           ["Complete research document", "Does not exist as a coherent whole. The current document is a composite with duplicate chapter numbering and no reference list."],
           ["Reference list (complete)", "Approximately 250 to 300 citations remain unverified; two cannot be located at all."],
           ["List of figures and list of tables", "Figure and table numbering repeats and runs out of sequence; the lists cannot be generated until renumbering is complete."],
           ["Figure files (PNG, SVG)", "The conceptual framework figures are low-resolution raster images embedded in the document, with no editable source. They must be redrawn, not exported."],
           ["EndNote library file", "Not created. A tested library file cannot be produced from nine records; the RIS file is supplied instead."]],
          [5.4, 10.5], fsize=9)

    h1(doc, "5. Approvals required")
    table(doc, "Table 3: Decisions and approvals",
          ["#", "Item", "Who", "Why it blocks delivery"],
          [["1", "Written position on permissible drafting support", "Supervisor", "The faculty guideline prohibits text generation; no chapter can be drafted until this is clarified in writing"],
           ["2", "Confirmation of referencing style", "Faculty", "The reference list cannot be finalised without it"],
           ["3", "University thesis or proposal template", "Faculty", "Front matter and heading hierarchy must follow it"],
           ["4", "Whether a proposal requires a results chapter", "Supervisor", "Determines the final chapter structure"],
           ["5", "Access to the four primary theory texts", "Client and library", "Chapter 3 cannot be written from secondary sources"]],
          [1.0, 6.0, 3.2, 5.7], fsize=9)

    h1(doc, "6. Statement of integrity")
    para(doc, "No research finding, reference, DOI, page number, quotation or institutional requirement in this "
              "pack has been invented. Where a fact could not be established, it is recorded as unresolved. Where "
              "a document does not exist, it is not included. No originality report has been produced, because no "
              "originality check has been run.")
    para(doc, "The pack should be read as an accurate statement of what has been completed, not as a claim about "
              "the whole project. The chapters remain to be written, and they will be written only when the "
              "dependencies in Section 5 are cleared.")
    save(doc, r"09_Final_Handover\Client_Handover_Summary.docx",
         "Client Handover Summary", "Deliverables overview")


# ==========================================================================
# RIS
# ==========================================================================
RIS = """TY  - JOUR
TI  - A posthumanist social approach to energy transition
AU  - Aboalez, K.
JO  - Procedia CIRP
VL  - 135
SP  - 8
EP  - 13
PY  - 2025
DO  - 10.1016/j.procir.2025.01.042
ER  - 

TY  - JOUR
TI  - Addressing gender dimensions in energy innovations: A gender analysis framework for informal urban settlements in Africa
AU  - Anditi, C.
AU  - Musango, J. K.
AU  - Smit, S.
AU  - Ceschin, F.
JO  - Energy Research & Social Science
VL  - 88
SP  - 102476
PY  - 2022
DO  - 10.1016/j.erss.2021.102476
ER  - 

TY  - JOUR
TI  - Modelling gendered innovation for the security of energy services in poor urban environments
AU  - Batinge, B.
AU  - Musango, J. K.
AU  - Ceschin, F.
AU  - Ambole, A.
AU  - Smit, S.
AU  - Petrulaityte, A.
JO  - Systems Research and Behavioral Science
VL  - 39
IS  - 2
SP  - 287
EP  - 304
PY  - 2022
DO  - 10.1002/sres.2793
ER  - 

TY  - JOUR
TI  - Analysis of barriers to South Africa's energy transition: Perspectives from industry experts
AU  - Molepo, P. M.
AU  - Aboalez, K.
AU  - Mathaba, T. N. D.
JO  - Energy for Sustainable Development
VL  - 88
SP  - 101777
PY  - 2025
DO  - 10.1016/j.esd.2025.101777
ER  - 

TY  - JOUR
TI  - Development of living labs to support gendered energy technology innovation in poor urban environments
AU  - Mukama, M.
AU  - Musango, J. K.
AU  - Smit, S.
AU  - Ceschin, F.
AU  - Petrulaityte, A.
JO  - Technology in Society
VL  - 68
SP  - 101850
PY  - 2022
DO  - 10.1016/j.techsoc.2021.101850
ER  - 

TY  - JOUR
TI  - Assessing gender and energy in urban household energy transitions in South Africa: A quantitative storytelling from Groenheuwel informal settlement
AU  - Musango, J. K.
JO  - Energy Research & Social Science
VL  - 88
SP  - 102525
PY  - 2022
DO  - 10.1016/j.erss.2022.102525
ER  - 

TY  - JOUR
TI  - Mainstreaming gender to achieve security of energy services in poor urban environments
AU  - Musango, J. K.
AU  - Smit, S.
AU  - Ceschin, F.
AU  - Ambole, A.
AU  - Batinge, B.
AU  - Anditi, C.
AU  - Petrulaityte, A.
AU  - Mukama, M.
JO  - Energy Research & Social Science
VL  - 70
SP  - 101715
PY  - 2020
DO  - 10.1016/j.erss.2020.101715
ER  - 

TY  - JOUR
TI  - Exploring factors that influence the mainstreaming of gendered energy interventions in poor urban environments: A structured literature review
AU  - Oosthuizen, L.
AU  - de Kock, I.
AU  - Musango, J. K.
JO  - South African Journal of Industrial Engineering
VL  - 31
IS  - 3
PY  - 2020
DO  - 10.7166/31-3-2421
ER  - 

TY  - JOUR
TI  - Supporting the development of gendered energy innovations for informal urban settlements: GENS codesign toolkit for multistakeholder collaboration
AU  - Petrulaityte, A.
AU  - Ceschin, F.
AU  - Musango, J. K.
AU  - Mwiti, E.
AU  - Anditi, C.
AU  - Njoroge, P.
JO  - Sustainability
VL  - 14
SP  - 6291
PY  - 2022
DO  - 10.3390/su14106291
ER  - 
"""


def build_ris():
    path = os.path.join(OUT, _map(r"06_References_EndNote\EndNote_References.ris"))
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(RIS)
    print("WROTE", path)


# ==========================================================================
# README + STATUS FILES
# ==========================================================================
README = """CLIENT DELIVERABLES PACK
===============================================================================
Project : PhD in Engineering Management
Topic   : Developing an Innovative Systemic Living Lab for Just Energy
          Transition in a South African Informal Settlement
Institution : University of Johannesburg
Date    : 12 November 2026
Status  : PARTIAL DELIVERY - see Section 5, Outstanding Work
===============================================================================

Read this file first. It explains exactly what is in the pack, what is not,
and why.


--------------------------------------------------------------------------
1. WHAT IS IN THIS PACK
--------------------------------------------------------------------------

The pack contains three kinds of deliverable:

  A. Quality-assurance reports (folder 08)
     These are finished analytical documents describing the state of the
     research, the revision required, and the verification status of the
     references. They are complete and can be circulated.

  B. Reference management (folder 06)
     A Verified Reference List and an importable RIS file containing the
     nine bibliographic records that have been checked field by field
     against publisher metadata.

  C. Front matter that could be completed (folder 07)
     A rebuilt List of Acronyms.

  D. Handover documentation (folder 09)
     This file and a Handover Summary in Word and PDF.


--------------------------------------------------------------------------
2. EDITABLE WORD FILES vs PDF COPIES
--------------------------------------------------------------------------

Editable masters are the .docx files:

  06_References_EndNote\\Reference_List.docx
  07_Figures_Tables_Acronyms\\List_of_Acronyms.docx
  08_Quality_Assurance\\Reference_Verification_Report.docx
  08_Quality_Assurance\\Academic_Revision_Report.docx
  08_Quality_Assurance\\Research_Consistency_Report.docx
  09_Final_Handover\\Client_Handover_Summary.docx

The .pdf files are read-only renderings of exactly the same content. They are
for circulation and review. They are generated from the Word file, so they are
identical at the time of delivery. If you edit a Word file, the PDF will not
change until it is regenerated.

All Word files use A4 paper, Times New Roman 12 pt, 1.5 line spacing,
justified body text, a consistent heading hierarchy, automatic page numbering
and an automatic table of contents. Press Ctrl+A then F9 in Word to refresh
the table of contents and page numbers after editing.

The .xlsx file (Literature Matrix) is a working analytical instrument and has
no PDF equivalent, which is appropriate for a spreadsheet.


--------------------------------------------------------------------------
3. HOW TO IMPORT THE RIS FILE INTO ENDNOTE
--------------------------------------------------------------------------

  1. Open EndNote.
  2. File > Import > File...
  3. Choose 06_References_EndNote\\EndNote_References.ris
  4. Set "Import Option" to  Reference Manager (RIS)
  5. Set "Text Translation" to  Unicode (UTF-8)
  6. Import into a NEW group, for example "Verified_Set_2026"

The file contains NINE records. It is not a complete library for the thesis.
The research document contains approximately 250-300 citations and no
reference list of its own; only nine have so far been verified. A complete
library cannot be produced until the remaining citations are resolved.

No EndNote .enl library file is included. A library file cannot be created or
tested outside EndNote, and it would be misleading to supply one built from
nine records. The RIS file is the correct interchange format and will import
cleanly.


--------------------------------------------------------------------------
4. REFERENCE PDFs
--------------------------------------------------------------------------

The folder 06_References_EndNote\\Reference_PDFs\\ is currently EMPTY.

Two of the nine verified sources are already held in the project workspace
(the supervisor's 2025 paper and the 2025 energy-transition barriers paper).
The remaining seven require access through the university library.

No source has been obtained from an unauthorised repository, and none will be.
Where full text cannot be lawfully obtained, only the DOI is recorded.


--------------------------------------------------------------------------
5. OUTSTANDING WORK
--------------------------------------------------------------------------

The following are NOT in this pack. They are not missing by accident.

  Chapter 3 (Theoretical Framework)
    Not written. It requires the four primary theory texts and a written
    supervisor position on permissible drafting support. Writing it without
    those texts would mean inventing interpretations of named theorists,
    which will not be done.

  Chapter 4 (Conceptual Framework)
    Depends on Chapter 3. The existing model is not derived from any theory.

  Chapter 1 and Chapter 2 revisions
    Depend on the question set, the terminology lock and Chapter 3.

  Abstract
    The document's abstract page is blank and no abstract can be written
    before the argument is settled.

  Complete Research Document
    Does not exist. The 199-page document is a composite with duplicate
    chapter numbering and no reference list.

  Complete reference list
    Approximately 250-300 citations remain unverified. Two heavily used
    sources (Okoye 2020; Kovacic et al. 2019) cannot be located at all.

  List of Figures and List of Tables
    Figure and table numbering in the source document repeats and runs out of
    sequence. The lists cannot be generated until renumbering is done.

  Figure files (PNG, SVG)
    The conceptual framework figures are low-resolution raster images
    embedded in the PDF, with no editable source. They must be REDRAWN, not
    exported. Exporting them would produce poor-quality files and would
    misrepresent the work as finished.

  EndNote library file (.enl)
    See Section 3.


--------------------------------------------------------------------------
6. DOCUMENTS REQUIRING APPROVAL
--------------------------------------------------------------------------

Before the remaining work can proceed, the following must be settled:

  1. SUPERVISOR - written position on permissible drafting support.
     The faculty guideline states that text generation is prohibited. No
     chapter will be drafted until this is clarified in writing.

  2. FACULTY - confirmation of the referencing style (Harvard or APA 7).
     The reference list cannot be finalised without it.

  3. FACULTY - the university thesis or proposal template.
     Front matter and heading hierarchy must follow it.

  4. SUPERVISOR - whether a proposal requires a results chapter.
     This determines the final chapter structure.

  5. CLIENT and LIBRARY - access to the four primary theory texts
     (Braidotti, Haraway, Barad, Hayles).
     Chapter 3 cannot be written from secondary sources.


--------------------------------------------------------------------------
7. STATEMENT OF INTEGRITY
--------------------------------------------------------------------------

No research finding, reference, DOI, page number, quotation or institutional
requirement in this pack has been invented. Where a fact could not be
established, it is recorded as unresolved. Where a document does not exist,
it is not included.

No originality or similarity report has been produced, because no such check
has been run. No claim is made about similarity scores.

The pack is an accurate statement of what has been completed. It is not a
claim about the whole project.

===============================================================================
End of file.
===============================================================================
"""


def build_readme():
    path = os.path.join(OUT, "09_Final_Handover", "DELIVERY_README.txt")
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(README)
    print("WROTE", path)


STATUS = {
    "01_Theoretical_Framework": (
        "STATUS: OUTSTANDING - Chapter 3 has not been written.\n"
        "=====================================================\n\n"
        "This folder is intentionally empty.\n\n"
        "Chapter 3 (Theoretical Framework) is the largest single piece of new work\n"
        "in the project: thirteen sections, approximately 28,000 words, with\n"
        "posthumanism as the central and mandated theory.\n\n"
        "It cannot be written yet, for two independent reasons:\n\n"
        "  1. The four primary theory texts have not been supplied or consulted\n"
        "     (Braidotti 2013 and 2019; Haraway 1988 and 2016; Barad 2007;\n"
        "     Hayles 1999). A theoretical framework built from secondary\n"
        "     summaries would misrepresent these authors and would not survive\n"
        "     examination.\n\n"
        "  2. The faculty guideline prohibits the use of text-generation tools.\n"
        "     A written supervisor position is required before any drafting.\n\n"
        "A complete thirteen-section structure and argument map has been prepared\n"
        "and is available to the supervisor on request. It is an internal working\n"
        "document and is not part of this pack.\n\n"
        "To release this work: supply the theory texts and the supervisor's written\n"
        "position.\n"
    ),
    "02_Conceptual_Framework": (
        "STATUS: OUTSTANDING - the conceptual framework has not been rebuilt.\n"
        "==================================================================\n\n"
        "This folder is intentionally empty.\n\n"
        "The existing conceptual model (currently mislabelled 'Chapter 1', pages\n"
        "117-140 of the research document) is not derived from any theory. Its\n"
        "components are asserted rather than justified, and its figures are\n"
        "low-resolution raster images with no editable source.\n\n"
        "A conceptual framework must follow from a theoretical framework. Because\n"
        "Chapter 3 does not yet exist, Chapter 4 cannot yet be re-derived.\n\n"
        "No PNG or SVG figure files are supplied. The existing figures cannot be\n"
        "usefully exported: they are embedded bitmaps of poor resolution. They must\n"
        "be redrawn as vector graphics once the model is settled. Exporting them\n"
        "now would produce poor files and would present unfinished work as final.\n"
    ),
    "03_Revised_Research_Chapters": (
        "STATUS: OUTSTANDING - no chapter has been revised.\n"
        "=================================================\n\n"
        "This folder is intentionally empty.\n\n"
        "Chapter 1 and Chapter 2 cannot be revised until three things are settled:\n\n"
        "  - the final research question set (the document is built on a\n"
        "    superseded set of four questions; the brief specifies five);\n"
        "  - the terminology lock (the model name alone appears in at least three\n"
        "    variants);\n"
        "  - Chapter 3, because Chapter 2 must lead into the theory that Chapter 3\n"
        "    establishes.\n\n"
        "The 35,163-word literature review also requires a structural rebuild, not\n"
        "an edit: roughly half of it answers questions that are no longer being\n"
        "asked.\n\n"
        "See 08_Quality_Assurance\\Academic_Revision_Report.docx for the full\n"
        "chapter-by-chapter plan.\n"
    ),
    "04_Abstract": (
        "STATUS: OUTSTANDING - no abstract exists.\n"
        "=========================================\n\n"
        "This folder is intentionally empty.\n\n"
        "The research document contains an abstract heading (page vi) with no text\n"
        "beneath it. There is nothing to copy.\n\n"
        "An abstract states what a study found and argues. This study has not yet\n"
        "been argued, because its theoretical framework does not exist. Writing an\n"
        "abstract now would mean describing work that has not been done.\n\n"
        "The faculty guideline requires 250-300 words.\n"
    ),
    "05_Complete_Research_Document": (
        "STATUS: OUTSTANDING - no complete document exists.\n"
        "=================================================\n\n"
        "This folder is intentionally empty.\n\n"
        "There is no complete research document to deliver. The 199-page source\n"
        "file is a composite of four separately written blocks:\n\n"
        "  pages 1-26    Introduction\n"
        "  pages 27-122  Literature review (35,163 words)\n"
        "  pages 123-146 Conceptual framework, mislabelled 'Chapter 1'\n"
        "  pages 147-186 Methodology, mislabelled 'Chapter 2'\n"
        "  pages 187-199 Proposal tail, ending in an empty References heading\n\n"
        "It has duplicate chapter numbering, no theoretical framework, no results\n"
        "chapter and no reference list. Assembling it into a 'complete document'\n"
        "would produce a misleading artefact.\n\n"
        "The complete document will be assembled once Chapters 1-7 exist.\n"
    ),
}


def build_status():
    for folder, text in STATUS.items():
        path = os.path.join(OUT, _map(folder), "STATUS_OUTSTANDING.txt")
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8", newline="\r\n") as f:
            f.write(text)
        print("WROTE", path)


def build_dirs():
    folders = [
        r"01_Theoretical_Framework",
        r"02_Conceptual_Framework",
        r"03_Revised_Research_Chapters",
        r"04_Abstract",
        r"05_Complete_Research_Document",
        r"06_References_EndNote\Reference_PDFs",
        r"07_Figures_Tables_Acronyms\Figures",
        r"07_Figures_Tables_Acronyms\Tables",
        r"08_Quality_Assurance",
        r"09_Final_Handover",
    ]
    for f in folders:
        os.makedirs(os.path.join(OUT, f), exist_ok=True)
    print("Structure created")


if __name__ == "__main__":
    build_dirs()
    build_reference_verification()
    build_academic_revision()
    build_consistency()
    build_reference_list()
    build_acronyms()
    build_handover()
    build_ris()
    build_readme()
    build_status()
    # copy literature matrix
    src = os.path.join(ROOT, "research_work", "literature_matrix.xlsx")
    dst = os.path.join(OUT, "08_Quality_Assurance", "Literature_Matrix.xlsx")
    if os.path.exists(src):
        shutil.copy2(src, dst)
        print("COPIED Literature_Matrix.xlsx")
    print("\nDONE")
