# -*- coding: utf-8 -*-
"""
Builds the SETILL thesis-format proposal dossier (DOCX) from the six part files,
curated supplements and the verified reference list. Internal build script.

Bindings: D1 (UJ), D2 (proposal only; proposed tense), D3 (chapter architecture),
D4 (RQs verbatim), D5 (end-user-centred reframe w/ gender as retained axis),
D6 (Tzaneen only), D7 (biogas), D8 (posthumanist paradigm).

Formatting spec: Times New Roman; body 12pt; H1 16pt bold; H2 14pt bold;
H3 12pt bold; 1.5 spacing; body justified; A4, 2.54cm margins; footer page
numbers; captioned tables/figures; auto-updatable TOC field.
"""
import io, os, re, sys
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

ROOT = r"C:\Users\Muhammad Fahad\Desktop\research-main\research-main"
PRO = os.path.join(ROOT, "proposal")
TB = os.path.join(ROOT, "johny", "thesis_build")
OUT = os.path.join(ROOT, "johny", "Final_Thesis")
ASSETS = os.path.join(PRO, "assets")
os.makedirs(OUT, exist_ok=True)

TNR = "Times New Roman"
TITLE = ("A Systemic Living Lab for a Just Energy Transition: An End-User-Centred "
         "Approach to Energy Technology Innovation in South African Informal Settlements")
SHORT = "SETILL — PhD research proposal presented in thesis format"

CHAPTERS = [
    (1, "Chapter_1_Introduction", "part1_intro.md"),
    (2, "Chapter_2_Literature_Review", "part2_lit.md"),
    (3, "Chapter_3_Theoretical_Framework", "part3_theory.md"),
    (4, "Chapter_4_Conceptual_Framework", "part4_concept.md"),
    (5, "Chapter_5_Research_Design_and_Methodology", "part5_method.md"),
    (6, "Chapter_6_Expected_Contributions", None),
    (7, "Chapter_7_Research_Plan_and_Timeline", None),
    (8, "Chapter_8_Preliminary_Chapter_Outline", None),
]

# ---------------------------------------------------------------------------
# remapping rules
# ---------------------------------------------------------------------------
def hsort(keys):
    def kf(k):
        return tuple(int(x) for x in k.split("."))
    return sorted(keys, key=kf, reverse=True)

def sec_head_rules(heading_map, sub_map=None):
    rules = []
    for src in hsort(heading_map):
        rules.append(("## " + src + " ", "## " + heading_map[src] + " "))
        rules.append(("### " + src + " ", "### " + heading_map[src] + " "))
    if sub_map:
        for src in hsort(sub_map):
            rules.append(("### " + src + " ", "### " + sub_map[src] + " "))
            rules.append(("#### " + src + " ", "#### " + sub_map[src] + " "))
    return rules

def token_rules(tokmap):
    """'Section x[.y]' -> mapped. Longest/most specific first."""
    rules = []
    for src in sorted(tokmap, key=lambda k: (-len(k), tuple(-int(x) for x in k.split(".")))):
        rules.append(("Section " + src, "Section " + tokmap[src]))
    return rules

PART1_HEAD = {"3.1": "1.1", "3.2": "1.3", "3.3": "1.4", "3.4": "1.5",
              "3.5": "1.6", "3.6": "1.7", "3.7": "1.8"}
PART1_SUB = {"3.1.1": "1.1.1", "3.1.2": "1.1.2", "3.1.3": "1.1.3", "3.1.4": "1.1.4",
             "3.5.1": "1.6.1", "3.5.2": "1.6.2", "3.5.3": "1.6.3", "3.5.4": "1.6.4"}
OLD_STRUCT = ("The proposal follows the twelve-section structure required by the Postgraduate "
"School of Engineering Management. Section 4 reviews the literature and identifies the gaps. "
"Section 5 sets out the theoretical framework, beginning with the general function of theory "
"in doctoral research and then developing critical posthumanism as the primary lens. Section 6 "
"presents the conceptual framework, including the ten components of the SETILL model and the "
"relationships between them. Section 7 describes the research design and methodology. Section 8 "
"describes the case site and its relevance to Engineering Management. Section 9 states the "
"expected contributions. Section 10 sets out the research plan and timeline. Section 11 "
"provides the preliminary chapter outline, and Section 12 lists the references. Appendices "
"contain the draft instruments, the consent summary, a glossary and the reference verification record.")
NEW_STRUCT = ("The document follows the structure described in this chapter. Chapter 2 reviews the "
"literature and identifies the gaps. Chapter 3 sets out the theoretical framework, beginning with "
"the general function of theory in doctoral research and then developing critical posthumanism as "
"the primary lens. Chapter 4 presents the conceptual framework, including the ten components of "
"the SETILL model and the relationships between them. Chapter 5 describes the research design and "
"methodology, including the proposed case site and its relevance to Engineering Management. "
"Chapter 6 states the expected contributions. Chapter 7 sets out the research plan and timeline, "
"and Chapter 8 provides the preliminary chapter outline of the eventual thesis. The references and "
"the appendices, which contain the draft instruments, the consent summary, a glossary and the "
"reference verification record, follow the numbered chapters.")

PART3_MAP = {"5.1": "3.1", "5.2": "3.2", "5.3": "3.3", "5.4": "3.4", "5.5": "3.5",
             "5.6": "3.6", "5.7": "3.7", "5.8": "3.9", "5.9": "3.10", "5.10": "3.11",
             "5.11": "3.12", "5.12": "3.13", "5.13": "3.14"}
PART4_MAP = {"6.1": "4.1", "6.2": "4.2", "6.3": "4.3", "6.4": "4.4", "6.5": "4.5",
             "6.6": "4.6", "6.7": "4.7", "6.8": "4.9", "6.9": "4.10"}
PART5_MAP = {"7.1": "5.1", "7.2": "5.2", "7.3": "5.3", "7.4": "5.4", "7.5": "5.5",
             "7.6": "5.6", "7.7": "5.7", "7.8": "5.8", "7.9": "5.9", "7.10": "5.10",
             "7.11": "5.11", "8.1": "5.12.1", "8.2": "5.12.2", "8.3": "5.12.3",
             "8.4": "5.12.4", "8.5": "5.12.5", "8.6": "5.12.6", "8.7": "5.12.7",
             "8.8": "5.12.8"}
PART2_MAP = {"4.1": "2.1", "4.2": "2.2", "4.3": "2.3", "4.4": "2.4", "4.5": "2.6",
             "4.6": "2.7", "4.7": "2.8", "4.8": "2.9"}

NOTE_TABLE11 = ("Chapter references in this and comparable tables follow the preliminary chapter "
                "outline of the eventual thesis, set out in Chapter 8 (Table 8.1).")

R1 = ([("# 3. Introduction", "# 1. Introduction")] +
      sec_head_rules(PART1_HEAD, PART1_SUB) + [
      ("## 3.8 Structure of the proposal", "## 1.9 Structure of the thesis"),
      (OLD_STRUCT, NEW_STRUCT),
      ("The methods described in Section 7", "The methods described in Chapter 5"),
      ("Section 3.7", "Section 1.8"), ("Section 3.6", "Section 1.7"),
      ("Section 3.5", "Section 1.6"), ("Section 3.4", "Section 1.5"),
      ("Section 3.3", "Section 1.4"), ("Section 3.2", "Section 1.3"),
      ("Section 3.1", "Section 1.1"),
      ("Key terms used in this proposal", "Key terms used in this study"),
      ("Several terms are used in this proposal", "Several terms are used in this study"),
])

R2 = ([("# 4. Literature Review", "# 2. Literature Review")] +
      sec_head_rules(PART2_MAP) + [
      ("Table 2.3", "Table 2.4__TMP"), ("Table 2.2", "Table 2.4"),
      ("Table 2.4__TMP", "Table 2.3"),
      ("Section 5.8", "Section 3.9"),
      ("The framework in Section 6", "The framework in Chapter 4"),
      ("The theoretical framework in Section 5", "The theoretical framework in Chapter 3"),
      ("Section 5", "Chapter 3"), ("Section 6", "Chapter 4"),
      ("Section 7", "Chapter 5"), ("Section 8", "Section 5.12"),
])

R3 = ([("# 5. Theoretical Framework", "# 3. Theoretical Framework")] +
      sec_head_rules(PART3_MAP, {"5.3.%d" % i: "3.3.%d" % i for i in range(1, 9)} |
                     {"5.5.1": "3.5.1"} |
                     {"5.7.%d" % i: "3.7.%d" % i for i in range(1, 7)}) +
      token_rules(PART3_MAP) + [
      ("Section 6.6", "Section 4.6"), ("Section 6", "Chapter 4"),
      ("Section 7", "Chapter 5"),
])

R4 = ([("# 6. Conceptual Framework", "# 4. Conceptual Framework")] +
      sec_head_rules(PART4_MAP, {"6.4.%d" % i: "4.4.%d" % i for i in range(1, 11)}) +
      token_rules(PART4_MAP) + [
      ("Figure 3.1", "Figure 4.1"), ("Figure 3.2", "Figure 4.2"),
      ("Section 5.8", "Section 3.9"),
      ("Section 5 ended with six propositions", "Chapter 3 ended with six propositions"),
      ("the distinction between Section 5 and this section", "the distinction between Chapter 3 and this chapter"),
      ("Section 5", "Chapter 3"),
])

R5 = ([("# 7. Research Design and Methodology", "# 5. Research Design and Methodology"),
       ("# 8. Context and Case Study", "## 5.12 Context and case study: Tzaneen")] +
      sec_head_rules(PART5_MAP, {**{"7.2.1": "5.2.1", "7.2.2": "5.2.2"},
                                 **{("7.5.%d" % i): ("5.5.%d" % i) for i in range(1, 8)},
                                 **{("7.6.%d" % i): ("5.6.%d" % i) for i in range(1, 6)}}) +
      [("## 5.12.%d" % i, "### 5.12.%d" % i) for i in range(1, 9)] +
      token_rules({k: v for k, v in PART5_MAP.items() if k.startswith("7.")}) + [
      ("Figure 1.1", "Figure 5.1"), ("Table 6.1", "Table 5.8"),
      ("Section 8", "Section 5.12"), ("Section 7", "Chapter 5"),
      ("Section 9 states", "Chapter 6 states"),
])

R_CH6 = ([("# 9. Expected Contributions", "# 6. Expected Contributions")] +
         sec_head_rules({"9.%d" % i: "6.%d" % i for i in range(1, 8)}) + [
         ("Table 7.1", "Table 6.1"),
         ("change in use | Chapters 5, 7 and 8 |", "change in use | Chapters 6, 8 and 9 |"),
         ("the design of the lab itself | Chapters 5 and 6 |", "the design of the lab itself | Chapters 3, 4 and 9 |"),
         ("in Section 10", "in Chapter 7"),
])

R_CH7 = ([("# 10. Research Plan and Timeline", "# 7. Research Plan and Timeline")] +
         sec_head_rules({"10.%d" % i: "7.%d" % i for i in range(1, 6)},
                        {"10.2.%d" % i: "7.2.%d" % i for i in range(1, 7)}) + [
         ("Table 8.1", "Table 7.1"), ("Figure 5.1", "Figure 7.1"),
])

R_CH8 = ([("# 11. Preliminary Chapter Outline", "# 8. Preliminary Chapter Outline")] +
         sec_head_rules({"11.%d" % i: "8.%d" % i for i in range(1, 4)}) + [
         ("Table 9.1", "Table 8.1"),
])

R_TAIL = [("# 12. References", "# References"),
          ("this proposal", "this thesis document")]

# ---------------------------------------------------------------------------
# assembly
# ---------------------------------------------------------------------------
def load(path):
    with io.open(path, "r", encoding="utf-8") as f:
        return f.read()

def apply_rules(text, rules):
    for old, new in rules:
        text = text.replace(old, new)
    return text

def clean_supp(fname):
    t = load(os.path.join(TB, fname))
    lines = t.splitlines()
    start = 0
    for i, ln in enumerate(lines):
        if re.match(r"^#{1,4} \d", ln):
            start = i
            break
    return "\n".join(lines[start:]).rstrip() + "\n"

def _marker_pos(t6, m):
    if t6.startswith(m + "\n"):
        return 0
    if m not in t6:
        raise ValueError("marker not found: " + m)
    return t6.index("\n" + m + "\n") + 1

def part6_segments():
    t6 = load(os.path.join(PRO, "part6_end.md"))
    marks = ["# 9. Expected Contributions", "# 10. Research Plan and Timeline",
             "# 11. Preliminary Chapter Outline", "# 12. References", "# Appendices"]
    idx = []
    for m in marks:
        idx.append(_marker_pos(t6, m))
    segs = []
    for i in range(len(marks)):
        end = idx[i + 1] if i + 1 < len(marks) else len(t6)
        segs.append(t6[idx[i]:end])
    return segs  # ch6src, ch7src, ch8src, refs, appendices

def insert_before(lines, marker_prefix, suppl_text):
    out, done = [], False
    for ln in lines:
        if not done and ln.startswith(marker_prefix):
            out.extend(suppl_text.splitlines())
            out.append("")
            done = True
        out.append(ln)
    if not done:
        out.extend(suppl_text.splitlines())
    return out

def chapter_text(num):
    """Returns the final (remapped, supplemented) markdown source for one chapter."""
    def rd(n):
        return load(os.path.join(PRO, n))
    if num == 1:
        t = apply_rules(rd("part1_intro.md"), R1)
        lines = t.splitlines()
        lines = insert_before(lines, "## 1.3 Problem statement",
                              clean_supp("supp_ch1.md").rstrip())
        out = []
        for ln in lines:
            if ln.startswith("The relationship between the questions is cumulative"):
                out.append(NOTE_TABLE11)
                out.append("")
            out.append(ln)
        return "\n".join(out)
    if num == 2:
        t = apply_rules(rd("part2_lit.md"), R2)
        lines = t.splitlines()
        lines = insert_before(lines, "## 2.6 End-user participation",
                              clean_supp("supp_ch2_ll.md").rstrip())
        lines = insert_before(lines, "## 2.8 Posthumanism",
                              clean_supp("supp_ch2_biogas.md").rstrip())
        return "\n".join(lines)
    if num == 3:
        t = apply_rules(rd("part3_theory.md"), R3)
        return "\n".join(insert_before(t.splitlines(), "## 3.9 Posthumanism and the Systems Living Lab",
                                       clean_supp("supp_ch3.md").rstrip()))
    if num == 4:
        t = apply_rules(rd("part4_concept.md"), R4)
        return "\n".join(insert_before(t.splitlines(), "## 4.9 Boundaries of the framework",
                                       clean_supp("supp_ch4.md").rstrip()))
    if num == 5:
        t = apply_rules(rd("part5_method.md"), R5)
        return "\n".join(insert_before(t.splitlines(), "## 5.6 Data analysis",
                                       clean_supp("supp_ch5.md").rstrip()))
    segs = part6_segments()
    if num == 6:
        return apply_rules(segs[0], R_CH6)
    if num == 7:
        return apply_rules(segs[1], R_CH7)
    if num == 8:
        return apply_rules(segs[2], R_CH8)
    if num == 12:
        return apply_rules(segs[3], R_TAIL)
    if num == 13:
        return apply_rules(segs[4], R_TAIL)
    if num == 9:
        return clean_supp("supp_conclusion.md")
    raise ValueError(num)

CHAPTER_TITLES = {1: "Introduction", 2: "Literature Review",
                  3: "Theoretical Framework", 4: "Conceptual Framework",
                  5: "Research Design and Methodology", 6: "Expected Contributions",
                  7: "Research Plan and Timeline", 8: "Preliminary Chapter Outline",
                  9: "Conclusion", 12: "References", 13: "Appendices"}

# ---------------------------------------------------------------------------
# DOCX rendering
# ---------------------------------------------------------------------------
def set_style_fonts(doc):
    st = doc.styles["Normal"]
    st.font.name = TNR
    st.font.size = Pt(12)
    st.element.rPr.rFonts.set(qn("w:eastAsia"), TNR)
    pf = st.paragraph_format
    pf.line_spacing = 1.5
    pf.space_after = Pt(6)
    specs = {"Heading 1": (16, True), "Heading 2": (14, True), "Heading 3": (12, True)}
    for name, (size, bold) in specs.items():
        s = doc.styles[name]
        s.font.name = TNR
        s.font.size = Pt(size)
        s.font.bold = bold
        s.font.color.rgb = RGBColor(0, 0, 0)
        s.font.italic = False
        rpr = s.element.get_or_add_rPr()
        rf = rpr.find(qn("w:rFonts"))
        if rf is None:
            rf = OxmlElement("w:rFonts")
            rpr.append(rf)
        for a in ("w:ascii", "w:hAnsi", "w:cs", "w:eastAsia"):
            rf.set(qn(a), TNR)

def setup_section(doc):
    sec = doc.sections[0]
    sec.page_width = Cm(21.0)
    sec.page_height = Cm(29.7)
    for m in ("top_margin", "bottom_margin", "left_margin", "right_margin"):
        setattr(sec, m, Cm(2.54))
    footer = sec.footer
    p = footer.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r = p.add_run()
    fb, it, fe = OxmlElement("w:fldChar"), OxmlElement("w:instrText"), OxmlElement("w:fldChar")
    fb.set(qn("w:fldCharType"), "begin")
    it.set(qn("xml:space"), "preserve")
    it.text = " PAGE "
    fe.set(qn("w:fldCharType"), "end")
    r._r.append(fb); r._r.append(it); r._r.append(fe)
    r.font.name = TNR
    r.font.size = Pt(11)
    header = sec.header
    ph = header.paragraphs[0]
    ph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    rh = ph.add_run(SHORT)
    rh.font.name = TNR
    rh.font.size = Pt(9)
    rh.font.italic = True
    rh.font.color.rgb = RGBColor(0x60, 0x60, 0x60)

def add_field(par, instr, placeholder=""):
    r = par.add_run()
    b = OxmlElement("w:fldChar"); b.set(qn("w:fldCharType"), "begin")
    i = OxmlElement("w:instrText"); i.set(qn("xml:space"), "preserve"); i.text = instr
    s = OxmlElement("w:fldChar"); s.set(qn("w:fldCharType"), "separate")
    t = OxmlElement("w:t"); t.text = placeholder
    e = OxmlElement("w:fldChar"); e.set(qn("w:fldCharType"), "end")
    r._r.append(b); r._r.append(i); r._r.append(s); r._r.append(t); r._r.append(e)
    return r

def add_md_runs(par, text, size=12, base_bold=False):
    """Handle *italic* and **bold** inline markers."""
    tokens = re.split(r"(\*\*[^*]+\*\*|\*[^*]+\*)", text)
    for tk in tokens:
        if not tk:
            continue
        if tk.startswith("**") and tk.endswith("**"):
            r = par.add_run(tk[2:-2]); r.bold = True
        elif tk.startswith("*") and tk.endswith("*"):
            r = par.add_run(tk[1:-1]); r.italic = True
        else:
            r = par.add_run(tk)
        r.font.name = TNR
        r.font.size = Pt(size)
        if base_bold:
            r.bold = True

MDIT = re.compile(r"\*([^*]+)\*")

def add_body(doc, text, size=12, align_justify=True, indent_first=None, hanging=False):
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY if align_justify else WD_ALIGN_PARAGRAPH.LEFT
    add_md_runs(par, text, size=size)
    if hanging:
        par.paragraph_format.left_indent = Cm(1.27)
        par.paragraph_format.first_line_indent = Cm(-1.27)
    return par

def add_heading(doc, text, level, page_break=False):
    style = {1: "Heading 1", 2: "Heading 2", 3: "Heading 3"}[level]
    par = doc.add_paragraph(style=style)
    if page_break:
        par.paragraph_format.page_break_before = True
    add_md_runs(par, text, size={1: 16, 2: 14, 3: 12}[level], base_bold=True)
    # ensure bold/colour despite style
    for r in par.runs:
        r.font.bold = True
        r.font.color.rgb = RGBColor(0, 0, 0)
        r.font.size = Pt({1: 16, 2: 14, 3: 12}[level])
        r.font.name = TNR
    return par

def add_caption(doc, text, before=False):
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.LEFT
    m = re.match(r"^((?:Table|Figure)\s+[A-Za-z0-9.\-]+:)(.*)$", text)
    if m:
        r1 = par.add_run(m.group(1) + " ")
        r1.bold = True
        add_md_runs(par, m.group(2).strip(), size=11)
        for r in par.runs[:1]:
            r.font.size = Pt(11)
    else:
        add_md_runs(par, text, size=11)
    for r in par.runs:
        if r.font.size is None:
            r.font.size = Pt(11)
        r.font.name = TNR
    par.paragraph_format.space_before = Pt(10)
    par.paragraph_format.space_after = Pt(4)
    par.paragraph_format.line_spacing = 1.15
    if before:
        par.paragraph_format.keep_with_next = True
    return par

def add_table(doc, caption, rows, widths):
    if widths:
        total = sum(widths)
        usable = 15.9
        if total > usable:
            widths = [w * usable / total for w in widths]
    add_caption(doc, caption, before=True)
    tab = doc.add_table(rows=len(rows), cols=len(rows[0]))
    tab.style = "Table Grid"
    tab.alignment = WD_TABLE_ALIGNMENT.CENTER
    tab.autofit = False
    tbl_pr = tab._tbl.tblPr
    layout = OxmlElement("w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tbl_pr.append(layout)
    for ri, row in enumerate(rows):
        for ci, cell_text in enumerate(row):
            cell = tab.cell(ri, ci)
            cell.paragraphs[0].text = ""
            par = cell.paragraphs[0]
            add_md_runs(par, cell_text.strip(),
                        size=11, base_bold=(ri == 0))
            for r in par.runs:
                r.font.size = Pt(11)
                r.font.name = TNR
            par.paragraph_format.line_spacing = 1.0
            par.paragraph_format.space_after = Pt(2)
            if ri == 0:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:fill"), "F2F2F2")
                tcPr.append(shd)
            if widths and ci < len(widths):
                cell.width = Cm(widths[ci])
    return tab

def add_note(doc, text):
    par = doc.add_paragraph()
    add_md_runs(par, text, size=11)
    for r in par.runs:
        r.italic = True
        r.font.color.rgb = RGBColor(0x40, 0x40, 0x40)
    par.paragraph_format.left_indent = Cm(0.75)
    return par

def add_figure(doc, fname, width_cm):
    path = os.path.join(ASSETS, fname)
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = par.add_run()
    run.add_picture(path, width=Cm(min(width_cm or 15.6, 15.5)))
    par.paragraph_format.space_before = Pt(10)
    return par

def add_quote(doc, text):
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    add_md_runs(par, text, size=12)
    for r in par.runs:
        r.italic = True
    par.paragraph_format.left_indent = Cm(1.0)
    par.paragraph_format.right_indent = Cm(1.0)
    return par

WIDTH_RE = re.compile(r"\{([0-9.,\s]+)\}\s*$")

def render_lines(doc, lines, first_h1_break=False):
    i, n = 0, len(lines)
    seen_h1 = False
    while i < n:
        raw = lines[i].rstrip()
        i += 1
        if not raw.strip():
            continue
        if raw.startswith("[TABLE]"):
            body = raw[len("[TABLE]"):].strip()
            wm = WIDTH_RE.search(body)
            widths = []
            cap = body
            if wm:
                cap = body[:wm.start()].strip()
                widths = [float(x) for x in wm.group(1).split(",") if x.strip()]
            rows = []
            while i < n and lines[i].lstrip().startswith("|"):
                row_line = lines[i].strip()
                i += 1
                if re.match(r"^\|[\s:\-|]+\|$", row_line):
                    continue
                cells = [c for c in row_line.strip("|").split("|")]
                rows.append(cells)
            add_table(doc, cap, rows, widths)
            continue
        if raw.startswith("[FIG]"):
            body = raw[len("[FIG]"):].strip()
            parts = [p.strip() for p in body.split("|")]
            fname = parts[0]
            w = float(parts[1]) if len(parts) > 1 else None
            add_figure(doc, fname, w)
            continue
        if raw.startswith("[CAPTION]"):
            add_caption(doc, raw[len("[CAPTION]"):].strip(), before=False)
            continue
        if raw.startswith("[NOTE]"):
            add_note(doc, raw[len("[NOTE]"):].strip())
            continue
        if raw.startswith("[REF]"):
            add_body(doc, raw[len("[REF]"):].strip(), hanging=True)
            continue
        if raw.startswith("[QUOTE]"):
            add_quote(doc, raw[len("[QUOTE]"):].strip())
            continue
        if raw.strip() == "[PAGEBREAK]":
            from docx.enum.text import WD_BREAK
            par = doc.add_paragraph()
            par.add_run().add_break(WD_BREAK.PAGE)
            continue
        if raw.startswith("|") or raw.lstrip().startswith("|"):
            # table block without a [TABLE] caption marker (e.g. Appendix D)
            rows = []
            while i < n and lines[i].lstrip().startswith("|"):
                row_line = lines[i].strip()
                i += 1
                if re.match(r"^\|[\s:\-|]+\|$", row_line):
                    continue
                rows.append([c for c in row_line.strip("|").split("|")])
            if rows:
                _bare_table(doc, rows)
            continue
        if raw.startswith("#### "):
            add_heading(doc, raw[5:].strip(), 3)
            continue
        if raw.startswith("### "):
            add_heading(doc, raw[4:].strip(), 3)
            continue
        if raw.startswith("## "):
            add_heading(doc, raw[3:].strip(), 2)
            continue
        if raw.startswith("# "):
            add_heading(doc, raw[2:].strip(), 1, page_break=(seen_h1 or first_h1_break))
            seen_h1 = True
            continue
        if raw.startswith("- ") or raw.startswith("– "):
            par = doc.add_paragraph(style="List Bullet")
            add_md_runs(par, raw[2:].strip(), size=12)
            for r in par.runs:
                r.font.name = TNR
                r.font.size = Pt(12)
            par.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            continue
        add_body(doc, raw)
    return seen_h1

def _bare_table(doc, rows):
    cols = len(rows[0])
    w = 15.9 / cols
    tab = doc.add_table(rows=len(rows), cols=cols)
    tab.style = "Table Grid"
    tab.alignment = WD_TABLE_ALIGNMENT.CENTER
    tab.autofit = False
    for ri, row in enumerate(rows):
        for ci, cell_text in enumerate(row):
            cell = tab.cell(ri, ci)
            par = cell.paragraphs[0]
            add_md_runs(par, cell_text.strip(), size=11, base_bold=(ri == 0))
            for r in par.runs:
                r.font.size = Pt(11)
                r.font.name = TNR
            par.paragraph_format.line_spacing = 1.0
            if ri == 0:
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:fill"), "F2F2F2")
                tcPr.append(shd)
            cell.width = Cm(w)
    return tab

def add_title_page(doc, chapter_label):
    def cline(text, size, bold=False, italic=False, space_before=0, space_after=8):
        par = doc.add_paragraph()
        par.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = par.add_run(text)
        r.font.name = TNR
        r.font.size = Pt(size)
        r.bold = bold
        r.italic = italic
        par.paragraph_format.space_before = Pt(space_before)
        par.paragraph_format.space_after = Pt(space_after)
        return par
    cline("UNIVERSITY OF JOHANNESBURG", 13, bold=True, space_before=90)
    cline("Faculty of Engineering and the Built Environment", 12)
    cline("Postgraduate School of Engineering Management", 12)
    cline(TITLE, 16, bold=True, space_before=48, space_after=16)
    cline("PhD research proposal presented in thesis format", 12, italic=True)
    cline(chapter_label, 14, bold=True, space_before=24)
    cline("November 2026", 12, space_before=48)
    par = doc.add_paragraph()
    par.paragraph_format.page_break_before = True
    par = doc.add_paragraph()
    par.alignment = WD_ALIGN_PARAGRAPH.LEFT
    r = par.add_run("Contents")
    r.font.name = TNR
    r.font.size = Pt(14)
    r.bold = True
    toc_par = doc.add_paragraph()
    add_field(toc_par, ' TOC \\o "1-3" \\h \\z \\u ',
              "Table of contents — right-click and choose \u201cUpdate Field\u201d to build.")

def build_document(chapters, out_name, title_from_chapter=None):
    doc = Document()
    set_style_fonts(doc)
    setup_section(doc)
    if title_from_chapter is not None:
        add_title_page(doc, title_from_chapter)
    for num in chapters:
        txt = chapter_text(num)
        render_lines(doc, txt.splitlines(), first_h1_break=True)
    path = os.path.join(OUT, out_name)
    doc.save(path)
    print("WROTE", path)
    return path

def build_references_doc():
    txt = chapter_text(12)
    doc = Document()
    set_style_fonts(doc)
    setup_section(doc)
    add_title_page(doc, "References")
    render_lines(doc, txt.splitlines(), first_h1_break=True)
    path = os.path.join(OUT, "References.docx")
    doc.save(path)
    print("WROTE", path)
    return path

if __name__ == "__main__":
    which = sys.argv[1] if len(sys.argv) > 1 else "all"
    if which in ("all", "chapters"):
        for num, fname, _src in CHAPTERS:
            build_document([num], fname + ".docx", "Chapter %d: %s" % (num, CHAPTER_TITLES[num]))
        build_document([9], "Conclusion.docx", "Conclusion")
        build_references_doc()
    if which in ("all", "complete"):
        build_document([1, 2, 3, 4, 5, 6, 7, 8, 9, 12, 13],
                       "Complete_Thesis.docx", "Complete thesis document")
    if which in ("all", "dump"):
        for num in [1, 2, 3, 4, 5, 6, 7, 8, 9, 12, 13]:
            with io.open(os.path.join(TB, "merged_%d.md" % num), "w", encoding="utf-8") as f:
                f.write(chapter_text(num))
        print("dumped merged chapters")
