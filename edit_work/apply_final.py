# -*- coding: utf-8 -*-
"""Second pass: Research_Proposal_Refined.docx -> Research_Proposal_Final.docx.

Keyed to refined_dump.txt paragraph indices. Four kinds of edit:
  1. REVS      - phrase-level polish on six paragraphs
  2. TABLE FIX - Table 3.6 cell 'Chapter 5 and Chapter 8' -> 'Chapters 4 and 8'
                 (original cell contradicts the original's own thesis outline,
                  where Chapter 4 is the methodology chapter)
  3. GLOSSARY  - Appendix F rows sorted alphabetically
  4. TYPOGRAPHY- apostrophes (\w'\w -> curly) and paired straight double quotes
                 -> curly, in body paragraphs and table cells, EXCEPT the
                 reference list (Harvard single quotes must stay) and the
                 reference-note box (tables[27]) and search-string syntax cells.
"""
import re
import docx
from docx.oxml.ns import qn

BASE = r"C:\Users\Muhammad Fahad\Desktop\research-main\research-main"
SRC = BASE + r"\Edited_Proposal_Output\Research_Proposal_Refined.docx"
OUT = BASE + r"\Edited_Proposal_Output\Research_Proposal_Final.docx"

REF_RANGE = set(range(533, 606))   # refined P533-P605 = reference entries

d = docx.Document(SRC)
paras = list(d.paragraphs)


def norm(s):
    for a, b in (("\u2019", "'"), ("\u2018", "'"), ("\u201c", '"'), ("\u201d", '"'),
                 ("\u2013", "-"), ("\u2014", "-")):
        s = s.replace(a, b)
    return s


def set_text(p, text):
    # keep first run's formatting, drop the rest
    runs = p.runs
    if not runs:
        p.add_run(text)
        return
    runs[0].text = text
    for r in runs[1:]:
        r.text = ""


# ------------------------------------------------------------------ 1. REVS
REVS = {
    "P22": ("Table of contents: after opening the document,",
            "The table of contents is a Microsoft Word field: select the whole document (Ctrl+A) and press F9, or right-click on the field and choose \u201cUpdate Field\u201d, to generate or refresh the list of contents."),
    "P92": ("The proposal follows the required twelve-section structure.",
            "The proposal follows the required twelve-section structure. Section 4 reviews the literature and identifies the gaps; Section 5 develops the theoretical framework; Section 6 presents the conceptual framework and the ten components of the SETILL model; Section 7 describes the design and methodology; Section 8 describes the case site; Section 9 states the expected contributions; Section 10 sets out the plan and timeline; Section 11 gives the preliminary chapter outline; and Section 12 lists the references. Appendices contain the draft instruments and workshop protocol, the consent summary, the diary template, a glossary and the reference verification record."),
    "P210": ("A TAM study would report that residents find the idea of biogas useful.",
             "A TAM study would report that residents find the idea of biogas useful. A posthumanist study would report that the system will fail where water is carried, succeed where a market provides a steady waste stream, and impose an additional burden on the women expected to feed it. The second account is less tidy, but more useful for design."),
    "P263": ("The theory is load-bearing for every question.",
             "The theory is load-bearing for every question. Table 3.5 sets out the relationship, and the paragraphs that follow show what it means for each question."),
    "P444": ("The study is a single embedded case study conducted through four DRM stages",
             "The study is a single embedded case study conducted through four DRM stages in an informal settlement in Greater Tzaneen. Data are collected through interviews, focus groups, transect walks, diaries, photovoice, co-design workshops and documents, and analysed through reflexive thematic analysis, situational mapping, causal loop diagramming and diffractive reading. Quality is addressed through the eight big-tent criteria and a documented audit trail. Ethics is addressed both through procedural clearance and through a relational practice that includes an exit plan and shared findings. The next section describes the case site."),
    "P489": ("A doctoral study is also a process of formation,",
             "A doctoral study is also a process of formation, and it is reasonable to state what the researcher expects to learn: to conduct qualitative research in a context of poverty and inequality; to apply posthumanist theory to engineering problems; and to manage a multi-stakeholder design process. The fieldwork requires working in a language and setting that are not the researcher\u2019s own; the theoretical work requires fluency in concepts not native to engineering; and the design work requires the management skills the study is, in part, about."),
}

mismatches = []
for k, (expect, _text) in REVS.items():
    cur = norm(paras[int(k[1:])].text)
    if not cur.startswith(norm(expect)):
        mismatches.append((k, norm(expect)[:40], cur[:60]))
if mismatches:
    for m in mismatches:
        print("MISMATCH:", m)
    raise SystemExit("Aborting: refined paragraph indices do not match.")

for k, (_expect, text) in sorted(REVS.items(), key=lambda x: int(x[0][1:])):
    set_text(paras[int(k[1:])], text)
print("paragraph edits:", len(REVS))

# ------------------------------------------------------------ 2. TABLE 3.6
cell = d.tables[11].rows[4].cells[2]
got = cell.text.strip()
if got != "RQ5; Chapter 5 and Chapter 8":
    raise SystemExit(f"Aborting: Table 3.6 cell unexpected: {got!r}")
set_text(cell.paragraphs[0], "RQ5; Chapters 4 and 8")
print("Table 3.6 cross-reference fixed: RQ5 row 'Chapter 5 and Chapter 8' -> 'Chapters 4 and 8'")

# ------------------------------------------------------------- 3. GLOSSARY
tbl = d.tables[30]
if tbl.rows[0].cells[0].text.strip() != "Term" or tbl.rows[1].cells[0].text.strip() != "Agential cut":
    raise SystemExit("Aborting: Appendix F glossary table not where expected.")
body_trs = [r._tr for r in list(tbl.rows)[1:]]


def key(tr):
    tcs = tr.findall(qn("w:tc"))
    if not tcs:
        return ""
    return "".join(n.text or "" for n in tcs[0].findall(".//" + qn("w:t")))


ordered = sorted(body_trs, key=key)
hdr = tbl.rows[0]._tr
for tr in body_trs:
    tr.getparent().remove(tr)
prev = hdr
for tr in ordered:
    prev.addnext(tr)
    prev = tr
print("Appendix F glossary rows sorted alphabetically:",
      [t.rows[i].cells[0].text.strip() for i, t in [(0, tbl)] for i in range(1, 4)])

# ----------------------------------------------------------- 4. TYPOGRAPHY
APO = re.compile(r"(?<=\w)'(?=\w)")


def normalize_element(el, skip_if=None, stats=None):
    """Normalise apostrophes and paired double quotes across every w:t node
    under el, preserving run boundaries via a paragraph-level toggle."""
    for p_el in el.iter(qn("w:p")):
        nodes = list(p_el.iter(qn("w:t")))
        full = "".join(n.text or "" for n in nodes)
        if skip_if and skip_if(full):
            continue
        if '"' in full and full.count('"') % 2 == 0:
            open_q = True
            for n in nodes:
                if not n.text:
                    continue
                out = []
                for ch in n.text:
                    if ch == '"':
                        out.append("\u201c" if open_q else "\u201d")
                        open_q = not open_q
                    else:
                        out.append(ch)
                n.text = "".join(out)
        changed = False
        for n in nodes:
            if n.text and "'" in n.text:
                new = APO.sub("\u2019", n.text)
                if new != n.text:
                    n.text = new
                    changed = True
        if stats is not None:
            stats[0] += 1


body_el = d.element.body
stats = [0]

# body paragraphs: skip the reference-entry range
for idx, p in enumerate(paras):
    if idx in REF_RANGE:
        continue
    normalize_element(p._p, stats=stats)

# table cells: skip the reference-note box and search-syntax cells
note_tbl_el = d.tables[27]._tbl


def skip_cell(txt):
    return " OR " in txt


for t in d.tables:
    if t._tbl is note_tbl_el:
        continue
    for row in t.rows:
        for c in row.cells:
            normalize_element(c._tc, skip_if=skip_cell, stats=stats)

d.save(OUT)
print("typography normalised across", stats[0], "paragraph/cell blocks")
print("Saved:", OUT)
