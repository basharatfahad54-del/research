# -*- coding: utf-8 -*-
"""Apply the revision layer to D13_Complete_Proposal.docx -> Research_Proposal_Refined.docx
Preserves styles, tables, images and heading structure; verifies each replacement against
the expected original text before applying."""
import sys, os, importlib
import docx
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

BASE = r"C:\Users\Muhammad Fahad\Desktop\research-main\research-main"
SRC  = os.path.join(BASE, "fd john", "deliverables", "D13_Complete_Proposal.docx")
OUT  = os.path.join(BASE, "Edited_Proposal_Output", "Research_Proposal_Refined.docx")
sys.path.insert(0, os.path.join(BASE, "edit_work"))

MODULES = ["rev_a_front_s3", "rev_b_s4", "rev_c_s5a", "rev_d_s5b", "rev_e_s6",
           "rev_f_s7", "rev_g_s8_s9", "rev_h_s10_s11", "rev_i_appendix"]
# Compression layer overrides the revision layer text where both target the same paragraph.
COMP_MODULES = ["comp_s3", "comp_s4", "comp_s5a", "comp_s5b", "comp_s6",
                "comp_s7", "comp_s8_s9", "comp_s10_s11"]

def norm(s):
    return (s.replace("\u2019", "'").replace("\u2018", "'")
             .replace("\u201c", '"').replace("\u201d", '"')
             .replace("\u2013", "-").replace("\u2014", "-").strip())

def set_text(p, text):
    runs = p.runs
    if runs:
        runs[0].text = text
        for r in runs[1:]:
            r._r.getparent().remove(r._r)
    else:
        p.add_run(text)

def make_toc_field(p):
    for r in list(p.runs):
        r._r.getparent().remove(r._r)
    def mk(tag, **attrs):
        el = OxmlElement(tag)
        for k, v in attrs.items():
            el.set(qn(k), v)
        return el
    r1 = mk('w:r');        r1.append(mk('w:fldChar', **{'w:fldCharType': 'begin'}))
    r2 = mk('w:r')
    instr = mk('w:instrText', **{'xml:space': 'preserve'})
    instr.text = ' TOC \\o "1-3" \\h \\z \\u '
    r2.append(instr)
    r3 = mk('w:r');        r3.append(mk('w:fldChar', **{'w:fldCharType': 'separate'}))
    r4 = mk('w:r')
    t = OxmlElement('w:t')
    t.text = ('Table of contents: after opening the document, select all (Ctrl+A) and press F9, '
              'or right-click here and choose "Update Field", to generate the contents list.')
    r4.append(t)
    r5 = mk('w:r');        r5.append(mk('w:fldChar', **{'w:fldCharType': 'end'}))
    for el in (r1, r2, r3, r4, r5):
        p._p.append(el)

def build_grid_table(doc, rows_data, ncols):
    tbl = doc.add_table(rows=len(rows_data), cols=ncols)
    try:
        tbl.style = 'Table Grid'
    except Exception:
        pass
    for i, row in enumerate(rows_data):
        for j, val in enumerate(row):
            cell = tbl.cell(i, j)
            cell.text = val
            if i == 0:
                for run in cell.paragraphs[0].runs:
                    run.bold = True
    return tbl

def main():
    doc = docx.Document(SRC)
    paras = list(doc.paragraphs)          # enumeration identical to the dump
    def P(idx):
        return paras[int(idx[1:])]

    # --- collect revisions -------------------------------------------------
    REVS, DELS = {}, set()
    for m in MODULES:
        mod = importlib.import_module(m)
        for k, pair in mod.REVS.items():
            if k in REVS:
                raise SystemExit(f"duplicate revision key {k}")
            REVS[k] = pair
        DELS.update(getattr(mod, "DELS", []))

    # --- collect compression layer (overrides rev text where both apply) ----
    comp = {}
    for m in COMP_MODULES:
        mod = importlib.import_module(m)
        for k, pair in mod.REVS.items():
            if k in comp:
                raise SystemExit(f"duplicate compression key {k}")
            comp[k] = pair
        DELS.update(getattr(mod, "DELS", []))

    # --- verify expected text (against the text each layer will see) --------
    mismatches = []
    for k, (expect, _text) in REVS.items():
        if k in comp:      # comp verified against the pass-1 (rev) text
            continue
        cur = norm(P(k).text)
        if not cur.startswith(norm(expect)):
            mismatches.append((k, norm(expect)[:40], cur[:60]))
    for k, (expect, _text) in comp.items():
        base = REVS[k][1] if k in REVS else P(k).text
        base = "" if base is None else base
        if not norm(base).startswith(norm(expect)):
            mismatches.append((k, norm(expect)[:40], norm(base)[:60]))
    for k in DELS:
        if not P(k).text.strip() and k not in REVS and k not in comp:
            mismatches.append((k, "<non-empty expected>", "<empty>"))
    if mismatches:
        for m in mismatches:
            print("MISMATCH:", m)
        raise SystemExit("Aborting: paragraph indices do not match the dump.")

    # --- apply text replacements (comp overrides rev) ------------------------
    final = {}
    for k, (_e, text) in REVS.items():
        final[k] = text
    for k, (_e, text) in comp.items():
        final[k] = text
    for k, text in sorted(final.items(), key=lambda x: int(x[0][1:])):
        if text is not None:
            set_text(P(k), text)

    # --- delete paragraphs ---------------------------------------------------
    for k in sorted(DELS, key=lambda x: int(x[1:]), reverse=True):
        el = P(k)._p
        el.getparent().remove(el)

    # --- delete the status-note box table ------------------------------------
    removed_tables = 0
    for tbl in list(doc.tables):
        try:
            first = tbl.rows[0].cells[0].text
        except Exception:
            continue
        if first.strip().startswith("AI-assisted working draft"):
            tbl._tbl.getparent().remove(tbl._tbl)
            removed_tables += 1

    # --- real TOC field ------------------------------------------------------
    make_toc_field(P("P23"))

    # --- Appendix D: markdown diary text -> real table ------------------------
    diary_rows = [
        ["Time", "What was cooked or used", "Fuel or device", "Approximate cost", "Who did the work", "Anything that went wrong"],
        ["Morning", "", "", "", "", ""],
        ["Midday", "", "", "", "", ""],
        ["Evening", "", "", "", "", ""],
        ["Night", "", "", "", "", ""],
    ]
    t = build_grid_table(doc, diary_rows, 6)
    P("P812")._p.addnext(t._tbl)
    el = P("P813")._p
    el.getparent().remove(el)

    # --- Appendix F: markdown glossary text -> real table ---------------------
    glossary_rows = [
        ["Term", "Plain-language definition"],
        ["Agential cut", "The boundary drawn within a phenomenon that separates what is observed from what observes; it is a decision with consequences."],
        ["Agential realism", "Barad's ontology in which entities emerge from relations rather than preceding them."],
        ["Assemblage", "A configuration of human and non-human elements whose relations produce capacities that none has alone."],
        ["Diffraction", "Reading one account through another to see what each makes visible, rather than judging one against the other."],
        ["Distributed agency", "The claim that the capacity to act is spread across people, artefacts, infrastructures and institutions."],
        ["Intra-action", "The relation through which entities acquire their boundaries and properties; distinct from interaction between pre-existing entities."],
        ["Potestas", "Braidotti's term for constraining, negating power."],
        ["Potentia", "Braidotti's term for enabling, capacity-building power."],
        ["Response-ability", "Haraway's term for the practised capacity to respond to others accountably."],
        ["Situated knowledge", "Knowledge produced from a particular position, which can therefore be located and checked."],
        ["Sympoiesis", "Making-with; collaboration as a condition of life rather than a technique."],
        ["Apparatus", "The arrangement of practices, instruments and concepts through which a phenomenon is produced and observed."],
    ]
    t = build_grid_table(doc, glossary_rows, 2)
    P("P824")._p.addnext(t._tbl)
    el = P("P825")._p
    el.getparent().remove(el)

    os.makedirs(os.path.dirname(OUT), exist_ok=True)
    doc.save(OUT)
    print(f"Saved: {OUT}")
    print(f"Replacements applied: {sum(1 for _, t in REVS.values() if t is not None)}")
    print(f"Paragraphs deleted: {len(DELS)}")
    print(f"Tables removed: {removed_tables} (status-note box)")

if __name__ == "__main__":
    main()
