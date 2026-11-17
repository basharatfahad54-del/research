# -*- coding: utf-8 -*-
"""QA checks for the built thesis dossier:
  1. file parity (DOCX/PDF pairs + reference exports)
  2. heading/table/figure numbering sequences per chapter
  3. leftover proposal-numbering / excluded-institution scans
  4. citation <-> reference-list parity (both directions)
  5. DOCX leftover marker scan ([TABLE]/[REF]/{widths} inside rendered XML)
  6. word counts per chapter (honest D11 reporting)
"""
import io, os, re, zipfile, sys
from collections import Counter

ROOT = r"C:\Users\Muhammad Fahad\Desktop\research-main\research-main"
TB = os.path.join(ROOT, "johny", "thesis_build")
OUT = os.path.join(ROOT, "johny", "Final_Thesis")

CHAPTER_FILES = {1: "merged_1.md", 2: "merged_2.md", 3: "merged_3.md", 4: "merged_4.md",
                 5: "merged_5.md", 6: "merged_6.md", 7: "merged_7.md", 8: "merged_8.md",
                 9: "merged_9.md", 12: "merged_12.md", 13: "merged_13.md"}

def md(n):
    with io.open(os.path.join(TB, "merged_%d.md" % n), "r", encoding="utf-8") as f:
        return f.read()

report = []
def rep(section, lines):
    report.append("## " + section)
    report.extend(lines)
    report.append("")

# --- 1. file parity ---------------------------------------------------------
expected = ["Chapter_1_Introduction", "Chapter_2_Literature_Review", "Chapter_3_Theoretical_Framework",
            "Chapter_4_Conceptual_Framework", "Chapter_5_Research_Design_and_Methodology",
            "Chapter_6_Expected_Contributions", "Chapter_7_Research_Plan_and_Timeline",
            "Chapter_8_Preliminary_Chapter_Outline", "Conclusion", "References",
            "Complete_Thesis", "Table_of_Contents"]
missing = []
for b in expected:
    for ext in (".docx", ".pdf"):
        if not os.path.exists(os.path.join(OUT, b + ext)):
            missing.append(b + ext)
for ext in ("References.ris", "References.bib", "References_EndNote.xml"):
    if not os.path.exists(os.path.join(OUT, ext)):
        missing.append(ext)
rep("1. File parity (DOCX/PDF pairs and reference exports)",
    ["Missing: " + (", ".join(missing) if missing else "none — all present")] +
    ["Files in Final_Thesis: %d" % len(os.listdir(OUT))])

# --- 2. numbering sequences -------------------------------------------------
seq_lines = []
problems = []
for n in sorted(CHAPTER_FILES):
    if n in (12, 13):
        continue
    t = md(n)
    h2 = re.findall(r"^## (\d+\.\d+)\s", t, flags=re.M)
    nums = [tuple(int(x) for x in h.split(".")) for h in h2]
    if nums != sorted(nums):
        problems.append("merged_%d: H2 sequence out of order: %s" % (n, h2))
    expected_first = "%d.1" % n
    if h2 and not h2[0].startswith(expected_first):
        problems.append("merged_%d: first H2 is %s (expected %s)" % (n, h2[0], expected_first))
    tabs = re.findall(r"^\[TABLE\]\s*(Table\s+[A-Za-z0-9.\-]+)", t, flags=re.M)
    figs = re.findall(r"^\[CAPTION\]\s*(Figure\s+[A-Za-z0-9.\-]+)", t, flags=re.M)
    seq_lines.append("merged_%d: %d H2 sections; tables %s; figures %s" %
                     (n, len(h2), ", ".join(x.replace("Table ", "") for x in tabs) or "-", ", ".join(y.replace("Figure ", "") for y in figs) or "-"))
rep("2. Numbering sequences", seq_lines +
    (["PROBLEMS: " + "; ".join(problems)] if problems else ["Sequence problems: none"]))

# --- 3. leftover scans -------------------------------------------------------
leftovers = []
for n in sorted(CHAPTER_FILES):
    t = md(n)
    for pat, label in [
        (r"Section (9|10|11|12)\b", "proposal-section leftover"),
        (r"Chapter (10|11|12)\b", "proposal-chapter leftover"),
        (r"Umuahia|Nigeria|Mater Pluviam|University of Pretoria|GSTM|Sphokazi", "excluded content"),
        (r"gender designs for|women empowerment'", "old gendered RQ phrasing"),
    ]:
        for m in re.finditer(pat, t):
            ctx = t[max(0, m.start() - 60):m.end() + 60].replace("\n", " ")
            leftovers.append("merged_%d [%s]: ...%s..." % (n, label, ctx))
    if n != 12 and n != 13:
        for m in re.finditer(r"Pretoria", t):
            ctx = t[max(0, m.start() - 60):m.end() + 60].replace("\n", " ")
            leftovers.append("merged_%d [Pretoria outside refs]: ...%s..." % (n, ctx))
rep("3. Leftover scans (proposal numbering, excluded institutions, old RQ phrasing)",
    leftovers if leftovers else ["Clean: no leftovers detected"])

# --- 4. citation <-> list parity ---------------------------------------------
ref_text = md(12)
refs = []
for ln in ref_text.splitlines():
    if ln.startswith("[REF] "):
        m = re.match(r"^\[REF\]\s*(.+?)\s*\((\d{4})\)", ln)
        if m:
            surname = m.group(1).split(",")[0].split(" and ")[0].strip()
            refs.append((surname, m.group(2)))
body_all = "\n".join(md(n) for n in sorted(CHAPTER_FILES))
cited = set()
pat_paren = re.compile(r"\(([^()]{2,120}?)\s*((?:19|20)\d{2}[a-z]?)\)")
for m in pat_paren.finditer(body_all):
    inner = m.group(1)
    for chunk in inner.split(";"):
        chunk = chunk.strip()
        if not chunk or re.match(r"^(see|cf\.|e\.g\.|i\.e\.|pp?\.|Section|Chapter|Table|Figure|Appendix|months?|no|p\.|n\.d\.)", chunk):
            continue
        first = re.split(r",|\s+and\s+|\bet al\b", chunk)[0].strip().strip("'\"")
        if first and re.match(r"^[A-Z]", first):
            cited.add((first, m.group(2)))
pat_narr = re.compile(r"([A-Z][A-Za-z'\-\u00C0-\u017F]+(?:'s)?)(?:\s+et\s+al\.?)?\s*\(((?:19|20)\d{2}[a-z]?)\)")
for m in pat_narr.finditer(body_all):
    name = m.group(1)
    if name.rstrip("'s") not in ("The", "In", "This", "Chapter", "Section", "South", "New", "Gen", "Phase", "SDG"):
        cited.add((name.rstrip("'s"), m.group(2)))
ref_set = set(refs)
cited_known = {c for c in cited if c[0].lower()[:5] in {r[0].lower()[:5] for r in refs}}
unmatched = sorted({c for c in cited if c not in ref_set and c[0].lower()[:6] not in {r[0].lower()[:6] for r in refs}})
uncited = sorted(r for r in refs if r not in cited)
rep("4. Citation <-> reference-list parity",
    ["Verified reference list entries: %d" % len(refs),
     "Distinct in-text citations detected: %d" % len(cited),
     "In-text citations with no matching list entry (investigate): %d" % len(unmatched)] +
    ["  MISSING LIST ENTRY: %s %s" % c for c in unmatched[:20]] +
    ["List entries never cited in body (report-only candidates): %d of %d" % (len(uncited), len(refs))] +
    ["  NOT CITED: %s %s" % r for r in uncited[:30]])

# --- 5. DOCX marker leftovers -------------------------------------------------
docx_left = []
for b in expected:
    p = os.path.join(OUT, b + ".docx")
    if not os.path.exists(p):
        continue
    with zipfile.ZipFile(p) as z:
        xml = z.read("word/document.xml").decode("utf-8", "replace")
        for pat in ("[TABLE]", "[REF]", "[CAPTION]", "[FIG]", "[NOTE]", "[PAGEBREAK]"):
            if pat in xml:
                docx_left.append("%s contains %s" % (b, pat))
        if re.search(r"\{[0-9.,\s]+\}", xml):
            docx_left.append("%s contains unreplaced width braces" % b)
rep("5. DOCX marker leftovers", docx_left if docx_left else ["Clean: no markers or width braces left in any DOCX"])

# --- 6. word counts -----------------------------------------------------------
wc = []
for n in sorted(CHAPTER_FILES):
    if n in (12, 13):
        continue
    t = md(n)
    t = re.sub(r"^\[.*$", "", t, flags=re.M)
    t = re.sub(r"^\|.*$", "", t, flags=re.M)
    wc.append("merged_%d: ~%d words" % (n, len(t.split())))
rep("6. Word counts (body text, excluding tables/markers)", wc)

out_path = os.path.join(TB, "qa_results.md")
with io.open(out_path, "w", encoding="utf-8") as f:
    f.write("# QA results (automated)\n\n" + "\n".join(report))
print("WROTE", out_path)
print("\n".join(report[:40]))
