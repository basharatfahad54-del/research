# -*- coding: utf-8 -*-
"""Check that every reference (original P684-P756) is still cited in the edited
document's body text or tables, and scan captions/numbering consistency."""
import re, sys
import docx

EDITED = sys.argv[1] if len(sys.argv) > 1 else r"Edited_Proposal_Output\Research_Proposal_Refined.docx"

# --- reference list from the original dump (unchanged in the edited version) ---
refs = []
with open(r"research-main\d13_full_dump.txt", encoding="utf-8") as f:
    lines = f.readlines()
grab = False
for ln in lines:
    m = re.match(r"^\[P(\d+)\|", ln)
    if not m:
        continue
    idx = int(m.group(1))
    if 684 <= idx <= 756:
        grab = True
    elif idx >= 757:
        grab = False
    if grab and idx >= 684:
        text = ln.split("] ", 1)[1].strip() if "] " in ln else ""
        if text:
            surname = text.split(",")[0].strip()
            surname = re.sub(r"^\d+\.\s*", "", surname)
            refs.append((idx, surname, text[:70]))

d = docx.Document(EDITED)
body = "\n".join(p.text for p in d.paragraphs)
tbl_text = "\n".join(c.text for t in d.tables for r in t.rows for c in r.cells)
haystack = body + "\n" + tbl_text

print(f"references parsed: {len(refs)}")
missing = []
for idx, surname, head in refs:
    token = surname.split(" ")[0] if not surname.startswith(("Statistics", "Presidential", "Department", "Intergovernmental", "Republic", "United", "World")) else surname
    if token.lower() not in haystack.lower():
        missing.append((idx, surname, head))
print("references NOT cited anywhere:", len(missing))
for idx, surname, head in missing:
    print(f"  P{idx}  {surname}  | {head}")

# --- caption / numbering consistency ---------------------------------------
def find(pattern):
    return re.findall(pattern, body)

fig_caps = find(r"Figure (\d+\.\d+)[.:]")
tab_caps = find(r"Table (\d+\.\d+)[.:]")
print("figure numbers used:", sorted(set(fig_caps)))
print("table numbers used:", sorted(set(tab_caps)))

# stale references to the old numbering
stale = []
for pat in [r"Figure 1\.1", r"Figure 3\.1", r"Figure 3\.2", r"Figure 5\.1", r"Table 3\.2: What posthumanism", r"Table 3\.1: Comparison"]:
    if re.search(pat, body):
        stale.append(pat)
print("stale old-numbering refs:", stale if stale else "none")

# placeholders
phits = [(i, p.text[:70]) for i, p in enumerate(d.paragraphs)
         if re.search(r"TODO|TBD|XXX|Lorem|\[insert|\bplaceholder", p.text, re.I)]
print("placeholder hits:", phits if phits else "none")
