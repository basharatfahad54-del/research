# -*- coding: utf-8 -*-
"""Extract recovery material from the concept note (D-01) for thesis gap-filling.

Writes block-structured excerpts to recovery_raw.md with a citation census per
block, so that only citations already present in the verified 73-entry
proposal reference list can be carried into the thesis.
"""
import io, os, re, sys

BASE = r"C:\Users\Muhammad Fahad\Desktop\research-main\research-main"
BIG = os.path.join(BASE, "00_extracted_text", "PhD DETAILED CONCEPT NOTE 2026_CHECK_261111_113441.txt")
CH4 = os.path.join(BASE, "ch4_dump.txt")
OUT = os.path.join(BASE, "johny", "thesis_build", "recovery_raw.md")

# (label, first_line, last_line) inclusive, 1-based, from the big extract
BLOCKS = [
    ("CH1_12_HISTORICAL_DEVELOPMENT", 261, 455),
    ("CH1_121_DEFINITIONS", 456, 586),
    ("LIT_222_HISTORICAL_ENERGY_ACCESS", 1179, 1219),
    ("LIT_223_ENERGY_TECH_INNOVATION", 1220, 1395),
    ("LIT_227_INDICATORS_ENERGY_POVERTY", 1707, 1865),
    ("LIT_231_LL_ORIGIN_DEFINITIONS", 2566, 2841),
    ("LIT_233_LL_TYPOLOGY", 2842, 3054),
    ("LIT_234_LL_CASE_STUDIES", 3055, 3373),
    ("LIT_235_LL_IN_ENERGY", 3374, 3521),
    ("LIT_236_SYSTEMS_LL_CONCEPTUALISING", 3522, 3673),
    ("LIT_237_REAL_WORLD_TESTING", 3674, 3945),
    ("LIT_241_SYSTEMS_APPROACHES_SD_SE", 3946, 4269),
    ("LIT_242_BIOGAS", 4270, 4526),
    ("LIT_243_POOR_INFORMAL_SETTLEMENTS", 4527, 4683),
    ("LIT_251_INTEGRATION", 4684, 4746),
    ("CF_13_LL_CONCEPTS_SETILL", 4927, 5001),
    ("CF_14_STRUCTURING_ADAPTING_SETILL", 5002, 5162),
    ("CF_15_STAKEHOLDERS_ROLES", 5163, 5452),
    ("CF_16_CONCEPTUALISING_SETILL", 5453, 5532),
    ("METH_INTERNET_RESEARCH", None, None),  # from ch4_dump around line 742
]

CIT = re.compile(r"\(([A-Z][A-Za-z\u00C0-\u017F'\-]+(?:\s+(?:et al\.?|and|&)\s*[A-Z][A-Za-z\u00C0-\u017F'\-]+)*)?,?\s*(?:19|20)\d{2}[a-z]?\)")

def read_lines(path):
    with io.open(path, "r", encoding="utf-8", errors="replace") as f:
        return f.read().splitlines()

big = read_lines(BIG)
ch4 = read_lines(CH4)

# locate internet-research block in ch4_dump
start = None
for i, ln in enumerate(ch4):
    if "Internet research" in ln:
        start = i
        break
internet = None
if start is not None:
    lo = max(0, start - 25)
    hi = min(len(ch4), start + 45)
    internet = (lo + 1, hi, "\n".join(ch4[lo:hi]))

with io.open(OUT, "w", encoding="utf-8") as f:
    f.write("# Recovery raw material (concept note D-01)\n\n")
    for label, a, b in BLOCKS:
        f.write("\n\n===== %s =====\n" % label)
        if label == "METH_INTERNET_RESEARCH":
            if internet:
                f.write("[ch4_dump lines %d-%d]\n" % (internet[0], internet[1]))
                f.write(internet[2] + "\n")
            continue
        chunk = big[a - 1:b]
        text = "\n".join(chunk)
        f.write(text + "\n")
        # citation census
        cits = CIT.findall(text)
        from collections import Counter
        c = Counter(x.strip() for x in cits)
        f.write("\n[CITATION CENSUS] " + ", ".join("%s(%d)" % (k, v) for k, v in sorted(c.items())) + "\n")

print("WROTE", OUT, os.path.getsize(OUT), "bytes")
