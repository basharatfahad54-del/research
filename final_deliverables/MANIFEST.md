# MANIFEST — Final Deliverables Package

**Project:** SETILL — Systemic Living Lab for Just Energy Transition
**Root:** `C:\Users\fahad\OneDrive\Desktop\research_work`
**Package:** `final_deliverables/`
**Date:** 14 November 2026
**AI status:** all prose is AI-assisted working draft; decision D10 outstanding.

---

## 1. File index

| # | File | Type | Words / records | Purpose |
|---|---|---|---|---|
| 00 | `00_PROJECT_STATUS_REPORT.md` | Report | — | Overall status, workstreams A–L, decisions, next steps |
| 01 | `01_CHAPTER3_COMPLIANCE_REVIEW.md` | Report | — | Chapter 3 checked against the school's four requirements |
| 02 | `02_SOURCE_VERIFICATION_REPORT.md` | Report | — | Verification state of all 49 matrix sources |
| 03 | `03_ABSTRACT.txt` | Draft | 267 words | Abstract to the school's 250–300-word, four-element requirement |
| 04 | `04_FRONT_MATTER.txt` | Draft | — | List of figures, list of tables, acronyms, numbering note |
| 05 | `05_TERMINOLOGY_CHANGE_LOG.md` | Report | — | Old concept note → proposal terminology |
| 06 | `06_SCHOOL_COMPLIANCE_MATRIX.md` | Report | — | The school's 12 sections mapped to the package |
| 07 | `07_QA_AND_OUTSTANDING_ISSUES.md` | Report | — | QA results, defects fixed, open issues |
| 08 | `08_CHAPTER4_CONCEPTUAL_FRAMEWORK.txt` | Draft | 2,132 words | Chapter 4: constructs, relationships, components, figures |
| 09 | `09_references.ris` | Data | 34 records | RIS import for EndNote/Zotero/Mendeley |
| 10 | `10_references.bib` | Data | 34 records | BibTeX |
| 11 | `11_references_endnote.xml` | Data | 34 records | EndNote XML |
| 12 | `12_reference_coverage_report.txt` | Report | — | Which matrix rows are and are not covered |
| — | `assets/fig4_1_setill_framework.svg` | Figure | — | Editable conceptual framework diagram |
| — | `assets/fig4_2_setill_loops.svg` | Figure | — | Editable feedback-loops diagram |
| — | `MANIFEST.md` | This file | — | Package index and instructions |

### Related files outside this package

| Path | Why it matters |
|---|---|
| `proposal_v3/03_chapter3_theoretical_framework.txt` | Current Chapter 3 (5,835 body words) |
| `proposal_v2/03_chapter3_theoretical_framework.txt` | Earlier Chapter 3; superseded |
| `proposal_v2/00_abstract_plan.txt` | Rationale for the abstract |
| `proposal/part1_intro.md` … `part6_end.md` | Long-form source for Chapters 1, 2 and 5 |
| `06_chapter_revisions/decision_record_and_claim_register.md` | Decisions D1–D11 |
| `literature_matrix.xlsx` | 49-source matrix, gap map, RQ-theory map |
| `doi_check.json` | Verified DOI metadata used to build the reference files |
| `README.md` | Workspace guide |

---

## 2. How to use the reference files

**EndNote (recommended by the school):** File → Import → File, choose
`11_references_endnote.xml`, import option "EndNote generated XML".
Alternatively import `09_references.ris` and check that volume, issue and pages landed
in the right fields.

**BibTeX/LaTeX:** add `10_references.bib` to the bibliography and use
`\bibliography{10_references}`.

**Important:** these files contain **only verified records** (34). They do **not** cover
the 40 matrix rows still marked "TBC". See `12_reference_coverage_report.txt` before
assuming the bibliography is complete.

---

## 3. How to regenerate the reference files

1. Add or correct DOIs in `literature_matrix.xlsx` (Sources sheet, column "DOI / identifier").
2. Re-run the DOI check (`check_dois.py`, or the existing Crossref lookup in the workspace).
3. Run the generator used for this package (see `12_reference_coverage_report.txt` header
   for its inputs). It reads `doi_check.json` + `literature_matrix.xlsx` and writes
   `09`, `10`, `11` and `12`.
4. Verify counts: 34 today; the count should equal the number of `ok: true` records.

---

## 4. How to edit the figures

Both figures are plain SVG with no embedded raster images and are fully editable in
Inkscape, Illustrator, Affinity Designer or a text editor. Text is live text, not paths.
Colours are set inline, so they can be changed with a find-and-replace:

| Element | Fill | Stroke |
|---|---|---|
| Context construct | `#DEEAF6` | `#1F3864` |
| Intervention construct | `#FFF2CC` | `#BF9000` |
| Outcome construct | `#FBE5D6` | `#C55A11` |
| Cross-cutting practice | `#E2EFDA` | `#548235` |

If figure content changes, update the notes in `08_CHAPTER4_CONCEPTUAL_FRAMEWORK.txt`
§4.6 in the same pass, so narrative and figure do not drift apart.

---

## 5. Encoding and formats

| File | Encoding | Notes |
|---|---|---|
| `.md` | UTF-8 | Plain markdown; safe to open in VS Code or Obsidian |
| `.txt` | UTF-8 | Line width kept to about 88 characters for review |
| `.ris` | UTF-8 with BOM | BOM is deliberate: EndNote on Windows reads non-ASCII names more reliably |
| `.bib` | UTF-8 | No BOM; BibTeX dislikes a BOM |
| `.xml` | UTF-8 | No BOM; declared in the XML prolog |
| `.svg` | UTF-8 | XML declaration present |

---

## 6. Package integrity

| Check | Result |
|---|---|
| All 13 numbered deliverables present | Yes |
| Both figures present and well-formed | Yes |
| Reference files contain 34 records each | Yes |
| No file contains a fabricated reference | Yes (built only from `doi_check.json`) |
| No file claims fieldwork was conducted | Yes |
| No file claims a Turnitin or AI-detection score | Yes |

---

## 7. Known gaps carried by this package

1. D10 (AI-use permission) — blocks finalisation.
2. D9 (citation style) — blocks formatting.
3. D11 (Chapter 3 length) — 5,835 words against a 3,500–4,500 working budget.
4. 40 of 49 matrix rows unverified.
5. Chapters 1 and 5 not re-checked.
6. Title page missing.
7. DOCX/PDF build blocked (`python-docx` and `pymupdf` not installed; stale hard-coded paths).

These are recorded in `07_QA_AND_OUTSTANDING_ISSUES.md`, not hidden.

---

## 8. Suggested order of action

1. Read `00_PROJECT_STATUS_REPORT.md`.
2. Read `07_QA_AND_OUTSTANDING_ISSUES.md`.
3. Resolve D9 and D10.
4. Decide D11.
5. Complete source verification (`02`, `12`).
6. Revise Chapters 1 and 5.
7. Build and check the DOCX/PDF.
