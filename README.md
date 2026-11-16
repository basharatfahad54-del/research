# SETILL PhD Research Workspace

Working repository for the PhD proposal:

**Developing an Innovative Systemic Living Lab for Just Energy Transition in a South African Informal Settlement (SETILL)**

University of Johannesburg, Engineering Management. Supervisor: Dr K. Aboalez.

This README explains what is in the workspace, how the parts relate, how to run the scripts, and which items are still unverified. It describes the workspace as it actually is, not as it was planned to be.

---

## 1. At a glance

| Item | Detail |
|---|---|
| Institution | University of Johannesburg (decision D1) |
| Document status | **Proposal only. No fieldwork has been conducted** (decision D2) |
| Thesis architecture | Option A: Ch1 Introduction, Ch2 Literature Review, Ch3 Theoretical Framework, Ch4 Conceptual Framework, Ch5 Methodology, Ch6 Results, Ch7 Discussion and Conclusion (D3) |
| Research questions | The brief's five end-user-centred RQs (D4) |
| Gender strategy | Three-tier retention; gender only where a cited source argues it (D5) |
| Case | Tzaneen, South Africa, only (D6) |
| Technology focus | Community-scale biogas (D7) |
| Paradigm | Posthumanist-informed (D8) |
| Primary theory | Critical posthumanism, with systems thinking, energy justice, socio-technical transitions and Living Lab theory as supporting perspectives |
| Referencing | EndNote temporary citations `{Author, Year #ID}`; final style still to be confirmed (D9) |
| AI-use position | Still to be confirmed in writing (D10). The school guideline prohibits AI-generated or AI-rephrased text; disclose as required |
| Core source documents | Concept Note (199 pp, 68,011 words), Molepo et al. (2025), Aboalez (2025), school proposal-structure note |

---

## 2. Status and binding decisions

The binding record is `06_chapter_revisions/decision_record_and_claim_register.md`. The short version:

- The old concept note is a **proposal draft**, not a completed study. It reports fieldwork in the past tense. Those claims are unsupported and are being removed or rewritten as proposed methodology.
- The old draft is **gender-centred** (about 505 mentions of "gender", 335 of "women") and is being rewritten around **end-users**, with exceptions only for cited titles, quotations and sources that genuinely argue gender.
- The old draft mixes **Tzaneen (South Africa)** and **Umuahia (Nigeria)**, and mixes **wind/solar IPP material** with **household biogas**. The current design is Tzaneen only, community-scale biogas.
- The old draft has **no reference list**, two competing chapter-numbering systems, and no theoretical-framework chapter. Chapter 3 is being written new.

Status of the work: Phase 1 (document analysis) is complete. Phase 2–3 (verified reading list and Chapter 3) is in progress. A full plain-text Chapter 3 and an abstract plan now exist in `proposal_v2/`.

---

## 3. Folder structure

```
research_work/
│
├── README.md                         This file
├── PROGRESS_REPORT.md                Phase 1 report: tasks, missing inputs, risks
├── research_topic_analysis.md        Anatomy of the topic; construct-by-construct analysis
├── research_gaps.md                  Gap-spotting analysis and defensibility ratings
├── chapter_revision_plan.md          Chapter-by-chapter issue → revision → justification
├── reference_verification.md         Verification protocol and verified records
├── literature_matrix.xlsx            49-source matrix (5 sheets)
│
├── Aboalez, K (2025)_261102_070901.pdf                  Original source (supervisor)
├── Molepo, P M, Aboalez, K, & Mathaba, T N D (2_...pdf  Original source (supervisor team)
├── PhD DETAILED CONCEPT NOTE 2026_CHECK_261111_113441.pdf  Original concept note (199 pp)
│
├── 00_extracted_text/                Plain-text copies of the PDFs (search only)
│   ├── Aboalez_ K (2025)_261102_070901.txt
│   ├── Molepo_ P M_ Aboalez_ K_ _ Mathaba_ T N D (2_261102_070808.txt
│   ├── PhD DETAILED CONCEPT NOTE 2026_CHECK_261111_113441.txt
│   └── Structure of a Proposal in Engineering Management_261102_135929.txt
│
├── 01_document_inventory/
│   └── document_inventory.md         Master inventory, page map, internal sub-documents,
│                                     figure/table/acronym defects
│
├── 04_theoretical_framework/
│   └── chapter3_structure_and_argument_map.md   13-section Chapter 3 plan and reading list
│
├── 06_chapter_revisions/
│   └── decision_record_and_claim_register.md    Binding decisions D1–D11 and
│                                                unsupported-claim register
│
├── proposal/                         Refined proposal draft (section-numbered 3–12)
│   ├── part1_intro.md                Sections 3.1–3.8: background, problem, aim, RQs,
│   │                                 justification, scope, key terms
│   ├── part2_lit.md                  Section 4: integrative literature review
│   ├── part3_theory.md               Section 5: theoretical framework (~10,000 words)
│   ├── part4_concept.md              Section 6: conceptual framework / SETILL model
│   ├── part5_method.md               Sections 7–8: methodology and Tzaneen case
│   ├── part6_end.md                  Sections 9–12 + appendices: contributions, plan,
│   │                                 thesis outline, references heading, instruments
│   ├── build_proposal.py             Builds the .docx from the markdown-like parts
│   ├── make_figures.py               Generates the five diagrams in assets/
│   ├── to_pdf.ps1                    Converts the .docx to PDF via Word COM
│   └── assets/
│       ├── fig1_design.png
│       ├── fig2_energy.png
│       ├── fig3_setill.png
│       ├── fig4_justice.png
│       └── fig5_timeline.png
│
├── proposal_v2/                      New plain-text deliverables (latest work)
│   ├── 00_abstract_plan.txt          Abstract plan to the school's four-part structure
│   └── 03_chapter3_theoretical_framework.txt   Chapter 3, ~4,490 words, brief-compliant
│
├── preview/                          PNG page previews for visual checks
│   ├── p001–p011, p021–p022, p053–p055, p059–p060, p067–p068
│   └── q001–q005, q009
│
├── ch1_dump.txt                      Page-range dumps of the concept note
├── ch2a_dump.txt                     (page markers: "########## PDF PAGE n ##########")
├── ch2b_dump.txt
├── ch3_dump.txt
├── ch4_dump.txt
├── front_dump.txt                    Front matter (pp. 3–5)
├── cites_raw.txt                     Raw in-text citation patterns and counts
├── setill_dump.json                  Paragraph-level JSON dump of a proposal .docx
│                                     (index, style, alignment, text)
│
├── crossref_verification.json/.txt   Crossref query results (contains HTTP 429
│                                     rate-limit errors — must be re-run)
├── openalex_verification.json/.txt   OpenAlex query results
├── doi_check.json                    DOI-by-DOI Crossref check (field-level metadata)
│
├── extract.py                        Extracts text from all PDFs into 00_extracted_text/
├── dump.py                           Dumps concept-note page ranges to ch*_dump.txt
├── pagemap.py                        Prints first lines and character count per page
├── headings.py                       Lists numbered headings per page
├── cite.py                           Extracts and counts in-text citation patterns
├── d2.py                             Dumps front matter pages
├── h2.py                             Lists chapter-2 headings
├── counts.py                         Counts key terms (posthuman, just transition, gender…)
├── stats.py                          Word counts by page range; figure/table counts
├── build_matrix.py                   Builds literature_matrix.xlsx
├── build_deliverables.py             Builds the client deliverables pack (.docx)
├── build_john_pack.py                Builds the John pack from build_deliverables.py
├── verify_sources.py                 Crossref verification batch 1
├── verify_sources2.py                Crossref verification batch 2
├── verify_openalex.py                OpenAlex verification batch
├── check_dois.py                     Crossref lookup for a fixed DOI list
├── qc_john.py                        QC of the John pack (PDFs, hashes, RIS, pages)
├── qc_ris.py                         QC of the RIS file (completeness, duplicate DOIs)
└── qc_tree.py                        QC of originals and deliverables tree
```

Note: some directories named in `PROGRESS_REPORT.md`'s planned tree (for example `02_literature_review/`, `03_reference_management/`, `05_conceptual_framework/`, `07_abstract/`, `08_figures_and_tables/`, `09_acronyms/`, `10_quality_assurance/`, `references/`, `final_documents/`) do **not** currently exist. Their content lives in the top-level files and in `proposal/` and `proposal_v2/`.

---

## 4. What each group of files is for

### 4.1 Original sources (do not edit)

| File | Pages / size | What it is |
|---|---|---|
| `Aboalez, K (2025)_261102_070901.pdf` | 6 pp, 580 KB | "A Posthumanist Social Approach to Energy Transition", *Procedia CIRP* 135, 8–13, DOI 10.1016/j.procir.2025.01.042. The supervisor's theoretical position. Context is engineering education and social life cycle assessment, **not** informal settlements. |
| `Molepo, P M, Aboalez, K, & Mathaba, T N D (2_...pdf` | 2.25 MB | "Analysis of Barriers to South Africa's Energy Transition", *Energy for Sustainable Development* 88, 101777, DOI 10.1016/j.esd.2025.101777. 15 expert interviews, AKTESP framework. About wind/solar IPPs and the grid, not households or Living Labs. |
| `PhD DETAILED CONCEPT NOTE 2026_CHECK_261111_113441.pdf` | 199 pp, 4.06 MB | The old proposal draft. About 68,000 words. Gender-centred; includes unsupported past-tense fieldwork claims. |

The original `Structure of a Proposal in Engineering Management_...pdf` is **not** in this folder. Its extracted text is in `00_extracted_text/`. The structure note fixes the abstract at 250–300 words covering problem, objectives, methodology and contributions.

### 4.2 Extracted text

`00_extracted_text/` holds plain-text copies with `=== PAGE BREAK ===` markers. They are for searching only. Page numbers in citations should be checked against the original PDFs.

### 4.3 Analysis and planning documents

| File | Purpose |
|---|---|
| `01_document_inventory/document_inventory.md` | Master inventory of the sources, page-by-page map of the concept note, internal sub-documents and numbering collisions, figure/table/acronym defects. |
| `research_topic_analysis.md` | Breaks the title into load-bearing constructs, states what each commits the study to, and where the old draft stands. Also addresses the "end-user" construct and the gender-removal problem. |
| `research_gaps.md` | Gap-spotting strategy (confusion, neglect, application spotting) and defensible gap statements. |
| `chapter_revision_plan.md` | Chapter-by-chapter plan with issue → revision → justification. Sets out Option A architecture. |
| `reference_verification.md` | Verification protocol and the records checked so far. **Only a small number of records were verified; most remain unverified.** |
| `06_chapter_revisions/decision_record_and_claim_register.md` | Binding decisions D1–D11 and the register of unsupported empirical claims that must be removed. |
| `04_theoretical_framework/chapter3_structure_and_argument_map.md` | 13-section plan for a long-form Chapter 3 (18,000–22,000 words), with the governing argument, section briefs, core reading list and risks. This is a planning document; the shorter brief-compliant Chapter 3 is in `proposal_v2/`. |

### 4.4 Proposal drafts (`proposal/`)

`proposal/` is the long-form refined draft, numbered against the school's structure items (3 Introduction … 12 References), with appendices A–F. Word counts: part1 ~3,999; part2 ~4,065; part3 ~10,082; part4 ~5,254; part5 ~7,121; part6 ~6,906 (about 37,400 words total). It is a draft for revision, not a submission copy.

`build_proposal.py` assembles the `.docx`; `make_figures.py` generates the diagrams; `to_pdf.ps1` exports the PDF using Microsoft Word.

### 4.5 Latest plain-text deliverables (`proposal_v2/`)

| File | Content |
|---|---|
| `00_abstract_plan.txt` | A plan, not the final abstract. Follows the school's four elements and allocates 250–300 words. Includes placeholders for the candidate's local observation and confirmed sample size. |
| `03_chapter3_theoretical_framework.txt` | Complete Chapter 3 at ~4,490 words including Table 3.1. Covers (a) why theory matters generally, (b) the main theories, (c) their relation to the problem, (d) justification against all five RQs. Includes reasoned reject/reject/accept for TAM, Lean Thinking and Complexity Theory, and a bridge to Chapter 4. Ends with editorial notes on `[VERIFY]` items and gender-term decisions. |

### 4.6 Verification data

| File | What it contains | Caveat |
|---|---|---|
| `crossref_verification.txt/.json` | Crossref query results for theory sources | Contains `HTTP 429 Too Many Requests` errors; re-run with delays |
| `openalex_verification.txt/.json` | OpenAlex results | Check each DOI resolves before import |
| `doi_check.json` | Field-level Crossref metadata for a DOI list | Trust only records with `"ok": true` |
| `literature_matrix.xlsx` | 49 sources across 5 sheets | Source of truth for the reading list, not for citation formats |

---

## 5. How the pieces fit together

```
Original PDFs
   │  extract.py / dump.py
   ▼
00_extracted_text  +  ch*_dump.txt  +  setill_dump.json
   │  inventory, counting, citation extraction
   ▼
Analysis docs:  document_inventory  ·  research_topic_analysis
                research_gaps        ·  chapter_revision_plan
                reference_verification · decision_record
   │
   ├──► 04_theoretical_framework/  (long-form Chapter 3 plan)
   │
   ├──► proposal/  (long-form refined draft, sections 3–12)
   │        └── build_proposal.py → .docx → to_pdf.ps1 → .pdf
   │
   └──► proposal_v2/  (brief-compliant plain text; latest)
            00_abstract_plan.txt
            03_chapter3_theoretical_framework.txt
```

---

## 6. Running the scripts

### Requirements

- Python 3.10+ (the existing scripts were run under a recent CPython).
- Packages: `pymupdf`, `openpyxl`, `python-docx`, `Pillow`.
- Microsoft Word (only for `proposal/to_pdf.ps1`, which uses Word COM).
- PowerShell 5.1 or later on Windows.

Install the Python packages if needed:

```powershell
python -m pip install pymupdf openpyxl python-docx Pillow
```

### Important path caveat

Most scripts contain **hard-coded absolute paths** pointing to the previous location:

```
C:\Users\Muhammad Fahad\Desktop\Research paper
```

The current workspace is:

```
C:\Users\fahad\OneDrive\Desktop\research_work
```

Before running any script, update its `ROOT`, `src`, `f`, `BASE` or `OUT` variable. `proposal/build_proposal.py`, `proposal/make_figures.py`, `proposal/to_pdf.ps1`, `build_deliverables.py`, `build_john_pack.py` and the `qc_*.py` scripts are the ones most affected.

### Typical commands

```powershell
# Re-extract text from the three PDFs (after fixing src in extract.py)
python extract.py

# Dump concept-note page ranges (after fixing f in dump.py)
python dump.py

# Rebuild the literature matrix
python build_matrix.py

# Rebuild the long-form proposal .docx (after fixing ROOT)
python proposal\build_proposal.py

# Export to PDF (requires Word)
powershell -ExecutionPolicy Bypass -File proposal\to_pdf.ps1

# Re-check DOIs (network required; Crossref rate-limits)
python check_dois.py
```

Run verification scripts with pauses between requests. Crossref returns `HTTP 429` if called too quickly; this is why `crossref_verification.txt` contains errors.

---

## 7. Conventions and integrity rules

1. **Originals are never modified.** All work products are new files. Do not overwrite the three source PDFs.
2. **No invented references.** If a detail (author order, year, volume, pages, DOI) is uncertain, mark it `[VERIFY]` rather than guessing.
3. **No invented findings.** There is no fieldwork. Methods are written as proposed, and any past-tense empirical claim in the old draft is a defect to be removed.
4. **Quotations are short, marked and page-referenced.** Paraphrase by default.
5. **EndNote temporary citations** use `{Author, Year #ID}`. The `#ID` numbers in `proposal_v2/` are placeholders and must be renumbered against the real EndNote library.
6. **No promises about Turnitin or AI-detection scores.** Run Turnitin through the University's licensed account; AI-detection scores are probabilistic and do not certify authorship.
7. **AI-use disclosure.** The school's proposal guideline prohibits AI tools for generating or rephrasing text. Confirm the supervisor's and faculty's written position and disclose assistance as required. Drafts in this workspace are assistance, not a substitute for the candidate's authorship.
8. **Gender language.** Use "end-users" or "household members" in the proposal's own text. Exceptions: titles of cited works, journal names, verbatim quotations, and sources that genuinely argue a gender-specific point. Flag any replacement that changes meaning.
9. **Encoding.** Some markdown files were written on a system that mangled em dashes and curly quotes (they may display as `�?"`). The underlying text is still readable; re-save as UTF-8 if you edit them.

---

## 8. Known gaps and next steps

### Decisions still outstanding

- D9: referencing style (Harvard or APA 7) for EndNote output.
- D10: written position on AI assistance.
- Faculty confirmation of chapter and total word limits. The 15,000–20,000-word budget in the current brief is a working budget, not a confirmed school limit.

### Work still to do

1. Reconcile the SETILL expansion. The old concept note says "Systems Energy Technology Innovation Living Lab"; the brief says "Systemic Living Lab for Just Energy Transition". Recommendation: use the brief's expansion, keep the acronym, and add one footnote.
2. Rebuild the reference list as an EndNote-importable `.ris`, verifying every DOI. Most of the roughly 250–300 old citations are still unverified.
3. Rebuild Chapter 4 (conceptual framework) from the five propositions in the new Chapter 3.
4. Revise Chapter 1 and Chapter 2 to the five new RQs and the shorter word budget.
5. Remove or rewrite all unsupported past-tense fieldwork claims (register in `decision_record_and_claim_register.md`).
6. Build the front matter: List of Figures, List of Tables, List of Acronyms, each on its own page, numbered to match the chapters.
7. Resolve the theory–method coherence issue: a posthumanist ontology with a constructivist-interpretivist epistemology needs an explicit reconciliation in the methodology chapter.
8. Define "acceptance" once. The old draft uses acceptance, adoption, uptake, take-up and acceptability interchangeably.

### Known document defects in the source

- No reference list in the concept note; section 8 is an empty heading.
- Two competing chapter-numbering systems.
- Broken table of contents and blank abstract.
- Figures 3.1–3.5 are raster images of variable legibility with no source notes.
- Corrupted citation strings, including "Sphokazi Phelokazi (2022)", "Mosango et al.", "Mamphweli (2016)", "Nations (2014)", "Agenda (2030)", "Filho et al. (2022)", "Agri (2023)", "Gates (2019)".
- A Melinda Gates quotation (p. 28) that cannot currently be traced.

---

## 9. Change log

| Date | Change |
|---|---|
| 12 Nov 2026 | Phase 1 complete: extraction, inventory, topic and gap analysis, revision plan, literature matrix, Chapter 3 plan. Decisions D1–D8 and D11 recorded. |
| 12 Nov 2026 | Long-form proposal draft (`proposal/`, sections 3–12) built and exported. |
| 14 Nov 2026 | New brief-compliant plain-text work added in `proposal_v2/`: abstract plan and a ~4,490-word Chapter 3. |
| 14 Nov 2026 | This README added. |
| 14 Nov 2026 | `proposal_v3/03_chapter3_theoretical_framework.txt` written: Chapter 3 revised against all four school requirements (importance of a framework, main theories, relation to the problem, justification tied to the RQs). ~5,835 body words; length decision (D11) open. |
| 14 Nov 2026 | `final_deliverables/` populated: status report, Chapter 3 compliance review, source verification report, 267-word abstract, front matter, terminology change log, school compliance matrix, QA register, Chapter 4 (~2,092 words), two editable SVG figures, RIS/BibTeX/EndNote reference files (34 verified records) and a coverage report. |
| 14 Nov 2026 | QA fixes: Chapter 4 construct codes C1–C10 no longer collide with SETILL components (renamed Component 1–10); figure notes reconciled with the SVGs; status-report reference counts corrected to 9 verified matrix rows / 34 verified records / 40 unverified rows. |

---

## 10. Quick reference: key facts to keep consistent

- Title: **Developing an Innovative Systemic Living Lab for Just Energy Transition in a South African Informal Settlement**.
- Acronym: **SETILL**.
- Case: Tzaneen, South Africa.
- Technology: community-scale biogas.
- Five RQs: exactly as stated in the brief (systems framing; participation factors; cultural and technological acceptance; mechanisms; Living Lab evaluation).
- Primary theory: critical posthumanism (Braidotti, Barad, Haraway, Hayles), with the supervisor's `potestas`/`potentia` and posthumanist cartography.
- Supporting: systems thinking, energy justice, socio-technical transitions, Living Lab theory, complexity theory.
- Rejected as core: Technology Acceptance Model, Lean Thinking.
