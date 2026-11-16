# QA and Outstanding Issues

**Date:** 14 November 2026
**Scope of QA:** Chapter 3 (v3), Chapter 4, abstract, front matter, reference files and
the two SVG figures.
**Status of all prose:** AI-assisted working draft. Nothing is submission-ready.

---

## 1. Checks performed

| Check | Method | Result |
|---|---|---|
| Chapter 3 covers the school's four requirements | Section-by-section mapping | **Pass** (`01_CHAPTER3_COMPLIANCE_REVIEW.md`) |
| Chapter 3 section numbering | Headings audit (§3.1–§3.12) | **Pass** |
| Chapter 3 table numbering | Caption audit | **Pass** (Table 3.1, 3.2 unique) |
| Chapter 3 word count | Word count of §3.1–§3.12 | **5,835 words** (budget 3,500–4,500; see §3 below) |
| Chapter 4 construct codes | Table 4.1 vs component list | **Fail → fixed** (see §2) |
| Chapter 4 figure notes vs figures | Note audit | **Fail → fixed** (see §2) |
| Chapter 4 word count | Count | **2,132 words** (within a typical chapter range) |
| Abstract word count | Count | **267 words** (limit 250–300) |
| Abstract covers four elements | Manual check | **Pass** |
| SVG files well-formed | `XmlDocument.Load` | **Pass** (both roots `svg`) |
| Reference files parse and count | Entry-count check | **Pass** (34 records in each of RIS, BibTeX, EndNote XML) |
| Reference files complete | Coverage check | **Fail** — 40 of 49 matrix rows unverified (`12_reference_coverage_report.txt`) |
| Fabricated empirical claims | Read-through | **Pass** — no fieldwork claimed |
| Fabricated references | Trace to `doi_check.json` | **Pass** — only Crossref-verified metadata used |
| Old-path dependencies in build scripts | File audit | **Fail** — hard-coded `C:\Users\Muhammad Fahad\...` |
| DOCX/PDF build | Import test | **Blocked** — `python-docx` and `pymupdf` not installed |

---

## 2. Defects found and corrected during QA

| # | Defect | Severity | Resolution |
|---|---|---|---|
| 1 | Chapter 4 used **C1–C10 twice**: once for constructs (Table 4.1) and once for SETILL components (§4.5 and Figure 4.2). This would have produced a genuinely confusing document. | High | Components recoded **LL1–LL10** (Living Lab components). Constructs keep **C1–C10**. §4.5, §4.6 and both figure notes updated; Figure 4.2 SVG relabelled. |
| 2 | Figure 4.1 note 3 said C10 had a "double outline"; the SVG uses a dashed outline. | Low | Note corrected to "dashed outline". |
| 3 | Figure 4.1 boundary note cited "C1" for boundary critique, which is Component 1, not construct C1. | Medium | Note and SVG label corrected to "Component 1". |
| 4 | Chapter 3 v3 initially carried a stale word count (5,450). | Low | Corrected to the measured **5,835** body words. |
| 5 | Status report initially implied 34 of 49 matrix rows were DOI-verified. | High | Corrected: **9** matrix rows DOI-verified; the 34 records in the files include 25 theory sources not yet mapped to a matrix row. |
| 6 | Source verification categories in the matrix mix "cited in draft", "verify" and "DOI verified". | Medium | Reported as three explicit levels (V1/V2/V3) in `02_SOURCE_VERIFICATION_REPORT.md`. |

---

## 3. Open issues register

| # | Issue | Decision | Severity | Owner | Effect if unresolved |
|---|---|---|---|---|---|
| 1 | AI-use permission not confirmed | **D10** | Blocking | Candidate + supervisor | Cannot declare any prose submission-ready |
| 2 | Citation style unconfirmed | **D9** | High | Candidate + school | Reference formatting may have to be redone |
| 3 | Chapter 3 length | **D11** | Medium | Candidate + supervisor | 5,835 words vs a 3,500–4,500 working budget; school sets no limit |
| 4 | 40 matrix rows unverified | — | High | Candidate | Full bibliography cannot be finalised |
| 5 | 2 DOI lookups failed | — | Low | Candidate | Two sources missing from reference files |
| 6 | In-text citation placeholders | — | High | Candidate | Citations will not render; cross-check impossible |
| 7 | SETILL expansion footnote (D8) | — | Low | Candidate | Acronym change unexplained to the reader |
| 8 | Chapters 1 and 5 not re-checked | — | High | Candidate | Possible contradiction with Chapter 3 and Chapter 4 |
| 9 | Figure/table numbering has two schemes | — | Medium | Candidate | Duplicate or missing numbers in the submitted document |
| 10 | Title page missing | — | Medium | Candidate | Section 1 of the required structure absent |
| 11 | DOCX/PDF build blocked | — | Medium | Candidate/IT | No formatted output |
| 12 | Predatory-source audit not done | — | High | Candidate | Section 12 requirement not met |
| 13 | Recency audit of Chapter 2 not done | — | Medium | Candidate | "Recent accredited" requirement may be questioned |

---

## 3b. Static scan of the long-form Chapters 1, 2 and 5

A pattern scan was run over `proposal/part1_intro.md` to `part6_end.md` (≈68,000 words)
for the defects that made the old concept note unsafe to reuse.

| Pattern | Occurrences |
|---|---|
| "interviews/workshops were conducted" and similar past-tense fieldwork claims | 0 |
| "the study found/showed/revealed/demonstrated" | 0 |
| "Umuahia" | 0 |
| "Nigeria" | 0 |
| "proves", "guarantees", "will demonstrate", "first study to" | 0 |
| Unyielded citation placeholders (`{Author, 2020 #12}`) | 0 |
| "gender" / "women" | 46 (substantive use, not contamination) |

**Reading:** the contamination found in the old concept note (Umuahia/Nigeria, wind and
solar IPP focus, past-tense fieldwork) is **not** present in the long-form draft. The
remaining alignment risk in Chapters 1 and 5 is conceptual (whether their framing
matches the revised Chapter 3 and new Chapter 4), not factual contamination. That risk
still requires a human read; it cannot be settled by pattern matching.

---

## 4. Claims I could not verify (and have not asserted)

These were left out of the drafts rather than guessed:

- Exact page numbers for many sources (only Crossref page ranges where available).
- Whether every "TBC" matrix row corresponds to the title described in the matrix.
- Whether the Aboalez (2025) paper supports every use made of it beyond the extraction
  already in the workspace. Chapter 3 §3.8 limits the claim accordingly.
- Quantitative claims about Tzaneen (population, household counts, energy mix) beyond
  what the case-study draft already states.
- Any Turnitin or AI-detection score. No such score is claimed anywhere.

---

## 5. What the candidate must do before submission

1. Obtain **D10** in writing.
2. Confirm **D9**.
3. Decide **D11** (keep the full Chapter 3 or cut to budget).
4. Verify the 40 unverified sources at source.
5. Replace all citation placeholders with EndNote field codes and cross-check.
6. Re-read Chapters 1, 2 and 5 against the revised Chapter 3 and new Chapter 4.
7. Insert the SETILL expansion footnote.
8. Add the title page.
9. Unify figure and table numbering at assembly.
10. Install `python-docx` and `pymupdf`, fix the hard-coded paths, build DOCX and PDF.
11. Read the school's AI statement again and comply with it literally.
