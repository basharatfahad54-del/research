# SETILL Thesis-Format Dossier — QA and Gap Report

**Document:** A Systemic Living Lab for a Just Energy Transition: An End-User-Centred Approach to Energy Technology Innovation in South African Informal Settlements
**Prepared:** November 2026
**Status:** Complete build, QA-verified. Unresolved items are listed transparently in §7 — none block submission of the dossier for supervisor review.

---

## 1. Deliverables and file inventory

All files are in `johny\Final_Thesis\`. Original research materials outside `johny` were not modified.

| File | DOCX | PDF |
|---|---|---|
| Chapter 1 — Introduction | ✔ | ✔ |
| Chapter 2 — Literature Review | ✔ | ✔ |
| Chapter 3 — Theoretical Framework | ✔ | ✔ |
| Chapter 4 — Conceptual Framework | ✔ | ✔ |
| Chapter 5 — Research Design and Methodology | ✔ | ✔ |
| Chapter 6 — Expected Contributions | ✔ | ✔ |
| Chapter 7 — Research Plan and Timeline | ✔ | ✔ |
| Chapter 8 — Preliminary Chapter Outline | ✔ | ✔ |
| Conclusion (Chapter 9) | ✔ | ✔ |
| References | ✔ | ✔ |
| Complete_Thesis (all chapters + appendices, auto-updatable TOC field) | ✔ | ✔ |
| Table_of_Contents (static, 172 entries, page numbers refer to Complete_Thesis) | ✔ | ✔ |
| Reference exports: `References.ris`, `References.bib`, `References_EndNote.xml` | — | — |

Total: 12 DOCX + 12 PDF + 3 reference exports.

Formatting applied throughout: Times New Roman; body 12 pt; chapter headings 16 pt bold; section headings 14 pt bold; subsections 12 pt bold; 1.5 line spacing; justified body text; A4 with 2.54 cm margins; centred page numbers in the footer; running header; captioned tables (shaded header row, fixed layout) and figures; per-document cover page and contents page.

## 2. Automated QA results

Full machine-readable results: `johny\thesis_build\qa_results.md`. Summary:

| Check | Result |
|---|---|
| File parity (every DOCX has a PDF; exports present) | Pass — nothing missing |
| Heading sequences (1.1–1.9, 2.1–2.9, 3.1–3.14, 4.1–4.10, 5.1–5.12.8, 6.1–6.7, 7.1–7.5, 8.1–8.3, 9.1–9.4) | Pass — no gaps, no out-of-order sections |
| Table numbering (1.1–1.2; 2.1–2.4; 3.1–3.6; 4.1–4.5; 5.1–5.8; 6.1; 7.1; 8.1) | Pass — sequential, no duplicates |
| Figure numbering (2.1; 4.1–4.2; 5.1; 7.1) | Pass |
| Leftover proposal numbering (Sections 9–12, Chapters 10–12) | Pass — none (one instance of "Section 9" found and fixed during QA; see §5) |
| Excluded content (Umuahia, Nigeria, University of Pretoria, GSTM, Mater Pluviam, Sphokazi) | Pass — none present |
| Build-marker leftovers (`[TABLE]`, `[REF]`, `{widths}` etc.) in rendered DOCX XML | Pass — none |
| Past-tense empirical claims (claim-register phrases, see §4) | Pass — none present |
| Compiled-thesis TOC field | Builds correctly in Word; 172 entries; Chapter 1 begins p. 9 |

## 3. Citation ↔ reference-list parity

**Reference list:** 73 entries, exactly the verified set checked against Crossref/OpenAlex/publisher records during preparation of the original proposal (verification record summarised in Appendix A of the thesis). No source was invented, and no source outside this verified set is cited anywhere in the dossier.

**Both directions verified:**

- *Every in-text citation resolves to a list entry.* An automated scan detected 93 distinct citation instances. The scanner's initial run flagged 19 apparent mismatches; each was manually traced and **all 19 are parse artifacts** of the automated matcher (second authors of two-author citations such as *Alvesson and Kärreman (2007)*, name prefixes such as *von Hippel (2005)*, and organisational authors such as *Statistics South Africa (2023)* and *Presidential Climate Commission (2022)*). Every flagged name/year pair was confirmed present in the list. Zero genuine mismatches remain.
- *Every list entry is cited in the body.* The scanner's "uncited" list (19 candidates) was likewise manually traced; all are first-author halves of narrative citations the regex could not see (e.g. *Braun and Clarke (2006, 2021)*, *Walker and Day (2012)*, *Hennink and Kaiser (2022)*, *Ravitch and Riggan (2017)*). All 73 entries are cited.

**EndNote-compatible exports** were regenerated from the final reference list in clean UTF-8 (73 records each; diacritics intact, e.g. Bergvall-Kåreborn, Ståhlbröst; ASCII-folded citation keys in the BibTeX for maximum tool compatibility; RIS with TY/AU/TI/T2/PY/DO fields; EndNote XML with ref-types for journal articles, books and reports).

## 4. Fieldwork claim register compliance (Decision D2)

The thesis is a **proposal**: no fieldwork has been conducted, no data exist, and every statement about the settlement, participants and outcomes is an intention or planning assumption. All register items C-01…C-19 were enforced:

| Disposition | Register items |
|---|---|
| Past-tense empirical claims converted to proposed tense or dropped | C-01–C-05, C-06–C-12, C-13, C-14 |
| Already-correct planning language retained unchanged | C-15, C-16, C-17 (reworded), C-18, C-19 (Gantt 2027–2029 retained as planning spine) |

An automated scan of the built chapters for the register's signature phrases ("50 households", "were conducted", "data were collected", "survey was done", "multi-stakeholder groups", etc.) returns **zero matches**. The case study is Tzaneen only (D6); the Umuahia (Nigeria) material is excluded entirely.

## 5. Consistency-pass log (edits applied during the build)

1. **Architecture (D3):** proposal Sections 3–12 remapped to thesis Chapters 1–8 + Conclusion + References + Appendices; the case-study chapter folded into methodology as §5.12 (subsections 5.12.1–5.12.8).
2. **§1.9 Structure paragraph** rewritten from the twelve-section proposal description to the actual nine-chapter document structure.
3. **Cross-references** updated throughout ("Section 5" → "Chapter 3", "Section 6" → "Chapter 4", "Section 8" → "Section 5.12", etc.); one residual "Section 9" in Chapter 5 found by QA and corrected to "Chapter 6".
4. **Table renumbering:** part-2 tables shifted to insert new Table 2.2 (Living Lab classifications and the position taken by SETILL); part-4 figures 3.1/3.2 → 4.1/4.2; Table 6.1 → 5.8; Figure 1.1 → 5.1; Tables 7.1/8.1/9.1 → 6.1/7.1/8.1; Figure 5.1 (timeline) → 7.1.
5. **Table 6.1 (contributions) rows 1–2** "Where reported" columns normalised to the eventual-thesis outline ("Chapters 6, 8 and 9"; "Chapters 3, 4 and 9").
6. **Explanatory note after Table 1.1:** chapter references in framework-mapping tables follow the preliminary chapter outline of the eventual thesis (Chapter 8, Table 8.1), not this document's nine-chapter layout — stated once, explicitly.
7. **New sections written from verified sources only:** §1.2 (historical development and conceptual foundations), §2.5 (Living Lab in depth, with Table 2.2), §2.7.1–2.7.2 (biogas technology and implementation experience), §3.8 (systems approaches: systems thinking, system dynamics, systems engineering), §4.8 (SETILL in practice, with Table 4.5 and the antecedents/attributes/consequences reading of components C1–C10), §5.5.8 (internet research), and Chapter 9 (Conclusion) — all drawn from existing research material, all cited within the verified 73-entry set.
8. **End-user-centred reframe (D5):** gender-focused framing replaced by the end-user-centred approach as primary framing, with gender retained as a necessary analytical axis (documented in §1.2.1, §4.8.1, §4.8.3); research questions stated verbatim per D4.
9. **Institutional naming per D1:** University of Johannesburg throughout; references heading de-numbered ("References"); "this proposal" → "this thesis document" in front-matter contexts where the document refers to itself.

## 6. Gap register (verified gaps, not fabricated content)

The following sources/concepts appear in the project's concept notes but **could not be verified** against the bibliographic databases during reference verification. They are **not cited** anywhere in the dossier; the relevant passages rest on general knowledge statements or were dropped. Each is a verification task, not a content invention:

| Gap | Disposition in the dossier |
|---|---|
| Okoye (2020) | Dropped from citations |
| Callaghan & Herselman (2015) | Dropped; the SETILL-structure table (Table 4.5) no longer attributes it, and the source is flagged for verification |
| Arnold & Wade (2015) systems-thinking definition; INCOSE (2011) systems-engineering definition | Dropped; §3.8.3 defines systems engineering without citation, flagged for verification |
| Biogas cluster (Mukumba et al.; Achinas & Euverink; Luostarinen & Normak; Munga; Torres) | Dropped; §2.7 carries the technology description as the study's own account; named case inventory flagged as unverified |
| African biogas case inventory (named implementations in South/East Africa) | §2.7.2 states explicitly that the inventory could not be verified at source and is not cited; targeted verification recorded as an outstanding task |
| Dell'Era & Landoni (2014) definitional quote — page number | The quote is retained (source verified); the page number is unverified and flagged for proofing (§7) |

## 7. Unresolved items requiring supervisor decision

1. **D9 — referencing style.** Harvard (author–date) is used throughout as the interim decision. If the faculty requires a different style (e.g., IEEE numeric), a reformat pass over citations and the reference list will be needed. The RIS/BibTeX/EndNote XML exports support restyling in EndNote without re-verification.
2. **D10 — AI-use position statement.** Not yet drafted; not included in this dossier. A one-paragraph declaration consistent with UJ policy should be added once the position is settled.
3. **UJ thesis template.** The official template was unavailable; formatting follows the agreed specification (TNR, sizes, spacing, margins, page numbers). If a template is issued, content can be pasted in without structural change.
4. **Word count.** The body text (excluding tables, figures, references and appendices) totals approximately **35,900 words** (Ch.1 ≈ 4,380; Ch.2 ≈ 5,280; Ch.3 ≈ 10,020; Ch.4 ≈ 5,960; Ch.5 ≈ 5,650; Ch.6 ≈ 1,280; Ch.7 ≈ 1,170; Ch.8 ≈ 350; Ch.9 ≈ 815). This is substantially below the aspirational 85,000–95,000 thesis target (D11), which is expected: this is a proposal-format dossier, and the eventual thesis expands primarily through findings chapters that do not yet exist (D2). No padding was added to close the gap.
5. **Dell'Era & Landoni (2014) quote page number** — verify against the source at proofing (the only direct quote retained).
6. **Unverified-source verification tasks** — the six rows of the gap register (§6), plus a refresh of the verification record (Appendix A) if sources are added.

## 8. Reproducibility

The dossier is fully regenerable: `johny\thesis_build\build_thesis.py` (DOCX from the six part files + curated supplements), `to_pdf_thesis.ps1` (field-updating PDF export), `dump_toc.ps1` + `make_toc_static.py` (static TOC), `export_refs.py` (reference exports), `qa_thesis.py` (automated QA). Merged chapter sources used for QA are in `johny\thesis_build\merged_*.md`; verification traces in `cite_triage.txt` and `uncited_triage.txt`.
