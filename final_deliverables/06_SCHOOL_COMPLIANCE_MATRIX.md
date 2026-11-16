# School Compliance Matrix

**Controlling document:** *Structure of a research proposal in Engineering Management
(PhD/M(RD))*, 12 numbered sections.
**Package audited:** long-form draft (`proposal/part1`–`part6`), new chapter drafts
(`proposal_v3/03`, `final_deliverables/08`), abstract, front matter, reference files.

**Two architectures are in play.** The long-form draft is numbered 3–12 because it
already follows the school's section order (sections 1 and 2 are the title page and
abstract). The adopted proposal architecture (Option A) regroups the same material into
five chapters. Both mappings are shown so nothing is lost in translation.

---

## 1. Section-by-section compliance

| # | School requirement | Long-form draft | 5-chapter location | Status |
|---|---|---|---|---|
| 1 | Title (cover) page | — | — | **Missing.** No cover page in the package. |
| 2 | Abstract, 250–300 words, four elements | `03_ABSTRACT.txt` | Front matter | **Draft complete** — 267 words, all four elements present. |
| 3 | Introduction: background/context; problem statement; aims and objectives; RQs/hypotheses; justification | `part1_intro.md` §§3.1–3.5 | Chapter 1 | **Draft exists; alignment not verified.** |
| 4 | Literature review: recent accredited peer-reviewed sources; theoretical, empirical and review studies; gap spotting; positioning | `part2_lit.md` §§4.1–4.8 | Chapter 2 | **Draft exists; source verification incomplete** (40 of 49 matrix rows unverified). |
| 5 | Theoretical framework: importance (general); main theories; relation to problem; justification tied to RQs and expected findings | `part3_theory.md` §§5.1–5.13; **superseded by** `proposal_v3/03` | Chapter 3 | **Draft complete** — all four requirements met (see `01_CHAPTER3_COMPLIANCE_REVIEW.md`). |
| 6 | Conceptual framework: visual and narrative; key variables/constructs and relationships; own labelled diagram with flow, feedback loops or process dynamics | `part4_concept.md` §§6.1–6.9; **superseded by** `final_deliverables/08` + two SVGs | Chapter 4 | **Draft complete** — narrative, Table 4.1, Figure 4.1, Figure 4.2. |
| 7 | Research design and methodology: paradigm and approach; strategy; data collection and instruments; sampling and participants; analysis; validity/trustworthiness; ethics; delimitations and limitations | `part5_method.md` §§7.1–7.11 | Chapter 5 | **Draft exists; not re-checked against the new theory.** |
| 8 | Context and case study (if applicable): description of setting; relevance to Engineering Management | `part5_method.md` §§8.1–8.8 | Chapter 5 (or Chapter 1) | **Draft exists.** |
| 9 | Expected contributions: theoretical; practical; policy/sector; methodological | `part6_end.md` §§9.1–9.7 | Chapter 5 / Chapter 1 | **Draft exists; must match Chapter 3 and 4 claims.** |
| 10 | Research plan / timeline | `part6_end.md` §§10.1–10.5 | Chapter 5 or appendix | **Draft exists.** |
| 11 | Preliminary chapter outline | `part6_end.md` §§11.1–11.3 | Chapter 5 or appendix | **Draft exists; must be updated to Option A.** |
| 12 | References: reference manager; cross-check citations; no AI-generated text or fake references; DOI-traceable; accredited UJ sources; no predatory sources | `09_references.ris`, `10_references.bib`, `11_references_endnote.xml`, `12_reference_coverage_report.txt` | References | **Partial / blocked** — files exist but cover only verified records; citation placeholders remain. |

---

## 2. Section 5 (theoretical framework) sub-requirements

| Sub-requirement | Where met | Status |
|---|---|---|
| General explanation of why a theoretical framework matters (not topic-specific), with references | Ch3 §3.1–§3.3 | Met (draft) |
| Explanation of the main theories/models | Ch3 §3.4, §3.5, §3.10 | Met |
| How the theories relate to the research problem | Ch3 §3.7 | Met |
| Justification of theory selection and how it informs RQs and expected findings | Ch3 §3.6, §3.9, §3.10, §3.11 | Met |
| The school's examples (TAM, Lean Thinking, Complexity Theory) explicitly handled | Ch3 §3.10 and §3.5.5 | Met |

---

## 3. Section 6 (conceptual framework) sub-requirements

| Sub-requirement | Where met | Status |
|---|---|---|
| Visual representation | Figure 4.1 (SVG) | Met |
| Narrative representation | Ch4 §4.1–§4.9 | Met |
| Key variables/constructs | Ch4 Table 4.1 (C1–C10) | Met |
| Presumed relationships | Ch4 §4.4 (R1–R8) | Met |
| Own well-labelled diagram showing flow of influence | Figure 4.1 | Met |
| Feedback loops or process dynamics | Figure 4.2 (learning and justice loops) | Met |
| Derived from literature review and theoretical framework | Ch4 §4.1; cross-references to Ch2 and Ch3 | Met |

---

## 4. Section 12 (references) sub-requirements

| Sub-requirement | Current state | Status |
|---|---|---|
| Use a reference manager (EndNote) to auto-format citations | EndNote XML provided; in-text placeholders not yet field codes | **Partial** |
| Cross-check all in-text citations against the bibliography | Not done; chapters 1, 2 and 5 still use placeholder IDs | **Not done** |
| No AI-generated or rephrased text | Prohibited by the school; D10 outstanding | **Blocked** |
| No fake references | Reference files built from verified Crossref metadata only | **Held** |
| References traceable through DOI links | 34 records carry DOIs; books will need ISBN/catalogue links | **Partial** |
| Only accredited UJ-database sources; no predatory or non-peer-reviewed sources | Book and journal sources only; predatory status not yet checked per source | **Partial** |

---

## 5. Other school requirements

| Requirement | Status |
|---|---|
| Recent accredited peer-reviewed literature | **Partial** — recent sources (2020–2025) are present; recency profile not yet audited across Chapter 2. |
| Gap spotting strategy | **Draft** — Chapter 2 §4.8 and the gap map in `literature_matrix.xlsx`. |
| Positioning in the existing body of knowledge | **Draft** — Chapter 2 §4.8. |
| Engineering Management relevance | **Draft** — Chapter 1 §3.1.4; Chapter 5 §8.6. |
| Ethics section | **Draft** — Chapter 5 §7.8; consent summary in Appendix C. |
| Delimitations and limitations | **Draft** — Chapter 1 §3.6; Chapter 5 §7.10. |

---

## 6. Compliance verdict

| Domain | Verdict |
|---|---|
| Structure (all 12 sections present somewhere in the package) | **Met at draft level** — only the title page is entirely missing. |
| Section 5 | **Met at draft level.** |
| Section 6 | **Met at draft level.** |
| Section 12 | **Not met** — blocked on D9, D10 and source verification. |
| Whole-document readiness | **Not submission-ready.** |

---

## 7. Actions required

1. Create the title page (section 1).
2. Resolve D9 (citation style) and D10 (AI position).
3. Complete verification for the 40 unverified matrix rows.
4. Replace all in-text citation placeholders with EndNote field codes.
5. Re-check Chapters 1, 2 and 5 against the revised Chapter 3 and new Chapter 4.
6. Update the preliminary chapter outline to the adopted 5-chapter architecture.
7. Audit recency and predatory status of every source.
