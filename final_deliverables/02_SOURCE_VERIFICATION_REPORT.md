# Source Verification Report

**Date:** 14 November 2026
**Sources:** `literature_matrix.xlsx` (Sources sheet, 49 rows), `doi_check.json`,
`crossref_verification.json`, `openalex_verification.json`, workspace PDFs
**Generated reference files:** `09_references.ris`, `10_references.bib`,
`11_references_endnote.xml` (34 records each)

---

## 1. Headline position

| Measure | Count |
|---|---|
| Matrix rows total | 49 |
| Rows with DOI-level verification (Crossref) | 9 |
| Rows with full text held in the workspace | 2 |
| Rows with no verified identifier ("TBC" / "verify") | 40 |
| DOI-verified records written to the reference files | 34 |
| … of which linked to a matrix row | 9 |
| … of which not yet mapped to a matrix row | 25 |
| DOI lookups that failed | 2 |

**Interpretation.** The reference files are valid as a set of verified records, but they
are **not** a complete bibliography of the matrix. 40 matrix rows cannot be cited
until their identifiers are checked at source. This is a deliberate constraint: no
reference is written into the files on the basis of memory or a guessed DOI.

---

## 2. What "verified" means here

Three different levels are in use in the matrix, and they should not be conflated:

| Level | Meaning | Rows |
|---|---|---|
| **V1 — Full text in workspace** | PDF obtained and read | 2 (L01, L02) |
| **V2 — Metadata verified** | DOI resolves through Crossref; title, authors, year, journal and pages confirmed | 9 (the two above plus L13, L14, L15, L16, L47, L48, L49) |
| **V3 — Unverified** | Cited in a draft but no identifier checked; title may be a working description rather than the exact published title | 40 |

Only V1 and V2 rows are safe to cite. V3 rows must not be treated as confirmed.

---

## 3. The nine verified matrix rows

| ID | Source | DOI | Level |
|---|---|---|---|
| L01 | Aboalez (2025), *A Posthumanist Social Approach to Energy Transition* | 10.1016/j.procir.2025.01.042 | V1 + V2 |
| L02 | Molepo, Aboalez & Mathaba (2025), *Analysis of Barriers to South Africa's Energy Transition* | 10.1016/j.esd.2025.101777 | V1 + V2 |
| L13 | Musango et al. (2020), *Mainstreaming gender to achieve security of energy services* | 10.1016/j.erss.2020.101715 | V2 |
| L14 | Mukama et al. (2022), *Development of living labs to support gendered energy transitions* | 10.1016/j.techsoc.2021.101850 | V2 |
| L15 | Anditi et al. (2022), *Addressing gender dimensions in energy innovations* | 10.1016/j.erss.2021.102476 | V2 |
| L16 | Petrulaityte et al. (2022), *Supporting the Development of Gendered Energy Innovations* | 10.3390/su14106291 | V2 |
| L47 | Batinge et al. (2022), *Modelling gendered innovation for the security of energy services* | 10.1002/sres.2793 | V2 |
| L48 | Musango (2022), *Assessing gender and energy in urban household energy transitions* | 10.1016/j.erss.2022.102525 | V2 |
| L49 | Oosthuizen, de Kock & Musango (2020), *Exploring factors that influence the mainstreaming of gender* | 10.7166/31-3-2421 | V2 |

**Note on L13–L16, L47–L49.** These are the gender-and-energy sources from the Musango
group. They are legitimate accredited sources and are useful for the recognition and
participation dimensions. They do **not** make gender the organising lens of this study;
see `05_TERMINOLOGY_CHANGE_LOG.md` for how they should be positioned.

---

## 4. The two failed DOI lookups

| DOI | Error | Likely cause | Action |
|---|---|---|---|
| 10.1086/345321 | HTTP 429 Too Many Requests | Rate limiting, not a bad DOI | Re-run after a delay; if it resolves, add the record |
| 10.1177/1350508410373519 | HTTP 404 Not Found | DOI does not resolve, or was transcribed incorrectly | Check the source document; find the correct DOI or drop the source |

Neither record appears in the generated reference files. That is correct: an
unresolved DOI must not be presented as verified.

---

## 5. The 40 unverified matrix rows

These fall into four groups. The grouping matters because the remedy differs.

### 5.1 Classic theory and methods books (12 rows)

L03 Braidotti (2013) *The Posthuman*; L04 Braidotti (2019) *Posthuman Knowledge*;
L06 Haraway (2016) *Staying with the Trouble*; L07 Barad (2007) *Meeting the Universe
Halfway*; L09 Hayles (1999) *How We Became Posthuman*; L10 Wolfe (2010) *What Is
Posthumanism?*; L30 Checkland (1981) *Systems Thinking, Systems Practice*; L31 Meadows
(2008) *Thinking in Systems*; L32 Midgley (2000) *Systemic Intervention*; L39 Lincoln &
Guba (1985) *Naturalistic Inquiry*; L41 Miles & Huberman (1994) *Qualitative Data
Analysis*; L46 Swilling & Annecke (2012) *Just Transitions*.

**Remedy:** verify publisher, edition and ISBN from the UJ library catalogue. These are
books, so a DOI is not always available; the ISBN or catalogue record is the correct
identifier. Chapter 3 relies on several of them, so this is the highest-priority group.

### 5.2 Journal articles with a known title (15 rows)

L05 Haraway (1988); L08 Barad (2003); L11 Forlano (2017); L12 Ulmer (2017);
L20 Stephenson et al. (2010); L22 Sovacool & Dworkin (2015); L23 Jenkins et al. (2016);
L24 Walker & Day (2012); L25 Geels (2002); L26 Geels (2011); L27 Davis (1989);
L28 Venkatesh et al. (2003); L29 Arnstein (1969); L37 Wüstenhagen et al. (2007);
L45 Heffron & McCauley (2017).

**Remedy:** locate the DOI through the journal or Crossref and record it. Several are
old and well known; the risk is a wrong volume/page, not a wrong paper.

### 5.3 Rows with a working title rather than the published title (6 rows)

L17 Okoye (2020); L18 Kovacic et al. (2019); L19 Callaghan & Herselman (2015);
L21 Shrestha & Acharya (2015); L36 Schuurman et al. (2015); L42 Booth & Bennett (2002).

**Remedy:** the matrix title is a description, not necessarily the published title. Find
the actual record before citing. L17 in particular is flagged as possibly a thesis.

### 5.4 Method and design-literature rows (7 rows)

L33 Leminen et al. (2012); L34 Dell'Era & Landoni (2014); L35 Hossain et al. (2019);
L38 Blessing & Chakrabarti (2009); L40 Easterby-Smith et al. (2012); L43 Postma (2016);
L44 Motala (2017).

**Remedy:** verify edition (for the books) and DOI (for the articles).

---

## 6. Cross-check against the old concept note (D-01)

Five rows are recorded as cited in D-01 but unverified (L03, L04, L07, L11, L12, L43,
L44). The old concept note has **no reference list**, so these citations cannot be
checked internally. Each must be verified against the published source before it is
carried into the new proposal. A source being cited in the old draft is not evidence
that it says what the old draft claims.

---

## 7. What must happen next

1. Re-run the two failed DOI lookups (429 is transient).
2. Verify the twelve classic theory and methods books against the UJ catalogue (ISBN/edition).
3. Resolve the six working-title rows (L17–L19, L21, L36, L42).
4. Verify the remaining journal articles and record the DOI.
5. Update the matrix Verification column for each row.
6. Re-run the generator so the reference files cover the full verified set.
7. Only then produce the final bibliography.

**Do not** state in the proposal that all sources are verified. The correct statement
at this stage is that the reference list is a working list of verified records plus a
set of sources still under verification.

---

## 8. Data-quality repairs made during generation

Three repairs were applied so that the files are usable as they stand:

1. **Blank author fields.** Crossref returned two empty author entries (`, `) for
   `10.5929/2014.4.2.9`. These were removed; the record now lists Grant and Osanloo
   only. An empty author field breaks EndNote and BibTeX parsing.
2. **Page ranges.** Where Crossref returned `404-428`, the RIS file now writes
   `SP - 404` and `EP - 428` separately rather than one ambiguous `SP` field.
3. **Stale header.** The BibTeX header now states the actual record count (34), not
   the earlier 36.

No title, author or year was altered. If a repair is disputed, correct it at source in
`doi_check.json` and regenerate; do not hand-edit the output files.
