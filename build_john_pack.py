# -*- coding: utf-8 -*-
"""
Builds JOHN_PHD_RESEARCH_DELIVERABLES by reusing the verified document
generators, remapping folder names to the structure requested for this pack.
"""
import os
import shutil
import sys

ROOT = r"C:\Users\Muhammad Fahad\Desktop\Research paper"
NEW = os.path.join(ROOT, "JOHN_PHD_RESEARCH_DELIVERABLES")
os.environ["DELIV_OUT"] = NEW

sys.path.insert(0, os.path.join(ROOT, "research_work"))
import build_deliverables as bd

bd.FOLDER_MAP.update({
    "01_Theoretical_Framework": "02_Chapter_3_Theoretical_Framework",
    "02_Conceptual_Framework": "03_Conceptual_Framework",
    "03_Revised_Research_Chapters": "01_Research_Chapters",
    "04_Abstract": "04_Abstract",
    "05_Complete_Research_Document": "05_Complete_Research_Document",
    "06_References_EndNote": "06_References",
    "07_Figures_Tables_Acronyms": "07_Front_Matter",
    "09_Final_Handover": "09_Handover",
})

DIRS = [
    "01_Research_Chapters",
    "02_Chapter_3_Theoretical_Framework",
    "03_Conceptual_Framework",
    "04_Abstract",
    "05_Complete_Research_Document",
    "06_References",
    "06_References\\Article_PDFs",
    "07_Front_Matter",
    "07_Front_Matter\\Figures",
    "07_Front_Matter\\Tables",
    "08_Quality_Assurance",
    "09_Handover",
]


def make_dirs():
    for d in DIRS:
        os.makedirs(os.path.join(NEW, d), exist_ok=True)


def copy_article_pdfs():
    """Copy the two lawfully held article PDFs. Originals are left untouched."""
    src = ROOT
    items = [
        ("Aboalez, K (2025)_261102_070901.pdf",
         "Aboalez_2025_Posthumanist_social_approach_to_energy_transition.pdf"),
        ("Molepo, P M, Aboalez, K, & Mathaba, T N D (2_261102_070808.pdf",
         "Molepo_Aboalez_Mathaba_2025_Barriers_to_South_Africas_energy_transition.pdf"),
    ]
    dest = os.path.join(NEW, "06_References", "Article_PDFs")
    for s, d in items:
        sp = os.path.join(src, s)
        if os.path.exists(sp):
            shutil.copy2(sp, os.path.join(dest, d))
            print("COPIED", d)


def copy_matrix():
    s = os.path.join(ROOT, "research_work", "literature_matrix.xlsx")
    d = os.path.join(NEW, "08_Quality_Assurance", "Literature_Matrix.xlsx")
    if os.path.exists(s):
        shutil.copy2(s, d)
        print("COPIED Literature_Matrix.xlsx")


# --------------------------------------------------------------------------
# Handover summary
# --------------------------------------------------------------------------
def build_handover():
    doc = bd.base_document()
    bd.cover(doc, "Client Handover", "Research Deliverables — Handover Summary",
             "Contents, status and outstanding items",
             [("Project", "PhD in Engineering Management"),
              ("Research topic", "Systemic Living Lab for Just Energy Transition"),
              ("Institution", "University of Johannesburg"),
              ("Date", bd.DATE),
              ("Package status", "Partial delivery — see Section 4")])
    bd.add_footer(doc, "Handover Summary")
    bd.add_toc(doc)
    doc.add_page_break()

    bd.h1(doc, "1. Purpose of this package")
    bd.para(doc, "This package collects the research deliverables that are finished, checked and safe to "
                 "circulate. It is organised so that a reader can tell immediately which documents are "
                 "editable, which are read-only copies, and which parts of the project remain outstanding.")
    bd.para(doc, "The package does not contain the chapter drafts. Those chapters have not been written. "
                 "Nothing in this package is a draft, and no placeholder document has been created to stand "
                 "in for work that does not exist.")

    bd.h1(doc, "2. Included documents")
    bd.table(doc, "Table 1: Delivered files",
             ["Folder", "Document", "Format", "Description"],
             [["06", "Verified Reference List", "Word, PDF", "Nine records confirmed against publisher metadata"],
              ["06", "EndNote References", "RIS", "The same nine records, ready for EndNote import"],
              ["06", "Article PDFs", "PDF", "Two lawfully held source articles"],
              ["07", "List of Acronyms", "Word, PDF", "Rebuilt working list with conflicts recorded"],
              ["08", "Reference Verification Report", "Word, PDF", "Bibliographic audit and verification protocol"],
              ["08", "Academic Revision Report", "Word, PDF", "Chapter-by-chapter revision plan and word budget"],
              ["08", "Research Consistency Report", "Word, PDF", "Audit of theory, terminology, questions and claims"],
              ["08", "Literature Matrix", "Excel", "49 sources mapped to themes, questions and evidence strength"],
              ["09", "Handover Summary", "Word, PDF", "This document"],
              ["09", "Handover README", "Text", "Plain-language guide to the package"]],
             [1.6, 4.8, 2.2, 7.3], fsize=8.5)

    bd.h1(doc, "3. Editable and read-only versions")
    bd.para(doc, "Every substantial document is supplied twice. The Word file is the editable master and should "
                 "be the working copy. The PDF is a read-only rendering of the same content for circulation. "
                 "Both are produced from a single source, so their content is identical at the time of delivery; "
                 "a PDF will not reflect later edits to the Word file until it is regenerated.")
    bd.para(doc, "The Excel workbook has no PDF counterpart. That is deliberate: it is an analytical instrument "
                 "for sorting and filtering, not a document for reading.")

    bd.h1(doc, "4. Work not included")
    bd.para(doc, "The following are absent by design. Each entry states why, so that the absence is a matter of "
                 "record rather than an oversight.")
    bd.table(doc, "Table 2: Outstanding items",
             ["Item", "Reason"],
             [["Complete editable research document", "Cannot be assembled. Chapters 1 to 7 do not all exist, so any combined file would be a misleading artefact."],
              ["Chapter 3 — Theoretical Framework (Word and PDF)", "Not written. Requires the four primary theory texts and a written supervisor position on permissible drafting support."],
              ["Conceptual framework (Word, PNG, SVG)", "Depends on Chapter 3. The existing diagrams are low-resolution embedded bitmaps with no editable source and must be redrawn."],
              ["Abstract (Word and PDF)", "The source document's abstract page is empty; no abstract can be written before the thesis argument is settled."],
              ["Individual chapter files", "No chapter has been revised or approved for inclusion."],
              ["List of Figures and List of Tables", "Figure and table numbering in the source repeats and runs out of sequence."],
              ["Complete reference list", "Approximately 250 to 300 citations remain unverified; two heavily used sources cannot be located."],
              ["EndNote library file", "Cannot be built or tested outside EndNote. The RIS file is supplied instead."]],
             [5.6, 10.3], fsize=9)

    bd.h1(doc, "5. Using the reference files")
    bd.h2(doc, "5.1 Importing into EndNote")
    bd.para(doc, "Open EndNote and choose File, then Import, then File. Select the RIS file from the reference "
                 "folder. Set the import option to Reference Manager (RIS) and the text translation to Unicode "
                 "(UTF-8). Import into a new group so the verified set stays distinguishable from later additions.")
    bd.para(doc, "The file contains nine records. It is not a complete bibliography for the thesis, and it "
                 "should not be described as one.")
    bd.h2(doc, "5.2 Article PDFs")
    bd.para(doc, "Two source articles are included: the supervisor's 2025 paper on a posthumanist approach to "
                 "energy transition, and the 2025 analysis of barriers to South Africa's energy transition. Both "
                 "were supplied with the project and are held lawfully. They are the only full texts available; "
                 "the remaining seven verified sources require library access.")

    bd.h1(doc, "6. Approvals required")
    bd.table(doc, "Table 3: Decisions needed before further work",
             ["#", "Item", "Who", "Why it blocks progress"],
             [["1", "Written position on permissible drafting support", "Supervisor", "The faculty guideline prohibits text generation; no chapter can be drafted until this is clarified"],
              ["2", "Referencing style", "Faculty", "The reference list cannot be finalised without it"],
              ["3", "Thesis or proposal template", "Faculty", "Front matter and heading hierarchy must follow it"],
              ["4", "Whether a proposal requires a results chapter", "Supervisor", "Determines the final chapter structure"],
              ["5", "Access to the four primary theory texts", "Client, library", "Chapter 3 cannot be written from secondary sources"]],
             [1.0, 6.0, 3.2, 5.7], fsize=9)

    bd.h1(doc, "7. Statement of integrity")
    bd.para(doc, "No research finding, reference, DOI, page number, quotation or institutional requirement in this "
                 "package has been invented. Where a fact could not be established it is recorded as unresolved. "
                 "Where a document does not exist it is not included.")
    bd.para(doc, "No originality or similarity report is included, because no such check has been run. No claim "
                 "is made about similarity scores.")
    bd.save(doc, r"09_Final_Handover\Handover_Summary.docx",
            "Handover Summary", "Package contents and outstanding work")


README = """JOHN PHD RESEARCH DELIVERABLES
===============================================================================
Project     : PhD in Engineering Management
Topic       : Developing an Innovative Systemic Living Lab for Just Energy
              Transition in a South African Informal Settlement
Institution : University of Johannesburg
Date        : 12 November 2026
Status      : PARTIAL DELIVERY - read Section 4
===============================================================================

Read this file first. It explains what is in the package, what is not, and why.


--------------------------------------------------------------------------
1. CONTENTS
--------------------------------------------------------------------------

  01_Research_Chapters
      Empty. No chapter has been revised and approved for inclusion.

  02_Chapter_3_Theoretical_Framework
      Empty. Chapter 3 has not been written.

  03_Conceptual_Framework
      Empty. Depends on Chapter 3, which does not exist.

  04_Abstract
      Empty. The source document's abstract page is blank.

  05_Complete_Research_Document
      Empty. A combined document is only possible once all chapters exist.

  06_References
      Reference_List.docx          editable master
      Reference_List.pdf           read-only copy
      EndNote_References.ris       nine verified records, EndNote importable
      Article_PDFs\\                two lawfully held source articles

  07_Front_Matter
      List_of_Acronyms.docx        editable master
      List_of_Acronyms.pdf         read-only copy
      Figures\\                     empty - see Section 4
      Tables\\                      empty - see Section 4

  08_Quality_Assurance
      Reference_Verification_Report.docx / .pdf
      Academic_Revision_Report.docx / .pdf
      Research_Consistency_Report.docx / .pdf
      Literature_Matrix.xlsx

  09_Handover
      Handover_Summary.docx / .pdf
      HANDOVER_README.txt          this file


--------------------------------------------------------------------------
2. EDITABLE vs READ-ONLY
--------------------------------------------------------------------------

Editable masters (.docx):
  06_References\\Reference_List.docx
  07_Front_Matter\\List_of_Acronyms.docx
  08_Quality_Assurance\\Reference_Verification_Report.docx
  08_Quality_Assurance\\Academic_Revision_Report.docx
  08_Quality_Assurance\\Research_Consistency_Report.docx
  09_Handover\\Handover_Summary.docx

Read-only copies (.pdf): the same six documents.

The PDFs are generated from the Word files, so content is identical at delivery.
Editing a Word file will not change its PDF until the PDF is regenerated.

Formatting throughout: A4 paper, Times New Roman 12 pt, 1.5 line spacing,
justified body text, consistent heading hierarchy, automatic page numbering,
automatic table of contents. Press Ctrl+A then F9 in Word to refresh fields.


--------------------------------------------------------------------------
3. ENDNOTE IMPORT
--------------------------------------------------------------------------

  1. Open EndNote.
  2. File > Import > File...
  3. Select  06_References\\EndNote_References.ris
  4. Import Option  : Reference Manager (RIS)
  5. Text Translation: Unicode (UTF-8)
  6. Import into a NEW group, e.g. "Verified_Set_2026"

The RIS file holds NINE records, each checked field by field against publisher
metadata. It is NOT a complete library. The research document contains roughly
250-300 citations and has no reference list of its own.

No .enl library file is supplied. A library file cannot be created or tested
outside EndNote, and supplying one built from nine records would misrepresent
it as a finished thesis library. The RIS file is the correct interchange
format and imports cleanly.


--------------------------------------------------------------------------
4. OUTSTANDING WORK
--------------------------------------------------------------------------

Absent by design, not by oversight:

  Complete research document
    Requires Chapters 1-7. They do not all exist.

  Chapter 3 - Theoretical Framework (Word and PDF)
    Requires the four primary theory texts (Braidotti 2013 and 2019;
    Haraway 1988 and 2016; Barad 2007; Hayles 1999) and a written supervisor
    position on permissible drafting support. Writing it without those texts
    would mean inventing interpretations of named theorists.

  Conceptual framework (Word, PNG, SVG)
    Depends on Chapter 3. The existing diagrams are low-resolution bitmaps
    embedded in a PDF with no editable source. They must be REDRAWN, not
    exported. Exporting them would produce poor files and would present
    unfinished work as final.

  Abstract (Word and PDF)
    The source abstract page is empty. An abstract describes findings and
    argument; neither is settled yet.

  Individual chapter files
    No chapter has been revised or approved.

  List of Figures / List of Tables
    Figure and table numbering repeats and runs out of sequence (three
    different figures are all numbered 2.1). The lists cannot be generated
    until the objects are renumbered.

  Complete reference list
    Roughly 250-300 citations remain unverified. Two heavily used sources
    (Okoye 2020; Kovacic et al. 2019) cannot be located at all.

  Article PDFs beyond the two included
    Seven of the nine verified sources require library access. No source has
    been taken from an unauthorised repository, and none will be.


--------------------------------------------------------------------------
5. APPROVALS REQUIRED
--------------------------------------------------------------------------

  1. SUPERVISOR - written position on permissible drafting support.
     The faculty guideline prohibits text generation. Nothing further can be
     drafted until this is settled in writing.

  2. FACULTY - referencing style (Harvard or APA 7).

  3. FACULTY - university thesis or proposal template.

  4. SUPERVISOR - whether a proposal requires a results chapter.

  5. CLIENT and LIBRARY - access to the four primary theory texts.


--------------------------------------------------------------------------
6. INTEGRITY STATEMENT
--------------------------------------------------------------------------

No research finding, reference, DOI, page number, quotation or institutional
requirement in this package has been invented. Where a fact could not be
established, it is recorded as unresolved. Where a document does not exist,
it is not included.

No Turnitin or similarity report is included, because none has been run. No
claim is made about similarity percentages.

This package is an accurate statement of what has been completed. It is not a
claim about the whole project.

===============================================================================
End of file.
===============================================================================
"""


def build_readme():
    p = os.path.join(NEW, "09_Handover", "HANDOVER_README.txt")
    os.makedirs(os.path.dirname(p), exist_ok=True)
    with open(p, "w", encoding="utf-8", newline="\r\n") as f:
        f.write(README)
    print("WROTE HANDOVER_README.txt")


if __name__ == "__main__":
    make_dirs()
    bd.build_reference_verification()
    bd.build_academic_revision()
    bd.build_consistency()
    bd.build_reference_list()
    bd.build_acronyms()
    build_handover()
    bd.build_ris()
    bd.build_status()
    copy_article_pdfs()
    copy_matrix()
    build_readme()
    print("\nDONE")
