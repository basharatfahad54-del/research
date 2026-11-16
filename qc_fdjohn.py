# -*- coding: utf-8 -*-
"""QC of the fd john deliverables pack."""
import os, re, glob
import pymupdf
from docx import Document

DELIV = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'fd john', 'deliverables')

print('=== DOCX <-> PDF pairing and content check ===')
issues = []
for docx in sorted(glob.glob(os.path.join(DELIV, '**', '*.docx'), recursive=True)):
    pdf = docx[:-5] + '.pdf'
    name = os.path.relpath(docx, DELIV)
    if not os.path.exists(pdf):
        print('MISSING PDF:', name)
        issues.append(name)
        continue
    doc = Document(docx)
    dwords = sum(len(p.text.split()) for p in doc.paragraphs)
    for t in doc.tables:
        for row in t.rows:
            for c in row.cells:
                dwords += len(c.text.split())
    d = pymupdf.open(pdf)
    txt = ''.join(p.get_text() for p in d)
    pwords = len(txt.split())
    imgs = sum(len(p.get_images()) for p in d)
    bad = [b for b in ('Lorem ipsum', 'XXX', '[insert', 'TODO') if b.lower() in txt.lower()]
    flag = ''
    if abs(dwords - pwords) > 0.12 * max(dwords, 1):
        flag += ' WORD-DIFF(%.0f%%)' % (abs(dwords - pwords) / max(dwords, 1) * 100)
    if bad:
        flag += ' BAD=%s' % bad
    print('%-52s docx=%5dw pdf=%2dp %5dw imgs=%d%s' % (name, dwords, d.page_count, pwords, imgs, flag))
    d.close()

print()
print('=== Reference file record counts ===')
ris = open(os.path.join(DELIV, 'D08_References', 'EndNote_References.ris'), encoding='utf-8-sig').read()
bib = open(os.path.join(DELIV, 'D08_References', 'references.bib'), encoding='utf-8').read()
xml = open(os.path.join(DELIV, 'D08_References', 'references_endnote.xml'), encoding='utf-8').read()
print('RIS records   :', len([r for r in ris.split('ER  -') if r.strip()]))
print('BibTeX entries:', len(re.findall(r'@\w+\{', bib)))
print('EndNote XML   :', len(re.findall(r'<record>', xml)))
print('Article PDFs  :', len(glob.glob(os.path.join(DELIV, 'D08_References', 'Article_PDFs', '*.pdf'))))

print()
print('=== D07 figures folder ===')
for f in sorted(glob.glob(os.path.join(DELIV, 'D07_Required_Diagrams', '*'))):
    print(' %-46s %6d KB' % (os.path.basename(f), os.path.getsize(f) // 1024))

print()
print('=== PDF outline spot-check (D01) ===')
d = pymupdf.open(os.path.join(DELIV, 'D01_Client_Requirements_Specification.pdf'))
for lvl, title, page in d.get_toc()[:8]:
    print('  L%d p%-3d %s' % (lvl, page, title[:60]))
d.close()
print()
print('QC COMPLETE. issues:', issues if issues else 'none')
