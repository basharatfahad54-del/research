import os, glob, hashlib, re
import pymupdf
import openpyxl

BASE = r'C:\Users\Muhammad Fahad\Desktop\Research paper\JOHN_PHD_RESEARCH_DELIVERABLES'
ORIG = r'C:\Users\Muhammad Fahad\Desktop\Research paper'

bad = ['TODO', 'TBD', 'Lorem ipsum', 'XXX', '[insert', 'placeholder']
print('=== PDF QC ===')
for pdf in sorted(glob.glob(os.path.join(BASE, '**', '*.pdf'), recursive=True)):
    d = pymupdf.open(pdf)
    txt = ''.join(p.get_text() for p in d)
    hits = [b for b in bad if b.lower() in txt.lower()]
    print(f'{os.path.relpath(pdf, BASE):66s} p={d.page_count:2d} words={len(txt.split()):5d} issues={hits}')
    d.close()

print()
print('=== ARTICLE PDF INTEGRITY ===')
pairs = [
    ('Aboalez, K (2025)_261102_070901.pdf',
     r'06_References\Article_PDFs\Aboalez_2025_Posthumanist_social_approach_to_energy_transition.pdf'),
    ('Molepo, P M, Aboalez, K, & Mathaba, T N D (2_261102_070808.pdf',
     r'06_References\Article_PDFs\Molepo_Aboalez_Mathaba_2025_Barriers_to_South_Africas_energy_transition.pdf'),
]
for src, dst in pairs:
    h1 = hashlib.sha256(open(os.path.join(ORIG, src), 'rb').read()).hexdigest()
    h2 = hashlib.sha256(open(os.path.join(BASE, dst), 'rb').read()).hexdigest()
    d = pymupdf.open(os.path.join(BASE, dst))
    print(f'{"IDENTICAL" if h1 == h2 else "DIFFERENT"}  {os.path.basename(dst)[:58]:58s} pages={d.page_count}')
    d.close()

print()
print('=== RIS ===')
t = open(os.path.join(BASE, r'06_References\EndNote_References.ris'), encoding='utf-8').read()
recs = [r for r in t.split('ER  - ') if r.strip()]
dois = re.findall(r'^DO  - (.+)$', t, re.M)
print(f'records={len(recs)} dois={len(dois)} unique={len(set(dois))}')

print()
print('=== XLSX ===')
wb = openpyxl.load_workbook(os.path.join(BASE, '08_Quality_Assurance', 'Literature_Matrix.xlsx'))
print('sheets:', wb.sheetnames, '| source rows:', wb['Sources'].max_row)

print()
print('=== PLACEHOLDER / EMPTY DOCX CHECK ===')
from docx import Document
for dx in sorted(glob.glob(os.path.join(BASE, '**', '*.docx'), recursive=True)):
    doc = Document(dx)
    n = len(' '.join(p.text for p in doc.paragraphs).split())
    print(f'{os.path.basename(dx):38s} paragraphs={len(doc.paragraphs):3d} tables={len(doc.tables):2d} words={n}')
