import re, sys
p = r'C:\Users\Muhammad Fahad\Desktop\Research paper\CLIENT_DELIVERABLES\06_References_EndNote\EndNote_References.ris'
t = open(p, encoding='utf-8').read()
recs = [r.strip() for r in t.split('ER  - ') if r.strip()]
print('records:', len(recs))
ok = True
seen = set()
for i, r in enumerate(recs, 1):
    f = dict(re.findall(r'^(AU|TI|JO|VL|IS|SP|EP|PY|DO)  - (.+)$', r, re.M))
    doi = f.get('DO', 'MISSING')
    seen.add(doi)
    missing = [k for k in ('TI', 'PY', 'JO', 'DO') if k not in f]
    print(f"{i}. {f.get('PY','?')} | {f.get('AU','?')[:24]:24s} | {doi:32s} | {f.get('JO','?')[:40]:40s} missing={missing}")
    if missing:
        ok = False
print('\nall records complete:', ok)
print('unique DOIs:', len(seen), 'of', len(recs))

# check the Error! Bookmark occurrence is quoted prose, not a live field
import pymupdf
d = pymupdf.open(r'C:\Users\Muhammad Fahad\Desktop\Research paper\CLIENT_DELIVERABLES\08_Quality_Assurance\Research_Consistency_Report.pdf')
for pg in d:
    tx = pg.get_text()
    if 'Error! Bookmark' in tx:
        i = tx.find('Error! Bookmark')
        print('\nCONTEXT:', repr(tx[max(0, i-140):i+60]))
        print('page', pg.number + 1)
d.close()
