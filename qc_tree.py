import os

os.chdir(r'C:\Users\Muhammad Fahad\Desktop\Research paper')
orig = {
    'Aboalez, K (2025)_261102_070901.pdf': 580278,
    'Description 2.txt': 660,
    'description.txt': 6226,
    'Molepo, P M, Aboalez, K, & Mathaba, T N D (2_261102_070808.pdf': 2253146,
    'PhD DETAILED CONCEPT NOTE 2026_CHECK_261111_113441.pdf': 4060980,
    'Structure of a Proposal in Engineering Management_261102_135929.pdf': 483983,
}
print('=== ORIGINALS UNCHANGED ===')
for f, sz in orig.items():
    cur = os.path.getsize(f) if os.path.exists(f) else None
    tag = 'OK     ' if cur == sz else 'CHANGED'
    print(f'{tag} {f[:60]:60s} {cur}')

print()
print('=== DELIVERABLES TREE ===')
tot = 0
size = 0
for root, dirs, files in os.walk('CLIENT_DELIVERABLES'):
    dirs.sort()
    rel = os.path.relpath(root, 'CLIENT_DELIVERABLES')
    depth = 0 if rel == '.' else rel.count(os.sep) + 1
    print('   ' * depth + (os.path.basename(root) + '/' if rel != '.' else 'CLIENT_DELIVERABLES/'))
    for f in sorted(files):
        p = os.path.join(root, f)
        tot += 1
        size += os.path.getsize(p)
        print('   ' * (depth + 1) + f + f'  ({os.path.getsize(p):,} bytes)')
print()
print('total files:', tot, '| total size:', f'{size:,}', 'bytes')
