import re, os
base = r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work"
for name in ["ch2a_dump.txt","ch2b_dump.txt"]:
    p = os.path.join(base, name)
    lines = open(p, encoding="utf-8").read().splitlines()
    print("=====", name)
    for i,l in enumerate(lines):
        s=l.strip()
        if re.match(r"^2\.\d(\.\d)*\s+[A-Z]", s) or re.match(r"^2\.\d(\.\d)*$", s):
            nxt = lines[i+1].strip()[:80] if i+1<len(lines) else ""
            print(s, "||", nxt)
