# -*- coding: utf-8 -*-
import re
import docx

d = docx.Document(r"research-main\Edited_Proposal_Output\Research_Proposal_Final.docx")
p = list(d.paragraphs)[22]
xml = p._p.xml
print("instrText occurrences:", xml.count("instrText"))
print("fldChar elements:", len(re.findall(r'<w:fldChar [^>]*fldCharType="(begin|separate|end)"', xml)))
m = re.search(r"instrText[^>]*>([^<]*)<", xml)
print("instruction:", repr(m.group(1)) if m else None)
print("cached text ok:", "Microsoft Word field" in p.text)
