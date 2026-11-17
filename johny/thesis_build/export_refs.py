# -*- coding: utf-8 -*-
"""Regenerates EndNote-compatible exports (RIS / BibTeX / EndNote XML) from the
verified 73-entry reference list in merged_12.md. UTF-8, full coverage, no
invented sources. Replaces the earlier 34-record mojibake exports."""
import io, os, re, sys

ROOT = r"C:\Users\Muhammad Fahad\Desktop\research-main\research-main"
SRC = os.path.join(ROOT, "johny", "thesis_build", "merged_12.md")
OUT = os.path.join(ROOT, "johny", "Final_Thesis")

ORG_AUTHORS = {"Department of Mineral Resources and Energy",
               "Statistics South Africa",
               "Presidential Climate Commission"}

def parse_refs(text):
    entries = []
    for ln in text.splitlines():
        if not ln.startswith("[REF] "):
            continue
        body = ln[len("[REF] "):].strip()
        m = re.match(r"^(.*?)\s*\((\d{4})\)\s*(.*)$", body)
        if not m:
            print("PARSE FAIL:", body[:80]); continue
        authors_str, year, rest = m.group(1).strip(), m.group(2), m.group(3).strip()
        e = {"authors": [], "year": year, "doi": None, "volume": None, "issue": None,
             "pages": None, "container": None, "publisher": None, "place": None,
             "rtype": None, "title": None}
        if rest.startswith("'"):
            e["rtype"] = "JOUR"
            mm = re.match(r"^'(.+?)',\s*", rest)
            e["title"] = mm.group(1)
            rest = rest[mm.end():]
            cm = re.match(r"^\*(.+?)\*\s*,?\s*", rest)
            if cm:
                e["container"] = cm.group(1)
                rest = rest[cm.end():]
            else:
                cm2 = re.match(r"^([^*]+?),\s*", rest)
                e["container"] = cm2.group(1).strip()
                rest = rest[cm2.end():]
            vm = re.match(r"^(\d+)(?:\(([^)]+)\))?\s*,?\s*", rest)
            if vm:
                e["volume"] = vm.group(1)
                e["issue"] = vm.group(2)
                rest = rest[vm.end():]
            pm = re.search(r"pp\.\s*([\d\u2013\u2014\-]+)", rest)
            if pm:
                e["pages"] = pm.group(1)
            am = re.search(r",\s*([0-9]+[A-Za-z0-9]*?)\.\s*(doi|$)", rest)
            if am and not e["pages"] and am.group(1) not in ("",):
                e["pages"] = am.group(1)  # article number (e.g. 102476, e0232076)
            dm = re.search(r"doi:\s*(\S+?)\.?\s*$", rest)
            if dm:
                e["doi"] = dm.group(1).rstrip(".")
        else:
            im = re.match(r"^\*(.+?)\*\s*\.?\s*(.*)$", rest)
            if not im or not im.group(2):
                im = re.match(r"^\*(.+?)\*\s*,\s*(.*)$", rest)
                if im:
                    rest = "*%s* %s" % (im.group(1), im.group(2))
                    im = re.match(r"^\*(.+?)\*\s*\.?\s*(.*)$", rest)
            e["title"] = im.group(1)
            tail = im.group(2)
            dm = re.search(r"doi:\s*(\S+?)\.?\s*$", tail)
            if dm:
                e["doi"] = dm.group(1).rstrip(".")
                tail = tail[:dm.start()].strip().rstrip(".")
            pl = re.match(r"^([^.]+):\s*([^.]+?)\.?\s*$", tail)
            if pl:
                e["place"] = pl.group(1).strip()
                e["publisher"] = pl.group(2).strip()
            if authors_str in ORG_AUTHORS:
                e["rtype"] = "RPRT"
            else:
                e["rtype"] = "BOOK"
        # split authors
        au = authors_str.replace(" and ", ", ")
        units, cur = [], []
        tokens = au.split(", ")
        i = 0
        while i < len(tokens):
            tok = tokens[i]
            if i + 1 < len(tokens) and re.match(r"^[A-Z]\.\s*[A-Z]?\.?$|^[A-Z]\.$", tokens[i + 1] or ""):
                units.append(tok + ", " + tokens[i + 1]); i += 2
            else:
                units.append(tok); i += 1
        e["authors"] = units if len(units) > 0 else [authors_str]
        entries.append(e)
    return entries

def ris_key(e):
    first = e["authors"][0].split(",")[0].replace(" ", "").replace("'", "")
    word = re.sub(r"[^A-Za-z]", "", e["title"].split()[0]) if e["title"] else "Item"
    return first + e["year"] + word

def write_ris(entries, path):
    out = []
    for e in entries:
        ty = {"JOUR": "JOUR", "BOOK": "BOOK", "RPRT": "RPRT"}[e["rtype"]]
        out.append("TY  - " + ty)
        for a in e["authors"]:
            out.append("AU  - " + a)
        out.append("TI  - " + e["title"])
        if e["container"]:
            out.append("T2  - " + e["container"])
            out.append("JO  - " + e["container"])
        if e["publisher"]:
            out.append("PB  - " + e["publisher"])
        if e["place"]:
            out.append("CY  - " + e["place"])
        out.append("PY  - " + e["year"])
        if e["volume"]:
            out.append("VL  - " + e["volume"])
        if e["issue"]:
            out.append("IS  - " + e["issue"])
        if e["pages"]:
            out.append("SP  - " + e["pages"])
        if e["doi"]:
            out.append("DO  - " + e["doi"])
            out.append("UR  - https://doi.org/" + e["doi"])
        out.append("ER  - ")
        out.append("")
    with io.open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(out))

FOLD = {"å": "a", "ä": "a", "ö": "o", "ü": "u", "é": "e", "è": "e", "ê": "e",
        "á": "a", "à": "a", "í": "i", "ó": "o", "ú": "u", "ñ": "n", "ç": "c",
        "Å": "A", "Ä": "A", "Ö": "O", "Ü": "U", "É": "E", "È": "E"}

def fold_ascii(s):
    for k, v in FOLD.items():
        s = s.replace(k, v)
    return s

def bibtex_escape(s):
    return s.replace("&", "\\&").replace("%", "\\%")

def write_bib(entries, path):
    out = []
    used = set()
    for e in entries:
        k = fold_ascii(ris_key(e)); base, n = k, 2
        while k in used:
            k = base + chr(96 + n); n += 1
        used.add(k)
        auth = " and ".join(e["authors"])
        if e["rtype"] == "JOUR":
            fields = ["author={%s}" % bibtex_escape(auth), "title={%s}" % bibtex_escape(e["title"]),
                      "journal={%s}" % bibtex_escape(e["container"] or ""), "year={%s}" % e["year"]]
            if e["volume"]: fields.append("volume={%s}" % e["volume"])
            if e["issue"]: fields.append("number={%s}" % e["issue"])
            if e["pages"]: fields.append("pages={%s}" % e["pages"].replace("\u2013", "--"))
            if e["doi"]: fields.append("doi={%s}" % e["doi"])
            out.append("@article{%s,\n  %s\n}" % (k, ",\n  ".join(fields)))
        elif e["rtype"] == "BOOK":
            fields = ["author={%s}" % bibtex_escape(auth), "title={%s}" % bibtex_escape(bibtex_escape(e["title"]))]
            if e["publisher"]: fields.append("publisher={%s}" % bibtex_escape(e["publisher"]))
            if e["place"]: fields.append("address={%s}" % bibtex_escape(e["place"]))
            fields.append("year={%s}" % e["year"])
            if e["doi"]: fields.append("doi={%s}" % e["doi"])
            out.append("@book{%s,\n  %s\n}" % (k, ",\n  ".join(fields)))
        else:
            fields = ["author={%s}" % bibtex_escape(auth), "title={%s}" % bibtex_escape(e["title"])]
            if e["publisher"]: fields.append("institution={%s}" % bibtex_escape(e["publisher"]))
            fields.append("year={%s}" % e["year"])
            if e["doi"]: fields.append("doi={%s}" % e["doi"])
            out.append("@techreport{%s,\n  %s\n}" % (k, ",\n  ".join(fields)))
        out.append("")
    with io.open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(out))

def xml_escape(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")

def write_xml(entries, path):
    out = ["<?xml version=\"1.0\" encoding=\"UTF-8\"?>", "<xml>", "<records>"]
    for i, e in enumerate(entries, 1):
        rt = {"JOUR": ("Journal Article", "17"), "BOOK": ("Book", "6"),
              "RPRT": ("Report", "27")}[e["rtype"]]
        out.append("<record>")
        out.append("<ref-type name=\"%s\">%s</ref-type>" % rt)
        out.append("<contributors><authors>")
        for a in e["authors"]:
            out.append("<author>%s</author>" % xml_escape(a))
        out.append("</authors></contributors>")
        out.append("<titles>")
        out.append("<title>%s</title>" % xml_escape(e["title"]))
        if e["container"]:
            out.append("<secondary-title>%s</secondary-title>" % xml_escape(e["container"]))
        out.append("</titles>")
        if e["container"]:
            out.append("<periodical><full-title>%s</full-title></periodical>" % xml_escape(e["container"]))
        if e["pages"]:
            out.append("<pages>%s</pages>" % xml_escape(e["pages"]))
        if e["volume"]:
            out.append("<volume>%s</volume>" % xml_escape(e["volume"]))
        if e["issue"]:
            out.append("<number>%s</number>" % xml_escape(e["issue"]))
        out.append("<dates><year>%s</year></dates>" % e["year"])
        if e["publisher"]:
            out.append("<pub-location>%s</pub-location>" % xml_escape(e["place"] or ""))
            out.append("<publisher>%s</publisher>" % xml_escape(e["publisher"]))
        if e["doi"]:
            out.append("<electronic-resource-num>%s</electronic-resource-num>" % xml_escape(e["doi"]))
        out.append("<label>L%02d</label>" % i)
        out.append("<rec-number>%d</rec-number>" % i)
        out.append("<remote-database-name>SETILL verified reference set</remote-database-name>")
        out.append("</record>")
    out += ["</records>", "</xml>"]
    with io.open(path, "w", encoding="utf-8", newline="\n") as f:
        f.write("\n".join(out))

if __name__ == "__main__":
    text = io.open(SRC, "r", encoding="utf-8").read()
    entries = parse_refs(text)
    print("PARSED", len(entries), "entries")
    write_ris(entries, os.path.join(OUT, "References.ris"))
    write_bib(entries, os.path.join(OUT, "References.bib"))
    write_xml(entries, os.path.join(OUT, "References_EndNote.xml"))
    print("WROTE References.ris / References.bib / References_EndNote.xml")
    for e in entries[:5]:
        print(" -", e["authors"][0], e["year"], "|", e["rtype"], "|", (e["title"] or "")[:60])
