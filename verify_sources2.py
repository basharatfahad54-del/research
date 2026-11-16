# -*- coding: utf-8 -*-
import json, time, urllib.parse, urllib.request

QUERIES = [
    "Sandberg Alvesson Ways of constructing research questions gap-spotting or problematization",
    "Geels Technological transitions as evolutionary reconfiguration processes a multi-level perspective and a case-study",
    "Geels The multi-level perspective on sustainability transitions responses to seven criticisms",
    "Jenkins McCauley Heffron Stephan Rehner Energy justice a conceptual review",
    "Newell Mulvaney The political economy of the just transition",
    "Winkler Towards a theory of just transition energy justice and the future of work",
    "Davis Perceived usefulness perceived ease of use and user acceptance of information technology",
    "Sanders Stappers Co-creation and the new landscapes of design",
    "Stephenson Barton-Johnson Energy cultures a framework for understanding energy behaviours",
    "Leminen Westerlund Nystrom Living Labs as open-innovation networks",
    "von Wirth Fuenfschilling Frantzeskaki Coenen Impacts of urban living labs on sustainability transitions mechanisms and strategies",
    "Sovacool What are we doing here Analyzing fifteen years of energy scholarship and proposing a social science research agenda",
    "Petrulaityte Ceschin Musango Mwiti Anditi Njoroge Supporting the development of gendered energy innovations for informal urban settlements GENS codesign toolkit",
    "Anditi Musango Smit Ceschin Addressing gender dimensions in energy innovations a gender analysis framework for informal urban settlements in Africa",
    "Oosthuizen de Kock Musango Exploring factors that influence the mainstreaming of gendered energy interventions in poor urban environments",
    "Musango Assessing gender and energy in urban household energy transitions in South Africa quantitative storytelling Groenheuwel",
    "Ricci Co-creating justice in housing energy transitions through energy living labs",
    "Coole Frost New Materialisms Ontology Agency and Politics",
    "Bennett Vibrant Matter A Political Ecology of Things",
    "Latour Reassembling the Social An Introduction to Actor-Network-Theory",
    "Winner Do artifacts have politics",
    "Shove Comfort Cleanliness and Convenience The Social Organization of Normality",
    "Maxwell Qualitative Research Design An Interactive Approach",
    "Ravitch Riggan Reason Rigor How Conceptual Frameworks Guide Research",
    "Cilliers Complexity and Postmodernism Understanding Complex Systems",
    "Checkland Systems Thinking Systems Practice",
    "Midgley Systemic Intervention Philosophy Methodology and Practice",
    "Puig de la Bellacasa Matters of Care Speculative Ethics in More than Human Worlds",
    "Haraway Simians Cyborgs and Women The Reinvention of Nature",
]


def q(s):
    url = ("https://api.crossref.org/works?rows=2&query.bibliographic="
           + urllib.parse.quote(s) + "&select=DOI,title,author,issued,container-title,volume,issue,page,type,publisher")
    req = urllib.request.Request(url, headers={"User-Agent": "PhDRefCheck/1.0 (mailto:research@example.org)"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r)["message"]["items"]


out = []
for s in QUERIES:
    for attempt in range(4):
        try:
            items = q(s)
            break
        except Exception as e:
            items = None
            err = str(e)
            time.sleep(5 + attempt * 5)
    if items is None:
        out.append({"query": s, "error": err})
        continue
    recs = []
    for it in items:
        au = it.get("author", [])
        recs.append({
            "doi": it.get("DOI"),
            "title": (it.get("title") or [""])[0],
            "authors": ", ".join((a.get("family", "") + " " + a.get("given", "")).strip() for a in au[:8]),
            "year": (it.get("issued", {}).get("date-parts") or [[None]])[0][0],
            "journal": (it.get("container-title") or [""])[0],
            "volume": it.get("volume"), "issue": it.get("issue"), "page": it.get("page"),
            "type": it.get("type"), "publisher": it.get("publisher"),
        })
    out.append({"query": s, "results": recs})
    time.sleep(3)

with open(r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\crossref_verification2.json", "w", encoding="utf-8") as f:
    json.dump(out, f, indent=2, ensure_ascii=False)
with open(r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\crossref_verification2.txt", "w", encoding="utf-8") as f:
    for rec in out:
        f.write("QUERY: " + rec["query"] + "\n")
        if "error" in rec:
            f.write("  ERROR " + rec["error"] + "\n\n"); continue
        for r in rec["results"]:
            f.write("  %s | %s | %s | %s %s %s\n" % (r["year"], r["authors"], r["title"], r["journal"], r["volume"], r["page"]))
            f.write("    DOI %s | %s | %s\n" % (r["doi"], r["type"], r["publisher"]))
        f.write("\n")
print("DONE", len(out))
