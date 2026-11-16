# -*- coding: utf-8 -*-
"""Queries Crossref for each candidate source and writes a verification table."""
import json, time, urllib.parse, urllib.request

QUERIES = [
    "Braidotti The Posthuman 2013",
    "Braidotti Posthuman Knowledge 2019",
    "Haraway Situated Knowledges The Science Question in Feminism and the Privilege of Partial Perspective",
    "Haraway Staying with the Trouble Making Kin in the Chthulucene",
    "Haraway Anthropocene Capitalocene Plantationocene Chthulucene Making Kin",
    "Barad Posthumanist Performativity Toward an Understanding of How Matter Comes to Matter",
    "Barad Meeting the Universe Halfway Quantum Physics and the Entanglement of Matter and Meaning",
    "Hayles How We Became Posthuman Virtual Bodies in Cybernetics Literature and Informatics",
    "Wolfe What Is Posthumanism",
    "Forlano Posthumanism and Design",
    "Ulmer Posthumanism as research methodology education in the Anthropocene",
    "Grant Osanloo Understanding selecting and integrating a theoretical framework in dissertation research",
    "Sandberg Alvesson Ways of constructing research questions gap-spotting or problematization",
    "Geels Technological transitions as evolutionary reconfiguration processes a multi-level perspective and a case-study",
    "Geels The multi-level perspective on sustainability transitions responses to seven criticisms",
    "Sovacool Dworkin Energy justice conceptual insights and practical applications",
    "Jenkins McCauley Heffron Stephan Rehner Energy justice a conceptual review",
    "Walker Day Fuel poverty as injustice integrating distribution recognition and procedure",
    "Heffron McCauley The concept of energy justice across the disciplines",
    "Newell Mulvaney The political economy of the just transition",
    "Winkler Towards a theory of just transition energy justice and the future of work",
    "Arnstein A ladder of citizen participation",
    "Davis Perceived usefulness perceived ease of use and user acceptance of information technology",
    "Venkatesh Morris Davis Davis User acceptance of information technology toward a unified view",
    "Wustenhagen Wolsink Burer Social acceptance of renewable energy innovation an introduction to the concept",
    "Sanders Stappers Co-creation and the new landscapes of design",
    "Stephenson Barton-Johnson Energy cultures a framework for understanding energy behaviours",
    "Leminen Westerlund Nystrom Living Labs as open-innovation networks",
    "Hossain Leminen Westerlund A systematic review of living lab literature",
    "DellEra Landoni Living Lab a methodology between user-centred design and participatory design",
    "von Wirth Fuenfschilling Frantzeskaki Coenen Impacts of urban living labs on sustainability transitions mechanisms and strategies",
    "Sovacool What are we doing here Analyzing fifteen years of energy scholarship and proposing a social science research agenda",
    "Musango Smit Ceschin Ambole Batinge Anditi Petrulaityte Mukama Mainstreaming gender to achieve security of energy services in poor urban environments",
    "Mukama Musango Smit Ceschin Petrulaityte Development of living labs to support gendered energy technology innovation in poor urban environments",
    "Petrulaityte Ceschin Musango Mwiti Anditi Njoroge Supporting the development of gendered energy innovations for informal urban settlements GENS codesign toolkit",
    "Anditi Musango Smit Ceschin Addressing gender dimensions in energy innovations a gender analysis framework for informal urban settlements in Africa",
    "Batinge Musango Ceschin Ambole Smit Petrulaityte Modelling gendered innovation for the security of energy services in poor urban environments",
    "Oosthuizen de Kock Musango Exploring factors that influence the mainstreaming of gendered energy interventions in poor urban environments",
    "Musango Assessing gender and energy in urban household energy transitions in South Africa quantitative storytelling Groenheuwel",
    "Molepo Aboalez Mathaba Analysis of barriers to South Africa's energy transition perspectives from industry experts",
    "Aboalez A posthumanist social approach to energy transition",
    "Systematic Review Application of the Living Lab Concept and Role of Stakeholders in the Energy Sector",
    "Ricci Co-creating justice in housing energy transitions through energy living labs",
]


def q(s):
    url = ("https://api.crossref.org/works?rows=3&query.bibliographic="
           + urllib.parse.quote(s) + "&select=DOI,title,author,issued,container-title,volume,issue,page,type,publisher")
    req = urllib.request.Request(url, headers={"User-Agent": "PhDRefCheck/1.0 (mailto:research@example.org)"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r)["message"]["items"]


out = []
for s in QUERIES:
    try:
        items = q(s)
    except Exception as e:
        out.append({"query": s, "error": str(e)})
        time.sleep(1)
        continue
    recs = []
    for it in items[:2]:
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
    time.sleep(0.4)

with open(r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\crossref_verification.json", "w", encoding="utf-8") as f:
    json.dump(out, f, indent=2, ensure_ascii=False)

with open(r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\crossref_verification.txt", "w", encoding="utf-8") as f:
    for rec in out:
        f.write("QUERY: " + rec["query"] + "\n")
        if "error" in rec:
            f.write("  ERROR " + rec["error"] + "\n\n")
            continue
        for r in rec["results"]:
            f.write("  %s | %s | %s | %s | %s %s %s\n" % (
                r["year"], r["authors"], r["title"], r["journal"], r["volume"], r["issue"], r["page"]))
            f.write("    DOI %s | %s | %s\n" % (r["doi"], r["type"], r["publisher"]))
        f.write("\n")
print("DONE", len(out))
