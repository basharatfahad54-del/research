# -*- coding: utf-8 -*-
import json, time, urllib.parse, urllib.request

Q = [
 "Writing integrative literature reviews using the past and present to explore the future Torraco",
 "Barriers to powering past coal implications for a just energy transition in South Africa Mirzania",
 "Constructing mystery empirical matters in theory development Alvesson Karreman",
 "Living Lab an open and citizen-centric approach for innovation Bergvall-Kareborn",
 "Using thematic analysis in psychology Braun Clarke",
 "A simple method to assess and report thematic saturation in qualitative research Guest Namey Chen",
 "Sample sizes for saturation in qualitative research Hennink Kaiser",
 "What are we doing here Analyzing fifteen years of energy scholarship Sovacool",
 "Technological transitions as evolutionary reconfiguration processes Geels",
 "The multi-level perspective on sustainability transitions responses to seven criticisms Geels",
 "Energy justice a conceptual review Jenkins McCauley Heffron Stephan Rehner",
 "The political economy of the just transition Newell Mulvaney",
 "Towards a theory of just transition neo-Gramscian Winkler",
 "A ladder of citizen participation Arnstein",
 "Perceived usefulness perceived ease of use and user acceptance of information technology Davis",
 "Co-creation and the new landscapes of design Sanders Stappers",
 "Energy cultures a framework for understanding energy behaviours Stephenson",
 "Living Labs as open-innovation networks Leminen Westerlund Nystrom",
 "A systematic review of living lab literature Hossain Leminen Westerlund",
 "Living Lab a methodology between user-centred design and participatory design DellEra Landoni",
 "Impacts of urban living labs on sustainability transitions mechanisms and strategies von Wirth",
 "Posthumanist performativity toward an understanding of how matter comes to matter Barad",
 "Staying with the Trouble Making Kin in the Chthulucene Haraway",
 "Meeting the Universe Halfway quantum physics and the entanglement of matter and meaning Barad",
]

def q(s):
    url = ("https://api.openalex.org/works?per-page=2&search=" + urllib.parse.quote(s)
           + "&select=id,doi,title,publication_year,authorships,primary_location,type,biblio")
    req = urllib.request.Request(url, headers={"User-Agent": "PhDRefCheck/1.0 (mailto:research@example.org)"})
    with urllib.request.urlopen(req, timeout=30) as r:
        return json.load(r)["results"]

out = []
for s in Q:
    items = None
    for attempt in range(3):
        try:
            items = q(s); break
        except Exception as e:
            err = str(e); time.sleep(3)
    if items is None:
        out.append({"q": s, "error": err}); continue
    recs = []
    for it in items:
        names = [a["author"]["display_name"] for a in it.get("authorships", [])][:8]
        loc = (it.get("primary_location") or {}).get("source") or {}
        b = it.get("biblio") or {}
        recs.append({"doi": it.get("doi"), "title": it.get("title"), "year": it.get("publication_year"),
                     "authors": names, "venue": loc.get("display_name"), "vol": b.get("volume"),
                     "issue": b.get("issue"), "pages": (b.get("first_page"), b.get("last_page")), "type": it.get("type")})
    out.append({"q": s, "recs": recs})
    time.sleep(1.2)

with open(r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\openalex_verification.json", "w", encoding="utf-8") as f:
    json.dump(out, f, indent=1, ensure_ascii=False)
with open(r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\openalex_verification.txt", "w", encoding="utf-8") as f:
    for r in out:
        f.write("Q: " + r["q"] + "\n")
        if "error" in r:
            f.write("  ERROR " + r["error"] + "\n\n"); continue
        for x in r["recs"]:
            f.write("  %s | %s | %s | %s %s %s | %s\n" % (x["year"], "; ".join(x["authors"]), x["title"], x["venue"], x["vol"], x["pages"], x["doi"]))
        f.write("\n")
print("DONE", len(out))
