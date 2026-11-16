# -*- coding: utf-8 -*-
import json, time, urllib.parse, urllib.request

DOIS = [
 "10.1177/1534484316671606", "10.5465/amr.2007.26586822", "10.1504/IJIRD.2009.022727",
 "10.1191/1478088706qp063oa", "10.1016/j.erss.2014.02.003", "10.1111/geoj.12008",
 "10.1080/01944366908977225", "10.2307/249008", "10.1080/15710880701875068",
 "10.22215/timreview/602", "10.1111/caim.12061", "10.1086/345321",
 "10.1215/9780822373780", "10.7208/chicago/9780226321394.001.0001",
 "10.1016/j.sheji.2017.08.001", "10.1080/09518398.2017.1336806",
 "10.1016/j.apenergy.2015.01.002", "10.1016/j.enpol.2012.01.044",
 "10.1016/j.enpol.2017.03.018", "10.1016/j.enpol.2006.12.001",
 "10.2307/30036540", "10.5929/2014.4.2.9", "10.1177/1350508410373519",
 "10.1016/j.enpol.2010.05.069", "10.1016/j.techsoc.2021.101850",
 "10.3390/su142114009", "10.1016/j.esd.2025.101777", "10.1016/j.procir.2025.01.042",
 "10.1016/j.erss.2020.101715", "10.1016/j.erss.2022.102525",
 "10.1016/j.erss.2021.102476", "10.1002/sres.2793", "10.3390/su14106291",
 "10.7166/31-3-2421", "10.1016/j.erss.2020.101789", "10.1016/j.erss.2023.103122",
]

out = []
for doi in DOIS:
    url = "https://api.crossref.org/works/" + urllib.parse.quote(doi)
    req = urllib.request.Request(url, headers={"User-Agent": "PhDRefCheck/1.0 (mailto:research@example.org)"})
    try:
        with urllib.request.urlopen(req, timeout=25) as r:
            m = json.load(r)["message"]
        out.append({"doi": doi, "ok": True,
                    "title": (m.get("title") or [""])[0],
                    "year": (m.get("issued", {}).get("date-parts") or [[None]])[0][0],
                    "journal": (m.get("container-title") or [""])[0],
                    "vol": m.get("volume"), "issue": m.get("issue"), "page": m.get("page"),
                    "authors": [a.get("family", "") + ", " + a.get("given", "") for a in m.get("author", [])]})
    except Exception as e:
        out.append({"doi": doi, "ok": False, "err": str(e)})
    time.sleep(1.0)

with open(r"C:\Users\Muhammad Fahad\Desktop\Research paper\research_work\doi_check.json", "w", encoding="utf-8") as f:
    json.dump(out, f, indent=1, ensure_ascii=False)
for x in out:
    if x["ok"]:
        print("OK  ", x["doi"], "|", x["year"], "|", x["title"][:95], "|", x["journal"], x["vol"], x["page"])
        print("     ", "; ".join(x["authors"][:6]))
    else:
        print("FAIL", x["doi"], x["err"])
