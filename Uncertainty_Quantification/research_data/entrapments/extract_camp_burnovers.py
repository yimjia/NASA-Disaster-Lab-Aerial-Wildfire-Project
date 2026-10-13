"""Extract the 19 Camp Fire burnover events from NIST TN 2135 Appendix B
(Maranghides et al 2021) into a CSV. Source PDF: ../../../sources/S08_NIST_TN2135_Camp_Fire_Progression_Timeline.pdf
Run: uv run --project ~/nasa_rag python extract_camp_burnovers.py"""
import re, subprocess, csv
from pathlib import Path
SRC = Path(__file__).resolve().parents[3] / "sources/S08_NIST_TN2135_Camp_Fire_Progression_Timeline.pdf"
FIRST = 198
txt = subprocess.run(["pdftotext", "-f", str(FIRST), "-l", "262", "-layout", str(SRC), "-"], capture_output=True, text=True).stdout
pages = txt.split("\f")
# duration minutes from NIST Table 38 (p.197 of PDF)
DUR = {1:40,2:70,4:360,6:150,7:60,8:140,9:60,10:120,11:90,12:90,15:15,16:120}
FR = re.compile(r"fire ?fighter|firefighter|engine|hand crew|inmate|shelter|dozer|strike team|CAL FIRE|BCFD|PPD|battalion|captain", re.I)
rows, cur = [], None
for pi, pg in enumerate(pages):
    for line in pg.splitlines():
        s = re.sub(r"\s+", " ", line).strip()
        m = re.match(r"^(\d{1,2})\. ([A-Z].+)$", s)
        if m and int(m.group(1)) == len(rows) + 1 and int(m.group(1)) <= 19:
            cur = {"id": int(m.group(1)), "name": m.group(2), "pdf_page": FIRST + pi, "text": []}; rows.append(cur); continue
        if cur is None: continue
        for key in ("Date/Time", "Location", "Coordinates"):
            if s.startswith(key + ":"): cur[key] = s.split(":", 1)[1].strip()
        cur["text"].append(s)
out = []
for r in rows:
    body = " ".join(r["text"])
    lat, lon = re.findall(r"-?\d+\.\d+", r.get("Coordinates", ""))[:2] or ("", "")
    summ = re.search(r"Summary: (.+?)(?: Time Observation| Time \(|$)", body)
    hits = sorted({h.lower() for h in FR.findall(body)})
    out.append({"id": r["id"], "event": r["name"], "date_time": r.get("Date/Time", "").replace("–", " to "),
        "duration_min_table38": DUR.get(r["id"], "n/d"), "location": r.get("Location", ""), "lat": lat, "lon": lon,
        "first_responders_mentioned": "yes" if hits else "no", "responder_keywords": "; ".join(hits),
        "summary": (summ.group(1)[:700] if summ else "").replace("–", " to ").replace("—", ", "),
        "citation": f"[Maranghides et al 2021, p.{r['pdf_page']}]"})
with open("Camp_2018_NIST_burnovers.csv", "w", newline="") as f:
    w = csv.DictWriter(f, fieldnames=list(out[0])); w.writeheader(); w.writerows(out)
print(len(out), "events;", sum(o["first_responders_mentioned"] == "yes" for o in out), "mention first responders")
for o in out: print(o["id"], o["event"], "|", o["date_time"], "|", o["responder_keywords"][:80])
