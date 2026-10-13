"""First-report chain, Camp ICS-209-PLUS time-area table, and v0.5 figures.

Adds to the v0.4 evaluation (perimeter_gap_evaluation.py, which is re-run here to load the
same observation series; its outputs are unchanged):
1. How fast the first information got out on each fire: ignition, first call, first aircraft
   observation, first satellite overpass with a detection (FIRMS archive), first perimeter,
   first ICS-209 report. -> first_report_eval/first_report_chain.csv
2. Camp 2018 perimeter and size information from ICS-209-PLUS, in the same column layout as the
   team's perimeter supplement (Camp was not in it).
   -> perimeter_supplements/Camp_2018_Perimeter_Time_Area_ICS209PLUS.xlsx
3. Figures for the v0.5 proposal. -> figures/

Every timed event names its source. Library citations use the project RAG page numbers.
Run: uv run --project ~/nasa_rag --with matplotlib python first_report_evaluation.py
"""
import csv, json, math, runpy, zipfile, io
from datetime import datetime, timedelta, timezone
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import matplotlib.dates as mdates

HERE = Path(__file__).parent
DS = HERE / "research_data" / "datasets"
OUT = HERE / "first_report_eval"; OUT.mkdir(exist_ok=True)
FIG = HERE / "figures"; FIG.mkdir(exist_ok=True)

g = runpy.run_path(str(HERE / "perimeter_gap_evaluation.py"))  # loads obs, IGN, HAZARD
obs, IGN, HAZARD = g["obs"], g["IGN"], g["HAZARD"]

# ---------------------------------------------------------------- ICS-209-PLUS rows
sit = list(csv.DictReader(open(DS / "05_incident_resource" / "ICS209PLUS_benchmark_sitreps.csv", encoding="utf-8")))
IDS = {"Pioneer 2016": "2016_4416704_PIONEER", "Carr 2018": "2018_9236519_CARR",
       "Camp 2018": "2018_9220077_CAMP", "LNU 2020": "2020_11850821_LNU LIGHTNING COMPLEX"}
def reports(fire):
    rs = [r for r in sit if r["INCIDENT_ID"] == IDS[fire] and r["REPORT_TO_DATE"]]
    return sorted(rs, key=lambda r: r["REPORT_TO_DATE"])
def ts(s): return datetime.fromisoformat(s[:19]) if s else None
def num(s):
    try: return float(s)
    except (TypeError, ValueError): return None

with zipfile.ZipFile(DS / "05_incident_resource" / "ics209plus-wildfire.zip") as z:
    inc_rows = list(csv.DictReader(io.TextIOWrapper(z.open("ics209plus-wildfire/ics209-plus-wf_incidents_1999to2020.csv"), encoding="utf-8")))
INC = {r["INCIDENT_ID"]: r for r in inc_rows if r["INCIDENT_ID"] in IDS.values()}

# ---------------------------------------------------------------- FIRMS first detection
ORIGIN = {  # point of origin from ICS-209-PLUS FOD coordinates (LNU: Hennessey origin)
 "Pioneer 2016": (43.9503, -115.7619, -6, "Pioneer_2016"), "Carr 2018": (40.67, -122.643, -7, "Carr_2018"),
 "Camp 2018": (39.8139, -121.4347, -8, "Camp_2018"), "LNU 2020": (38.5497, -122.3797, -7, "LNU_2020")}
def km(a, b, c, d):
    p = math.pi / 180
    h = math.sin((c - a) * p / 2) ** 2 + math.cos(a * p) * math.cos(c * p) * math.sin((d - b) * p / 2) ** 2
    return 12742 * math.asin(math.sqrt(h))
def firms_first(fire, radius_km=15):
    la, lo, off, stem = ORIGIN[fire]; ign = IGN[fire]["ignition"]; res = {}
    for inst in ("MODIS", "VIIRS_SNPP"):
        f = DS / "02_active_fire" / f"{stem}_FIRMS_{inst}_archive_subset.csv"
        if not f.exists(): res[inst] = None; continue
        passes = {}
        for r in csv.DictReader(open(f)):
            t = datetime.fromisoformat(r["acq_date"]) + timedelta(hours=int(r["acq_time"]) // 100 + off, minutes=int(r["acq_time"]) % 100)
            if t >= ign - timedelta(hours=1) and km(la, lo, float(r["latitude"]), float(r["longitude"])) < (30 if fire == "LNU 2020" else radius_km):
                passes[t] = passes.get(t, 0) + 1
        res[inst] = sorted(passes.items())[:3]
    return res

# ---------------------------------------------------------------- first-report chain
def hrs(t, fire): return round((t - IGN[fire]["ignition"]).total_seconds() / 3600, 1) if t else None
def first(fire, fam, after=None):
    rs = [r for r in obs[fire] if r["fam"] == fam and r["acres"] and (after is None or r["t"] >= after)]
    return rs[0] if rs else None

EARLY = {  # documented first reports and aircraft observations (event records with sources)
 "Pioneer 2016": dict(call=("17:17 first report by the Casner Fire lead plane, 1 to 1.5 ac", "Gabbert 2016, p.2 (perimeter supplement)"),
                      air=("17:17 lead plane, 1 to 1.5 ac; about 100 ac by 21:27", "Gabbert 2016, p.2 (perimeter supplement)")),
 "Carr 2018": dict(call=("13:10 to 13:11 first calls", "USFS Carr origin investigation 2021, pp.3, 6, 10 (perimeter supplement)"),
                   air=("13:32 air attack, 20 ac", "USFS Carr origin investigation 2021, p.10 (perimeter supplement)")),
 "Camp 2018": dict(call=("06:25 first 911 call; dispatched 06:31; first engine in sight 06:44", "[Maranghides et al 2023, pp.87-88]"),
                   air=("07:52 air attack, about 1,000 ac; 08:01 about 2,000 ac and in Paradise", "[Maranghides et al 2021, pp.81-82]")),
 "LNU 2020": dict(call=("16 to 17 Aug, many lightning ignitions; complex start 06:00 (GOFER), 06:40 or 08:00 (CAL FIRE)", "perimeter supplement (CAL FIRE 2020, 2023; Liu et al. 2024)"),
                  air=("not documented in the project records", "")),
}
chain = []
for fire in IDS:
    ign = IGN[fire]["ignition"]
    fr = firms_first(fire)
    mapped = first(fire, "mapped", after=ign)
    if fire == "Carr 2018":  # the 11:30 and 13:02 GeoMAC stamps precede ignition (v0.4, Section 5.5)
        mapped = first(fire, "mapped", after=datetime(2018, 7, 23, 13, 9))
    sat = first(fire, "satellite", after=ign)
    rep = reports(fire)[0]
    disc = ts(INC[IDS[fire]]["DISCOVERY_DATE"]) if INC.get(IDS[fire]) else None
    def fmt(p):
        return f"{p[0][0]:%d %b %H:%M} ({hrs(p[0][0], fire)} h, {p[0][1]} px)" if p else "none in archive subset"
    chain.append({
        "fire": fire, "ignition_reference": f"{ign:%d %b %Y %H:%M}",
        "first_call_or_report": EARLY[fire]["call"][0], "first_call_source": EARLY[fire]["call"][1],
        "first_aircraft_observation": EARLY[fire]["air"][0], "first_aircraft_source": EARLY[fire]["air"][1],
        "first_MODIS_detection": fmt(fr["MODIS"]), "first_VIIRS_detection": fmt(fr["VIIRS_SNPP"]) if fr.get("VIIRS_SNPP") is not None else "VIIRS not downloaded",
        "first_aircraft_perimeter": f"{mapped['t']:%d %b %H:%M}, {mapped['acres']:,.0f} ac" if mapped else "no aircraft series",
        "h_to_first_aircraft_perimeter": hrs(mapped["t"], fire) if mapped else None,
        "first_satellite_area": f"{sat['t']:%d %b %H:%M}, {sat['acres']:,.0f} ac ({sat['src']})" if sat else "none",
        "h_to_first_satellite_area": hrs(sat["t"], fire) if sat else None,
        "first_ICS209_report": f"{ts(rep['REPORT_TO_DATE']):%d %b %H:%M}, {num(rep['CURR_INCIDENT_AREA']):,.0f} ac",
        "h_to_first_ICS209": hrs(ts(rep["REPORT_TO_DATE"]), fire),
        "ICS209_discovery_time": f"{disc:%d %b %H:%M}" if disc else "", "h_ignition_to_209_discovery": hrs(disc, fire) if disc else None,
    })
with open(OUT / "first_report_chain.csv", "w", newline="") as f:
    w = csv.DictWriter(f, fieldnames=list(chain[0])); w.writeheader(); w.writerows(chain)

# ---------------------------------------------------------------- Camp time-area table (supplement layout)
camp_rows = []
def add(t, a, src, typ, com): camp_rows.append((t, a, src, typ, com))
D = datetime
EV = [
 (D(2018,11,8,6,25), None, "Maranghides et al 2023, pp.87-88 (NIST TN 2252)", "Event: First 911 call", "First 911 call reporting the fire at 06:25; response dispatched 06:31. Use 06:25 as the earliest documented detection, not the 08:29 ICS-209 discovery time."),
 (D(2018,11,8,6,29), None, "CAL FIRE 2018 Camp Fire Green Sheet, p.7", "Event: Agency start-time record", "CAL FIRE start time 06:29. Keep separate from the 06:25 call time; the four-minute spread is the start-time uncertainty for Camp."),
 (D(2018,11,8,6,44), None, "Maranghides et al 2023, pp.87-88", "Event: First engine on scene", "First engine in sight of the fire at 06:44, essentially driving time from the nearest station. Reports about 35 mph sustained wind and declares potential for a major incident (Maranghides et al 2021, p.91)."),
 (D(2018,11,8,6,45), None, "Maranghides et al 2021, p.91", "Event: Second ignition", "Investigators place a second power-line ignition near Rim Road and Concow Road at 06:45; it merged into the Camp Fire. A single origin and start time understate the early fire area."),
 (D(2018,11,8,6,55), None, "Maranghides et al 2023, p.95", "Event: First evacuation order", "IC orders Pulga evacuated at 06:55, three minutes before fire was recorded in Pulga; CodeRED followed 18 minutes after the order (07:13)."),
 (D(2018,11,8,7,10), 250, "Maranghides et al 2021, pp.91-92", "Observation: Ground estimate (engine)", "Engine reports the fire at 80 to 120 ha (200 to 300 ac) with rapid spread toward Concow Reservoir. Area entered as the midpoint; carry the range."),
 (D(2018,11,8,7,50), None, "Maranghides et al 2021, p.197; Camp_2018_NIST_burnovers.csv", "Event: Crew hazard", "First of 19 NIST-documented burnovers (Hoffman Road and Concow Road, 07:50). Concow evacuation order was 07:37, 13 minutes before (Maranghides et al 2023, Table 12)."),
 (D(2018,11,8,7,52), 1000, "Maranghides et al 2021, pp.81-82", "Observation: Aircraft visual estimate (air attack radio)", "Air attack off Chico airport: about 1,000 ac; orders 6 tankers, 1 lead plane, 6 Type 1 helicopters. First aerial size estimate on record; radio only, no mapped boundary."),
 (D(2018,11,8,8,1), 2000, "Maranghides et al 2021, pp.81-82", "Observation: Aircraft visual estimate (air attack radio)", "Air attack: 'The fire is now in Paradise ... About 2000 ac along Pentz Rd.' NIST notes information on spread was coming from 911 calls because firefighters were saving lives."),
 (D(2018,11,8,8,29), None, "ICS-209-PLUS v3, incident table (DISCOVERY_DATE)", "Event: ICS-209 discovery time", "ICS-209 discovery time 08:29, two hours after the first 911 call. Do not use it as T_detect."),
 (D(2018,11,8,9,3), None, "Maranghides et al 2023, pp.87-88", "Event: Evacuation orders", "Nearly all of Paradise and Magalia under evacuation orders by 09:03."),
 (D(2018,11,8,10,45), None, "Maranghides et al 2021, pp.156-157", "Observation: Satellite image (Landsat 8)", "Landsat 8 imaged the fire at 10:45; NIST digitized a perimeter from it (Map 6). Area not given in the text; the image validates NIST's ground observations."),
 (D(2018,11,8,14,15), None, "CAL FIRE 2018 Green Sheet, p.8", "Event: Crew hazard (Rattlesnake Flats)", "Firing operation organised on Rattlesnake Flats Road with the fire about 1/4 mile north, backing south. Start of the documented exposure window."),
 (D(2018,11,8,14,45), None, "CAL FIRE 2018 Green Sheet, pp.8-9", "Event: Crew hazard (Rattlesnake Flats)", "Wind shifts suddenly to about 25 mph, 10 to 15 ft flames; fire crosses the road both sides of the crew. NIST logs the burnover at 15:15 (Maranghides et al 2021, p.197), a 30 minute source disagreement."),
 (D(2018,11,8,15,0), None, "CAL FIRE 2018 Green Sheet, p.9 (via Research Data Index)", "Event: Notification failure", "Emergency traffic call went unanswered because of radio traffic volume. A T_notify failure, not a delay."),
 (D(2018,11,8,18,29), 55000, "CAL FIRE 2018 Green Sheet, p.7", "Observation: Agency retrospective estimate", "Green Sheet: in the first 12 hours the fire ran about 15 miles downwind and exceeded 55,000 ac, spotting up to a mile ahead. Retrospective; it was not the number on record that evening (18:30 ICS-209: 16,000 ac)."),
]
for e in EV: add(*e)
for r in reports("Camp 2018"):
    t = ts(r["REPORT_TO_DATE"]); a = num(r["CURR_INCIDENT_AREA"])
    fb = ", ".join(x for x in (r["FIRE_BEHAVIOR_1"], r["FIRE_BEHAVIOR_2"], r["FIRE_BEHAVIOR_3"]) if x)
    add(t, a, "ICS-209-PLUS v3, sitreps (CURR_INCIDENT_AREA)", "Observation: Incident report (method unspecified)",
        f"Report period {r['REPORT_FROM_DATE'][5:16]} to {r['REPORT_TO_DATE'][5:16]}. {num(r['PCT_CONTAINED_COMPLETED']):.0f}% contained; "
        f"fire behaviour: {fb or 'not given'} ({r['GEN_FIRE_BEHAVIOR'] or 'n/a'}). No mapped boundary in the 209.")
for r in obs["Camp 2018"]:
    if r["fam"] == "mapped": add(r["t"], round(r["acres"], 1), "NIFC Open Data / GeoMAC", r["typ"] or "Aircraft perimeter", "Archive time stamp may record processing rather than flight time (v0.4, Section 5.5).")
    if r["fam"] == "satellite": add(r["t"], round(r["acres"], 1), "NASA FEDS 12 h (research_data)", "Reconstruction: VIIRS satellite reconstruction", "375 m VIIRS pixels; reads about 2% below the official final area on Camp.")
inc = INC[IDS["Camp 2018"]]
add(D(2018,11,25,19,30), num(inc["FIRED_ACRES"]), "ICS-209-PLUS v3, incident table (FIRED link)", "Reconstruction: FIRED daily MODIS event (incident-level)",
    f"FIRED event {inc['FIRED_ID'][:6]} linked in ICS-209-PLUS: {num(inc['FIRED_ACRES']):,.0f} ac (23% below the official 153,336), largest one-day growth "
    f"{num(inc['FIRED_MAX_GROWTH']):,.0f} ac. Its ignition point ({float(inc['FIRED_IG_LAT']):.3f}, {float(inc['FIRED_IG_LON']):.3f}) is the first burned MODIS pixel, "
    f"about 17 km west of the origin. Incident-level only: no daily geometry in ICS-209-PLUS.")
add(D(2018,11,9,7,45), 52000, "ICS-209-PLUS v3, incident table (WF_MAX_FSR, WF_MAX_GROWTH_DATE)", "Derived: largest growth between reports",
    "ICS-209-PLUS dates the largest growth (52,000 ac) to the report ending 9 Nov 07:45, because the overnight report was the first to record 8 Nov growth. "
    "The fire actually made that run on the morning of 8 Nov. Report-derived growth dates lag the fire by a reporting cycle.")
camp_rows.sort(key=lambda r: r[0])

wb = Workbook(); ws = wb.active; ws.title = "Camp 2018"
hdr = ["Local time / date\nPST (UTC-08:00)", "Area (acres)", "Source", "Data type", "Comment"]
ws.append(hdr)
for t, a, s, ty, c in camp_rows: ws.append([t, a, s, ty, c])
for c in ws[1]: c.font = Font(bold=True); c.alignment = Alignment(wrap_text=True, vertical="top")
for row in ws.iter_rows(min_row=2):
    row[0].number_format = "yyyy-mm-dd hh:mm"; row[1].number_format = "#,##0"
    for c in row: c.alignment = Alignment(wrap_text=True, vertical="top")
for col, wdt in zip("ABCDE", (18, 12, 38, 34, 90)): ws.column_dimensions[col].width = wdt
ws.freeze_panes = "A2"
w2 = wb.create_sheet("ICS-209-PLUS reports")
cols = ["REPORT_FROM_DATE", "REPORT_TO_DATE", "CURR_INCIDENT_AREA", "ACRES", "PCT_CONTAINED_COMPLETED", "PCT_PERIM_TO_BE_CONTAINED",
        "FIRE_BEHAVIOR_1", "FIRE_BEHAVIOR_2", "FIRE_BEHAVIOR_3", "GEN_FIRE_BEHAVIOR", "TOTAL_PERSONNEL", "TOTAL_AERIAL",
        "STR_DESTROYED", "FATALITIES", "PROJ_INCIDENT_AREA", "SIGNIF_EVENTS_SUMMARY", "REMARKS"]
w2.append(cols)
for r in reports("Camp 2018"):
    w2.append([num(r[c]) if c in ("CURR_INCIDENT_AREA", "ACRES", "PCT_CONTAINED_COMPLETED", "PCT_PERIM_TO_BE_CONTAINED", "TOTAL_PERSONNEL", "TOTAL_AERIAL", "STR_DESTROYED", "FATALITIES", "PROJ_INCIDENT_AREA") else r[c] for c in cols])
for c in w2[1]: c.font = Font(bold=True)
w2.freeze_panes = "C2"
w3 = wb.create_sheet("Read me")
for line in [
    "Camp Fire 2018 perimeter and size record, built to match the layout of 'Wildfire Perimeter Time Area Sources.xlsx' (which covers Pioneer, Carr and LNU).",
    f"Built {datetime.now():%Y-%m-%d} by first_report_evaluation.py from: ICS-209-PLUS v3 (St. Denis et al. 2023; research_data/datasets/05_incident_resource), NIFC GeoMAC and NASA FEDS histories (research_data/datasets/01_progression_perimeters), and NIST / CAL FIRE documents in the project reference library.",
    "What ICS-209-PLUS holds for perimeter: no geometry. It gives the reported size per report period (CURR_INCIDENT_AREA, as reported; ACRES is the cleaned field), percent contained, percent of perimeter to be contained (100% on every Camp report, so uninformative), fire behaviour codes, narratives, and an incident-level link to the FIRED satellite event.",
    "Times are local PST. ICS-209 rows are placed at the end of the report period (REPORT_TO_DATE), which is when the number became available.",
    "Sheet 'ICS-209-PLUS reports' lists all 38 Camp reports with the fields used.",
]: w3.append([line])
w3.column_dimensions["A"].width = 160
for row in w3.iter_rows():
    for c in row: c.alignment = Alignment(wrap_text=True, vertical="top")
wb.save(HERE / "perimeter_supplements" / "Camp_2018_Perimeter_Time_Area_ICS209PLUS.xlsx")

# ---------------------------------------------------------------- figures
BLUE, ORANGE, AQUA, INK, MUTED, GRID = "#2a78d6", "#eb6834", "#1baf7a", "#0b0b0b", "#52514e", "#e4e3df"
HAZ = "#8e6ad8"  # violet hazard band, readable against the white page
plt.rcParams.update({"font.family": "Helvetica", "font.size": 8.5, "axes.edgecolor": MUTED, "axes.labelcolor": INK,
                     "xtick.color": MUTED, "ytick.color": MUTED, "axes.spines.top": False, "axes.spines.right": False})
WINDOW = {"Pioneer 2016": 50, "Carr 2018": 10, "Camp 2018": 6, "LNU 2020": 8}
fig, axes = plt.subplots(2, 2, figsize=(7.2, 5.2))
for ax, fire in zip(axes.flat, IDS):
    ign = IGN[fire]["ignition"]; end = ign + timedelta(days=WINDOW[fire])
    dd = lambda t: (t - ign).total_seconds() / 86400
    for fam, col, lab, style in (("satellite", AQUA, "Satellite (FEDS / GOFER)", "-"), ("mapped", BLUE, "Aircraft perimeter", "o")):
        rs = [r for r in obs[fire] if r["fam"] == fam and r["acres"] and ign <= r["t"] <= end]
        if fire == "Carr 2018" and fam == "mapped": rs = [r for r in rs if r["t"] >= datetime(2018, 7, 23, 13, 9)]
        if rs:
            x = [dd(r["t"]) for r in rs]; y = [r["acres"] / 1000 for r in rs]
            if style == "o": ax.plot(x, y, "o", ms=3.2, color=col, mec=col, mew=0, label=lab, zorder=3)
            else: ax.plot(x, y, "-", lw=1.4, color=col, label=lab, zorder=2)
    rp = [(ts(r["REPORT_TO_DATE"]), num(r["CURR_INCIDENT_AREA"])) for r in reports(fire)]
    rp = [(t, a) for t, a in rp if a and t <= end]
    ax.step([0] + [dd(t) for t, _ in rp], [0] + [a / 1000 for _, a in rp], where="post", lw=1.6, color=ORANGE, label="ICS-209 size on record", zorder=4)
    for name, h0, h1 in HAZARD.get(fire, []):
        if h0 <= end: ax.axvspan(dd(h0), dd(min(h1, end)), color=HAZ, alpha=0.28, lw=0, zorder=0)
    ax.set_title(fire, loc="left", fontsize=9.5, color=INK, fontweight="bold")
    ax.set_xlim(0, WINDOW[fire]); ax.set_ylim(bottom=0)
    ax.grid(axis="y", color=GRID, lw=0.6); ax.set_xlabel("Days since ignition"); ax.set_ylabel("Thousand acres")
h, l = axes.flat[1].get_legend_handles_labels()
from matplotlib.patches import Patch
h.append(Patch(color=HAZ, alpha=0.28)); l.append("Documented crew hazard window")
fig.legend(h, l, loc="lower center", ncol=4, frameon=False, fontsize=8)
fig.tight_layout(rect=(0, 0.06, 1, 1))
fig.savefig(FIG / "fig1_size_on_record_vs_fire.png", dpi=220); plt.close(fig)

# Camp first 36 hours: what was on record versus what happened
ign = datetime(2018, 11, 8, 6, 25); end = ign + timedelta(hours=36)
fig, ax = plt.subplots(figsize=(7.2, 3.6))
hh = lambda t: (t - ign).total_seconds() / 3600
for fam, col, lab in (("satellite", AQUA, "Satellite area (FEDS)"), ("mapped", BLUE, "Aircraft perimeter (GeoMAC)")):
    rs = [r for r in obs["Camp 2018"] if r["fam"] == fam and r["acres"] and ign <= r["t"] <= end]
    ax.plot([hh(r["t"]) for r in rs], [r["acres"] / 1000 for r in rs], "o-" if fam == "mapped" else "s-", ms=4, lw=1.2, color=col, mec=col, mew=0, label=lab, zorder=3)
rp = [(ts(r["REPORT_TO_DATE"]), num(r["CURR_INCIDENT_AREA"])) for r in reports("Camp 2018") if ts(r["REPORT_TO_DATE"]) <= end]
ax.step([0] + [hh(t) for t, _ in rp] + [36], [0] + [a / 1000 for _, a in rp] + [rp[-1][1] / 1000], where="post", lw=1.8, color=ORANGE, label="ICS-209 size on record", zorder=4)
radio = [(D(2018,11,8,7,10), 0.25), (D(2018,11,8,7,52), 1.0), (D(2018,11,8,8,1), 2.0)]
ax.plot([hh(t) for t, _ in radio], [a for _, a in radio], "D", ms=4, color=INK, label="Ground / air attack radio estimate", zorder=5)
ax.axvspan(hh(D(2018,11,8,7,50)), hh(D(2018,11,8,13,10)), color=HAZ, alpha=0.28, lw=0, zorder=0)
ax.axvspan(hh(D(2018,11,8,14,45)), hh(D(2018,11,8,15,30)), color=HAZ, alpha=0.28, lw=0, zorder=0)
for t, txt, y in ((D(2018,11,8,6,25), "06:25 first 911 call", 119), (D(2018,11,8,7,50), "07:50 first burnovers (14 by 13:10)", 111),
                  (D(2018,11,8,14,45), "14:45 Rattlesnake Flats burnover", 119), (D(2018,11,8,18,30), "18:30 first ICS-209: 16,000 ac", 24)):
    ax.axvline(hh(t), color=MUTED, lw=0.6, ls=":", zorder=1); ax.text(hh(t) + 0.3, y, txt, fontsize=7.5, color=INK)
ax.set_xlim(0, 36); ax.set_ylim(0, 128); ax.set_xticks(range(0, 37, 6))
ax.set_xlabel("Hours since first 911 call (06:25 PST, 8 Nov 2018)"); ax.set_ylabel("Thousand acres")
ax.grid(axis="y", color=GRID, lw=0.6)
ax.legend(loc="lower right", bbox_to_anchor=(1, 0.22), frameon=False, fontsize=7.5)
fig.tight_layout(); fig.savefig(FIG / "fig2_camp_first_36h.png", dpi=220); plt.close(fig)

json.dump({"chain": chain, "camp_rows": len(camp_rows)}, open(OUT / "results.json", "w"), indent=1, default=str)
print(json.dumps(chain, indent=1, default=str)); print("camp rows", len(camp_rows))
