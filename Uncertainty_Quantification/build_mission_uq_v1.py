"""Build the five mission uncertainty documents, version 1.0 (M1 to M5).

One Word document per Week 6 mission, each under two pages, each answering: how the mission's uncertainty
is quantified, with a worked calculation; which drone types serve it and with what parameters; the weather
that was actually on the four fires (Weather Data folder, via weather_evaluation.py); and who uses the
number (incident commander or crew level and lookouts). GOFER is not used anywhere in v1.0.

Run after weather_evaluation.py:
    uv run --with python-docx --with pandas python build_mission_uq_v1.py
"""
import json, math, random
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn

HERE = Path(__file__).parent
OUT = HERE / "Mission UQ v1"; OUT.mkdir(exist_ok=True)
DATE = "20 October 2026"
WX = json.load(open(HERE / "weather_eval" / "results.json"))
BLACK = RGBColor(0, 0, 0)

# Representative station per fire: the station nearest the critical run with gust data.
REP = {"Camp": "JBGC1", "Carr": "MMOC1", "Pioneer": "BCFI1", "LNU": "ATLC1"}
WIN = {"Camp": "8 Nov 2018, 06:00 to 18:00 PST", "Carr": "26 Jul 2018, 12:00 to 23:00 PDT",
       "Pioneer": "29 Aug 2016, 08:00 to 23:00 MDT", "LNU": "18 Aug 12:00 to 19 Aug 2020 12:00 PDT"}
DIRS = {"Camp": "NNE to NE (14 to 42°)", "Carr": "variable, W to NW by evening", "Pioneer": "E in the morning, SW to W after 15:00",
        "LNU": "SW to WSW (213 to 255°) to 20:00, then WNW"}


def g(fire, drone, kind="crit", gate="p_fly_gust"):
    key = {"crit": "crit_gate", "perim": "perimeter_gate", "night": "perimeter_night_gate"}[kind]
    return WX[fire][REP[fire]][key][drone][gate]


def pct(x):
    return "under 1%" if 0 < x < 0.005 else f"{round(100 * x):d}%"


def age_h(R, p):
    """Expected information age (h) when a pass is due every R hours and succeeds with probability p.
    Gaps between successes are R x Geometric(p); renewal theory gives E[A] = R (2 - p) / (2 p)."""
    return float("inf") if p == 0 else R * (2 - p) / (2 * p)


# ============================================================ document helpers
def new_doc():
    d = Document()
    st = d.styles["Normal"]
    st.font.name = "Times New Roman"; st.font.size = Pt(11); st.font.color.rgb = BLACK
    st.element.rPr.rFonts.set(qn("w:eastAsia"), "Times New Roman")
    st.paragraph_format.space_after = Pt(3); st.paragraph_format.space_before = Pt(0)
    for s in d.sections:
        s.left_margin = s.right_margin = Inches(0.7); s.top_margin = s.bottom_margin = Inches(0.5)
    return d


def run(p, text, bold=False, italic=False, size=None):
    r = p.add_run(text); r.bold = bold; r.italic = italic
    r.font.name = "Times New Roman"; r.font.color.rgb = BLACK
    if size: r.font.size = Pt(size)
    return r


def H(d, text, size=12):
    p = d.add_paragraph(); p.paragraph_format.space_before = Pt(5); p.paragraph_format.space_after = Pt(2)
    p.paragraph_format.keep_with_next = True
    run(p, text, bold=True, size=size)


def P(d, *parts, after=3):
    """parts: strings, or (text, 'b') / (text, 'i') tuples."""
    p = d.add_paragraph(); p.paragraph_format.space_after = Pt(after)
    for x in parts:
        if isinstance(x, tuple):
            run(p, x[0], bold="b" in x[1], italic="i" in x[1])
        else:
            run(p, x)
    return p


def B(d, lead, text):
    p = d.add_paragraph(style="List Bullet"); p.paragraph_format.space_after = Pt(1)
    run(p, lead, bold=True); run(p, text)


def T(d, header, rows, widths, size=9.5):
    t = d.add_table(rows=1, cols=len(header)); t.style = "Table Grid"; t.autofit = False
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]; c.text = ""
        run(c.paragraphs[0], h, bold=True, size=size)
    for r in rows:
        cells = t.add_row().cells
        for i, v in enumerate(r):
            cells[i].text = ""
            run(cells[i].paragraphs[0], str(v), size=size)
    # Pages and some Word builds read the grid, not the cell widths, so set both
    for gc, w in zip(t._tbl.tblGrid.findall(qn("w:gridCol")), widths):
        gc.set(qn("w:w"), str(int(w * 1440)))
    for row in t.rows:
        for i, w in enumerate(widths):
            row.cells[i].width = Inches(w)
            for para in row.cells[i].paragraphs:
                para.paragraph_format.space_after = Pt(0)
    d.add_paragraph().paragraph_format.space_after = Pt(0)
    return t


def header_block(d, mission, title, version_note):
    p = d.add_paragraph(); p.paragraph_format.space_after = Pt(1)
    run(p, f"{mission} {title}: Quantified Uncertainty, Version 1.0", bold=True, size=14)
    p = d.add_paragraph(); p.paragraph_format.space_after = Pt(4)
    run(p, f"NASA Aerial Wildfire Project, ACERO Disaster Lab working group. {DATE}. {version_note}", italic=True, size=11)


def weather_table(d, fires, extra_col=None):
    rows = []
    for f in fires:
        r = WX[f][REP[f]]
        row = [f, WIN[f], f"{REP[f]} {r['name'].title()}", f"{r['crit_sust_max_mph']:.0f}", f"{r['crit_gust_max_mph']:.0f}",
               f"{r['crit_rh_min']:.0f}", "n/a" if r.get("crit_fuel_moisture_min") is None else f"{r['crit_fuel_moisture_min']:.0f}", DIRS[f]]
        if extra_col: row.append(extra_col(f))
        rows.append(row)
    hdr = ["Fire", "Critical window", "Station", "Sust. max mph", "Gust max mph", "Min RH %", "Min 10-h fuel moist. %", "Wind from"]
    w = [0.6, 1.45, 1.2, 0.55, 0.55, 0.5, 0.65, 1.1]
    if extra_col:
        hdr.append(extra_col.__doc__); w = [0.55, 1.25, 1.05, 0.5, 0.5, 0.45, 0.6, 0.95, 0.95]
    T(d, hdr, rows, w, size=9)


SRC = ("Sources: Weather Data folder via weather_evaluation.py (RAWS 10-min mean at 20 ft; gust = prior-hour peak); "
       "Week 6 workbook; UQ proposal v0.5 (W5 = Week 5 Platform tab). Assumed values are labelled.")

# ============================================================ shared numbers
CAMP_GROWTH_ACH = 55000 / 12  # CAL FIRE: more than 55,000 ac in the first 12 h (v0.5)
AX, SB, T30, T40 = "Alta X (22 mph)", "SwitchBlade-Elite (25 mph)", "Type 1 FW if 30 mph", "Type 1 FW if 40 mph"

# ============================================================ M1 Fireline mapping
d = new_doc()
header_block(d, "M1", "Fireline Mapping", "First version of the M1 uncertainty calculation; builds on UQ proposal v0.5.")
H(d, "Bottom line")
pA = g("Camp", AX); ageA = age_h(1, pA); pS = g("Camp", AX, gate="p_fly_sustained")
p30 = g("Camp", T30); age30 = age_h(1, p30)
B(d, f"{ageA:.1f} h versus 0.5 h. ", f"Expected age of the newest Camp fireline map for an Alta X flying a 1 h revisit on the first day: "
  f"0.5 h if the 22 mph limit is read against sustained wind, {ageA:.1f} h if it is read against gusts (flyable in {pct(pA)} of hours at Jarbo Gap). "
  "The gap between the two readings is larger than the drone's benefit.")
B(d, "Most of the first day unseen. ", f"Read against gusts, the Alta X flies in about 1 of 13 hours, so its newest map would typically be most of the day old; "
  f"the steady-state {ageA:.1f} h is an upper bound longer than the window itself. At Camp's average first-day growth (more than 55,000 ac in 12 h, about {CAMP_GROWTH_ACH:,.0f} ac/h) "
  f"a Type 1 fixed wing rated to 30 mph leaves about {round(CAMP_GROWTH_ACH * age30, -2):,.0f} ac unmapped at any moment.")
B(d, "Gusts ground the drones, and the Type 1 limit is the largest unknown. ", "At Jarbo Gap, on the Camp fire, sustained wind peaked at 18 mph and gusts at 40; other Camp stations reached 25 to 29 mph sustained and 59 mph gusts, so the Alta X gust gate runs 0% to 100% by station. "
  f"No FVR-90 class limit is published; Camp flyable hours move from {pct(g('Camp', T30))} to {pct(g('Camp', T40))} between a 30 and a 40 mph limit.")
H(d, "Who uses it")
P(d, ("Incident commander and planning (situation unit, fire behavior analyst): ", "b"),
  "map age decides whether resource orders and evacuation triggers rest on a current line. ",
  ("Division supervisors and lookouts: ", "b"), "head and flank position on the active edge, refreshed within the hour.")
H(d, "Drone types serving M1 and their parameters")
T(d, ["Drone", "Type", "Parameters", "Wind limit", "M1 role"], [
    ["L3Harris FVR-90", "NWCG Type 1 FW, VTOL", "12 to 18 h with 8 to 22 lb thermal video; Type 1 class: 3,500 to 8,000 ft AGL, 50 mi range", "Not found (30 and 40 mph run)", "Whole-fire perimeter every pass"],
    ["SuperVolo (NASA Langley)", "VTOL FW research", "8 h; 45 to 75 mph cruise; flown as relay in TCL 1", "Not found", "Second mapping tier, relay"],
    ["Alta X (Freefly)", "NWCG Type 3 rotor", "50 min empty, 22 min at 22 lb; FY22 average sortie 17 min; 2,000 ft AGL and below", "22 mph (secondary)", "Head and spot detail close in"],
    ["SwitchBlade-Elite", "Type 3 band", "Up to 50 min, average sortie 18 min; 2 to 2.3 kg payload", "25 mph (W5)", "Close-in edge mapping"],
], [1.25, 1.1, 2.6, 1.0, 1.35])
H(d, "Live weather on the fires and what it means for tracking the fireline")
weather_table(d, ["Camp", "Carr", "Pioneer", "LNU"])
P(d, "Wind direction places the head and sets where the drone flies first. At Camp the NE wind at Jarbo Gap drove the head southwest "
  "toward Paradise (air attack, 08:01); a 1 h revisit maps that head every pass and the flanks every other pass. "
  "At LNU, SW to WSW wind at Atlas Peak rose to 29 mph sustained, 41 mph gusts by 18:00 on 18 August, hours before the overnight merge. "
  "Pioneer's 29 August run happened with surface gusts of 25 mph or less: it had pyroconvective updrafts (v0.5), so a surface wind gate alone would wrongly clear drones to fly near it.")
P(d, ("Tracking upslope runs and fuels. ", "b"), "Slope from USGS 3DEP marks the flanks where a run will accelerate upslope; those flanks get the next pass after the head. "
  "Fuels were dry (10-h moisture 2 to 5%, table above). Next: LANDFIRE fuel loading against drone-derived fuel maps, burn history from Landsat in Google Earth Engine, and the burning index from the RAWS records.")
H(d, "How the uncertainty is quantified")
P(d, ("Fireline map error at time t: ", "b"), "E(t) = G x A(t), the area grown since the newest map, where G is the growth rate (ac/h) and A(t) the information age (h). "
  "With a pass due every R hours that succeeds with probability p, the gaps between good maps are R x Geometric(p), so the expected age is ",
  ("E[A] = R (2 - p) / (2p).", "b"), " p is the hourly wind gate, the share of hours the station wind is at or under the drone's limit: "
  "sustained wind gives the optimistic bound, gusts the conservative one. Both are 20 ft winds; flight-level wind is stronger, so both overstate availability.")
rows = []
for name, lab in [(AX, "Alta X, 22 mph"), (SB, "SwitchBlade-Elite, 25 mph"), (T30, "Type 1 FW if 30 mph"), (T40, "Type 1 FW if 40 mph")]:
    ps, pg = g("Camp", name, gate="p_fly_sustained"), g("Camp", name)
    a_s, a_g = age_h(1, ps), age_h(1, pg)
    rows.append([lab, pct(ps), pct(pg), f"{a_s:.1f} to {a_g:.1f}", f"{round(CAMP_GROWTH_ACH * a_s, -2):,.0f} to {round(CAMP_GROWTH_ACH * a_g, -2):,.0f}",
                 pct(g('Camp', name, 'perim'))])
T(d, ["Camp (Jarbo Gap), R = 1 h", "p sustained", "p gust", "E[A] h", "Unseen ac", "p gust, whole Camp period"], rows, [1.6, 0.85, 0.7, 0.9, 1.35, 1.6])
P(d, "E[A] above 12 h exceeds the 13 h window and unseen area above 55,000 ac exceeds the fire's own 12 h growth: read those cells as 'no useful map'. Growth was front-loaded, so a constant G understates the early hours.")
P(d, ("Kinds of uncertainty. ", "b"), "p: aleatory (weather), measured hourly. Type 1 limit and gust versus sustained reading: epistemic, fixed by asking Overwatch Aero or NASA. "
  "Assumes passes independent hour to hour; gusty hours cluster, which lengthens the worst gaps.")
P(d, ("v1.1 needs: ", "b"), "Yiming's hourly perimeters as the reference for G (replaces GOFER); the FVR-90 wind limit; flight-level wind (HRRR); a plume term for Pioneer-type runs; perimeter length and swath to turn R into a fleet size (Week 6 Q1).", after=2)
P(d, (SRC, "i"))
d.save(OUT / "M1_Fireline_Mapping_UQ_v1.0.docx")

# ============================================================ M2 Spot fire detection
d = new_doc()
header_block(d, "M2", "Spot Fire Detection", "First version of the M2 uncertainty calculation; builds on UQ proposal v0.5.")


def tdet_min(R_min, p_fly, Pd):
    q = p_fly * Pd
    return float("inf") if q == 0 else R_min / 2 + R_min * (1 - q) / q


R2 = 30


def gsd(h_m, hfov_deg=40, n=640):
    return 2 * h_m * math.tan(math.radians(hfov_deg / 2)) / n

H(d, "Bottom line")
t_best = tdet_min(R2, g("Camp", AX, gate="p_fly_sustained"), 0.9); t_worst = tdet_min(R2, g("Camp", AX), 0.9)
B(d, f"{t_best:.0f} to {t_worst:.0f} minutes. ", f"Expected time to detect a new spot fire at Camp on the first day for an Alta X searching every {R2} min with a 90% per-pass detection probability, "
  "depending on whether its 22 mph limit is read against sustained wind or gusts.")
B(d, "Against Camp's 85 minutes. ", f"Only the Alta X read against gusts ({tdet_min(R2, g('Camp', AX), 0.9):.0f} min) misses it; the SwitchBlade-Elite ({tdet_min(R2, g('Camp', SB), 0.9):.0f} min) and a Type 1 at 30 mph ({tdet_min(R2, g('Camp', T30), 0.9):.0f} min) detect inside it at Pd 0.9. The 85 minutes runs from the first 911 call to the first burnovers, so it is a comparison, not a deadline.")
B(d, "Sensor probability is the second unknown. ", "No drone in the library has an operational spot fire detection record (Week 6 Platform, M2 PARTIAL), so per-pass detection Pd runs as 50%, 80% and 95%.")
H(d, "Who uses it")
P(d, ("Lookouts and division supervisors: ", "b"), "a spot fire across the line changes escape routes and safety zones within minutes, so the alert must reach them first. ",
  ("Incident commander and operations: ", "b"), "detection time sets whether initial attack resources are ordered before the spot becomes a new front.")
H(d, "Drone types serving M2 and their parameters")
T(d, ["Drone", "Type", "Parameters", "Wind limit", "M2 role"], [
    ["Type 1 and 2 FW (FVR-90, ScanEagle, Aerosonde class)", "NWCG Type 1 / 2", "Type 1: 6 to 14 h, mid-wave IR; Type 2: 1 to 6 h, long-wave IR; above all other incident aircraft", "Not found", "Wide-area IR search"],
    ["Alta X / SwitchBlade-Elite", "NWCG Type 3", "17 to 18 min average sorties; 5 mi range", "22 / 25 mph", "Confirm and geolocate flagged hot spots, cut false alarms"],
    ["DJI Mavic 2 Enterprise Advanced (sensor reference)", "Type 4 band, research", "640 x 512 uncooled thermal, 8 to 14 µm, plus 4K RGB; flown at 60 m AGL for FLAME 2", "Not in library", "Fused RGB-IR flame detection benchmark"],
], [1.6, 1.0, 2.55, 0.8, 1.35])
H(d, "Live weather on the fires and what it means for spotting")
weather_table(d, ["Camp", "Pioneer", "LNU", "Carr"])
P(d, "Low humidity and gusts are what make spot fires, and they are also what ground the small drones. At Jarbo Gap relative humidity fell from 23% at 06:00 to 11% by 14:00 "
  "with gusts of 21 to 40 mph; at LNU Hawkeye RAWS held 6 to 13% overnight. The search must therefore be planned for the hours the Type 3 tier cannot fly, which points M2 to the Type 1 tier if its wind limit is 30 mph or higher.")
P(d, ("What heat sensors tell us. ", "b"), "Thermal sees flame and hot ground through smoke when visibility defeats RGB and lookouts; fused RGB-IR is the FLAME 2 benchmark. "
  "The Mavic 2 EA thermal range is -40 to 550 °C in low gain, so it locates flame but saturates before measuring intensity. "
  "Smoke detection suits early detection and flame detection suits larger fires (Keerthinathan et al. 2023, Week 6 M2 Q5).")
H(d, "How the uncertainty is quantified")
P(d, ("Time to detect: ", "b"), "with a search pass every R minutes, a per-pass chance the drone flies p (wind gate) and a per-pass chance the sensor sees the spot Pd, the spot is seen on pass G ~ Geometric(q), q = p x Pd. "
  "Expected detection time ", ("E[T_detect] = R/2 + R (1 - q)/q.", "b"), " Smallest detectable spot: ", ("A_min = (k x GSD)^2, GSD = 2 h tan(HFOV/2) / N", "b"),
  " for an N-pixel array at height h. Example (assumed 40° HFOV, k = 3 pixels): a 640-pixel array at 60 m AGL gives GSD "
  f"{gsd(60) * 100:.1f} cm and A_min about {(3 * gsd(60)) ** 2:.2f} m²; at 3,500 ft AGL (Type 1 floor) GSD {gsd(1066.8):.2f} m and A_min about {(3 * gsd(1066.8)) ** 2:.0f} m².")
rows = []
for name, lab in [(AX, "Alta X, 22 mph"), (SB, "SwitchBlade-Elite, 25 mph"), (T30, "Type 1 FW if 30 mph"), (T40, "Type 1 FW if 40 mph")]:
    ps, pg = g("Camp", name, gate="p_fly_sustained"), g("Camp", name)
    rows.append([lab] + [f"{tdet_min(R2, ps, Pd):.0f} to {tdet_min(R2, pg, Pd):.0f}" for Pd in (0.5, 0.8, 0.95)])
T(d, [f"Camp day 1, R = {R2} min", "Pd 50%: minutes", "Pd 80%: minutes", "Pd 95%: minutes"], rows, [2.2, 1.5, 1.5, 1.5])
P(d, ("Kinds of uncertainty. ", "b"), "p: aleatory, from the station gusts. Pd: epistemic until flight-tested in smoke (Week 6 M2 Q1). "
  "HFOV and k: assumed, to be replaced by the chosen payload's specification. Ranges in the table run from the sustained-wind to the gust reading.")
P(d, ("v1.1 needs: ", "b"), "Pd and false alarm rate from flight tests by smoke class; search area per fire (Week 6 Q2) to set R from fleet size; the Type 1 wind limit.", after=2)
P(d, (SRC, "i"))
d.save(OUT / "M2_Spot_Fire_Detection_UQ_v1.0.docx")

# ============================================================ M3 Firefighter safety
d = new_doc()
header_block(d, "M3", "Firefighter Safety", "First version of the M3 uncertainty calculation; builds on the safety margin in UQ proposal v0.5.")
random.seed(20261020)
N = 20000


def geom(q):
    n = 1
    while random.random() > q:
        n += 1
        if n > 10000: break
    return n


def p_margin_neg(T_fire, p_fly, egress, R=15, Pd=0.9, s=0.52, retry=2):
    neg = 0
    for _ in range(N):
        q = p_fly * Pd
        td = random.uniform(0, R) + (R * (geom(q) - 1) if q > 0 else 1e9)
        tn = retry * geom(s)
        neg += (T_fire - (td + tn + egress)) < 0
    return neg / N


H(d, "Bottom line")
cases = [("Camp first burnovers, Alta X, sustained", 85, g("Camp", AX, gate="p_fly_sustained")),
         ("Camp first burnovers, Alta X, gusts", 85, g("Camp", AX)),
         ("Camp first burnovers, Type 1 if 30 mph, gusts", 85, g("Camp", T30)),
         ("Carr 26 July, any drone listed, gusts", 62, g("Carr", AX))]
res = {c[0]: [p_margin_neg(c[1], c[2], e) for e in (30, 45, 60)] for c in cases}
B(d, "P(M < 0), the chance the fire reaches the crew before they are out, ", f"for a 45-minute escape at Camp's first burnovers is {pct(res[cases[0][0]][1])} with an Alta X if gusts are ignored and {pct(res[cases[1][0]][1])} if they are not.")
B(d, "Carr: flyable, still at risk. ", f"Mule Mountain gusts stayed at or below 21 mph before the 26 July fire whirl, so every listed drone could fly. P(M < 0) for a 45-minute escape is still {pct(res[cases[3][0]][1])}: about half from a missed detection pass (Pd 0.9), half from relay retries and egress.")
B(d, "Egress time is the missing measurement. ", "Crew positions and escape routes at the burnovers are not in project records, so results are shown for 30, 45 and 60 minutes.")
H(d, "Who uses it")
P(d, ("Crew level, lookouts and crew bosses: ", "b"), "this is their number; LCES depends on a warning arriving with more time than the escape takes. ",
  ("Division supervisors and the safety officer: ", "b"), "they set trigger points from it. ", ("Incident commander: ", "b"), "decides whether a crew is committed in the conditions.")
H(d, "Drone types serving M3 and their parameters")
T(d, ["Drone", "Type", "Parameters", "Wind limit", "M3 role"], [
    ["FVR-90", "NWCG Type 1 FW", "12 to 18 h; thermal video; flew as airborne relay in TCL 1", "Not found", "Overwatch of crews and escape routes; radio relay"],
    ["SuperVolo", "VTOL FW research", "8 h; DoodleLabs relay at 1,200 to 1,500 ft MSL", "Not found", "Alternate relay behind terrain"],
    ["Alta X", "NWCG Type 3 rotor", "17 min average sortie; under 2,000 ft AGL", "22 mph (secondary)", "Close checks of escape routes and safety zones"],
    ["DJI Matrice 200 v2 (research)", "Type 3 / 4 band", "Sonic anemometer profiles near the fire front", "Not in library", "Measures local wind for trigger points"],
], [1.35, 1.1, 2.4, 0.95, 1.5])
H(d, "Live weather during the crew hazards")
weather_table(d, ["Camp", "Carr", "LNU"])
P(d, "Camp: in its 07:00 hourly report, before the 07:50 burnovers, Jarbo Gap read 14 mph sustained, 28 mph gusts and 21% humidity; at the Rattlesnake Flats burnover hour (15:00) it read 10 mph, 28 mph gusts, 12%. "
  "Carr: at 18:00 on 26 July, before the 18:38 aerial warning, Mule Mountain read 3 mph sustained, 21 mph gusts, 44 °C and 7% humidity: light wind at this station during a wind-driven run with a fire whirl (v0.5).")
P(d, ("Crew and fireline tracking. ", "b"), "Avenza Maps GPS tracks give each firefighter's position hour by hour; set against the drone fireline they give distance to the edge and real T_egress. "
  "Where smoke compromises visibility and escape routes, the drone's thermal view stands in for the lookout and marks no-work zones. "
  "Firing operations at the start of a shift need the same gate: at Camp's Rattlesnake Flats a firing operation was planned for 14:15, the wind shifted at 14:45 and the burnover came at 15:15.")
H(d, "How the uncertainty is quantified")
P(d, ("Safety margin: ", "b"), ("M = T_fire - (T_detect + T_notify + T_egress),", "b"), " headline P(M < 0). T_detect: pass every R = 15 min (assumed), flies with wind gate p, sees the hazard with Pd = 0.9 (assumed). "
  "T_notify: retries every 2 min (assumed), each delivered with s = 52%, the TCL 1 telemetry track sharing rate with four ground nodes (Wu et al 2025, pp.9-11), used as a stand-in for alert delivery. "
  "T_fire: 85 min at Camp (first 911 call 06:25 to first burnovers 07:50) and 62 min at Carr (aerial warning 18:38 to mayday 19:40). Monte Carlo, 20,000 draws per cell.")
T(d, ["P(M < 0) by escape time", "30 min", "45 min", "60 min"],
  [[c[0]] + [pct(x) for x in res[c[0]]] for c in cases], [3.2, 1.1, 1.1, 1.1])
P(d, ("Kinds of uncertainty. ", "b"), "p: aleatory (gusts). Pd, R, retry: epistemic, assumed for v1.0. s: measured once in TCL 1. T_egress: epistemic and the largest; it needs crew positions and routes. "
  "T_fire is treated as fixed per case; v1.1 samples it from the spread model ensemble.")
P(d, ("v1.1 needs: ", "b"), "Avenza crew tracks or other crew positions and escape routes at the Camp and Carr burnovers; radio and dispatch logs for real T_notify; a plume-turbulence gate (Pioneer); a measured alert delivery rate (TCL 1 telemetry was over 90% with three ground nodes).", after=2)
P(d, (SRC, "i"))
d.save(OUT / "M3_Firefighter_Safety_UQ_v1.0.docx")

# ============================================================ M4 Water drop suppression
d = new_doc()
header_block(d, "M4", "Water Drop Suppression", "First version of the M4 uncertainty calculation; mission remains a GAP in the library.")
G0 = 9.81; HREL = 30.0
tfall = math.sqrt(2 * HREL / G0)


def drift_m(mph):
    return mph / 2.236936 * tfall


camp = WX["Camp"]["JBGC1"]
HAWK = 34.5  # 30 kt
hawk_p = None
H(d, "Bottom line")
B(d, "No drone in the project library has dropped water. ", "Every M4 candidate is a test or demonstrator (W5), so M4 v1.0 quantifies what the weather would have allowed, not what was done.")
B(d, f"{drift_m(camp['crit_sust_max_mph']):.0f} to {drift_m(camp['crit_gust_max_mph']):.0f} m of drift. ", f"A release from {HREL:.0f} m at Camp's first-day winds (sustained {camp['crit_sust_max_mph']:.0f} mph, gusts {camp['crit_gust_max_mph']:.0f} mph) moves the load that far downwind before any aiming correction.")
B(d, "Throughput is the other wall. ", "Week 6 records about 35 L/h for a small drone against 10,345 L/h for a large airtanker; swarms of 20 to 100 gal aircraft close part of that only if many fly at once.")
H(d, "Who uses it")
P(d, ("Incident commander, operations section chief and air tactical group supervisor: ", "b"), "they order and sequence drops and need the chance a drop lands on target when crewed aircraft are grounded. ",
  ("Crews on the line: ", "b"), "they work under the drop, so drift and timing are also a safety number for them.")
H(d, "Drone types for M4 and their parameters (W5 unless noted)")
T(d, ["Aircraft", "Status", "Load", "Wind", "M4 role"], [
    ["Rain and Sikorsky autonomous Black Hawk / Firehawk", "Demonstrator; CAL FIRE agreement 2025", "324 gal bucket (demo); 1,000 gal tank class", "Flew hands-off in gusts to 30 kt", "Direct attack at helicopter scale"],
    ["Seneca Argo-1 swarm (5 drones)", "Tested by CAL FIRE, July 2026", "20 gal water plus foam each", "Not found; IR targeting", "Initial attack on small fires"],
    ["FireSwarm Thunder Wasp GT", "Tested (XPRIZE finalist)", "Up to 100 gal", "Not found", "Initial attack; bulk water"],
    ["Large UAS or optionally piloted aircraft (Joby class)", "Concept (ConOps p.19)", "Bulk water", "Not found", "Water drop instead of retardant"],
], [1.6, 1.25, 1.35, 1.6, 1.25])
H(d, "Live weather on the fires")


def hawk(f):
    """Firehawk 30 kt gate, gusts"""
    import csv
    rows = [r for r in csv.DictReader(open(HERE / "weather_eval" / "critical_window_hourly.csv")) if r["fire"] == f and r["station"] == REP[f] and r["gust"]]
    return pct(sum(float(r["gust"]) <= HAWK for r in rows) / len(rows))


weather_table(d, ["Camp", "Pioneer", "LNU", "Carr"], hawk)
P(d, "Camp's critical day is the hard case: the Alta X class flies in 1 of 13 hours by gusts, while a 30 kt (34.5 mph) aircraft flies in all but the 06:00 hour. "
  "Carr and Pioneer stayed within the 30 kt Firehawk figure; wind limits for the other listed aircraft were not found. Pioneer's run had pyroconvective updrafts, which a surface wind gate does not see.")
H(d, "How the uncertainty is quantified")
P(d, ("Effective water on target per hour: ", "b"), ("V_eff = n x S x V x P(hit) x p,", "b"), " with n aircraft, S sorties per hour = 60 / (2 transit + drop + turnaround), V load per sortie, p the wind gate. "
  "P(hit) is the chance the drop centroid lands within the target half-width w: with drift d = u x t_fall, t_fall = sqrt(2h/g) = "
  f"{tfall:.2f} s at h = {HREL:.0f} m (assumed release height), and aiming error σ, P(hit) = Φ((w - d_residual)/σ) - Φ((-w - d_residual)/σ), where d_residual is the drift not corrected by the aim. "
  "Wind uncertainty enters twice: through p and through d.")
rows = []
for f in ["Camp", "LNU", "Carr", "Pioneer"]:
    r = WX[f][REP[f]]
    rows.append([f, f"{r['crit_sust_max_mph']:.0f} / {r['crit_gust_max_mph']:.0f}", f"{drift_m(r['crit_sust_max_mph']):.0f} to {drift_m(r['crit_gust_max_mph']):.0f}",
                 f"{drift_m(r['crit_gust_max_mph']) - drift_m(r['crit_sust_max_mph']):.0f}", hawk(f)])
T(d, ["Fire", "Sust. / gust mph", "Drift m (h = 30 m)", "Gust spread m (uncorrectable)", "p, 30 kt gate"], rows, [1.0, 1.4, 1.5, 1.9, 1.3])
P(d, "The gust spread is the part of the drift an aim made on sustained wind cannot remove. This is a no-drag, instant-wind estimate from 20 ft winds; drag and stronger flight-level wind push it in opposite directions, so the sign of the bias is unknown.")
P(d, ("v1.1 needs: ", "b"), "measured drop accuracy against wind from the Seneca and Firehawk tests; release heights and loads per aircraft; target width by fuel type and fuel loading (LANDFIRE).", after=2)
P(d, (SRC, "i"))
d.save(OUT / "M4_Water_Drop_Suppression_UQ_v1.0.docx")

# ============================================================ M5 Logistics support
d = new_doc()
header_block(d, "M5", "Logistics Support", "First version of the M5 uncertainty calculation; questions as on the Week 6 M5 tab.")
V = 50.0; END, RES = 22.0, 5.0; HOLD, TURN = 2.0, 10.0; DIST = 3.0


def radius_km(w_kmh):
    if w_kmh >= V: return 0.0
    tt = (END - RES) / 60
    return tt / (1 / (V - w_kmh) + 1 / (V + w_kmh))


def cycle_min(dkm, w_kmh):
    if w_kmh >= V: return float("inf")
    return 60 * (dkm / (V - w_kmh) + dkm / (V + w_kmh)) + HOLD + TURN


H(d, "Bottom line")
sus = camp["crit_sust_max_mph"] * 1.609; gus = camp["crit_gust_max_mph"] * 1.609
B(d, f"{radius_km(0):.1f} km becomes {radius_km(sus):.1f} km, then zero. ", f"Loaded radius of an Alta X carrying 22 lb (10 kg, about 10 L of drinking water): {radius_km(0):.1f} km in still air, "
  f"{radius_km(sus):.1f} km into Camp's {camp['crit_sust_max_mph']:.0f} mph sustained headwind, and nothing into its {camp['crit_gust_max_mph']:.0f} mph gusts (above the assumed {V:.0f} km/h cruise; the airframe top speed is 60 mph, and its 22 mph limit would ground it first).")
import csv as _csv
HOURLY = list(_csv.DictReader(open(HERE / "weather_eval" / "critical_window_hourly.csv")))


def p_on_time(f, field, window=30.0, limit=22.0):
    """Share of critical-window hours a 3 km delivery is flown (wind at or under the limit) and back inside the window."""
    hrs = [r for r in HOURLY if r["fire"] == f and r["station"] == REP[f] and r[field]]
    ok = [float(r[field]) <= limit and cycle_min(DIST, float(r[field]) * 1.609) <= window for r in hrs]
    return sum(ok) / len(ok)


pon = {f: (p_on_time(f, "sust"), p_on_time(f, "gust")) for f in ["Camp", "LNU", "Pioneer", "Carr"]}
B(d, "P(on time) on Camp's first day: ", f"{pct(pon['Camp'][1])} to {pct(pon['Camp'][0])} for a 3 km delivery inside a 30-minute window; across the whole Camp perimeter period, gusts ground the Alta X in {pct(1 - g('Camp', AX, 'perim'))} of hours.")
B(d, "Small drones carry litres, crews need loads. ", f"At {60 / cycle_min(DIST, 0):.1f} deliveries per hour one Alta X moves about {10 * 60 / cycle_min(DIST, 0):.0f} kg/h; heavy-lift aircraft in development (Elroy Chaparral, 500 lb over 300 nm; Piasecki KARGO II, 1,500 lb) are the M5 class.")
H(d, "Who uses it")
P(d, ("Crew bosses and division supervisors: ", "b"), "whether water, hose, pumps and food arrive when the crew needs them, especially far from roads (Pioneer) or when roads are jammed by evacuation (Camp). ",
  ("Logistics section chief and incident commander: ", "b"), "how many aircraft to order and which crews to supply by air.")
H(d, "Drone types serving M5 and their parameters")
T(d, ["Aircraft", "Status", "Payload and endurance", "Wind limit", "M5 role"], [
    ["Alta X (Freefly)", "In service, flown with ACERO", "35 lb max; 50 min empty, 35 min at 8 lb, 22 min at 22 lb; 60 mph top speed", "22 mph (secondary)", "Light resupply under 10 kg"],
    ["Type 3 / 4 agency small UAS", "In service", "Light loads; Type 3 range 5 mi, Type 4 under 2 mi", "Not published", "Radios, batteries, first aid"],
    ["Elroy Air Chaparral", "In development (W5)", "500 lb over 300 nm", "Not found", "Crew resupply from staging"],
    ["Piasecki KARGO II", "In development (W5)", "1,500 lb", "Not found", "Water, hose and pumps in bulk"],
], [1.4, 1.3, 2.35, 0.95, 1.3])
H(d, "Live weather on the fires")
weather_table(d, ["Camp", "LNU", "Pioneer", "Carr"])
P(d, "Daytime heat also cuts payload: Carr reached 44 °C at Mule Mountain and LNU 38 °C in the afternoon at Atlas Peak, which raises density altitude and shortens loaded endurance (Week 6 M5 Q2 derate).")
H(d, "How the uncertainty is quantified")
P(d, ("Probability a requested delivery arrives in its window: ", "b"), ("P(on time) = p x P(complete | fly) x P(cycle <= window).", "b"),
  " Loaded radius into a headwind w at airspeed V: r = t_usable / (1/(V - w) + 1/(V + w)), t_usable = endurance loaded - reserve. "
  "Cycle = round trip + hold + turnaround (Week 6 M5 Q4). Inputs: endurance 22 min at 22 lb (W5), reserve 5 min, V = 50 km/h, hold 2 min, turnaround 10 min, crew at 3 km (all four assumed). "
  "P(complete | fly) = 1 until a delivery success rate is measured.")
rows = []
for f in ["Camp", "LNU", "Pioneer", "Carr"]:
    r = WX[f][REP[f]]; ws, wg = r["crit_sust_max_mph"] * 1.609, r["crit_gust_max_mph"] * 1.609
    cs, cg = cycle_min(DIST, ws), cycle_min(DIST, wg)
    rows.append([f, f"{radius_km(ws):.1f} / {radius_km(wg):.1f}", f"{cs:.0f} / {'none' if math.isinf(cg) else f'{cg:.0f}'}",
                 f"{pct(pon[f][1])} to {pct(pon[f][0])}", pct(g(f, AX, "perim"))])
T(d, ["Fire, Alta X 10 kg", "Radius km at peak sust. / gust", "Cycle min at peak sust. / gust", "P(on time), share of day-1 hours", "p gust, whole fire"], rows, [1.5, 1.6, 1.6, 1.4, 1.3])
P(d, ("Kinds of uncertainty. ", "b"), "p and w: aleatory, from station wind. Airspeed, turnaround, success rate: epistemic, assumed for v1.0. Crew demand (Week 6 M5 Q4) is not yet in project records, so fleet size is not computed.")
P(d, ("v1.1 needs: ", "b"), "crew demand per shift (water, hose, pumps); measured Alta X loaded cruise speed and wind limit; distances from staging to crews on Camp and Pioneer; heavy-lift specifications.", after=2)
P(d, (SRC, "i"))
d.save(OUT / "M5_Logistics_Support_UQ_v1.0.docx")

json.dump({"M3": res}, open(OUT / "mission_uq_v1_numbers.json", "w"), indent=1)
print("wrote", sorted(p.name for p in OUT.iterdir()))
