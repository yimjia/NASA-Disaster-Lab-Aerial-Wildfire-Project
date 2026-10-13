"""Builds UQ proposal v0.2 (docx) and evaluation workbook v0.2 (xlsx).

v0.2 adds the Fire Case Studies evaluation (Pioneer 2016, Carr 2018, LNU 2020) to the
v0.1 framework and the research data gathered on 2026-09-30. Literature claims keep the
[Doc Year, p.N] citations from ~/nasa_rag used in v0.1. Case study numbers come from
case_study_eval/ (run case_study_evaluation.py first) and name their data source.
v0.1 files are not modified.

Run: uv run --project ~/nasa_rag python build_uq_v02.py
"""
import json, csv
from pathlib import Path
from docx import Document
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

OUT = Path(__file__).parent
EV = OUT / "case_study_eval"
DATE = "6 October 2026"
R = json.load(open(EV / "results.json"))
def rows(name): return list(csv.DictReader(open(EV / f"{name}.csv")))
P16, C18, L20 = R["Pioneer 2016"], R["Carr 2018"], R["LNU 2020"]
f0 = lambda x: f"{x:,.0f}"
f1 = lambda x: f"{x:,.1f}"

# derived figures used in text and sheets
lag = {k: R[k]["nicc_lag_test"] for k in R}
chain = {k: R[k].get("latency_chain_geomac_to_209") for k in R}
res = {k: R[k]["resources_nicc"] for k in R}
grow = {k: R[k]["growth"] for k in R}
short_carr = {s["date"]: s for s in C18["blowup_day_shortfall"]}
lnu_day = {r["date"]: r for r in rows("size_by_day") if r["fire"] == "LNU 2020"}
lnu_819 = (float(lnu_day["2020-08-19"]["ics209_last"]), float(lnu_day["2020-08-19"]["feds_last"]))
lnu_feds_stable = float(lnu_day["2020-09-16"]["feds_last"])
res_rows = rows("resources")
def peak_growth_personnel(fire):
    top = [t["date"] for t in grow[fire]["top3_days"]]
    out = []
    for d in top:
        from datetime import date, timedelta
        nd = str(date.fromisoformat(d) + timedelta(days=1))
        m = [r for r in res_rows if r["fire"] == fire and r["imsr_date"] == nd]
        if m and m[0]["personnel"]: out.append(f"{d[5:]}: {int(float(m[0]['personnel'])):,}")
    return out
CS_REF = "Fire Case Studies workbooks (ICS-209-PLUS, NICC IMSR), case_study_eval/results.json"

# ======================================================================== document
d = Document()
st = d.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(11)
def H(t, l=1): d.add_heading(t, level=l)
def P(t, italic=False, bold=False):
    p = d.add_paragraph(); r = p.add_run(t); r.italic = italic; r.bold = bold; return p
def B(items):
    for t in items: d.add_paragraph(t, style="List Bullet")
def N(items):
    for t in items: d.add_paragraph(t, style="List Number")
def T(header, rws):
    t = d.add_table(rows=1, cols=len(header)); t.style = "Light Grid Accent 1"
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]; c.text = ""; c.paragraphs[0].add_run(h).bold = True
    for row in rws:
        cells = t.add_row().cells
        for i, v in enumerate(row): cells[i].text = str(v)
    for row in t.rows:
        for c in row.cells:
            for p in c.paragraphs:
                for r in p.runs: r.font.size = Pt(9)
    d.add_paragraph()

tp = d.add_paragraph(); tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = tp.add_run("Quantifying Uncertainty for Firefighter Safety and Aerial Logistics\nin 24/7 Wildland Fire Aerial Operations"); r.bold = True; r.font.size = Pt(18)
sp = d.add_paragraph(); sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
sp.add_run("Evaluation proposal with fire case study evidence\n").italic = True
sp.add_run(f"NASA Aerial Wildfire Project, ACERO Disaster Lab working group\nDraft v0.2, {DATE}")

P("Status of this draft. Preliminary thinking for team discussion, not a submission. v0.2 keeps the v0.1 "
  "framework (29 September 2026) and adds an evaluation of three fire case studies built from operational "
  "records: Pioneer 2016, Carr 2018 and the LNU Lightning Complex 2020. Literature claims are cited as "
  "[Document Year, p.N] from the project reference library. Case study numbers come from the project's own "
  "analysis of public records and name the record they come from. Methods marked \"general practice\" are "
  "standard uncertainty quantification methodology, not drawn from the library. This draft makes no claim "
  "that the approach or the findings are new; a dated literature search is required before any novelty framing.",
  italic=True)

H("What changed from v0.1")
B([
 "New Section 4: an evaluation of the Pioneer, Carr and LNU case studies across four records of fire size "
 "(ICS-209 reports, NICC daily situation reports, NIFC GeoMAC perimeter histories, NASA FEDS satellite perimeters).",
 "Perimeter age (U5) now has an empirical prior instead of a literature range, and five new uncertain inputs "
 "(U12 to U16) come from the case studies: reporting lag, blowup-day size shortfall, satellite area bias, "
 "operational revisions, and logistics demand scale.",
 "A data quality register: errors and definitional traps in the operational archives that would bias a "
 "validation if not handled.",
 "The experiment design now uses four fires with distinct roles, and the surveillance cadence levels are "
 "anchored to observed cadences.",
 "Carr 2018 joins Camp as a validation case for the crew safety margin, since it has timed aerial warning "
 "and entrapment records.",
])

H("Abstract")
P("NASA's ACERO ConOps sets the goal of 24/7 aerial operations for wildland fire, organised in this project around "
  "surveillance (M1 fireline mapping, M2 spot fire detection, M3 firefighter safety) and suppression (M4 water drop, "
  "M5 water delivery). v0.1 proposed turning the uncertainty in those missions into two decision numbers: the "
  "probability that a crew's safety margin goes negative, P(M<0), and the probability that a delivery arrives on "
  "time. Its inputs came from the literature, for example a 47% mean absolute percentage error in operational rate "
  "of spread predictions [Cardil et al 2023, pp.6-7] and perimeters that can be hours old when a forecast starts "
  "[Kochanski et al 2023, p.12]. This version tests those inputs against what operations actually recorded on three "
  f"large fires. The official fire size was about one day behind the mapped perimeter: the ICS-209 size matched the "
  f"previous night's perimeter on {chain['Pioneer 2016']['match_next_day_209']} of {chain['Pioneer 2016']['days_tested']} "
  f"changing days on Pioneer and {chain['Carr 2018']['match_next_day_209']} of {chain['Carr 2018']['days_tested']} on Carr, "
  f"and the national situation report repeated the previous evening's 209 on {lag['Pioneer 2016']['exact_match_prev_day_209']} "
  f"of {lag['Pioneer 2016']['days_compared']}, {lag['Carr 2018']['exact_match_prev_day_209']} of {lag['Carr 2018']['days_compared']} "
  f"and {lag['LNU 2020']['exact_match_prev_day_209']} of {lag['LNU 2020']['days_compared']} days. On blowup days that lag "
  f"became a large shortfall: on 27 July 2018 the Carr 209 gave {f0(short_carr['2018-07-27']['ics209'])} acres while the "
  f"same-day perimeter showed {f0(short_carr['2018-07-27']['geomac'])} and the satellite {f0(short_carr['2018-07-27']['feds'])}. "
  f"Growth was concentrated: the three largest days held {grow['Pioneer 2016']['top3_share_of_final']:.0%}, "
  f"{grow['Carr 2018']['top3_share_of_final']:.0%} and {grow['LNU 2020']['top3_share_of_final']:.0%} of final area. These "
  "findings set the baseline that 24/7 aerial surveillance would have to beat (RQ4), give empirical priors for perimeter "
  "age and observation error, size the logistics demand for M5 from recorded crew and aircraft counts, and flag archive "
  "errors that would corrupt a validation. Camp 2018 and Carr 2018 are proposed as the validation cases for P(M<0).")

H("1. Motivation")
H("1.1 Why crew safety needs a number, not a narrative", 2)
P("The wildland fire airspace \"has repeatedly proven to be among the most hazardous airspace environments in government "
  "aviation operations\" [NASA ACERO ConOps 2026, p.12], and the ConOps names improving the safety of current crewed "
  "operations as \"the top priority\" of the future state [NASA ACERO ConOps 2026, p.16]. NWCG frames UAS safety as "
  "risk, set by \"the severity of possible events and the probability that they will occur\" [NWCG 2025, pp.20-21]. A "
  "24/7 aerial capability changes the probability term for ground crews, and the team needs a way to measure by how "
  "much, with honest error bars.")
H("1.2 Why logistics needs a number", 2)
P("The ConOps describes drone-enabled BVLOS logistics (food, water, radios, tools, hose, fuel) delivered to crews where "
  "crewed flight is not acceptable, with crews able to move the rendezvous point mid-flight [NASA ACERO ConOps 2026, p.19]. "
  "Whether that concept is useful depends on reliability, not capability: how often a delivery arrives on time, given "
  "wind, endurance and a network that is not yet robust [Wu et al 2025, p.11].")
H("1.3 What the case studies add", 2)
P("v0.1 drew its priors from the literature. The case studies show what decision makers actually had in hand on three "
  "large fires: how often the fire was measured, how old the public number was, how far it fell behind on the worst "
  "days, and how many people and aircraft were on the incident. Those are the quantities that 24/7 surveillance and "
  "aerial logistics would change.")

H("2. Research questions")
N([
 "RQ1 (M3, crew safety). For a crew at a given position and time, what is the probability that fire arrival precedes "
 "the crew's completed egress to a safety zone, and how does that probability change with and without 24/7 aerial "
 "surveillance (M1, M2)?",
 "RQ2 (M5, logistics). What is the probability that a requested delivery arrives within its required window, and what "
 "is the distribution of sortie turnaround time, under day, night and smoke (DVE) conditions, at the crew and "
 "helicopter counts recorded on real incidents?",
 "RQ3 (drivers). Which input uncertainties dominate the variance of the RQ1 and RQ2 outputs?",
 "RQ4 (value of information). How much does fresher aerial observation shrink the uncertainty in RQ1 and RQ2, measured "
 "against the observed operational baseline of Section 4 (about one day of reporting lag, perimeters 17 to 19 hours "
 "apart at the median)?",
 "RQ5 (observation error, new). How large and how biased are the operational records of fire size and growth, and how "
 "should they enter calibration and validation so that model error is not confused with reporting error?",
])

H("3. Where the uncertainty comes from")
P("Aleatory uncertainty (true variability) is separated from epistemic uncertainty (lack of knowledge) because RQ4 "
  "only applies to the epistemic part (general practice). Rows marked CS are new in v0.2 and come from Section 4.")
T(["Source", "Type", "Evidence", "Citation or record"], [
 ["Fire spread model structure", "Epistemic", "Operational ROS MAPE 47%, best in shrub (26%); timber underpredicts", "[Cardil et al 2023, pp.6-7]"],
 ["Fuel type mapping", "Epistemic", "Key contributor to ROS bias; wind secondary; moisture smaller", "[Cardil et al 2023, p.6]"],
 ["Correlated input errors", "Epistemic", "Weather error is spatially correlated; use multivariate perturbations", "[Pyregence Consortium n.d. GridFire, p.113]"],
 ["Initial fire state (perimeter)", "Epistemic", "Spin-up initialisation cut growth error 3 to 35 times; VIIRS latency up to 4 h", "[Kochanski et al 2023, p.12]"],
 ["Communication network", "Aleatory in operation", "PAMS field data: TFR 100%, operational intent 87%, telemetry 52%", "[Wu et al 2025, pp.9-10]"],
 ["Benefit assumptions", "Epistemic", "Expert opinion varied widely; PERT in Monte Carlo", "[Hope et al 2024, pp.7-8]"],
 ["CS: perimeter cadence", "Epistemic", f"GeoMAC perimeters a median {P16['geomac_interval_h']['median']:.1f} h apart (Pioneer), {C18['geomac_interval_h']['median']:.1f} h (Carr); {P16['geomac_share_interval_over_24h']:.0%} and {C18['geomac_share_interval_over_24h']:.0%} of gaps over 24 h", "NIFC GeoMAC perimeter histories"],
 ["CS: reporting lag", "Epistemic", "209 size equals previous night's perimeter; NICC repeats previous evening's 209", CS_REF],
 ["CS: blowup-day shortfall", "Epistemic", "Official size 1.5 to 1.7 times below same-day perimeter on Carr's worst days", CS_REF],
 ["CS: satellite area bias", "Epistemic", "FEDS (375 m VIIRS) final area above the official perimeter on Carr", "NASA FEDS 12 h perimeters"],
 ["CS: archive cleaning", "Epistemic", "Cleaned fields replace as-reported values (size, containment date)", "ICS-209-PLUS"],
])

H("4. Evaluation of the fire case studies")
H("4.1 Data", 2)
P("Each case study combines four records of fire size and one of resources. ICS-209 incident status summaries come "
  "from the ICS-209-PLUS archive (St. Denis et al. 2023). NICC Incident Management Situation Reports are the national "
  "daily summaries, recovered from Internet Archive copies of the NWCG archive. NIFC GeoMAC perimeter histories and "
  "NASA FEDS 12-hourly VIIRS perimeters were gathered for the UQ research data on 30 September 2026. Final perimeters "
  "come from NIFC (Pioneer, Carr) and CAL FIRE FRAP (LNU, which NIFC's history service does not cover for 2020).")
T(["Fire", "ICS-209 reports", "NICC reports listing it", "GeoMAC perimeters", "FEDS overpasses", "Final acres"], [
 ["Pioneer 2016 (Boise NF, ID)", P16["209_reports"], P16["nicc_reports_with_row"], P16.get("geomac_perimeters", "none"), "none (archive starts 2018)", f0(grow["Pioneer 2016"]["final_acres_209"])],
 ["Carr 2018 (Shasta / Whiskeytown, CA)", C18["209_reports"], C18["nicc_reports_with_row"], C18.get("geomac_perimeters"), C18.get("feds_overpasses"), f0(grow["Carr 2018"]["final_acres_209"])],
 ["LNU Lightning Complex 2020 (CA)", L20["209_reports"], L20["nicc_reports_with_row"], "none", L20.get("feds_overpasses"), f0(grow["LNU 2020"]["final_acres_209"])],
])
P("Method. For each calendar day the last ICS-209 size is the reference. The NICC report published the next morning, "
  "the last GeoMAC perimeter of the day and the last FEDS overpass of the day (local time) are compared with it. "
  "Sizes use the as-reported ICS-209 field (CURR_INCIDENT_AREA), because the cleaned ACRES field in ICS-209-PLUS is "
  "capped at the final size (Section 4.6). Reproduce with case_study_evaluation.py.")

H("4.2 The official fire size is about one day old", 2)
T(["Fire", "209 matches previous night's perimeter", "209 matches same-day perimeter", "NICC matches previous evening's 209", "NICC matches same-day 209"], [
 [k, f"{chain[k]['match_next_day_209']} of {chain[k]['days_tested']} days" if chain[k] else "no GeoMAC", f"{chain[k]['match_same_day_209']} of {chain[k]['days_tested']}" if chain[k] else "n/a",
  f"{lag[k]['exact_match_prev_day_209']} of {lag[k]['days_compared']} days", f"{lag[k]['exact_match_same_day_209']} of {lag[k]['days_compared']}"] for k in R])
P("Days counted are those where the 209 size changed; a match is within 0.5% (perimeters) or exact (NICC). The "
  "reading is a reporting chain: an overnight infrared perimeter becomes the next day's 209 size, and the 209 becomes "
  "the next morning's national number. Two consequences follow. First, the national situation report is not an "
  "independent measurement; agreement between NICC and the 209 says nothing about accuracy. Second, the size a "
  "decision maker reads is commonly one to two days behind the fire, which is the baseline RQ4 has to improve on.")

H("4.3 On blowup days the lag becomes a large shortfall", 2)
T(["Fire, date", "ICS-209 acres", "Same-day perimeter", "Same-day satellite (FEDS)", "Perimeter / 209", "Satellite / 209"], [
 [f"Carr, {s['date']}", f0(s["ics209"]), f0(s["geomac"]) if s["geomac"] else "", f0(s["feds"]) if s["feds"] else "",
  f"{s['ratio_geomac']:.2f}" if s["ratio_geomac"] else "", f"{s['ratio_feds']:.2f}" if s["ratio_feds"] else ""] for s in sorted(C18["blowup_day_shortfall"], key=lambda s: s["date"])] +
 [[f"Pioneer, {s['date']}", f0(s["ics209"]), f0(s["geomac"]), "", f"{s['ratio_geomac']:.2f}", ""] for s in sorted(P16["blowup_day_shortfall"], key=lambda s: s["date"])] +
 [["LNU, 2020-08-19", f0(lnu_819[0]), "", f0(lnu_819[1]), "", f"{lnu_819[1]/lnu_819[0]:.2f}"]] +
 [[f"LNU, {s['date']}", f0(s["ics209"]), "", f0(s["feds"]), "", f"{s['ratio_feds']:.2f}"] for s in sorted(L20["blowup_day_shortfall"], key=lambda s: s["date"])])
P("The fatal Carr burnovers happened on the evening of 26 July 2018 (Research Data Index, Margin Timing Anchors: aerial "
  f"warning to Branch III at 18:38, mayday at 19:40). That day the 209 recorded {f0(short_carr['2018-07-26']['ics209'])} acres, "
  f"against {f0(short_carr['2018-07-26']['geomac'])} on the perimeter and {f0(short_carr['2018-07-26']['feds'])} from the satellite. "
  "The day crews most needed current information is the day the official number was furthest behind. Part of the "
  "satellite excess is pixel bias (Section 4.5), so the perimeter ratio is the more conservative measure of the shortfall.")

H("4.4 Growth is concentrated in a few days", 2)
T(["Fire", "Days with growth", "Median daily growth (acres)", "Largest day (acres)", "Top 3 days, share of final", "Days from first 209 to 50% / 90% of final", "NICC personnel on top 3 days (date: people)"], [
 [k, grow[k]["days_with_growth"], f0(grow[k]["daily_growth_acres"]["median"]), f0(grow[k]["daily_growth_acres"]["max"]),
  f"{grow[k]['top3_share_of_final']:.0%}", f"{grow[k]['days_to_50pct']} / {grow[k]['days_to_90pct']}",
  "; ".join(peak_growth_personnel(k))] for k in R])
P("The three fires span three regimes: a long timber fire with late runs (Pioneer), a wind-driven run early in a "
  "WUI fire (Carr), and a lightning complex that reached half its size in a day (LNU). Daily growth is heavy tailed, "
  "so an ensemble that matches the median day can still miss the days that matter for crew safety. The T_fire term "
  "of the safety margin should be evaluated separately on blowup days, and validation scores should be reported by "
  "regime, not pooled (proposed).")

H("4.5 How much do the records disagree about size?", 2)
T(["Record", "Pioneer", "Carr", "LNU"], [
 ["Official final sizes (209, NICC, agency perimeters)", "within 0.05%", "229,475 to 229,686 acres, within 0.1%", "362,921 (3 CAL FIRE perimeters) vs 363,220, within 0.1%"],
 ["Satellite (FEDS) after growth stopped", "n/a", f"{f0(C18['final_size_sources']['NASA FEDS last overpass'])}, about +{100*(C18['final_size_sources']['NASA FEDS last overpass']/C18['growth']['final_acres_209']-1):.0f}%", f"{f0(lnu_feds_stable)} on 16 Sep, about +{100*(lnu_feds_stable/L20['growth']['final_acres_209']-1):.1f}%; later overpasses drop to {f0(L20['final_size_sources']['NASA FEDS last overpass'])} (not explained, treat as artifact)"],
 ["Same-day perimeter vs 209, median (p90)", f"+{P16['size_diff_vs_209_pct']['GeoMAC']['signed']['median']:.1f}% (+{P16['size_diff_vs_209_pct']['GeoMAC']['signed']['p90']:.0f}%)", f"+{C18['size_diff_vs_209_pct']['GeoMAC']['signed']['median']:.1f}% (+{C18['size_diff_vs_209_pct']['GeoMAC']['signed']['p90']:.0f}%)", "no GeoMAC"],
 ["Operational downward revision", "none", "none", "375,209 to 363,220 (-3.2%) on 8 to 9 Sep, \"more accurate mapping\" (NICC narrative)"],
])
P("Final sizes agree closely once the fire is out. During the fire, the disagreement is mostly timing (the lag of 4.2) "
  "plus a satellite area bias that depends on fire and resolution. For model validation this means scoring against a "
  "time-matched perimeter, not against the 209 size of the same date, and treating FEDS area as biased high by an "
  "amount that has to be estimated per fire (proposed).")

H("4.6 Data quality register (summary)", 2)
B([
 "ICS-209-PLUS ACRES is cleaned and capped at the final size; on LNU 26 of 61 reports differ from the as-reported "
 "CURR_INCIDENT_AREA, hiding the 375,209 acre peak. Use the as-reported field for anything about real-time knowledge.",
 f"Personnel has two incompatible definitions: NICC counts all assigned people (peaks {f0(P16['personnel_definition_gap']['nicc_peak'])}, "
 f"{f0(C18['personnel_definition_gap']['nicc_peak'])}, {f0(L20['personnel_definition_gap']['nicc_peak'])}); the ICS-209-PLUS "
 f"derived field is {P16['personnel_definition_gap']['ratio']:.1f}, {C18['personnel_definition_gap']['ratio']:.1f} and "
 f"{L20['personnel_definition_gap']['ratio']:.1f} times smaller at peak.",
 "Cause disagrees across records: all three are U (undetermined or under investigation) in the 209s and NICC, while "
 "the Boise dispatch annual report gives Pioneer as human-caused and CAL FIRE codes all three LNU fires as lightning.",
 "Structures damaged totals track structures destroyed almost exactly on Carr and LNU while residential damaged is far "
 "lower; treat the damaged totals as suspect.",
 "Reporting gaps: the LNU 209 series jumps from 16 to 30 September; Pioneer has a 169 h gap in October. Transcription "
 "errors exist in NICC PDFs (personnel printed as 1.506 and 3.831).",
])

H("5. Quantities of interest")
H("5.1 Crew safety (M3)", 2)
B([
 "Safety margin M = T_fire - (T_detect + T_notify + T_egress) (proposed definition, unchanged from v0.1).",
 "Primary metric P(M < 0); secondary the 5th percentile of M and expected warning lead time.",
 "New: information age at decision, the age of the newest fire size or perimeter available to the crew boss, with "
 "the case study baseline of roughly one day for official size and a median 17 to 19 h between perimeters.",
])
H("5.2 Logistics (M5)", 2)
B([
 "P(on time), sortie turnaround p50 and p90, P(abort), link-loss minutes per sortie (unchanged).",
 f"New: demand scale from recorded resources. Median assigned personnel {f0(res['Pioneer 2016']['personnel']['median'])} (Pioneer), "
 f"{f0(res['Carr 2018']['personnel']['median'])} (Carr), {f0(res['LNU 2020']['personnel']['median'])} (LNU); median crews "
 f"{f0(res['Pioneer 2016']['crews']['median'])}, {f0(res['Carr 2018']['crews']['median'])}, {f0(res['LNU 2020']['crews']['median'])}; "
 f"median helicopters {f0(res['Pioneer 2016']['helicopters']['median'])}, {f0(res['Carr 2018']['helicopters']['median'])}, "
 f"{f0(res['LNU 2020']['helicopters']['median'])}. Crews per helicopter, a first proxy for how many crews one aerial "
 f"delivery asset would serve, had medians of {res['Pioneer 2016']['crews_per_helicopter']['median']:.1f}, "
 f"{res['Carr 2018']['crews_per_helicopter']['median']:.1f} and {res['LNU 2020']['crews_per_helicopter']['median']:.1f}.",
])
H("5.3 Suppression (M4)", 2)
P("Still out of scope: no drone in the library has flown water drops [Boroujeni et al 2024, p.38]. The NICC records "
  f"give helicopter-days on each incident ({res['Pioneer 2016']['helicopter_days']}, {res['Carr 2018']['helicopter_days']}, "
  f"{res['LNU 2020']['helicopter_days']}) as a scale for any future M4 comparison.")

H("6. Methods")
P("Sections 6.1 to 6.6 keep the v0.1 methods; changes are noted.")
H("6.1 Ensemble fire spread with correlated perturbations", 2)
P("Ensembles in ELMFIRE or GridFire with spatially correlated perturbations of fuel, wind and moisture "
  "[Pyregence Consortium n.d. GridFire, p.113]; ELMFIRE and GridFire suit extended attack, FARSITE short horizons "
  "[Altintas 2025, p.8]. Change: ensembles are scored separately on blowup days (Section 4.4).")
H("6.2 Empirical model-error and observation-error layers", 2)
P("v0.1 added a model residual by fuel group using Cardil's MAE, MBE and MAPE [Cardil et al 2023, pp.5-6], with the "
  "timber underprediction bias carried forward [Cardil et al 2023, p.7]. v0.2 adds a separate observation-error layer "
  "for the records used to score the model: a lag (one reporting cycle), a blowup-day shortfall, and a satellite area "
  "bias, each with the empirical ranges in the Uncertainty Register (U12 to U15). Keeping the two layers separate is "
  "what RQ5 asks for (proposed).")
H("6.3 Data assimilation as the lever for RQ4", 2)
P("Particle filters assimilate observations into spread models [Boroujeni et al 2024, p.46], and perimeter "
  "initialisation strongly affects forecast error [Kochanski et al 2023, p.12]. Change: the cadence levels are now "
  "the observed ones. Baseline A is the official record (one day of lag); baseline B is the perimeter record (median "
  "17 to 19 h, a quarter of gaps over a day on Pioneer); baseline C is satellite (12 h, biased area); the 24/7 aerial "
  "case is 30 minutes. The experiment measures how much the spread of T_fire shrinks from A to the aerial case.")
H("6.4 Stochastic operations simulation for logistics", 2)
P("Discrete-event sorties (dispatch, transit, rendezvous, drop, return, recharge) with random wind availability, link "
  "success from PAMS [Wu et al 2025, pp.9-10] and reroute requests from the ConOps logic [NASA ACERO ConOps 2026, p.19]. "
  "Change: the demand side is driven by the recorded daily crew counts from the NICC sheets, so the simulation runs "
  "against a real incident's ramp up and ramp down instead of a constant crew number.")
H("6.5 Expert elicitation", 2)
P("PERT distributions from practitioner minimum, most likely and maximum [Hope et al 2024, pp.7-8]. Still needed for "
  "notification latency, delivery window tolerance and human factor delays. Egress now has published travel-rate "
  "models in the research data (Sullivan 2020; Campbell 2017 and 2019), so elicitation for egress becomes a check.")
H("6.6 Sensitivity analysis", 2)
P("One-at-a-time first pass, then Sobol indices with an emulator (general practice). Prior ordering for ROS: fuel type, "
  "wind, moisture [Cardil et al 2023, p.6]. New check: whether information age (U5, U12) outranks fuel type for P(M<0) "
  "on blowup days.")

H("7. Experiment and simulation design")
T(["Factor", "Levels", "Basis"], [
 ["Fire case", "Camp 2018 (validation); Carr 2018 (validation); Pioneer 2016 (long timber); LNU 2020 (complex, extreme early growth)", "Case study roles, Section 4.4"],
 ["Time of day / visibility", "Day VMC; night DVE; smoke DVE", "Crewed window about 30% of 24 h [NASA ACERO ConOps 2026, pp.30-31]"],
 ["Information cadence", "Official record (about 24 h lag); perimeter record (median 17 to 19 h); satellite 12 h; aerial 30 min", "Section 4.2 and cadence sheet"],
 ["Network reliability", "52%; 87%; 95% message success", "[Wu et al 2025, pp.9-10]"],
 ["Model", "ELMFIRE; GridFire", "[Altintas 2025, p.8]"],
 ["Demand profile (M5)", "Recorded daily crews per fire", "NICC Daily IMSR sheets"],
])
P("A full factorial is 4 x 3 x 4 x 3 x 2 = 288 cells before demand profiles. Start with a fractional design or Latin "
  "hypercube (general practice), running Carr 26 July and Camp 8 November first, since both have timed entrapment events.")
H("7.1 Validation", 2)
P("Camp: NIST minute-level progression and 19 burnovers [Maranghides et al 2021, p.148]. Carr: timed aerial warning "
  "(18:38) and mayday (19:40) on 26 July from the CAL FIRE Green Sheet, plus 63 GeoMAC and 88 FEDS perimeters. Pioneer "
  "and LNU have no firefighter entrapment in one day's search (Research Data Index), so they validate spread and "
  "information age, not P(M<0). Forecasts are scored against time-matched perimeters with the observation-error "
  "layer applied, and P(M<0) is checked for elevation where burnovers occurred.")

H("8. Preliminary feasibility")
T(["Component", "Data and tools on hand", "Main risk", "Verdict"], [
 ["Ensemble spread", "ELMFIRE, GridFire; perimeters for all four fires", "Compute cost; restart support unconfirmed", "Feasible now"],
 ["Model-error layer", "Cardil residuals by fuel group", "California sample; Pioneer is Idaho timber", "Feasible now"],
 ["Observation-error layer", "This evaluation: lag, shortfall, satellite bias", "Three fires only; no GeoMAC for LNU", "Feasible now"],
 ["Data assimilation", "Perimeter histories at observed cadence", "Engineering effort", "Feasible, higher effort"],
 ["Logistics DES", "NICC daily crews and helicopters; PAMS link priors", "No flown delivery data", "Feasible with elicited priors"],
 ["Validation of P(M<0)", "Camp burnovers; Carr timed warning and mayday", "Two events with full timing", "Feasible, thin"],
 ["Archive quality", "Errors found and logged", "Silent errors in cleaned fields", "Fix database before analysis"],
])

H("9. Limitations")
B([
 "Three fires is a small sample; the lag and shortfall figures describe these incidents, not a national rate.",
 "The national situation report is derived from the 209, so the sources are not independent; true size uncertainty "
 "during growth is only bounded by the perimeter and satellite records, which have their own errors.",
 "FEDS satellite perimeters are biased by pixel size and show an unexplained late drop on LNU.",
 "The latency analysis compares calendar-day records in local time; exact report times would sharpen it.",
 "The \"no entrapment found\" results for Pioneer and LNU come from one day of searching (30 September 2026).",
 "Novelty is not assessed. Run the dated literature gate before any proposal language says this is new.",
])

H("10. Next steps")
N([
 "Team agrees on P(M<0) and P(on time) as headline metrics and on information age as a reported covariate.",
 "Build the observation-error layer from the Uncertainty Register (U12 to U15) and apply it to Camp and Carr scoring.",
 "Run a 100-member ELMFIRE ensemble for Carr 26 July and Camp 8 November, day and night, at the four cadence levels.",
 "Drive the logistics DES with Carr's recorded daily crew counts; compare 52% and 95% link success.",
 "Pursue the FOIA request for Pioneer dispatch and radio records for T_notify evidence; request NIROPS access for night perimeters.",
 "Run the literature gate on observation error in operational fire records before any novelty claim.",
])

H("References (project library)")
for ref in [
 "Altintas 2025. Actionable fire modeling in Firemap for extended attack decision support.",
 "Boroujeni et al 2024. A comprehensive survey of research towards AI-enabled unmanned aerial systems in pre-, active-, and post-wildfire management.",
 "Cardil et al 2023. Performance of operational fire spread models in California. International Journal of Wildland Fire.",
 "Hope et al 2024. A cost-benefit analysis of WildFireSat, a wildfire monitoring satellite mission for Canada. PLOS ONE.",
 "Kochanski et al 2023. Analysis of methods for assimilating fire perimeters into a coupled fire-atmosphere model.",
 "Maranghides et al 2021. A case study of the Camp Fire: Fire progression timeline. NIST TN 2135.",
 "NASA ACERO ConOps 2026. Kopardekar et al. 24/7 Wildland Fire Aerial Operations Airspace Management System ConOps.",
 "NWCG 2025. Standards for fire unmanned aircraft systems operations.",
 "Pyregence Consortium n.d. GridFire documentation.",
 "Wu et al 2025. Development of a portable airspace management system for wildland fire operations.",
]: d.add_paragraph(ref, style="List Bullet")
H("Data sources (case study evaluation)")
for ref in [
 "St. Denis et al. 2023. All-hazards dataset mined from the US National Incident Management System 1999 to 2020 (ICS-209-PLUS). Scientific Data.",
 "NICC Incident Management Situation Reports, 2016, 2018, 2020 (Internet Archive copies of famprod.nwcg.gov).",
 "NICC Wildland Fire Summary and Statistics Annual Reports 2016, 2018, 2020.",
 "NIFC InterAgencyFirePerimeterHistory and Historic GeoMAC Perimeters; CAL FIRE FRAP 2020 California Fire Perimeters.",
 "NASA FEDS 12-hourly VIIRS perimeters (Chen et al. 2022 algorithm), via NASA VEDA.",
 "Boise Interagency Dispatch Center 2016 Annual Report; CAL FIRE Carr Fire Green Sheet (in research_data/entrapments).",
]: d.add_paragraph(ref, style="List Bullet")
P("Page numbers in library citations are PDF page numbers as indexed by the project RAG.", italic=True)
docx_path = OUT / "UQ_Crew_Safety_Logistics_Proposal_v0.2.docx"
d.save(docx_path)

# ======================================================================== workbook
wb = Workbook()
HDR = PatternFill("solid", fgColor="1F3864"); NEW = PatternFill("solid", fgColor="E2EFDA")
def sheet(ws, header, rws, widths, new_col=None):
    ws.append(header)
    for r in rws: ws.append(r)
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws[1]: c.font = Font(bold=True, color="FFFFFF"); c.fill = HDR
    for row in ws.iter_rows():
        for c in row: c.alignment = Alignment(wrap_text=True, vertical="top")
    if new_col is not None:
        for row in ws.iter_rows(min_row=2):
            if str(row[new_col].value).startswith("v0.2"):
                for c in row: c.fill = NEW
    ws.freeze_panes = "A2"

sheet(wb.active, ["Item", "Detail"], [
 ["Version", f"v0.2, {DATE}. Supersedes UQ_Evaluation_Workbook_v0.1.xlsx (kept unchanged)."],
 ["What is new", "Fire Case Studies evaluation (Pioneer 2016, Carr 2018, LNU 2020): sheets Case Study Summary to Data Quality. Uncertainty Register rows U12 to U16 and the updated U5 are shaded green."],
 ["Built by", "case_study_evaluation.py (analysis) then build_uq_v02.py (this workbook and the v0.2 proposal)."],
 ["Inputs", "Fire Case Studies/*/ (ICS-209-PLUS sheets, NICC Daily IMSR sheets, final perimeters); research_data/datasets/01_progression_perimeters (GeoMAC, FEDS); research_data/Research_Data_Index.xlsx (timing anchors)."],
 ["Size field", "All size comparisons use the as-reported ICS-209 CURR_INCIDENT_AREA. The ICS-209-PLUS ACRES field is cleaned and capped at the final size."],
 ["Day alignment", "Each calendar day (local time) uses the last 209 report of the day; NICC is the report published the next morning; GeoMAC and FEDS are the last of the day."],
 ["Citations", "Literature: [Doc Year, p.N] from the project RAG. Case study figures: the named record and case_study_eval/results.json."],
], [18, 120])
wb.active.title = "README"

U = [
 ["U1", "Rate of spread model error", "M3, M4", "Epistemic", "Residual by fuel group", "MAPE 47% overall, shrub 26%, timber underpredicts", "Partly", "[Cardil et al 2023, pp.6-7]", "v0.1"],
 ["U2", "Fuel type map", "M3", "Epistemic", "Categorical swap probabilities", "Key driver of ROS bias", "No", "[Cardil et al 2023, p.6]", "v0.1"],
 ["U3", "Wind speed and direction", "M3, M5", "Mixed", "Correlated multivariate perturbation", "High wind overpredicts ROS, low wind underpredicts", "Yes (in-situ profiling)", "[Cardil et al 2023, p.7]; [Pyregence Consortium n.d. GridFire, p.113]", "v0.1"],
 ["U4", "Fuel moisture", "M3", "Epistemic", "Correlated perturbation", "Smaller contribution than fuel and wind", "Partly", "[Cardil et al 2023, p.6]", "v0.1"],
 ["U5", "Perimeter age at forecast start", "M1, M3", "Epistemic", "Empirical interval distribution",
  f"Median gap {P16['geomac_interval_h']['median']:.1f} h Pioneer, {C18['geomac_interval_h']['median']:.1f} h Carr; p90 {P16['geomac_interval_h']['p90']:.1f} h and {C18['geomac_interval_h']['p90']:.1f} h; max {P16['geomac_interval_h']['max']:.0f} h and {C18['geomac_interval_h']['max']:.0f} h; {P16['geomac_share_interval_over_24h']:.0%} / {C18['geomac_share_interval_over_24h']:.0%} of gaps over 24 h. Literature: VIIRS up to 4 h",
  "Yes, directly", "GeoMAC perimeter histories; [Kochanski et al 2023, p.12]", "v0.2 updated"],
 ["U6", "Message delivery success", "M3, M5", "Aleatory", "Bernoulli per message", "TFR 100%, OI 87%, telemetry 52%", "n/a", "[Wu et al 2025, pp.9-10]", "v0.1"],
 ["U7", "UAS path deviation from wind", "M2, M5", "Aleatory", "Transit time distribution", "Not quantified in library", "n/a", "[Keerthinathan et al 2023, p.35]", "v0.1"],
 ["U8", "Crew egress time", "M3", "Mixed", "Travel-rate model plus PERT check", "Sullivan 2020; Campbell 2017, 2019 travel rates (research_data)", "No", "research_data/07_crew_egress; entrapments/", "v0.1 (source added)"],
 ["U9", "Notification latency to crew", "M3", "Mixed", "PERT from elicitation + U6", "Carr: 62 min from aerial warning (18:38) to mayday (19:40); Camp: emergency call unanswered", "Partly", "Research_Data_Index, Margin Timing Anchors", "v0.1 (anchor added)"],
 ["U10", "Delivery window tolerance", "M5", "Epistemic", "PERT from elicitation", "NOT IN LIBRARY", "No", "Elicit (questionnaire)", "v0.1"],
 ["U11", "Human factors delay", "M3, M5", "Epistemic", "Multiplier on decision times", "Tempo, saturation, fatigue listed as factors", "No", "[NWCG 2025, p.21]", "v0.1"],
 ["U12", "Official size reporting lag", "M1, M3", "Epistemic", "Fixed lag of one reporting cycle, with probability of same-day update",
  f"209 equals previous night's perimeter on {chain['Pioneer 2016']['match_next_day_209']}/{chain['Pioneer 2016']['days_tested']} (Pioneer) and {chain['Carr 2018']['match_next_day_209']}/{chain['Carr 2018']['days_tested']} (Carr) changing days; NICC equals previous evening's 209 on {lag['Pioneer 2016']['exact_match_prev_day_209']}/{lag['Pioneer 2016']['days_compared']}, {lag['Carr 2018']['exact_match_prev_day_209']}/{lag['Carr 2018']['days_compared']}, {lag['LNU 2020']['exact_match_prev_day_209']}/{lag['LNU 2020']['days_compared']}",
  "Yes, directly", CS_REF, "v0.2 new"],
 ["U13", "Blowup-day size shortfall", "M1, M3", "Epistemic", "Ratio distribution conditional on growth regime",
  f"Same-day perimeter / 209: Carr 1.55 (26 Jul), 1.68 (27 Jul), 1.14 (28 Jul); Pioneer 1.12 (30 Aug). Satellite / 209: Carr up to 2.27; LNU {lnu_819[1]/lnu_819[0]:.2f} (19 Aug)",
  "Yes, directly", CS_REF, "v0.2 new"],
 ["U14", "Satellite area bias (FEDS, 375 m)", "M1", "Epistemic", "Multiplicative bias per fire",
  f"After growth stopped: Carr +{100*(C18['final_size_sources']['NASA FEDS last overpass']/C18['growth']['final_acres_209']-1):.0f}%, LNU +{100*(lnu_feds_stable/L20['growth']['final_acres_209']-1):.1f}%; LNU late drop to {f0(L20['final_size_sources']['NASA FEDS last overpass'])} unexplained",
  "n/a (sensor property)", "NASA FEDS 12 h perimeters", "v0.2 new"],
 ["U15", "Operational size revisions", "M1", "Epistemic", "Occasional downward jump", "LNU 375,209 to 363,220 (-3.2%), better mapping; none on Pioneer or Carr", "Yes", "NICC IMSR 2020-09-10 narrative; ICS-209-PLUS CURR_INCIDENT_AREA", "v0.2 new"],
 ["U16", "Logistics demand scale", "M5", "Aleatory (incident-driven)", "Recorded daily crew and helicopter series",
  f"Median crews {f0(res['Pioneer 2016']['crews']['median'])} / {f0(res['Carr 2018']['crews']['median'])} / {f0(res['LNU 2020']['crews']['median'])}; peak {res['Pioneer 2016']['crews']['max']} / {res['Carr 2018']['crews']['max']} / {res['LNU 2020']['crews']['max']}; crews per helicopter median {res['Pioneer 2016']['crews_per_helicopter']['median']:.1f} / {res['Carr 2018']['crews_per_helicopter']['median']:.1f} / {res['LNU 2020']['crews_per_helicopter']['median']:.1f}",
  "n/a", "NICC Daily IMSR sheets", "v0.2 new"],
]
sheet(wb.create_sheet("Uncertainty Register"), ["ID", "Uncertain input", "Mission", "Type", "How to represent", "Starting value / prior", "Reducible by aerial obs?", "Source", "Version"],
      U, [6, 28, 9, 12, 30, 60, 16, 40, 14], new_col=8)

sheet(wb.create_sheet("QoI Metrics"), ["Metric", "Mission", "Definition (proposed)", "Decision use", "Source for framing", "Version"], [
 ["Safety margin M", "M3", "T_fire - (T_detect + T_notify + T_egress)", "Per crew position", "Proposed; [NWCG 2025, pp.20-21]", "v0.1"],
 ["P(M < 0)", "M3", "Share of ensemble members with negative margin", "Probability axis of NWCG risk assessment", "[NWCG 2025, pp.20-21]", "v0.1"],
 ["5th percentile of M", "M3", "Bad-day margin", "Go / no-go threshold", "Proposed", "v0.1"],
 ["Warning lead time", "M1, M2, M3", "T_fire - T_notify", "Value of surveillance", "Proposed", "v0.1"],
 ["Information age at decision", "M1, M3", "Age of newest size or perimeter available to the crew boss", "Baseline for RQ4", "Section 4.2; U5, U12", "v0.2 new"],
 ["Size shortfall ratio", "M1", "Observed (perimeter) size / official size, same day", "How far behind the official picture is", "Section 4.3; U13", "v0.2 new"],
 ["P(on time)", "M5", "Deliveries inside window / total", "Dependability", "Proposed; [NASA ACERO ConOps 2026, p.19]", "v0.1"],
 ["Turnaround p50 / p90", "M5", "Sortie cycle time distribution", "Fleet sizing", "Proposed", "v0.1"],
 ["Crews served per aerial asset", "M5", "Daily crews / delivery aircraft", "Fleet sizing against real demand", "U16", "v0.2 new"],
 ["P(abort)", "M5", "Sorties cancelled / total", "Reliability", "Proposed", "v0.1"],
 ["ROS MAE / MBE / MAPE", "Validation", "Predicted vs time-matched observed", "Model skill", "[Cardil et al 2023, pp.5-6]", "v0.1 (time-matched in v0.2)"],
], [28, 11, 46, 34, 34, 22], new_col=5)

sheet(wb.create_sheet("Experiment Matrix"), ["Factor", "Levels", "Count", "Source"], [
 ["Fire case", "Camp 2018; Carr 2018; Pioneer 2016; LNU 2020", 4, "Case study roles (proposal 4.4, 7)"],
 ["Visibility", "Day VMC; night DVE; smoke DVE", 3, "[NASA ACERO ConOps 2026, pp.30-31]"],
 ["Information cadence", "Official ~24 h lag; perimeter median 17 to 19 h; satellite 12 h; aerial 30 min", 4, "Reporting Cadence sheet"],
 ["Network success", "52%; 87%; 95%", 3, "[Wu et al 2025, pp.9-10]"],
 ["Model", "ELMFIRE; GridFire", 2, "[Altintas 2025, p.8]"],
 ["Total cells (full factorial)", "", "=PRODUCT(C2:C6)", "Use fractional design first; run Carr 26 Jul and Camp 8 Nov first"],
], [26, 60, 12, 46])

sheet(wb.create_sheet("Feasibility"), ["Component", "On hand", "Main risk", "Verdict"], [
 ["Ensemble spread", "ELMFIRE, GridFire; perimeters for 4 fires", "Compute; restart support unconfirmed", "Feasible now"],
 ["Model-error layer", "Cardil residual metrics", "California sample", "Feasible now"],
 ["Observation-error layer", "U12 to U15 from this evaluation", "Three fires; no GeoMAC for LNU", "Feasible now"],
 ["Data assimilation", "Perimeter histories at observed cadence", "Engineering effort", "Feasible, higher effort"],
 ["Logistics DES", "NICC daily crews and helicopters; PAMS priors", "No flown delivery data", "Feasible with elicited priors"],
 ["Validation of P(M<0)", "Camp burnovers; Carr warning and mayday", "Two timed events", "Feasible, thin"],
 ["Archive quality", "Data Quality sheet", "Silent errors in cleaned fields", "Fix database first"],
], [24, 40, 34, 26])

cs = []
for k, x in R.items():
    cs += [
     [k, "ICS-209 reports / median interval (h) / gaps over 24 h", f"{x['209_reports']} / {x['209_interval_h']['median']:.1f} / {x['209_gaps_over_24h']}", "ICS-209-PLUS"],
     [k, "Longest 209 gap (h, from, to)", " / ".join(str(v) for v in x["209_longest_gap"]), "ICS-209-PLUS"],
     [k, "NICC reports listing the fire", x["nicc_reports_with_row"], "NICC IMSR"],
     [k, "Final size, all sources", "; ".join(f"{a}: {f0(b)}" for a, b in x["final_size_sources"].items()), "see sheet Size Reconciliation"],
     [k, "Peak as-reported vs cleaned 209 size", f"{f0(x['209_max_reported_vs_cleaned'][0])} vs {f0(x['209_max_reported_vs_cleaned'][1])}", "ICS-209-PLUS CURR_INCIDENT_AREA vs ACRES"],
     [k, "Top 3 growth days (acres)", "; ".join(f"{t['date']}: {f0(t['acres'])}" for t in x["growth"]["top3_days"]), "ICS-209-PLUS"],
     [k, "Personnel median / peak; personnel-days (NICC)", f"{f0(x['resources_nicc']['personnel']['median'])} / {f0(x['resources_nicc']['personnel']['max'])}; {f0(x['resources_nicc']['personnel_days'])}", "NICC IMSR"],
     [k, "Helicopters median / peak; helicopter-days", f"{f0(x['resources_nicc']['helicopters']['median'])} / {x['resources_nicc']['helicopters']['max']}; {x['resources_nicc']['helicopter_days']}", "NICC IMSR"],
     [k, "Cost to date at last NICC listing ($)", f0(x["resources_nicc"]["final_cost"]), "NICC IMSR"],
    ]
sheet(wb.create_sheet("Case Study Summary"), ["Fire", "Measure", "Value", "Record"], cs, [14, 46, 90, 36])

sbd = rows("size_by_day")
sheet(wb.create_sheet("Size Reconciliation"), ["Fire", "Date (local)", "ICS-209 last (as reported)", "NICC next morning", "GeoMAC last perimeter", "FEDS last overpass", "GeoMAC / 209", "FEDS / 209"],
      [[r["fire"], r["date"]] + [float(r[c]) if r[c] else None for c in ("ics209_last", "nicc_next_morning", "geomac_last", "feds_last")] +
       [round(float(r["geomac_last"]) / float(r["ics209_last"]), 3) if r["geomac_last"] and r["ics209_last"] else None,
        round(float(r["feds_last"]) / float(r["ics209_last"]), 3) if r["feds_last"] and r["ics209_last"] else None] for r in sbd],
      [14, 12, 18, 16, 18, 16, 12, 12])
for row in wb["Size Reconciliation"].iter_rows(min_row=2):
    for c in row[2:6]: c.number_format = "#,##0"

cad = rows("cadence")
lat = [[k, chain[k]["days_tested"] if chain[k] else None, chain[k]["match_next_day_209"] if chain[k] else None, chain[k]["match_same_day_209"] if chain[k] else None,
        lag[k]["days_compared"], lag[k]["exact_match_prev_day_209"], lag[k]["exact_match_same_day_209"]] for k in R]
ws = wb.create_sheet("Reporting Cadence")
sheet(ws, ["Fire", "Record", "Intervals (n)", "Median interval (h)", "p90 interval (h)", "Max interval (h)"],
      [[r["fire"], r["source"], int(float(r["n"])), round(float(r["median"]), 1), round(float(r["p90"]), 1), round(float(r["max"]), 1)] for r in cad], [14, 22, 14, 18, 16, 16])
ws.append([]); ws.append(["Latency chain", "Days 209 changed", "209 = previous night's perimeter", "209 = same-day perimeter", "Days NICC compared", "NICC = previous evening's 209", "NICC = same-day 209"])
for c in ws[ws.max_row]: c.font = Font(bold=True)
for r in lat: ws.append(r)
ws.append(["Match = within 0.5% for perimeters, exact for NICC. LNU has no GeoMAC history."])

bl = []
for k in R:
    for s in sorted(R[k]["blowup_day_shortfall"], key=lambda s: s["date"]):
        bl.append([k, s["date"], s["ics209"], s["geomac"], s["feds"], round(s["ratio_geomac"], 2) if s["ratio_geomac"] else None, round(s["ratio_feds"], 2) if s["ratio_feds"] else None])
bl.append(["LNU 2020", "2020-08-19", lnu_819[0], None, lnu_819[1], None, round(lnu_819[1] / lnu_819[0], 2)])
ws = wb.create_sheet("Blowup Shortfall")
sheet(ws, ["Fire", "Date", "ICS-209 acres", "Same-day perimeter", "Same-day FEDS", "Perimeter / 209", "FEDS / 209"], bl, [14, 12, 14, 18, 16, 14, 12])
ws.append([]); ws.append(["Carr 26 Jul 2018: fatal burnovers that evening; aerial warning 18:38, mayday 19:40 (Research_Data_Index, Margin Timing Anchors). FEDS includes pixel bias (U14); the perimeter ratio is the conservative measure."])

sheet(wb.create_sheet("Daily Growth"), ["Fire", "Date", "Growth (acres)", "Growth / previous size"],
      [[r["fire"], r["date"], float(r["acres"]), round(float(r["rel"]), 3) if r["rel"] else None] for r in rows("growth")], [14, 12, 16, 22])

sheet(wb.create_sheet("Resources M5"), ["Fire", "NICC report date", "Personnel", "Crews", "Engines", "Helicopters", "Cost to date ($)", "% contained"],
      [[r["fire"], r["imsr_date"]] + [float(r[c]) if r[c] not in ("", None) else None for c in ("personnel", "crews", "engines", "helicopters", "cost", "pct_contained")] for r in res_rows],
      [14, 16, 12, 10, 10, 12, 18, 12])

sheet(wb.create_sheet("Data Quality"), ["ID", "Issue", "Fires", "Effect on UQ", "Handling", "Evidence"], [
 ["DQ1", "ICS-209-PLUS ACRES is cleaned and capped at final size", f"LNU ({L20['209_rows_reported_ne_cleaned']} of {L20['209_reports']} reports); Carr 1; Pioneer 0", "Hides real-time overestimates and revisions", "Use CURR_INCIDENT_AREA", "All 209 fields sheets"],
 ["DQ3", "Two personnel definitions", "All three", f"Demand off by {L20['personnel_definition_gap']['ratio']:.1f} to {C18['personnel_definition_gap']['ratio']:.1f} times", "Use NICC counts for M5 demand", "Case Study Summary"],
 ["DQ4", "NICC is derived from the 209", "All three", "Not an independent size check", "Count as one source", "Reporting Cadence sheet"],
 ["DQ5", "Cause disagrees across records", "All three", "Matters only if cause enters a model", "Record both", "209/NICC U vs BDC human (Pioneer) and CAL FIRE lightning (LNU)"],
 ["DQ6", "Damaged structures total suspect", "Carr, LNU", "Not used in this evaluation", "Use residential split or DINS", "Carr 1,610 vs 192 residential; LNU 1,469 vs 112"],
 ["DQ7", "Reporting gaps", "LNU (16 to 30 Sep), Pioneer (169 h Oct)", "Gaps inflate apparent perimeter age late in incidents", "Exclude post-containment periods from cadence", "Reporting Cadence sheet"],
 ["DQ8", "NICC PDF transcription errors", "Pioneer (1.506), Carr (3.831)", "Small", "Corrected and flagged in case study sheets", "NICC Daily IMSR sheets"],
 ["DQ9", "FEDS late drop", "LNU (372,519 to 312,526 from 18 Sep)", "Would understate final size", "Use last stable overpass", "Size Reconciliation"],
 ["DQ10", "Detection time conflict", "Pioneer: reported 1717 by Casner Fire lead plane (Wildfire Today, 2 Sep 2016) vs 209 discovery 18:09", "T_detect anchor uncertain by about 50 min", "Carry as timing uncertainty", "Pioneer case study Summary sheet"],
], [7, 42, 34, 36, 34, 50])

xlsx_path = OUT / "UQ_Evaluation_Workbook_v0.2.xlsx"
wb.save(xlsx_path)
print(docx_path); print(xlsx_path)
