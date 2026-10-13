"""Builds the UQ proposal (docx) and the companion evaluation workbook (xlsx).
Every factual claim carries a [Doc Year, p.N] citation from ~/nasa_rag.
Run: uv run --project ~/nasa_rag python build_uq.py"""
from pathlib import Path
from docx import Document
from docx.shared import Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

OUT = Path(__file__).parent
DATE = "29 September 2026"

# ---------------------------------------------------------------- document
d = Document()
st = d.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(11)

def H(t, l=1): d.add_heading(t, level=l)
def P(t, italic=False, bold=False):
    p = d.add_paragraph(); r = p.add_run(t); r.italic = italic; r.bold = bold; return p
def B(items):
    for t in items: d.add_paragraph(t, style="List Bullet")
def N(items):
    for t in items: d.add_paragraph(t, style="List Number")
def T(header, rows, widths=None):
    t = d.add_table(rows=1, cols=len(header)); t.style = "Light Grid Accent 1"
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]; c.text = ""; c.paragraphs[0].add_run(h).bold = True
    for row in rows:
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
sp.add_run("A preliminary feasibility proposal for experiments and simulations\n").italic = True
sp.add_run(f"NASA Aerial Wildfire Project, ACERO Disaster Lab working group\nDraft v0.1, {DATE}")

P("Status of this draft. Preliminary thinking for team discussion, not a submission. Every "
  "factual claim is grounded in the project reference library (NASA RAG, 41 indexed documents) "
  "and cited as [Document Year, p.N]. Methods marked \"general practice\" come from standard "
  "uncertainty quantification methodology and are not drawn from the library. This draft makes "
  "no claim that the approach is novel; the library answers \"what do our references say\", not "
  "\"what has the field done\". A dated literature search is required before any novelty framing.",
  italic=True)

H("Abstract")
P("NASA's ACERO ConOps sets the goal of 24/7 aerial operations for wildland fire, organised in this "
  "project around surveillance (M1 fireline mapping, M2 spot fire detection, M3 firefighter safety) and "
  "suppression (M4 water drop suppression, M5 water delivery). Two of these missions are ultimately "
  "judged by outcomes for people on the ground: whether crews stay out of the fire's path (M3) and "
  "whether supplies and water reach them when needed (M5). Both outcomes depend on chains of uncertain "
  "quantities: fire spread predictions whose operational error in California averaged a 47% mean absolute "
  "percentage error in rate of spread [Cardil et al 2023, pp.6-7], communication links that delivered "
  "only 52% of telemetry messages in NASA's own field data collection [Wu et al 2025, pp.9-10], and "
  "perimeter observations that can be hours old when a forecast starts [Kochanski et al 2023, p.12]. "
  "This proposal outlines how to turn those uncertainties into decision-grade numbers: a probability "
  "that a crew's safety margin goes negative, and a probability that a delivery arrives inside its "
  "required window. We propose a layered approach: ensemble fire spread simulation with correlated "
  "input perturbations, an empirical model-error layer calibrated against observed fires, a stochastic "
  "operations simulation for aerial logistics, and expert elicitation with PERT distributions where no "
  "data exist, followed by global sensitivity analysis to find which uncertainties matter most. The "
  "Camp Fire, with its NIST minute-level progression and burnover record, is proposed as the "
  "validation case. We assess each component for feasibility with the data and models the team has now.")

H("1. Motivation")
H("1.1 Why crew safety needs a number, not a narrative", 2)
P("The wildland fire airspace \"has repeatedly proven to be among the most hazardous airspace environments "
  "in government aviation operations\", with near misses, controlled flight into terrain, and fatal mid-air "
  "collisions documented in DOI safety reports from 2008 to 2022 [NASA ACERO ConOps 2026, p.12]. The "
  "ConOps names improving the safety of current crewed operations as \"the top priority\" of the future "
  "state [NASA ACERO ConOps 2026, p.16]. On the ground, crews avoid night operations for safety reasons in "
  "low visibility [NASA ACERO ConOps 2026, p.19], and the Camp Fire record shows 19 burnovers, one of which "
  "involved hand crews and caused serious injuries; burnovers typically lasted 1 to 2 hours "
  "[Maranghides et al 2021, p.148].")
P("NWCG already frames UAS safety as risk: \"Risk levels are established by the severity of possible "
  "events and the probability that they will occur\" [NWCG 2025, pp.20-21], and notes that hazards may "
  "include \"hazards to personnel if the flight is not performed\" [NWCG 2025, p.21]. That second point "
  "is the core of this proposal. A 24/7 aerial capability changes the probability term for ground crews, "
  "and we need a way to measure by how much, with honest error bars.")
H("1.2 Why logistics needs a number", 2)
P("The ConOps describes drone-enabled BVLOS logistics as a prime uncrewed use case: food, drinking water, "
  "radios, saws, hand tools, hose, and fuel delivered to crews in terrain and smoke where crewed flight is "
  "not acceptable, with crews able to move the rendezvous point mid-flight, and even \"emergency extraction "
  "of an injured firefighter amid a potential burn over\" [NASA ACERO ConOps 2026, p.19]. Whether that "
  "concept is useful depends on reliability, not capability: how often a delivery arrives on time, given "
  "wind, endurance, and a network that is not yet robust [Wu et al 2025, p.11].")
H("1.3 The time window at stake", 2)
P("Crewed aircraft can currently work safely in roughly 30% of a 24-hour cycle, and a DVE solution could "
  "expand this by almost 250% [NASA ACERO ConOps 2026, pp.30-31]. Nearly 20% of wildland fire ignitions "
  "in a 2015 to 2017 DOI assessment occurred outside aerial operating hours, some growing uncontested for "
  "up to 16 hours [NASA ACERO ConOps 2026, pp.15-16]. Uncertainty quantification lets the team state the "
  "benefit of closing that window as a distribution rather than a single optimistic figure.")

H("2. Research questions")
N([
 "RQ1 (M3, crew safety). For a crew at a given position and time, what is the probability that fire "
 "arrival precedes the crew's completed egress to a safety zone, and how does that probability change with "
 "and without 24/7 aerial surveillance (M1, M2)?",
 "RQ2 (M5, logistics and water delivery). What is the probability that a requested delivery arrives within "
 "its required window, and what is the distribution of sortie turnaround time, under day, night, and "
 "smoke (DVE) conditions?",
 "RQ3 (drivers). Which input uncertainties (fuel, wind, fuel moisture, perimeter age, communication "
 "success, aircraft endurance) dominate the variance of the RQ1 and RQ2 outputs?",
 "RQ4 (value of information). How much does fresher aerial observation (shorter perimeter latency, "
 "continuous spot fire detection) shrink the uncertainty in RQ1 and RQ2?",
])

H("3. Where the uncertainty comes from")
P("We separate aleatory uncertainty (true variability such as wind gusts and ember spotting, which "
  "cannot be reduced) from epistemic uncertainty (lack of knowledge, such as fuel maps and model bias, "
  "which better data can reduce). This split is general practice; it matters here because RQ4 is only "
  "answerable for the epistemic part.")
T(["Source", "Type", "What the library says", "Citation"], [
 ["Fire spread model structure", "Epistemic", "Operational Rothermel-based simulations of 1,853 California fires: average ROS MAPE 47%, best in shrub (26%); underprediction bias in most timber fuels", "[Cardil et al 2023, pp.6-7]; [Cardil et al 2023, p.7]"],
 ["Fuel type mapping", "Epistemic", "Fuel type was the key contributor to ROS bias; wind secondary; fuel moisture smaller", "[Cardil et al 2023, p.6]"],
 ["Wind input", "Mixed", "High wind led to ROS overestimation, low wind to underestimation; terrain-driven wind field error could be corrected with nearby station data", "[Cardil et al 2023, p.7]; [Cardil et al 2023, p.9]"],
 ["Correlated input errors", "Epistemic", "Independent uniform perturbations fall short; weather error is spatially correlated, not i.i.d. noise; use multivariate distributions upstream", "[Pyregence Consortium n.d. GridFire, p.113]"],
 ["Initial fire state (perimeter)", "Epistemic", "Forecast growth error 3 to 35 times smaller with spin-up initialisation; satellite perimeter latency up to 4 h (VIIRS); overnight IR perimeters often outdated", "[Kochanski et al 2023, p.12]; [Kochanski et al 2023, pp.11-12]"],
 ["Physics model fidelity", "Epistemic", "FDS predictions range from within experimental uncertainty to about 20% off measurements in validation cases", "[McGrattan et al 2013, pp.30-32]"],
 ["Communication network", "Aleatory in operation", "PAMS field data: TFR 100%, operational intent 87%, telemetry 52% success; >90% in shakedown with fewer, closer nodes", "[Wu et al 2025, pp.9-10]; [Wu et al 2025, p.11]"],
 ["UAS path and endurance", "Mixed", "Path planning studies have not incorporated wind; wind disturbance can cause incomplete detection; endurance is limited", "[Keerthinathan et al 2023, p.35]"],
 ["Human factors", "Epistemic", "NWCG lists tempo, task saturation, fatigue, normalised risk, lack of SA, rushing as risk factors", "[NWCG 2025, p.21]"],
 ["Benefit assumptions", "Epistemic", "Expert opinions on benefit varied widely; represented with PERT distributions in Monte Carlo", "[Hope et al 2024, pp.7-8]"],
])

H("4. Quantities of interest")
P("Uncertainty is only useful if it is attached to a quantity a decision maker cares about. We propose "
  "the following outputs. Formulas are proposed definitions for this project, not taken from the library.")
H("4.1 Crew safety (M3)", 2)
B([
 "Safety margin M = T_fire - (T_detect + T_notify + T_egress), where T_fire is fire arrival time at the crew "
 "position from the spread ensemble, T_detect is time for surveillance to see the threat (M1, M2), T_notify is "
 "message latency over the network, and T_egress is crew travel time to the safety zone.",
 "Primary metric: P(M < 0), the probability of an entrapment-exposure event. Secondary: the 5th percentile of M "
 "(the \"bad day\" margin), and expected warning lead time.",
 "Mapping to NWCG practice: P(M < 0) supplies the probability axis of the severity by probability risk "
 "assessment NWCG already requires [NWCG 2025, pp.20-21].",
])
H("4.2 Logistics and water delivery (M5)", 2)
B([
 "P(on time): probability a delivery lands within its requested window.",
 "Sortie turnaround time distribution (median and 90th percentile), and deliveries per crew-hour.",
 "P(abort): probability a sortie is cancelled for wind, link loss, airspace conflict, or reroute.",
 "Link-loss exposure: expected minutes per sortie without telemetry, using field success rates "
 "[Wu et al 2025, pp.9-10] as the starting prior.",
])
H("4.3 Suppression (M4), noted but out of scope for now", 2)
P("Drop accuracy and coverage are measurable from ortho-rectified airborne infrared imagery of drops "
  "[Boroujeni et al 2024, p.38], but no drone in the current library has flown water drops, so M4 remains "
  "a data gap for this proposal.")

H("5. Proposed methods")
H("5.1 Ensemble fire spread with correlated perturbations", 2)
P("Run each scenario as an ensemble (hundreds of members) in ELMFIRE or GridFire, perturbing fuel, wind, "
  "and moisture from correlated, spatially structured distributions rather than independent uniform noise, "
  "as GridFire's own documentation recommends [Pyregence Consortium n.d. GridFire, p.113]. ELMFIRE and "
  "GridFire suit extended attack timescales, FARSITE the short-term immediate behaviour "
  "[Altintas 2025, p.8]. Output: a map of fire arrival time distributions, from which T_fire at any crew "
  "position is read directly.")
H("5.2 Empirical model-error layer", 2)
P("Ensembles over inputs do not capture model structural error. We propose adding a residual term drawn "
  "from observed-versus-predicted ROS residuals by fuel group, following the metrics Cardil et al. used "
  "(MAE, MBE, MAPE by Scott and Burgan fuel group) [Cardil et al 2023, pp.5-6]. Timber cases should carry "
  "the documented underprediction bias [Cardil et al 2023, p.7]; this is the direction that matters most "
  "for crew safety, since underpredicted spread means overestimated margin.")
H("5.3 Data assimilation as the lever for RQ4", 2)
P("Particle filters (sequential Monte Carlo) are used to assimilate observations into spread models and to "
  "feed back uncertainty on state variables [Boroujeni et al 2024, p.46]; adaptive particle numbers scale with "
  "inferred state uncertainty [Boroujeni et al 2024, pp.46-47]. Bayesian calibration with a Gaussian process "
  "emulator has been used to keep the cost manageable [Boroujeni et al 2024, p.46]. Perimeter initialisation "
  "method strongly affects forecast error [Kochanski et al 2023, p.12]. The experiment: rerun the same ensemble "
  "with perimeter updates every 30 minutes (24/7 aerial surveillance) versus every 12 hours or with a 4 hour "
  "satellite lag, and measure how much the spread of T_fire shrinks.")
H("5.4 Stochastic operations simulation for logistics", 2)
P("Model each sortie as a discrete-event process (dispatch, climb, transit, rendezvous, drop, return, "
  "recharge) with random draws for wind-limited availability, transit speed, link success, and reroute "
  "requests. The ConOps supplies the mission logic, including mid-flight rendezvous changes "
  "[NASA ACERO ConOps 2026, p.19]; the PAMS field test supplies the first link-reliability priors "
  "[Wu et al 2025, pp.9-10]. The discrete-event formulation itself is general practice.")
H("5.5 Expert elicitation where there is no data", 2)
P("Several inputs (crew egress speed in steep terrain at night, delivery window tolerance, human factor "
  "delays) have no measured distribution in the library. Follow the WildFireSat approach: elicit minimum, "
  "most likely, and maximum from practitioners and fit PERT distributions, which weight the most likely "
  "value more than triangular ones [Hope et al 2024, pp.7-8]; run to convergence (they used 50,000 "
  "iterations) [Hope et al 2024, p.3]. Christensen used triangular distributions with 5,000 Monte Carlo runs "
  "for a similar UAV evaluation with limited data [Christensen 2015, p.4]. This fits the team's planned "
  "stakeholder questionnaire.")
H("5.6 Global sensitivity analysis", 2)
P("Answer RQ3 with variance-based sensitivity indices (Sobol) or, as a cheaper first pass, one-at-a-time "
  "tornado plots (general practice). Cardil et al. already give a prior expectation: fuel type first, "
  "wind second, moisture smaller [Cardil et al 2023, p.6]. Checking whether that ordering holds for "
  "the safety margin, not just ROS, is a concrete early result.")
H("5.7 Communicating the result", 2)
P("Firemap's authors list refined uncertainty visualisation for risk communication as future work "
  "[Altintas 2025, p.8]. We propose three outputs a crew boss or air operations branch director can "
  "read: a probability-of-arrival-by-time map, a traffic-light P(M<0) per crew position, and a delivery "
  "reliability curve by time of day.")

H("6. Experiment and simulation design")
T(["Factor", "Levels", "Why"], [
 ["Fire case", "Camp Fire (validation); Pioneer Fire", "NASA team candidate fires; Camp has NIST minute-level progression and burnover data [Maranghides et al 2021, p.148]"],
 ["Time of day / visibility", "Day VMC; night DVE; smoke DVE", "Crewed window is ~30% of 24 h [NASA ACERO ConOps 2026, pp.30-31]"],
 ["Surveillance cadence", "None; 12 h IR; 4 h satellite lag; 30 min aerial", "Tests RQ4, perimeter latency effect [Kochanski et al 2023, p.12]"],
 ["Network reliability", "52%; 87%; 95% message success", "Field range from PAMS [Wu et al 2025, pp.9-10]"],
 ["Model", "ELMFIRE; GridFire (FARSITE for short horizon)", "Timescale fit [Altintas 2025, p.8]"],
 ["Fleet mix (M5)", "1 to N delivery UAS", "Delivery reliability versus fleet size"],
])
P("A full factorial is 2 x 3 x 4 x 3 x 2 = 144 cells before fleet mix, each an ensemble. Recommended "
  "start: a fractional design or Latin hypercube over the continuous factors (general practice), with "
  "the Camp Fire day and night cells run first.")
H("6.1 Validation", 2)
P("Score ensemble forecasts against the Camp Fire progression timeline using Cardil's MAE, MBE, and MAPE "
  "[Cardil et al 2023, pp.5-6], plus a check that observed arrival times fall inside the predicted "
  "intervals at the stated rate (calibration, general practice). Burnover locations and times "
  "[Maranghides et al 2021, p.148] provide ground truth events for testing whether P(M<0) was elevated "
  "where burnovers actually happened.")

H("7. Preliminary feasibility")
T(["Component", "Data and tools on hand", "Main risk", "Verdict"], [
 ["Ensemble spread (5.1)", "ELMFIRE, GridFire documented; ensemble perturbation built into GridFire", "Compute cost; SimFIRE/ELMFIRE restart support still to confirm", "Feasible now"],
 ["Model-error layer (5.2)", "Published residual metrics by fuel group", "Residuals are California operational runs; transfer to other regions unverified", "Feasible now"],
 ["Data assimilation (5.3)", "Methods described in the survey literature; Camp perimeters", "Engineering effort; emulator may be needed", "Feasible, higher effort"],
 ["Logistics DES (5.4)", "ConOps mission logic; PAMS link priors", "No flown delivery drone data in library; M5 is a gap", "Feasible with elicited priors"],
 ["Expert elicitation (5.5)", "Questionnaire already planned", "Small, variable expert pool", "Feasible now"],
 ["Sensitivity (5.6)", "Standard methods", "Run count multiplies ensemble cost", "Feasible with emulator or OAT first pass"],
 ["Crew egress data", "Not found in the library", "Core input to RQ1", "Gap: elicit, then seek data"],
])

H("8. Limitations and open questions")
B([
 "The model-error layer is only as transferable as Cardil's California sample; transfer to the Pioneer Fire's fuels and region is untested.",
 "Crew egress times, notification latency on the fireline, and delivery tolerances are not in the library; "
 "the first version will rest on elicitation.",
 "Network figures come from one early field campaign and the authors attribute the drop to relay geometry "
 "[Wu et al 2025, p.11]; they are priors, not stable rates.",
 "No library drone has flown water drops or deliveries, so M4 and M5 results are simulation only.",
 "Restart from a saved fire front is required for assimilation cycles; FARSITE cannot, SimFIRE and ELMFIRE "
 "support is still to be confirmed.",
 "Novelty: not assessed. Before any public or proposal language says this is new, run the dated literature "
 "gate (at least four sources, five phrasings, grey literature, agency tools).",
])

H("9. Proposed next steps")
N([
 "Team agrees on the two headline metrics: P(M<0) for M3, P(on time) for M5.",
 "Confirm ELMFIRE restart capability; set up a 100-member Camp Fire ensemble for one day and one night period.",
 "Add egress, notification, and delivery-window questions to the stakeholder questionnaire (min, most likely, max).",
 "Build the logistics DES with PAMS link priors; compare 52% and 95% link success.",
 "First sensitivity pass (one-at-a-time) to check whether fuel type still dominates for the safety margin.",
 "Run the literature gate on \"uncertainty quantification firefighter safety margin\" and \"UAS logistics reliability wildfire\" before writing any novelty claim.",
])

H("References (project library)")
for ref in [
 "Altintas 2025. Actionable fire modeling in Firemap for extended attack decision support.",
 "Boroujeni et al 2024. A comprehensive survey of research towards AI-enabled unmanned aerial systems in pre-, active-, and post-wildfire management. Information Fusion 108.",
 "Cardil et al 2023. Performance of operational fire spread models in California. International Journal of Wildland Fire. doi:10.1071/WF22128.",
 "Christensen 2015. Use of UAV or remotely piloted aircraft and forward-looking infrared in forest, rural and wildland fire management: evaluation using simple economic analysis.",
 "Hope et al 2024. A cost-benefit analysis of WildFireSat, a wildfire monitoring satellite mission for Canada. PLOS ONE. doi:10.1371/journal.pone.0302699.",
 "Keerthinathan et al 2023. Exploring unmanned aerial systems operations in wildfire management: data types, processing algorithms and navigation.",
 "Kochanski et al 2023. Analysis of methods for assimilating fire perimeters into a coupled fire-atmosphere model. Frontiers in Forests and Global Change. doi:10.3389/ffgc.2023.1203578.",
 "Maranghides et al 2021. A case study of the Camp Fire: Fire progression timeline. NIST TN 2135. doi:10.6028/NIST.TN.2135.",
 "McGrattan et al 2013. Fire Dynamics Simulator, technical reference guide, sixth edition.",
 "NASA ACERO ConOps 2026. Kopardekar et al. NASA ACERO 24/7 Wildland Fire Aerial Operations Airspace Management System ConOps (20260002902).",
 "NWCG 2025. Standards for fire unmanned aircraft systems operations.",
 "Pyregence Consortium n.d. GridFire documentation.",
 "Wu et al 2025. Development of a portable airspace management system for wildland fire operations.",
]:
    d.add_paragraph(ref, style="List Bullet")
P("Page numbers in citations are PDF page numbers as indexed by the project RAG, which can differ from the "
  "printed page number.", italic=True)

docx_path = OUT / "UQ_Crew_Safety_Logistics_Proposal_v0.1.docx"
d.save(docx_path)

# ---------------------------------------------------------------- workbook
wb = Workbook()
def sheet(ws, header, rows, widths):
    ws.append(header)
    for r in rows: ws.append(r)
    for i, w in enumerate(widths, 1): ws.column_dimensions[get_column_letter(i)].width = w
    for c in ws[1]:
        c.font = Font(bold=True, color="FFFFFF"); c.fill = PatternFill("solid", fgColor="1F3864")
    for row in ws.iter_rows():
        for c in row: c.alignment = Alignment(wrap_text=True, vertical="top")
    ws.freeze_panes = "A2"

sheet(wb.active, ["ID", "Uncertain input", "Mission", "Type", "How to represent", "Starting value / prior", "Reducible by aerial obs?", "Source"], [
 ["U1", "Rate of spread model error", "M3, M4", "Epistemic", "Residual by fuel group added to each member", "MAPE 47% overall, shrub 26%, timber underpredicts", "Partly (assimilation)", "[Cardil et al 2023, pp.6-7]"],
 ["U2", "Fuel type map", "M3", "Epistemic", "Categorical swap probabilities", "Key driver of ROS bias", "No (needs ground/fuel data)", "[Cardil et al 2023, p.6]"],
 ["U3", "Wind speed and direction", "M3, M5", "Mixed", "Spatially correlated multivariate perturbation", "High wind overpredicts ROS, low wind underpredicts", "Yes (in-situ profiling)", "[Cardil et al 2023, p.7]; [Pyregence Consortium n.d. GridFire, p.113]"],
 ["U4", "Fuel moisture", "M3", "Epistemic", "Correlated perturbation", "Smaller contribution than fuel and wind", "Partly", "[Cardil et al 2023, p.6]"],
 ["U5", "Perimeter age at forecast start", "M1, M3", "Epistemic", "Scenario factor", "Up to 4 h (VIIRS); overnight IR often outdated", "Yes, directly", "[Kochanski et al 2023, p.12]"],
 ["U6", "Message delivery success", "M3, M5", "Aleatory", "Bernoulli per message", "TFR 100%, OI 87%, telemetry 52%", "n/a (network design)", "[Wu et al 2025, pp.9-10]"],
 ["U7", "UAS path deviation from wind", "M2, M5", "Aleatory", "Transit time distribution", "Not quantified in library", "n/a", "[Keerthinathan et al 2023, p.35]"],
 ["U8", "Crew egress time", "M3", "Mixed", "PERT from elicitation", "NOT IN LIBRARY", "No", "Elicit (questionnaire)"],
 ["U9", "Notification latency to crew", "M3", "Mixed", "PERT from elicitation + U6", "NOT IN LIBRARY", "Partly", "Elicit (questionnaire)"],
 ["U10", "Delivery window tolerance", "M5", "Epistemic", "PERT from elicitation", "NOT IN LIBRARY", "No", "Elicit (questionnaire)"],
 ["U11", "Human factors delay", "M3, M5", "Epistemic", "Multiplier on decision times", "Tempo, saturation, fatigue listed as factors", "No", "[NWCG 2025, p.21]"],
], [6, 30, 10, 12, 34, 38, 22, 40])
wb.active.title = "Uncertainty Register"

sheet(wb.create_sheet("QoI Metrics"), ["Metric", "Mission", "Definition (proposed)", "Decision use", "Source for framing"], [
 ["Safety margin M", "M3", "T_fire - (T_detect + T_notify + T_egress)", "Per crew position", "Proposed; risk framing [NWCG 2025, pp.20-21]"],
 ["P(M < 0)", "M3", "Share of ensemble members with negative margin", "Probability axis of NWCG risk assessment", "[NWCG 2025, pp.20-21]"],
 ["5th percentile of M", "M3", "Bad-day margin", "Go / no-go threshold", "Proposed"],
 ["Warning lead time", "M1, M2, M3", "T_fire - T_notify", "Value of surveillance", "Proposed"],
 ["P(on time)", "M5", "Deliveries inside requested window / total", "Is the logistics concept dependable", "Proposed; concept [NASA ACERO ConOps 2026, p.19]"],
 ["Turnaround p50 / p90", "M5", "Sortie cycle time distribution", "Fleet sizing", "Proposed"],
 ["P(abort)", "M5", "Sorties cancelled / total", "Reliability", "Proposed"],
 ["Link-loss minutes per sortie", "M5", "Expected minutes without telemetry", "Network requirement", "[Wu et al 2025, pp.9-10]"],
 ["ROS MAE / MBE / MAPE", "Validation", "Predicted vs observed ROS", "Model skill", "[Cardil et al 2023, pp.5-6]"],
], [26, 12, 44, 34, 44])

sheet(wb.create_sheet("Experiment Matrix"), ["Factor", "Levels", "Count", "Source"], [
 ["Fire case", "Camp; Pioneer", 2, "NASA team candidate fires"],
 ["Visibility", "Day VMC; night DVE; smoke DVE", 3, "[NASA ACERO ConOps 2026, pp.30-31]"],
 ["Surveillance cadence", "None; 12 h IR; 4 h satellite; 30 min aerial", 4, "[Kochanski et al 2023, p.12]"],
 ["Network success", "52%; 87%; 95%", 3, "[Wu et al 2025, pp.9-10]"],
 ["Model", "ELMFIRE; GridFire", 2, "[Altintas 2025, p.8]"],
 ["Total cells (full factorial)", "", "=PRODUCT(C2:C6)", "Use fractional design first"],
], [26, 46, 12, 36])

sheet(wb.create_sheet("Feasibility"), ["Component", "On hand", "Main risk", "Verdict"], [
 ["Ensemble spread", "ELMFIRE, GridFire perturbations", "Compute; restart support unconfirmed", "Feasible now"],
 ["Model-error layer", "Cardil residual metrics", "California-only sample", "Feasible now"],
 ["Data assimilation", "Survey methods; Camp perimeters", "Engineering effort", "Feasible, higher effort"],
 ["Logistics DES", "ConOps logic; PAMS priors", "No flown delivery data (M5 gap)", "Feasible with elicited priors"],
 ["Expert elicitation", "Questionnaire planned", "Small expert pool", "Feasible now"],
 ["Sensitivity analysis", "Standard methods", "Run count", "OAT first, Sobol with emulator"],
 ["Crew egress data", "None in library", "Core RQ1 input", "GAP"],
], [24, 34, 34, 28])

xlsx_path = OUT / "UQ_Evaluation_Workbook_v0.1.xlsx"
wb.save(xlsx_path)
print(docx_path); print(xlsx_path)
