"""Build the v0.5 key findings and data needs brief.

A short companion to UQ_Crew_Safety_Logistics_Proposal_v0.5.docx: the findings and decision points of
v0.5 in two pages, then the data still missing for a complete study (radio and dispatch records,
night records, wind, aircraft, crew position, fuels, drone performance), with what is already in
research_data/ and how each missing item can be obtained. Numbers are copied from v0.5; nothing new
is computed here except the DOI UAS name search noted in Section 3.

    uv run --project ~/nasa_rag python build_uq_v05_brief.py
"""
from pathlib import Path
from docx import Document
from docx.shared import Pt, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH

HERE = Path(__file__).parent
OUT = HERE / "UQ_v0.5_Key_Findings_and_Data_Needs.docx"
DATE = "13 October 2026"

d = Document()
st = d.styles["Normal"]; st.font.name = "Calibri"; st.font.size = Pt(10.5)
for s in d.sections:
    s.left_margin = s.right_margin = Inches(0.8); s.top_margin = s.bottom_margin = Inches(0.8)
def H(t, l=1): d.add_heading(t, level=l)
def P(t, italic=False, size=None):
    p = d.add_paragraph(); r = p.add_run(t); r.italic = italic
    if size: r.font.size = Pt(size)
    return p
def LB(lead, text, style="List Bullet"):
    p = d.add_paragraph(style=style); p.add_run(lead).bold = True; p.add_run(text)
def T(header, rws, size=8.5, widths=None):
    t = d.add_table(rows=1, cols=len(header)); t.style = "Light Grid Accent 1"
    for i, h in enumerate(header):
        c = t.rows[0].cells[i]; c.text = ""; c.paragraphs[0].add_run(h).bold = True
    for row in rws:
        cells = t.add_row().cells
        for i, v in enumerate(row): cells[i].text = str(v)
    for row in t.rows:
        for i, c in enumerate(row.cells):
            if widths: c.width = Inches(widths[i])
            for p in c.paragraphs:
                for r in p.runs: r.font.size = Pt(size)
    d.add_paragraph()

# ------------------------------------------------------------------ title
tp = d.add_paragraph(); tp.alignment = WD_ALIGN_PARAGRAPH.CENTER
r = tp.add_run("UQ Proposal v0.5: Key Findings and Data Needs"); r.bold = True; r.font.size = Pt(16)
sp = d.add_paragraph(); sp.alignment = WD_ALIGN_PARAGRAPH.CENTER
sp.add_run("Companion brief to UQ_Crew_Safety_Logistics_Proposal_v0.5.docx\n").italic = True
sp.add_run(f"NASA Aerial Wildfire Project, ACERO Disaster Lab working group, {DATE}")
P("Numbers below come from the v0.5 proposal and its workbook, which name the record behind each one. "
  "Data availability was checked against the project folder (research_data/dataset_catalog.json, accessed "
  "30 September 2026) and is not a full survey of what exists.", italic=True, size=9)

# ------------------------------------------------------------------ 1
H("1. The picture in one paragraph")
P("On all four fires (Pioneer 2016, Carr 2018, Camp 2018, LNU 2020) people saw the fire within minutes, but the "
  "written record trailed it by hours to a day, and the gap was widest exactly when the fire grew fastest and crews "
  "were in danger. The official size was usually the previous night's perimeter. On the first day of the Camp Fire, "
  "the first ICS-209 at 18:30 gave 16,000 acres, after 15 of the 19 documented burnovers had already happened; CAL "
  "FIRE later put the first 12 hours at more than 55,000 acres. Long-endurance drones could close much of this "
  "information gap, but wind on the worst days exceeded the published limit of every drone in the agency fleet.")

# ------------------------------------------------------------------ 2
H("2. Key findings")
H("The first report and the egress clock", 2)
LB("People report fast. ", "First 911 calls, lead plane or air attack reports came within minutes of ignition; the "
   "first aircraft size estimate came within 0.4 to 1.5 hours where one is documented. That information lived on the radio.")
LB("Records report slowly. ", "The first mapped perimeter arrived 3.5 to 14 hours after the first report and the first "
   "ICS-209 5 to 12 hours after (50 hours for the LNU complex). Satellites saw Carr within about 1 hour and Camp within "
   "3.8 hours (MODIS), but delivery to users adds up to 4 more hours for VIIRS.")
LB("The ICS-209 discovery time is not the detection time. ", "It sits 1 to 8 hours after the first report, so it "
   "cannot be used as T_detect in the safety margin.")
LB("Warnings arrived with the fire, or failed. ", "On Camp, evacuation orders went out at about the time the fire "
   "arrived in each zone, 13 to 59 minutes before the early burnovers, and in Concow the fire arrived first. At "
   "Rattlesnake Flats an emergency call went unanswered because of radio traffic. On Carr, 62 minutes passed between "
   "the aerial warning to Branch III and the mayday.")
LB("Camp's first 85 minutes decided the day. ", "First 911 call 06:25; first burnovers 07:50; Paradise reached by "
   "08:01; nothing on record but radio until the first perimeter (stamped 09:54, time uncertain).")

H("When information arrived versus when the fire grew", 2)
LB("The record ran about one day behind. ", "The ICS-209 size equalled the previous night's perimeter on 48 of 56 "
   "days (Pioneer) and 20 of 28 (Carr). The national situation report repeated the previous evening's ICS-209 almost "
   "every day.")
LB("Most growth was unseen for over 12 hours. ", "59% to 91% of the final area on Carr, Camp and Pioneer grew while "
   "the newest perimeter was more than 12 hours old. Aircraft perimeters were a median 14.5 to 18.8 hours apart, with "
   "gaps up to 140 hours.")
LB("Worst mismatch: Camp, 8 November. ", "The same-day perimeter was 3.91 times the ICS-209 size. On Carr, 27 July, "
   "the ICS-209 gave 48,312 acres against a perimeter of 80,968.")
LB("Night matters. ", "37% of LNU growth happened between 20:00 and 06:00, when crewed aircraft do not fly, and only "
   "4% to 21% of aircraft perimeters carry a night time stamp.")
LB("The surveillance requirement is a 2 to 3 hour revisit, day and night. ", "On the hourly LNU record the last-known "
   "size falls short by 2.6% on average at a 2 hour revisit, 5.1% at 3 hours, 20.0% at 12 hours and 32.2% at 24 hours.")
LB("The records carry their own errors. ", "ICS-209-PLUS ACRES is capped (use CURR_INCIDENT_AREA); its growth date "
   "lags the run by one report (Camp's largest growth is dated 9 November, a day late); its linked satellite event is "
   "23% below Camp's final size; satellite area is biased per fire (+9% Carr, -2% Camp); some archive time stamps are "
   "processing times.")

H("Drones", 2)
LB("Today's agency drones fly short sorties. ", "USFS drones averaged 17 to 18 minutes per flight in FY22 (Week 5). "
   "Holding one Alta X station for 24 hours takes about 85 sorties.")
LB("Long-endurance fixed wing fits the requirement. ", "NWCG Type 1 fixed wing (6 to 14 hours; FVR-90 12 to 18 hours, "
   "Week 5) can hold the 2 to 3 hour revisit with 2 to 4 sorties and 3 pilot shifts a day. FVR-90 and SuperVolo have "
   "flown with NASA ACERO as radio relays.")
LB("Wind is the binding limit. ", "Camp's first morning (20 to 30 mph sustained, gusts over 40), Pioneer's and LNU's "
   "runs all exceeded the published limits of the agency fleet. Only Carr (up to 21 mph) was within the Alta X limit, "
   "and marginally. No wind limit for Type 1 fixed wing was found in project records.")
LB("Warning relay is the weakest link. ", "Relay message success was 52% with four ground nodes and over 90% with "
   "three; no drone has demonstrated crew tracking or alert delivery.")
LB("Water drop (M4) and logistics support (M5) are future only. ", "No library drone has flown either; Week 5 lists tests and "
   "demonstrators carrying 20 to 100 gallons per sortie, against 300 to 1,000 gallons for helicopters.")
LB("Night flight needs approval. ", "Night drone flights need a DOI/FAA agreement, and night BVLOS has not been shown "
   "in the library.")

# ------------------------------------------------------------------ 3
H("3. Key points of the proposal")
LB("Two decision numbers. ", "Crew safety: P(M < 0), where the safety margin M = T_fire - (T_detect + T_notify + "
   "T_egress). Logistics support (M5): P(on time), the chance a requested delivery of water, hose, pumps or supplies "
   "reaches the crew within its window. It answers the four Week 6 M5 questions, unchanged: payload (Q1), loaded range "
   "to crews far from roads (Q2), accurate and safe delivery (Q3), and enough fleet capacity to matter (Q4).", style="List Number")
LB("Five research questions. ", "Effect of 24/7 surveillance on P(M < 0) (RQ1); delivery reliability by day, night "
   "and smoke (RQ2); which inputs dominate (RQ3); how much fresher observation helps against the measured baseline "
   "(RQ4); how biased the records themselves are (RQ5).", style="List Number")
LB("Four new uncertain inputs. ", "U21 drone availability by wind; U22 time to first drone perimeter; U23 night and "
   "BVLOS authorisation; U24 warning relay success.", style="List Number")
LB("Validation runs. ", "Camp 8 November and Carr 26 July first, because both have timed crew hazards. Pioneer and "
   "LNU test spread and information age. ELMFIRE and GridFire ensembles at the observed cadence and at 6, 3, 2 and "
   "1 hour drone revisits, with and without the wind gate.", style="List Number")
LB("Honest limits. ", "Four fires, not a national rate; the revisit curve rests on one fire; Week 5 drone figures are "
   "agency and vendor claims; notification times rest on published reports because radio and dispatch logs are not "
   "in hand; novelty has not been assessed.", style="List Number")

# ------------------------------------------------------------------ 4
H("4. Data we still need for a complete study")
P("Each row names the model term it would fix. Priority 1 items are needed before the Camp and Carr validation runs; "
  "priority 2 completes the four-fire picture; priority 3 strengthens it.")

H("4.1 Radio and dispatch records (T_detect, T_notify)", 2)
T(["Data", "Why it matters", "Fires", "What we have", "How to get it", "Pri."], [
 ["Radio traffic audio and transcripts (dispatch, air-to-ground, air tactical, command and tactical nets)",
  "Times when fire position and warnings were spoken versus heard; failed calls (U9)",
  "All four", "None. Camp events quoted in NIST reports; Carr and Camp Green Sheets",
  "Pioneer: FOIA to USFS R4, draft ready (Fire Case Studies/Pioneer_2016). Camp, Carr: records request to CAL FIRE (California Public Records Act); Carr also NPS Whiskeytown (FOIA)", "1"],
 ["Dispatch logs and computer-aided dispatch record (e.g. WildCAD)",
  "True first report time, resource dispatch and arrival (T_detect, U17)",
  "All four", "First call times from published reports only", "Same requests as above (item 3 of the Pioneer draft)", "1"],
 ["ICS-214 activity logs (Communications Unit, Air Operations, Air Tactical) and ICS-205 radio plans",
  "Who knew what, when; which frequencies were saturated", "All four", "None", "Same requests (items 4 and 5 of the Pioneer draft)", "2"],
 ["Evacuation and crew notification times outside Camp",
  "Notification prior for Carr and LNU (U9)", "Carr, LNU", "Camp only (NIST NETTRA, minute level)",
  "County OES and sheriff alert logs; CAL FIRE after-action reviews", "2"],
], widths=[1.45, 1.35, 0.6, 1.25, 1.95, 0.35])

H("4.2 Night records (information age at night)", 2)
T(["Data", "Why it matters", "Fires", "What we have", "How to get it", "Pri."], [
 ["NIROPS night infrared perimeters and heat products",
  "The only night perimeters crews used; tests the night gap measured from GeoMAC", "All four",
  "None; behind FAMAuth", "Partner access through NASA or a NWCG account", "1"],
 ["Night operations logs: night shift ICS-214s, Incident Action Plan night operations, night aviation flight logs",
  "What was known and flown at night, especially LNU's 18 to 19 August overnight run", "LNU, Camp, Carr",
  "None", "Records requests above, scoped to night operational periods", "2"],
 ["GOES-16 fire detections (FDC), every 5 to 15 min, day and night",
  "Night detection and growth timing for Camp and LNU; first detection without a polar overpass", "Camp, Carr, LNU",
  "Camp and LNU days verified on public AWS; not downloaded. No GOES-R fire product before 2017 (Pioneer)", "Public download, noaa-goes16 bucket", "2"],
 ["FIRMS VIIRS 2020 and the LNU component-fire ICS-209s",
  "Completes LNU's first-report chain (MODIS +17.1 h is the only satellite time now)", "LNU",
  "MODIS only; VIIRS file skipped for size", "FIRMS free MAP_KEY; ICS-209-PLUS sitreps by component fire name", "2"],
 ["GOFER hourly reconstruction for Carr and Camp", "Hourly reference for the revisit curve on a second and third fire",
  "Carr, Camp", "LNU only (team supplement)", "Ask the GOFER authors whether Carr and Camp are covered", "2"],
], widths=[1.45, 1.35, 0.6, 1.25, 1.95, 0.35])

H("4.3 Wind and weather (U21 drone availability, fire spread)", 2)
T(["Data", "Why it matters", "Fires", "What we have", "How to get it", "Pri."], [
 ["Hourly wind and gusts at the critical hours, compared with drone limits",
  "Builds U21: could a drone have flown, hour by hour", "All four",
  "In hand, not yet analysed: ASOS hourly (two to three stations per fire), ERA5 hourly point, gridMET daily",
  "Run the comparison (next step in v0.5)", "1"],
 ["RAWS hourly from stations near each fire",
  "Wind near the fire line, not at a valley airport", "All four",
  "None; WRCC blocks pre-2026 hourly data and needs an access code", "Request a WRCC access code, or Synoptic API token (free tier terms not checked)", "1"],
 ["On-fire weather: spot forecasts, incident meteorologist observations, belt weather readings",
  "Wind the crews actually saw at the burnover sites", "Camp, Carr", "Wind values quoted in NIST and Green Sheets",
  "Records requests (ICS-214s, IAP weather pages); NWS spot forecast archive", "2"],
 ["Gridded hourly wind above the surface (HRRR)",
  "Drone limits apply at flight altitude; ASOS and ERA5 extracts are 10 m surface wind", "Pioneer, Carr, Camp, LNU",
  "Listing verified on public AWS; not downloaded", "Public download, only wind fields for a small box", "2"],
 ["Published wind limit for Type 1 fixed wing (FVR-90, ScanEagle class)",
  "Without it U21 is a 25 to 40 mph sensitivity range", "n/a", "Not found in project records",
  "Ask Overwatch Aero (L3Harris) or NASA ACERO", "1"],
], widths=[1.45, 1.35, 0.6, 1.25, 1.95, 0.35])

H("4.4 Aircraft, airspace and drones on the fire", 2)
T(["Data", "Why it matters", "Fires", "What we have", "How to get it", "Pri."], [
 ["Crewed aircraft flight logs and grounding times (aircraft tracking, lead plane and air attack logs)",
  "When air surveillance actually existed; Camp aircraft were grounded by wind and smoke", "All four",
  "Narrative only. NIFC Historical_Flight_Use_Concurrent service listed, not inspected", "Inspect the NIFC service; records requests", "2"],
 ["Drone flights during the four fires", "Real sortie lengths and night use on these incidents", "All four",
  "DOI UAS flight-use file in hand: a name search on 13 Oct 2026 found only post-fire mapping at Whiskeytown (6 to 27 Oct 2018), no flights during the fires. USFS and CAL FIRE drone logs not in hand",
  "Ask USFS and CAL FIRE aviation for incident UAS logs", "3"],
 ["Historical TFR records", "Airspace delay for U23", "All four", "None; no public archive found (one day's search, 30 Sep 2026)",
  "FAA request; DOI UAS file has a per-flight TFR flag", "3"],
 ["Night and BVLOS approval times; time to first drone on scene",
  "Priors for U22 and U23", "n/a", "None", "Stakeholder questionnaire (elicitation)", "2"],
], widths=[1.45, 1.35, 0.6, 1.25, 1.95, 0.35])

H("4.5 Crews, fuels and the fire record", 2)
T(["Data", "Why it matters", "Fires", "What we have", "How to get it", "Pri."], [
 ["Crew positions, escape routes and safety zones at each burnover", "T_egress per hazard, not a generic travel rate",
  "Camp, Carr", "Burnover locations from NIST and Green Sheets; travel rates from Sullivan 2020; crew GPS tracks not public",
  "Records requests; Campbell 2019 Escape Route Index (PDF via firescience.gov)", "1"],
 ["Fuels and terrain for the ensembles (LANDFIRE FBFM40, canopy; 3DEP)", "Required inputs for ELMFIRE and GridFire",
  "All four", "None (folder empty on purpose)", "LANDFIRE clip-and-ship (LF2016 for Camp, Carr, LNU); Pioneer LF2014 has no public URL", "1"],
 ["True time of Camp's first perimeter (stamped 09:54)", "Whether the first perimeter came at +3.5 h or much later",
  "Camp", "GeoMAC stamp only", "Compare with NIST progression maps", "1"],
 ["Firefighter entrapments on Pioneer and LNU", "Adds hazard cases beyond Camp and Carr", "Pioneer, LNU",
  "None found in one day of searching (30 Sep 2026)", "Wildland Fire Lessons Learned Center; agency safety reports", "3"],
], widths=[1.45, 1.35, 0.6, 1.25, 1.95, 0.35])

# ------------------------------------------------------------------ 5
H("5. Suggested order")
for lead, text in [
 ("This week, no outside help needed: ", "run the wind comparison (U21) with the ASOS and ERA5 data already in hand; check Camp's 09:54 perimeter against NIST maps; download LANDFIRE fuels for Camp and Carr."),
 ("Send now, because replies take weeks: ", "the Pioneer FOIA (draft ready), records requests for Camp and Carr radio, dispatch and ICS-214s, the WRCC access code, and the FVR-90 wind limit question."),
 ("Through partners: ", "NIROPS access, GOFER coverage of Carr and Camp, and the stakeholder questionnaire for U22 and U23."),
 ("Then: ", "GOES and VIIRS for night detection, HRRR winds aloft, and the entrapment search for Pioneer and LNU."),
]:
    LB(lead, text, style="List Number")

d.save(OUT)
print("wrote", OUT.name)
