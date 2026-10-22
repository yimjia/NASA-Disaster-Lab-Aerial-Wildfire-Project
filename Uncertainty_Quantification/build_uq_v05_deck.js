const pptxgen = require("pptxgenjs");
const { applyTheme } = require("/Users/vickylopez/.claude/skills/synced/e4c1bbb6-f014-42d0-a969-856eecea68c8_f69e8114-8d9f-4fc0-a4a9-97e023ca6068/pptx/scripts/apply_theme.js");

const FIG = "/Users/vickylopez/Desktop/NASA_Aerial_Wildfire_Project/Uncertainty Quantification/figures/";
const OUT = "UQ_v0.5_Findings_Presentation.pptx";

const THEME = {
  name: "Ember Charcoal",
  headFontFace: "Arial",
  bodyFontFace: "Calibri",
  colors: {
    dk1: "1D2125", lt1: "FFFFFF", dk2: "3A424B", lt2: "F1F2F4",
    accent1: "E8642F", // ember: the record (matches ICS-209 orange in the figures)
    accent2: "2A78D6", // aircraft perimeter blue
    accent3: "18A878", // satellite green
    accent4: "8E7BD0", // crew hazard purple
    accent5: "E9A821", // amber: partial
    accent6: "6B7480", // slate
    hlink: "2A78D6", folHlink: "6B7480",
  },
};
const HEX = THEME.colors;

const pres = new pptxgen();
pres.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pres.theme = { headFontFace: THEME.headFontFace, bodyFontFace: THEME.bodyFontFace };
pres.title = "UQ v0.5 Findings";
pres.author = "ACERO Disaster Lab working group";
pres.subject = "Uncertainty quantification for crew safety and aerial logistics";
const C = pres.SchemeColor;

const W = 13.333, M = 0.6;
const FOOT = "NASA Aerial Wildfire Project  |  ACERO Disaster Lab  |  UQ proposal v0.5";

// ---------- layouts ----------
pres.defineSlideMaster({
  title: "TITLE",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 2.0, w: 11.4, h: 1.9, fontSize: 40, bold: true, color: C.background1, valign: "bottom", fontFace: "Arial", align: "left" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: M, y: 4.1, w: 11.0, h: 0.9, fontSize: 20, color: C.background2, valign: "top" }, text: "" } },
  ],
});
pres.defineSlideMaster({
  title: "SECTION",
  background: { color: C.text1 },
  objects: [
    { placeholder: { options: { name: "title", type: "title", x: M, y: 2.2, w: 12.1, h: 1.6, fontSize: 40, bold: true, color: C.background1, valign: "bottom", align: "left" }, text: "" } },
    { placeholder: { options: { name: "body", type: "body", x: M, y: 3.95, w: 10.5, h: 1.2, fontSize: 18, color: C.background2, valign: "top" }, text: "" } },
    { text: { text: FOOT, options: { x: M, y: 7.0, w: 9, h: 0.3, fontSize: 10, color: C.accent6, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 7.0, w: 0.6, h: 0.3, fontSize: 10, color: C.accent6, align: "right" },
});
pres.defineSlideMaster({
  title: "CONTENT",
  background: { color: C.background1 },
  objects: [
    { placeholder: { options: { name: "kicker", type: "body", x: M, y: 0.35, w: 9, h: 0.3, fontSize: 12, bold: true, color: C.accent1, margin: 0, charSpacing: 2 }, text: "" } },
    { placeholder: { options: { name: "title", type: "title", x: M, y: 0.62, w: 12.1, h: 0.75, fontSize: 32, bold: true, color: C.text1, margin: 0, valign: "top", align: "left" }, text: "" } },
    { text: { text: FOOT, options: { x: M, y: 7.05, w: 9, h: 0.28, fontSize: 10, color: C.accent6, margin: 0 } } },
  ],
  slideNumber: { x: 12.2, y: 7.05, w: 0.55, h: 0.28, fontSize: 10, color: C.accent6, align: "right" },
});

// ---------- helpers ----------
let SECTION = "";
function section(title) { SECTION = title; pres.addSection({ title }); }
function content(kicker, title) {
  const s = pres.addSlide({ masterName: "CONTENT", sectionTitle: SECTION });
  s.addText(kicker.toUpperCase(), { placeholder: "kicker" });
  s.addText(title, { placeholder: "title" });
  return s;
}
function divider(title, sub) {
  const s = pres.addSlide({ masterName: "SECTION", sectionTitle: SECTION });
  s.addText(title, { placeholder: "title" });
  s.addText(sub, { placeholder: "body" });
  return s;
}
const T = (text, o = {}) => ({ text, options: o });
// card: soft tinted rounded rectangle
function card(s, x, y, w, h, fill = C.background2, name = "card") {
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h, fill: { color: fill }, line: { color: fill }, rectRadius: 0.08, objectName: name });
}
// chip: small rounded label, the deck motif
function chip(s, x, y, text, fill, color = C.background1, w = null) {
  const cw = w || Math.max(0.9, 0.11 * text.length + 0.35);
  s.addText(text.toUpperCase(), { x, y, w: cw, h: 0.3, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.15, fill: { color: fill }, color, fontSize: 10, bold: true, align: "center", valign: "middle", margin: 0, charSpacing: 1, isTextBox: true, objectName: "chip " + text });
  return cw;
}
function numCircle(s, x, y, n, fill = C.accent1, d = 0.5) {
  s.addText(String(n), { x, y, w: d, h: d, shape: pres.shapes.OVAL, fill: { color: fill }, color: C.background1, fontSize: 16, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true, objectName: "number " + n });
}
function stat(s, x, y, w, big, label, color = C.accent1, bigSize = 44) {
  s.addText(big, { x, y, w, h: 0.85, fontSize: bigSize, bold: true, color, fontFace: "Arial", margin: 0, valign: "bottom", isTextBox: true, objectName: "stat " + big });
  s.addText(label, { x, y: y + 0.9, w, h: 0.9, fontSize: 14, color: C.text2, margin: 0, valign: "top", isTextBox: true, objectName: "stat label" });
}
// equation runs with subscripts
function eqRuns(size, color) {
  const b = { fontSize: size, color, bold: true };
  const sub = { fontSize: size, color, bold: true, subscript: true };
  return [T("M = T", b), T("fire", sub), T("  −  ( T", b), T("detect", sub), T("  +  T", b), T("notify", sub), T("  +  T", b), T("egress", sub), T(" )", b)];
}
const AX = {
  catAxisLabelColor: HEX.dk2, valAxisLabelColor: HEX.accent6, catAxisLabelFontFace: "+mn-lt", valAxisLabelFontFace: "+mn-lt",
  catAxisLabelFontSize: 12, valAxisLabelFontSize: 11, dataLabelFontFace: "+mn-lt", dataLabelFontSize: 12, dataLabelColor: HEX.dk1,
  titleFontFace: "+mn-lt", titleFontSize: 14, titleColor: HEX.dk1, legendFontFace: "+mn-lt", legendFontSize: 12, legendColor: HEX.dk2,
  catAxisLineShow: false, valAxisLineShow: false,
};
function grid() { return { color: "E1E4E8", size: 0.75 }; }

// =====================================================================
section("Opening");
{
  const s = pres.addSlide({ masterName: "TITLE", sectionTitle: SECTION });
  s.addText("The Fire Outran the Record", { placeholder: "title" });
  s.addText("Uncertainty quantification for crew safety and aerial logistics: what four fires show, and where 24/7 drones close the gap", { placeholder: "body" });
  chip(s, M, 1.45, "UQ proposal v0.5 findings", C.accent1);
  s.addText("NASA Aerial Wildfire Project  |  ACERO Disaster Lab working group  |  October 2026", { x: M, y: 6.6, w: 11, h: 0.4, fontSize: 14, color: C.accent6, margin: 0, isTextBox: true });
  s.addNotes("Title. This deck walks through the v0.5 uncertainty quantification proposal: what happened on Pioneer 2016, Carr 2018, Camp 2018 and the LNU Lightning Complex 2020, how fast information got out, where the record lagged the fire, which drones could close those gaps, and how all of that enters the uncertainty model. It ends with what we are missing, next steps, and suggestions to strengthen the method. Draft for team discussion; novelty has not been assessed.");
}

{
  const s = content("Bottom line", "Decision makers read a number about one day old");
  const y = 1.75, w = 2.8, g = 0.27;
  stat(s, M, y, w, "16,000 ac", "First Camp ICS-209 at 18:30, day 1. CAL FIRE later put the first 12 h at more than 55,000 acres", C.accent1, 40);
  stat(s, M + (w + g), y, w, "15 of 19", "Camp burnovers that had already happened when that first ICS-209 was filed", C.accent4, 40);
  stat(s, M + 2 * (w + g), y, w, "59 to 91%", "Share of final area on Carr, Camp and Pioneer that grew while the newest perimeter was over 12 h old", C.accent2, 40);
  stat(s, M + 3 * (w + g), y, w, "2 to 3 h", "Revisit that keeps last-known size within about 3 to 5% during growth, day and night", C.accent3, 40);
  card(s, M, 4.35, 12.1, 2.35, C.background2, "summary card");
  s.addText([
    T("People detected each fire within minutes; the records took hours. ", { bold: true }),
    T("First mapped perimeter 3.5 to 14 h after ignition, first ICS-209 5 to 12 h (50 h for LNU).", { breakLine: true }),
    T("Long-endurance drones can, in principle, hold the revisit around the clock. ", { bold: true }),
    T("Small agency drones average about 17 minutes a sortie and cannot.", { breakLine: true }),
    T("Wind is the binding limit. ", { bold: true }),
    T("On Camp's first morning and on Pioneer's and LNU's critical days, wind exceeded the published limit of every drone in the agency fleet. So drone availability enters the model as an uncertain input conditioned on wind and time of day."),
  ], { x: M + 0.3, y: 4.5, w: 11.5, h: 2.05, fontSize: 16, color: C.text1, valign: "middle", paraSpaceAfter: 8, margin: 0, isTextBox: true });
  s.addNotes("Bottom line up front. On all four fires the fire outran the record. The official ICS-209 size was usually the previous night's perimeter, so it was about one day old, and on blowup days it was far too small. The clearest case is Camp: the first ICS-209 at 18:30 on 8 November said 16,000 acres; CAL FIRE later estimated more than 55,000 acres in the first 12 hours, and 15 of the 19 documented burnovers had already happened. Most growth on Carr, Camp and Pioneer (59 to 91 percent) happened while the newest perimeter was more than 12 hours old. The hourly LNU record shows a 2 to 3 hour revisit around the clock keeps the last-known size within about 3 to 5 percent during growth, against 20 to 32 percent at the 12 to 24 hour cadence of the records. Type 1 fixed-wing drones can meet that cadence; small drones cannot; and wind on the worst days grounds the current fleet.");
}

{
  const s = content("What we are quantifying", "Two decision numbers, five research questions");
  card(s, M, 1.65, 6.9, 2.45, C.background2, "M3 card");
  chip(s, M + 0.3, 1.85, "M3 crew safety", C.accent1);
  s.addText(eqRuns(22, C.text1), { x: M + 0.3, y: 2.3, w: 6.4, h: 0.6, margin: 0, isTextBox: true });
  s.addText("Time between fire arrival and a crew's completed escape. Headline: P(M < 0), the chance the fire gets there first.", { x: M + 0.3, y: 3.0, w: 6.4, h: 0.9, fontSize: 15, color: C.text2, margin: 0, isTextBox: true });
  card(s, M, 4.3, 6.9, 2.45, C.background2, "M5 card");
  chip(s, M + 0.3, 4.5, "M5 logistics support", C.accent2);
  s.addText("P(on time)", { x: M + 0.3, y: 4.92, w: 6.4, h: 0.55, fontSize: 24, bold: true, color: C.text1, margin: 0, isTextBox: true });
  s.addText("Chance a requested delivery of water, hose, pumps or supplies reaches crews within its window. Built from the Week 6 M5 questions, unchanged: Q1 payload, Q2 loaded range, Q3 delivery accuracy, Q4 fleet throughput.", { x: M + 0.3, y: 5.5, w: 6.4, h: 1.15, fontSize: 14, color: C.text2, margin: 0, isTextBox: true });
  const rq = [
    ["RQ1", "How does P(M < 0) change with and without 24/7 aerial surveillance?"],
    ["RQ2", "What is P(on time) by day, night and smoke, at real incident demand?"],
    ["RQ3", "Which inputs dominate the uncertainty?"],
    ["RQ4", "How much does fresher observation shrink it, against the measured baseline?"],
    ["RQ5", "How large and biased are the records, so reporting error is not read as model error?"],
  ];
  rq.forEach(([k, v], i) => {
    const y = 1.7 + i * 1.02;
    s.addText(k, { x: 7.85, y, w: 0.8, h: 0.42, shape: pres.shapes.ROUNDED_RECTANGLE, rectRadius: 0.1, fill: { color: C.text1 }, color: C.background1, fontSize: 13, bold: true, align: "center", valign: "middle", margin: 0, isTextBox: true });
    s.addText(v, { x: 8.85, y: y - 0.08, w: 3.9, h: 0.9, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  });
  s.addNotes("We already know the mission, so this is only the frame for the numbers. Crew safety is the margin M: time of fire arrival minus the sum of detection, notification and egress time. We report P(M < 0), the probability the fire wins. Logistics support is P(on time), built directly from the four Week 6 M5 questions. Five research questions: RQ1 the effect of 24/7 surveillance on P(M < 0); RQ2 delivery reliability by day, night and smoke; RQ3 which inputs dominate; RQ4 how much fresher observation helps against the baseline measured in this work; RQ5 how biased the records themselves are.");
}

// =====================================================================
section("What happened on four fires");
divider("What happened on four fires", "Pioneer 2016, Carr 2018, Camp 2018 and the LNU Lightning Complex 2020, read from ICS-209-PLUS, NIFC GeoMAC, NASA FEDS, GOFER and FIRMS");

{
  const s = content("Figure 1", "The fire versus the record, on all four fires");
  const ih = 5.45, iw = ih * 1584 / 1144;
  s.addImage({ path: FIG + "fig1_size_on_record_vs_fire.png", x: M, y: 1.5, w: iw, h: ih, altText: "Figure 1: fire size on record against the fire for Pioneer, Carr, Camp and LNU", objectName: "figure 1" });
  const x = M + iw + 0.35, w = W - M - x;
  s.addText("How to read it", { x, y: 1.55, w, h: 0.4, fontSize: 18, bold: true, color: C.text1, margin: 0, isTextBox: true });
  const keys = [[C.accent1, "Orange step", "ICS-209 size on record at the end of each report period"], [C.accent2, "Blue dots", "Aircraft perimeters (NIFC GeoMAC)"], [C.accent3, "Green line", "Satellite area (FEDS 12-hourly VIIRS; GOFER hourly GOES for LNU)"], [C.accent4, "Purple band", "Documented crew hazard windows"]];
  keys.forEach(([col, k, v], i) => {
    const y = 2.05 + i * 0.62;
    s.addShape(pres.shapes.OVAL, { x, y: y + 0.07, w: 0.22, h: 0.22, fill: { color: col }, line: { color: col }, objectName: "key " + k });
    s.addText([T(k + ". ", { bold: true }), T(v)], { x: x + 0.35, y, w: w - 0.35, h: 0.58, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  });
  card(s, x, 4.6, w, 2.25, C.background2, "takeaway");
  s.addText([
    T("The gap between orange and the others is the information gap. ", { bold: true }),
    T("It is widest exactly when growth is fastest, and the Carr and Camp crew hazards fall inside it. Pioneer, a slow timber fire, stays within about a day."),
  ], { x: x + 0.2, y: 4.72, w: w - 0.4, h: 2.0, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  s.addNotes("Figure 1 plots thousand acres against days since ignition for each fire. The orange step line is what was on record: the ICS-209 size as reported (CURR_INCIDENT_AREA) at the end of each report period. Blue dots are aircraft perimeters from NIFC GeoMAC. The green line is satellite area: FEDS 12-hourly VIIRS for Carr and Camp, and GOFER hourly GOES for LNU. The dips in the LNU green line are reconstruction artifacts, not real shrinkage. Purple bands are documented crew hazard windows: the 19 NIST burnovers on Camp and the 26 July evening run on Carr. How to read it: the vertical distance between the orange step and the blue or green is how far the record trailed the fire. On Pioneer, a long timber fire, the record kept within about a day. On Carr, Camp and LNU, wind-driven runs put the fire tens of thousands of acres ahead of the official size for one to two days, and the crew hazards fall inside those windows. Caveat: satellite area reads high on Carr and LNU (per-fire bias, see the records slide).");
}

{
  const s = content("Fire by fire", "Same pattern, four very different fires");
  const fires = [
    ["Pioneer 2016", "Boise NF, Idaho  |  188,404 ac", C.accent6, ["Lead plane reported 1 to 1.5 ac at 17:17, 18 July; first perimeter 13.7 h later", "Slow for a month, then about 29,000 ac in one run on 29 Aug, with pyroconvective updrafts", "91% of growth unseen for over 12 h; one 140 h gap with no perimeter", "No entrapment found"]],
    ["Carr 2018", "Shasta and Whiskeytown, CA  |  229,651 ac", C.accent2, ["Ignition 13:09, calls within 2 min, air attack 20 ac at 13:32", "26 July fire whirl: warning to Branch III 18:38, mayday 19:40, 62 min later", "Newest ICS-209 then was 10.5 h old at 20,000 ac; perimeter 44,484, satellite about 57,700", "Next day: 48,312 ac on record against a perimeter of 80,968"]],
    ["Camp 2018", "Butte County, CA  |  153,336 ac, 85 deaths", C.accent1, ["First 911 call 06:25, first engine 06:44 in about 35 mph wind", "First burnovers 07:50, 85 min after the call; Paradise by 08:01", "First ICS-209 at 18:30: 16,000 ac", "Same-day perimeter was 3.91 times the ICS-209 size"]],
    ["LNU Complex 2020", "Napa, Sonoma, Lake, Solano, Yolo  |  363,220 ac", C.accent3, ["Over 2,000 lightning strikes on 16 to 17 Aug; start times disagree between records", "Fires merged overnight 18 to 19 Aug with a 12 mile run", "First complex ICS-209 50 h after start, already 33,025 ac", "37% of growth between 20:00 and 06:00; no entrapment found"]],
  ];
  const w = 2.85, g = 0.233;
  fires.forEach(([name, meta, col, pts], i) => {
    const x = M + i * (w + g);
    card(s, x, 1.6, w, 5.2, C.background2, "fire " + name);
    s.addShape(pres.shapes.OVAL, { x: x + 0.25, y: 1.85, w: 0.32, h: 0.32, fill: { color: col }, line: { color: col }, objectName: "dot " + name });
    s.addText(name, { x: x + 0.7, y: 1.8, w: w - 0.85, h: 0.42, fontSize: 18, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
    s.addText(meta, { x: x + 0.25, y: 2.3, w: w - 0.45, h: 0.55, fontSize: 12, color: C.accent6, margin: 0, valign: "top", isTextBox: true });
    s.addText(pts.map((p, j) => T(p, { bullet: { indent: 12 }, breakLine: j < pts.length - 1 })), { x: x + 0.2, y: 2.95, w: w - 0.4, h: 3.7, fontSize: 13, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 7, isTextBox: true });
  });
  s.addNotes("One card per fire. Pioneer: reported by the Casner Fire lead plane, poor access, first perimeter 13.7 hours after the first report; slow for a month then a 29,000 acre run on 29 August; 91 percent of its area grew while the newest perimeter was over 12 hours old and the archive has a 140 hour gap. Carr: calls within two minutes of ignition, air attack sized it at 20 acres within 23 minutes; on 26 July the air tactical supervisor warned Branch III to leave at 18:38 and the mayday came 62 minutes later; the newest ICS-209 at that moment was 10.5 hours old and said 20,000 acres while the satellite showed about 57,700. Camp: burnovers began 85 minutes after the first 911 call; the first ICS-209 came 12 hours in at 16,000 acres. LNU: a lightning complex with conflicting start times, an overnight 12 mile run, and its first complex ICS-209 50 hours after the start. 37 percent of LNU growth happened at night. No entrapments were found for Pioneer or LNU in one day of searching.");
}

{
  const s = content("Figure 2", "Camp Fire: the first 36 hours");
  const iw = 8.3, ih = iw / 2;
  s.addImage({ path: FIG + "fig2_camp_first_36h.png", x: M, y: 1.55, w: iw, h: ih, altText: "Figure 2: Camp Fire first 36 hours, record against aircraft, satellite and radio estimates", objectName: "figure 2" });
  s.addText("Hours since first 911 call (06:25 PST, 8 Nov 2018). Purple: burnover windows (NIST TN 2135 p.197; CAL FIRE Green Sheet).", { x: M, y: 1.6 + ih, w: iw, h: 0.5, fontSize: 11, italic: true, color: C.accent6, margin: 0, isTextBox: true });
  const x = M + iw + 0.35, w = W - M - x;
  const pts = [
    ["Radio knew first", "Black diamonds: engine and air attack estimates, 200 to 2,000 ac by 08:01, lived only on the radio"],
    ["Hazards came before the record", "The early burnover window and Rattlesnake Flats both fall before any ICS-209 existed"],
    ["The record lagged a full day", "Orange reached 70,000 ac only at 07:45 on 9 Nov, a day after the run"],
    ["One time stamp is uncertain", "The first perimeter is stamped 09:54; its size matches the 12 h estimate, so it may be a processing time"],
  ];
  pts.forEach(([h, b], i) => {
    const y = 1.6 + i * 1.32;
    numCircle(s, x, y, i + 1, i === 3 ? C.accent5 : C.accent1, 0.42);
    s.addText([T(h, { bold: true, breakLine: true }), T(b, { color: C.text2 })], { x: x + 0.55, y: y - 0.04, w: w - 0.55, h: 1.25, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  });
  s.addNotes("Figure 2 zooms in on Camp's first 36 hours, measured from the first 911 call at 06:25. Orange is the ICS-209 size on record: zero until the first report at 18:30 (16,000 acres, 18,000 at 19:20), then 70,000 acres only at 07:45 the next morning. Blue is GeoMAC aircraft perimeters, starting at about 54,600 acres. Green squares are FEDS satellite area. Black diamonds are ground and air attack radio estimates, which are the only information in the first 90 minutes. Purple bands are burnover windows: the first burnovers at 07:50, 14 by 13:10, and Rattlesnake Flats in the afternoon. The point: every documented crew hazard on day one happened before the official record said anything. Two cautions. The first GeoMAC perimeter is stamped 09:54, but its size matches the Green Sheet's first-12-hour estimate, so the stamp may be a processing time, not the flight time; we need to check it against NIST maps. And the figure's Rattlesnake Flats label reads 14:45, which is the wind shift; the Green Sheet logs the burnover at 15:15.");
}

{
  const s = content("Camp timeline", "Every early hazard happened while only the radio knew");
  const rows = [
    ["06:25 to 06:44", "First 911 call 06:25; dispatch 06:31; first engine in sight 06:44, about 35 mph wind", "Nothing. ICS-209 discovery time later set to 08:29"],
    ["06:55 to 07:37", "Pulga evacuation order 06:55, 3 min before fire arrived; engine reports 200 to 300 ac at 07:10; Concow order 07:37", "Radio only"],
    ["07:50 to 08:01", "First burnovers (Hoffman Rd, Concow Rd) 07:50; air attack about 1,000 ac at 07:52, 'now in Paradise, about 2,000 ac' at 08:01", "Radio only"],
    ["08:30 to 13:10", "Burnovers 3 to 14 in Paradise and Concow; nearly all of Paradise and Magalia under orders by 09:03", "First aircraft perimeter (stamped 09:54, 54,586 ac); first satellite detection 10:14 (MODIS)"],
    ["14:15 to 15:15", "Rattlesnake Flats: firing operation planned 14:15, wind shift 14:45, burnover logged 15:15; an emergency call went unanswered in radio traffic", "Newest perimeter 4.3 to 5.4 h old; no ICS-209 yet"],
    ["18:30", "CAL FIRE later estimates more than 55,000 ac in the first 12 h", "First ICS-209: 16,000 ac (18,000 at 19:20)"],
    ["9 Nov 07:45", "", "ICS-209: 70,000 ac. ICS-209-PLUS dates the largest growth (52,000 ac) to this report, a day after the run"],
  ];
  const hdr = ["Time, 8 Nov", "What happened", "What was on record"].map((t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 } } }));
  const body = rows.map((r, i) => r.map((t, j) => ({ text: t, options: { fill: { color: i % 2 ? C.background1 : C.background2 }, bold: j === 0, color: j === 2 && /^(Radio only|Nothing)/.test(t) ? C.accent1 : C.text1 } })));
  s.addTable([hdr, ...body], { x: M, y: 1.6, w: 12.1, colW: [1.7, 6.0, 4.4], fontSize: 12, fontFace: "Calibri", valign: "middle", border: { type: "solid", pt: 0.5, color: "E1E4E8" }, margin: [4, 8, 4, 8], rowH: 0.62, objectName: "camp timeline table" });
  s.addNotes("This is the Camp time-area table built from ICS-209-PLUS, GeoMAC, FEDS and the NIST and CAL FIRE timelines (128 rows, in the team supplement format: perimeter_supplements/Camp_2018_Perimeter_Time_Area_ICS209PLUS.xlsx). Read left to right: what was happening on the ground versus what any record held. For the first 90 minutes the only information was on the radio, and the first burnovers came at 07:50. Two lessons for the model. First, the derived growth fields in ICS-209-PLUS inherit the reporting lag: it dates Camp's largest growth to 9 November, a day after the run, so those fields cannot time a blowup. Second, the linked FIRED satellite event gives 117,968 acres, 23 percent below the official final size, and places ignition about 17 km west of the origin, so satellite products need a per-fire bias term (U14), and their ignition points are not detection times.");
}

// =====================================================================
section("How fast information got out");
divider("How fast the first information got out", "Detection and notification are the first two terms of the safety margin");

{
  const s = content("First-report chain", "People detect in minutes; records lag by hours");
  const labels = ["Pioneer 2016", "Carr 2018", "Camp 2018"];
  s.addChart(pres.charts.BAR, [
    { name: "First aircraft estimate (radio)", labels, values: [0, 0.4, 1.5] },
    { name: "First satellite detection (MODIS)", labels, values: [6.3, 1.0, 3.8] },
    { name: "First ICS-209", labels, values: [7.5, 5.3, 12.1] },
    { name: "First aircraft perimeter", labels, values: [13.7, 14.2, 3.5] },
  ], {
    x: M, y: 1.55, w: 7.6, h: 5.3, barDir: "col", barGrouping: "clustered", barGapWidthPct: 60,
    chartColors: [HEX.dk1, HEX.accent3, HEX.accent1, HEX.accent2],
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0.0",
    showTitle: true, title: "Hours from first documented start to first entry", showLegend: true, legendPos: "b",
    valAxisTitle: "Hours", showValAxisTitle: true, valAxisTitleFontSize: 11, valAxisTitleColor: HEX.accent6, valAxisTitleFontFace: "+mn-lt",
    valGridLine: grid(), catGridLine: { style: "none" }, valAxisMaxVal: 16, valAxisMajorUnit: 4, ...AX,
    objectName: "first report chart",
  });
  const x = 8.55, w = W - M - x;
  card(s, x, 1.6, w, 2.05, C.background2, "LNU card");
  s.addText([T("LNU 2020, not charted", { bold: true, breakLine: true }), T("First MODIS detection +17.1 h; GOFER hourly area +3.0 h; first complex ICS-209 +50.0 h at 33,025 ac. VIIRS not yet downloaded.", { color: C.text2 })], { x: x + 0.2, y: 1.72, w: w - 0.4, h: 1.85, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  card(s, x, 3.85, w, 3.0, C.background2, "discovery card");
  chip(s, x + 0.2, 4.0, "U17", C.accent1);
  s.addText([T("Never use the ICS-209 discovery time as T detect. ", { bold: true }), T("It sits 0.9 to 8 h after the first report (Pioneer +0.9, Carr +2.1, Camp +2.1, LNU +8.0). Satellite times are acquisition times; delivery adds up to 4 h for VIIRS.", { color: C.text2 })], { x: x + 0.2, y: 4.42, w: w - 0.4, h: 2.35, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  s.addNotes("How to read this chart: each cluster is a fire; each bar is the hours from the earliest documented start to the first entry in one information source. Black is the first aircraft size estimate, which lived on the radio: zero for Pioneer because the lead plane both reported and sized it, 0.4 hours on Carr, 1.5 on Camp. Green is the first satellite detection from the FIRMS MODIS archive, using the first overpass with a detection within 15 km of the origin. Orange is the first ICS-209. Blue is the first aircraft perimeter. The pattern: people and radio are fast, records are slow. Camp's perimeter bar is short (3.5 hours) only because of the uncertain 09:54 stamp. LNU is left off the chart because its 50 hour first ICS-209 would flatten the others, and it is a complex whose component fires may have reported separately. Key modelling consequence: the ICS-209 discovery time is 1 to 8 hours after the first report, so it can never stand in for T detect. Starts used: Pioneer 17:17, Carr 13:09, Camp 06:25, LNU 06:00. One VIIRS pixel 6 minutes after the Carr ignition falls in the search radius but is not confirmed as this fire.");
}

{
  const s = content("Notification", "Warnings arrived with the fire, or failed outright");
  const y = 1.75, w = 2.8, g = 0.27;
  stat(s, M, y, w, "13 to 59 min", "How far Camp's evacuation orders preceded the early burnovers (NIST NETTRA, Table 12)", C.accent1, 30);
  stat(s, M + (w + g), y, w, "Fire first", "In Concow the fire arrived before the warning did", C.accent1, 30);
  stat(s, M + 2 * (w + g), y, w, "Unanswered", "Rattlesnake Flats emergency call lost in radio traffic (Camp Green Sheet)", C.accent4, 30);
  stat(s, M + 3 * (w + g), y, w, "62 min", "Carr: aerial warning to Branch III at 18:38, mayday at 19:40", C.accent2, 30);
  card(s, M, 4.3, 12.1, 2.45, C.background2, "model card");
  chip(s, M + 0.3, 4.5, "What this means for the model", C.text1);
  s.addText([
    T("U9, T notify: ", { bold: true }), T("priors come from these events, and a notification can fail outright, not only arrive late.", { breakLine: true }),
    T("Why it failed: ", { bold: true }), T("NIST notes spread information came in from 911 calls because firefighters were busy saving lives.", { breakLine: true }),
    T("U8, T egress: ", { bold: true }), T("crew travel time uses published GPS travel rates (Sullivan 2020; Campbell 2017, 2019)."),
  ], { x: M + 0.3, y: 4.95, w: 11.5, h: 1.7, fontSize: 15, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 6, isTextBox: true });
  s.addNotes("Notification is the second link in the margin. On Camp, between 07:30 and 08:40 evacuation orders went out at about the time the fire arrived in each zone; in Concow the fire arrived before the warning; and the orders preceded the early burnovers by only 13 to 59 minutes (NIST TN 2252, pp.97 to 101). At Rattlesnake Flats an emergency call went unanswered because of radio traffic. On Carr, 62 minutes passed between the aerial warning and the mayday. These set the priors for T notify (U9) and, importantly, show that notification can fail outright. Egress time uses published GPS travel-rate models. All of these rest on published reports; radio and dispatch logs are not yet in hand.");
}

// =====================================================================
section("The information gap");
divider("When information arrived versus when the fire grew", "The measured baseline that drones have to beat (RQ4)");

{
  const s = content("Information gap, measured", "The record ran about one day behind the fire");
  const hdr = ["Measure", "Pioneer", "Carr", "Camp", "LNU", "Record"].map((t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 } } }));
  const rows = [
    ["ICS-209 size equals previous night's perimeter", "48 of 56 days", "20 of 28", "4 of 13", "no GeoMAC", "ICS-209-PLUS, GeoMAC"],
    ["National report repeats previous evening's 209", "71 of 72 days", "37 of 38", "6 of 10", "28 of 28", "NICC IMSR"],
    ["Worst same-day perimeter / ICS-209 ratio", "1.12", "1.68 (27 Jul)", "3.91 (8 Nov)", "2.07 (satellite)", "GeoMAC, FEDS"],
    ["Median hours between aircraft perimeters (max)", "18.8 (140)", "17.9 (76)", "14.5 (93)", "none", "GeoMAC"],
    ["Share of final area grown while newest perimeter > 12 h old", "91%", "59%", "71%", "0% (hourly GOFER)", "gap analysis"],
    ["Share of aircraft perimeters with a night time stamp", "21%", "10%", "4%", "n/a", "GeoMAC"],
    ["Top three growth days, share of final area", "32%", "34%", "65%", "55%", "ICS-209-PLUS"],
  ];
  const body = rows.map((r, i) => r.map((t, j) => ({ text: t, options: { fill: { color: i % 2 ? C.background1 : C.background2 }, bold: (i === 2 && j === 3), color: (i === 2 && j === 3) ? C.accent1 : (j === 5 ? C.accent6 : C.text1), align: j === 0 ? "left" : "center" } })));
  s.addTable([hdr, ...body], { x: M, y: 1.6, w: 12.1, colW: [4.3, 1.45, 1.45, 1.45, 1.65, 1.8], fontSize: 13, fontFace: "Calibri", valign: "middle", border: { type: "solid", pt: 0.5, color: "E1E4E8" }, margin: [4, 8, 4, 8], rowH: 0.6, objectName: "gap measures table" });
  s.addText([T("Read the Camp ratio: ", { bold: true }), T("on 8 November the same-day aircraft perimeter was 3.91 times the size on record. Decision makers were planning against about a quarter of the fire.")], { x: M, y: 6.5, w: 12.1, h: 0.45, fontSize: 14, color: C.text1, margin: 0, isTextBox: true });
  s.addNotes("This is the v0.4 evaluation summarised: the measured baseline. Row 1: on most Pioneer and Carr days the ICS-209 size was just the previous night's perimeter. Row 2: the national situation report repeated the previous evening's ICS-209 almost every day on all four fires. Row 3: the worst same-day ratio of perimeter to ICS-209 size, 3.91 on Camp, 1.68 on Carr, 2.07 on LNU using satellite area. Row 4: aircraft perimeters were a median 14.5 to 18.8 hours apart, with gaps as long as 140 hours. Row 5: the share of final area that grew while the newest perimeter was more than 12 hours old; LNU is 0 percent only because GOFER is hourly. Row 6: very few aircraft perimeters carry a night time stamp. Row 7: growth is concentrated, the top three days carry a third to two thirds of the final area, which is why missing a single day matters.");
}

{
  const s = content("Growth and night", "Most growth happened while nobody had a fresh map");
  s.addChart(pres.charts.BAR, [{ name: "Share of final area", labels: ["Pioneer 2016", "Camp 2018", "Carr 2018"], values: [91, 71, 59] }], {
    x: M, y: 1.55, w: 5.9, h: 4.3, barDir: "bar", chartColors: [HEX.accent1], barGapWidthPct: 45,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"', showLegend: false,
    showTitle: true, title: "Grown while the newest perimeter was over 12 h old", catAxisOrientation: "maxMin",
    valAxisMaxVal: 100, valAxisMinVal: 0, valAxisMajorUnit: 25, valAxisLabelFormatCode: '0"%"', valGridLine: grid(), catGridLine: { style: "none" }, ...AX,
    objectName: "growth unseen chart",
  });
  s.addChart(pres.charts.BAR, [{ name: "Night-stamped perimeters", labels: ["Pioneer 2016", "Carr 2018", "Camp 2018"], values: [21, 10, 4] }], {
    x: 6.85, y: 1.55, w: 5.9, h: 4.3, barDir: "bar", chartColors: [HEX.accent2], barGapWidthPct: 45,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0"%"', showLegend: false,
    showTitle: true, title: "Aircraft perimeters with a night time stamp", catAxisOrientation: "maxMin",
    valAxisMaxVal: 100, valAxisMinVal: 0, valAxisMajorUnit: 25, valAxisLabelFormatCode: '0"%"', valGridLine: grid(), catGridLine: { style: "none" }, ...AX,
    objectName: "night perimeter chart",
  });
  card(s, M, 6.05, 12.1, 0.85, C.background2, "night card");
  s.addText([T("37% ", { bold: true, color: C.accent1, fontSize: 20 }), T("of LNU growth happened between 20:00 and 06:00, when crewed aircraft do not fly. Crewed aircraft can work about 30% of a 24 h cycle (ACERO ConOps pp.30-31).")], { x: M + 0.3, y: 6.1, w: 11.5, h: 0.75, fontSize: 15, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  s.addNotes("Left chart: for each fire with an aircraft perimeter series, the share of final area that grew while the newest perimeter was more than 12 hours old. Pioneer 91 percent, Camp 71, Carr 59. LNU is not shown because its reference is the hourly GOFER reconstruction, so by construction nothing is older than an hour. Right chart, on the same 0 to 100 percent scale: the share of aircraft perimeters with a night time stamp, only 4 to 21 percent. Together: growth happens when maps are stale, and maps are almost never made at night. The LNU figure ties it to the mission: 37 percent of LNU growth was at night, and the ConOps puts crewed aircraft at about 30 percent of the day.");
}

{
  const s = content("The surveillance requirement", "Revisit every 2 to 3 hours, around the clock");
  const labels = ["2 h", "3 h", "6 h", "12 h", "24 h"];
  s.addChart(pres.charts.BAR, [
    { name: "Mean shortfall", labels, values: [2.6, 5.1, 10.9, 20.0, 32.2] },
    { name: "90th percentile", labels, values: [5.6, 16.1, 39.5, 62.3, 86.4] },
  ], {
    x: M, y: 1.55, w: 7.9, h: 5.3, barDir: "col", barGrouping: "clustered", barGapWidthPct: 55,
    chartColors: [HEX.accent1, "F4B79C"], showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: '0.0"%"',
    showTitle: true, title: "Shortfall of last-known size against the hourly reference, by revisit interval", showLegend: true, legendPos: "b",
    valAxisMaxVal: 100, valAxisMajorUnit: 25, valAxisLabelFormatCode: '0"%"', valGridLine: grid(), catGridLine: { style: "none" }, ...AX,
    catAxisTitle: "Revisit interval", showCatAxisTitle: true, catAxisTitleFontSize: 11, catAxisTitleColor: HEX.accent6, catAxisTitleFontFace: "+mn-lt",
    objectName: "revisit chart",
  });
  const x = 8.85, w = W - M - x;
  s.addText("How it was measured", { x, y: 1.6, w, h: 0.4, fontSize: 18, bold: true, color: C.text1, margin: 0, isTextBox: true });
  s.addText("Sample the hourly LNU GOFER record at a fixed revisit interval, then compare the last-known size with the reference every hour, over the growth phase, 17 to 21 August.", { x, y: 2.05, w, h: 1.4, fontSize: 14, color: C.text2, margin: 0, valign: "top", isTextBox: true });
  card(s, x, 3.6, w, 1.45, C.background2, "result");
  s.addText([T("2 to 3 h: ", { bold: true, color: C.accent1 }), T("mean shortfall 2.6 to 5.1%. The records' 12 to 24 h cadence: 20 to 32%.")], { x: x + 0.2, y: 3.7, w: w - 0.4, h: 1.25, fontSize: 14, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  chip(s, x, 5.3, "Uncertainty: U18", C.accent5, C.text1);
  s.addText("One fire, 2 km GOES pixels. Measures how fast size goes stale, not shape accuracy.", { x, y: 5.7, w, h: 1.0, fontSize: 13, italic: true, color: C.accent6, margin: 0, valign: "top", isTextBox: true });
  s.addNotes("This is the most important chart for the drone question. Using the hourly LNU GOFER record as truth, we pretend we can only see the fire every N hours, carry the last-known size forward, and compare it with the reference every hour during the growth phase, 17 to 21 August. Dark bars are the mean shortfall; light bars are the 90th percentile, the bad hours. At a 2 hour revisit the last-known size is 2.6 percent short on average and 5.6 percent at the 90th percentile. At 24 hours, the cadence of the official record, it is 32 percent short on average and 86 percent short in the worst tenth of hours. The curve bends sharply between 3 and 6 hours, which is why the requirement is set at 2 to 3 hours, day and night, from the first hours. Limits: it is one fire, the pixels are 2 km, and it measures staleness of size, not where the edge is. This curve becomes input U18.");
}

{
  const s = content("Observation error (RQ5)", "The records carry their own errors");
  const items = [
    ["ICS-209-PLUS ACRES is capped", "The cleaned field is capped at the final size and hid LNU's 375,209 ac peak. Use CURR_INCIDENT_AREA."],
    ["Derived growth dates lag", "Camp's largest growth is dated 9 Nov, one report after the run. These fields cannot time a blowup."],
    ["Satellite area is biased per fire", "FEDS +9% on Carr, -2% on Camp. The FIRED event is 23% below Camp's final size; its ignition point sits 17 km west of the origin."],
    ["Time stamps may be processing times", "Two Carr GeoMAC perimeters are stamped before ignition; Camp's first perimeter stamp is uncertain."],
    ["Counts disagree between sources", "Personnel counts differ by a factor of 2 to 4 between NICC and ICS-209-PLUS."],
  ];
  items.forEach(([h, b], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = M + col * 6.15, y = 1.6 + row * 1.6, w = 5.95;
    card(s, x, y, w, 1.42, C.background2, "error " + (i + 1));
    numCircle(s, x + 0.22, y + 0.22, i + 1, C.accent5, 0.42);
    s.addText([T(h, { bold: true, breakLine: true }), T(b, { color: C.text2 })], { x: x + 0.8, y: y + 0.16, w: w - 1.0, h: 1.15, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  });
  const x = M + 6.15, y = 1.6 + 2 * 1.6;
  card(s, x, y, 5.95, 1.42, C.text1, "so what");
  s.addText([T("So the model carries a separate observation-error layer, ", { bold: true }), T("so reporting error is never mistaken for model error. Full register: workbook, Data Quality sheet.")], { x: x + 0.25, y: y + 0.12, w: 5.45, h: 1.2, fontSize: 14, color: C.background1, margin: 0, valign: "middle", isTextBox: true });
  s.addNotes("RQ5. Before scoring any model against the records, we have to know how wrong the records are. Five classes found so far. One: the ICS-209-PLUS cleaned ACRES field is capped at the final size, which hid LNU's 375,209 acre peak; use CURR_INCIDENT_AREA. Two: derived growth fields inherit the reporting lag. Three: satellite area has a per-fire bias, plus the FIRED event undercounts Camp and misplaces its ignition. Four: some archive time stamps are processing times, not flight times. Five: personnel counts differ 2 to 4 times between sources. These become an observation-error layer in the model, kept separate from model error. The full register is in the workbook Data Quality sheet.");
}

// =====================================================================
section("Where drones close the gap");
divider("Where drones can close the gap", "Current and future capability mapped gap by gap. Week 5 (W5) figures are agency or vendor claims, not tests");

{
  const s = content("Drone capability", "Three tiers: persistent, close-in, and future");
  const tiers = [
    ["Persistent wide area", "NWCG Type 1 fixed wing", C.accent3, "Fits M1, M2 wide area; M3 overwatch and relay", [
      "L3Harris FVR-90: 12 to 18 h (W5); flown with ACERO as relay; maps perimeters on USFS and BLM fires (W5)",
      "Type 1 class: 6 to 14 h, flies above all other aircraft (NWCG 2025)",
      "NASA SuperVolo VTOL: 8 h, flown with ACERO as alternate relay, day only in tests",
      "Wind limit: not found in project records"]],
    ["Close-in tier", "NWCG Type 3 and 4 rotorcraft", C.accent2, "Fits M1, M2 close in; M3 escape route checks", [
      "Alta X: 50 min empty, 22 min at 22 lb; FY22 average sortie 17 min; wind limit 22 mph (W5)",
      "SwitchBlade-Elite: largest USFS type (66), 18 min average sortie, 25 mph (W5)",
      "DOI fleet: 531 aircraft and 10,342 flights in FY2018",
      "Night flight needs a DOI/FAA agreement"]],
    ["Future: tests only", "M4 water drop and M5 logistics", C.accent1, "No library drone has flown either mission", [
      "Seneca Argo-1 swarm: 20 gal each, CAL FIRE test July 2026",
      "FireSwarm Thunder Wasp: up to 100 gal; Dryad Silvanet: XPRIZE test",
      "Rain and Sikorsky autonomous Firehawk: demo, hands-off in gusts to 30 kt",
      "KARGO II 1,500 lb; Elroy Chaparral 500 lb over 300 nm (in development)"]],
  ];
  const w = 3.9, g = 0.2;
  tiers.forEach(([name, cls, col, fit, pts], i) => {
    const x = M + i * (w + g);
    card(s, x, 1.6, w, 5.25, C.background2, "tier " + name);
    chip(s, x + 0.25, 1.82, name, col, C.background1);
    s.addText(cls, { x: x + 0.25, y: 2.25, w: w - 0.5, h: 0.4, fontSize: 16, bold: true, color: C.text1, margin: 0, isTextBox: true });
    s.addText(fit, { x: x + 0.25, y: 2.65, w: w - 0.5, h: 0.5, fontSize: 12, italic: true, color: C.accent6, margin: 0, valign: "top", isTextBox: true });
    s.addText(pts.map((p, j) => T(p, { bullet: { indent: 12 }, breakLine: j < pts.length - 1 })), { x: x + 0.2, y: 3.2, w: w - 0.4, h: 3.55, fontSize: 13, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 7, isTextBox: true });
  });
  s.addNotes("The 15-row capability table in Section 5.1 condensed into three tiers. Persistent wide area: NWCG Type 1 fixed wing such as the FVR-90 and contract ScanEagle, Aerosonde and Penguin C class aircraft, 6 to 18 hours of endurance; the FVR-90 and SuperVolo have flown with NASA ACERO as radio relays, though those flights were daytime VLOS. Close-in: Type 3 and 4 rotorcraft like the Alta X and SwitchBlade-Elite. They are in service in large numbers, but average sorties of 17 to 18 minutes and published wind limits of 22 to 25 mph. Future: water drop and logistics drones exist only as tests, demonstrators and development programs; payloads are 20 to 100 gallons per sortie against 300 to 1,000 for helicopters. Also tracked: Joby optionally piloted aircraft (flown with ACERO, high orbit telemetry only, a future night water-drop class), Ondas drone-in-a-box, and the Data Blanket swarm with FAA approval for one pilot flying many drones. Every Week 5 figure is an agency or vendor statement.");
}

{
  const s = content("What 24/7 takes", "Holding one station for 24 hours");
  s.addChart(pres.charts.BAR, [{ name: "Sorties per 24 h", labels: ["FVR-90 (Type 1), 12 h", "SuperVolo, 8 h", "Type 1 class low end, 6 h", "Alta X (Type 3), 17 min avg"], values: [2, 3, 4, 85] }], {
    x: M, y: 1.55, w: 7.6, h: 4.6, barDir: "bar", chartColors: [HEX.accent3, HEX.accent3, HEX.accent3, HEX.accent2], barGapWidthPct: 45,
    showValue: true, dataLabelPosition: "outEnd", dataLabelFormatCode: "0", showLegend: false,
    showTitle: true, title: "Sorties needed for one continuous station, per 24 h", catAxisOrientation: "maxMin",
    valAxisMaxVal: 100, valAxisMajorUnit: 25, valGridLine: grid(), catGridLine: { style: "none" }, ...AX,
    objectName: "sorties chart",
  });
  s.addText("Lower bounds: transit, turnaround and spares not included. Remote pilots fly at most 8 h per duty day (DOI OPM-11, W5), so every option needs 3 pilot shifts; Alta X needs 3 per launch site, each flying about 28 sorties.", { x: M, y: 6.25, w: 7.6, h: 0.7, fontSize: 12, italic: true, color: C.accent6, margin: 0, isTextBox: true });
  const x = 8.55, w = W - M - x;
  card(s, x, 1.6, w, 2.4, C.background2, "type 1");
  chip(s, x + 0.2, 1.78, "Fits the requirement", C.accent3);
  s.addText("A Type 1 fixed wing holds the 2 to 3 h revisit with 2 to 4 sorties a day.", { x: x + 0.2, y: 2.2, w: w - 0.4, h: 1.7, fontSize: 15, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  card(s, x, 4.2, w, 2.65, C.background2, "type 3");
  chip(s, x + 0.2, 4.38, "Close-in only", C.accent2);
  s.addText("Type 3 rotorcraft cannot hold persistent cover. They belong in spot fire confirmation and escape route checks, as in the Week 6 fleet mix.", { x: x + 0.2, y: 4.8, w: w - 0.4, h: 1.95, fontSize: 15, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  s.addNotes("This chart is simple arithmetic: 24 hours divided by usable endurance gives the number of sorties needed to keep one aircraft always on station over a fire. FVR-90 at 12 hours needs 2; SuperVolo at 8 hours needs 3; the low end of the Type 1 class at 6 hours needs 4; the Alta X at its FY22 average sortie of 17 minutes needs about 85. These are lower bounds because transit, turnaround and spares are not included. Pilot duty limits mean three shifts a day in every case. The conclusion: Type 1 fixed wing for persistent coverage, Type 3 for the close-in tier. How many stations a fire needs depends on perimeter length and sensor swath, which the platform analysis still has to set.");
}

{
  const s = content("Gap by gap", "Wind and warnings remain the hard gaps");
  const hdr = ["Gap", "What the fires show", "Best current fit", "What still blocks it", "Status"].map((t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 } } }));
  const rows = [
    ["First hours", "First perimeter 3.5 to 14 h; Camp burnovers 85 min after the call", "Type 3 first look (20 to 60 min); Type 1 VTOL for the whole fire", "Dispatch and airspace: NOTAM, TFR before smoke or night work", ["Closable", C.accent3]],
    ["Revisit", "Perimeters 14.5 to 18.8 h apart; gaps to 140 h. Need 2 to 3 h", "Type 1 fixed wing, 6 to 18 h, above other aircraft", "Pilots and airframes for continuous cover; swath per fire not set", ["Closable", C.accent3]],
    ["Night", "37% of LNU growth at night; 4 to 21% of perimeters at night", "FVR-90 day and night (W5); drones have mapped at night", "DOI/FAA night agreement; night BVLOS not shown in the library", ["Approvals", C.accent5]],
    ["Warning", "Orders about at fire arrival; unanswered call; 62 min to mayday", "FVR-90, SuperVolo flew as relays (TCL 1)", "Relay success 52 to over 90%; no crew tracking or alert delivery shown", ["Weakest link", C.accent1]],
    ["Wind", "Camp, Pioneer, LNU critical days above every fleet limit", "Only Carr within Alta X limit, marginally", "The worst days are the windiest days", ["Binding", C.accent1]],
    ["M4 and M5", "No library drone has flown either; Camp crews needed water and hose", "None", "20 to 100 gal per sortie vs 300 to 1,000; approvals pending", ["Future", C.accent6]],
  ];
  const body = rows.map((r, i) => r.map((t, j) => {
    const bg = i % 2 ? C.background1 : C.background2;
    if (j === 4) return { text: t[0].toUpperCase(), options: { fill: { color: t[1] }, color: j === 4 && t[1] === C.accent5 ? C.text1 : C.background1, bold: true, align: "center", fontSize: 11 } };
    return { text: t, options: { fill: { color: bg }, bold: j === 0, color: C.text1 } };
  }));
  s.addTable([hdr, ...body], { x: M, y: 1.6, w: 12.1, colW: [1.35, 3.45, 3.0, 3.0, 1.3], fontSize: 12, fontFace: "Calibri", valign: "middle", border: { type: "solid", pt: 0.5, color: "E1E4E8" }, margin: [4, 7, 4, 7], rowH: 0.78, objectName: "gap by gap table" });
  s.addNotes("Section 5.2 condensed. Each row is a gap measured on the four fires, the drone that fits best today, what still blocks it, and a status. First hours and revisit are closable in principle with aircraft classes already flying on fires; they enter the model as faster T detect and fresher information age. Night is closable but gated by approvals: night flight needs a DOI/FAA agreement and night BVLOS is not demonstrated in the library. Warning delivery is the weakest link: relays have flown, but message success was 52 percent with four ground nodes and over 90 percent with three, a single relay does not scale, and no drone has demonstrated crew tracking or alert delivery. Wind is binding, next slide. Water drop and logistics support are future capability, evaluated as scenarios against the M5 questions, not as validated performance.");
}

{
  const s = content("The binding constraint", "Wind on the worst days exceeds every fleet limit");
  const labels = ["Carr 2018", "Camp 2018 (day 1)", "Pioneer 2016", "LNU 2020"];
  s.addChart([
    { type: pres.charts.BAR, data: [{ name: "Upper end of reported wind, mph", labels, values: [21, 40, 40, 40] }], options: { barDir: "col", chartColors: [HEX.accent6], barGapWidthPct: 70 } },
    { type: pres.charts.LINE, data: [
      { name: "Alta X limit, 22 mph", labels, values: [22, 22, 22, 22] },
      { name: "SwitchBlade-Elite limit, 25 mph", labels, values: [25, 25, 25, 25] },
    ], options: { chartColors: [HEX.accent2, HEX.accent1], lineSize: 2, lineDataSymbol: "none" } },
  ], {
    x: M, y: 1.55, w: 7.9, h: 5.3, showLegend: true, legendPos: "b", showTitle: true, title: "Critical-period wind against published drone limits",
    valAxisMaxVal: 50, valAxisMinVal: 0, valAxisMajorUnit: 10, valGridLine: grid(), catGridLine: { style: "none" }, ...AX,
    valAxisTitle: "mph", showValAxisTitle: true, valAxisTitleFontSize: 11, valAxisTitleColor: HEX.accent6, valAxisTitleFontFace: "+mn-lt",
    objectName: "wind chart",
  });
  const x = 8.85, w = W - M - x;
  const facts = [
    ["Camp, first morning", "20 to 30 mph sustained, gusts over 40. Aircraft grounded; air attack 'not very effective due to smoke and wind' (NIST TN 2135 p.403)"],
    ["Pioneer and LNU", "Gusts 28 to 40 and 20 to 40 mph on the runs"],
    ["Type 1 fixed wing", "No published limit found. U21 runs a 25 to 40 mph sensitivity range"],
  ];
  facts.forEach(([h, b], i) => {
    const y = 1.6 + i * 1.55;
    card(s, x, y, w, 1.4, C.background2, "wind " + h);
    s.addText([T(h, { bold: true, breakLine: true }), T(b, { color: C.text2 })], { x: x + 0.2, y: y + 0.1, w: w - 0.4, h: 1.2, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  });
  s.addNotes("How to read it: grey columns are the upper end of the wind reported for each fire's critical period, from the Week 6 wind table (Carr up to 21 mph; Camp first morning 20 to 30 sustained with gusts over 40; Pioneer gusts 28 to 40; LNU 20 to 40). Columns show the top of each range, so read them as the worst hours, not the whole day. The two flat lines are the published limits of the two most common agency drones: Alta X at 22 mph and SwitchBlade-Elite at 25 mph. Only Carr sits under the Alta X line, and only marginally. The days that matter most are the windiest days, and no Type 1 fixed-wing wind limit was found, so the model treats drone availability as an uncertain input conditioned on hourly wind (U21), with a 25 to 40 mph sensitivity range for Type 1. Note these are surface winds; flight-level wind is usually higher.");
}

// =====================================================================
section("Uncertainty model");
divider("The uncertainty model", "Where each finding enters the safety margin, and where the gaps still are");

{
  const s = content("Uncertainty inputs", "Every term in the margin now has a measured prior");
  card(s, M, 1.55, 12.1, 0.9, C.text1, "equation band");
  s.addText(eqRuns(24, C.background1), { x: M, y: 1.6, w: 12.1, h: 0.8, align: "center", valign: "middle", margin: 0, isTextBox: true });
  const cols = [
    ["T fire", C.accent1, [["U1", "Spread model error: MAPE 47%, timber underpredicts (Cardil 2023)"], ["U5, U12", "Information age: official size about 1 day old; perimeters 14.5 to 18.8 h apart"], ["U18", "Last-known size shortfall: 2.6% at 2 h to 32.2% at 24 h"]]],
    ["T detect", C.accent3, [["U17", "ICS-209 discovery 1 to 8 h after first report: not a detection time"], ["U22 new", "Time to first drone perimeter: launch, transit, mapping; 20 to 60 min first look (elicit)"]]],
    ["T notify", C.accent4, [["U9", "Latency and outright failure: Camp 13 to 59 min, unanswered call; Carr 62 min"], ["U24 new", "Warning relay success: 52% to over 90% by ground nodes (Wu 2025)"]]],
    ["T egress", C.accent2, [["U8", "Crew travel time from GPS travel-rate models (Sullivan 2020; Campbell 2017, 2019)"]]],
  ];
  const w = 2.9, g = 0.167;
  cols.forEach(([name, col, items], i) => {
    const x = M + i * (w + g);
    s.addText(name, { x, y: 2.6, w, h: 0.4, fontSize: 16, bold: true, color: col, margin: 0, isTextBox: true });
    items.forEach(([id, txt], j) => {
      const y = 3.05 + j * 1.12;
      card(s, x, y, w, 1.02, C.background2, id);
      s.addText([T(id, { bold: true, color: col, breakLine: true }), T(txt)], { x: x + 0.15, y: y + 0.06, w: w - 0.3, h: 0.92, fontSize: 12, color: C.text1, margin: 0, valign: "top", isTextBox: true });
    });
  });
  card(s, M, 6.25, 12.1, 0.65, C.background2, "gates");
  s.addText([T("Gates on every drone-dependent term:  ", { bold: true }), T("U21 drone availability by wind (fly or no fly per airframe, hourly)   |   U23 probability a night or BVLOS mission is approved in time")], { x: M + 0.25, y: 6.27, w: 11.6, h: 0.6, fontSize: 13, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  s.addNotes("This maps the findings onto the margin. T fire carries spread model error (U1, Cardil's 47 percent MAPE, with timber fires underpredicted), information age (U5 and U12), and the revisit shortfall curve (U18). T detect carries the discovery-time trap (U17) and the new time to first drone perimeter (U22), which has to be elicited. T notify carries latency and outright failure (U9) and the new relay success term (U24). T egress uses GPS travel-rate models (U8). Two new gates sit across every drone-dependent term: U21, drone availability by hourly wind against airframe limits, and U23, the probability night or BVLOS flight is approved in time. U1 to U20 are in the workbook Uncertainty Register.");
}

{
  const s = content("Uncertainty gaps", "Where the uncertainty is still largest");
  const hdr = ["Input", "What we have", "What is missing", "Status"].map((t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 } } }));
  const G = [C.accent3, "Measured"], P = [C.accent5, "Partial"], O = [C.accent1, "Open"];
  const rows = [
    ["Information age (U5, U12)", "Measured on all four fires", "Night perimeters (NIROPS)", G],
    ["Revisit shortfall (U18)", "LNU hourly curve", "A second and third fire (GOFER for Carr, Camp); shape accuracy", P],
    ["Detection (U17, T detect)", "First-report chain for three fires", "LNU VIIRS and component ICS-209s; dispatch logs", P],
    ["Notification (U9)", "Camp minute-level; Carr one interval", "Radio and dispatch logs; Carr and LNU alert times", P],
    ["Spread error (U1)", "Published MAPE 47%", "Fuels for the ensembles (LANDFIRE) not yet downloaded", P],
    ["Egress (U8)", "Generic GPS travel rates", "Crew positions, escape routes and safety zones at each burnover", P],
    ["Drone wind gate (U21)", "ASOS and ERA5 hourly in hand, not analysed", "Type 1 wind limit; RAWS; wind at flight altitude", O],
    ["First drone, approvals (U22, U23)", "Nothing", "Stakeholder elicitation", O],
    ["Relay success (U24)", "One field test, 52 to over 90%", "Alert delivery to crews never demonstrated", O],
    ["M5 delivery, P(on time)", "Vendor and test payloads (W5)", "Any flown delivery performance", O],
  ];
  const body = rows.map((r, i) => r.map((t, j) => {
    if (j === 3) return { text: t[1].toUpperCase(), options: { fill: { color: t[0] }, color: t[0] === C.accent5 ? C.text1 : C.background1, bold: true, align: "center", fontSize: 11 } };
    return { text: t, options: { fill: { color: i % 2 ? C.background1 : C.background2 }, bold: j === 0, color: C.text1 } };
  }));
  s.addTable([hdr, ...body], { x: M, y: 1.55, w: 12.1, colW: [3.0, 3.6, 4.3, 1.2], fontSize: 12, fontFace: "Calibri", valign: "middle", border: { type: "solid", pt: 0.5, color: "E1E4E8" }, margin: [3, 7, 3, 7], rowH: 0.47, objectName: "uncertainty gaps table" });
  s.addNotes("Our read of where the uncertainty sits today, input by input. This status column is a summary made for the deck, not a table from the proposal. Measured means we have it from records on all four fires. Partial means we have a measured value but it rests on one fire, one event, or published reports only. Open means we have a range or nothing, and the term has to come from elicitation, a vendor, or new data. The pattern: the information-gap side of the model is well measured; the drone side (wind gate, approvals, relay, delivery) is mostly open. That is where the next round of data collection should go.");
}

{
  const s = content("Methods and experiment", "Ensembles, two error layers, and a factorial design");
  const steps = [
    ["Spread ensembles", "ELMFIRE or GridFire with spatially correlated perturbations of fuel, wind and moisture"],
    ["Two error layers", "Model error from Cardil's residuals; observation error from the records analysis"],
    ["Assimilation", "Perimeter assimilation is the lever for RQ4 (Kochanski 2023)"],
    ["Delivery simulation", "Discrete-event model of sorties, driven by recorded daily crew counts (M5)"],
    ["Elicitation", "PERT for terms with no data (Hope 2024)"],
    ["Sensitivity", "One at a time, then Sobol (RQ3)"],
  ];
  steps.forEach(([h, b], i) => {
    const y = 1.6 + i * 0.86;
    numCircle(s, M, y, i + 1, C.accent1, 0.42);
    s.addText([T(h + ". ", { bold: true }), T(b, { color: C.text2 })], { x: M + 0.6, y: y - 0.04, w: 5.1, h: 0.8, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  });
  const hdr = ["Factor", "Levels"].map((t) => ({ text: t, options: { bold: true, color: C.background1, fill: { color: C.text1 } } }));
  const rows = [
    ["Fire case", "Camp 8 Nov and Carr 26 Jul (validation); Pioneer; LNU"],
    ["Information source", "Official record (1 day); aircraft (14 to 19 h); satellite 12 h; drone 6, 3, 2, 1 h"],
    ["Drone availability", "Day only; 24 h; 24 h with wind limit; plus authorisation delay"],
    ["Time of day", "Day VMC; night DVE; smoke DVE"],
    ["Warning relay", "52%; 87%; 95% message success"],
    ["Model", "ELMFIRE; GridFire"],
  ];
  const body = rows.map((r, i) => r.map((t, j) => ({ text: t, options: { fill: { color: i % 2 ? C.background1 : C.background2 }, bold: j === 0, color: C.text1 } })));
  s.addTable([hdr, ...body], { x: 6.5, y: 1.6, w: 6.23, colW: [1.85, 4.38], fontSize: 12, fontFace: "Calibri", valign: "middle", border: { type: "solid", pt: 0.5, color: "E1E4E8" }, margin: [4, 7, 4, 7], rowH: 0.56, objectName: "experiment table" });
  card(s, 6.5, 5.75, 6.23, 1.1, C.background2, "run order");
  s.addText([T("Run Camp 8 Nov and Carr 26 Jul first: ", { bold: true }), T("both have timed crew hazards (19 NIST burnovers; Carr warning and mayday). Pioneer and LNU validate spread and information age, not P(M < 0).")], { x: 6.7, y: 5.8, w: 5.85, h: 1.0, fontSize: 13, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
  s.addNotes("Methods are unchanged in substance from v0.4. Left: the six method components. Right: the experimental factors. Camp 8 November and Carr 26 July run first because they are the only cases with timed crew hazards to validate P(M < 0) against: Camp's minute-level NIST progression and 19 burnovers plus NETTRA notification times, and Carr's timed warning and mayday plus 63 GeoMAC and 88 FEDS perimeters. Forecasts are scored against time-matched perimeters with the observation-error layer applied. A fractional design or Latin hypercube keeps the run count manageable.");
}

{
  const s = content("Limitations", "What these findings can and cannot claim");
  const lim = [
    ["Four fires", "The figures describe these incidents, not a national rate"],
    ["One revisit curve", "From LNU only, with 2 km GOES pixels"],
    ["Time stamps", "Archive stamps may record processing, not flight time (Camp's first perimeter, two Carr perimeters)"],
    ["Vendor claims", "Week 5 drone figures are agency and vendor statements; Type 1 wind limits not found"],
    ["Satellite search", "FIRMS search is by distance from origin and may catch unrelated fires in the first hour"],
    ["Published reports only", "Radio and dispatch logs not in hand; notification times rest on NIST and Green Sheets"],
    ["Entrapments", "None found for Pioneer or LNU in one day of searching (30 Sep 2026)"],
    ["Novelty", "Not assessed. Run the literature gate before any claim that this is new"],
  ];
  lim.forEach(([h, b], i) => {
    const col = i % 2, row = Math.floor(i / 2);
    const x = M + col * 6.15, y = 1.6 + row * 1.32, w = 5.95;
    card(s, x, y, w, 1.17, C.background2, "limit " + h);
    s.addText([T(h, { bold: true, breakLine: true }), T(b, { color: C.text2 })], { x: x + 0.25, y: y + 0.1, w: w - 0.5, h: 1.0, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  });
  s.addNotes("The proposal's limitations, as written. Two to stress in discussion: the revisit requirement rests on one fire, and every drone figure from Week 5 is a claim, not a test. Also, novelty has not been assessed: before anyone says this has not been done, run the literature gate.");
}

// =====================================================================
section("Next steps");
divider("What we are missing, and what comes next", "Data needs by priority, the next steps in order, and suggestions to strengthen the method");

{
  const s = content("What we are missing", "Data needs, by priority");
  const pri = [
    ["Priority 1", "Before the Camp and Carr runs", C.accent1, [
      "Radio audio and transcripts (U9)", "Dispatch and CAD logs: true first report (U17)", "NIROPS night infrared perimeters", "Hourly wind vs drone limits (in hand)", "RAWS hourly: WRCC access code", "Type 1 wind limit (Overwatch Aero or NASA)", "Crew positions, escape routes (T egress)", "LANDFIRE fuels for the ensembles", "True time of Camp's 09:54 perimeter"]],
    ["Priority 2", "Completes the four-fire picture", C.accent5, [
      "ICS-214 logs and ICS-205 radio plans", "Alert times for Carr and LNU", "Night operations logs (LNU 18 to 19 Aug)", "GOES-16 fire detections", "FIRMS VIIRS 2020 and LNU component 209s", "GOFER hourly for Carr and Camp", "On-fire weather; HRRR winds aloft", "Crewed aircraft logs and grounding times", "Elicitation for U22, U23"]],
    ["Priority 3", "Strengthens it", C.accent6, [
      "USFS and CAL FIRE drone logs during the fires", "Historical TFR records (U23)", "Entrapment search for Pioneer and LNU"]],
  ];
  const w = 3.9, g = 0.2;
  pri.forEach(([p, sub, col, items], i) => {
    const x = M + i * (w + g);
    card(s, x, 1.6, w, 5.25, C.background2, p);
    chip(s, x + 0.25, 1.82, p, col, col === C.accent5 ? C.text1 : C.background1);
    s.addText(sub, { x: x + 0.25, y: 2.2, w: w - 0.5, h: 0.4, fontSize: 14, italic: true, color: C.accent6, margin: 0, isTextBox: true });
    s.addText(items.map((t, j) => T(t, { bullet: { indent: 12 }, breakLine: j < items.length - 1 })), { x: x + 0.2, y: 2.65, w: w - 0.4, h: 4.1, fontSize: 13, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 5, isTextBox: true });
  });
  s.addNotes("From the companion Key Findings and Data Needs brief. Priority 1 is needed before the Camp and Carr validation runs: radio and dispatch records for T detect and T notify, NIROPS night perimeters, the wind analysis we can run now with ASOS and ERA5 already in hand, RAWS hourly data behind a WRCC access code, a Type 1 wind limit, crew positions and escape routes, LANDFIRE fuels, and the true time of Camp's first perimeter. Priority 2 completes the four-fire picture. Priority 3 strengthens it. Data availability was checked against the project folder on 30 September 2026 and is not a full survey of what exists.");
}

{
  const s = content("Next steps", "In order: what we can do now, and what takes weeks");
  const cols = [
    ["This week", "No outside help", C.accent3, ["Agree P(M < 0) and P(on time) as headline metrics, with information age and drone availability as covariates", "Build U21: hourly ASOS and ERA5 wind against airframe limits", "Check Camp's 09:54 perimeter against NIST maps", "Download LANDFIRE fuels for Camp and Carr"]],
    ["Send now", "Replies take weeks", C.accent1, ["Pioneer FOIA to USFS R4 (draft ready)", "Records requests to CAL FIRE for Camp and Carr radio, dispatch, ICS-214s; NPS for Carr", "WRCC access code for RAWS", "FVR-90 wind limit: Overwatch Aero or NASA"]],
    ["Through partners", "", C.accent2, ["NIROPS access for night infrared perimeters", "Ask GOFER authors about Carr and Camp", "Stakeholder questionnaire for U22 and U23"]],
    ["Then", "", C.accent6, ["100-member ELMFIRE ensembles, Camp 8 Nov and Carr 26 Jul: observed cadence and 6, 3, 2, 1 h drone revisit, with and without the wind gate", "FIRMS VIIRS and component 209s for LNU", "Add radio log times when they arrive"]],
  ];
  const w = 2.9, g = 0.167;
  cols.forEach(([h, sub, col, items], i) => {
    const x = M + i * (w + g);
    numCircle(s, x, 1.6, i + 1, col, 0.5);
    s.addText(h, { x: x + 0.62, y: 1.58, w: w - 0.62, h: 0.32, fontSize: 17, bold: true, color: C.text1, margin: 0, valign: "middle", isTextBox: true });
    if (sub) s.addText(sub, { x: x + 0.62, y: 1.9, w: w - 0.62, h: 0.26, fontSize: 12, italic: true, color: C.accent6, margin: 0, isTextBox: true });
    card(s, x, 2.35, w, 4.5, C.background2, "step " + h);
    s.addText(items.map((t, j) => T(t, { bullet: { indent: 12 }, breakLine: j < items.length - 1 })), { x: x + 0.15, y: 2.5, w: w - 0.3, h: 4.25, fontSize: 13, color: C.text1, margin: 0, valign: "top", paraSpaceAfter: 8, isTextBox: true });
  });
  s.addNotes("The proposal's next steps, ordered by the suggested sequence in the companion brief. This week, with no outside help: agree the headline metrics, run the wind comparison for U21, check the Camp perimeter time against NIST maps, and pull LANDFIRE fuels. Send now, because replies take weeks: the Pioneer FOIA, which is drafted; CPRA requests for Camp and Carr; a WRCC access code; and the FVR-90 wind limit question. Through partners: NIROPS, GOFER coverage, and the stakeholder questionnaire for time to first drone and night approval. Then the first real ensemble runs on Camp and Carr at the observed cadence and at drone revisit intervals, with and without the wind gate.");
}

{
  const s = content("Suggestions: modelling", "Five ways to make the uncertainty model more honest");
  const items = [
    ["Sample wind once, use it twice", "The same wind draw should drive both ELMFIRE spread and the U21 drone gate. Sampled independently, the model will put drones in the air on the very hours the fire runs, and understate P(M < 0) on the bad days."],
    ["Model notification as fail or delay", "Use a mixture for U9 and U24: a probability the warning never arrives, plus a latency distribution when it does. Concow and the unanswered Rattlesnake Flats call show failure is real."],
    ["Make observation error hierarchical", "Fit per-record lag and bias (ICS-209 about 1 day; FEDS +9%, -2%) with partial pooling across the four fires, so a fifth fire gets an honest prior instead of a fixed correction."],
    ["Separate what we don't know from what varies", "Two-loop Monte Carlo: outer loop for epistemic terms (Type 1 wind limit, U22, U23), inner loop for wind and moisture. Report P(M < 0) with an interval and show which part shrinks with data."],
    ["Gate on gusts at flight altitude", "Correct 10 m ASOS and ERA5 wind to flight level (HRRR) and use gusts, not sustained wind. Treat each limit as a soft availability curve rather than a hard cutoff."],
  ];
  items.forEach(([h, b], i) => {
    const y = 1.55 + i * 1.07;
    numCircle(s, M, y + 0.05, i + 1, C.accent2, 0.45);
    s.addText([T(h + ". ", { bold: true }), T(b, { color: C.text2 })], { x: M + 0.65, y, w: 11.45, h: 1.0, fontSize: 14, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  });
  s.addNotes("These are suggestions, not part of v0.5; the team should weigh them. The first one matters most. Wind drives both fire spread and drone availability. If the ensemble samples wind for ELMFIRE and separately samples whether a drone could fly, the two will be uncorrelated, and the model will often have drones flying during the wind event that drives the run. That biases P(M < 0) low exactly where it matters. Use one joint wind draw for both. Second, notification has two failure modes, late and never; a single latency distribution hides the never. Third, partial pooling across fires turns the records analysis into priors that transfer to new fires. Fourth, separating epistemic from aleatory uncertainty tells the team which uncertainty more data can buy down. Fifth, drone limits apply at flight altitude and to gusts, so surface sustained wind will understate groundings.");
}

{
  const s = content("Suggestions: validation and data", "Five ways to make the results defensible");
  const items = [
    ["Score calibration, not only error", "CRPS, reliability diagrams and PIT histograms for ensemble perimeters against time-matched perimeters. For P(M < 0), backcast the 19 Camp burnovers and the Carr mayday: did the model give each real probability?"],
    ["Use an emulator for Sobol", "Train a Gaussian process or polynomial chaos surrogate on the Latin hypercube ELMFIRE runs. Sobol on raw ensembles needs thousands of runs."],
    ["Rank data requests by value of information", "Expected value of partial perfect information (EVPPI): which of radio logs, NIROPS or RAWS would move P(M < 0) most. Chase that one first."],
    ["Replicate the revisit curve first", "It is about to become a requirement on one fire. Add GOFER Carr and Camp, bootstrap intervals, and a shape metric (IoU or Hausdorff distance), since area alone hides where the fire went."],
    ["Structure the elicitation", "Run U22 and U23 with a protocol such as SHELF or IDEA, with calibration questions (Cooke's classical model) so expert weights are earned, not assumed."],
  ];
  items.forEach(([h, b], i) => {
    const col = i < 3 ? 0 : 1, row = i < 3 ? i : i - 3;
    const x = M + col * 6.15, y = 1.6 + row * 1.75, w = 5.95;
    card(s, x, y, w, 1.6, C.background2, "suggestion " + (i + 6));
    numCircle(s, x + 0.2, y + 0.2, i + 6, C.accent2, 0.42);
    s.addText([T(h, { bold: true, breakLine: true }), T(b, { color: C.text2 })], { x: x + 0.78, y: y + 0.14, w: w - 0.98, h: 1.4, fontSize: 13, color: C.text1, margin: 0, valign: "top", isTextBox: true });
  });
  const x = M + 6.15, y = 1.6 + 2 * 1.75;
  card(s, x, y, 5.95, 1.6, C.text1, "close");
  s.addText([T("The aim: ", { bold: true }), T("a P(M < 0) the team can defend, with an interval, validated on Camp and Carr, and a clear list of which data would shrink it.")], { x: x + 0.25, y: y + 0.12, w: 5.45, h: 1.36, fontSize: 15, color: C.background1, margin: 0, valign: "middle", isTextBox: true });
  s.addNotes("More suggestions, these about validation and data. Calibration scoring checks whether the ensemble spread is honest: if the truth falls outside the ensemble too often, the uncertainty is understated. The burnover backcast is the most direct test of P(M < 0) we have. An emulator makes Sobol sensitivity affordable. Value of information analysis turns the data-needs list into a ranked order based on what actually moves the answer. The revisit curve should be replicated before it is presented as a requirement, since right now it comes from one fire. Structured elicitation with calibration questions makes the drone timing and approval priors defensible. Methods to cite if adopted: Gneiting and Raftery 2007 on proper scoring rules; Saltelli on Sobol indices; O'Hagan et al. on SHELF; Cooke 1991 on the classical model. Check these against the library before they go into the proposal.");
}

{
  const s = pres.addSlide({ masterName: "SECTION", sectionTitle: SECTION });
  s.addText("Discussion", { placeholder: "title" });
  s.addText("1. Agree the headline metrics.   2. Assign the priority 1 data requests.   3. Pick which suggestions go into v0.6.", { placeholder: "body" });
  s.addNotes("Close on three decisions for the team: agree P(M < 0) and P(on time) as the headline metrics, assign owners for the priority 1 data requests that take weeks, and choose which method suggestions fold into v0.6.");
}

(async () => {
  await pres.writeFile({ fileName: OUT });
  await applyTheme(OUT, THEME);
  console.log("wrote", OUT);
})();
