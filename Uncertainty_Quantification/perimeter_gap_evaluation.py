"""Observation-gap and time-area uncertainty evaluation for the four case study fires.

Inputs
- perimeter_supplements/Wildfire Perimeter Time Area Sources.xlsx (Pioneer, Carr, LNU: ICS-209, GeoMAC,
  GOFER hourly GOES reconstruction, FEDS objects, agency and published observations, dated events)
- Fire Case Studies/Camp_2018 workbook + research_data GeoMAC / FEDS histories (Camp is not in the supplement)
- research_data/datasets/01_progression_perimeters/Carr_2018_NASA_FEDS_perimeter_12h.geojson (Carr FEDS series)

Outputs: perimeter_eval/results.json and CSV tables, read by build_uq_v04.py.
Run: uv run --project ~/nasa_rag python perimeter_gap_evaluation.py
"""
import json, csv, statistics as st
from datetime import datetime, timedelta, timezone
from pathlib import Path
from openpyxl import load_workbook

HERE = Path(__file__).parent
SUP = HERE / "perimeter_supplements" / "Wildfire Perimeter Time Area Sources.xlsx"
PER = HERE / "research_data" / "datasets" / "01_progression_perimeters"
CAMP_WB = HERE.parent / "Fire Case Studies" / "Camp_2018" / "Camp_Fire_2018_ICS209.xlsx"
OUT = HERE / "perimeter_eval"; OUT.mkdir(exist_ok=True)
KM2 = 247.105

def q(xs, p):
    xs = sorted(xs)
    if not xs: return None
    k = (len(xs) - 1) * p; f = int(k); c = min(f + 1, len(xs) - 1)
    return xs[f] + (xs[c] - xs[f]) * (k - f)
def summ(xs):
    xs = [x for x in xs if x is not None]
    return dict(n=len(xs), median=q(xs, .5), p90=q(xs, .9), max=max(xs), mean=st.mean(xs)) if xs else dict(n=0)

def family(src, typ):
    s, t = (src or "").lower(), (typ or "").lower()
    if t.startswith("event") or t.startswith("suppression"): return "event"
    if "gofer" in s or "feds" in s or "satellite" in t or "modis" in t or "viirs" in t: return "satellite"
    if "geomac" in s or "nirops" in t or "cal fire 20200828" in s or "progression map" in t: return "mapped"
    if "209" in s or "report" in t or "estimate" in t or "acreage" in t: return "report"
    return "other"

# ---------------------------------------------------------------- load observations
obs = {}
wb = load_workbook(SUP, data_only=True)
for fire in ("Pioneer 2016", "Carr 2018", "LNU 2020"):
    rows = []
    for r in wb[fire].iter_rows(min_row=2, values_only=True):
        t, a, src, typ, com = (list(r) + [None] * 5)[:5]
        if not isinstance(t, datetime): continue
        rows.append(dict(t=t, acres=a if isinstance(a, (int, float)) else None, src=src, typ=typ, comment=com or "",
                         fam=family(src, typ)))
    obs[fire] = rows

# FEDS objects in the LNU sheet are separate fires of the complex: sum objects at the same time
lnu = obs["LNU 2020"]; feds_t = {}
for r in lnu:
    if r["src"] and r["src"].startswith("FEDS 2020") and r["acres"]:
        feds_t.setdefault(r["t"], 0); feds_t[r["t"]] += r["acres"]
obs["LNU 2020"] = [r for r in lnu if not (r["src"] or "").startswith("FEDS 2020")] + \
    [dict(t=t, acres=a, src="FEDS 2020 (sum of objects)", typ="Reconstruction: VIIRS satellite reconstruction", comment="", fam="satellite") for t, a in feds_t.items()]

# LNU component rows (single fires of the complex) are not comparable with whole-complex areas
obs["LNU 2020"] = [r for r in obs["LNU 2020"] if not ("component" in (r["src"] or "") or ("Larson" in (r["src"] or "") and "Whole" not in r["src"]))]
_cf = [r for r in obs["LNU 2020"] if (r["src"] or "").startswith("CAL FIRE 20200828") and r["acres"]]
if _cf:
    obs["LNU 2020"] = [r for r in obs["LNU 2020"] if not (r["src"] or "").startswith("CAL FIRE 20200828")] + \
        [dict(t=_cf[0]["t"], acres=sum(r["acres"] for r in _cf), src="CAL FIRE 20200828 (3 polygons summed; layer date only)", typ="Observation: Mixed methods / infrared", comment="", fam="mapped_dateonly")]

# Carr FEDS 12 h series from research_data (supplement lists FEDS for Carr as a reference only)
for f in json.load(open(PER / "Carr_2018_NASA_FEDS_perimeter_12h.geojson"))["features"]:
    t = datetime.fromisoformat(f["properties"]["t"]) - timedelta(hours=7)
    if t >= datetime(2018, 7, 23, 13):
        obs["Carr 2018"].append(dict(t=t, acres=f["properties"]["farea"] * KM2, src="NASA FEDS 12 h (research_data)", typ="Reconstruction: VIIRS satellite reconstruction", comment="", fam="satellite"))
# merge FEDS overpasses at the same time (several fire objects)
def merge_same_time(rows, tag):
    keep = [r for r in rows if r["src"] != tag]; agg = {}
    for r in rows:
        if r["src"] == tag: agg[r["t"]] = agg.get(r["t"], 0) + r["acres"]
    return keep + [dict(t=t, acres=a, src=tag, typ="Reconstruction: VIIRS satellite reconstruction", comment="", fam="satellite") for t, a in agg.items()]
obs["Carr 2018"] = merge_same_time(obs["Carr 2018"], "NASA FEDS 12 h (research_data)")

# Camp from the case study and research_data (PST = UTC-8)
camp = []
cwb = load_workbook(CAMP_WB, data_only=True); ws = cwb["All 209 fields"]; h = [c.value for c in ws[1]]
for r in ws.iter_rows(min_row=2, values_only=True):
    d = dict(zip(h, r))
    if d.get("REPORT_TO_DATE"):
        a = d.get("CURR_INCIDENT_AREA") or d.get("ACRES")
        camp.append(dict(t=datetime.fromisoformat(str(d["REPORT_TO_DATE"])), acres=float(a) if a else None, src="ICS-209-PLUS v3", typ="Observation: Incident report", comment="", fam="report"))
for f in json.load(open(PER / "Camp_2018_NIFC_GeoMAC_perimeter_history.geojson"))["features"]:
    p = f["properties"]
    if p.get("perimeterdatetime"):
        camp.append(dict(t=datetime.fromtimestamp(p["perimeterdatetime"] / 1000, timezone.utc).replace(tzinfo=None) - timedelta(hours=8),
                         acres=p["gisacres"], src="NIFC Open Data / GeoMAC", typ=p.get("mapmethod") or "", comment="", fam="mapped"))
for f in json.load(open(PER / "Camp_2018_NASA_FEDS_perimeter_12h.geojson"))["features"]:
    t = datetime.fromisoformat(f["properties"]["t"]) - timedelta(hours=8)
    if t >= datetime(2018, 11, 8, 6):
        camp.append(dict(t=t, acres=f["properties"]["farea"] * KM2, src="NASA FEDS 12 h (research_data)", typ="Reconstruction: VIIRS satellite reconstruction", comment="", fam="satellite"))
camp = merge_same_time(camp, "NASA FEDS 12 h (research_data)")
camp += [dict(t=datetime(2018, 11, 8, 6, 29), acres=None, src="CAL FIRE Green Sheet (research_data)", typ="Event: Agency start-time record", comment="Fire start 06:29, report 06:33", fam="event"),
         dict(t=datetime(2018, 11, 8, 6, 25), acres=None, src="Maranghides et al 2021, p.19", typ="Event: First observations 06:25 to 06:40", comment="", fam="event")]
obs["Camp 2018"] = camp
for k in obs: obs[k].sort(key=lambda r: r["t"])

# ---------------------------------------------------------------- reference times
IGN = {  # earliest credible start, and all start-time records (minutes are what matter)
 "Pioneer 2016": dict(ignition=datetime(2016, 7, 18, 17, 17), records=[("First report by Casner Fire lead plane (Gabbert 2016, p.2)", datetime(2016, 7, 18, 17, 17)),
                   ("ICS-209 discovery time", datetime(2016, 7, 18, 18, 9))]),
 "Carr 2018": dict(ignition=datetime(2018, 7, 23, 13, 9), records=[("Investigator ignition estimate (USFS 2021, pp.3, 6)", datetime(2018, 7, 23, 13, 9)),
                   ("CAL FIRE administrative start", datetime(2018, 7, 23, 13, 15)), ("ICS-209 discovery time", datetime(2018, 7, 23, 15, 0))]),
 "Camp 2018": dict(ignition=datetime(2018, 11, 8, 6, 25), records=[("NIST first observations (Maranghides et al 2021, p.19)", datetime(2018, 11, 8, 6, 25)),
                   ("CAL FIRE start (Green Sheet)", datetime(2018, 11, 8, 6, 29)), ("ICS-209 discovery time", datetime(2018, 11, 8, 8, 29))]),
 "LNU 2020": dict(ignition=datetime(2020, 8, 17, 6, 0), records=[("GOFER reference ignition time", datetime(2020, 8, 17, 6, 0)),
                   ("CAL FIRE incident update, 1 Sep 2020, p.1", datetime(2020, 8, 17, 6, 40)), ("CAL FIRE incident webpage", datetime(2020, 8, 17, 8, 0)),
                   ("ICS-209 discovery time", datetime(2020, 8, 17, 13, 57))]),
}
CONTAIN = {
 "Pioneer 2016": [("Rodriguez 2020 / 209 at 100%", "2016-11-03"), ("Clements 2018", "2016-11-04")],
 "Carr 2018": [("CAL FIRE webpage 09:37", "2018-08-30"), ("ICS-209 F report", "2018-08-30")],
 "Camp 2018": [("ICS-209 at 100%", "2018-11-25"), ("NICC annual table", "2018-11-25")],
 "LNU 2020": [("ICS-209 at 100%", "2020-09-30"), ("ICS-209 F report", "2020-10-01"), ("CAL FIRE webpage 10:38", "2020-10-02")],
}
HAZARD = {  # documented crew hazard windows (research_data Margin Timing Anchors / Camp Burnovers)
 "Carr 2018": [("Dozer burnover, fire whirl, 18:38 aerial warning, 19:40 mayday", datetime(2018, 7, 26, 17, 30), datetime(2018, 7, 26, 19, 40))],
 "Camp 2018": [("19 NIST burnovers", datetime(2018, 11, 8, 7, 50), datetime(2018, 11, 9, 7, 15)),
               ("Rattlesnake Flats hand crew burnover", datetime(2018, 11, 8, 14, 15), datetime(2018, 11, 8, 15, 30))],
}

R, T = {}, {"gaps": [], "staleness": [], "cadence_curve": [], "hourly_cover": [], "timestamp_flags": [], "source_pairs": []}
for fire, rows in obs.items():
    ign = IGN[fire]["ignition"]
    area = [r for r in rows if r["acres"] is not None]
    # timestamp sanity: area observations stamped before the earliest credible start
    flags = [r for r in area if r["t"] < ign]
    for r in flags: T["timestamp_flags"].append(dict(fire=fire, time=str(r["t"]), acres=round(r["acres"]), source=r["src"], method=r["typ"]))
    area = [r for r in area if r["t"] >= ign]
    spatial = [r for r in area if r["fam"] in ("mapped", "satellite")]
    mapped = [r for r in area if r["fam"] == "mapped"]
    def intervals(xs):
        ts = sorted({r["t"] for r in xs}); return [(a, b, (b - a).total_seconds() / 3600) for a, b in zip(ts, ts[1:])]
    cov = {}
    for name, xs in (("any area observation", area), ("spatial (mapped or satellite)", spatial), ("mapped perimeter only", mapped)):
        iv = intervals(xs)
        cov[name] = dict(n=len({r["t"] for r in xs}), interval_h=summ([g for _, _, g in iv]),
                         share_over_12h=(sum(1 for *_, g in iv if g > 12) / len(iv)) if iv else None)
    first_spatial = min((r["t"] for r in spatial), default=None)
    cov["hours_ignition_to_first_spatial"] = (first_spatial - ign).total_seconds() / 3600 if first_spatial else None
    fams = {}
    for r in area: fams[r["fam"]] = fams.get(r["fam"], 0) + 1
    cov["observations_by_family"] = fams
    # per-family series (one source family at a time, so source biases do not create fake jumps)
    def series(pred):
        d = {}
        for r in area:
            if pred(r): d[r["t"]] = max(d.get(r["t"], 0), r["acres"])
        return d
    fam_series = {"mapped (aircraft / GeoMAC)": series(lambda r: r["fam"] == "mapped"),
                  "satellite (GOFER or FEDS)": series(lambda r: r["fam"] == "satellite"),
                  "incident report (ICS-209)": series(lambda r: (r["src"] or "").startswith("ICS-209"))}
    base_name = "mapped (aircraft / GeoMAC)" if len(fam_series["mapped (aircraft / GeoMAC)"]) > 5 else "satellite (GOFER or FEDS)"
    base = fam_series[base_name]; ts = sorted(base); gaps = []
    for a, b in zip(ts, ts[1:]):
        g = (b - a).total_seconds() / 3600
        if g > 12:
            inside = [r for r in rows if a < r["t"] < b]
            ev = [f"{r['t']:%d %b %H:%M} {r['src']}: {r['comment'][:110]}" for r in inside if r["fam"] in ("event", "other") or (r["fam"] == "report" and not (r["src"] or "").startswith("ICS-209") and r["comment"])]
            hz = [h[0] for h in HAZARD.get(fire, []) if h[1] < b and h[2] > a]
            gaps.append(dict(fire=fire, record=base_name, start=str(a), end=str(b), hours=round(g, 1), acres_start=round(base[a]), acres_end=round(base[b]),
                             growth_acres=round(base[b] - base[a]),
                             satellite_obs_inside=sum(1 for r in inside if r["fam"] == "satellite") if base_name.startswith("mapped") else None,
                             reports_inside=sum(1 for r in inside if r["fam"] == "report" and r["acres"]),
                             documented_events_inside=" | ".join(ev[:4]), crew_hazard_inside="; ".join(hz)))
    T["gaps"] += sorted(gaps, key=lambda x: x["start"])
    R[fire] = dict(coverage=cov, gap_record=base_name, gaps_over_12h=len(gaps), longest_gaps=sorted(gaps, key=lambda x: -x["hours"])[:5],
                   gap_hours_total=sum(x["hours"] for x in gaps), growth_in_gaps=sum(max(0, x["growth_acres"]) for x in gaps),
                   gaps_uncovered_by_satellite=sum(1 for x in gaps if x["satellite_obs_inside"] == 0),
                   final_acres=max(fam_series["incident report (ICS-209)"].values()), timestamp_flags=len(flags))
    # staleness: at each new observation in a family, growth since the previous one in the same family
    R[fire]["staleness_pct_by_family"] = {}
    for nm, ser in fam_series.items():
        tt = sorted(ser); vals = []
        for a, b in zip(tt, tt[1:]):
            gr = ser[b] - ser[a]
            if ser[a] > 0 and gr > 0: vals.append(100 * gr / ser[b])
            T["staleness"].append(dict(fire=fire, family=nm, prev=str(a), next=str(b), hours=round((b - a).total_seconds() / 3600, 2), growth_acres=round(gr), pct_of_next=round(100 * gr / ser[b], 2) if ser[b] else None))
        R[fire]["staleness_pct_by_family"][nm] = summ(vals)
    # information age at each documented crew hazard: newest observation of each family before the hazard began
    R[fire]["hazard_information_age"] = []
    for name, h0, h1 in HAZARD.get(fire, []):
        row = dict(fire=fire, hazard=name, hazard_start=str(h0), hours_since_ignition=round((h0 - ign).total_seconds() / 3600, 2))
        for nm, ser in fam_series.items():
            before = [t for t in ser if t <= h0]; after = [t for t in ser if t > h0]
            row[f"{nm}: newest before"] = str(max(before)) if before else "none since ignition"
            row[f"{nm}: age h"] = round((h0 - max(before)).total_seconds() / 3600, 2) if before else round((h0 - ign).total_seconds() / 3600, 2)
            row[f"{nm}: size then"] = round(ser[max(before)]) if before else 0
            row[f"{nm}: next after"] = str(min(after)) if after else None
            row[f"{nm}: size next"] = round(ser[min(after)]) if after else None
        R[fire]["hazard_information_age"].append(row); T.setdefault("hazard_info_age", []).append(row)
    # time-of-day coverage of spatial observations vs growth (local hour)
    hours = [0] * 24
    for t in {r["t"] for r in mapped}: hours[t.hour] += 1
    night = sum(hours[h] for h in list(range(0, 6)) + list(range(20, 24)))
    R[fire]["mapped_obs_night_share"] = night / sum(hours) if sum(hours) else None
    T["hourly_cover"] += [dict(fire=fire, hour=hh, mapped_perimeters=n) for hh, n in enumerate(hours)]
    # start-time and containment uncertainty
    recs = IGN[fire]["records"]
    R[fire]["start_records"] = [(n, str(t)) for n, t in recs]
    R[fire]["start_spread_min"] = (max(t for _, t in recs) - min(t for _, t in recs)).total_seconds() / 60
    R[fire]["start_spread_excl_209_min"] = (max(t for n, t in recs if "209" not in n) - min(t for n, t in recs if "209" not in n)).total_seconds() / 60 if len([1 for n, _ in recs if "209" not in n]) > 1 else 0
    cd = sorted(datetime.fromisoformat(d) for _, d in CONTAIN[fire])
    R[fire]["containment_records"] = CONTAIN[fire]; R[fire]["containment_spread_days"] = (cd[-1] - cd[0]).days

# ---------------------------------------------------------------- revisit (cadence) error curve from GOFER hourly (LNU)
g = sorted([r for r in obs["LNU 2020"] if r["src"] == "GOFER-Combined v0.2"], key=lambda r: r["t"])
A = {r["t"]: r["acres"] for r in g}; t0, t1 = g[0]["t"], g[-1]["t"]
hrs = [t0 + timedelta(hours=i) for i in range(int((t1 - t0).total_seconds() // 3600) + 1)]
curve = []
for c in (1, 2, 3, 6, 12, 24):
    for window, lo, hi in (("whole record 17 to 28 Aug", t0, t1), ("growth phase 17 to 21 Aug", t0, datetime(2020, 8, 21, 23))):
        errs, pcts = [], []
        for phase in range(c):
            obs_times = [t for i, t in enumerate(hrs) if (i - phase) % c == 0]
            for t in hrs:
                if not (lo <= t <= hi): continue
                prev = [o for o in obs_times if o <= t]
                if not prev: continue
                e = A[t] - A[prev[-1]]; errs.append(e); pcts.append(100 * e / A[t] if A[t] else 0)
        row = dict(cadence_h=c, window=window, mean_acres=st.mean(errs), p90_acres=q(errs, .9), max_acres=max(errs), mean_pct=st.mean(pcts), p90_pct=q(pcts, .9), max_pct=max(pcts))
        curve.append(row); T["cadence_curve"].append({k: (round(v, 2) if isinstance(v, float) else v) for k, v in row.items()})
growth = [(b["t"], b["acres"] - a["acres"]) for a, b in zip(g, g[1:])]
night_g = sum(x for t, x in growth if t.hour >= 20 or t.hour < 6); tot_g = sum(x for _, x in growth)
R["GOFER_LNU"] = dict(hours=len(g), start=str(t0), end=str(t1), end_acres=g[-1]["acres"], peak_hourly_growth=max(growth, key=lambda x: x[1]),
                      top_hours=sorted(growth, key=lambda x: -x[1])[:6], night_share_of_growth=night_g / tot_g,
                      night_def="20:00 to 06:00 PDT", cadence_curve=curve)
# GOFER vs other sources at matched times (+/- 1 h)
def nearest(series, t, tol_h):
    c = [r for r in series if abs((r["t"] - t).total_seconds()) <= tol_h * 3600 and r["acres"]]
    return min(c, key=lambda r: abs((r["t"] - t).total_seconds())) if c else None
pairs = []
for r in obs["LNU 2020"]:
    if r["src"] == "GOFER-Combined v0.2" or not r["acres"]: continue
    gg = nearest(g, r["t"], 1)
    if gg:
        pairs.append(dict(fire="LNU 2020", time=str(r["t"]), source=r["src"][:60], family=r["fam"], acres=round(r["acres"]), gofer=round(gg["acres"]), gofer_over_source=round(gg["acres"] / r["acres"], 3)))
T["source_pairs"] += pairs
by = {}
for p in pairs: by.setdefault(p["family"] if "FEDS" not in p["source"] else "satellite (FEDS)", []).append(p["gofer_over_source"])
R["GOFER_LNU"]["ratio_vs_sources"] = {k: summ(v) for k, v in by.items()}

json.dump(R, open(OUT / "results.json", "w"), indent=1, default=str)
for name, rows in T.items():
    if rows:
        with open(OUT / f"{name}.csv", "w", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys())); w.writeheader(); w.writerows(rows)
print(json.dumps({k: (v if k == "GOFER_LNU" else {kk: vv for kk, vv in v.items() if kk != "longest_gaps"}) for k, v in R.items()}, indent=1, default=str)[:9000])
