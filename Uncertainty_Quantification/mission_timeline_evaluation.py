"""Hour-by-hour mission timelines and night statistics, from project data only (no GOFER).

Inputs, all inside the project folder:
  perimeter_supplements/Wildfire Perimeter Time Area Sources.xlsx  (Pioneer, Carr, LNU; GOFER rows dropped)
  perimeter_supplements/Camp_2018_Perimeter_Time_Area_ICS209PLUS.xlsx
  weather_eval/critical_window_hourly.csv and results.json        (from weather_evaluation.py)
  ../Weather Data/<fire>/<station>_Weather.xlsx                  (hourly series for the timeline windows)

For each mission one fire window is laid out hour by hour: the newest aircraft perimeter and its age, the newest
ICS-209 size and its age, what the records say happened, the station wind and humidity, whether an Alta X (22 mph)
and a Type 1 fixed wing at 30 mph could fly by gusts, and the drone task proposed for that hour.

Night statistics per fire (night = 20:00 to 06:00 local, as in v0.5): share of aircraft perimeters stamped at night,
and share of growth between consecutive aircraft perimeters that falls in night hours (growth spread evenly over
each interval, an assumption). LNU has no aircraft perimeter series in the record, so its night growth is not computed.

Run after weather_evaluation.py:
    uv run --with pandas --with openpyxl python mission_timeline_evaluation.py
"""
import json
from pathlib import Path
import pandas as pd

HERE = Path(__file__).parent
SUP = HERE / "perimeter_supplements"
WX = HERE.parent / "Weather Data"
OUT = HERE / "timeline_eval"; OUT.mkdir(exist_ok=True)
MPH = 2.236936


def load_records():
    recs = {}
    for sheet in ["Pioneer 2016", "LNU 2020", "Carr 2018"]:
        recs[sheet.split()[0]] = pd.read_excel(SUP / "Wildfire Perimeter Time Area Sources.xlsx", sheet_name=sheet)
    recs["Camp"] = pd.read_excel(SUP / "Camp_2018_Perimeter_Time_Area_ICS209PLUS.xlsx", sheet_name="Camp 2018")
    for k, df in recs.items():
        df = df.iloc[:, :5].copy(); df.columns = ["t", "area", "src", "type", "comment"]
        gofer = df["src"].astype(str).str.contains("GOFER", case=False) | df["comment"].astype(str).str.contains("GOFER", case=False)
        df = df[~gofer]
        df["t"] = pd.to_datetime(df["t"], errors="coerce"); df["area"] = pd.to_numeric(df["area"], errors="coerce")
        recs[k] = df.dropna(subset=["t"]).sort_values("t")
    return recs


def family(df, kind):
    s = df["src"].astype(str)
    m = {"aircraft": s.str.contains("GeoMAC"), "ics209": s.str.contains("ICS-209-PLUS") & df["type"].astype(str).str.startswith("Observation"),
         "satellite": s.str.contains("FEDS")}[kind]
    return df[m & df["area"].notna()]


def hourly_weather(fire, station):
    df = pd.read_excel(WX / fire / f"{station}_Weather.xlsx", sheet_name="Weather")
    t = df.columns[0]; df[t] = pd.to_datetime(df[t]); d = df.set_index(t)
    sp = [c for c in df.columns if "wind speed" in c.lower()][0]
    gu = [c for c in df.columns if "peak wind" in c.lower()][0]
    rh = [c for c in df.columns if "relative humidity" in c.lower()][0]
    h = pd.DataFrame({"sust": d[sp] * MPH, "gust": d[gu] * MPH, "rh": d[rh]}).resample("1h").agg({"sust": "max", "gust": "max", "rh": "min"})
    return h


def night(ts):
    return ts.hour >= 20 or ts.hour < 6


# Mission windows, event labels (condensed from the record comments at those times) and proposed drone tasks.
MISSIONS = {  # every window is the fire's standard day, 06:00 to 05:00 local (24 hourly rows)
    "M1": dict(fire="Carr", station="MMOC1", start="2018-07-26 06:00", end="2018-07-27 05:00",
               events={"2018-07-26 07:00": "ICS-209 20,000 ac, 10% contained",
                       "2018-07-26 14:00": "Aircraft perimeter 48,237 ac",
                       "2018-07-26 17:00": "17:30 run east at 2.5 mph, spotting over 1 mile; 17:34 perimeter 44,484 ac",
                       "2018-07-26 19:00": "19:17 spots east of Sacramento River; 19:20 to 20:00 vortex on radar; 19:45 ICS-209 28,763 ac",
                       "2018-07-26 20:00": "ICS-209 still 28,763 ac (20:15, 21:00)",
                       "2018-07-27 00:00": "Aircraft perimeter 50,050 ac; spots south of Whiskeytown Lake"},
               tasks=[(6, 12, "Fresh perimeter for the morning plan; decide where the line and resources go"),
                      (12, 17, "Type 1 holds whole-fire perimeter; Type 3 maps the active east flank"),
                      (17, 20, "Track the run hourly: head, spots across the river; push the line to IC and divisions"),
                      (20, 24, "Night IR perimeter by Type 1 so the morning plan starts from a current line"),
                      (0, 6, "Night IR passes on spots and flanks; fresh perimeter before the briefing")]),
    "M2": dict(fire="Camp", station="JBGC1", start="2018-11-08 06:00", end="2018-11-09 05:00",
               events={"2018-11-08 06:00": "06:25 first 911 call; 06:44 first engine, about 35 mph wind; 06:45 second ignition",
                       "2018-11-08 07:00": "07:10 engine estimate 200 to 300 ac; 07:37 Concow order; 07:50 first burnovers; 07:52 air attack 1,000 ac",
                       "2018-11-08 08:00": "08:01 'fire is now in Paradise', 2,000 ac; 08:29 ICS-209 discovery time",
                       "2018-11-08 09:00": "09:03 Paradise and Magalia under orders; 09:54 first aircraft perimeter 54,586 ac",
                       "2018-11-08 10:00": "10:45 Landsat 8 image (no area given)",
                       "2018-11-08 16:00": "Aircraft perimeter 70,454 ac; satellite 81,810 ac",
                       "2018-11-08 18:00": "Green Sheet: first 12 h ran 15 mi, over 55,000 ac, spotting up to a mile ahead; 18:30 first ICS-209 16,000 ac",
                       "2018-11-08 19:00": "19:20 ICS-209 18,000 ac; long-range spotting reported",
                       "2018-11-09 04:00": "Satellite 92,880 ac"},
               tasks=[(6, 7, "Initial attack: first look and size-up from the engine or station drone"),
                      (7, 9, "Search ahead of the head for spots toward Concow and Paradise; geolocate each for engines"),
                      (9, 18, "Wide-area IR search downwind; confirm spots that threaten evacuation routes"),
                      (18, 24, "Night IR search for spots outside the line; send positions to night operations"),
                      (0, 6, "Night IR search; list confirmed spots for the morning briefing")]),
    "M3": dict(fire="Camp", station="JBGC1", start="2018-11-08 06:00", end="2018-11-09 05:00",
               events={"2018-11-08 06:00": "06:44 first engine; about 35 mph wind",
                       "2018-11-08 07:00": "07:37 Concow order; 07:50 first burnovers (Hoffman Rd, Concow Rd)",
                       "2018-11-08 08:00": "08:01 fire in Paradise; burnovers continue",
                       "2018-11-08 09:00": "09:03 Paradise and Magalia under orders; 09:54 first aircraft perimeter",
                       "2018-11-08 14:00": "14:15 firing operation, Rattlesnake Flats Rd; 14:45 wind shift, fire crosses road around crew",
                       "2018-11-08 15:00": "15:00 emergency call unanswered; 15:15 burnover",
                       "2018-11-08 16:00": "Aircraft perimeter 70,454 ac",
                       "2018-11-08 18:00": "18:30 first ICS-209: 16,000 ac"},
               tasks=[(6, 8, "Watch crews and engines on Concow Rd; check escape routes ahead of the head"),
                      (8, 14, "Overwatch of crews in Paradise; mark safety zones and blocked routes; relay warnings"),
                      (14, 16, "Before and during firing: wind and fire position over Rattlesnake Flats; relay when radios saturate"),
                      (16, 20, "Reassess safety zones and helicopter landing zones for the night"),
                      (20, 24, "Night overwatch of crews on the line; thermal view replaces the lookout"),
                      (0, 6, "Night overwatch; confirm escape routes are open for the morning shift")]),
    "M4": dict(fire="LNU", station="ATLC1", start="2020-08-18 06:00", end="2020-08-19 05:00",
               events={"2020-08-18 12:00": "Satellite (VIIRS) 7,328 ac for the main object",
                       "2020-08-18 18:00": "Atlas Peak 29 mph sustained, 41 mph gusts",
                       "2020-08-18 20:00": "Overnight 18 to 19 Aug: fires merge in a 12-mile run (agency chronology)",
                       "2020-08-19 00:00": "Satellite (VIIRS) 41,602 ac for the main object"},
               tasks=[(6, 14, "Morning: drones map which drops held and where to drop next"),
                      (14, 20, "Daylight: helicopter and airtanker drops; drones map where drops held"),
                      (20, 24, "Night: crewed aircraft stand down; night water drop is a large UAS or OPA concept only"),
                      (0, 6, "Night: drone IR shows where line held, ready for first light drops")]),
    "M5": dict(fire="Pioneer", station="BCFI1", start="2016-08-29 06:00", end="2016-08-30 05:00",
               events={"2016-08-29 06:00": "00:00 NIROPS boundary labelled (Kingsmill et al.); about 29,000 ac of growth on 29 Aug",
                       "2016-08-29 18:00": "18:16 aircraft perimeter 140,778 ac; 19:00 ICS-209 111,604 ac, 58% contained",
                       "2016-08-30 00:00": "00:16 NIROPS boundary labelled (Kingsmill et al.)"},
               tasks=[(6, 12, "Resupply crews on remote lines (water, batteries, radios) before the afternoon"),
                      (12, 16, "Scout helicopter landing zones and drop points ahead of the run"),
                      (16, 21, "Hold resupply; keep landing zones and safety zones current as the run builds"),
                      (21, 24, "Night resupply to crews holding line, if authorised"),
                      (0, 6, "Night: stage loads and check landing zones for first light")]),
}


def task_for(tasks, h):
    for a, b, txt in tasks:
        if a <= h < b:
            return txt
    return ""


def build_timeline(mid, m, recs):
    df = recs[m["fire"]]
    air, ics = family(df, "aircraft"), family(df, "ics209")
    wx = hourly_weather(m["fire"], m["station"])
    rows = []
    for ts in pd.date_range(m["start"], m["end"], freq="1h"):
        end = ts + pd.Timedelta(hours=1)
        a = air[air.t < end].tail(1); i = ics[ics.t < end].tail(1)
        w = wx.loc[ts] if ts in wx.index else None
        ev = m["events"].get(ts.strftime("%Y-%m-%d %H:%M"), "")
        rows.append({
            "hour": ts.strftime("%d %b %H:00"), "period": "Night" if night(ts) else "Day",
            "aircraft_perimeter": "none yet" if a.empty else f"{a.area.iloc[0]:,.0f} ac, {(end - a.t.iloc[0]).total_seconds() / 3600:.0f} h old",
            "ics209": "none yet" if i.empty else f"{i.area.iloc[0]:,.0f} ac, {(end - i.t.iloc[0]).total_seconds() / 3600:.0f} h old",
            "events": ev,
            "sust": None if w is None or pd.isna(w["sust"]) else round(float(w["sust"])),
            "gust": None if w is None or pd.isna(w["gust"]) else round(float(w["gust"])),
            "rh": None if w is None or pd.isna(w["rh"]) else round(float(w["rh"])),
            "alta_x": None if w is None or pd.isna(w["gust"]) else ("fly" if w["gust"] <= 22 else "no"),
            "type1_30": None if w is None or pd.isna(w["gust"]) else ("fly" if w["gust"] <= 30 else "no"),
            "task": task_for(m["tasks"], ts.hour),
        })
    out = pd.DataFrame(rows); out.to_csv(OUT / f"{mid}_{m['fire']}_hourly.csv", index=False)
    return out


def night_stats(recs):
    res = {}
    for fire, df in recs.items():
        air = family(df, "aircraft").drop_duplicates("t")
        if air.empty:
            res[fire] = {"aircraft_perimeters": 0, "night_stamped_share": None, "night_growth_share": None}
            continue
        ns = float(air.t.apply(night).mean())
        grow_n = grow_t = 0.0
        for (t0, a0), (t1, a1) in zip(air[["t", "area"]].values[:-1], air[["t", "area"]].values[1:]):
            g = a1 - a0
            if g <= 0:
                continue
            hrs = pd.date_range(pd.Timestamp(t0), pd.Timestamp(t1), freq="15min")[:-1]
            if len(hrs) == 0:
                continue
            frac = sum(night(h) for h in hrs) / len(hrs)
            grow_n += g * frac; grow_t += g
        res[fire] = {"aircraft_perimeters": int(len(air)), "night_stamped_share": round(ns, 3),
                     "night_growth_share": round(grow_n / grow_t, 3) if grow_t else None}
    return res


recs = load_records()
for mid, m in MISSIONS.items():
    t = build_timeline(mid, m, recs)
    print(mid, m["fire"], len(t), "rows; Alta X fly hours", (t.alta_x == "fly").sum(), "; Type 1 at 30 fly hours", (t.type1_30 == "fly").sum())
ns = night_stats(recs)
json.dump(ns, open(OUT / "night_stats.json", "w"), indent=1)
print(json.dumps(ns))
