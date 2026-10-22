"""Evaluate the station weather in ../Weather Data against drone wind limits (U21 drone wind gate).

For every station: build an hourly local-time series of sustained wind (10-min mean at nominal 20 ft),
peak gust, direction, temperature and RH. RAWS rows are used as reported (one 10-min mean per hour, gust
= peak wind in the prior hour). ASOS one-minute rows are reduced to the hourly maximum of the complete
10-min estimates (README method) and the hourly maximum gust; METAR rows only fill hours with no
one-minute data. Nothing is interpolated.

The gate is reported as a band, which is the uncertainty:
  optimistic  = fly if the hourly sustained wind is at or under the limit
  conservative = fly if the hourly peak gust is at or under the limit
Both are surface (20 ft) winds; wind at flight altitude is usually stronger, so P(fly) is an upper bound.

Run:
    uv run --with pandas --with openpyxl python weather_evaluation.py
"""
import json
from pathlib import Path
import pandas as pd

HERE = Path(__file__).parent
WX = HERE.parent / "Weather Data"
OUT = HERE / "weather_eval"; OUT.mkdir(exist_ok=True)
MS_TO_MPH = 2.236936

# Published or project-recorded limits (Week 5 / Week 6 Platform tabs). Type 1 fixed wing has no published
# limit in project records, so it runs as a 30 and 40 mph sensitivity.
LIMITS = {"Alta X (22 mph)": 22, "SwitchBlade-Elite (25 mph)": 25, "Type 1 FW if 30 mph": 30, "Type 1 FW if 40 mph": 40}

# Critical windows, local time, from the v0.5 case studies.
# Standard evaluation unit (20 Oct 2026): one 24-hour day, 06:00 to 06:00 local, on each fire's peak-run day.
# Same length for every fire and mission, one full night included. End is inclusive to 05:59.
WINDOWS = {
    "Camp": ("2018-11-08 06:00", "2018-11-09 05:59", "8 to 9 Nov: first-day run and initial attack, 19 burnovers 07:50 to 15:15"),
    "Carr": ("2018-07-26 06:00", "2018-07-27 05:59", "26 to 27 Jul: run and fire whirl, warning 18:38, mayday 19:40"),
    "Pioneer": ("2016-08-29 06:00", "2016-08-30 05:59", "29 to 30 Aug: run of about 29,000 ac"),
    "LNU": ("2020-08-18 06:00", "2020-08-19 05:59", "18 to 19 Aug: overnight merge and 12 mile run"),
}
PERIM = {  # perimeter windows from the Weather Data README
    "Pioneer": ("2016-07-19 07:00", "2016-09-28 18:00"),
    "LNU": ("2020-08-17 09:00", "2020-08-28 14:00"),
    "Carr": ("2018-07-23 11:30", "2018-08-31 00:54"),
    "Camp": ("2018-11-08 09:54", "2018-11-25 10:22"),
}


def col(df, *keys):
    for c in df.columns:
        if all(k.lower() in str(c).lower() for k in keys):
            return c
    return None


def hourly(path: Path) -> pd.DataFrame:
    df = pd.read_excel(path, sheet_name="Weather")
    t = df.columns[0]
    df[t] = pd.to_datetime(df[t])
    est = col(df, "estimated wind speed")
    if est:  # ASOS
        prod = col(df, "observation product")
        one = df[df[prod].astype(str).str.lower().str.contains("1min")] if prod else df
        met = df.drop(one.index)
        g = col(df, "gust speed")  # absent in the RDD, APC, STS and VCB files
        dc = col(df, "direction")
        orig = col(df, "original wind speed")

        def agg(part, speed):
            a = part.set_index(t).resample("1h")
            return pd.DataFrame({"sust": a[speed].max(), "gust": a[g].max() if g else float("nan"), "dir": a[dc].median()})
        h1 = agg(one, est)
        hm = agg(met, orig)
        hm["sust"] = hm["sust"] / 1.15  # same 10 m to 20 ft divisor as the README; METAR is a 2-min mean
        h = h1.combine_first(hm)
        tc, rc = col(df, "air temperature"), col(df, "relative humidity")
        if tc: h["temp"] = df.set_index(t)[tc].resample("1h").mean()
        if rc: h["rh"] = df.set_index(t)[rc].resample("1h").min()
    else:  # RAWS or DCP
        s = col(df, "wind speed")
        g = col(df, "peak wind") or col(df, "gust")
        d = df.set_index(t)
        h = pd.DataFrame({"sust": d[s], "gust": d[g] if g else None, "dir": d[col(df, "direction")]})
        tc, rc = col(df, "air temperature"), col(df, "relative humidity")
        h["temp"] = d[tc] if tc else None
        h["rh"] = d[rc] if rc else None
        fm = col(df, "fuel moisture")  # 10-h fuel stick on RAWS; absent at Pioneer
        h["fuel"] = d[fm] if fm else None
        h = h.resample("1h").agg({"sust": "max", "gust": "max", "dir": "median", "temp": "mean", "rh": "min", "fuel": "min"})
    h[["sust", "gust"]] = h[["sust", "gust"]] * MS_TO_MPH
    return h


def gate(h: pd.DataFrame) -> dict:
    v = h.dropna(subset=["sust"])
    out = {"hours_with_wind": int(len(v))}
    for name, lim in LIMITS.items():
        opt = (v["sust"] <= lim).mean() if len(v) else None
        gv = v.dropna(subset=["gust"])
        con = (gv["gust"] <= lim).mean() if len(gv) else None
        out[name] = {"p_fly_sustained": None if opt is None else round(float(opt), 3),
                     "p_fly_gust": None if con is None else round(float(con), 3)}
    return out


def nightshare(h):  # 20:00 to 06:00 local
    return h[(h.index.hour >= 20) | (h.index.hour < 6)]


results, rows = {}, []
for fire in WINDOWS:
    st = pd.read_excel(WX / fire / "Stations.xlsx")
    dist = col(st, "distance")
    results[fire] = {}
    for _, r in st.iterrows():
        sid = r["Station ID"]
        f = WX / fire / f"{sid}_Weather.xlsx"
        if not f.exists():
            continue
        h = hourly(f)
        a, b, why = WINDOWS[fire]
        cw = h.loc[a:b]
        pa, pb = PERIM[fire]
        pw = h.loc[pa:pb]
        rec = {
            "name": r["Station name"], "network": str(r.get("Network", r.get("Network / data type", ""))),
            "distance_km": None if dist is None or pd.isna(r[dist]) else round(float(r[dist]), 1),
            "critical_window": f"{a} to {b} ({why})",
            "crit_hours_with_wind": int(cw["sust"].notna().sum()),
            "crit_sust_max_mph": None if cw["sust"].isna().all() else round(float(cw["sust"].max()), 1),
            "crit_sust_mean_mph": None if cw["sust"].isna().all() else round(float(cw["sust"].mean()), 1),
            "crit_gust_max_mph": None if cw["gust"].isna().all() else round(float(cw["gust"].max()), 1),
            "crit_rh_min": None if "rh" not in cw or cw["rh"].isna().all() else round(float(cw["rh"].min()), 0),
            "crit_fuel_moisture_min": None if "fuel" not in cw or cw["fuel"].isna().all() else round(float(cw["fuel"].min()), 0),
            "crit_dir_median": None if cw["dir"].isna().all() else round(float(cw["dir"].median()), 0),
            "crit_gate": gate(cw),
            "perimeter_gate": gate(pw),
            "perimeter_night_gate": gate(nightshare(pw)),
        }
        results[fire][sid] = rec
        cw2 = cw.copy(); cw2["fire"] = fire; cw2["station"] = sid
        rows.append(cw2.reset_index().rename(columns={cw2.index.name or "index": "local_time"}))
        print(fire, sid, rec["distance_km"], rec["crit_hours_with_wind"], rec["crit_sust_max_mph"], rec["crit_gust_max_mph"], rec["crit_rh_min"],
              "AltaX crit sust/gust", rec["crit_gate"]["Alta X (22 mph)"], "perim AltaX", rec["perimeter_gate"]["Alta X (22 mph)"])

json.dump(results, open(OUT / "results.json", "w"), indent=1, default=str)
pd.concat(rows).to_csv(OUT / "critical_window_hourly.csv", index=False, float_format="%.2f")
print("wrote", OUT)
