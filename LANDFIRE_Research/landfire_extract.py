"""LANDFIRE LF 2016 Remap inside each benchmark fire's final perimeter.

Layers (lfps.usgs.gov/arcgis/rest/services/Landfire_LF2016/LF2016_<layer>_CONUS/ImageServer, public, 30 m, EPSG:5070):
FBFM40 (Scott and Burgan fuel models), FCCS (fuelbeds), EVT (existing vegetation type), CC (canopy cover %),
CBD (canopy bulk density, kg m-3 x 100), CBH (canopy base height, m x 10), CH (canopy height, m x 10),
FDist (fuel disturbance code; lookup LF16_FDst_200.csv). Units from the LANDFIRE Data Dictionary (lookup/LF_Data_Dictionary.pdf).

Code tables: lookup/LF16_F40_200.csv, LF16_FCCS_200.csv, LF16_EVT_200.csv, LF16_FDst_200.csv (landfire.gov CSV).
Surface fuel loading per pixel = Scott and Burgan (2005) Table 7 loads for that pixel's fuel model
(lookup/scott_burgan_2005_table7.csv). This is the model's assigned load, not a measured load.

LF 2016 Remap represents circa 2016 conditions. It predates Carr, Camp (2018) and LNU (2020). For Pioneer (2016) it
may include the fire's own disturbance; check the FDst summary before reading Pioneer fuels as pre-fire.

Run: uv run --with numpy --with tifffile --with shapely --with pyproj python landfire_extract.py
Outputs: landfire_summary.json, landfire_<fire>_<table>.csv, cache/ (GeoTIFFs, deletable, re-downloaded on rerun).
"""
import csv, json, urllib.parse, urllib.request
from collections import Counter
from pathlib import Path
import numpy as np, tifffile
from shapely.geometry import shape
from shapely.ops import unary_union
from shapely import contains_xy, transform as stransform
from pyproj import Transformer

HERE = Path(__file__).parent
LK = HERE / "lookup"
CACHE = HERE / "cache"; CACHE.mkdir(exist_ok=True)
PER = HERE.parents[1] / "sources" / "perimeters"
SVC = "https://lfps.usgs.gov/arcgis/rest/services/Landfire_LF2016/LF2016_{}_CONUS/ImageServer/exportImage"
FIRES = {"Pioneer 2016": "Pioneer_2016", "Carr 2018": "Carr_2018", "Camp 2018": "Camp_2018", "LNU 2020": "LNU_2020"}
LAYERS = ["FBFM40", "FCCS", "EVT", "CC", "CBD", "CBH", "CH", "FDist"]
PIX = 30.0
to5070 = Transformer.from_crs(4326, 5070, always_xy=True).transform
log = open(HERE / "landfire_query_log.txt", "a")


def table(name, key="VALUE"):
    return {int(float(r[key])): r for r in csv.DictReader(open(LK / name, encoding="latin-1"))}


F40, FCCS, EVT, FDST = table("LF16_F40_200.csv"), table("LF16_FCCS_200.csv"), table("LF16_EVT_200.csv"), table("LF16_FDst_200.csv")
SB = {r["FBFM40"]: r for r in csv.DictReader(open(LK / "scott_burgan_2005_table7.csv"))}
GROUP = {"NB": "Non-burnable", "GR": "Grass", "GS": "Grass-shrub", "SH": "Shrub", "TU": "Timber-understory", "TL": "Timber litter", "SB": "Slash-blowdown"}


def fetch(stem, layer, bbox, w, h, tile=2000):
    """Raw codes for bbox at 30 m. Large areas are requested in tiles (the service returns HTTP 500 above a few
    million pixels) and stitched; the stitched array is cached as .npy."""
    f = CACHE / f"{stem}_{layer}.npy"
    if f.exists():
        return np.load(f)
    out = np.full((h, w), -9999, dtype=np.int32)
    for r0 in range(0, h, tile):
        for c0 in range(0, w, tile):
            tw, th = min(tile, w - c0), min(tile, h - r0)
            tb = (bbox[0] + c0 * PIX, bbox[3] - (r0 + th) * PIX, bbox[0] + (c0 + tw) * PIX, bbox[3] - r0 * PIX)
            q = {"bbox": ",".join(f"{v:.1f}" for v in tb), "bboxSR": 5070, "imageSR": 5070, "size": f"{tw},{th}", "format": "tiff",
                 "pixelType": "S32" if layer == "FCCS" else "S16", "noData": -9999, "interpolation": "RSP_NearestNeighbor", "f": "image"}
            url = SVC.format(layer) + "?" + urllib.parse.urlencode(q)
            log.write(url + "\n")
            tmp = CACHE / "_tile.tif"
            urllib.request.urlretrieve(url, tmp)
            out[r0:r0 + th, c0:c0 + tw] = tifffile.imread(tmp)
    np.save(f, out)
    return out


def load_tac(code):
    r = SB.get(code)
    return None if r is None else sum(float(r[k]) for k in ("load_1h_tac", "load_10h_tac", "load_100h_tac", "load_live_herb_tac", "load_live_woody_tac"))


def share_table(vals, names, fire, tag, top=15, extra=None):
    c = Counter(vals.tolist()); n = sum(c.values()); rows = []
    for v, k in c.most_common(top):
        row = {"value": v, "name": names(v), "pixels": k, "share": round(k / n, 4), "acres": round(k * PIX * PIX / 4046.86)}
        if extra: row.update(extra(v))
        rows.append(row)
    with open(HERE / f"landfire_{fire.split()[0]}_{tag}.csv", "w", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=list(rows[0])); w.writeheader(); w.writerows(rows)
    return rows


out = {}
for label, stem in FIRES.items():
    gj = json.load(open(PER / f"{stem}.geojson"))
    geom = stransform(unary_union([shape(f["geometry"]) for f in gj["features"]]), lambda xy: np.column_stack(to5070(xy[:, 0], xy[:, 1])))
    x0, y0, x1, y1 = geom.bounds
    w, h = int(np.ceil((x1 - x0) / PIX)), int(np.ceil((y1 - y0) / PIX))
    bbox = (x0, y0, x0 + w * PIX, y0 + h * PIX)
    xs = bbox[0] + (np.arange(w) + 0.5) * PIX; ys = bbox[3] - (np.arange(h) + 0.5) * PIX
    inside = contains_xy(geom, *np.meshgrid(xs, ys))
    R = {L: fetch(stem, L, bbox, w, h) for L in LAYERS}
    np.save(CACHE / f"{stem}_mask.npy", inside); json.dump({"bbox": bbox, "w": w, "h": h}, open(CACHE / f"{stem}_grid.json", "w"))
    m = inside & (R["FBFM40"] > 0)
    f40 = R["FBFM40"][m]
    codes = np.vectorize(lambda v: F40.get(int(v), {}).get("FBFM40", "?"))(f40)
    grp = Counter(GROUP.get(c[:2], "Other") for c in codes.tolist()); n = sum(grp.values())
    loads = np.array([load_tac(c) if load_tac(c) is not None else 0.0 for c in codes.tolist()])
    burnable = np.array([not c.startswith("NB") for c in codes.tolist()])
    rows_f40 = share_table(f40, lambda v: F40.get(v, {}).get("FBFM40", "?"), label, "FBFM40", top=20,
                           extra=lambda v: {"surface_load_tac": load_tac(F40.get(v, {}).get("FBFM40", "")) or 0.0})
    rows_fccs = share_table(R["FCCS"][inside & (R["FCCS"] >= 0)], lambda v: FCCS.get(v, {}).get("FUELBED", "?"), label, "FCCS", top=12)
    rows_evt = share_table(R["EVT"][inside & (R["EVT"] > 0)], lambda v: EVT.get(v, {}).get("EVT_NAME", "?"), label, "EVT", top=12)
    fdn = lambda v: " / ".join(FDST.get(v, {}).get(k, "?") for k in ("D_TYPE", "D_SEVERITY", "D_TIME"))
    rows_fd = share_table(R["FDist"][inside & (R["FDist"] >= 0)], fdn, label, "FDist", top=10)
    forest = inside & (R["CC"] > 0)
    def stat(a, scale):
        a = a.astype(float) / scale
        return {"median": round(float(np.median(a)), 2), "p10": round(float(np.percentile(a, 10)), 2), "p90": round(float(np.percentile(a, 90)), 2)}
    out[label] = {
        "pixels": int(inside.sum()), "acres": round(inside.sum() * PIX * PIX / 4046.86),
        "fuel_group_share": {k: round(v / n, 3) for k, v in grp.most_common()},
        "top_fuel_models": [(r["name"], r["share"]) for r in rows_f40[:6]],
        "surface_load_tac_burnable": {"mean": round(float(loads[burnable].mean()), 2), "median": round(float(np.median(loads[burnable])), 2),
                                      "p90": round(float(np.percentile(loads[burnable], 90)), 2)},
        "top_fccs": [(r["name"], r["share"]) for r in rows_fccs[:5]],
        "top_evt": [(r["name"], r["share"]) for r in rows_evt[:6]],
        "canopy_cover_pct_where_canopy": stat(R["CC"][forest], 1), "canopy_share_of_area": round(float(forest.sum() / inside.sum()), 3),
        "cbd_kg_m3": stat(R["CBD"][forest & (R["CBD"] > 0)], 100), "cbh_m": stat(R["CBH"][forest & (R["CBH"] > 0)], 10),
        "ch_m": stat(R["CH"][forest & (R["CH"] > 0)], 10),
        "fdst_top": [(r["name"], r["share"]) for r in rows_fd[:5]],
    }
    print(label, json.dumps(out[label])[:900])
json.dump(out, open(HERE / "landfire_summary.json", "w"), indent=1)
