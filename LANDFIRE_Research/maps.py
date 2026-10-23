"""Figures for the LANDFIRE and perimeter deck. Run after landfire_extract.py (uses cache/*.npy and *_grid.json).

Per fire:
  figures/<fire>_fuels_progression.png  LANDFIRE LF2016 fuel groups (FBFM40) with dated perimeters
      (NIFC GeoMAC history for Pioneer, Carr, Camp; NASA FEDS 12 h satellite perimeters for LNU; no GOFER)
  figures/<fire>_canopy_operations.png  canopy cover with completed fireline, safety zones, helispots and landing areas
      from the NIFC Operational Data Archive (Carr, Camp 2018; LNU 2020). Pioneer: fuel disturbance map instead.
figures/compare_fuel_groups.png  fuel group shares and key canopy and load numbers across the four fires.

Run: uv run --with numpy --with matplotlib --with pyproj python maps.py
"""
import csv, json
from datetime import datetime, timezone
from pathlib import Path
import numpy as np
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.colors import ListedColormap, BoundaryNorm
from matplotlib.patches import Patch
from matplotlib.lines import Line2D
from pyproj import Transformer

HERE = Path(__file__).parent
ROOT = HERE.parents[1]
CACHE, FIG = HERE / "cache", HERE / "figures"; FIG.mkdir(exist_ok=True)
PROG = ROOT / "Uncertainty Quantification" / "research_data" / "datasets" / "01_progression_perimeters"
OPS = ROOT / "Research" / "Fireline_Operations"
PER = ROOT / "sources" / "perimeters"
to5070 = Transformer.from_crs(4326, 5070, always_xy=True).transform
F40 = {int(float(r["VALUE"])): r["FBFM40"] for r in csv.DictReader(open(HERE / "lookup" / "LF16_F40_200.csv"))}
GROUPS = [("Grass", "GR", "#E3B341"), ("Grass-shrub", "GS", "#C98B2B"), ("Shrub", "SH", "#9C5B1F"),
          ("Timber-understory", "TU", "#3E8E5E"), ("Timber litter", "TL", "#1F4E3D"), ("Slash-blowdown", "SB", "#7A2E2E"),
          ("Non-burnable", "NB", "#C8CCD0")]
FIRES = {"Pioneer": ("Pioneer_2016", "Pioneer 2016", "MDT", -6), "Carr": ("Carr_2018", "Carr 2018", "PDT", -7),
         "Camp": ("Camp_2018", "Camp 2018", "PST", -8), "LNU": ("LNU_2020", "LNU Lightning Complex 2020", "PDT", -7)}
plt.rcParams.update({"font.family": "DejaVu Sans", "font.size": 10, "axes.edgecolor": "#5F6670", "axes.labelcolor": "#2B3036",
                     "xtick.color": "#5F6670", "ytick.color": "#5F6670"})


def grid(stem):
    g = json.load(open(CACHE / f"{stem}_grid.json"))
    b = g["bbox"]
    return np.load(CACHE / f"{stem}_mask.npy"), (b[0], b[2], b[1], b[3])


def rings(geom):
    polys = geom["coordinates"] if geom["type"] == "MultiPolygon" else [geom["coordinates"]]
    for poly in polys:
        for ring in poly:
            xy = np.array(ring)
            x, y = to5070(xy[:, 0], xy[:, 1])
            yield x, y


def lines(geom):
    if not geom:
        return
    parts = geom["coordinates"] if geom["type"] == "MultiLineString" else [geom["coordinates"]]
    for c in parts:
        xy = np.array(c)
        yield to5070(xy[:, 0], xy[:, 1])


def progression(key):
    """[(datetime UTC, geometry)] sorted by time."""
    if key == "LNU":
        feats = json.load(open(PROG / "LNU_2020_NASA_FEDS_perimeter_12h.geojson"))["features"]
        out = [(datetime.fromisoformat(f["properties"]["t"]).replace(tzinfo=timezone.utc), f["geometry"]) for f in feats]
    else:
        feats = json.load(open(PROG / f"{FIRES[key][0]}_NIFC_GeoMAC_perimeter_history.geojson"))["features"]
        out = [(datetime.fromtimestamp(float(f["properties"]["perimeterdatetime"]) / 1000, timezone.utc), f["geometry"]) for f in feats]
    return sorted([o for o in out if o[1]], key=lambda o: o[0])


def fuel_groups(stem):
    f = np.load(CACHE / f"{stem}_FBFM40.npy")
    lut = np.zeros(f.max() + 2, dtype=np.int8) - 1
    for v, code in F40.items():
        if v >= 0:
            for i, (_, pre, _) in enumerate(GROUPS):
                if code.startswith(pre) and v < len(lut):
                    lut[v] = i
    g = np.where(f >= 0, lut[np.clip(f, 0, len(lut) - 1)], -1).astype(float)
    g[g < 0] = np.nan
    return g


def frame(ax, ext):
    ax.set_xlim(ext[0], ext[1]); ax.set_ylim(ext[2], ext[3]); ax.set_aspect("equal")
    ax.set_xticks([]); ax.set_yticks([])
    span = ext[1] - ext[0]; km = 10 if span > 40000 else 5
    x0, y0 = ext[0] + span * 0.04, ext[2] + (ext[3] - ext[2]) * 0.04
    ax.plot([x0, x0 + km * 1000], [y0, y0], color="#2B3036", lw=3)
    ax.text(x0 + km * 500, y0 + (ext[3] - ext[2]) * 0.015, f"{km} km", ha="center", fontsize=9, color="#2B3036")


def final_outline(ax, stem, **kw):
    for f in json.load(open(PER / f"{stem}.geojson"))["features"]:
        for x, y in rings(f["geometry"]):
            ax.plot(x, y, **kw)


# ---------------------------------------------------------------- fuels + progression
for key, (stem, title, tz, off) in FIRES.items():
    mask, ext = grid(stem)
    g = fuel_groups(stem)
    shade = np.where(mask, g, np.nan); outside = np.where(~mask, g, np.nan)
    cmap = ListedColormap([c for _, _, c in GROUPS]); norm = BoundaryNorm(np.arange(-0.5, len(GROUPS)), len(GROUPS))
    fig, ax = plt.subplots(figsize=(9, 7.2), dpi=170)
    ax.imshow(outside, extent=ext, cmap=cmap, norm=norm, alpha=0.28, interpolation="nearest")
    ax.imshow(shade, extent=ext, cmap=cmap, norm=norm, interpolation="nearest")
    prog = progression(key)
    t0, t1 = prog[0][0].timestamp(), prog[-1][0].timestamp()
    pc = plt.get_cmap("magma_r")
    for t, geom in prog:
        c = pc(0.15 + 0.8 * (t.timestamp() - t0) / max(1, t1 - t0))
        for x, y in rings(geom):
            ax.plot(x, y, color=c, lw=0.7, alpha=0.9)
    final_outline(ax, stem, color="#111111", lw=1.4)
    frame(ax, ext)
    loc = lambda t: (t.timestamp() + off * 3600)
    d0 = datetime.fromtimestamp(loc(prog[0][0]), timezone.utc).strftime("%d %b"); d1 = datetime.fromtimestamp(loc(prog[-1][0]), timezone.utc).strftime("%d %b %Y")
    sm = plt.cm.ScalarMappable(cmap=matplotlib.colors.LinearSegmentedColormap.from_list("p", [pc(0.15), pc(0.55), pc(0.95)]))
    cb = fig.colorbar(sm, ax=ax, fraction=0.035, pad=0.01, ticks=[0, 1]); cb.ax.set_yticklabels([d0, d1]); cb.outline.set_visible(False)
    cb.set_label("Perimeter date (local)", color="#2B3036")
    src = "NASA FEDS 12 h satellite perimeters" if key == "LNU" else "NIFC GeoMAC perimeter history"
    ax.legend(handles=[Patch(color=c, label=n) for n, _, c in GROUPS if n != "Slash-blowdown" or key == "LNU"] +
              [Line2D([0], [0], color="#111111", lw=1.4, label="Final perimeter")], loc="upper right", fontsize=8, frameon=True, framealpha=0.9)
    note = " [post-fire classification, see disturbance map]" if key == "Pioneer" else ""
    ax.set_title(f"{title}: LANDFIRE 2016 fuel groups{note} and {len(prog)} dated perimeters ({src})", fontsize=11, color="#2B3036", loc="left")
    fig.tight_layout(); fig.savefig(FIG / f"{key}_fuels_progression.png", bbox_inches="tight"); plt.close(fig)
    print("wrote", key, "fuels")

# ---------------------------------------------------------------- canopy + operations (or disturbance for Pioneer)
OPSFILES = {"Carr": ("nifs2018_Carr_lines.geojson", "nifs2018_Carr_points.geojson", "2018-08-20T20:00"),
            "Camp": ("nifs2018_Camp_lines.geojson", "nifs2018_Camp_points.geojson", "2018-11-23T20:00"),
            "LNU": ("nifs2020_LNU_lines.geojson", "nifs2020_LNU_points.geojson", "2020-09-10T20:00")}
LINESTY = {"Completed Dozer Line": ("#C0392B", 0.9), "Completed Hand Line": ("#1F6FB2", 0.9), "Completed Line": ("#6C3483", 0.8),
           "Road as Completed Line": ("#5F6670", 0.6)}
PTSTY = {"Safety Zone": ("#27AE60", "o"), "Helispot": ("#F39C12", "^"), "Unimproved Landing Area": ("#F39C12", "s"), "Helibase": ("#D35400", "*")}
for key, (stem, title, tz, off) in FIRES.items():
    mask, ext = grid(stem)
    fig, ax = plt.subplots(figsize=(9, 7.2), dpi=170)
    if key == "Pioneer":
        fd = np.load(CACHE / f"{stem}_FDist.npy").astype(float)
        sev = np.full(fd.shape, np.nan)
        lut = {r["VALUE"]: r for r in csv.DictReader(open(HERE / "lookup" / "LF16_FDst_200.csv"))}
        for v, r in lut.items():
            v = int(float(v))
            if r["D_TYPE"] == "Fire":
                sev[fd == v] = {"Low": 1, "Moderate": 2, "High": 3}.get(r["D_SEVERITY"], np.nan)
            elif r["D_TYPE"] == "No Disturbance":
                sev[fd == v] = 0
        cm = ListedColormap(["#DDE3E8", "#F2C36B", "#E07B39", "#9E2A1F"])
        ax.imshow(np.where(mask, sev, np.nan), extent=ext, cmap=cm, norm=BoundaryNorm([-0.5, 0.5, 1.5, 2.5, 3.5], 4), interpolation="nearest")
        ax.imshow(np.where(~mask, sev, np.nan), extent=ext, cmap=cm, norm=BoundaryNorm([-0.5, 0.5, 1.5, 2.5, 3.5], 4), alpha=0.3, interpolation="nearest")
        final_outline(ax, stem, color="#111111", lw=1.4)
        ax.legend(handles=[Patch(color=c, label=l) for c, l in zip(["#DDE3E8", "#F2C36B", "#E07B39", "#9E2A1F"],
                  ["No disturbance", "Fire, low severity", "Fire, moderate", "Fire, high"])], loc="upper right", fontsize=8, framealpha=0.9)
        ax.set_title("Pioneer 2016: LANDFIRE 2016 already records the fire (fuel disturbance layer)", fontsize=11, color="#2B3036", loc="left")
    else:
        cc = np.load(CACHE / f"{stem}_CC.npy").astype(float); cc[cc < 0] = np.nan
        im = ax.imshow(np.where(mask, cc, np.nan), extent=ext, cmap="Greens", vmin=0, vmax=90, interpolation="nearest")
        ax.imshow(np.where(~mask, cc, np.nan), extent=ext, cmap="Greys", vmin=0, vmax=150, alpha=0.35, interpolation="nearest")
        cb = fig.colorbar(im, ax=ax, fraction=0.035, pad=0.01); cb.set_label("Canopy cover %", color="#2B3036"); cb.outline.set_visible(False)
        final_outline(ax, stem, color="#111111", lw=1.2)
        lf, pf, snap = OPSFILES[key]
        t = datetime.fromisoformat(snap).replace(tzinfo=timezone.utc).timestamp() * 1000
        for f in json.load(open(OPS / lf))["features"]:
            p = f["properties"]; cat = p.get("FeatureCategory")
            if cat in LINESTY and (p.get("GDB_FROM_DATE") or 0) <= t < (p.get("GDB_TO_DATE") or 4e12):
                col, lw = LINESTY[cat]
                for x, y in lines(f["geometry"]):
                    ax.plot(x, y, color=col, lw=lw)
        seen = set()
        for f in json.load(open(OPS / pf))["features"]:
            p = f["properties"]; cat = p.get("FeatureCategory")
            g_ = f.get("geometry")
            if cat in PTSTY and g_ and p.get("GlobalID") not in seen:
                seen.add(p.get("GlobalID"))
                x, y = to5070(*g_["coordinates"][:2]); col, mk = PTSTY[cat]
                ax.plot(x, y, marker=mk, color=col, ms=5, mec="#111111", mew=0.5, ls="none")
        ax.legend(handles=[Line2D([0], [0], color=c, lw=2, label=k) for k, (c, _) in LINESTY.items()] +
                  [Line2D([0], [0], marker=m, color=c, mec="#111111", ls="none", label=k) for k, (c, m) in PTSTY.items()],
                  loc="upper right", fontsize=8, framealpha=0.9)
        day = datetime.fromisoformat(snap).strftime("%d %b %Y")
        ax.set_title(f"{title}: canopy cover with completed fireline on the incident map ({day}) and operations points", fontsize=11, color="#2B3036", loc="left")
    frame(ax, ext)
    fig.tight_layout(); fig.savefig(FIG / f"{key}_canopy_operations.png", bbox_inches="tight"); plt.close(fig)
    print("wrote", key, "canopy/ops")

# ---------------------------------------------------------------- comparison
S = json.load(open(HERE / "landfire_summary.json"))
names = list(S.keys())
fig, axes = plt.subplots(1, 2, figsize=(12, 4.6), dpi=170, gridspec_kw={"width_ratios": [1.5, 1]})
left = np.zeros(len(names))
for n, _, c in GROUPS:
    vals = np.array([S[k]["fuel_group_share"].get(n, 0) for k in names]) * 100
    axes[0].barh(names, vals, left=left, color=c, label=n, edgecolor="white", linewidth=0.6)
    for i, v in enumerate(vals):
        if v >= 8:
            axes[0].text(left[i] + v / 2, i, f"{v:.0f}", ha="center", va="center", fontsize=8, color="white" if n in ("Timber litter", "Shrub", "Timber-understory") else "#2B3036")
    left += vals
axes[0].invert_yaxis(); axes[0].set_xlim(0, 100); axes[0].set_xlabel("Share of burned area, %")
axes[0].legend(ncol=4, fontsize=8, loc="upper center", bbox_to_anchor=(0.5, -0.16), frameon=False)
axes[0].set_title("LANDFIRE 2016 fuel groups inside each perimeter", loc="left", fontsize=11)
for s_ in ("top", "right"): axes[0].spines[s_].set_visible(False)
load = [S[k]["surface_load_tac_burnable"]["median"] for k in names]
axes[1].barh(names, load, color="#8C6D46")
for i, v in enumerate(load):
    axes[1].text(v + 0.2, i, f"{v:.1f} t/ac", va="center", fontsize=9)
axes[1].invert_yaxis(); axes[1].set_xlim(0, max(load) * 1.3); axes[1].set_yticklabels([])
axes[1].set_xlabel("Median assigned surface fuel load, t/ac (Scott and Burgan 2005)")
axes[1].set_title("Surface fuel load by fuel model", loc="left", fontsize=11)
for s_ in ("top", "right"): axes[1].spines[s_].set_visible(False)
fig.tight_layout(); fig.savefig(FIG / "compare_fuel_groups.png", bbox_inches="tight"); plt.close(fig)
print("wrote comparison")
