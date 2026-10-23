# LANDFIRE evaluation of the benchmark fires (2026-10-22)

- `LANDFIRE_Evaluation_Benchmark_Fires_2026-10-22.docx`: the documentation (what was evaluated, method, results per fire, data quality findings, next steps).
- `Perimeters_and_LANDFIRE_2026-10-22.pptx`: 14-slide deck, perimeter progression and fireline over LANDFIRE fuels and canopy, with what each picture shows.
- `landfire_extract.py` -> `landfire_summary.json`, `landfire_<fire>_<FBFM40|FCCS|EVT|FDist>.csv`: LF 2016 Remap layers inside each final perimeter (30 m, lfps.usgs.gov image services, tiled requests; URLs in `landfire_query_log.txt`).
- `maps.py` -> `figures/`: fuel groups with dated perimeters; canopy cover with fireline and operations points (disturbance map for Pioneer); cross-fire comparison chart.
- `build_landfire_outputs.py`: builds the Word document and the deck from the same numbers.
- `lookup/`: LANDFIRE LF2016 code tables, LANDFIRE Data Dictionary, Scott and Burgan (2005) RMRS-GTR-153 and its Table 7 as CSV.
- `cache/`: downloaded rasters as .npy (about 450 MB). Safe to delete; rerunning landfire_extract.py downloads them again.

Key cautions: Pioneer's LF 2016 map is post-fire (99% fire-disturbed); loads are Scott and Burgan model loads, not measurements; FCCS value 0 ("Bare Ground") is mostly grass here; Camp has a map-zone seam.

## In the Disaster Lab repo

Published as `LANDFIRE_Research/` without `cache/` (rebuilt by landfire_extract.py). The scripts read the final perimeters from `sources/perimeters/`, dated perimeters from `Uncertainty Quantification/research_data/datasets/01_progression_perimeters/`, fireline from `Research/Fireline_Operations/` and terrain from `Research/Fuels_and_Topography/` in the project folder; those inputs are not in this folder, so the figures and numbers here are the record and rerunning needs the full project.
