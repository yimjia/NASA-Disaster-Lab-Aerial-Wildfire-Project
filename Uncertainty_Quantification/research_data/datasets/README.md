# UQ research datasets (ACERO benchmark fires)

Accessed 2026-09-30. Everything here is public and was fetched with no login and no key.
Full metadata for every entry, including the ones that were only cataloged, is in
`dataset_catalog.json` (49 entries, 26 downloaded). The exact request URL for every
downloaded file is logged one per line in `_query_log.jsonl`.

Benchmark fires: Camp 2018, Carr 2018, LNU Lightning Complex 2020, Pioneer 2016,
Eaton 2025, Cedar 2003. The final perimeters, CAL FIRE DINS and the NIST TN2135 timeline
already live in `../../../sources/` and were not downloaded again.

Size on disk is about 164 MB. About 120 MB more was downloaded as national FIRMS yearly
files, filtered locally, then deleted.

## Folders

| Folder | Contents |
|---|---|
| `01_progression_perimeters/` | Time-stamped perimeters: NIFC GeoMAC histories (Camp, Carr, Pioneer, Cedar), NIFC Operational Data Archive versions and FIRIS IR perimeters (Eaton), NASA FEDS 12-hourly VIIRS perimeters and fire lines (Camp, Carr, LNU), Chen et al. 2022 FEDS Readme and Summary |
| `02_active_fire/` | FIRMS MODIS and VIIRS S-NPP archive detections, subset per fire |
| `03_weather/` | gridMET daily point series, ERA5 hourly point series (Open-Meteo), ASOS hourly station observations (IEM), `weather_points.json` |
| `04_fuels_terrain/` | Empty on purpose. LANDFIRE and 3DEP access is recorded in the catalog only |
| `05_incident_resource/` | ICS-209-PLUS subsets and original zip, MTBS fire points subset and original zip |
| `06_aviation_ops/` | DOI UAS flight-use records (personal fields removed), USFS Aviation Annual Reports CY2018 and CY2020 |
| `07_crew_egress/` | Sullivan et al. 2020 travel-rate paper (CC BY 4.0) |

## How each file was obtained

### Perimeters (ArcGIS REST, `f=geojson`, `outSR=4326`, `geometryPrecision=6`, paged by objectIds)

Base: `https://services3.arcgis.com/T4QMspbfLg3qTGWY/arcgis/rest/services`

- Camp: `Historic_Geomac_Perimeters_2018/FeatureServer/0/query?where=uniquefireidentifier='2018-CABTU-016737'` (28 features)
- Carr: same layer, `uniquefireidentifier='2018-CAWNP-001084'` (63)
- Pioneer: `Historic_Geomac_Perimeters_2016/FeatureServer/0/query?where=uniquefireidentifier='2016-IDBOF-000539'` (89)
- Cedar: `Historic_Geomac_Perimeters_2003/FeatureServer/0/query?where=uniquefireidentifier='2003-CACNF-003056'` (8)
- Eaton: `Operational_Data_Archive_2025/FeatureServer/6/query?where=IRWINID in ('{72660ADC-B5EF-4D96-A33F-B4EA3740A4E3}','{19A869B7-6CB6-4EF7-AD31-94F13BD1C622}') and FeatureCategory='Wildfire Daily Fire Perimeter'` (44). GDB_FROM_DATE is an edit time, not an observation time.
- Eaton FIRIS: `https://services5.arcgis.com/VNhSlpl1umSknM3q/arcgis/rest/services/FIRIS_perimeters_(Production_ReadOnly)/FeatureServer/0/query?where=1=1&geometry=-118.162,34.162,-118.013,34.238&geometryType=esriGeometryEnvelope&inSR=4326&spatialRel=esriSpatialRelIntersects&outFields=status,description,date_added,source` (20). This is a third-party (Watch Duty) public view, not an official archive.

### NASA FEDS (OGC API Features, `limit=40` pages; `limit=100` returns HTTP 500)

`https://openveda.cloud/api/features/collections/public.eis_fire_lf_perimeter_archive/items?fireid=<id>&datetime=<window>`,
and the same query against `public.eis_fire_lf_fireline_archive`.

- Camp: fireid F17028, `2018-11-07T00:00:00Z/2018-12-01T00:00:00Z`
- Carr: F9276, `2018-07-20T00:00:00Z/2018-09-05T00:00:00Z`
- LNU: F11529, F11572, F11709, `2020-08-15T00:00:00Z/2020-09-25T00:00:00Z`

Fire IDs repeat across years, so always filter by datetime. The archive covers 2018 to 2021 only.

Chen et al. 2022 (figshare collection 5601537): `https://ndownloader.figshare.com/files/34382966` (Readme) and `https://ndownloader.figshare.com/files/34062410` (Summary.tar.gz).

### FIRMS (no key)

`https://firms.modaps.eosdis.nasa.gov/data/country/<modis|viirs-snpp>/<year>/<sensor>_<year>_United_States.csv`
was downloaded for MODIS 2003, 2016, 2018, 2020 and VIIRS S-NPP 2016, 2018. Each file was filtered to
the fire bbox plus 0.15 deg and to the fire date window, then deleted. The windows were:

- Camp: 2018-11-07 to 2018-11-26
- Carr: 2018-07-22 to 2018-09-02
- LNU: 2020-08-16 to 2020-10-03
- Pioneer: 2016-07-17 to 2016-11-15
- Cedar: 2003-10-24 to 2003-11-06

VIIRS 2020 (53.8 MB) was skipped because it is over the 50 MB cap. There are no 2025 yearly files, so Eaton needs a MAP_KEY.

### Weather

Points are the bbox centers listed in `03_weather/weather_points.json`.

- gridMET: `http://thredds.northwestknowledge.net:8080/thredds/ncss/agg_met_<var>_1979_CurrentYear_CONUS.nc?var=<name>&latitude=<lat>&longitude=<lon>&time_start=<t0>T00:00:00Z&time_end=<t1>T00:00:00Z&accept=csv`
  for vs, th, rmin, rmax, tmmx, tmmn, fm100, fm1000, erc, bi, pr, vpd. NCSS returns packed integers. The scale_factor and add_offset from the OPeNDAP `.das` were applied, and temperatures are in K.
- ERA5 hourly: `https://archive-api.open-meteo.com/v1/archive?latitude=<lat>&longitude=<lon>&start_date=<t0>&end_date=<t1>&hourly=temperature_2m,relative_humidity_2m,wind_speed_10m,wind_direction_10m,wind_gusts_10m,surface_pressure,precipitation,vapour_pressure_deficit&wind_speed_unit=ms&timezone=UTC&models=era5`
- ASOS: `https://mesonet.agron.iastate.edu/cgi-bin/request/asos.py?station=<ID>&data=tmpf&data=relh&data=sknt&data=drct&data=gust&data=peak_wind_gust&data=peak_wind_drct&data=p01i&data=vsby&year1=..&month1=..&day1=..&year2=..&month2=..&day2=..&tz=Etc/UTC&format=onlycomma&latlon=yes&elev=yes&missing=empty&trace=T&report_type=3&report_type=4`.
  Stations: OVE and CIC (Camp), RDD and RBL (Carr), STS, APC and VCB (LNU), BOI and SNT (Pioneer), EMT and BUR (Eaton), RNM and SEE (Cedar).

### Incident and resource data

- ICS-209-PLUS: `https://ndownloader.figshare.com/files/38766504` (wildfire zip, kept) and `https://ndownloader.figshare.com/files/38766492` (reference). The subsets were filtered by these INCIDENT_IDs:
  - 2003_CA-CNF-003056_CEDAR
  - 2016_4416704_PIONEER
  - 2018_9220077_CAMP
  - 2018_9236519_CARR
  - 2020_11850821_LNU LIGHTNING COMPLEX

  The sitreps contain multi-line text fields, so read them with a real CSV parser.
- MTBS points: `https://edcintl.cr.usgs.gov/downloads/sciweb1/shared/MTBS_Fire/data/composite_data/fod_pt_shapefile/mtbs_fod_pts_data.zip`. The DBF was filtered by name, year and state (7 rows).

### Aviation

- DOI UAS flight use: `https://services3.arcgis.com/T4QMspbfLg3qTGWY/arcgis/rest/services/Historic_Flight_Use_20251204/FeatureServer/0/query?where=1=1&outFields=<30 non-personal fields>&returnGeometry=false&f=json`, paged 2,000 at a time (23,142 records). Operator and observer names, FAA registration, battery ID and free-text notes were not requested.
- USFS aviation reports: URLs are in the catalog.

### Crew egress

- Sullivan et al. 2020: `https://research.fs.usda.gov/download/treesearch/61249.pdf`

## Gaps (details in the catalog notes)

- **LNU 2020:** no public agency daily perimeter series was found. NASA FEDS is the only time series.
- **Eaton 2025:** no public perimeter captures the first night's run. FIRMS 2025 needs a key, and there is no ICS-209-PLUS coverage.
- **Cedar 2003:** only 8 GeoMAC perimeters, starting about 2 days after ignition. There is no FEDS, VIIRS, GOES-R or HRRR data, and the fuels are LANDFIRE 2001 Refresh only.
- **Pioneer 2016:** not in the FEDS archive. LANDFIRE 2014 fuels have no public direct URL.
- **Blocked by access:**
  - NIROPS raw IR needs FAMAuth.
  - WRCC RAWS needs an access code.
  - Synoptic needs a token.
- **Aviation:** no public historical FAA TFR archive was found. Crew GPS travel tracks are not public, so only the published model coefficients are available.
