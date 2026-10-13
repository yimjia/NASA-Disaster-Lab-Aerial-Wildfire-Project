# Research data for the UQ proposal (crew safety and logistics)

Built 2026-09-30 for UQ_Crew_Safety_Logistics_Proposal_v0.1.docx. Start with
**Research_Data_Index.xlsx**. Its "Margin Timing Anchors" tab lists the real timings
that feed the safety margin M = T_fire - (T_detect + T_notify + T_egress). I checked each
one against the report text.

## entrapments/ (96 MB, 49 PDFs + 3 data files)
- `Camp_2018_NIST_burnovers.csv`: the 19 Camp Fire burnovers from NIST TN 2135 Appendix B,
  with times, durations, coordinates and page citations. Rebuild it with `extract_camp_burnovers.py`.
  The first-responder flag comes from a keyword search, and "PPD" (Paradise Police) counts
  as a keyword, so read the summary before relying on the flag.
- `entrapment_sources.json`: 32 entries with a report, timing, weather and outcome for
  each event, plus the national datasets.
- Fire reports: Camp Blue Sheet, Carr Green and Blue Sheets plus the AAR, four Cedar
  Engine 6162 reports, Eaton evacuation and AAR reports, and Pioneer's 72-hour report,
  learning review and close call write-up. The Camp Green Sheet is already in References/.
- National files (NATIONAL_*): NWCG fatality reports, Lessons Learned Center annual summaries
  for 2017 to 2025, NWCG Safety Grams, Page et al 2019 (x2), Page and Butler, Butler 2014,
  Campbell 2017 and 2019, Sullivan 2020, Gibos 2022 and JFSP 18-S-01-1.

## datasets/ (164 MB, see README.md and dataset_catalog.json inside)
- 01_progression_perimeters: time-stamped perimeters for all six fires (NIFC and GeoMAC,
  NASA FEDS 12 h, FIRIS for Eaton).
- 02_active_fire: FIRMS MODIS and VIIRS detection subsets.
- 03_weather: gridMET daily, ERA5 hourly and ASOS hourly point series.
- 05_incident_resource: ICS-209-PLUS with per-fire subsets (320 daily reports), and MTBS.
- 06_aviation_ops: DOI UAS flight-use records for FY2017 to 2026 with personal fields removed,
  plus the USFS aviation reports.
- 07_crew_egress: Sullivan 2020 travel-rate model.
- `_query_log.jsonl` has the exact request URL behind every downloaded file.

## Caveats
- The "not found" results (no firefighter entrapment on Pioneer, LNU or Eaton, and no public
  firefighter GPS tracks) come from one day of web searching on 2026-09-30. They are not a
  systematic search.
- Source documents disagree on some times: 30 minutes at Camp Rattlesnake Flats, and
  10 minutes on the Carr Blue Sheet versus the Green Sheet. Treat those gaps as timing
  uncertainty, not as errors to resolve.
- Entrapment reports name real people. Keep this folder local, and quote only the
  public report titles.
