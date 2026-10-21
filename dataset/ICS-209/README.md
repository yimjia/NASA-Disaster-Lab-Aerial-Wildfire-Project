# ICS-209 dataset: NASA Disaster Lab benchmark fires

Core incident record for the four benchmark fires (Pioneer 2016, Carr 2018, Camp 2018, LNU Lightning Complex 2020) used in the UQ research.

## Contents

| Path | What it is |
|---|---|
| `per_fire_workbooks/` | One workbook per fire built from ICS-209-PLUS and NICC daily situation reports: every ICS-209 field by report period (sheet `All 209 fields`), plus a summary sheet. Includes fuel model fields (FUEL_MODEL, SECNDRY_FUEL_MODEL, ADDNTL_FUEL_MODEL), size (CURR_INCIDENT_AREA), containment, resources and fire behavior. |
| `ICS-209-PLUS_subsets/ICS209PLUS_benchmark_incidents.csv` | Incident-level rows for the benchmark incidents |
| `ICS-209-PLUS_subsets/ICS209PLUS_benchmark_sitreps.csv` | All situation reports for those incidents (multi-line text fields: read with a real CSV parser) |
| `ICS-209-PLUS_subsets/ICS209PLUS_benchmark_complex_associations.csv` | Complex to component incident links (LNU) |
| `ICS-209-PLUS_subsets/ics209plus-reference.zip` | ICS-209-PLUS reference tables (codes and field definitions) |
| `Camp_2018_Perimeter_Time_Area_ICS209PLUS.xlsx` | Camp 2018 time-area table: ICS-209-PLUS sizes alongside NIFC GeoMAC, NASA FEDS and NIST/CAL FIRE timeline events |

Incident IDs used for the subsets: 2016_4416704_PIONEER, 2018_9220077_CAMP, 2018_9236519_CARR, 2020_11850821_LNU LIGHTNING COMPLEX (plus 2003_CA-CNF-003056_CEDAR in the CSV subsets).

## Source

ICS-209-PLUS wildfire dataset on figshare: wildfire tables https://ndownloader.figshare.com/files/38766504 (48.7 MB zip, not stored here; download from the link) and reference tables https://ndownloader.figshare.com/files/38766492. Accessed 30 September 2026. Check the figshare record for the dataset citation and license before reuse.

## Reading notes

- Use `CURR_INCIDENT_AREA` for what was on record in real time; `ACRES` is cleaned and capped at the final size.
- The ICS-209 size usually equals the previous night's perimeter, so it runs about a day behind the fire on growth days.
- The NICC daily situation report republishes the previous evening's ICS-209; it is not an independent observation.
