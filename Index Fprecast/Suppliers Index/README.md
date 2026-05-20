# ICIS Resin Index Reference Tables

Source workbook:

`ICIS Dashboard Price Forecast 2026-05-04 152801_v1.xls`

Source sheet:

`ICIS Price History`

Extracted range:

`B13:D46`

Note: the originally requested `B13:C46` range contains the ICIS Low series in column `C`. The ICIS Mid series is in column `D`, so the extractor reads `B13:D46` and keeps both Low and Mid in the same reference table.

## Outputs

Combined long table:

`icis_resin_index_reference_table.csv`

`icis_resin_index_reference_table.xlsx`

Monthly wide table:

`icis_resin_index_reference_wide.csv`

## Resin Index Types

- `ICIS Asia SE Low`
- `ICIS China Mid`

Use `resin_index_type = ICIS China Mid` for supplier forecast logic that specifically references ICIS Mid.

Current source coverage:

- `ICIS Asia SE Low`: January 2025 through September 2027
- `ICIS China Mid`: January 2025 through April 2026

## Command

```powershell
python scripts\extract_icis_resin_index_reference.py
```
