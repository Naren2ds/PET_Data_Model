# PET_Data_Model

## Pipeline behavior

Each extraction script is built to run as a small pipeline:

- By default, it picks the latest matching workbook from `data`.
- Use `--file` to force a specific workbook.
- Outputs are written to `output`.
- Use `--write-csv` when CSV files are also needed.

## Bavaria monthly extraction

Place the monthly Bavaria price workbook in `data`, then run:

```powershell
python scripts\extract_bavaria_formula.py
```

The pipeline will:

- Pick the latest workbook in `data` matching `*Bavaria*.xlsx`.
- Match the workbook month to the `Formula` sheet, for example `Precio Marzo` to `Formula Marzo 26`.
- Extract `B13:CV26`.
- Save the final Excel output in `output` as `<source workbook>_formula_final.xlsx`.

To run a specific file:

```powershell
python scripts\extract_bavaria_formula.py --file "data\1 Bavaria - Precio Marzo.xlsx"
```

To also generate CSV extracts:

```powershell
python scripts\extract_bavaria_formula.py --write-csv
```

## Ecuador Formula Virgen monthly extraction

Place the monthly Ecuador workbook in `data`, then run:

```powershell
python scripts\extract_ecuador_formula_virgen.py
```

The pipeline will:

- Pick the latest workbook in `data` matching `*ECUADOR*.xlsx`.
- Use the `Formula Virgen` sheet.
- Extract `A2:AT18`.
- Add `time_period`, `time_period_year`, and `time_period_month` in `final_data`.
- Use row 3 as the month source for `time_period`, for example `AT3 = Marzo` maps to `March 2026`.
- Save the final Excel output in `output` as `<source workbook>_formula_virgen_final.xlsx`.

To run a specific file:

```powershell
python scripts\extract_ecuador_formula_virgen.py --file "data\3 Precios AB Inbev_Abril 26 ECUADOR.xlsx"
```

To override the rightmost period anchor:

```powershell
python scripts\extract_ecuador_formula_virgen.py --anchor-period 2026-04
```

To also generate CSV extracts:

```powershell
python scripts\extract_ecuador_formula_virgen.py --write-csv
```

## Cristalpet indices extraction

Place the monthly Cristalpet workbook in `data`, then run:

```powershell
python scripts\extract_cristalpet_indices.py
```

The pipeline will:

- Pick the latest workbook in `data` matching `*Cristalpet*.xlsx`.
- Use the `Indices 2025-26-27` sheet.
- Extract `B2:N8` and `B19:N19`.
- Use row 2 as the period header for `B2:N8`.
- Use row 17 as the period header for the freight values in `B19:N19`.
- Add Total Landing Cost rows from `Precos Cristalpet` row 5, using row 4 as the period header.
- Save the final Excel output in `output` as `<source workbook>_indices_final.xlsx`.

To run a specific file:

```powershell
python scripts\extract_cristalpet_indices.py --file "data\04. Cristalpet.xlsx"
```

To also generate CSV extracts:

```powershell
python scripts\extract_cristalpet_indices.py --write-csv
```

## Engepack indices extraction

Place the monthly Engepack workbook in `data`, then run:

```powershell
python scripts\extract_engepack_indices.py
```

The pipeline will:

- Pick the latest workbook in `data` matching `*Engepack*.xlsx`.
- Use the `Indices` sheet.
- Extract `B2:AT5`.
- Extract `B23:AT23` as field name `Freight`, using row 20 as the freight period header.
- Add Total Landing Cost rows from `Precos Engepack!B8:AP8`.
- Map TLC periods from `Indices!G2:AT2`, because the TLC formulas reference those index columns.
- Save the final Excel output in `output` as `<source workbook>_indices_final.xlsx`.

To run a specific file:

```powershell
python scripts\extract_engepack_indices.py --file "data\04. Engepack.xlsx"
```

To also generate CSV extracts:

```powershell
python scripts\extract_engepack_indices.py --write-csv
```

## Valgroup month extraction

Place the monthly Valgroup workbook in `data`, then run:

```powershell
python scripts\extract_valgroup_months.py
```

The pipeline will:

- Pick the latest workbook in `data` matching `*Valgroup*.xlsx`.
- Process all month-named sheets, such as `JAN.26`, `FEV.26`, `MAR.26`, and `ABR.26`.
- Extract `U2:V12` from each month sheet.
- Add `month_english`, `time_period`, `time_period_year`, and `time_period_month`.
- Translate Portuguese labels into English.
- For sheets where column `U` contains labels, use column `U` as the translation source.
- For `JAN.26`, where `U:V` are product value columns, use column `T` as the label source while still preserving the raw `U2:V12` extract.
- Save the final Excel output in `output` as `<source workbook>_months_final.xlsx`.

To run a specific file:

```powershell
python scripts\extract_valgroup_months.py --file "data\04. Valgroup.xlsx"
```

To also generate CSV extracts:

```powershell
python scripts\extract_valgroup_months.py --write-csv
```

## Sell Side ICIS Price History extraction

Place the monthly Sell Side ABI workbook in `data`, then run:

```powershell
python scripts\extract_sell_side_icis_price_history.py
```

The pipeline will:

- Pick the latest workbook in `data` matching `Sell_Side*ABI*.xlsx`.
- Use the `ICIS Price History` sheet.
- Extract `I214:R217`.
- Use `I13:R13` as the final field names.
- Preserve the raw range in `raw_I214_R217`.
- Add `price_date`, `time_period`, `time_period_year`, and `time_period_month` from the same source rows.
- Save the final Excel output in `output` as `<source workbook>_icis_price_history_final.xlsx`.

To run a specific file:

```powershell
python scripts\extract_sell_side_icis_price_history.py --file "data\Sell_Side_2026-ABI Mar 2026.xlsx"
```

To also generate CSV extracts:

```powershell
python scripts\extract_sell_side_icis_price_history.py --write-csv
```

## Pricing Pref Inputs extraction

Place the monthly Pricing Pref workbook in `data`, then run:

```powershell
python scripts\extract_pricing_pref_inputs.py
```

The pipeline will:

- Pick the latest workbook in `data` matching `Pricing_Pref*ABI*Cliente*.xlsx`.
- Use the `Inputs` sheet.
- Extract `G9:H13`.
- Use column `G` as the metric label source and translate those labels into English.
- Add `time_period`, `time_period_year`, and `time_period_month`.
- Default the period from `Inputs!C6:D6`, which is `April 2026` for the current file.
- Save the final Excel output in `output` as `<source workbook>_inputs_final.xlsx`.

To run a specific file:

```powershell
python scripts\extract_pricing_pref_inputs.py --file "data\Pricing_Pref-Bot_ABI_PER_Men_4_26_Cliente.xlsx"
```

To override the period:

```powershell
python scripts\extract_pricing_pref_inputs.py --time-period 2026-04
```

To also generate CSV extracts:

```powershell
python scripts\extract_pricing_pref_inputs.py --write-csv
```

## Pricing Pref RD Inputs extraction

Place the monthly Pricing Pref RD workbook in `data`, then run:

```powershell
python scripts\extract_pricing_pref_rd_inputs.py
```

The pipeline will:

- Pick the latest workbook in `data` matching `Pricing_Pref_ABI_RD*Cliente*.xlsx`.
- Use the `Inputs` sheet.
- Extract `G8:I12`.
- Use row 8 as the country header row.
- Use column `G` as the metric label source and translate those labels into English.
- Transform the country columns into long format with `country` as the column name.
- Add `time_period`, `time_period_year`, and `time_period_month`.
- Default the period from `Inputs!C6:D6`, which is `April 2026` for the current file.
- Save the final Excel output in `output` as `<source workbook>_inputs_final.xlsx`.

To run a specific file:

```powershell
python scripts\extract_pricing_pref_rd_inputs.py --file "data\Pricing_Pref_ABI_RD_Men_4_26_Cliente.xlsx"
```

To override the period:

```powershell
python scripts\extract_pricing_pref_rd_inputs.py --time-period 2026-04
```

To also generate CSV extracts:

```powershell
python scripts\extract_pricing_pref_rd_inputs.py --write-csv
```

## Pricing Pref PAN Inputs extraction

Place the monthly Pricing Pref PAN workbook in `data`, then run:

```powershell
python scripts\extract_pricing_pref_pan_inputs.py
```

The pipeline will:

- Pick the latest workbook in `data` matching `Pricing_Pref_ABI_PAN*Cliente*.xlsx`.
- Use the `Inputs` sheet.
- Extract `G7:H11`.
- Use column `G` as the metric label source and translate those labels into English.
- Add `time_period`, `time_period_year`, and `time_period_month`.
- Default the period from `Inputs!C6:D6`, which is `April 2026` for the current file.
- Save the final Excel output in `output` as `<source workbook>_inputs_final.xlsx`.

To run a specific file:

```powershell
python scripts\extract_pricing_pref_pan_inputs.py --file "data\Pricing_Pref_ABI_PAN_Men_04_26_Cliente.xlsx"
```

To override the period:

```powershell
python scripts\extract_pricing_pref_pan_inputs.py --time-period 2026-04
```

To also generate CSV extracts:

```powershell
python scripts\extract_pricing_pref_pan_inputs.py --write-csv
```
