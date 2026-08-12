# Feasibility Dashboard — Excel With Rules V3

This version fixes the important export behavior:

## What the button must download

`Download Excel with Rules` must export the **CURRENT dashboard state** — i.e. the
values the user has filled/changed in the Data Entry section — not an untouched
copy of the originally loaded Excel file.

The export flow must therefore be:

1. Load the original/master Excel only as the RULE/TEMPLATE BASE.
2. Read the CURRENT dashboard `data` / `model` state.
3. Rebuild/update the three report sheets from that CURRENT state.
4. Keep all supporting sheets required for formulas, lookups, dropdowns,
   validations, named ranges, Auto GP%, Auto GP Share, etc.
5. Hide all supporting sheets, including MASTER/master when needed.
6. Export the merged/current workbook as `_with_rules.xlsx`.

## Visible sheets

- Sales forecasting tools
- INFORMATION
- AUTO GENERATED FEASIBILITY

## Hidden sheets

Any sheet needed by workbook logic, including MASTER/master/source/lookup/list/
helper/calculation/rules sheets.

## Key correction

Do **not** simply call `workbook.xlsx.load(originalBuffer)` and immediately
download that workbook. That only returns the original file.

After loading the base workbook, the exporter MUST run the existing report-sheet
builders with the live dashboard `data` and `model` before writing the XLSX.
