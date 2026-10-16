# Portability checklist (Excel and Google Sheets)

The workbook was recalculated in Python (`pycel`), never in Excel or Google Sheets. This checklist is how to close that gap. Tick each box and note any difference. Full expected values for every metric and scenario are in `tests/expected_values.csv`.

## Excel

- [ ] Open `model/keel_saas_model.xlsx` and let it calculate (it is set to recalculate on open). No `#NAME?`, `#REF!`, `#VALUE!` or `#DIV/0!` anywhere (Ctrl+F for `#`).
- [ ] `Cover!B4` reads **ALL CHECKS PASS** and `Cover!B7` reads 0 outlier warnings.
- [ ] Formulas > Name Manager shows 355 names (101 `inp_*`, 249 `out_*`, plus `chk_*` and `scen_*`).
- [ ] Change the dropdown on `Inputs_Drivers!C4` to each scenario and compare with the table below (Downside: master check reads **CHECK FAILURE** by design, because pro-forma cash goes negative).
- [ ] Type 0.9 into the tax-rate input: Excel should refuse it (data validation, allowed range 0 to 0.5).
- [ ] `MA_Valuation`: the football field chart renders, with floating bars.
- [ ] Columns C:M are grouped and collapsed on the time-series tabs; the outline buttons expand them. `Rev_Rec` shows them expanded.
- [ ] Switch to **Custom**, change a value in `Inputs_Scenarios` column F, and confirm the model responds.
- [ ] Save as .xlsx and re-open: nothing changes.

## Google Sheets

- [ ] File > Import > Upload > "Replace spreadsheet". Wait for calculation. No error values.
- [ ] Same Cover, scenario table, and data-validation checks as above.
- [ ] Data > Named ranges lists the names; `Outputs_Web` links resolve.
- [ ] The football field chart renders (Sheets may restyle it).
- [ ] Note any function Sheets treats differently. The workbook uses only 21 classic functions (ABS, AND, AVERAGE, COLUMN, COUNT, COUNTIF, EOMONTH, IF, INDEX, ISNUMBER, LARGE, MATCH, MAX, MIN, OR, ROUNDUP, ROW, SUM, SUMIF, SUMPRODUCT, YEAR).

## Expected values (from the independent recalculation; GBP k unless a ratio)

| Metric | Base | Upside | Downside |
|---|---|---|---|
| FY1 revenue | 48,367 | 51,168 | 43,789 |
| FY1 EBITDA | -6,972 | -5,720 | -8,479 |
| ARR, month 24 | 68,321 | 80,599 | 50,209 |
| NRR, month 12 | 1.0277 | 1.0677 | 0.9384 |
| Cash, month 24 | 22,579 | 31,968 | 9,508 |
| Lowest month-end cash | 20,882 | 24,328 | 9,508 |
| Deal value created | 7,333 | 16,827 | -12,592 |
| Pro-forma cash, month 24 | 8,913 | 19,125 | -6,226 |

## Result

| Engine | Date | Matches? | Differences found |
|---|---|---|---|
| Excel | | | |
| Google Sheets | | | |
| HyperFormula (Phase 6) | | | |
