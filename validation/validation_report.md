# Validation report

*Generated 2026-10-15 22:21 by `validation/validate_model.py` on `keel_saas_model.xlsx`. Recalculation engine: pycel (not Excel).*

**32 passed, 0 failed, 2 for review, 1 expected.**


## A. Static scans

| Status | Test | Detail |
|---|---|---|
| PASS | No hardcoded numbers in forecast ranges (months 1-24, 11 time-series tabs) | none found |
| PASS | One consistent formula across all 24 periods (309 rows tested) | every row consistent |
| PASS | No non-portable or volatile functions (16,658 formulas scanned) | none found |
| PASS | Every inp_* / chk_* name used in a formula is defined | 355 defined names |
| PASS | Every defined name points at an existing sheet | ok |
| PASS | Every typed number outside the input tabs is marked: blue (input), grey (engine-run snapshot) or white (header) | all marked |
| REVIEW | Numeric literals inside formulas other than unit conversions (0, 1, 2, 3, 6, 11, 12, 30, 100, 1000, 0.5) | 10 x19 (e.g. Lever_Analysis!E6) |
| PASS | Every input value, in every scenario column, is inside its allowed range | all inside |

## B. Independent re-computation

| Status | Test | Detail |
|---|---|---|
| PASS | Base: independent model vs workbook (13 outputs + goodwill) | largest gap 1.46e-11 GBP k; Python balance sheet gap 9.1e-12; goodwill IFRS 19,170.0 / US GAAP 19,995.0 (gap 0.0e+00) |
| PASS | Upside: independent model vs workbook (13 outputs + goodwill) | largest gap 1.46e-11 GBP k; Python balance sheet gap 9.1e-12; goodwill IFRS 19,170.0 / US GAAP 19,995.0 (gap 0.0e+00) |
| PASS | Downside: independent model vs workbook (13 outputs + goodwill) | largest gap 7.28e-12 GBP k; Python balance sheet gap 9.8e-12; goodwill IFRS 19,170.0 / US GAAP 19,995.0 (gap 0.0e+00) |
| PASS | Custom: independent model vs workbook (13 outputs + goodwill) | largest gap 1.46e-11 GBP k; Python balance sheet gap 9.1e-12; goodwill IFRS 19,170.0 / US GAAP 19,995.0 (gap 0.0e+00) |

## B2. The workbook's own master check, by scenario

| Status | Test | Detail |
|---|---|---|
| PASS | Base: all 39 checks pass | outlier warnings: 0 |
| PASS | Upside: all 39 checks pass | outlier warnings: 0 |
| EXPECTED | Downside: master check red because the deal leaves the pro-forma business with negative cash | Intended: the model refuses to hide a funding gap. See docs/acquisition_memo.md |
| PASS | Custom: all 39 checks pass | outlier warnings: 0 |

## C. Test the tests (deliberately break the model)

| Status | Test | Detail |
|---|---|---|
| PASS | Overwrite one month of the receivables cash-flow line with +100 | master=0; failing checks: Balance sheet balances, Pro-forma balance sheet balances, Static tables current (Base only) |
| PASS | Add 50 to one month of closing Enterprise ARR (breaks the roll-forward) | master=0; failing checks: ARR roll-forward, ARR segments sum to total, Long-term ARR roll-forward, Static tables current (Base only) |
| PASS | Set the tax rate to 90% (outside its allowed range) | master=0; failing checks: Inputs within allowed range |
| PASS | Set one month's billings negative | master=0; failing checks: Deferred revenue roll-forward, Billings never negative, Static tables current (Base only) |
| PASS | Add 1 to one month of total assets | master=0; failing checks: Balance sheet balances |
| PASS | Outlier warning fires when the revenue threshold is set to 0.01% | 23 months flagged (Base threshold 15%: none) |

## D. Error scan

| Status | Test | Detail |
|---|---|---|
| PASS | No error values on 12 output tabs (5,816 formula cells evaluated in 13s) | none |

## E. Web contract (Outputs_Web and named ranges)

| Status | Test | Detail |
|---|---|---|
| PASS | 249 out_* named ranges, all on Outputs_Web | ok |
| PASS | Input registry lists every inp_* name (99) | complete |
| PASS | Each scenario-driven input points at the right Inputs_Scenarios row (the Custom column is where the website writes) | ok |
| PASS | Every inp_* name sits on Inputs_Drivers | ok |
| PASS | Outputs_Web holds only pure links (no calculation of its own) | all links |
| PASS | Every Outputs_Web cell evaluates without error (5,469 cells, 36s) | none |
| PASS | out_master_check reads 1 in Base | 1 |

## F. Portability inventory (Excel, Google Sheets, HyperFormula)

| Status | Test | Detail |
|---|---|---|
| PASS | Only 21 classic functions used (21 distinct) | ABS x356, AND x19, AVERAGE x99, COLUMN x922, COUNT x2, COUNTIF x2, EOMONTH x396, IF x1628, INDEX x352, ISNUMBER x59, LARGE x15, MATCH x16, MAX x191, MIN x153, OR x98, ROUNDUP x792, ROW x19, SUM x751, SUMIF x364, SUMPRODUCT x7, YEAR x37 |
| PASS | SUMPRODUCT is only ever called with two plain ranges (no array arithmetic, which HyperFormula disables by default) | ok |
| PASS | INDEX is always called with explicit row and column (HyperFormula reads a 2-argument INDEX on a one-row range differently from Excel and Sheets) | ok |
| PASS | Formula length and nesting are within Excel limits | longest 144 characters (limit 8,192); deepest nesting 4 (limit 64) |
| REVIEW | Engines not yet tried: Excel and Google Sheets | HyperFormula (the website's engine) reproduces every Outputs_Web cell: see web/tests. Run tests/portability_checklist.md in Excel and Google Sheets to close the remaining gap. |

## Independent model: scenario comparison (GBP k)

| Output | Base workbook / Python | Upside workbook / Python | Downside workbook / Python | Custom workbook / Python |
|---|---|---|---|---|
| FY1 revenue | 48,367.1 / 48,367.1 | 51,168.0 / 51,168.0 | 43,788.9 / 43,788.9 | 48,367.1 / 48,367.1 |
| FY2 revenue | 62,987.3 / 62,987.3 | 71,813.5 / 71,813.5 | 49,504.0 / 49,504.0 | 62,987.3 / 62,987.3 |
| FY1 EBITDA | -6,972.5 / -6,972.5 | -5,719.7 / -5,719.7 | -8,479.3 / -8,479.3 | -6,972.5 / -6,972.5 |
| FY2 EBITDA | -3,129.3 / -3,129.3 | 1,191.6 / 1,191.6 | -7,575.0 / -7,575.0 | -3,129.3 / -3,129.3 |
| FY1 net income | -6,790.5 / -6,790.5 | -5,477.5 / -5,477.5 | -8,413.0 / -8,413.0 | -6,790.5 / -6,790.5 |
| Closing ARR, month 12 | 53,403.6 / 53,403.6 | 58,460.0 / 58,460.0 | 45,146.3 / 45,146.3 | 53,403.6 / 53,403.6 |
| Closing ARR, month 24 | 68,321.2 / 68,321.2 | 80,599.1 / 80,599.1 | 50,209.1 / 50,209.1 | 68,321.2 / 68,321.2 |
| Deferred revenue, month 24 | 31,720.7 / 31,720.7 | 37,258.7 / 37,258.7 | 23,574.8 / 23,574.8 | 31,720.7 / 31,720.7 |
| Receivables, month 24 | 8,821.8 / 8,821.8 | 9,288.4 / 9,288.4 | 8,679.0 / 8,679.0 | 8,821.8 / 8,821.8 |
| Payables, month 24 | 1,593.7 / 1,593.7 | 1,698.8 / 1,698.8 | 1,438.2 / 1,438.2 | 1,593.7 / 1,593.7 |
| Cash, month 12 | 21,629.5 / 21,629.5 | 24,981.4 / 24,981.4 | 15,991.3 / 15,991.3 | 21,629.5 / 21,629.5 |
| Cash, month 24 | 22,579.1 / 22,579.1 | 31,968.1 / 31,968.1 | 9,507.8 / 9,507.8 | 22,579.1 / 22,579.1 |
| Equity, month 24 | 5,740.5 / 5,740.5 | 11,042.9 / 11,042.9 | -611.0 / -611.0 | 5,740.5 / 5,740.5 |

## What this does not cover

- The Vaultline monthly model, the DCF and the pro-forma are not re-implemented in Python (only the Vaultline goodwill is). Their integrity rests on the workbook's checks, which the mutation tests above exercise only for the Northgate model.
- pycel is the recalculation engine here, not Excel or Google Sheets. A difference in how those engines treat a function would not show up.
- Static tables (levers, grids, scenario snapshot) are checked for staleness inside the workbook, not re-derived here.
- The independent Python model was written from the same understanding of the business as the workbook. It catches wiring, formula and arithmetic errors; it cannot catch a shared conceptual mistake (for example a wrong accounting treatment).
- The one review item (the literal 10 in the Lever_Analysis staleness tolerance) is the cut-off between ratio metrics and GBP k metrics. It is deliberate and visible.
