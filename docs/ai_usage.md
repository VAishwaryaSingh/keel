# How AI was used on this project, and where it went wrong

The build was directed with Claude Code. This note says what was automated, what stays a human judgement, and every material mistake that was caught (22 so far), with what caught it. It is meant to be read alongside `validation/validation_report.md`.

## What the AI automated

| Task | How |
|---|---|
| Laying out the workbook | `model/build_model.py` writes every formula, named range, format and check with openpyxl. All calculation lives in the workbook, not the script. |
| Recalculating without Excel | `pycel` evaluates the workbook so every scenario can be run and every check read (`validation/recalc_report.py`). |
| Lever, sensitivity and scenario snapshots | `validation/snapshots.py` recalculates the whole workbook once per lever and grid point (about 50 runs, under 10 seconds). |
| Independent verification | `validation/validate_model.py`: a second implementation of the Northgate model in plain Python, a goodwill re-computation, scans for hardcodes and inconsistent formulas, "test the tests" mutations, an error scan. |
| First drafts of documents | Assumption log, Q3 business review, acquisition memo, this note. Every figure in them was read from the model, then checked again. |

## What stays human judgement

- **Assumptions:** churn, expansion, hiring, the offer multiple, synergy size, the share of price allocated to intangibles, WACC. The AI proposed values; none is derived from real data. They are the owner's to set and defend.
- **Accounting policy:** the treatment of commission (IFRS 15), deferred revenue at fair value (IFRS 3 vs ASU 2021-08), and goodwill is a modelling choice a finance professional must understand and be able to defend line by line.
- **The recommendation:** `docs/acquisition_memo.md` is a draft argument, not advice. Its thresholds (4.0x ARR, funding headroom) are judgement calls.
- **Hand verification:** the brief requires rebuilding the ARR bridge, deferred revenue roll-forward and goodwill independently on a blank sheet (Phase 5). That has **not** been done. A kit is ready in `tests/` (blank template, inputs, 20 target numbers, instructions in words, no formulas): it is the model owner's task. The Python re-implementation is not a substitute: it was written from the same understanding as the workbook.
- **Looking at the website:** the UI was tested against a simulated DOM only. Layout, colour, dark mode, the phone layout and chart appearance have not been seen in a real browser.
- **Opening the file in Excel and Google Sheets:** not yet done. All results come from `pycel`. `tests/portability_checklist.md` is the procedure.

## Where the AI's output was wrong, and what caught it

| # | Mistake | Caught by | Fixed |
|---|---|---|---|
| 1 | Tried a general spreadsheet library for recalculation first: one run took **95 minutes** and crashed at the end on a formatting bug. | Waiting for it; the tool was the wrong choice, not the model. | Switched to `pycel` (about 6 seconds). |
| 2 | First calibration made the "not yet profitable" company profitable within 24 months, and NRR implausibly low. | Reading the outputs against the brief. | Re-tuned hiring, marketing and expansion twice. |
| 3 | **Annual "opening" columns were wrong** (FY1 opening ARR showed 0; FY2 showed month 12's opening). | Building the budget bridge in Phase 2. None of the Phase 1 checks tested annual columns. | Fixed; three annual-column checks added to `Checks`, and the validation script now scans for the class of error. |
| 4 | The long-term view referenced rows before they were registered, so the build crashed. | Running the build. | Two-pass layout. |
| 5 | Q3 review draft overstated a churn sensitivity (said £0.7m; the lever table says about £1.2m per point) and mis-worded a hiring sensitivity. | Checking each sentence against `Lever_Analysis`. | Corrected. |
| 6 | README quoted the wrong Downside FY1 EBITDA (−9,164 instead of −8,479). | Re-reading against the snapshot file. | Corrected. |
| 7 | **US GAAP goodwill column was identical to IFRS** (the fair-value row reused the IFRS deferred revenue adjustment). | Reading the PPA output side by side. The reconciliation check passed anyway, because it was circular for that column. | Fixed; the validation script recomputes both goodwill figures independently for every scenario. |
| 8 | Acquisition memo said goodwill and intangibles were "64% of the price". They are about 100%. | Re-checking each memo claim against the PPA. | Corrected. |
| 9 | Used `MEDIAN` for the comps mid-point; the verification engine does not support it, so the cell could not be checked. | The error scan output. | Replaced with the average, and relabelled. |
| 10 | Five axis values on `Sensitivity` were typed numbers in black, breaking the colour convention; several thresholds (24-month payback, the ±0.5% "on budget" band, DCF bands) were hardcoded inside formulas. | The validation script's marking scan and literal scan (its first run). | Marked blue; thresholds moved to visible input cells. |
| 11 | The first Phase 5 build crashed: a regular expression in my generator was written with a doubled backslash, so a cell reference could not be parsed. | Running the build. | Fixed. |
| 12 | The first version of the test-kit script failed with a Python syntax error (a backslash inside an f-string). | Running it. | Fixed. |
| 13 | The input registry on `Outputs_Web` carried typed "scenario row" numbers that could drift from the real rows. | The validation script's marking scan (second run). | Replaced by `ROW()` formulas, and the validator now cross-checks each one against the label it points at. |
| 14 | Two checks used array arithmetic inside `SUMPRODUCT`: fine in Excel and Sheets, but HyperFormula disables array arithmetic by default, so they would have broken in the website. | My own portability review, then a validator rule so it cannot return. | Replaced by plain helper rows. |
| 15 | Wrote the website loader assuming SheetJS keeps formula cells. It silently **drops formula cells that have no cached value** (the workbook has none), so HyperFormula received an almost empty workbook and every output was blank. | The first engine test: 4,898 of 5,261 cells differed. | Read with `sheetStubs: true`. |
| 16 | Assumed HyperFormula's named expressions behave like Excel's. Pointing a name at a cell that itself holds a formula (`inp_scen_idx`) made evaluation fail outright. | Bisecting the workbook sheet by sheet. | The loader substitutes direct addresses for single-cell names. |
| 17 | Wrote `INDEX(row_range, n)` throughout the workbook, which Excel and Sheets read as a column index and HyperFormula reads as a row. Outputs silently became errors in Upside and Downside only. | The engine test, scenario by scenario (Base worked by luck because the index was 1). | Fixed in the workbook (`INDEX(range,1,n)`), and the validator now rejects two-argument `INDEX`. |
| 18 | Assumed ExcelJS could read any workbook this project produces. It crashes on the football-field chart. | The first download test. | The build also writes a chart-free copy for the site; the full workbook keeps the chart. |
| 19 | A UI assertion expected named inputs inside the "workings" for the revenue section, but the rows chosen were built from other rows. | The UI test. | Exposed the driver rows on `Outputs_Web`. |
| 20 | Read workbook values **inside** a HyperFormula batch, where evaluation is suspended. It happened twice: once in the input editor, again in the what-if session start. | The tests (an `EvaluationSuspendedError` each time). | Read first, then batch. |
| 21 | Shipped the planner module with one extra closing parenthesis. The UI test then died silently, because the test itself swallowed the load error and exited with no output. | Noticing a test run that printed nothing. | Fixed the syntax; the shared test set-up now fails loudly. |
| 22 | The tools tests contained wrong test code three times (counting both the comparison and the diff table's headers, an operator-precedence slip in a "direction" assertion, using the same driver on both heatmap axes). Each looked like a product failure. | Reading the failing assertion against the intent, rather than loosening it. | Tests corrected; the product behaviour was right each time. |

## Limits of the checks, stated plainly

- The independent model shares the author's understanding. It would not catch a conceptual error that both versions make.
- The "test the tests" mutations cover the Northgate model. The Vaultline model, DCF and pro-forma are covered by the workbook's own checks and are not re-implemented in Python.
- A passing check proves that things tie, not that the assumptions are right. A tie-out (the PPA reconciliation, in mistake 7) can be circular and pass while the number is wrong.
- Everything here is synthetic. No real company, market data or filing was used.
