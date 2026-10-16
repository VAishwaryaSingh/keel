# Website acceptance tests (brief section 5.9)

Run with `cd web && npm install && npm test` (154 automated checks), then `python validation/verify_goalseek.py web/model/keel_saas_model.xlsx model/snap_spec.json web/tests/fixtures/goalseek_result.json` for the independent half of test 5.
"Reference" means the independent pycel recalculation of the workbook, not Excel.

| # | Test | Status | Evidence |
|---|---|---|---|
| 1 | Base, Upside and Downside outputs on the site match the spreadsheet exactly for 10 named outputs | **Passes against the reference.** Not yet against Excel | `engine.test.mjs`: 10 named outputs per scenario match `expected_values.csv`; all 5,469 `Outputs_Web` cells match to 1e-6 relative. Also with the exact browser (UMD) builds (`umd.test.mjs`) |
| 2 | Changing any input and downloading gives an `.xlsx` that, opened in Excel, shows the same outputs and a green master check | **Partly.** Round trip through HyperFormula identical; **Excel not tried** | `download.test.mjs`: 16,450 formulas intact, inputs written, outputs and master check identical after reopening. **Open: do this in Excel.** |
| 3 | The tornado ranking matches a manual +/-10% test on the top 3 drivers in the sheet | **Passes against the reference** | `tools.test.mjs`: all 27 flexed drivers x 8 outputs match the independent engine (worst relative gap 4e-15); top 3 identical for ARR (new Enterprise logos, Enterprise ACV, Enterprise expansion), cash and FY2 EBITDA (R&D cost per FTE, S&M cost per FTE, Enterprise ACV) and deal value (WACC, offer multiple, synergies achieved) |
| 4 | One two-way grid cell matches a manual recalculation in the sheet | **Passes against the reference** | `tools.test.mjs`: logos 2.5 per month x churn 8% gives ARR 63,922.6, cash 17,760.1 and lowest cash matching the independent recalculation |
| 5 | Goal seek returns a churn rate that, typed back into the sheet, gives the target within tolerance | **Passes** (typed back into the model by two different engines) | Target: lowest month-end cash 15,000. Answer: Enterprise churn 16.2759%, 22 bisections. HyperFormula: 15,000.0009. pycel, typed into the workbook: 15,000.000922. A second goal seek finds the break-even offer price (5.70x ARR) |
| 6a | Out-of-range input is blocked | **Passes** | `ui.test.mjs`: 90% churn rejected ("Maximum is 50%"), model unchanged; the same for grid limits and the tornado flex in `ui_tools.test.mjs` |
| 6b | A deliberately broken check shows red on the site | **Passes** | Downside breaks "Pro-forma cash never negative": red light, banner naming the check, one failing row. Workbook-level mutation tests are in `validation/validation_report.md` |
| 7 | The site works with browser storage disabled | **Passes** | `ui.test.mjs` (theme toggle with storage throwing) and `ui_storage.test.mjs`: Save is disabled with an explanation; presets and comparison still work |

## Non-functional targets (5.8)

| Target | Result |
|---|---|
| Recalculation under 300 ms for a single input change | About 100 ms (Node 20, full engine) |
| 7 x 7 grid (49 recalculations) in under 5 s | **3.8-3.9 s** (Node 20, lean engine at about 63 ms per run). A tornado of 27 drivers (55 runs) takes 2.4 s; a 19-driver bridge 1.1 s |
| First load under 3 s on broadband | **Unmeasured in a browser.** Node: 1.8 s to build the model, 2.9 s to first render, plus about 2.7 MB of libraries, plus about 0.6 s the first time a what-if tool is used |
| Works on phone and laptop, light and dark mode | **Built, not seen.** No real browser has rendered the CSS |
| Accessible: labelled fields, keyboard navigation, text summaries, colour never the only signal | Labels, text summaries, `aria-live` status, symbols beside colour (heatmap numbers, value-destroying cells marked with a triangle) are tested; keyboard flow and contrast are **not** tested in a browser |
| No sign-in, no tracking, no personal data | True by construction; storage holds only an optional theme preference and up to four saved scenarios |
