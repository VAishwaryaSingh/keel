# Keel website

A static, no-backend site that runs the Keel workbook's **own formulas** in the browser. Nothing is re-implemented in JavaScript: change an input and HyperFormula recalculates the same cells that Excel would.

## Run it

```
cd web
python3 -m http.server 8000        # then open http://localhost:8000
npm install && npm test            # engine, download, browser-build and UI tests (Node 20)
```

Opening `index.html` straight from disk will not work (the page fetches the workbook), hence the server. The page loads four pinned libraries from jsDelivr with integrity hashes: HyperFormula 2.7.1, SheetJS 0.18.5, ExcelJS 4.4.0, Chart.js 4.4.1.

## How it is built

| File | Role |
|---|---|
| `engine.js` | Parses the `.xlsx` with SheetJS, builds a HyperFormula instance, resolves named ranges, reads `out_*`, writes `inp_*`, handles scenarios, validation, trace and formula display. No financial logic. |
| `app.js` | Page, control panel, status bar, 13 sections, charts, trace panel, "show the workings", download. |
| `download.js` | ExcelJS: writes current inputs into the original workbook, sets recalculate-on-open, adds `Run_Log`. |
| `scenarios.js`, `sensitivity.js` | Sections 9 and 10: planner (presets, saved scenarios, comparison, input diff, flags, acquisition overlay, bridge) and sensitivity tool (tornado, two-way grid, goal seek, acquisition grid, CSV export). |
| `tools.js` | The calculations behind them: scenario evaluation, bridge, tornado, grid, bisection. No spreadsheet logic: it sets inputs and reads the workbook. |
| `dom.js`, `charts.js`, `format.js` | DOM helpers, Chart.js builders (light/dark aware), number formatting. |
| `model/keel_saas_model.xlsx` | The workbook the site runs. Identical to `../model/keel_saas_model.xlsx` except the football-field chart is removed (ExcelJS cannot read openpyxl chart drawings). |
| `model/named_ranges.json` | Every named range and its address. |
| `tests/` | Seven suites (`engine`, `download`, `umd`, `tools`, `ui`, `ui_tools`, `ui_storage`; 154 checks) and the pycel fixtures they compare against. |

### The contract

- The site **reads** only the `out_*` ranges on `Outputs_Web` (scalars, 24-month series, bridges, tables) and **writes** only `inp_*` inputs.
- A scenario-driven input is a formula on `Inputs_Drivers`. Editing one copies the scenario you were viewing into the **Custom** column of `Inputs_Scenarios`, writes your value there, and selects Custom. Static inputs are written directly.
- Out-of-range values are blocked in the browser and never sent to the engine. Limits come from the workbook.
- If any workbook check fails, the status bar goes red, a banner names the failing checks, and nothing is hidden.

### The what-if engine

The planner and sensitivity tools need dozens of full recalculations. They run on a second, lean HyperFormula instance (built on first use, about 0.6 s) that omits the 5,000+ link cells on `Outputs_Web` and reads each output from the cell its link points at. A run costs about 63 ms instead of 91 ms, so a 7 x 7 grid takes about 3.8 s. A what-if session sets every varied input in one batch, then restores the instance, and never touches the model behind the page: your inputs stay exactly as they were.

### Two engine workarounds (both found by the tests)

1. **Named ranges in formulas.** HyperFormula's named expressions failed to evaluate when a name pointed at a cell that itself held a formula. The loader replaces each single-cell name (`inp_*`, `chk_*`) with its direct address before loading. The semantics are identical; the "workings" view maps addresses back to names.
2. **`INDEX` on a one-row range.** Excel and Sheets treat `INDEX(row_range, n)` as the n-th column; HyperFormula treats it as a row. The workbook now always uses `INDEX(range, 1, n)`. This was fixed in the workbook, not in website code, and the validator enforces it.

## What has been verified

- HyperFormula reproduces all 5,469 `Outputs_Web` cells in Base, Upside and Downside to within one part in a million of an independent recalculation (pycel), using both the npm builds and the exact browser (UMD) builds.
- A single input change recalculates in about 100 ms, a scenario switch in about 180-260 ms (Node 20; target under 300 ms).
- The downloaded `.xlsx` keeps all 16,450 formulas, 355 defined names, data validation and the recalculate flag; reopening it in HyperFormula gives identical outputs.
- Sections 9 and 10 are verified against independent recalculations: tornado (27 drivers x 8 outputs, worst relative gap 4e-15, same top three), one grid cell, goal seek typed back into the workbook by a second engine, and a scenario bridge whose steps sum exactly to the scenario figure.
- 41 + 42 + 4 UI checks pass against a simulated DOM: debounce, validation, Custom scenario, reset, changed-input marker, red status bar in Downside, statement tabs, trace, workings, download, and a site that still works with browser storage blocked.

## What has **not** been verified

- **No real browser has run this.** Layout, colours, dark mode, the phone layout, keyboard flow and chart appearance are untested. The UI tests check behaviour, not looks. Please open it and look.
- **First-load time is not measured in a browser.** In Node the workbook builds in about 1.8 s and first render completes in about 2.9 s, plus roughly 2.7 MB of libraries to download. The 3-second target on broadband is uncertain.
- The recalculation reference is pycel, not Excel. Opening the workbook in Excel and Sheets is still open (`../tests/portability_checklist.md`).
- Deployed to GitHub Pages by `.github/workflows/pages.yml`: https://vaishwaryasingh.github.io/keel/

## Licence

GPL-3.0 (see `../LICENSE`), because HyperFormula is GPL-3.0-only. Anyone deploying this under another licence needs a HyperFormula commercial licence.
