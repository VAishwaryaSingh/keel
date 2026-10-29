// app.js: the Keel website. Every number shown comes from the workbook's own formulas, run in the browser by
// HyperFormula (engine.js). This file only reads out_* named ranges and writes inp_* inputs.
import { loadModel } from './engine.js';
import { fmtK, fmtMoney, fmtPct, fmtX, fmtMo, fmtNum, fmtByUnit, inputDisplay, inputParse, isPercentUnit, excelDate, isNum } from './format.js';
import { ChartHost } from './charts.js';
import { buildDownload } from './download.js';
import { h, $, table, row, downloadText } from './dom.js';
import { mountPlanner } from './scenarios.js';
import { mountSensitivity } from './sensitivity.js';

// ------------------------------------------------------------------------------------------------
// configuration
// ------------------------------------------------------------------------------------------------
const SECTIONS = [
  ['overview', '1. Overview'], ['assumptions', '2. Inputs'], ['revenue', '3. Revenue build'], ['revrec', '4. Revenue recognition'],
  ['costs', '5. Costs & headcount'], ['statements', '6. Three statements'], ['health', '7. SaaS health'], ['budget', '8. Budget vs forecast'],
  ['planner', '9. Scenario planner'], ['sensitivity', '10. Sensitivity tool'], ['acquisition', '11. Acquisition'], ['checks', '12. Checks & AI note'], ['download', '13. Download'],
];
const SHEET_SECTION = { Rev_Build: 'revenue', Rev_Rec: 'revrec', Headcount_Opex: 'costs', Contract_Costs: 'costs', COGS: 'costs', IS: 'statements', BS: 'statements', CF: 'statements',
  SaaS_Metrics: 'health', Budget_vs_Forecast: 'budget', LT_Model: 'overview', MA_Target: 'acquisition', MA_PPA: 'acquisition', MA_ProForma: 'acquisition', MA_Valuation: 'acquisition',
  MA_Returns: 'acquisition', Checks: 'checks', Inputs_Drivers: 'assumptions', Inputs_Scenarios: 'assumptions', Outputs_Web: 'overview' };
const KPIS = [
  ['out_arr12', 'ARR, month 12', fmtMoney], ['out_arr24', 'ARR, month 24', fmtMoney], ['out_rev_fy1', 'Revenue, FY1', fmtMoney], ['out_ebitda_fy1', 'EBITDA, FY1', fmtMoney],
  ['out_nrr12', 'Net revenue retention (month 12)', fmtPct], ['out_payback12', 'CAC payback (month 12)', fmtMo], ['out_cash24', 'Cash, month 24', fmtMoney], ['out_ma_value', 'Acquisition: value created', fmtMoney],
];
const IS_ROWS = [['rev_ent'], ['rev_mid'], ['rev_impl'], ['rev_tot', 1], ['c_host'], ['c_lic'], ['c_support'], ['c_svc'], ['cogs_tot', 1], ['gp', 1], ['gm'], ['o_rd'], ['o_sm'], ['o_ga'], ['opex_tot', 1], ['ebitda', 1], ['ebitda_m'], ['da'], ['ebit', 1], ['interest'], ['pbt', 1], ['tax'], ['ni', 1]];
const BS_ROWS = [['cash'], ['ar'], ['ppe'], ['cca'], ['ta', 1], ['ap'], ['defrev'], ['tl', 1], ['equity', 1], ['check']];
const CF_ROWS = [['ni'], ['da'], ['d_ar'], ['d_ap'], ['d_def'], ['d_cca'], ['cfo', 1], ['capex'], ['cfi', 1], ['net', 1], ['cash_open'], ['cash_close', 1]];
const HEALTH = [
  ['out_sm_nrr', 'Net revenue retention', fmtPct, 'NRR', 'Share of opening ARR kept after expansion and churn, annualised.'],
  ['out_sm_grr', 'Gross revenue retention', fmtPct, 'GRR', 'Share of opening ARR kept after churn only, annualised.'],
  ['out_sm_payback', 'CAC payback', fmtMo, 'CAC payback', 'Months of gross margin on new and expansion ARR needed to repay sales and marketing spend.'],
  ['out_sm_ltv_cac', 'LTV / CAC', fmtX, 'LTV / CAC', 'Lifetime gross-margin value of a customer relative to its acquisition cost.'],
  ['out_sm_burn_mult', 'Burn multiple', fmtX, 'Burn multiple', 'Net cash burned per pound of net new ARR. Lower is better; n/a when not burning.'],
  ['out_sm_rule40', 'Rule of 40', fmtPct, 'Rule of 40', 'ARR growth plus EBITDA margin. 40% or more is the usual benchmark.'],
  ['out_sm_runway', 'Cash runway', fmtMo, 'Cash runway', 'Cash divided by trailing three-month average burn; n/a when cash generative.'],
  ['out_sm_magic', 'Magic number', fmtX, null, 'Three-month net new ARR divided by prior three-month sales and marketing spend.'],
];
const AI_NOTE = [
  'The workbook was laid out by a script directed with Claude Code, then checked by a second, independent implementation of the model written in plain Python.',
  'The independent model agrees with the workbook to within 0.01 GBP k on 13 outputs in every scenario; five deliberate breakages are each caught by the workbook\'s checks.',
  'Assumptions, accounting policy choices and the acquisition recommendation stay human judgements. The manual rebuild of the ARR bridge, deferred revenue and goodwill is the model owner\'s task and has not been recorded here.',
  'Fourteen mistakes were made and caught along the way, including a US GAAP goodwill column that first showed the IFRS number. They are listed in docs/ai_usage.md.',
  'The browser engine (HyperFormula) reproduces all 5,261 Outputs_Web cells in Base, Upside and Downside to within a millionth. It has not yet been compared with Microsoft Excel or Google Sheets.',
];

// ------------------------------------------------------------------------------------------------
// app state
// ------------------------------------------------------------------------------------------------
const S = { listeners: [], ctx: null, model: null, bytes: null, ExcelJS: null, charts: null, fy: 'fy1', stmt: 'is', period: 'fy1', healthMonth: 12, open: {}, fields: {}, cells: {}, debounce: {}, lastRender: null };
const q = name => S.model.get(name);
const ser = name => S.model.series(name);
const monthLabels = () => ser('out_period_end').months.map(excelDate);

// ------------------------------------------------------------------------------------------------
// status bar, panel
// ------------------------------------------------------------------------------------------------
function buildStatus() {
  const sel = h('select', { id: 'scenario-select', 'aria-label': 'Scenario', onchange: e => { S.model.setScenario(e.target.value); render(); } },
    ['Base', 'Upside', 'Downside', 'Custom'].map(n => h('option', { value: n }, n)));
  return h('div', { id: 'statusbar', role: 'status', 'aria-live': 'polite' },
    h('button', { id: 'panel-toggle', class: 'btn', 'aria-controls': 'panel', onclick: () => togglePanel() }, 'Inputs'),
    h('label', { class: 'sb-item' }, 'Scenario ', sel),
    h('a', { id: 'master-light', class: 'light', href: '#checks' }, h('span', { class: 'dot', 'aria-hidden': 'true' }), h('span', { id: 'master-text' }, '')),
    h('button', { id: 'changed-count', class: 'btn linklike', onclick: () => togglePanel(true) }, '0 inputs changed'),
    h('span', { id: 'recalc', class: 'sb-item muted' }, ''),
    h('button', { id: 'theme-toggle', class: 'btn', 'aria-label': 'Toggle light and dark mode', onclick: toggleTheme }, 'Theme'));
}
function toggleTheme() {
  const root = document.documentElement;
  const dark = root.dataset.theme ? root.dataset.theme === 'dark' : globalThis.matchMedia?.('(prefers-color-scheme: dark)').matches;
  root.dataset.theme = dark ? 'light' : 'dark';
  try { localStorage.setItem('keel-theme', root.dataset.theme); } catch { /* storage may be unavailable */ }
  render();
}
function togglePanel(force) {
  const p = $('#panel');
  const open = force === undefined ? !p.classList.contains('open') : force;
  p.classList.toggle('open', open);
  $('#panel-toggle').setAttribute('aria-expanded', String(open));
  // On wide screens the panel is always shown, so "open" means bring it into view.
  if (open) p.scrollIntoView?.({ block: 'nearest', behavior: 'smooth' });
}

function buildPanel() {
  const panel = h('aside', { id: 'panel', 'aria-label': 'Control panel: inputs' });
  panel.append(h('div', { class: 'panel-head' }, h('h2', {}, 'Inputs'),
    h('button', { class: 'btn', id: 'reset-all', onclick: () => { S.model.resetAll(); render(); } }, 'Reset all to Base'),
    h('button', { class: 'btn close', 'aria-label': 'Close inputs', onclick: () => togglePanel(false) }, 'Close')));
  panel.append(h('p', { class: 'muted small' }, 'Editing a scenario-driven input switches to the Custom scenario, starting from the scenario you were viewing. Out-of-range values are blocked and never reach the model.'));
  const groups = new Map();
  for (const r of S.model.registry) {
    if (!groups.has(r.section)) groups.set(r.section, []);
    groups.get(r.section).push(r);
  }
  let first = true;
  for (const [sec, items] of groups) {
    const key = sec.startsWith('Acquisition') || sec.startsWith('Long-term') || sec.startsWith('Timing') || sec.startsWith('People (static)') || sec.startsWith('Costs') || sec.startsWith('Accounting');
    const det = h('details', { class: 'group', open: !key || first }, h('summary', {}, sec, h('span', { class: 'count' }, ` ${items.length}`)));
    first = false;
    for (const r of items) det.append(buildField(r));
    panel.append(det);
  }
  return panel;
}

function buildField(r) {
  const id = 'in-' + r.name;
  const pct = isPercentUnit(r.unit);
  const range = (r.max ?? 0) - (r.min ?? 0);
  const toggle = r.unit === '1/0';
  const step = pct ? 0.001 : range <= 20 ? 0.1 : range <= 200 ? 1 : Math.pow(10, Math.floor(Math.log10(range / 200)));
  const numStep = pct ? 0.1 : step;
  const unitText = pct ? r.unit.replace(/^pts$/, 'pts') : r.unit;
  const err = h('div', { class: 'err', id: id + '-err', role: 'alert' });
  let control;
  if (toggle) {
    control = h('input', { type: 'checkbox', id, onchange: e => apply(r, e.target.checked ? 1 : 0, e.target, err) });
  } else {
    const slider = h('input', { type: 'range', id: id + '-range', min: r.min ?? 0, max: r.max ?? 1, step, 'aria-label': r.label + ' slider', tabindex: '-1' });
    const box = h('input', { type: 'number', id, step: numStep, inputmode: 'decimal' });
    slider.addEventListener('input', () => { box.value = inputDisplay(+slider.value, r.unit); schedule(r, +slider.value, box, err); });
    box.addEventListener('input', () => schedule(r, inputParse(box.value, r.unit), box, err));
    control = h('div', { class: 'ctl' }, slider, box, h('span', { class: 'unit' }, unitText));
    S.fields[r.name] = { r, box, slider, err };
  }
  if (toggle) S.fields[r.name] = { r, box: control, err };
  const lim = r.min === null ? '' : `range ${inputDisplay(r.min, r.unit)} to ${inputDisplay(r.max, r.unit)}${pct ? '%' : ''}`;
  const field = h('div', { class: 'field', dataset: { name: r.name } },
    h('label', { for: id }, r.label, toggle ? h('span', { class: 'unit' }, ' (tick = yes)') : null),
    control,
    h('div', { class: 'meta' }, h('span', { class: 'base' }, ''), ' ', lim ? h('span', {}, '· ' + lim) : '', ' ',
      h('button', { class: 'reset', type: 'button', 'aria-label': `Reset ${r.label} to Base`, onclick: () => { S.model.resetInput(r.name); render(); } }, 'reset')),
    r.tip ? h('div', { class: 'tip' }, r.tip) : null, err);
  S.fields[r.name].field = field;
  return field;
}

function schedule(r, value, box, err) {
  const res = validate(r, value);
  err.textContent = res || '';
  box.setAttribute('aria-invalid', res ? 'true' : 'false');
  clearTimeout(S.debounce[r.name]);
  if (res) return;                       // invalid values never reach the engine
  S.debounce[r.name] = setTimeout(() => apply(r, value, box, err), 300);
}
function validate(r, v) {
  if (!Number.isFinite(v)) return 'Enter a number.';
  const shown = x => inputDisplay(x, r.unit) + (isPercentUnit(r.unit) ? '%' : '');
  if (r.min !== null && v < r.min - 1e-12) return `Minimum is ${shown(r.min)}.`;
  if (r.max !== null && v > r.max + 1e-12) return `Maximum is ${shown(r.max)}.`;
  return '';
}
function apply(r, value, box, err) {
  const res = S.model.setInput(r.name, value);
  err.textContent = res.ok ? '' : res.message;
  if (res.ok) render(box);
}
function updatePanel(skip) {
  for (const [name, f] of Object.entries(S.fields)) {
    const v = S.model.value(name);
    const base = S.model.baseValue(name);
    const pct = isPercentUnit(f.r.unit);
    if (f.r.unit === '1/0') f.box.checked = v === 1;
    else if (f.box !== skip) {
      f.box.value = inputDisplay(v, f.r.unit);
      f.slider.value = v;
    }
    const changed = Math.abs(v - base) > 1e-12;
    f.field.classList.toggle('changed', changed);
    $('.base', f.field).textContent = `Base ${inputDisplay(base, f.r.unit)}${pct ? '%' : ''}${changed ? ' · changed' : ''}`;
  }
  for (const [name, c] of Object.entries(S.cells)) {
    const v = S.model.value(name);
    if (c.r.unit === '1/0') c.box.checked = v === 1;
    else if (c.box !== skip) c.box.value = inputDisplay(v, c.r.unit);
    c.tr.classList.toggle('changed', Math.abs(v - S.model.baseValue(name)) > 1e-12);
  }
}

// ------------------------------------------------------------------------------------------------
// section 2: every input as an editable table (same validation and engine path as the panel)
// ------------------------------------------------------------------------------------------------
function buildInputTables() {
  const groups = new Map();
  for (const r of S.model.registry) {
    if (!groups.has(r.section)) groups.set(r.section, []);
    groups.get(r.section).push(r);
  }
  return [...groups].map(([sec, items]) => h('div', { class: 'in-group' }, h('h3', {}, sec),
    h('table', { class: 'tbl in-tbl' },
      h('thead', {}, h('tr', {}, ['Input', 'Value', 'Unit', 'Base', 'Allowed range', ''].map(x => h('th', { scope: 'col' }, x)))),
      h('tbody', {}, items.map(buildInputRow)))));
}
function buildInputRow(r) {
  const id = 'sec-in-' + r.name;
  const pct = isPercentUnit(r.unit);
  const toggle = r.unit === '1/0';
  const err = h('div', { class: 'err', id: id + '-err', role: 'alert' });
  const box = toggle
    ? h('input', { type: 'checkbox', id, onchange: e => apply(r, e.target.checked ? 1 : 0, e.target, err) })
    : h('input', { type: 'number', id, class: 'in-val', step: pct ? 0.1 : 'any', inputmode: 'decimal' });
  if (!toggle) box.addEventListener('input', () => schedule(r, inputParse(box.value, r.unit), box, err));
  const shown = x => inputDisplay(x, r.unit) + (pct ? '%' : '');
  const tr = h('tr', { dataset: { name: r.name } },
    h('th', { scope: 'row' }, h('label', { for: id }, r.label), r.tip ? h('div', { class: 'tip' }, r.tip) : null),
    h('td', {}, box, err),
    h('td', { class: 'unit' }, toggle ? 'tick = yes' : (pct ? '%' : r.unit)),
    h('td', { class: 'num' }, toggle ? (S.model.baseValue(r.name) === 1 ? 'yes' : 'no') : shown(S.model.baseValue(r.name))),
    h('td', { class: 'muted small' }, r.min === null || toggle ? '' : `${shown(r.min)} to ${shown(r.max)}`),
    h('td', {}, h('button', { class: 'reset', type: 'button', 'aria-label': `Reset ${r.label} to Base`, onclick: () => { S.model.resetInput(r.name); render(); } }, 'reset')));
  S.cells[r.name] = { r, box, err, tr };
  return tr;
}

// ------------------------------------------------------------------------------------------------
// page skeleton
// ------------------------------------------------------------------------------------------------
function workingsBlock(id) {
  return h('div', { class: 'workings' },
    h('button', { class: 'btn small', 'aria-expanded': 'false', 'aria-controls': id + '-w', onclick: e => { S.open[id] = !S.open[id]; e.target.setAttribute('aria-expanded', String(!!S.open[id])); renderWorkings(id); } }, 'Show the workings'),
    h('div', { id: id + '-w', class: 'workings-box', hidden: true }));
}
function buildPage() {
  const app = $('#app');
  app.replaceChildren();
  const nav = h('nav', { id: 'sidenav', 'aria-label': 'Sections' }, h('strong', { class: 'brand' }, 'Keel'),
    h('ul', {}, SECTIONS.map(([id, label]) => h('li', {}, h('a', { href: '#' + id }, label)))));
  const main = h('main', { id: 'main' });
  main.append(h('div', { id: 'banner', role: 'alert', hidden: true }));
  for (const [id, label] of SECTIONS) {
    const sec = h('section', { id, 'aria-labelledby': id + '-h' }, h('h2', { id: id + '-h' }, label));
    main.append(sec);
  }
  app.append(buildStatus(), h('div', { id: 'layout' }, nav, main, buildPanel()), buildTrace());
  buildStatic();
}
function buildTrace() {
  return h('div', { id: 'trace', role: 'dialog', 'aria-modal': 'false', 'aria-labelledby': 'trace-h', hidden: true },
    h('div', { class: 'trace-head' }, h('h3', { id: 'trace-h' }, 'Trace a number'), h('button', { class: 'btn', onclick: () => { $('#trace').hidden = true; } }, 'Close')),
    h('div', { id: 'trace-body' }));
}
function sec(id) { return $('#' + id); }

function buildStatic() {
  // 1 overview
  sec('overview').append(
    h('p', { class: 'lead' }, 'Northgate Controls Ltd is a synthetic UK B2B SaaS company selling compliance and surveillance software to financial institutions. This site runs the Keel workbook\'s own formulas in your browser: change an input and every section recalculates.'),
    h('p', {}, 'The model answers: how fast can ARR grow, what does it cost, when does cash run out, and does buying Vaultline Archive Ltd create value? Click any headline figure to trace where it comes from.'),
    h('div', { id: 'kpis', class: 'kpis' }),
    h('h3', {}, 'Model map'),
    h('ol', { class: 'map' }, [
      'Inputs (named, validated) and scenarios', 'ARR build by segment', 'Billings, ratable revenue and deferred revenue (IFRS 15)', 'Headcount, payroll, commission (capitalised) and COGS',
      'Integrated income statement, balance sheet and cash flow', 'SaaS metrics and the levers behind them', 'Budget variance bridges', 'Vaultline: standalone, IFRS 3 purchase price allocation, pro-forma, valuation and returns', '39 integrity checks'].map(t => h('li', {}, t))),
    workingsBlock('overview'));
  // 2 assumptions
  sec('assumptions').append(
    h('p', {}, 'This is the source of every number in the model. Edit any value below and every section recalculates. Changing a scenario-driven input switches to the Custom scenario; changed rows are highlighted. The same inputs are also in the side panel, so you can edit while looking at another section.'),
    h('div', { class: 'row-ctl' }, h('button', { class: 'btn', type: 'button', onclick: () => { S.model.resetAll(); render(); } }, 'Reset all to Base')),
    ...buildInputTables(), h('div', { id: 'scen-table' }), workingsBlock('assumptions'));
  // 3 revenue
  sec('revenue').append(
    h('p', {}, 'ARR starts at the opening balance. New logos (volume x deal size), expansion and churn move it each month, separately for Enterprise and Mid-market. Churn and expansion are annual rates converted to monthly compounding and applied to opening ARR.'),
    h('div', { class: 'row-ctl' }, h('label', {}, 'Year ', h('select', { id: 'fy-select', onchange: e => { S.fy = e.target.value; render(); } }, h('option', { value: 'fy1' }, 'FY1'), h('option', { value: 'fy2' }, 'FY2')))),
    chartBox('c-waterfall', 'ARR waterfall for the selected year'), chartBox('c-seg', 'Closing ARR by segment, by month'),
    h('div', { id: 'cohort' }), workingsBlock('revenue'));
  // 4 revrec
  sec('revrec').append(
    h('p', {}, 'Both segments are invoiced annually upfront. Under IFRS 15 each subscription invoice is recognised evenly over twelve months, so billings lead revenue and the difference sits in deferred revenue. Implementation fees are a separate performance obligation recognised over three months.'),
    chartBox('c-bill', 'Billings and revenue, by month'), chartBox('c-def', 'Deferred revenue, by month'), h('div', { id: 'dr-roll' }), workingsBlock('revrec'));
  // 5 costs
  sec('costs').append(
    h('p', {}, 'Headcount follows a hiring plan by function, costed at a fully loaded rate with an annual pay rise. Customer Success sits in cost of revenue with hosting and licences; sales commission is capitalised and amortised (IFRS 15).'),
    chartBox('c-heads', 'Closing headcount by function'), chartBox('c-cost', 'Cost of revenue and operating expense, by month'), h('div', { id: 'cost-table' }), workingsBlock('costs'));
  // 6 statements
  sec('statements').append(
    h('p', {}, 'Integrated statements that balance in every period. The balance check shows assets less liabilities less equity.'),
    h('div', { class: 'row-ctl' },
      h('div', { role: 'tablist', 'aria-label': 'Statement' }, [['is', 'Income statement'], ['bs', 'Balance sheet'], ['cf', 'Cash flow']].map(([k, l]) =>
        h('button', { class: 'btn tab', role: 'tab', dataset: { k }, onclick: () => { S.stmt = k; render(); } }, l))),
      h('label', {}, 'Period ', h('select', { id: 'period-select', onchange: e => { S.period = e.target.value; render(); } },
        [h('option', { value: 'fy1' }, 'FY1'), h('option', { value: 'fy2' }, 'FY2'), ...Array.from({ length: 24 }, (_, i) => h('option', { value: 'm' + (i + 1) }, `Month ${i + 1}`))]))),
    h('div', { id: 'stmt' }), h('div', { id: 'stmt-check', class: 'check-line' }), workingsBlock('statements'));
  // 7 health
  sec('health').append(
    h('p', {}, 'Each card shows the metric for the selected month and the drivers that move it.'),
    h('div', { class: 'row-ctl' }, h('label', {}, 'Month ', h('select', { id: 'health-month', onchange: e => { S.healthMonth = +e.target.value; render(); } },
      Array.from({ length: 22 }, (_, i) => h('option', { value: i + 3, selected: i + 3 === 12 }, String(i + 3)))))),
    h('div', { id: 'health-cards', class: 'kpis' }), workingsBlock('health'));
  // 8 budget
  sec('budget').append(
    h('p', {}, 'FY1 board budget mapped to the model. Variance = forecast minus budget, so a positive number is favourable on every line.'),
    h('div', { id: 'bvf' }), chartBox('c-b1', 'FY1 closing ARR: budget to forecast'), chartBox('c-b2', 'FY1 revenue: budget to forecast'), chartBox('c-b3', 'FY1 EBITDA: budget to forecast'), workingsBlock('budget'));
  // 9, 10 (Phase 7)
  mountPlanner(sec('planner'), S.ctx);
  mountSensitivity(sec('sensitivity'), S.ctx);
  // 11 acquisition
  sec('acquisition').append(
    h('p', {}, 'Vaultline Archive Ltd: a smaller, profitable archiving business. The deal terms (offer multiple, debt, synergies, integration cost) and the include-acquisition switch are in the control panel under the Acquisition groups.'),
    h('div', { id: 'ma-kpis', class: 'kpis' }),
    chartBox('c-ff', 'Football field: enterprise value, GBP m'), h('div', { id: 'ppa' }),
    chartBox('c-pfarr', 'ARR: pro-forma against Northgate standalone'), chartBox('c-pfcash', 'Cash: pro-forma against Northgate standalone'), workingsBlock('acquisition'));
  // 12 checks
  sec('checks').append(h('div', { id: 'checks-table' }), h('h3', {}, 'How AI was used'), h('ul', {}, AI_NOTE.map(t => h('li', {}, t))), workingsBlock('checks'));
  // 13 download
  sec('download').append(
    h('p', {}, 'Download the original workbook with your current inputs written into the named input cells. Every formula is kept; Excel recalculates when you open it. A Run_Log tab records the date, scenario and every input changed from Base.'),
    h('button', { id: 'dl-xlsx', class: 'btn primary', onclick: doDownload }, 'Download your model (.xlsx)'), h('p', { id: 'dl-status', class: 'muted', role: 'status' }, ''),
    h('p', { class: 'muted small' }, 'The download omits the football-field chart that exists in the full workbook (model/keel_saas_model.xlsx in the repository). The scenario comparison, bridge, tornado, grids and goal-seek iterations each have their own CSV button in sections 9 and 10.'));
}
function chartBox(id, title) {
  return h('figure', { class: 'chart' }, h('figcaption', {}, title), h('div', { class: 'canvas-wrap' }, h('canvas', { id, role: 'img', 'aria-describedby': id + '-sum', 'aria-label': title })), h('p', { id: id + '-sum', class: 'summary' }, ''));
}
const setSummary = (id, text) => { const e = $('#' + id + '-sum'); if (e) e.textContent = text; };

// ------------------------------------------------------------------------------------------------
// rendering
// ------------------------------------------------------------------------------------------------
function render(skip) {
  const m = S.model;
  renderStatus();
  updatePanel(skip);
  renderOverview();
  renderAssumptions();
  renderRevenue();
  renderRevRec();
  renderCosts();
  renderStatements();
  renderHealth();
  renderBudget();
  renderAcquisition();
  renderChecks();
  for (const id of Object.keys(S.open)) renderWorkings(id);
  S.lastRender = new Date();
  S.listeners.forEach(fn => fn());
  $('#recalc').textContent = `Recalculated ${S.lastRender.toLocaleTimeString('en-GB')} (${m.lastRecalcMs.toFixed(0)} ms)`;
}

function renderStatus() {
  const m = S.model;
  $('#scenario-select').value = m.scenario;
  const okk = m.get('out_master_check') === 1;
  const fails = m.failingChecks();
  const light = $('#master-light');
  light.classList.toggle('ok', okk);
  light.classList.toggle('fail', !okk);
  $('#master-text').textContent = okk ? 'All checks pass' : `Check failure (${fails.length})`;
  const n = m.changedInputs().length;
  $('#changed-count').textContent = `${n} input${n === 1 ? '' : 's'} changed`;
  const banner = $('#banner');
  banner.hidden = okk;
  banner.replaceChildren();
  if (!okk) {
    banner.append(h('strong', {}, 'The model is not internally consistent in this scenario. '), 'Failing: ', fails.map(f => f.name).join('; '), '. ',
      h('a', { href: '#checks' }, 'See all checks'));
  }
}

function renderOverview() {
  const box = $('#kpis');
  box.replaceChildren(...KPIS.map(([name, label, f]) => kpiCard(name, label, f(q(name)))));
}
function kpiCard(name, label, value, sub, index = 0) {
  return h('button', { class: 'kpi', dataset: { trace: name }, onclick: () => showTrace(name, index), 'aria-label': `${label}: ${value}. Trace this number.` },
    h('span', { class: 'kpi-label' }, label), h('span', { class: 'kpi-value' }, value), sub ? h('span', { class: 'kpi-sub' }, sub) : null);
}

function renderAssumptions() {
  const m = S.model;
  const reg = m.registry.filter(r => r.kind === 'scenario');
  const cols = ['Base', 'Upside', 'Downside', 'Custom'];
  const rows = reg.map(r => {
    const base = isPercentUnit(r.unit);
    return row([r.label + (r.unit ? ` (${r.unit})` : ''), ...['C', 'D', 'E', 'F'].map(c => { const v = m.cell('Inputs_Scenarios', `${c}${r.scenarioRow}`); return isNum(v) ? (base ? fmtPct(v) : fmtNum(v, 2)) : ''; })]);
  });
  $('#scen-table').replaceChildren(h('h3', {}, 'Scenario values'), table(['Driver', ...cols], rows, { caption: 'Values of every scenario-driven input; the live scenario is highlighted in the status bar.' }));
}

function renderRevenue() {
  const m = S.model;
  $('#fy-select').value = S.fy;
  const g = n => ser(n)[S.fy];
  const open = g('out_sm_arr_open'), nw = g('out_sm_arr_new'), ex = g('out_sm_arr_exp'), ch = g('out_sm_arr_churn'), cl = g('out_sm_arr_close');
  S.charts.waterfall($('#c-waterfall'), [
    { label: 'Opening ARR', from: 0, to: open, kind: 'total' }, { label: 'New logos', from: open, to: open + nw }, { label: 'Expansion', from: open + nw, to: open + nw + ex },
    { label: 'Churn', from: open + nw + ex, to: open + nw + ex + ch }, { label: 'Closing ARR', from: 0, to: cl, kind: 'total' }]);
  setSummary('c-waterfall', `${S.fy.toUpperCase()}: ARR rises from ${fmtMoney(open)} to ${fmtMoney(cl)}: new logos add ${fmtMoney(nw)}, expansion ${fmtMoney(ex)}, churn removes ${fmtMoney(-ch)}.`);
  const ent = ser('out_rb_arr_close_ent').months, mid = ser('out_rb_arr_close_mid').months;
  S.charts.stacked($('#c-seg'), monthLabels(), [{ label: 'Enterprise', data: ent }, { label: 'Mid-market', data: mid }]);
  setSummary('c-seg', `By month 24 Enterprise ARR is ${fmtMoney(ent[23])} and Mid-market ${fmtMoney(mid[23])}.`);
  const c = n => ser(n);
  const rows = [['Enterprise', 'out_rb_cust_close_ent', 'out_rb_arr_close_ent'], ['Mid-market', 'out_rb_cust_close_mid', 'out_rb_arr_close_mid']].map(([l, cu, ar]) => {
    const cs = c(cu).months, as = c(ar).months;
    return row([l, fmtNum(cs[11], 0), fmtMoney(as[11]), fmtK(as[11] / cs[11], 0), fmtNum(cs[23], 0), fmtMoney(as[23]), fmtK(as[23] / cs[23], 0)]);
  });
  $('#cohort').replaceChildren(h('h3', {}, 'Customers and ARR by segment'), table(['Segment', 'Customers M12', 'ARR M12', 'ARR / customer M12 (£k)', 'Customers M24', 'ARR M24', 'ARR / customer M24 (£k)'], rows));
}

function renderRevRec() {
  const labels = monthLabels();
  const bill = ser('out_rr_bill_tot'), rev = ser('out_rr_rev_tot'), def = ser('out_rr_def_tot');
  S.charts.line($('#c-bill'), labels, [{ label: 'Billings', data: bill.months }, { label: 'Revenue', data: rev.months }]);
  setSummary('c-bill', `Over 24 months billings total ${fmtMoney(bill.months.reduce((a, b) => a + b, 0))} against revenue of ${fmtMoney(rev.months.reduce((a, b) => a + b, 0))}; billings run ahead because customers pay a year upfront.`);
  S.charts.line($('#c-def'), labels, [{ label: 'Deferred revenue', data: def.months }]);
  setSummary('c-def', `Deferred revenue grows from ${fmtMoney(def.opening)} to ${fmtMoney(def.months[23])}.`);
  const mk = (fy, open) => row([fy.toUpperCase(), fmtK(open), fmtK(bill[fy]), fmtK(-rev[fy]), fmtK(def[fy])]);
  $('#dr-roll').replaceChildren(h('h3', {}, 'Deferred revenue roll-forward (£k)'), table(['Year', 'Opening', '+ Billings', '- Revenue recognised', 'Closing'], [mk('fy1', def.opening), mk('fy2', def.fy1)],
    { caption: 'Closing = opening + billings - revenue. Includes implementation fees.' }),
  h('p', { class: 'note' }, 'IFRS 15: subscription revenue is recognised ratably as the service is delivered; the unrecognised part of each upfront invoice is a contract liability (deferred revenue).'));
}

function renderCosts() {
  const labels = monthLabels();
  const hs = [['R&D', 'out_hc_h_close_rd'], ['Sales & marketing', 'out_hc_h_close_sm'], ['G&A', 'out_hc_h_close_ga'], ['Customer Success', 'out_hc_h_close_cs']];
  S.charts.stacked($('#c-heads'), labels, hs.map(([l, n]) => ({ label: l, data: ser(n).months })));
  const fte = ser('out_hc_fte_close');
  setSummary('c-heads', `Headcount grows from ${fmtNum(fte.opening, 0)} to ${fmtNum(fte.months[23], 0)} FTE over 24 months.`);
  const cs = [['Cost of revenue', 'out_is_cogs_tot'], ['R&D', 'out_is_o_rd'], ['Sales & marketing', 'out_is_o_sm'], ['G&A', 'out_is_o_ga']];
  S.charts.stacked($('#c-cost'), labels, cs.map(([l, n]) => ({ label: l, data: ser(n).months })));
  setSummary('c-cost', `Total cost and expense in month 24 is ${fmtMoney(cs.reduce((a, [, n]) => a + ser(n).months[23], 0))} against revenue of ${fmtMoney(ser('out_is_rev_tot').months[23])}.`);
  const rowsDef = [['Revenue', 'out_is_rev_tot', fmtK], ['Cost of revenue', 'out_is_cogs_tot', fmtK], ['Gross margin', 'out_is_gm', fmtPct], ['R&D', 'out_is_o_rd', fmtK], ['Sales & marketing', 'out_is_o_sm', fmtK], ['G&A', 'out_is_o_ga', fmtK],
    ['EBITDA', 'out_is_ebitda', fmtK], ['EBITDA margin', 'out_is_ebitda_m', fmtPct], ['Closing headcount (FTE)', 'out_hc_fte_close', v => fmtNum(v, 0)]];
  $('#cost-table').replaceChildren(h('h3', {}, 'Cost structure by year (£k)'), table(['', 'FY1', 'FY2'], rowsDef.map(([l, n, f]) => row([l, f(ser(n).fy1), f(ser(n).fy2)], /EBITDA$/.test(l) ? 'bold' : ''))));
}

function periodValue(name) {
  const s = ser(name);
  if (S.period === 'fy1') return s.fy1;
  if (S.period === 'fy2') return s.fy2;
  return s.months[+S.period.slice(1) - 1];
}
function renderStatements() {
  const defs = { is: ['Income statement', IS_ROWS, 'is'], bs: ['Balance sheet', BS_ROWS, 'bs'], cf: ['Cash flow statement', CF_ROWS, 'cf'] };
  const [title, rows, pre] = defs[S.stmt];
  document.querySelectorAll('#statements .tab').forEach(b => b.setAttribute('aria-selected', String(b.dataset.k === S.stmt)));
  $('#period-select').value = S.period;
  const trs = rows.map(([k, bold]) => {
    const name = `out_${pre}_${k}`;
    const unit = S.model.seriesUnit(name);
    return row([S.model.seriesLabel(name), fmtByUnit(periodValue(name), unit)], bold ? 'bold' : '');
  });
  const lab = S.period.startsWith('fy') ? S.period.toUpperCase() : `Month ${S.period.slice(1)}`;
  $('#stmt').replaceChildren(table(['£k unless stated', lab], trs, { caption: `${title}, ${lab}` }));
  const chk = periodValue('out_bs_check');
  $('#stmt-check').textContent = `Balance check (assets - liabilities - equity), ${lab}: ${fmtNum(chk, 4)} ${Math.abs(chk) < 0.01 ? '✓ balances' : '✗ DOES NOT BALANCE'}`;
  $('#stmt-check').classList.toggle('bad', Math.abs(chk) >= 0.01);
}

function renderHealth() {
  const i = S.healthMonth - 1;
  const levers = Object.fromEntries(S.model.get('out_levers').map(r => [r[0], r[1]]));
  const cards = HEALTH.map(([name, label, f, lever, def]) => {
    const v = ser(name).months[i];
    const b = kpiCard(name, `${label} (month ${S.healthMonth})`, isNum(v) ? f(v) : (v === '' ? 'n/a' : String(v)), null, S.healthMonth);
    b.append(h('span', { class: 'kpi-def' }, def), lever && levers[lever] ? h('span', { class: 'kpi-lev' }, h('b', {}, 'What moves this: '), levers[lever]) : null);
    return b;
  });
  $('#health-cards').replaceChildren(...cards);
}

function bridgeSteps(rows) {
  const steps = [];
  rows.forEach((r, idx) => {
    if (idx === 0 || idx === rows.length - 1) steps.push({ label: r[0], from: 0, to: r[1], kind: 'total' });
    else steps.push({ label: r[0], from: rows[idx - 1][2], to: r[2] });
  });
  return steps;
}
function renderBudget() {
  const pl = S.model.get('out_bvf_pl');
  const trs = pl.map(r => row([r[0], fmtK(r[1]), fmtK(r[2]), fmtK(r[3]), fmtPct(r[4]), (r[5] === 'Favourable' ? '▲ ' : r[5] === 'Adverse' ? '▼ ' : '● ') + r[5], r[6] || ''],
    /^(Total|Gross|EBITDA)/.test(r[0]) ? 'bold' : ''));
  $('#bvf').replaceChildren(table(['FY1 (£k)', 'Budget', 'Forecast', 'Variance', 'Var %', 'Flag', 'Commentary'], trs, { cls: 'wide' }));
  [['c-b1', 'out_bridge_arr', 'ARR'], ['c-b2', 'out_bridge_rev', 'revenue'], ['c-b3', 'out_bridge_ebitda', 'EBITDA']].forEach(([id, nm, what]) => {
    const rows = S.model.get(nm).filter(r => r[0] !== '');
    const steps = bridgeSteps(rows);
    S.charts.waterfall($('#' + id), steps, { truncate: true });
    const first = rows[0], last = rows[rows.length - 1];
    setSummary(id, `Budget ${what} ${fmtMoney(first[1])} to forecast ${fmtMoney(last[1])}: ${rows.slice(1, -1).map(r => `${r[0].toLowerCase()} ${r[1] >= 0 ? '+' : ''}${fmtK(r[1])}`).join('; ')}. Vertical axis is truncated to show the steps.`);
  });
}

function renderAcquisition() {
  const m = S.model;
  const k = [['out_ma_value', 'Value created (NPV)', fmtMoney, q('out_ma_verdict')], ['out_ma_breakeven', 'Synergies needed to break even', fmtPct, 'of plan'],
    ['out_ma_payback', 'Payback', v => (isNum(v) ? fmtNum(v, 1) + ' years' : String(v)), ''], ['out_ma_accretion', 'ARR accretion, month 24', fmtPct, ''],
    ['out_ma_premium', 'Premium over standalone DCF', fmtMoney, ''], ['out_pf_cash24', 'Pro-forma cash, month 24', fmtMoney, ''],
    ['out_pf_mincash', 'Pro-forma lowest cash', fmtMoney, q('out_ma_funding_flag')], ['out_goodwill_ifrs', 'Goodwill (IFRS 3)', fmtMoney, `US GAAP ${fmtMoney(q('out_goodwill_usgaap'))}`]];
  $('#ma-kpis').replaceChildren(...k.map(([n, l, f, sub]) => kpiCard(n, l, f(q(n)), sub)));
  const ff = m.get('out_football').filter(r => r[0] !== '');
  S.charts.range($('#c-ff'), ff.map(r => ({ label: r[0], low: r[1], high: r[3], highlight: /^Offer/.test(r[0]) })));
  setSummary('c-ff', ff.map(r => `${r[0]}: ${fmtNum(r[1], 1)} to ${fmtNum(r[3], 1)} (mid ${fmtNum(r[2], 1)})`).join('; ') + '. Figures in GBP m.');
  const ppa = m.get('out_ppa').filter(r => r[0] !== '');
  $('#ppa').replaceChildren(h('h3', {}, 'Purchase price allocation (£k)'), table(['Line', 'IFRS 3', 'US GAAP (ASU 2021-08)'], ppa.map(r => row([r[0], isNum(r[1]) ? (Math.abs(r[1]) < 20 && !Number.isInteger(r[1]) ? fmtNum(r[1], 3) : fmtK(r[1])) : '', isNum(r[2]) ? (Math.abs(r[2]) < 20 && !Number.isInteger(r[2]) ? fmtNum(r[2], 3) : fmtK(r[2])) : ''],
    /^(Equity|Net assets|Fair value|Goodwill =)/.test(r[0]) ? 'bold' : ''))));
  const labels = monthLabels();
  const pf = ser('out_pf_pf_arr'), sa = ser('out_sm_arr_close');
  S.charts.line($('#c-pfarr'), labels, [{ label: 'Pro-forma', data: pf.months }, { label: 'Northgate standalone', data: sa.months, dash: [6, 4] }]);
  setSummary('c-pfarr', `Month-24 ARR: pro-forma ${fmtMoney(pf.months[23])} against standalone ${fmtMoney(sa.months[23])}${m.value('inp_ma_include') === 0 ? ' (acquisition switched off: the lines coincide)' : ''}.`);
  const pc = ser('out_pf_pf_cash'), sc = ser('out_bs_cash');
  S.charts.line($('#c-pfcash'), labels, [{ label: 'Pro-forma', data: pc.months }, { label: 'Northgate standalone', data: sc.months, dash: [6, 4] }]);
  setSummary('c-pfcash', `Month-24 cash: pro-forma ${fmtMoney(pc.months[23])} against standalone ${fmtMoney(sc.months[23])}; the lowest pro-forma month-end balance is ${fmtMoney(q('out_pf_mincash'))}.`);
}

function renderChecks() {
  const rows = S.model.get('out_checks').filter(r => r[0] !== '');
  const trs = rows.map(r => row([r[0], fmtNum(r[1], 4), fmtNum(r[2], 4), r[3] === 1 ? '✓ pass' : '✗ FAIL'], r[3] === 1 ? '' : 'bad'));
  $('#checks-table').replaceChildren(
    h('p', {}, `${rows.filter(r => r[3] === 1).length} of ${rows.length} checks pass. Outlier warnings: ${q('out_warnings')} series flagged for review (they do not fail the model).`),
    table(['Check', 'Value', 'Tolerance', 'Result'], trs, { caption: 'Every integrity check in the workbook. A failure is never hidden.' }));
}

// ------------------------------------------------------------------------------------------------
// show the workings, trace a number, download
// ------------------------------------------------------------------------------------------------
const WORKINGS = {
  overview: ['out_lt_arr_close', 'out_arr24'], assumptions: [], revenue: ['out_rb_logos_ent', 'out_rb_acv_ent', 'out_rb_churn_m_ent', 'out_rb_exp_m_ent', 'out_rb_arr_new_ent', 'out_rb_arr_exp_ent', 'out_rb_arr_churn_ent', 'out_rb_arr_close_ent', 'out_rb_cust_close_ent'],
  revrec: ['out_rr_inv_tot', 'out_rr_rev_sub', 'out_rr_def_sub_close', 'out_rr_ar'], costs: ['out_hc_pay_tot', 'out_hc_opex_sm', 'out_cc_a_amort', 'out_cogs_host'],
  statements: ['out_is_rev_tot', 'out_is_ebitda', 'out_is_ni', 'out_bs_cash', 'out_bs_defrev', 'out_bs_check', 'out_cf_cfo'], health: ['out_sm_nrr', 'out_sm_payback', 'out_sm_burn_mult', 'out_sm_runway'],
  budget: [], acquisition: ['out_pf_pf_cash', 'out_pf_pb_check', 'out_mt_ufcf'], checks: [],
};
function renderWorkings(id) {
  const box = $('#' + id + '-w');
  if (!box) return;
  box.hidden = !S.open[id];
  if (!S.open[id]) return;
  const m = S.model;
  const names = WORKINGS[id] || [];
  const rows = names.map(n => {
    const isSeries = m.addr[n].c1 > m.addr[n].c0;
    const s1 = m.source(n, isSeries ? 1 : 0), s13 = isSeries ? m.source(n, 13) : null;
    if (!s1) return null;
    const f1 = m.prettyFormula(m.formula(s1.sheet, s1.a1)) ?? String(m.cell(s1.sheet, s1.a1));
    const f13 = s13 ? m.prettyFormula(m.formula(s13.sheet, s13.a1)) : '';
    return h('tr', {}, h('th', { scope: 'row' }, m.seriesLabel(n)), h('td', {}, `${s1.sheet}!${s1.a1}`), h('td', { class: 'code' }, f1), h('td', { class: 'code' }, f13 || ''));
  }).filter(Boolean);
  box.replaceChildren(
    h('p', { class: 'small muted' }, 'These are the actual cells and formulas in the workbook. Month 1 and month 13 are shown so you can see the same formula runs across periods (references shift by column; named inputs stay fixed).'),
    names.length ? h('table', { class: 'tbl wide' }, h('thead', {}, h('tr', {}, ['Line', 'Source cell (month 1)', 'Formula, month 1', 'Formula, month 13'].map(x => h('th', { scope: 'col' }, x)))), h('tbody', {}, rows)) : h('p', {}, 'This section shows workbook outputs directly; see the Revenue build and Three statements workings for formulas.'));
}

function showTrace(name, index = 0) {
  const m = S.model;
  const src = m.source(name, index);
  const body = $('#trace-body');
  body.replaceChildren();
  const label = KPIS.find(k => k[0] === name)?.[1] || HEALTH.find(k => k[0] === name)?.[1] || name;
  if (!src) { body.append(h('p', {}, `${label} has no single source cell.`)); $('#trace').hidden = false; return; }
  const pre = m.precedents(src.sheet, src.a1);
  const f = m.prettyFormula(pre.formula);
  const sectionId = SHEET_SECTION[src.sheet] || 'overview';
  body.append(
    h('p', {}, h('strong', {}, label), ': ', String(m.cell(src.sheet, src.a1))),
    h('p', {}, 'Lives in ', h('code', {}, `${src.sheet}!${src.a1}`), ' as ', h('code', {}, f || '(a value)')),
    pre.refs.length ? table(['Direct precedent', 'Current value'], pre.refs.map(r => row([r.kind === 'input' ? `${r.text} (input)` : m.prettyFormula(r.text), isNum(r.value) ? fmtNum(r.value, 4) : String(r.value ?? '')]))) : h('p', {}, 'This cell is a typed value or has no precedents.'),
    h('p', {}, h('a', { href: '#' + sectionId, onclick: () => { $('#trace').hidden = true; } }, `Jump to the ${SECTIONS.find(s => s[0] === sectionId)[1]} section`)));
  $('#trace').hidden = false;
}

async function doDownload() {
  const st = $('#dl-status');
  try {
    st.textContent = 'Preparing…';
    const out = await buildDownload(S.ExcelJS, S.bytes, S.model);
    const blob = new Blob([out], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
    const url = URL.createObjectURL(blob);
    const a = h('a', { href: url, download: `keel_model_${S.model.scenario.toLowerCase()}.xlsx` });
    document.body.append(a);
    a.click();
    a.remove();
    setTimeout(() => URL.revokeObjectURL(url), 5000);
    st.textContent = `Downloaded ${(out.length / 1024).toFixed(0)} KB with ${S.model.changedInputs().length} changed input(s) and a Run_Log tab.`;
  } catch (e) {
    st.textContent = `Download failed: ${e.message}`;
  }
}

// ------------------------------------------------------------------------------------------------
// start
// ------------------------------------------------------------------------------------------------
export async function startApp({ bytes, names, Chart, ExcelJS }) {
  S.bytes = bytes;
  S.ExcelJS = ExcelJS;
  S.model = await loadModel(bytes, names);
  S.charts = new ChartHost(Chart);
  S.ctx = { model: () => S.model, charts: () => S.charts, rerender: () => render(), monthLabels, onChange: fn => S.listeners.push(fn) };
  buildPage();
  render();
  $('#loading')?.setAttribute('hidden', '');
  globalThis.keel = { model: S.model, state: S, render };
  return S;
}

async function boot() {
  const loading = $('#loading');
  try {
    try { const t = localStorage.getItem('keel-theme'); if (t) document.documentElement.dataset.theme = t; } catch { /* no storage */ }
    const [bytes, names] = await Promise.all([fetch('model/keel_saas_model.xlsx').then(r => { if (!r.ok) throw new Error('model file not found'); return r.arrayBuffer(); }),
      fetch('model/named_ranges.json').then(r => r.json())]);
    await startApp({ bytes: new Uint8Array(bytes), names, Chart: globalThis.Chart, ExcelJS: globalThis.ExcelJS });
  } catch (e) {
    if (loading) loading.textContent = `Could not start: ${e.message}. Serve this folder over HTTP (for example: python -m http.server) rather than opening the file directly.`;
    console.error(e);
  }
}
if (!globalThis.__KEEL_NO_AUTOSTART && typeof document !== 'undefined') boot();
