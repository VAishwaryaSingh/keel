// scenarios.js: the scenario planner (section 9). Presets, saved custom scenarios (browser storage, optional),
// side-by-side comparison, input diff, downside flags, acquisition overlay and the scenario bridge.
// All figures come from the workbook through the lean what-if engine (tools.js); this file only lays them out.
import { h, $, table, row, progressBar, downloadText } from './dom.js';
import { fmtMoney, fmtPct, fmtX, fmtNum, isNum, isPercentUnit, inputDisplay } from './format.js';
import * as T from './tools.js';

const KEY = 'keel-scenarios';
export const MAX_SAVED = 4;

export function storageOk() {
  try { localStorage.setItem('__keel', '1'); localStorage.removeItem('__keel'); return true; } catch { return false; }
}
export function loadSaved() {
  try { const v = JSON.parse(localStorage.getItem(KEY) || '[]'); return Array.isArray(v) ? v.slice(0, MAX_SAVED) : []; } catch { return []; }
}
function persist(list) {
  try { localStorage.setItem(KEY, JSON.stringify(list)); return true; } catch { return false; }
}

const fmtNo = (v, kind) => {
  if (!isNum(v)) return v === null || v === undefined || v === '' ? 'n/a' : String(v);
  if (kind === 'pct') return fmtPct(v);
  if (kind === 'x') return fmtX(v);
  if (kind === 'mo') return v >= T.NO_LIMIT ? 'no limit' : `${fmtNum(v, 1)} mo`;
  if (kind === 'ok') return v === 1 ? 'pass' : 'FAIL';
  return fmtMoney(v);
};
const ROWS = [
  ['ARR', 'arr_y1', 'arr_y3', 'money'], ['Revenue', 'rev_y1', 'rev_y3', 'money'], ['EBITDA', 'ebitda_y1', 'ebitda_y3', 'money'],
  ['Closing cash', 'cash_y1', 'cash_y3', 'money'], ['Net revenue retention', 'nrr_y1', 'nrr_y3', 'pct'], ['Burn multiple', 'burn_y1', 'burn_y3', 'x'],
  ['Cash runway (month 12 | month 24)', 'runway_m12', 'runway_m24', 'mo'], ['CAC payback (month 12 | month 24)', 'payback_m12', 'payback_m24', 'mo'],
  ['Pro-forma ARR | cash, month 24', 'pf_arr24', 'pf_cash24', 'money'], ['Deal value created | model checks', 'ma_value', 'master', 'mixed'],
];

export function mountPlanner(root, ctx) {
  const st = { saved: loadSaved(), selected: new Set(['Base', 'Downside']), overlay: 'with', results: null, bridge: null, stamp: null };
  const canSave = storageOk();
  const prog = progressBar();
  const msg = h('p', { id: 'pl-msg', class: 'small', role: 'status' }, '');
  const savedBox = h('div', { id: 'pl-saved' });
  const picker = h('fieldset', { class: 'picker' }, h('legend', {}, 'Compare 2 to 4 scenarios'));
  const results = h('div', { id: 'pl-results' });
  const bridgeBox = h('div', { id: 'pl-bridge' });
  const nameBox = h('input', { type: 'text', id: 'pl-name', maxlength: 30, placeholder: 'e.g. Slow hiring', 'aria-label': 'Name for the scenario' });
  const saveBtn = h('button', { class: 'btn', id: 'pl-save', disabled: !canSave, onclick: save }, 'Save current inputs as a scenario');
  const compareBtn = h('button', { class: 'btn primary', id: 'pl-run', onclick: runCompare }, 'Compare');

  root.append(
    h('p', {}, 'Build a custom scenario by choosing a preset in the status bar and editing drivers in the control panel, then save it here (up to 4, kept in this browser only). Compare any 2 to 4 scenarios side by side; every figure is recalculated by the workbook\'s own formulas.'),
    h('div', { class: 'saver' }, h('label', { for: 'pl-name' }, 'Name '), nameBox, ' ', saveBtn,
      canSave ? null : h('span', { class: 'err' }, ' Saving is unavailable: this browser is blocking storage. Presets and comparison still work.')),
    savedBox, msg,
    picker,
    h('div', { class: 'row-ctl' },
      h('fieldset', { class: 'inline' }, h('legend', {}, 'Acquisition overlay'),
        ...['with', 'without'].map(v => h('label', {}, h('input', { type: 'radio', name: 'pl-overlay', value: v, checked: v === 'with', onchange: () => { st.overlay = v; } }), v === 'with' ? ' With acquisition' : ' Without acquisition'))),
      compareBtn, prog.el),
    results, bridgeBox);
  const note = h('p', { id: 'pl-stale', class: 'err small', hidden: true }, 'Your inputs have changed since "Current inputs" was compared. Compare again to refresh.');
  results.before(note);

  function items() {
    const out = [];
    for (const k of st.selected) {
      if (['Base', 'Upside', 'Downside'].includes(k)) out.push({ key: k, label: k, values: T.valuesFor(ctx.model(), { preset: k }) });
      else if (k === '__current') out.push({ key: k, label: 'Current inputs', values: ctx.model().currentValues() });
      else { const s = st.saved.find(x => 'S:' + x.name === k); if (s) out.push({ key: k, label: s.name, values: T.valuesFor(ctx.model(), { changes: s.changes }) }); }
    }
    return out;
  }
  function options() {
    return [['Base', 'Base'], ['Upside', 'Upside'], ['Downside', 'Downside'], ...st.saved.map(s => ['S:' + s.name, s.name]), ['__current', 'Current inputs (live)']];
  }
  function renderPicker() {
    picker.replaceChildren(h('legend', {}, 'Compare 2 to 4 scenarios'));
    const n = st.selected.size;
    for (const [k, label] of options()) {
      picker.append(h('label', { class: 'chk' }, h('input', { type: 'checkbox', value: k, checked: st.selected.has(k), disabled: !st.selected.has(k) && n >= 4,
        onchange: e => { e.target.checked ? st.selected.add(k) : st.selected.delete(k); renderPicker(); } }), ' ' + label));
    }
    compareBtn.disabled = st.selected.size < 2;
  }
  function renderSaved() {
    const m = ctx.model();
    savedBox.replaceChildren(st.saved.length ? h('ul', { class: 'saved' }, st.saved.map(s => h('li', {},
      h('strong', {}, s.name), ` · ${Object.keys(s.changes).length} change${Object.keys(s.changes).length === 1 ? '' : 's'} from Base `,
      h('button', { class: 'btn small', onclick: () => load(s) }, 'Load into the control panel'), ' ',
      h('button', { class: 'btn small', onclick: () => remove(s) }, 'Delete')))) : h('p', { class: 'muted small' }, canSave ? 'No saved scenarios yet.' : ''));
  }
  function save() {
    const name = nameBox.value.trim();
    if (!name) { msg.textContent = 'Give the scenario a name first.'; return; }
    if (st.saved.some(s => s.name === name)) { msg.textContent = `A scenario called "${name}" already exists.`; return; }
    if (st.saved.length >= MAX_SAVED) { msg.textContent = `You can keep ${MAX_SAVED} saved scenarios. Delete one first.`; return; }
    const changes = T.diffFromBase(ctx.model(), ctx.model().currentValues());
    st.saved.push({ name, changes });
    if (!persist(st.saved)) { st.saved.pop(); msg.textContent = 'Could not save: storage is unavailable.'; return; }
    nameBox.value = '';
    msg.textContent = `Saved "${name}" with ${Object.keys(changes).length} change(s) from Base.`;
    renderSaved(); renderPicker();
  }
  function remove(s) {
    st.saved = st.saved.filter(x => x !== s);
    st.selected.delete('S:' + s.name);
    persist(st.saved);
    renderSaved(); renderPicker();
  }
  function load(s) {
    const m = ctx.model();
    m.resetAll();
    for (const [n, v] of Object.entries(s.changes)) m.setInput(n, v);
    msg.textContent = `Loaded "${s.name}" into the control panel.`;
    ctx.rerender();
  }

  async function runCompare() {
    const its = items();
    if (its.length < 2) return;
    compareBtn.disabled = true;
    prog.start('Preparing the what-if engine…');
    await T.tick();
    const include = st.overlay === 'without' ? 0 : null;
    const res = await T.evaluateItems(ctx.model(), its, { include, onProgress: (d, t) => prog.update(d, t) });
    prog.done(`Compared ${its.length} scenarios`);
    st.results = { its, res, include };
    st.stamp = JSON.stringify(ctx.model().currentValues());
    renderResults();
    compareBtn.disabled = false;
  }

  function renderResults() {
    const { its, res } = st.results;
    const heads = ['Metric', ...its.flatMap(i => [`${i.label}: Year 1`, `${i.label}: Year 3`])];
    const trs = ROWS.map(([label, a, b, kind]) => {
      const cells = [label];
      res.forEach(r => {
        if (kind === 'mixed') { cells.push(fmtNo(r[a], 'money'), fmtNo(r[b], 'ok')); }
        else cells.push(fmtNo(r[a], kind), fmtNo(r[b], kind));
      });
      return row(cells);
    });
    const flagItems = its.map((i, k) => ({ i, f: T.flags(res[k]) }));
    const arrSeries = its.flatMap((i, k) => [{ label: i.label, data: res[k].arr_series }, ...(st.results.include === 0 ? [] : [{ label: `${i.label} (pro-forma)`, data: res[k].pf_arr_series, dash: [6, 4] }])]);
    const cashSeries = its.flatMap((i, k) => [{ label: i.label, data: res[k].cash_series }, ...(st.results.include === 0 ? [] : [{ label: `${i.label} (pro-forma)`, data: res[k].pf_cash_series, dash: [6, 4] }])]);
    const diff = T.inputDiff(ctx.model(), its);
    const fmtIn = (r, v) => isPercentUnit(r.unit) ? `${inputDisplay(v, r.unit)}%` : `${fmtNum(v, 2)}`;
    results.replaceChildren(
      h('h3', {}, 'Headline outputs'),
      table(heads, trs, { caption: st.results.include === 0 ? 'Overlay: acquisition switched OFF in every scenario (pro-forma equals standalone).' : 'Overlay: acquisition ON. Year 3 comes from the five-year view; runway and CAC payback are monthly metrics, shown at months 12 and 24.', cls: 'wide' }),
      h('button', { class: 'btn small', id: 'pl-csv', onclick: () => downloadText('keel_scenario_comparison.csv', T.toCsv([heads, ...trs.map(r => r.cells)])) }, 'Download comparison (.csv)'),
      h('h3', {}, 'Downside flags'),
      h('ul', { class: 'flags' }, flagItems.map(({ i, f }) => h('li', { class: f.length ? 'bad' : '' }, h('strong', {}, i.label), f.length ? ': ⚠ ' + f.join('; ') : ': ✓ no flags (runway ≥ 12 months, NRR ≥ 100%, CAC payback ≤ 24 months, all checks pass)'))),
      chartFigure('c-cmp-arr', 'Closing ARR, by month'), chartFigure('c-cmp-cash', 'Cash, by month'),
      h('h3', {}, 'Input diff: which drivers differ'),
      diff.length ? table(['Driver', ...its.map(i => i.label)], diff.map(d => row([d.label, ...d.values.map(v => fmtIn(d, v))])), { cls: 'wide' }) : h('p', {}, 'These scenarios use identical inputs.'));
    const labels = ctx.monthLabels();
    ctx.charts().line($('#c-cmp-arr'), labels, arrSeries);
    ctx.charts().line($('#c-cmp-cash'), labels, cashSeries);
    for (const [id, ser] of [['c-cmp-arr', arrSeries], ['c-cmp-cash', cashSeries]]) $('#' + id + '-sum').textContent = ser.filter(s => !/pro-forma/.test(s.label)).map(s => `${s.label}: ${fmtMoney(s.data[0])} in month 1 to ${fmtMoney(s.data[23])} in month 24`).join('; ') + '.';
    renderBridgeControls();
    note.hidden = true;
  }
  function chartFigure(id, title) {
    return h('figure', { class: 'chart' }, h('figcaption', {}, title), h('div', { class: 'canvas-wrap' }, h('canvas', { id, role: 'img', 'aria-label': title, 'aria-describedby': id + '-sum' })), h('p', { id: id + '-sum', class: 'summary' }, ''));
  }

  function renderBridgeControls() {
    const targets = st.results.its.filter(i => i.key !== 'Base');
    const sel = h('select', { id: 'br-target', 'aria-label': 'Scenario to bridge to' }, targets.map(i => h('option', { value: i.key }, i.label)));
    const out = h('select', { id: 'br-output', 'aria-label': 'Output to bridge' }, ['arr24', 'ebitda_fy1', 'cash24'].map(id => h('option', { value: id }, T.OUTPUT[id].label)));
    const go = h('button', { class: 'btn', id: 'br-run', onclick: () => runBridge(sel.value, out.value) }, 'Build bridge');
    bridgeBox.replaceChildren(h('h3', {}, 'Scenario bridge: Base to a scenario'),
      h('div', { class: 'row-ctl' }, h('label', {}, 'Scenario ', sel), h('label', {}, 'Output ', out), go),
      h('div', { id: 'br-out' }));
    if (!targets.length) { go.disabled = true; $('#br-out').textContent = 'Select a scenario other than Base to bridge to.'; }
  }
  async function runBridge(key, outputId) {
    const it = st.results.its.find(i => i.key === key);
    const base = T.valuesFor(ctx.model(), { preset: 'Base' });
    prog.start('Building the bridge…');
    const br = await T.scenarioBridge(ctx.model(), base, it.values, outputId, { onProgress: (d, t) => prog.update(d, t) });
    prog.done(`Bridge built: ${br.steps.length} changed drivers`);
    let steps = br.steps.slice();
    let other = null;
    if (steps.length > 12) {
      const top = steps.slice().sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta)).slice(0, 10);
      const rest = steps.filter(s => !top.includes(s));
      other = { label: `${rest.length} smaller drivers`, delta: rest.reduce((a, s) => a + s.delta, 0) };
      steps = [...steps.filter(s => top.includes(s)), other];
    }
    let cum = br.base;
    const wf = [{ label: 'Base', from: 0, to: br.base, kind: 'total' }];
    for (const s of steps) { wf.push({ label: s.label, from: cum, to: cum + s.delta }); cum += s.delta; }
    wf.push({ label: it.label, from: 0, to: br.final, kind: 'total' });
    const box = $('#br-out');
    box.replaceChildren(
      h('figure', { class: 'chart' }, h('figcaption', {}, `${T.OUTPUT[outputId].label}: Base to ${it.label}`), h('div', { class: 'canvas-wrap' }, h('canvas', { id: 'c-bridge', role: 'img', 'aria-label': 'Scenario bridge', 'aria-describedby': 'c-bridge-sum' })), h('p', { id: 'c-bridge-sum', class: 'summary' }, '')),
      h('p', { class: 'note' }, 'The order in which drivers are applied affects how the total is split between them: each contribution is measured after the drivers above it have already been changed. The total is exact.'),
      table(['Driver (in order applied)', 'Base', it.label, 'Contribution'], br.steps.map(s => { const r = ctx.model().registry.find(x => x.name === s.name); const f = v => isPercentUnit(r.unit) ? `${inputDisplay(v, r.unit)}%` : fmtNum(v, 2); return row([s.label, f(s.from), f(s.to), `${s.delta >= 0 ? '+' : ''}${fmtNum(s.delta, 0)}`]); }), { cls: 'wide' }),
      h('button', { class: 'btn small', onclick: () => downloadText('keel_scenario_bridge.csv', T.toCsv([['driver', 'base', 'scenario', 'contribution'], ...br.steps.map(s => [s.label, s.from, s.to, s.delta]), ['Total change', '', '', br.final - br.base]])) }, 'Download bridge (.csv)'));
    ctx.charts().waterfall($('#c-bridge'), wf, { truncate: true });
    $('#c-bridge-sum').textContent = `${T.OUTPUT[outputId].label} moves from ${fmtMoney(br.base)} (Base) to ${fmtMoney(br.final)} (${it.label}); largest contributors: ${br.steps.slice().sort((a, b) => Math.abs(b.delta) - Math.abs(a.delta)).slice(0, 3).map(s => `${s.label} ${s.delta >= 0 ? '+' : ''}${fmtMoney(s.delta)}`).join(', ')}.`;
    st.bridge = br;
  }

  ctx.onChange(() => {
    if (st.results && st.results.its.some(i => i.key === '__current') && JSON.stringify(ctx.model().currentValues()) !== st.stamp) note.hidden = false;
    const live = $('#pl-live');
  });
  renderSaved();
  renderPicker();
  return { state: st };
}
