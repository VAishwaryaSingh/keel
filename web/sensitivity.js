// sensitivity.js: the sensitivity tool (section 10): tornado, two-way heatmap, goal seek, acquisition grid.
// Every number is produced by recalculating the workbook (tools.js); this file lays the results out.
import { h, $, table, row, progressBar, downloadText } from './dom.js';
import { fmtMoney, fmtNum, fmtPct, isNum, isPercentUnit, inputDisplay, inputParse } from './format.js';
import * as T from './tools.js';

const fmtOut = (id, v) => {
  if (!isNum(v)) return v === null || v === undefined ? 'n/a' : String(v);
  const o = T.OUTPUT[id];
  if (o.kind === 'months') return v >= T.NO_LIMIT ? 'no limit' : `${fmtNum(v, 1)} mo`;
  return fmtMoney(v);
};
const fmtIn = (r, v) => isPercentUnit(r.unit) ? `${inputDisplay(v, r.unit)}%` : fmtNum(v, Math.abs(v) < 10 ? 2 : 0);
const stamp = ctx => JSON.stringify(ctx.model().currentValues());

export function mountSensitivity(root, ctx) {
  const reg = () => ctx.model().registry.filter(r => r.unit !== '1/0' && r.min !== null);
  const st = { tornado: null, heat: null, goal: null, acq: null };
  const prog = progressBar();
  const stale = id => h('p', { id, class: 'err small', hidden: true }, 'Your inputs have changed since this was calculated. Run it again to refresh.');
  const outSelect = (id, skip = []) => h('select', { id, 'aria-label': 'Output' }, T.OUTPUTS.filter(o => !skip.includes(o.id)).map(o => h('option', { value: o.id }, o.label)));
  const driverSelect = (id, value) => h('select', { id, 'aria-label': 'Driver' }, reg().map(r => h('option', { value: r.name, selected: r.name === value }, r.label)));
  const busy = (btn, on) => { btn.disabled = on; };

  // ---------------- tornado ----------------
  const tnOut = outSelect('tn-output', ['runway24']);
  const tnPct = h('input', { type: 'number', id: 'tn-pct', value: 10, min: 1, max: 50, step: 1, 'aria-label': 'Flex percentage' });
  const tnRun = h('button', { class: 'btn primary', id: 'tn-run', onclick: runTornado }, 'Run tornado');
  const tnErr = h('span', { class: 'err', role: 'alert' }, '');
  const tnBox = h('div', { id: 'tn-box' });
  const tnStale = stale('tn-stale');
  async function runTornado() {
    const pct = Number(tnPct.value);
    if (!(pct >= 1 && pct <= 50)) { tnErr.textContent = 'Enter a flex between 1 and 50.'; return; }
    tnErr.textContent = '';
    busy(tnRun, true);
    prog.start('Preparing the what-if engine…');
    await T.tick();
    st.tornado = await T.tornado(ctx.model(), pct / 100, { onProgress: (d, t) => prog.update(d, t) });
    st.tornado.stamp = stamp(ctx);
    prog.done(`Flexed ${st.tornado.rows.filter(r => r.atLo).length} drivers by ±${pct}%`);
    renderTornado();
    tnStale.hidden = true;
    busy(tnRun, false);
  }
  function renderTornado() {
    const t = st.tornado;
    if (!t) return;
    const id = tnOut.value;
    const ranked = T.rankTornado(t, id);
    const base = t.base[id];
    const top = ranked.slice(0, 15);
    const skipped = t.rows.filter(r => r.skipped);
    tnBox.replaceChildren(
      h('figure', { class: 'chart' }, h('figcaption', {}, `${T.OUTPUT[id].label}: effect of flexing each driver ±${(t.pct * 100).toFixed(0)}% (top ${top.length}; base ${fmtOut(id, base)})`),
        h('div', { class: 'canvas-wrap tall' }, h('canvas', { id: 'c-tornado', role: 'img', 'aria-label': 'Tornado chart', 'aria-describedby': 'c-tornado-sum' })), h('p', { id: 'c-tornado-sum', class: 'summary' }, '')),
      table(['Rank', 'Driver', 'Flexed from → to', `${T.OUTPUT[id].label} at low`, 'at high', 'Swing'],
        ranked.map((r, i) => row([String(i + 1), r.label, `${fmtIn(r, r.lo)} → ${fmtIn(r, r.hi)}${r.clamped ? ' (limited by allowed range)' : ''}`, fmtOut(id, r.outLo), fmtOut(id, r.outHi), fmtOut(id, r.swing)])), { cls: 'wide' }),
      skipped.length ? h('p', { class: 'muted small' }, `Not flexed (base value is 0): ${skipped.map(r => r.label).join(', ')}.`) : null,
      h('button', { class: 'btn small', id: 'tn-csv', onclick: () => downloadText('keel_tornado.csv', T.toCsv([['rank', 'driver', 'input', 'base', 'low', 'high', `${T.OUTPUT[id].label} at low`, `${T.OUTPUT[id].label} at high`, 'swing'], ...ranked.map((r, i) => [i + 1, r.label, r.name, r.base, r.lo, r.hi, r.outLo, r.outHi, r.swing])])) }, 'Download tornado table (.csv)'));
    ctx.charts().range($('#c-tornado'), top.map(r => ({ label: r.label, low: Math.min(r.outLo, r.outHi), high: Math.max(r.outLo, r.outHi) })).reverse(), { axisTitle: `${T.OUTPUT[id].label} (GBP k unless stated). Base ${fmtOut(id, base)}`, name: 'Output range' });
    $('#c-tornado-sum').textContent = `Biggest swings for ${T.OUTPUT[id].label}: ${top.slice(0, 3).map(r => `${r.label} (${fmtOut(id, r.outLo)} to ${fmtOut(id, r.outHi)})`).join('; ')}.`;
  }
  tnOut.addEventListener('change', renderTornado);

  // ---------------- heatmap ----------------
  const hmX = driverSelect('hm-x', 'inp_logos_ent'), hmY = driverSelect('hm-y', 'inp_churn_ent');
  const hmOut = outSelect('hm-output');
  const num = (id, label) => h('input', { type: 'number', id, step: 'any', 'aria-label': label });
  const hx0 = num('hm-x0', 'X from'), hx1 = num('hm-x1', 'X to'), hy0 = num('hm-y0', 'Y from'), hy1 = num('hm-y1', 'Y to');
  const steps = (id, label, v) => h('select', { id, 'aria-label': label }, [2, 3, 4, 5, 6, 7].map(n => h('option', { value: n, selected: n === v }, String(n))));
  const hnx = steps('hm-nx', 'X steps', 5), hny = steps('hm-ny', 'Y steps', 5);
  const hmRun = h('button', { class: 'btn primary', id: 'hm-run', onclick: runHeat }, 'Run grid');
  const hmErr = h('span', { class: 'err', role: 'alert', id: 'hm-err' }, '');
  const hmBox = h('div', { id: 'hm-box' });
  const hmStale = stale('hm-stale');
  function defaults(sel, a, b) {
    const r = reg().find(x => x.name === sel.value);
    const v = ctx.model().value(r.name);
    let lo = Math.max(r.min, v * 0.8), hi = Math.min(r.max, v * 1.2);
    if (v === 0) { lo = r.min; hi = Math.min(r.max, r.min + (r.max - r.min) * 0.2); }
    a.value = inputDisplay(lo, r.unit);
    b.value = inputDisplay(hi, r.unit);
  }
  const resetDefaults = () => { defaults(hmX, hx0, hx1); defaults(hmY, hy0, hy1); };
  hmX.addEventListener('change', () => defaults(hmX, hx0, hx1));
  hmY.addEventListener('change', () => defaults(hmY, hy0, hy1));
  async function runHeat() {
    const rx = reg().find(r => r.name === hmX.value), ry = reg().find(r => r.name === hmY.value);
    if (rx.name === ry.name) { hmErr.textContent = 'Pick two different drivers.'; return; }
    const parse = (el, r) => inputParse(el.value, r.unit);
    const [x0, x1, y0, y1] = [parse(hx0, rx), parse(hx1, rx), parse(hy0, ry), parse(hy1, ry)];
    const bad = [[x0, rx, 'X from'], [x1, rx, 'X to'], [y0, ry, 'Y from'], [y1, ry, 'Y to']].find(([v, r]) => !Number.isFinite(v) || v < r.min - 1e-12 || v > r.max + 1e-12);
    if (bad) { hmErr.textContent = `${bad[2]} must be a number between ${fmtIn(bad[1], bad[1].min)} and ${fmtIn(bad[1], bad[1].max)}.`; return; }
    hmErr.textContent = '';
    busy(hmRun, true);
    prog.start('Preparing the what-if engine…');
    await T.tick();
    const nx = +hnx.value, ny = +hny.value;
    st.heat = await T.heatmap(ctx.model(), rx.name, T.linspace(x0, x1, nx), ry.name, T.linspace(y0, y1, ny), { onProgress: (d, t) => prog.update(d, t) });
    st.heat.rx = rx; st.heat.ry = ry; st.heat.stamp = stamp(ctx); st.heat.cur = [ctx.model().value(rx.name), ctx.model().value(ry.name)];
    prog.done(`Ran ${nx * ny} recalculations`);
    renderHeat();
    hmStale.hidden = true;
    busy(hmRun, false);
  }
  function scaleColour(t) {   // 0 = lowest value, 1 = highest: red to green via the theme colours, never the only signal (numbers stay visible)
    const p = Math.round(t * 100);
    return `color-mix(in srgb, var(--good) ${p}%, var(--bad) ${100 - p}%)`;
  }
  function renderHeat() {
    const g = st.heat;
    if (!g) return;
    const id = hmOut.value;
    const vals = g.cells.flat().map(c => c[id]).filter(isNum);
    const lo = Math.min(...vals), hi = Math.max(...vals);
    const near = (arr, v) => arr.reduce((b, x, i) => (Math.abs(x - v) < Math.abs(arr[b] - v) ? i : b), 0);
    const ci = near(g.xVals, g.cur[0]), cj = near(g.yVals, g.cur[1]);
    const trs = g.cells.map((rowOut, j) => row([fmtIn(g.ry, g.yVals[j])], '', {
      styles: rowOut.map((c, i) => `background:${isNum(c[id]) && hi > lo ? `color-mix(in srgb, ${scaleColour((c[id] - lo) / (hi - lo))} 38%, var(--surface))` : 'transparent'}`),
      tdCls: rowOut.map((c, i) => (i === ci && j === cj ? 'current' : '')).map((x, i) => x), }));
    trs.forEach((tr, j) => { tr.cells = [tr.cells[0], ...g.cells[j].map(c => fmtOut(id, c[id]))]; tr.styles = [null, ...tr.styles]; tr.tdCls = [null, ...tr.tdCls]; });
    hmBox.replaceChildren(
      h('p', { class: 'small' }, `${T.OUTPUT[id].label}. Rows: ${g.ry.label}. Columns: ${g.rx.label}. Colour runs from lowest (red) to highest (green); the numbers carry the meaning. The outlined cell is the closest to your current inputs (${fmtIn(g.rx, g.cur[0])}, ${fmtIn(g.ry, g.cur[1])}).`),
      table([`${g.ry.label} \\ ${g.rx.label}`, ...g.xVals.map(x => fmtIn(g.rx, x))], trs, { cls: 'hm', caption: `Two-way sensitivity of ${T.OUTPUT[id].label}` }),
      h('p', { class: 'summary' }, `Range across the grid: ${fmtOut(id, lo)} to ${fmtOut(id, hi)}.`),
      h('button', { class: 'btn small', id: 'hm-csv', onclick: () => downloadText('keel_sensitivity_grid.csv', T.toCsv([[`${g.ry.label} (rows) / ${g.rx.label} (columns)`, ...g.xVals.map(x => x)], ...g.cells.map((rowOut, j) => [g.yVals[j], ...rowOut.map(c => c[id])])])) }, 'Download grid (.csv)'));
  }
  hmOut.addEventListener('change', renderHeat);

  // ---------------- goal seek ----------------
  const gsOut = outSelect('gs-output');
  gsOut.value = 'mincash';
  const gsTarget = h('input', { type: 'number', id: 'gs-target', step: 'any', value: 0, 'aria-label': 'Target value' });
  const gsDriver = driverSelect('gs-driver', 'inp_churn_ent');
  const gsRun = h('button', { class: 'btn primary', id: 'gs-run', onclick: runGoal }, 'Solve');
  const gsBox = h('div', { id: 'gs-box' });
  const gsStale = stale('gs-stale');
  async function runGoal() {
    const target = Number(gsTarget.value);
    if (!Number.isFinite(target)) { gsBox.textContent = 'Enter a numeric target.'; return; }
    busy(gsRun, true);
    prog.start('Preparing the what-if engine…');
    await T.tick();
    const r = await T.goalSeek(ctx.model(), gsOut.value, target, gsDriver.value, { onProgress: (d, t) => prog.update(d, t) });
    r.stamp = stamp(ctx);
    st.goal = r;
    prog.done(r.solved ? `Solved in ${r.iterations} iterations` : 'No solution');
    renderGoal();
    gsStale.hidden = true;
    busy(gsRun, false);
  }
  function renderGoal() {
    const r = st.goal;
    if (!r) return;
    const o = T.OUTPUT[r.outputId];
    const log = table(['Iteration', `${r.driver.label}`, `${o.label}`, 'Gap to target'], r.iterationLog.map(it => row([it.note ? `${it.note}` : String(it.i), fmtIn(r.driver, it.x) + (isPercentUnit(r.driver.unit) ? '' : ''), fmtOut(r.outputId, it.value), fmtNum(it.gap, 4)])), { cls: 'wide', caption: 'Bisection: the interval is halved each step towards the target.' });
    gsBox.replaceChildren(
      r.solved ? h('p', { id: 'gs-answer', class: 'answer' }, h('strong', {}, `${r.driver.label} = ${isPercentUnit(r.driver.unit) ? inputDisplay(r.x, r.driver.unit).toFixed(4) + '%' : fmtNum(r.x, 4)}`),
        ` makes ${o.label} ${fmtOut(r.outputId, r.value)} (target ${fmtOut(r.outputId, r.target)}, gap ${fmtNum(r.gap, 4)}) after ${r.iterations} iterations. Other inputs stay as they are now.`)
        : h('p', { id: 'gs-answer', class: 'err' }, r.reason),
      r.solved ? h('button', { class: 'btn small', id: 'gs-apply', onclick: () => { ctx.model().setInput(r.driver.name, r.x); ctx.rerender(); } }, 'Apply this value to the inputs') : null,
      log, h('p', { class: 'muted small' }, 'Bisection finds a point where the output crosses the target between the input\'s limits. If the relationship is not monotonic there may be other solutions.'),
      h('button', { class: 'btn small', id: 'gs-csv', onclick: () => downloadText('keel_goal_seek.csv', T.toCsv([['iteration', r.driver.name, r.outputId, 'gap'], ...r.iterationLog.map(it => [it.note || it.i, it.x, it.value, it.gap])])) }, 'Download iterations (.csv)'));
  }

  // ---------------- acquisition grid ----------------
  const MULT = [3.5, 4.0, 4.5, 5.0, 5.5], ACH = [0, 0.4, 0.7, 1.0, 1.3];
  const agRun = h('button', { class: 'btn primary', id: 'ag-run', onclick: runAcq }, 'Run acquisition grid');
  const agBox = h('div', { id: 'ag-box' });
  const agStale = stale('ag-stale');
  async function runAcq() {
    busy(agRun, true);
    prog.start('Preparing the what-if engine…');
    await T.tick();
    st.acq = await T.acquisitionGrid(ctx.model(), MULT, ACH, { onProgress: (d, t) => prog.update(d, t) });
    st.acq.stamp = stamp(ctx);
    prog.done('Acquisition grid complete');
    renderAcq();
    agStale.hidden = true;
    busy(agRun, false);
  }
  function renderAcq() {
    const g = st.acq;
    if (!g) return;
    const heads = ['Synergies achieved \\ offer EV / ARR', ...g.multiples.map(m => `${m.toFixed(1)}x`)];
    const valRows = g.cells.map((r, j) => row([fmtPct(g.achieved[j], 0), ...r.map(c => (c.value >= 0 ? '' : '▼ ') + fmtMoney(c.value))], '', { tdCls: [null, ...r.map(c => (c.value < 0 ? 'neg' : 'pos'))] }));
    const payRows = g.cells.map((r, j) => row([fmtPct(g.achieved[j], 0), ...r.map(c => (isNum(c.payback) ? `${fmtNum(c.payback, 1)} yrs` : String(c.payback)))]));
    agBox.replaceChildren(h('h4', {}, 'Value created (NPV)'), table(heads, valRows, { cls: 'hm', caption: '▼ marks a deal that destroys value. Inputs other than the offer multiple and synergies achieved stay as they are now.' }),
      h('h4', {}, 'Payback (undiscounted, five years)'), table(heads, payRows, { cls: 'hm' }),
      h('button', { class: 'btn small', id: 'ag-csv', onclick: () => downloadText('keel_acquisition_grid.csv', T.toCsv([['achieved', 'offer_multiple', 'value_created', 'payback'], ...g.cells.flatMap((r, j) => r.map((c, i) => [g.achieved[j], g.multiples[i], c.value, c.payback]))])) }, 'Download acquisition grid (.csv)'));
  }

  // ---------------- layout ----------------
  const wrap = (title, intro, ...kids) => h('div', { class: 'tool' }, h('h3', {}, title), h('p', {}, intro), ...kids);
  root.append(
    h('p', {}, 'Each tool recalculates the workbook many times (a what-if copy of the model, so your inputs and the rest of the page are not disturbed). Results reflect your inputs at the moment you press the button.'),
    prog.el,
    wrap('Tornado: which assumption matters most?', 'Flexes each key driver up and down by the same percentage, one at a time, and ranks them by how far the chosen output moves.',
      h('div', { class: 'row-ctl' }, h('label', {}, 'Output ', tnOut), h('label', {}, 'Flex ± ', tnPct, ' %'), tnRun, tnErr), tnStale, tnBox),
    wrap('Two-way sensitivity grid', 'Pick any two drivers, their ranges and step counts (2 to 7 each, up to 49 recalculations) and an output.',
      h('div', { class: 'row-ctl' }, h('label', {}, 'Columns ', hmX), h('label', {}, 'from ', hx0), h('label', {}, 'to ', hx1), h('label', {}, 'steps ', hnx)),
      h('div', { class: 'row-ctl' }, h('label', {}, 'Rows ', hmY), h('label', {}, 'from ', hy0), h('label', {}, 'to ', hy1), h('label', {}, 'steps ', hny)),
      h('div', { class: 'row-ctl' }, h('label', {}, 'Output ', hmOut), hmRun, hmErr), hmStale, hmBox),
    wrap('Goal seek: what value of a driver reaches a target?', 'For example: what churn rate takes the lowest month-end cash down to a given level? Solved by bisection on the workbook, within the input\'s allowed range.',
      h('div', { class: 'row-ctl' }, h('label', {}, 'Output ', gsOut), h('label', {}, 'target ', gsTarget), h('label', {}, 'by changing ', gsDriver), gsRun), gsStale, gsBox),
    wrap('Acquisition sensitivity: offer price against synergies achieved', 'A preset grid: the offer multiple (price) across the columns and the share of planned synergies achieved down the rows.',
      h('div', { class: 'row-ctl' }, agRun), agStale, agBox));
  resetDefaults();
  ctx.onChange(() => {
    const now = stamp(ctx);
    for (const [k, el] of [['tornado', tnStale], ['heat', hmStale], ['goal', gsStale], ['acq', agStale]]) if (st[k] && st[k].stamp !== now) el.hidden = false;
  });
  return { state: st };
}
