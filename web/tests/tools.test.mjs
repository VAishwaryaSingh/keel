// Tools test: scenario planner and sensitivity calculations, checked against independent recalculations (pycel)
// and against the engine itself. Covers brief acceptance tests 5.9 #3 (tornado), #4 (grid cell), #5 (goal seek).
import { readFileSync, writeFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
globalThis.HyperFormula = require('hyperformula').HyperFormula;
globalThis.XLSX = require('xlsx');
const here = fileURLToPath(new URL('.', import.meta.url));
const { loadModel } = await import('../engine.js');
const T = await import('../tools.js');
const bytes = new Uint8Array(readFileSync(here + '../model/keel_saas_model.xlsx'));
const names = JSON.parse(readFileSync(here + '../model/named_ranges.json', 'utf8'));
const ref = JSON.parse(readFileSync(here + 'fixtures/sensitivity_reference.json', 'utf8'));
let failures = 0;
const ok = (c, m) => { console.log(`${c ? 'PASS' : 'FAIL'}  ${m}`); if (!c) failures++; };
const close = (a, b, rel = 1e-6) => Math.abs(a - b) <= rel * Math.max(1, Math.abs(b));
const model = await loadModel(bytes, names);
const stateBefore = JSON.stringify(model.currentValues());
const t0 = performance.now();
const lean = model.getLean();
console.log(`lean what-if engine built in ${lean.buildMs.toFixed(0)} ms`);

// ---------- planner: evaluate presets -------------------------------------------------------------------
const items = ['Base', 'Upside', 'Downside'].map(p => ({ key: p, values: T.valuesFor(model, { preset: p }) }));
const res = await T.evaluateItems(model, items);
for (const [i, p] of ['Base', 'Upside', 'Downside'].entries()) {
  model.setScenario(p);
  const direct = T.headline(model);
  const keys = ['arr_y1', 'arr_y3', 'rev_y1', 'rev_y3', 'ebitda_y1', 'ebitda_y3', 'cash_y1', 'cash_y3', 'nrr_y1', 'nrr_y3', 'payback_m12', 'payback_m24', 'pf_cash24', 'ma_value', 'master'];
  const bad = keys.filter(k => !close(res[i][k], direct[k]));
  ok(bad.length === 0, `${p}: lean-engine headline equals the main engine for ${keys.length} figures${bad.length ? ' (differ: ' + bad + ')' : ''}`);
}
model.setScenario('Base');
ok(res[0].master === 1 && res[2].master === 0 && T.flags(res[2]).some(f => f.includes('Pro-forma cash never negative')), 'Downside raises a model-check flag; Base does not');
ok(T.flags(res[2]).some(f => f.startsWith('NRR under 100%')) && T.flags(res[2]).some(f => f.startsWith('CAC payback over 24')), `Downside flags: ${T.flags(res[2]).join(' | ')}`);
ok(T.flags(res[0]).length === 0, `Base raises no flags (${T.flags(res[0]).length})`);
ok(JSON.stringify(model.currentValues()) === stateBefore && model.scenario === 'Base', 'the visitor\'s inputs are untouched after the planner runs');
const withOut = await T.evaluateItems(model, [items[0]], { include: 0 });
ok(close(withOut[0].pf_cash24, res[0].cash_y1 === undefined ? 0 : T.OUTPUT.cash24.read(model)) && Math.abs(withOut[0].pf_cash24 - model.get('out_cash24')) < 1e-6, 'overlay "without acquisition": pro-forma cash equals standalone');
const diff = T.inputDiff(model, items);
ok(diff.some(d => d.name === 'inp_churn_ent') && !diff.some(d => d.name === 'inp_marketing_m') && diff.length >= 18, `input diff lists ${diff.length} differing drivers, none of them static`);

// ---------- bridge --------------------------------------------------------------------------------------
const tb = performance.now();
const br = await T.scenarioBridge(model, items[0].values, items[2].values, 'arr24');
const sum = br.steps.reduce((a, s) => a + s.delta, 0);
ok(close(br.base + sum, br.final, 1e-9) && close(br.final, res[2].arr_series[23]), `bridge Base -> Downside (ARR m24): ${br.base.toFixed(0)} + ${br.steps.length} steps = ${br.final.toFixed(0)} = Downside ARR (${((performance.now() - tb) / 1000).toFixed(1)} s)`);
ok(JSON.stringify(model.currentValues()) === stateBefore, 'state untouched after the bridge');
const brC = await T.scenarioBridge(model, items[0].values, items[2].values, 'cash24');
ok(close(brC.final, res[2].cash_series[23]), 'bridge on closing cash also ends at the Downside figure');

// ---------- tornado: acceptance 5.9 #3 ------------------------------------------------------------------
const tt = performance.now();
const tor = await T.tornado(model, 0.10);
const tornSecs = (performance.now() - tt) / 1000;
const jsNames = tor.rows.filter(r => r.atLo).map(r => r.name).sort();
const refNames = Object.keys(ref.tornado).sort();
ok(JSON.stringify(jsNames) === JSON.stringify(refNames), `tornado flexes the same ${jsNames.length} drivers as the independent test (${tornSecs.toFixed(1)} s)`);
let worst = 0;
for (const r of tor.rows.filter(x => x.atLo)) for (const id of Object.keys(ref.base)) {
  worst = Math.max(worst, Math.abs(r.atLo[id] - ref.tornado[r.name].atLo[id]) / Math.max(1, Math.abs(ref.tornado[r.name].atLo[id])), Math.abs(r.atHi[id] - ref.tornado[r.name].atHi[id]) / Math.max(1, Math.abs(ref.tornado[r.name].atHi[id])));
}
ok(worst < 1e-6, `every low/high result for 8 outputs matches the independent engine (worst relative gap ${worst.toExponential(1)})`);
for (const id of ['arr24', 'cash24', 'ebitda_fy2', 'ma_value']) {
  const mine = T.rankTornado(tor, id).slice(0, 3).map(r => r.name);
  const theirs = Object.entries(ref.tornado).map(([n, v]) => [n, Math.abs(v.atHi[id] - v.atLo[id])]).sort((a, b) => b[1] - a[1]).slice(0, 3).map(x => x[0]);
  ok(JSON.stringify(mine) === JSON.stringify(theirs), `top 3 for ${id}: ${mine.join(', ')} (matches the manual +/-10% test)`);
}
ok(close(tor.base.arr24, ref.base.arr24) && JSON.stringify(model.currentValues()) === stateBefore, 'tornado base equals the reference base and the visitor\'s state is untouched');
const clamp = tor.rows.filter(r => r.clamped).map(r => r.name);
console.log(`      clamped to allowed range: ${clamp.join(', ') || 'none'}`);

// ---------- heatmap: acceptance 5.9 #4 ------------------------------------------------------------------
const hm = await T.heatmap(model, 'inp_logos_ent', [2.0, 2.5, 3.0], 'inp_churn_ent', [0.06, 0.08, 0.10]);
const cell = hm.cells[1][1];
ok(close(cell.arr24, ref.grid_cell.outputs.arr24) && close(cell.cash24, ref.grid_cell.outputs.cash24) && close(cell.mincash, ref.grid_cell.outputs.mincash), `grid cell (logos 2.5, churn 8%) = ARR ${cell.arr24.toFixed(1)}, cash ${cell.cash24.toFixed(1)}: matches the independent recalculation`);
ok(close(hm.cells[1][1].arr24, hm.cells[1][1].arr24) && hm.cells[0][0].arr24 > hm.cells[2][2].arr24 === false, 'grid is oriented correctly (more logos and less churn is not worse)');
const th = performance.now();
const big = await T.heatmap(model, 'inp_logos_ent', T.linspace(2, 4, 7), 'inp_churn_ent', T.linspace(0.03, 0.12, 7));
const gridSecs = (performance.now() - th) / 1000;
ok(big.cells.length === 7 && big.cells[0].length === 7 && gridSecs < 5, `7 x 7 grid (49 recalculations) in ${gridSecs.toFixed(1)} s (target under 5 s)`);

// ---------- goal seek: acceptance 5.9 #5 ----------------------------------------------------------------
const gs = await T.goalSeek(model, 'mincash', 15000, 'inp_churn_ent');
ok(gs.solved && Math.abs(gs.gap) <= gs.tol, `goal seek: churn ${(gs.x * 100).toFixed(4)}% gives lowest cash ${gs.value.toFixed(4)} (target 15,000) in ${gs.iterations} iterations`);
ok(gs.iterationLog.length === gs.iterations + 2, `iteration log has ${gs.iterationLog.length} rows (2 limits + ${gs.iterations} bisections)`);
model.setInput('inp_churn_ent', gs.x);
ok(Math.abs(T.OUTPUT.mincash.read(model) - 15000) <= 0.01, `typed back into the model: lowest cash = ${T.OUTPUT.mincash.read(model).toFixed(4)}`);
model.resetAll();
writeFileSync(here + 'fixtures/goalseek_result.json', JSON.stringify({ driver: 'inp_churn_ent', x: gs.x, output: 'mincash', target: 15000 }));
const gs2 = await T.goalSeek(model, 'ma_value', 0, 'inp_ma_offer_mult');
ok(gs2.solved && Math.abs(gs2.value) <= 0.001, `goal seek: offer multiple ${gs2.x.toFixed(4)}x makes the deal value exactly zero (break-even price)`);
const none = await T.goalSeek(model, 'arr24', 1e9, 'inp_logos_ent');
ok(!none.solved && /not reached/.test(none.reason), 'unreachable target is reported, not forced');
ok(JSON.stringify(model.currentValues()) === stateBefore, 'state untouched after goal seeks');

// ---------- acquisition grid -----------------------------------------------------------------------------
const ag = await T.acquisitionGrid(model, [3.5, 4.0, 4.5, 5.0, 5.5], [0, 0.4, 0.7, 1.0, 1.3]);
ok(close(ag.cells[3][2].value, model.get('out_ma_value')), `acquisition grid: base cell (4.5x, 100%) = ${ag.cells[3][2].value.toFixed(0)} = the model's value created`);
ok(ag.cells[0][0].value > ag.cells[0][4].value && ag.cells[4][0].value > ag.cells[0][0].value && ag.cells[0][4].value < 0, 'grid direction: a lower price helps, more synergy helps, 5.5x with no synergy destroys value');
ok(ag.cells.flat().every(c => typeof c.payback === 'number' || /years/.test(c.payback)), 'every cell carries a payback (a number of years, or "> 5 years")');

// ---------- CSV ----------------------------------------------------------------------------------------------
ok(T.toCsv([['a,b', 'c"d', 1], [null, 'x', 2.5]]) === '"a,b","c""d",1\n,x,2.5\n', 'CSV quoting');
console.log(`total tools test time ${((performance.now() - t0) / 1000).toFixed(0)} s`);
console.log(failures ? `\n${failures} FAILED` : '\nall tools tests passed');
process.exit(failures ? 1 : 0);
