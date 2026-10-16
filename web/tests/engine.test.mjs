// Engine test: HyperFormula (the browser engine) must reproduce the workbook's own numbers.
// Reference values come from the independent pycel recalculation (validation/export_fixtures.py).
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
globalThis.HyperFormula = require('hyperformula').HyperFormula;
globalThis.XLSX = require('xlsx');
const { loadModel } = await import('../engine.js');

const here = fileURLToPath(new URL('.', import.meta.url));
const bytes = readFileSync(here + '../model/keel_saas_model.xlsx');
const names = JSON.parse(readFileSync(here + '../model/named_ranges.json', 'utf8'));
let failures = 0;
const ok = (cond, msg) => { console.log(`${cond ? 'PASS' : 'FAIL'}  ${msg}`); if (!cond) failures++; };

const model = await loadModel(new Uint8Array(bytes), names);
console.log(`workbook loaded into HyperFormula in ${model.buildMs.toFixed(0)} ms`);
ok(model.buildMs < 8000, `build time ${model.buildMs.toFixed(0)} ms (target: well under the 3 s first-load budget on a laptop; Node is not a browser)`);
ok(Object.keys(names).length > 340 && model.addr.out_input_registry !== undefined, `${Object.keys(names).length} named ranges registered`);
ok(model.registry.length === 99, `${model.registry.length} inputs in the registry`);

function compare(scen) {
  const ref = JSON.parse(readFileSync(here + `fixtures/outputs_${scen.toLowerCase()}.json`, 'utf8'));
  let bad = [], n = 0;
  for (const [coord, want] of Object.entries(ref)) {
    n++;
    const got = model.cell('Outputs_Web', coord);
    const w = want === null || want === undefined ? '' : want;
    const g = got === null || got === undefined ? '' : got;
    if (typeof w === 'number' && typeof g === 'number') {
      if (Math.abs(w - g) > 1e-6 * Math.max(1, Math.abs(w))) bad.push(`${coord}: pycel ${w} vs HF ${g}`);
    } else if (typeof w === 'number' && g === '' && Math.abs(w) < 1e-9) {
      // empty vs zero: fine
    } else if (String(w) !== String(g)) {
      bad.push(`${coord}: pycel ${JSON.stringify(w)} vs HF ${JSON.stringify(g)}`);
    }
  }
  return { n, bad };
}

for (const scen of ['Base', 'Upside', 'Downside']) {
  model.setScenario(scen);
  const { n, bad } = compare(scen);
  ok(bad.length === 0, `${scen}: all ${n} Outputs_Web cells match the independent recalculation` + (bad.length ? `; ${bad.length} differ, e.g. ${bad.slice(0, 5).join(' | ')}` : ''));
  console.log(`      scenario switch recalculated in ${model.lastRecalcMs.toFixed(0)} ms`);
}

// ten named outputs against the expected-values table (acceptance test 5.9 #1, against the independent engine)
const csv = readFileSync(here + '../../tests/expected_values.csv', 'utf8').trim().split('\n').slice(1).map(l => l.split(','));
const pick = ['rev_fy1', 'ebitda_fy1', 'arr24', 'nrr12', 'cash24', 'mincash', 'ma_value', 'pf_cash24', 'payback12', 'ltvcac12'];
for (const [i, scen] of ['Base', 'Upside', 'Downside'].entries()) {
  model.setScenario(scen);
  const bad = pick.filter(k => {
    const row = csv.find(r => r[0] === k);
    const want = parseFloat(row[1 + i]);
    const got = model.get('out_' + k);
    return Math.abs(got - want) > 1e-3 * Math.max(1, Math.abs(want)) * (Math.abs(want) < 5 ? 0.1 : 1);
  });
  ok(bad.length === 0, `${scen}: 10 named outputs match expected_values.csv` + (bad.length ? ` (differ: ${bad})` : ''));
}

// master check semantics
model.setScenario('Base');
ok(model.get('out_master_check') === 1, 'Base: master check = 1');
model.setScenario('Downside');
ok(model.get('out_master_check') === 0 && model.failingChecks().some(c => c.name === 'Pro-forma cash never negative'), 'Downside: master check = 0, failing check is "Pro-forma cash never negative"');
model.setScenario('Base');

// input validation: out-of-range values never reach the engine
const before = model.get('out_arr24');
const bad1 = model.setInput('inp_churn_ent', 0.9);
ok(!bad1.ok && model.get('out_arr24') === before && model.scenario === 'Base', `churn 90% blocked: "${bad1.message}"`);
ok(!model.setInput('inp_churn_ent', 'abc').ok, 'non-numeric input blocked');

// writing a scenario-driven input switches to Custom, starting from the current scenario
model.setScenario('Downside');
const r1 = model.setInput('inp_churn_ent', 0.07);
ok(r1.ok && model.scenario === 'Custom', 'editing a scenario-driven input switches to Custom');
ok(Math.abs(model.get('inp_logos_ent') - 2.2) < 1e-12, 'Custom starts from the previously selected scenario (Downside logos kept)');
ok(model.changedInputs().some(c => c.name === 'inp_churn_ent'), 'changed-input marker sees the edit');
console.log(`      single input change recalculated in ${model.lastRecalcMs.toFixed(0)} ms (target < 300 ms)`);
ok(model.lastRecalcMs < 300, `single-input recalculation ${model.lastRecalcMs.toFixed(0)} ms`);

// reset
model.resetAll();
ok(model.scenario === 'Base' && model.changedInputs().length === 0, 'reset all returns to Base with no changed inputs');
ok(Math.abs(model.get('out_arr24') - before) < 1e-6, 'after reset the outputs equal the original Base outputs');

// static input edit
const gw0 = model.get('out_goodwill_ifrs');
ok(model.setInput('inp_ma_offer_mult', 5.5).ok && model.get('out_goodwill_ifrs') > gw0, 'a static input (offer multiple) flows through to goodwill');
model.resetAll();

// trace
const src = model.source('out_arr24');
ok(src && src.sheet === 'SaaS_Metrics', `trace: out_arr24 comes from ${src && src.sheet}!${src && src.a1}`);
const pre = model.precedents(src.sheet, src.a1);
ok(pre.refs.length > 0 && pre.formula.startsWith('='), `trace: one level of precedents (${pre.refs.length}) from ${pre.formula}`);

console.log(failures ? `\n${failures} FAILED` : '\nall engine tests passed');
process.exit(failures ? 1 : 0);
