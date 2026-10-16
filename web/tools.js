// tools.js: scenario and sensitivity calculations. No spreadsheet logic lives here: every figure is read from the
// workbook (through the lean what-if engine) after setting inputs. Pure functions over a Model, so they run in Node tests.
export const tick = () => new Promise(r => setTimeout(r, 0));
const num = v => (typeof v === 'number' && Number.isFinite(v) ? v : null);

// ---- outputs the tools can report on ------------------------------------------------------------------
export const OUTPUTS = [
  { id: 'arr24', label: 'ARR, month 24', kind: 'money', read: m => m.get('out_arr24') },
  { id: 'arr_y3', label: 'ARR, year 3 (five-year view)', kind: 'money', read: m => m.get('out_lt_arr_close')[2] },
  { id: 'rev_fy1', label: 'Revenue, FY1', kind: 'money', read: m => m.get('out_rev_fy1') },
  { id: 'ebitda_fy1', label: 'EBITDA, FY1', kind: 'money', read: m => m.get('out_ebitda_fy1') },
  { id: 'ebitda_fy2', label: 'EBITDA, FY2', kind: 'money', read: m => m.get('out_ebitda_fy2') },
  { id: 'cash24', label: 'Cash, month 24', kind: 'money', read: m => m.get('out_cash24') },
  { id: 'mincash', label: 'Lowest month-end cash', kind: 'money', read: m => m.get('out_mincash') },
  { id: 'ma_value', label: 'Acquisition: value created', kind: 'money', read: m => m.get('out_ma_value') },
  { id: 'pf_cash24', label: 'Pro-forma cash, month 24', kind: 'money', read: m => m.get('out_pf_cash24') },
  { id: 'runway24', label: 'Cash runway, month 24 (months; no limit when cash generative)', kind: 'months', read: m => runway(m, 23) },
];
export const OUTPUT = Object.fromEntries(OUTPUTS.map(o => [o.id, o]));
export const NO_LIMIT = 1e6;                       // runway when the business is not burning cash
export function runway(m, i) {
  const v = m.series('out_sm_runway').months[i];
  return num(v) !== null ? v : NO_LIMIT;
}
export function readAll(m) {
  const out = {};
  for (const o of OUTPUTS) { const v = o.read(m); out[o.id] = num(v) !== null ? v : null; }
  return out;
}

/** Headline figures for the scenario comparison. Year 1 = FY1 / month 12; year 3 from the five-year view. */
export function headline(m) {
  const s = n => m.series(n).months;
  const lt = n => m.get(n);
  const burn = s('out_sm_burn_mult');
  return {
    arr_y1: m.get('out_arr12'), arr_y3: lt('out_lt_arr_close')[2],
    rev_y1: m.get('out_rev_fy1'), rev_y3: lt('out_lt_rev')[2],
    ebitda_y1: m.get('out_ebitda_fy1'), ebitda_y3: lt('out_lt_ebitda')[2],
    cash_y1: s('out_bs_cash')[11], cash_y3: lt('out_lt_cash_close')[2],
    nrr_y1: s('out_sm_nrr')[11], nrr_y3: lt('out_lt_nrr')[2],
    burn_y1: burn[11], burn_y3: lt('out_lt_burn_mult')[2],
    runway_m12: runway(m, 11), runway_m24: runway(m, 23),
    payback_m12: s('out_sm_payback')[11], payback_m24: s('out_sm_payback')[23],
    pf_arr24: m.get('out_pf_arr24') ?? m.series('out_pf_pf_arr').months[23], pf_cash24: m.get('out_pf_cash24'), ma_value: m.get('out_ma_value'),
    master: m.get('out_master_check'), failing: m.failingChecks().map(f => f.name),
    arr_series: s('out_sm_arr_close'), cash_series: s('out_bs_cash'), pf_arr_series: m.series('out_pf_pf_arr').months, pf_cash_series: m.series('out_pf_pf_cash').months,
  };
}
export const THRESHOLDS = { runwayMonths: 12, nrr: 1.0, paybackMonths: 24 };
export function flags(h) {
  const f = [];
  if (h.runway_m24 < THRESHOLDS.runwayMonths) f.push(`Runway under ${THRESHOLDS.runwayMonths} months at month 24`);
  if (h.runway_m12 < THRESHOLDS.runwayMonths) f.push(`Runway under ${THRESHOLDS.runwayMonths} months at month 12`);
  if (num(h.nrr_y1) !== null && h.nrr_y1 < THRESHOLDS.nrr) f.push('NRR under 100% at month 12');
  if (num(h.nrr_y3) !== null && h.nrr_y3 < THRESHOLDS.nrr) f.push('NRR under 100% in year 3');
  if (num(h.payback_m12) !== null && h.payback_m12 > THRESHOLDS.paybackMonths) f.push(`CAC payback over ${THRESHOLDS.paybackMonths} months at month 12`);
  if (num(h.payback_m24) !== null && h.payback_m24 > THRESHOLDS.paybackMonths) f.push(`CAC payback over ${THRESHOLDS.paybackMonths} months at month 24`);
  if (h.master !== 1) f.push(`Model check failure: ${h.failing.join('; ')}`);
  return [...new Set(f)];
}

// ---- scenario sets -------------------------------------------------------------------------------------------
export function diffFromBase(model, values) {
  const base = model.scenarioValues('Base');
  return Object.fromEntries(Object.entries(values).filter(([n, v]) => Math.abs(v - base[n]) > 1e-12));
}
/** Full input map for a preset or for a saved scenario ({from?, values: {name: value}} = changes from Base). */
export function valuesFor(model, spec) {
  if (spec.preset) return model.scenarioValues(spec.preset);
  if (spec.current) return model.currentValues();
  return { ...model.scenarioValues('Base'), ...spec.changes };
}
export function inputDiff(model, items) {
  const rows = [];
  for (const r of model.registry) {
    const vals = items.map(it => it.values[r.name]);
    if (Math.max(...vals) - Math.min(...vals) > 1e-12) rows.push({ name: r.name, label: r.label, unit: r.unit, values: vals });
  }
  return rows;
}
/** Evaluate several full input maps in one session. Returns results aligned with items. */
export async function evaluateItems(model, items, { include = null, read = headline, onProgress } = {}) {
  const names = model.registry.map(r => r.name);
  const maps = items.map(it => (include === null ? it.values : { ...it.values, inp_ma_include: include }));
  const results = [];
  await model.withSession(names, async (run, eng) => {
    for (let i = 0; i < maps.length; i++) {
      run(maps[i]);
      results.push(read(eng));
      onProgress?.(i + 1, maps.length);
      await tick();
    }
  }, maps[0]);
  return results;
}

/** Waterfall from Base to a target scenario, one changed driver at a time (order of application matters). */
export async function scenarioBridge(model, baseValues, targetValues, outputId, { onProgress } = {}) {
  const changed = model.registry.filter(r => Math.abs(targetValues[r.name] - baseValues[r.name]) > 1e-12);
  const names = changed.map(r => r.name);
  const read = OUTPUT[outputId].read;
  const outs = [];
  await model.withSession(names, async (run, eng) => {
    for (let k = 0; k <= changed.length; k++) {
      const ov = {};
      for (let j = 0; j < k; j++) ov[names[j]] = targetValues[names[j]];
      run(ov);
      outs.push(read(eng));
      onProgress?.(k + 1, changed.length + 1);
      await tick();
    }
  }, baseValues);
  const steps = changed.map((r, i) => ({ name: r.name, label: r.label, from: baseValues[r.name], to: targetValues[r.name], delta: outs[i + 1] - outs[i] }));
  return { base: outs[0], final: outs[outs.length - 1], steps };
}

// ---- tornado --------------------------------------------------------------------------------------------------
const EXTRA_KEY = ['inp_marketing_m', 'inp_sal_rd', 'inp_sal_sm', 'inp_host_ent', 'inp_ma_offer_mult', 'inp_ma_cost_syn', 'inp_ma_integration_total', 'inp_ma_wacc'];
export function keyDrivers(model) {
  return model.registry.filter(r => r.unit !== '1/0' && (r.kind === 'scenario' || EXTRA_KEY.includes(r.name)));
}
export async function tornado(model, pct, { onProgress } = {}) {
  const live = model.currentValues();
  const drivers = keyDrivers(model);
  const rows = [];
  const flex = [];
  for (const r of drivers) {
    const v = live[r.name];
    if (!Number.isFinite(v) || v === 0) { rows.push({ name: r.name, label: r.label, unit: r.unit, base: v, skipped: 'base value is 0, a percentage flex has no effect' }); continue; }
    const a = v * (1 - pct), b = v * (1 + pct);
    let lo = Math.min(a, b), hi = Math.max(a, b);
    const clampedLo = r.min !== null && lo < r.min, clampedHi = r.max !== null && hi > r.max;
    if (r.min !== null) lo = Math.max(lo, r.min);
    if (r.max !== null) hi = Math.min(hi, r.max);
    flex.push({ r, v, lo, hi, clamped: clampedLo || clampedHi });
  }
  const total = flex.length * 2 + 1;
  let done = 0;
  let baseOut;
  await model.withSession(flex.map(f => f.r.name), async (run, eng) => {
    run({});
    baseOut = readAll(eng);
    onProgress?.(++done, total);
    await tick();
    for (const f of flex) {
      run({ [f.r.name]: f.lo });
      const low = readAll(eng);
      run({ [f.r.name]: f.hi });
      const high = readAll(eng);
      rows.push({ name: f.r.name, label: f.r.label, unit: f.r.unit, base: f.v, lo: f.lo, hi: f.hi, clamped: f.clamped, atLo: low, atHi: high });
      done += 2;
      onProgress?.(done, total);
      await tick();
    }
  });
  return { base: baseOut, rows, pct };
}
/** Rank the flexed drivers for one output by swing = |high - low|. */
export function rankTornado(t, outputId) {
  return t.rows.filter(r => r.atLo).map(r => {
    const lo = r.atLo[outputId], hi = r.atHi[outputId];
    return { ...r, outLo: lo, outHi: hi, swing: Math.abs((hi ?? 0) - (lo ?? 0)) };
  }).sort((a, b) => b.swing - a.swing);
}

// ---- two-way grid ------------------------------------------------------------------------------------------------
export function linspace(a, b, n) {
  return n === 1 ? [a] : Array.from({ length: n }, (_, i) => a + (b - a) * i / (n - 1));
}
export async function heatmap(model, xName, xVals, yName, yVals, { onProgress } = {}) {
  const cells = [];
  const total = xVals.length * yVals.length;
  let done = 0;
  await model.withSession([xName, yName], async (run, eng) => {
    for (const y of yVals) {
      const rowOut = [];
      for (const x of xVals) {
        run({ [xName]: x, [yName]: y });
        rowOut.push(readAll(eng));
        onProgress?.(++done, total);
        await tick();
      }
      cells.push(rowOut);
    }
  });
  return { xName, yName, xVals, yVals, cells };
}

// ---- goal seek (bisection) ------------------------------------------------------------------------------------------
export async function goalSeek(model, outputId, target, driverName, { onProgress, maxIter = 60 } = {}) {
  const reg = model.registry.find(r => r.name === driverName);
  if (!reg || reg.min === null) throw new Error('Pick a numeric input with an allowed range.');
  const tol = Math.max(Math.abs(target) * 1e-9, OUTPUT[outputId].kind === 'months' ? 1e-4 : 1e-3);
  const iterations = [];
  let result;
  await model.withSession([driverName], async (run, eng) => {
    const f = x => { run({ [driverName]: x }); const v = OUTPUT[outputId].read(eng); return (num(v) === null ? NaN : v) - target; };
    let lo = reg.min, hi = reg.max;
    let flo = f(lo), fhi = f(hi);
    iterations.push({ i: 0, x: lo, value: flo + target, gap: flo, note: 'lower limit' }, { i: 0, x: hi, value: fhi + target, gap: fhi, note: 'upper limit' });
    if (Number.isNaN(flo) || Number.isNaN(fhi) || Math.sign(flo) === Math.sign(fhi) && flo !== 0 && fhi !== 0) {
      result = { solved: false, reason: 'The target is not reached anywhere inside the input\'s allowed range (the output does not cross the target between the limits).', lo, hi, flo: flo + target, fhi: fhi + target };
      return;
    }
    for (let i = 1; i <= maxIter; i++) {
      const mid = (lo + hi) / 2;
      const fm = f(mid);
      iterations.push({ i, x: mid, value: fm + target, gap: fm });
      onProgress?.(i, maxIter);
      await tick();
      if (Math.abs(fm) <= tol || (hi - lo) / 2 < 1e-12) { result = { solved: true, x: mid, value: fm + target, gap: fm, iterations: i }; return; }
      if (Math.sign(fm) === Math.sign(flo)) { lo = mid; flo = fm; } else { hi = mid; fhi = fm; }
    }
    result = { solved: false, reason: `Did not converge within ${maxIter} iterations.`, x: (lo + hi) / 2 };
  });
  return { ...result, iterationLog: iterations, driver: reg, outputId, target, tol };
}

// ---- acquisition grid -----------------------------------------------------------------------------------------------
export async function acquisitionGrid(model, multiples, achieved, { onProgress } = {}) {
  const cells = [];
  let done = 0;
  await model.withSession(['inp_ma_offer_mult', 'inp_ma_syn_achieved'], async (run, eng) => {
    for (const a of achieved) {
      const row = [];
      for (const m of multiples) {
        run({ inp_ma_offer_mult: m, inp_ma_syn_achieved: a });
        row.push({ value: eng.get('out_ma_value'), payback: eng.get('out_ma_payback') });
        onProgress?.(++done, multiples.length * achieved.length);
        await tick();
      }
      cells.push(row);
    }
  });
  return { multiples, achieved, cells };
}

// ---- CSV ---------------------------------------------------------------------------------------------------------------
export function toCsv(rows) {
  const esc = v => { const s = v === null || v === undefined ? '' : String(v); return /[",\n]/.test(s) ? `"${s.replace(/"/g, '""')}"` : s; };
  return rows.map(r => r.map(esc).join(',')).join('\n') + '\n';
}
