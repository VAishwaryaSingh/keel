// engine.js: loads the Keel workbook into HyperFormula and exposes the named-range contract.
// The website never re-implements a calculation: every number comes from the workbook's own formulas.
// Reads only out_* names; writes only inp_* names (and the Custom scenario column / scenario selector).
// Needs globalThis.HyperFormula and globalThis.XLSX (loaded from CDN in the browser, from npm in tests).

const A1 = /^\$?([A-Z]{1,3})\$?(\d+)$/;

export function colToIndex(letters) {
  let n = 0;
  for (const ch of letters) n = n * 26 + (ch.charCodeAt(0) - 64);
  return n - 1;
}
export function indexToCol(i) {
  let s = '';
  for (i += 1; i > 0; i = Math.floor((i - 1) / 26)) s = String.fromCharCode(65 + ((i - 1) % 26)) + s;
  return s;
}

export function parseAddress(text) {
  // "Outputs_Web!$D$12:$AD$12" or "Inputs_Drivers!$C$8" -> {sheet, r0, c0, r1, c1}
  const [sheetRaw, rest] = text.split('!');
  const sheet = sheetRaw.replace(/^'|'$/g, '');
  const [a, b] = rest.split(':');
  const m1 = A1.exec(a);
  const m2 = b ? A1.exec(b) : m1;
  return { sheet, r0: +m1[2] - 1, c0: colToIndex(m1[1]), r1: +m2[2] - 1, c1: colToIndex(m2[1]) };
}

export function excelSerialToDate(serial) {
  return new Date(Date.UTC(1899, 11, 30) + Math.round(serial) * 86400000);
}

export class Model {
  constructor(hf, names, sheetIds, parent = null) {
    this.hf = hf;
    this.names = names;                // name -> "Sheet!$A$1:$B$2"
    this.sheetIds = sheetIds;
    this.parent = parent;              // a lean what-if instance reads Outputs_Web through its parent's links
    this._linkCache = {};
    this.addr = {};
    for (const [n, t] of Object.entries(names)) this.addr[n] = parseAddress(t);
    this.buildMs = 0;
    this.lastRecalcMs = 0;
    if (parent) {
      this.registry = parent.registry;
      this.initial = parent.initial;
    } else {
      this.registry = this._readRegistry();
      this.initial = {};                  // static input name -> value when loaded
      for (const r of this.registry) if (r.kind === 'static') this.initial[r.name] = this.value(r.name);
    }
    this.selectorCell = this.addr.inp_scenario;
    // direct address -> input name, so formulas read as the workbook shows them (inp_churn_ent, not Inputs_Drivers!$C$20)
    this.addrToName = {};
    for (const [n, t] of Object.entries(names)) {
      const a = parseAddress(t);
      if (a.r0 === a.r1 && a.c0 === a.c1 && (n.startsWith('inp_') || n.startsWith('chk_'))) this.addrToName[`${a.sheet}!$${indexToCol(a.c0)}$${a.r0 + 1}`] = n;
    }
  }
  prettyFormula(f) {
    if (!f) return f;
    return f.replace(/(?:'?[A-Za-z_ ]+'?!)\$[A-Z]{1,3}\$\d+/g, m => this.addrToName[m.replace(/'/g, '')] || m);
  }
  seriesLabel(name) {
    const a = this.addr[name];
    return String(this.cell('Outputs_Web', `B${a.r0 + 1}`) ?? name).replace(/^[^:]+: /, '');
  }
  seriesUnit(name) {
    const a = this.addr[name];
    return this.cell('Outputs_Web', `C${a.r0 + 1}`);
  }

  // ---- reading --------------------------------------------------------------------------------
  _range(a) {
    if (this.parent && a.sheet === 'Outputs_Web') return this._leanRange(a);
    const id = this.sheetIds[a.sheet];
    return this.hf.getRangeValues({ start: { sheet: id, row: a.r0, col: a.c0 }, end: { sheet: id, row: a.r1, col: a.c1 } });
  }
  // Lean instance: Outputs_Web is not loaded (5,000+ link cells that a what-if run does not need). Each
  // Outputs_Web cell is a pure link, so read the cell it points at instead.
  _leanCell(r, c) {
    const key = r + ',' + c;
    let link = this._linkCache[key];
    if (link === undefined) {
      const f = this.parent.hf.getCellFormula({ sheet: this.parent.sheetIds.Outputs_Web, row: r, col: c });
      const m = f ? /^=?\s*'?([A-Za-z_ ]+)'?!\$?([A-Z]+)\$?(\d+)$/.exec(f) : null;
      link = m && this.sheetIds[m[1]] !== undefined ? { id: this.sheetIds[m[1]], row: +m[3] - 1, col: colToIndex(m[2]) } : null;
      this._linkCache[key] = link;
    }
    if (link) return this.hf.getCellValue({ sheet: link.id, row: link.row, col: link.col });
    return this.parent.hf.getCellValue({ sheet: this.parent.sheetIds.Outputs_Web, row: r, col: c });
  }
  _leanRange(a) {
    const out = [];
    for (let r = a.r0; r <= a.r1; r++) {
      const row = [];
      for (let c = a.c0; c <= a.c1; c++) row.push(this._leanCell(r, c));
      out.push(row);
    }
    return out;
  }
  /** Scalar -> value; one row -> 1-D array; otherwise 2-D array. */
  get(name) {
    const a = this.addr[name];
    if (!a) throw new Error(`Unknown named range: ${name}`);
    const v = this._range(a);
    if (a.r0 === a.r1 && a.c0 === a.c1) return v[0][0];
    if (a.r0 === a.r1) return v[0];
    return v;
  }
  value(name) {
    return this.get(name);
  }
  cell(sheet, a1) {
    if (this.parent && sheet === 'Outputs_Web') return this.parent.cell(sheet, a1);
    const m = A1.exec(a1);
    return this.hf.getCellValue({ sheet: this.sheetIds[sheet], row: +m[2] - 1, col: colToIndex(m[1]) });
  }
  formula(sheet, a1) {
    const m = A1.exec(a1);
    return this.hf.getCellFormula({ sheet: this.sheetIds[sheet], row: +m[2] - 1, col: colToIndex(m[1]) }) || null;
  }
  _readRegistry() {
    const rows = this.get('out_input_registry');
    return rows.filter(r => r[0]).map(r => ({
      name: r[0], label: r[1], unit: r[2], live: r[3], min: r[4] === '' ? null : r[4], max: r[5] === '' ? null : r[5],
      kind: r[6], scenarioRow: r[7] || null, section: r[8], tip: r[9],
    }));
  }
  /** Series rows are [opening, m1..m24, FY1, FY2]. */
  series(name) {
    const v = this.get(name);
    return { opening: v[0], months: v.slice(1, 25), fy1: v[25], fy2: v[26] };
  }
  /** Where does an Outputs_Web cell come from? Parses the pure link in the cell. */
  source(name, index = 0) {
    const a = this.addr[name];
    const col = a.c0 + index;
    const f = this.formula('Outputs_Web', `${indexToCol(col)}${a.r0 + 1}`);
    if (!f) return null;
    const m = /^=?'?([A-Za-z_ ]+)'?!\$?([A-Z]+)\$?(\d+)$/.exec(f);
    return m ? { sheet: m[1], a1: `${m[2]}${m[3]}`, row: +m[3], col: m[2] } : null;
  }
  /** One level of precedents of a source cell: references and their current values. */
  precedents(sheet, a1) {
    const f = this.formula(sheet, a1);
    if (!f) return { formula: null, refs: [] };
    const refs = [];
    const seen = new Set();
    const re = /(?:'?([A-Za-z_ ]+)'?!)?(\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?)|\b((?:inp|chk)_[A-Za-z0-9_]+)\b/g;
    let m;
    while ((m = re.exec(f))) {
      const key = m[0];
      if (seen.has(key)) continue;
      seen.add(key);
      if (m[3]) {
        refs.push({ text: m[3], kind: 'input', value: this.safeValue(m[3]) });
      } else {
        const sh = m[1] || sheet;
        const ref = m[2].replace(/\$/g, '');
        const single = !ref.includes(':');
        refs.push({ text: `${sh}!${ref}`, sheet: sh, ref, kind: 'cell', value: single ? this.cell(sh, ref) : '(range)' });
      }
    }
    return { formula: f, refs };
  }
  safeValue(name) {
    try { return this.get(name); } catch { return null; }
  }
  rowLabel(sheet, row) {
    return this.cell(sheet, `A${row}`);
  }

  // ---- writing ----------------------------------------------------------------------------------
  _set(sheet, a1, value) {
    const m = A1.exec(a1);
    this.hf.setCellContents({ sheet: this.sheetIds[sheet], row: +m[2] - 1, col: colToIndex(m[1]) }, [[value]]);
  }
  _timed(fn) {
    const t0 = performance.now();
    this.hf.batch(fn);
    this.lastRecalcMs = performance.now() - t0;
  }
  get scenario() { return this.get('inp_scenario'); }
  scenarioColumn(name) { return { Base: 'C', Upside: 'D', Downside: 'E', Custom: 'F' }[name]; }
  _cellOf(name) {
    const a = this.addr[name];
    return { sheet: a.sheet, a1: `${indexToCol(a.c0)}${a.r0 + 1}` };
  }
  /** Values of the live scenario's column, read BEFORE any batch (HyperFormula suspends evaluation inside one). */
  _liveScenarioValues() {
    const from = this.scenarioColumn(this.scenario);
    if (from === 'F') return null;
    return this.registry.filter(r => r.kind === 'scenario').map(r => [r.scenarioRow, this.cell('Inputs_Scenarios', `${from}${r.scenarioRow}`)]);
  }
  setScenario(name) {
    const c = this._cellOf('inp_scenario');
    this._timed(() => this._set(c.sheet, c.a1, name));
  }
  /** Set an input. Returns {ok, message}. Invalid values never reach the engine. */
  setInput(name, value) {
    const r = this.registry.find(x => x.name === name);
    if (!r) return { ok: false, message: `Unknown input ${name}` };
    const v = Number(value);
    if (!Number.isFinite(v)) return { ok: false, message: 'Enter a number.' };
    if (r.min !== null && v < r.min) return { ok: false, message: `Minimum is ${r.min}.` };
    if (r.max !== null && v > r.max) return { ok: false, message: `Maximum is ${r.max}.` };
    const copy = r.kind === 'scenario' ? this._liveScenarioValues() : null;
    this._timed(() => {
      if (r.kind === 'scenario') {
        // start from the scenario the visitor was looking at, then apply the edit
        if (copy) for (const [row, val] of copy) this._set('Inputs_Scenarios', `F${row}`, val);
        this._set('Inputs_Scenarios', `F${r.scenarioRow}`, v);
        const c = this._cellOf('inp_scenario');
        this._set(c.sheet, c.a1, 'Custom');
      } else {
        const c = this._cellOf(name);
        this._set(c.sheet, c.a1, v);
      }
    });
    return { ok: true };
  }
  baseValue(name) {
    const r = this.registry.find(x => x.name === name);
    return r.kind === 'scenario' ? this.cell('Inputs_Scenarios', `C${r.scenarioRow}`) : this.initial[name];
  }
  isChanged(name) {
    const r = this.registry.find(x => x.name === name);
    return Math.abs(this.value(name) - this.baseValue(name)) > 1e-12 || (r.kind === 'scenario' && false);
  }
  changedInputs() {
    return this.registry.filter(r => this.isChanged(r.name)).map(r => ({ name: r.name, label: r.label, base: this.baseValue(r.name), now: this.value(r.name) }));
  }
  resetInput(name) {
    return this.setInput(name, this.baseValue(name));
  }
  resetAll() {
    const bases = this.registry.map(r => [r, r.kind === 'scenario' ? this.cell('Inputs_Scenarios', `C${r.scenarioRow}`) : this.initial[r.name]]);
    this._timed(() => {
      for (const [r, v] of bases) {
        if (r.kind === 'scenario') this._set('Inputs_Scenarios', `F${r.scenarioRow}`, v);
        else {
          const c = this._cellOf(r.name);
          this._set(c.sheet, c.a1, v);
        }
      }
      const c = this._cellOf('inp_scenario');
      this._set(c.sheet, c.a1, 'Base');
    });
  }
  /** The lean what-if engine (built on first use; about a second). */
  getLean() {
    if (this.parent) return this;
    if (!this._lean) {
      const sheets = {};
      for (const [n, rows] of Object.entries(this.parsed)) if (n !== 'Outputs_Web') sheets[n] = rows;
      const t0 = performance.now();
      const hf = globalThis.HyperFormula.buildFromSheets(sheets, this.config);
      const ids = {};
      for (const name of hf.getSheetNames()) ids[name] = hf.getSheetId(name);
      this._lean = new Model(hf, this.names, ids, this);
      this._lean.buildMs = performance.now() - t0;
    }
    return this._lean;
  }

  // ---- what-if sessions: many recalculations without disturbing what the visitor sees -------------------------
  /** Current value of every input, by name. */
  currentValues() {
    return Object.fromEntries(this.registry.map(r => [r.name, this.value(r.name)]));
  }
  /** Every input's value in a preset scenario (Base/Upside/Downside/Custom); static inputs take their loaded value. */
  scenarioValues(preset) {
    const col = this.scenarioColumn(preset);
    return Object.fromEntries(this.registry.map(r => [r.name, r.kind === 'scenario' ? this.cell('Inputs_Scenarios', `${col}${r.scenarioRow}`) : this.initial[r.name]]));
  }
  /**
   * Open a session that varies the named inputs. The current state is saved and restored by endSession().
   * run(overrides) sets every named input to its override (or its value at session start) in ONE batch,
   * so each run costs one recalculation.
   */
  beginSession(names, liveValues = null) {
    if (this._sess) throw new Error('A what-if session is already open.');
    const regs = Object.fromEntries(this.registry.map(r => [r.name, r]));
    const live = liveValues || this.currentValues();
    const saved = { scenario: this.scenario, custom: {}, statics: {} };
    for (const r of this.registry) {
      if (r.kind === 'scenario') saved.custom[r.scenarioRow] = this.cell('Inputs_Scenarios', `F${r.scenarioRow}`);
      else saved.statics[r.name] = this.value(r.name);
    }
    const base = {};
    for (const n of names) {
      if (!regs[n]) throw new Error(`Unknown input ${n}`);
      base[n] = live[n];
    }
    // make this instance equal to the visitor's state: Custom column = live scenario values, statics = live statics
    const staticWrites = this.registry.filter(r => r.kind === 'static' && this.value(r.name) !== live[r.name]);   // read before the batch
    this._timed(() => {
      for (const r of this.registry) if (r.kind === 'scenario') this._set('Inputs_Scenarios', `F${r.scenarioRow}`, live[r.name]);
      for (const r of staticWrites) { const c = this._cellOf(r.name); this._set(c.sheet, c.a1, live[r.name]); }
      const c = this._cellOf('inp_scenario');
      this._set(c.sheet, c.a1, 'Custom');
    });
    this._sess = { saved, base, last: { ...base }, regs };
  }
  run(overrides = {}) {
    const s = this._sess;
    if (!s) throw new Error('No what-if session is open.');
    const todo = [];
    for (const n of Object.keys(s.base)) {
      const v = n in overrides ? overrides[n] : s.base[n];
      if (v !== s.last[n]) todo.push([n, v]);
    }
    for (const n of Object.keys(overrides)) if (!(n in s.base)) throw new Error(`${n} was not declared when the session began.`);
    if (todo.length) {
      this._timed(() => {
        for (const [n, v] of todo) {
          const r = s.regs[n];
          if (r.kind === 'scenario') this._set('Inputs_Scenarios', `F${r.scenarioRow}`, v);
          else { const c = this._cellOf(n); this._set(c.sheet, c.a1, v); }
          s.last[n] = v;
        }
      });
    }
    return this;
  }
  endSession() {
    const s = this._sess;
    if (!s) return;
    this._timed(() => {
      for (const [row, v] of Object.entries(s.saved.custom)) this._set('Inputs_Scenarios', `F${row}`, v);
      for (const [n, v] of Object.entries(s.saved.statics)) { const c = this._cellOf(n); this._set(c.sheet, c.a1, v); }
      const c = this._cellOf('inp_scenario');
      this._set(c.sheet, c.a1, s.saved.scenario);
    });
    this._sess = null;
  }
  /** Run fn(run) inside a session; always restores the visitor's state. */
  async withSession(names, fn, live = null, useLean = true) {
    const eng = useLean ? this.getLean() : this;
    eng.beginSession(names, live || this.currentValues());
    try { return await fn(o => eng.run(o), eng); } finally { eng.endSession(); }
  }

  /** Failing checks: [{name, value, tol}] */
  failingChecks() {
    return this.get('out_checks').filter(r => r[3] !== 1).map(r => ({ name: r[0], value: r[1], tol: r[2] }));
  }
}

export async function loadModel(xlsxBytes, names) {
  const HF = globalThis.HyperFormula;
  const XLSX = globalThis.XLSX;
  if (!HF || !XLSX) throw new Error('HyperFormula and XLSX must be loaded first.');
  const t0 = performance.now();
  const wb = XLSX.read(xlsxBytes, { type: 'array', cellFormula: true, cellDates: false, cellStyles: false, sheetStubs: true });
  const sheets = {};
  for (const name of wb.SheetNames) {
    const ws = wb.Sheets[name];
    if (!ws['!ref']) { sheets[name] = [[]]; continue; }
    const rg = XLSX.utils.decode_range(ws['!ref']);
    const rows = [];
    for (let r = 0; r <= rg.e.r; r++) {
      const row = [];
      for (let c = 0; c <= rg.e.c; c++) {
        const cell = ws[XLSX.utils.encode_cell({ r, c })];
        if (!cell) row.push(null);
        else if (cell.f) row.push('=' + cell.f);
        else row.push(cell.t === 'e' ? null : cell.v);
      }
      rows.push(row);
    }
    sheets[name] = rows;
  }
  // HyperFormula's named expressions fail to evaluate when a name points at a cell that itself holds a formula
  // (inp_scen_idx, chk_master, ...). Single-cell names are therefore replaced by their direct address, which is
  // semantically identical. Range names (out_*) are not used inside formulas; the site resolves them itself.
  const single = {};
  for (const [n, t] of Object.entries(names)) {
    const a = parseAddress(t);
    if (a.r0 === a.r1 && a.c0 === a.c1) single[n] = t.replace(/^'?([^'!]+)'?!/, '$1!');
  }
  const nameRe = new RegExp('\\b(' + Object.keys(single).map(n => n.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('|') + ')\\b', 'g');
  for (const rows of Object.values(sheets)) {
    for (const row of rows) {
      for (let i = 0; i < row.length; i++) {
        const v = row[i];
        if (typeof v === 'string' && v.charCodeAt(0) === 61 && (v.includes('inp_') || v.includes('chk_'))) row[i] = v.replace(nameRe, m => single[m]);
      }
    }
  }
  const config = { licenseKey: 'gpl-v3', evaluateNullToZero: true, useArrayArithmetic: false, useColumnIndex: false };
  const hf = HF.buildFromSheets(sheets, config);
  const sheetIds = {};
  for (const name of hf.getSheetNames()) sheetIds[name] = hf.getSheetId(name);
  const model = new Model(hf, names, sheetIds);
  model.parsed = sheets;
  model.config = config;
  model.buildMs = performance.now() - t0;
  return model;
}
