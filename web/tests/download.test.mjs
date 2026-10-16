// Download test: the .xlsx written by the website keeps every formula and carries the visitor's inputs.
import { readFileSync, writeFileSync, mkdtempSync } from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
import { tmpdir } from 'node:os';
import { join } from 'node:path';
const require = createRequire(import.meta.url);
globalThis.HyperFormula = require('hyperformula').HyperFormula;
globalThis.XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const { loadModel } = await import('../engine.js');
const { buildDownload } = await import('../download.js');
const here = fileURLToPath(new URL('.', import.meta.url));
const bytes = new Uint8Array(readFileSync(here + '../model/keel_saas_model.xlsx'));
const names = JSON.parse(readFileSync(here + '../model/named_ranges.json', 'utf8'));
let failures = 0;
const ok = (c, m) => { console.log(`${c ? 'PASS' : 'FAIL'}  ${m}`); if (!c) failures++; };

const model = await loadModel(bytes, names);
model.setScenario('Downside');
model.setInput('inp_churn_ent', 0.07);   // scenario-driven edit -> Custom
model.setInput('inp_marketing_m', 600);  // static edit
const out = await buildDownload(ExcelJS, bytes, model, new Date('2026-10-15T12:00:00Z'));
const dir = mkdtempSync(join(tmpdir(), 'keel-'));
const file = join(dir, 'download.xlsx');
writeFileSync(file, out);
console.log('written', file, out.length, 'bytes');

// 1. re-read the written file with SheetJS: formulas intact, inputs carried
const XLSX = globalThis.XLSX;
const orig = XLSX.read(bytes, { type: 'array', cellFormula: true, sheetStubs: true });
const down = XLSX.read(out, { type: 'array', cellFormula: true, sheetStubs: true });
function formulas(wb) {
  const m = new Map();
  for (const n of wb.SheetNames) for (const [k, c] of Object.entries(wb.Sheets[n])) if (k[0] !== '!' && c.f) m.set(n + '!' + k, c.f);
  return m;
}
const fo = formulas(orig), fd = formulas(down);
let diff = 0, missing = 0;
for (const [k, f] of fo) { if (!fd.has(k)) missing++; else if (fd.get(k).replace(/\s/g, '') !== f.replace(/\s/g, '')) diff++; }
ok(fo.size > 15000 && missing === 0 && diff === 0, `all ${fo.size.toLocaleString()} formulas kept intact (missing ${missing}, changed ${diff})`);
ok(down.SheetNames.includes('Run_Log') && down.SheetNames.length === orig.SheetNames.length + 1, `Run_Log added (${down.SheetNames.length} tabs)`);

// 2. load the downloaded file into HyperFormula: same outputs as the live site state
const model2 = await loadModel(out, names);
ok(model2.scenario === 'Custom', 'downloaded file opens on the Custom scenario');
ok(Math.abs(model2.value('inp_churn_ent') - 0.07) < 1e-12 && model2.value('inp_marketing_m') === 600, 'inputs written into the named input cells');
const keys = ['out_arr24', 'out_cash24', 'out_ebitda_fy1', 'out_nrr12', 'out_ma_value', 'out_pf_cash24', 'out_rev_fy2'];
const bad = keys.filter(k => Math.abs(model2.get(k) - model.get(k)) > 1e-6 * Math.max(1, Math.abs(model.get(k))));
ok(bad.length === 0, `re-opened file reproduces the site's outputs (${keys.length} checked)${bad.length ? ' differ: ' + bad : ''}`);
ok(model2.get('out_master_check') === model.get('out_master_check'), 'master check identical after round trip');

// 3. what ExcelJS did and did not carry over (reported, not hidden)
const zipNames = (b) => Object.keys(require('jszip').loadAsync ? {} : {});
const JSZip = require('jszip');
const zo = await JSZip.loadAsync(bytes), zd = await JSZip.loadAsync(out);
const has = (z, p) => Object.keys(z.files).some(f => f.startsWith(p));
console.log(`      charts in original: ${has(zo, 'xl/charts/')}, in download: ${has(zd, 'xl/charts/')}`);
const dn = (await zd.file('xl/workbook.xml').async('string')).match(/<definedName /g) || [];
ok(dn.length >= 340, `${dn.length} defined names kept`);
const dv = (await zd.file('xl/worksheets/sheet2.xml').async('string')).includes('dataValidation');
ok(dv, 'data validation (scenario dropdown, input ranges) kept on Inputs_Drivers');
const wbx = await zd.file('xl/workbook.xml').async('string');
ok(/fullCalcOnLoad="(1|true)"/.test(wbx), 'recalculate-on-open flag set');
console.log(failures ? `\n${failures} FAILED` : '\nall download tests passed');
console.log('FILE=' + file);
process.exit(failures ? 1 : 0);
