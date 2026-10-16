// Loads the exact browser (UMD) builds the page pulls from the CDN, the way a <script> tag would, then runs the engine on them.
import { readFileSync } from 'node:fs';
import vm from 'node:vm';
import { fileURLToPath } from 'node:url';
const here = fileURLToPath(new URL('.', import.meta.url));
const lib = p => readFileSync(here + '../node_modules/' + p, 'utf8');
// a minimal "window" so UMD wrappers attach to globalThis the way they do in a browser
globalThis.window = globalThis; globalThis.self = globalThis;
for (const p of ['hyperformula/dist/hyperformula.full.min.js', 'xlsx/dist/xlsx.full.min.js', 'exceljs/dist/exceljs.min.js']) vm.runInThisContext(lib(p), { filename: p });
let failures = 0;
const ok = (c, m) => { console.log(`${c ? 'PASS' : 'FAIL'}  ${m}`); if (!c) failures++; };
ok(typeof globalThis.HyperFormula === 'function' && typeof globalThis.XLSX?.read === 'function' && typeof globalThis.ExcelJS?.Workbook === 'function', 'UMD builds expose HyperFormula, XLSX and ExcelJS as globals');
const { loadModel } = await import('../engine.js');
const { buildDownload } = await import('../download.js');
const bytes = new Uint8Array(readFileSync(here + '../model/keel_saas_model.xlsx'));
const names = JSON.parse(readFileSync(here + '../model/named_ranges.json', 'utf8'));
const model = await loadModel(bytes, names);
ok(Math.abs(model.get('out_arr24') - 68321.2394) < 0.01 && model.get('out_master_check') === 1, `UMD HyperFormula: ARR month 24 = ${model.get('out_arr24').toFixed(2)}, master check 1`);
model.setScenario('Downside');
ok(Math.abs(model.get('out_cash24') - 9507.778) < 0.01, `UMD HyperFormula: Downside cash month 24 = ${model.get('out_cash24').toFixed(2)}`);
model.setScenario('Base');
const out = await buildDownload(globalThis.ExcelJS, bytes, model);
ok(out.length > 100000, `UMD ExcelJS wrote a ${(out.length / 1024).toFixed(0)} KB file`);
console.log(failures ? `\n${failures} FAILED` : '\nall UMD tests passed');
process.exit(failures ? 1 : 0);
