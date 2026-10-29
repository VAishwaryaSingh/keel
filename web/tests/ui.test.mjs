// UI test: drives the real app.js against a simulated DOM (jsdom). Chart.js is replaced by a recorder.
// This tests behaviour and wiring, not visual layout: layout needs a real browser and a human eye.
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const { JSDOM } = require('jsdom');
globalThis.HyperFormula = require('hyperformula').HyperFormula;
globalThis.XLSX = require('xlsx');
const ExcelJS = require('exceljs');
const here = fileURLToPath(new URL('.', import.meta.url));
const html = readFileSync(here + '../index.html', 'utf8');
const dom = new JSDOM(html, { url: 'http://localhost/', pretendToBeVisual: true });
const w = dom.window;
Object.assign(globalThis, { document: w.document, window: w, Node: w.Node, getComputedStyle: w.getComputedStyle.bind(w), HTMLElement: w.HTMLElement });
globalThis.__KEEL_NO_AUTOSTART = true;
w.HTMLAnchorElement.prototype.click = function () { globalThis.__lastDownload = { href: this.href, name: this.getAttribute('download') }; };
const errors = [];
process.on('uncaughtException', e => errors.push(e));
class Chart { static calls = []; constructor(c, cfg) { Chart.calls.push({ id: c.id, cfg }); } destroy() {} }

const bytes = new Uint8Array(readFileSync(here + '../model/keel_saas_model.xlsx'));
const names = JSON.parse(readFileSync(here + '../model/named_ranges.json', 'utf8'));
const { startApp } = await import('../app.js');
let failures = 0;
const ok = (c, m) => { console.log(`${c ? 'PASS' : 'FAIL'}  ${m}`); if (!c) failures++; };
const $ = s => document.querySelector(s), $$ = s => [...document.querySelectorAll(s)];
const sleep = ms => new Promise(r => setTimeout(r, ms));
const fire = (el, type) => el.dispatchEvent(new w.Event(type, { bubbles: true }));

const t0 = performance.now();
const S = await startApp({ bytes, names, Chart, ExcelJS });
console.log(`app started and first render complete in ${(performance.now() - t0).toFixed(0)} ms`);

// structure
ok(['overview', 'assumptions', 'revenue', 'revrec', 'costs', 'statements', 'health', 'budget', 'planner', 'sensitivity', 'acquisition', 'checks', 'download'].every(id => $('#' + id) && $('#' + id + ' h2')), 'all 13 sections present, each with a heading');
ok($$('#sidenav a').length === 13, 'side navigation lists 13 sections');
ok($$('.kpi').length >= 24, `${$$('.kpi').length} metric cards rendered (overview, health, acquisition)`);
const arrCard = $('[data-trace="out_arr24"] .kpi-value').textContent;
ok(/^£\d+\.\dm$/.test(arrCard), `ARR month 24 card reads ${arrCard}`);
ok($('#master-text').textContent === 'All checks pass' && $('#master-light').classList.contains('ok') && $('#banner').hidden, 'status bar: green light, "All checks pass", no banner');
ok($('#scenario-select').value === 'Base' && $('#changed-count').textContent === '0 inputs changed', 'status bar: Base scenario, 0 inputs changed');
ok(Chart.calls.length >= 12, `${Chart.calls.length} charts drawn`);
ok($$('.chart .summary').every(p => p.textContent.length > 20), 'every chart has a text summary for screen readers');
ok($$('canvas').every(c => c.getAttribute('aria-label')), 'every chart canvas has an accessible label');

// panel
const fields = $$('#panel .field');
ok(fields.length === 99, `${fields.length} input fields in the panel (one per registry entry)`);
ok(fields.every(f => { const l = f.querySelector('label'); const i = f.querySelector('input[type=number],input[type=checkbox]'); return l && i && l.getAttribute('for') === i.id; }), 'every field has a label bound to its control');
ok($$('input[type=range]').length > 80, `${$$('input[type=range]').length} sliders`);
ok($$('#panel input[type=checkbox]').length === 2, 'two on/off toggles (include acquisition, capitalise commission)');
ok($$('details.group').length >= 8, `${$$('details.group').length} input groups`);

// edit a scenario-driven input (churn 6% -> 8%)
const box = $('#in-inp_churn_ent');
ok(box.value === '6', `churn box shows 6 (percent): "${box.value}"`);
const before = $('[data-trace="out_arr24"] .kpi-value').textContent;
box.value = '8'; fire(box, 'input');
await sleep(100);
ok(S.model.value('inp_churn_ent') === 0.06, 'debounce: nothing applied after 100 ms');
await sleep(300);
ok(Math.abs(S.model.value('inp_churn_ent') - 0.08) < 1e-12 && S.model.scenario === 'Custom', 'after 300 ms the edit reaches the model and the scenario becomes Custom');
ok($('#scenario-select').value === 'Custom' && $('#changed-count').textContent === '1 input changed', 'status bar shows Custom and "1 input changed"');
ok($('.field[data-name="inp_churn_ent"]').classList.contains('changed'), 'changed field is highlighted');
ok($('[data-trace="out_arr24"] .kpi-value').textContent !== before, `ARR card updated: ${before} -> ${$('[data-trace="out_arr24"] .kpi-value').textContent}`);
ok($('#recalc').textContent.includes('ms'), `last recalculated: "${$('#recalc').textContent}"`);

// invalid input is blocked
const arrNow = S.model.get('out_arr24');
box.value = '90'; fire(box, 'input');
await sleep(450);
ok($('#in-inp_churn_ent-err').textContent.includes('Maximum is 50%') && S.model.get('out_arr24') === arrNow && box.getAttribute('aria-invalid') === 'true', `invalid 90% blocked: "${$('#in-inp_churn_ent-err').textContent}"`);

// per-field reset, then reset all
$('.field[data-name="inp_churn_ent"] .reset').click();
ok(Math.abs(S.model.value('inp_churn_ent') - 0.06) < 1e-12, 'per-field reset restores Base value');
box.value = '9'; fire(box, 'input'); await sleep(450);
$('#reset-all').click();
ok(S.model.scenario === 'Base' && $('#changed-count').textContent === '0 inputs changed', 'reset all returns to Base');

// section 2: every input editable in a table, kept in sync with the panel
ok($('#assumptions-h').textContent === '2. Inputs' && $$('#assumptions .in-tbl tbody tr').length === 99, `${$$('#assumptions .in-tbl tbody tr').length} editable rows in the Inputs section`);
ok($$('#assumptions .in-tbl tbody tr').every(tr => { const l = tr.querySelector('label'); return l && $('#' + l.getAttribute('for')); }), 'every Inputs-section row has a label bound to its control');
const cell = $('#sec-in-inp_churn_ent');
ok(cell.value === '6', `Inputs-section churn shows 6: "${cell.value}"`);
cell.value = '8'; fire(cell, 'input'); await sleep(450);
ok(Math.abs(S.model.value('inp_churn_ent') - 0.08) < 1e-12 && S.model.scenario === 'Custom' && box.value === '8', 'editing in the Inputs section reaches the model and updates the panel');
ok($('#assumptions tr[data-name="inp_churn_ent"]').classList.contains('changed'), 'changed row is highlighted in the Inputs section');
cell.value = '90'; fire(cell, 'input'); await sleep(450);
ok($('#sec-in-inp_churn_ent-err').textContent.includes('Maximum is 50%') && Math.abs(S.model.value('inp_churn_ent') - 0.08) < 1e-12, 'Inputs section blocks out-of-range values');
box.value = '7'; fire(box, 'input'); await sleep(450);
ok(cell.value === '7', 'editing in the panel updates the Inputs section');
$('#reset-all').click();
ok(cell.value === '6' && !$('#assumptions tr[data-name="inp_churn_ent"]').classList.contains('changed'), 'reset all clears the Inputs section');

// toggle
const tog = $('#in-inp_ma_include');
const ev0 = S.model.get('out_pf_cash24');
tog.checked = false; fire(tog, 'change');
ok(S.model.value('inp_ma_include') === 0 && Math.abs(S.model.get('out_pf_cash24') - S.model.get('out_cash24')) < 1e-6, 'include-acquisition off: pro-forma cash equals standalone');
ok($('#c-pfarr-sum').textContent.includes('coincide'), 'acquisition summary notes that the lines coincide');
$('#reset-all').click();
ok(S.model.get('out_pf_cash24') === ev0, 'toggle reset');

// Downside: broken model is never hidden
const sel = $('#scenario-select');
sel.value = 'Downside'; fire(sel, 'change');
ok($('#master-light').classList.contains('fail') && !$('#banner').hidden && $('#banner').textContent.includes('Pro-forma cash never negative'), 'Downside: red light and banner naming the failing check');
ok($$('#checks-table tr.bad').length === 1, 'checks table shows exactly one failing row');
sel.value = 'Base'; fire(sel, 'change');
ok($('#banner').hidden, 'back to Base: banner hidden');

// statements
$$('#statements .tab').find(b => b.dataset.k === 'bs').click();
ok($('#stmt caption').textContent.startsWith('Balance sheet'), 'balance sheet tab');
const per = $('#period-select'); per.value = 'm12'; fire(per, 'change');
ok($('#stmt caption').textContent.includes('Month 12') && $('#stmt-check').textContent.includes('balances'), `period selector + balance check: "${$('#stmt-check').textContent}"`);
$$('#statements .tab').find(b => b.dataset.k === 'cf').click();
ok($('#stmt caption').textContent.startsWith('Cash flow'), 'cash flow tab');

// trace
$('[data-trace="out_nrr12"]').click();
ok(!$('#trace').hidden && $('#trace-body').textContent.includes('SaaS_Metrics') && $$('#trace-body tbody tr').length >= 2, 'trace panel shows source cell and its precedents');
ok($('#trace-body').textContent.includes('Jump to'), 'trace panel offers a jump to the section');
$('#trace .btn').click();

// workings
$$('#revenue .workings .btn')[0].click();
ok(!$('#revenue-w').hidden && $('#revenue-w').textContent.includes('inp_') && $$('#revenue-w tbody tr').length >= 4, 'show the workings: real formulas with named inputs');

// download
$('#dl-xlsx').click();
for (let t = 0; t < 200 && $('#dl-status').textContent.startsWith('Preparing'); t++) await sleep(100);   // slow machines need more than 2.5 s
ok($('#dl-status').textContent.startsWith('Downloaded'), `download: "${$('#dl-status').textContent}"`);
ok(globalThis.__lastDownload && /^keel_model_base\.xlsx$/.test(globalThis.__lastDownload.name), 'download file name');

// browser storage disabled: the site still works (only the theme preference uses storage)
Object.defineProperty(w, 'localStorage', { get() { throw new Error('storage blocked'); }, configurable: true });
Object.defineProperty(globalThis, 'localStorage', { get() { throw new Error('storage blocked'); }, configurable: true });
let threw = false;
try { $('#theme-toggle').click(); } catch { threw = true; }
ok(!threw && document.documentElement.dataset.theme && S.model.get('out_master_check') === 1, 'storage disabled: theme toggle still works and the model is unaffected');

// budget table
ok($$('#bvf tbody tr').length === 10 && $('#bvf').textContent.includes('Adverse'), 'budget table: 10 lines with flags');
ok(errors.length === 0, `no uncaught errors${errors.length ? ': ' + errors[0].message : ''}`);
console.log(failures ? `\n${failures} FAILED` : '\nall UI tests passed');
process.exit(failures ? 1 : 0);
