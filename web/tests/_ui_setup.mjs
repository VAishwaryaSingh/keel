// Shared set-up for the UI tests: a simulated DOM, a chart recorder, download interception.
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { fileURLToPath } from 'node:url';
const require = createRequire(import.meta.url);
const { JSDOM } = require('jsdom');
globalThis.HyperFormula = require('hyperformula').HyperFormula;
globalThis.XLSX = require('xlsx');
export const ExcelJS = require('exceljs');
const here = fileURLToPath(new URL('.', import.meta.url));
export function makeDom({ blockStorage = false } = {}) {
  const dom = new JSDOM(readFileSync(here + '../index.html', 'utf8'), { url: 'http://localhost/', pretendToBeVisual: true });
  const w = dom.window;
  Object.assign(globalThis, { document: w.document, window: w, Node: w.Node, getComputedStyle: w.getComputedStyle.bind(w), HTMLElement: w.HTMLElement });
  if (blockStorage) {
    Object.defineProperty(globalThis, 'localStorage', { get() { throw new Error('storage blocked'); }, configurable: true });
  } else {
    Object.defineProperty(globalThis, 'localStorage', { value: w.localStorage, configurable: true });
  }
  globalThis.__KEEL_NO_AUTOSTART = true;
  globalThis.__downloads = [];
  const orig = URL.createObjectURL;
  URL.createObjectURL = blob => { globalThis.__lastBlob = blob; return orig.call(URL, blob); };
  w.HTMLAnchorElement.prototype.click = function () { globalThis.__downloads.push({ name: this.getAttribute('download'), blob: globalThis.__lastBlob }); };
  return w;
}
export class Chart { static calls = []; constructor(c, cfg) { Chart.calls.push({ id: c.id, cfg }); } destroy() {} }
export const bytes = () => new Uint8Array(readFileSync(here + '../model/keel_saas_model.xlsx'));
export const names = () => JSON.parse(readFileSync(here + '../model/named_ranges.json', 'utf8'));
export const $ = s => document.querySelector(s);
export const $$ = s => [...document.querySelectorAll(s)];
export const fire = (el, type) => el.dispatchEvent(new window.Event(type, { bubbles: true }));
export async function waitFor(cond, ms = 60000) {
  const t0 = Date.now();
  while (!cond()) { if (Date.now() - t0 > ms) throw new Error('timeout waiting for condition'); await new Promise(r => setTimeout(r, 25)); }
}
export function reporter() {
  let failures = 0;
  return { ok: (c, m) => { console.log(`${c ? 'PASS' : 'FAIL'}  ${m}`); if (!c) failures++; }, done: label => { console.log(failures ? `\n${failures} FAILED` : `\nall ${label} tests passed`); process.exit(failures ? 1 : 0); } };
}
