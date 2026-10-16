// dom.js: small DOM helpers shared by app.js, scenarios.js and sensitivity.js.
export function h(tag, attrs = {}, ...kids) {
  const e = document.createElement(tag);
  for (const [k, v] of Object.entries(attrs || {})) {
    if (v === false || v === null || v === undefined) continue;
    if (k === 'class') e.className = v;
    else if (k === 'dataset') Object.assign(e.dataset, v);
    else if (k.startsWith('on')) e.addEventListener(k.slice(2), v);
    else e.setAttribute(k, v === true ? '' : v);
  }
  for (const kid of kids.flat()) {
    if (kid === null || kid === undefined || kid === false) continue;
    e.append(kid instanceof Node ? kid : document.createTextNode(String(kid)));
  }
  return e;
}
export const $ = (sel, root = document) => root.querySelector(sel);
export function table(headers, rows, { caption, cls = '' } = {}) {
  return h('table', { class: `tbl ${cls}` },
    caption ? h('caption', {}, caption) : null,
    h('thead', {}, h('tr', {}, headers.map((x, i) => h('th', { scope: 'col', class: i ? 'num' : '' }, x)))),
    h('tbody', {}, rows.map(r => h('tr', { class: r.cls || '' }, r.cells.map((c, i) => (i === 0 ? h('th', { scope: 'row' }, c) : h('td', { class: `num ${r.tdCls?.[i] || ''}`, style: r.styles?.[i] || null }, c)))))));
}
export const row = (cells, cls, extra = {}) => ({ cells, cls, ...extra });

/** A progress indicator for long what-if runs: <progress> plus text. */
export function progressBar() {
  const bar = h('progress', { max: 1, value: 0, 'aria-label': 'Progress' });
  const txt = h('span', { class: 'muted small', role: 'status' }, '');
  const el = h('div', { class: 'progress', hidden: true }, bar, ' ', txt);
  let t0 = 0;
  return {
    el,
    start(label = '') { t0 = performance.now(); el.hidden = false; bar.value = 0; txt.textContent = label; },
    update(done, total) { bar.max = total; bar.value = done; txt.textContent = `${done} of ${total} recalculations, ${((performance.now() - t0) / 1000).toFixed(1)} s`; },
    done(msg) { txt.textContent = msg; bar.value = bar.max; },
    hide() { el.hidden = true; },
  };
}
/** Offer a text file for download (CSV exports). */
export function downloadText(filename, text, type = 'text/csv') {
  const url = URL.createObjectURL(new Blob([text], { type }));
  const a = h('a', { href: url, download: filename });
  document.body.append(a);
  a.click();
  a.remove();
  setTimeout(() => URL.revokeObjectURL(url), 5000);
}
