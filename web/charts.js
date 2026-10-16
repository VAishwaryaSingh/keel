// charts.js: Chart.js builders. Colours come from CSS custom properties so light and dark mode both work.
const css = n => getComputedStyle(document.documentElement).getPropertyValue(n).trim() || '#888';
export const palette = () => [css('--c1'), css('--c2'), css('--c3'), css('--c4'), css('--c5'), css('--c6')];

function base(extra = {}) {
  return {
    responsive: true, maintainAspectRatio: false, animation: false,
    plugins: { legend: { labels: { color: css('--text'), boxWidth: 12 } }, tooltip: { mode: 'index', intersect: false } },
    scales: { x: { ticks: { color: css('--muted'), maxRotation: 0, autoSkip: true }, grid: { color: css('--grid') } },
              y: { ticks: { color: css('--muted') }, grid: { color: css('--grid') } } },
    ...extra,
  };
}

export class ChartHost {
  constructor(Chart) { this.Chart = Chart; this.charts = new Map(); }
  draw(canvas, config) {
    const old = this.charts.get(canvas);
    if (old) old.destroy();
    this.charts.set(canvas, new this.Chart(canvas, config));
  }
  line(canvas, labels, series, opts = {}) {
    const p = palette();
    this.draw(canvas, { type: 'line', data: { labels, datasets: series.map((s, i) => ({ label: s.label, data: s.data, borderColor: p[i % 6], backgroundColor: p[i % 6], pointRadius: 0, borderWidth: 2, borderDash: s.dash || [], tension: 0.15, yAxisID: s.axis || 'y' })) },
      options: base({ scales: { ...base().scales, ...(opts.y2 ? { y2: { position: 'right', ticks: { color: css('--muted') }, grid: { display: false } } } : {}) } }) });
  }
  stacked(canvas, labels, series, opts = {}) {
    const p = palette();
    this.draw(canvas, { type: 'bar', data: { labels, datasets: series.map((s, i) => ({ label: s.label, data: s.data, backgroundColor: p[i % 6], stack: 's' })) },
      options: base({ scales: { x: { stacked: true, ticks: { color: css('--muted'), maxRotation: 0, autoSkip: true }, grid: { display: false } }, y: { stacked: true, ticks: { color: css('--muted') }, grid: { color: css('--grid') } } } }) });
  }
  /** Waterfall from [{label, from, to, kind: 'total'|'up'|'down'}] */
  waterfall(canvas, steps, opts = {}) {
    const ok = css('--good'), bad = css('--bad'), tot = css('--c1');
    const lo = Math.min(...steps.flatMap(s => [s.from, s.to]));
    const hi = Math.max(...steps.flatMap(s => [s.from, s.to]));
    const yMin = opts.truncate ? Math.floor((lo - (hi - lo) * 0.15) / 1000) * 1000 : undefined;
    this.draw(canvas, { type: 'bar', data: { labels: steps.map(s => s.label), datasets: [{ label: 'GBP k', data: steps.map(s => [s.from, s.to]),
      backgroundColor: steps.map(s => s.kind === 'total' ? tot : (s.to >= s.from ? ok : bad)) }] },
      options: base({ plugins: { legend: { display: false }, tooltip: { callbacks: { label: c => `${steps[c.dataIndex].to - steps[c.dataIndex].from >= 0 ? '+' : ''}${Math.round(steps[c.dataIndex].to - steps[c.dataIndex].from).toLocaleString('en-GB')}` } } },
        scales: { x: { ticks: { color: css('--muted'), maxRotation: 0 }, grid: { display: false } }, y: { min: yMin, ticks: { color: css('--muted') }, grid: { color: css('--grid') } } } }) });
  }
  /** Horizontal floating bars [{label, low, high, highlight}] (football field, tornado) */
  range(canvas, rows, { axisTitle = 'GBP m', name = 'Enterprise value, GBP m' } = {}) {
    const p = palette();
    this.draw(canvas, { type: 'bar', data: { labels: rows.map(r => r.label), datasets: [{ label: name, data: rows.map(r => [r.low, r.high]), backgroundColor: rows.map(r => r.highlight ? css('--c2') : p[0]) }] },
      options: base({ indexAxis: 'y', plugins: { legend: { display: false } }, scales: { x: { ticks: { color: css('--muted') }, grid: { color: css('--grid') }, title: { display: true, text: axisTitle, color: css('--muted') } }, y: { ticks: { color: css('--text') }, grid: { display: false } } } }) });
  }
}
