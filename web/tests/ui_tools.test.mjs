// UI test for sections 9 and 10: planner and sensitivity tool, driven through the DOM.
import { makeDom, Chart, bytes, names, ExcelJS, $, $$, fire, waitFor, reporter } from './_ui_setup.mjs';
makeDom();
const { startApp } = await import('../app.js');
const { ok, done } = reporter();
const S = await startApp({ bytes: bytes(), names: names(), Chart, ExcelJS });
const m = S.model;
const csvOf = async name => { const d = globalThis.__downloads.filter(x => x.name === name).pop(); return d ? await d.blob.text() : null; };
const drawn = id => Chart.calls.some(c => c.id === id);

// ---- planner: save, pick, compare
const inputsBefore = JSON.stringify(m.currentValues());
m.setInput('inp_churn_ent', 0.08); S.ctx.rerender();
$('#pl-name').value = 'Slow churn fix'; $('#pl-save').click();
ok(JSON.parse(localStorage.getItem('keel-scenarios')).length === 1 && $('#pl-saved').textContent.includes('Slow churn fix'), 'saved scenario stored in the browser and listed');
$('#pl-name').value = 'Slow churn fix'; $('#pl-save').click();
ok($('#pl-msg').textContent.includes('already exists'), 'duplicate name rejected');
$('#pl-name').value = ''; $('#pl-save').click();
ok($('#pl-msg').textContent.includes('name'), 'empty name rejected');
m.resetAll(); S.ctx.rerender();
const boxes = () => $$('.picker input[type=checkbox]');
ok(boxes().length === 5 && boxes().filter(b => b.checked).length === 2, 'picker offers Base, Upside, Downside, the saved scenario and live inputs; Base and Downside preselected');
boxes().find(b => b.value === 'Upside').click(); boxes().find(b => b.value === 'S:Slow churn fix').click();
ok(boxes().filter(b => b.checked).length === 4 && boxes().filter(b => !b.checked).every(b => b.disabled), 'maximum of 4 scenarios enforced (the rest are disabled)');
boxes().find(b => b.value === 'Upside').click();
$('#pl-run').click();
await waitFor(() => $('#pl-results table'));
const head = $$('#pl-results table:first-of-type thead th').map(t => t.textContent);
ok(head.length === 1 + 3 * 2 && head[1] === 'Base: Year 1' && head.includes('Slow churn fix: Year 3'), `comparison table has Year 1 and Year 3 for each scenario (${head.length} columns)`);
const rowsTxt = $$('#pl-results table:first-of-type tbody tr').map(r => r.textContent);
ok(rowsTxt.length === 10 && rowsTxt[0].startsWith('ARR') && /£\d+\.\dm/.test(rowsTxt[0]), 'ten headline rows, formatted');
ok($('#pl-results .flags').textContent.includes('Downside') && $('#pl-results .flags li.bad') && $('#pl-results .flags').textContent.includes('Pro-forma cash never negative'), 'downside flags name the failed model check');
ok(drawn('c-cmp-arr') && drawn('c-cmp-cash') && $('#c-cmp-arr-sum').textContent.includes('Base'), 'overlaid ARR and cash charts with a text summary');
const diffRows = $$('#pl-results table').pop().querySelectorAll('tbody tr').length;
ok(diffRows >= 18, `input diff lists ${diffRows} differing drivers`);
ok(JSON.stringify(m.currentValues()) === inputsBefore && m.scenario === 'Base', 'the visitor\'s own inputs were not disturbed by the comparison');
$('#pl-csv').click();
const csv = await csvOf('keel_scenario_comparison.csv');
ok(csv && csv.split('\n')[0].includes('Base: Year 1') && csv.split('\n').length >= 11, 'comparison CSV export');
// overlay without acquisition
$$('input[name=pl-overlay]').find(r => r.value === 'without').click(); fire($$('input[name=pl-overlay]').find(r => r.value === 'without'), 'change');
$('#pl-run').click();
await waitFor(() => $('#pl-results caption')?.textContent.includes('OFF'));
ok(true, 'overlay "without acquisition" is reflected in the table caption');
const pfRow = $$('#pl-results table tbody tr').find(r => r.textContent.startsWith('Pro-forma'));
ok(pfRow, 'pro-forma row present in the overlay table');
// bridge
$('#br-target').value = 'Downside'; $('#br-output').value = 'cash24'; $('#br-run').click();
await waitFor(() => $('#c-bridge-sum') && $('#c-bridge-sum').textContent.length > 10);
ok(drawn('c-bridge') && $('#br-out').textContent.includes('order') && $$('#br-out table tbody tr').length >= 18, 'scenario bridge: chart, order-of-application note and one row per changed driver');
$('#br-out button').click();
const bcsv = await csvOf('keel_scenario_bridge.csv');
ok(bcsv && bcsv.includes('Total change'), 'bridge CSV export');
// load and delete a saved scenario
$('#pl-saved button').click();
ok(Math.abs(m.value('inp_churn_ent') - 0.08) < 1e-12 && m.scenario === 'Custom', 'loading a saved scenario puts its inputs in the control panel');
$$('#pl-saved button').find(b => b.textContent === 'Delete').click();
ok(JSON.parse(localStorage.getItem('keel-scenarios')).length === 0 && boxes().length === 4, 'deleting a saved scenario removes it');
m.resetAll(); S.ctx.rerender();

// ---- tornado
$('#tn-pct').value = '0'; $('#tn-run').click();
ok($('#tn-run').parentElement.textContent.includes('between 1 and 50'), 'invalid flex percentage rejected');
$('#tn-pct').value = '10'; $('#tn-run').click();
await waitFor(() => $('#tn-box table'));
const tRows = $$('#tn-box table tbody tr');
ok(tRows.length === 27 && drawn('c-tornado'), `tornado ranks ${tRows.length} drivers with a chart`);
ok(/logos|ACV|expansion/i.test(tRows[0].textContent), `top of the ARR tornado: ${tRows[0].children[1].textContent}`);
const callsBefore = Chart.calls.length;
$('#tn-output').value = 'cash24'; fire($('#tn-output'), 'change');
ok(Chart.calls.length === callsBefore + 1 && $$('#tn-box table tbody tr')[0].children[1].textContent !== tRows[0].children[1].textContent, 'switching the output re-ranks from stored results without recalculating');
$('#tn-csv').click();
ok((await csvOf('keel_tornado.csv'))?.startsWith('rank,driver'), 'tornado CSV export');
ok($('#tn-stale').hidden, 'no stale warning yet');
m.setInput('inp_marketing_m', 500); S.ctx.rerender();
ok(!$('#tn-stale').hidden, 'changing an input marks the tornado results as out of date');
m.resetAll(); S.ctx.rerender();

// ---- heatmap
$('#hm-nx').value = '3'; $('#hm-ny').value = '3';
$('#hm-x').value = 'inp_acv_ent'; fire($('#hm-x'), 'change');
$('#hm-y').value = 'inp_logos_ent'; fire($('#hm-y'), 'change');
$('#hm-run').click();
await new Promise(r => setTimeout(r, 50));
ok(true, 'heatmap with default ranges accepted');
await waitFor(() => $('#hm-box table'));
ok($$('#hm-box tbody tr').length === 3 && $$('#hm-box tbody tr')[0].children.length === 4, 'heatmap: 3 x 3 grid');
ok($$('#hm-box td.current').length === 1, 'exactly one cell is outlined as the current case');
$('#hm-x').value = 'inp_logos_ent'; fire($('#hm-x'), 'change');
$('#hm-y').value = 'inp_churn_ent'; fire($('#hm-y'), 'change');
$('#hm-nx').value = '7'; $('#hm-ny').value = '7';
const t0 = performance.now();
$('#hm-run').click();
await waitFor(() => $$('#hm-box tbody tr').length === 7);
const secs = (performance.now() - t0) / 1000;
ok($$('#hm-box tbody tr')[0].children.length === 8 && secs < 8, `7 x 7 grid through the interface in ${secs.toFixed(1)} s (49 recalculations)`);
$('#hm-output').value = 'cash24'; fire($('#hm-output'), 'change');
ok($('#hm-box').textContent.includes('Cash, month 24'), 'switching the grid output redraws from stored results');
$('#hm-csv').click();
ok((await csvOf('keel_sensitivity_grid.csv'))?.split('\n').length >= 8, 'grid CSV export');
$('#hm-x0').value = '99'; $('#hm-run').click();
ok($('#hm-err').textContent.includes('X from must be'), `out-of-range grid limit blocked: "${$('#hm-err').textContent}"`);
$('#hm-x0').value = ''; fire($('#hm-x'), 'change');
$('#hm-y').value = $('#hm-x').value; $('#hm-run').click();
ok($('#hm-err').textContent.includes('two different'), 'same driver on both axes rejected');

// ---- goal seek
$('#gs-output').value = 'mincash'; $('#gs-target').value = '15000'; $('#gs-driver').value = 'inp_churn_ent';
$('#gs-run').click();
await waitFor(() => $('#gs-answer'));
ok($('#gs-answer').textContent.includes('16.27') && $('#gs-answer').textContent.includes('iterations'), `goal seek answer: "${$('#gs-answer').textContent.slice(0, 110)}…"`);
ok($$('#gs-box table tbody tr').length >= 20, `iteration table shows ${$$('#gs-box table tbody tr').length} rows`);
$('#gs-apply').click();
ok(Math.abs(m.value('inp_churn_ent') - 0.16276) < 1e-4 && Math.abs(m.get('out_mincash') - 15000) < 0.01, 'applying the answer sets the input and the model reaches the target');
$('#gs-csv').click();
ok((await csvOf('keel_goal_seek.csv'))?.startsWith('iteration,'), 'goal seek CSV export');
m.resetAll(); S.ctx.rerender();
$('#gs-target').value = '1e9'; $('#gs-output').value = 'arr24'; $('#gs-driver').value = 'inp_logos_ent'; $('#gs-run').click();
await waitFor(() => $('#gs-answer')?.classList.contains('err'));
ok($('#gs-answer').textContent.includes('not reached'), 'unreachable target explained');

// ---- acquisition grid
$('#ag-run').click();
await waitFor(() => $('#ag-box table'));
const agT = $$('#ag-box table');
ok(agT.length === 2 && agT[0].querySelectorAll('tbody tr').length === 5 && agT[0].querySelectorAll('tbody tr')[0].children.length === 6, 'acquisition grid: value created and payback, 5 x 5');
ok($$('#ag-box td.neg').length > 0 && agT[0].textContent.includes('▼'), 'value-destroying cells are marked with a symbol as well as colour');
$('#ag-csv').click();
ok((await csvOf('keel_acquisition_grid.csv'))?.split('\n').length === 27, 'acquisition grid CSV (25 cells + header)');
ok(JSON.stringify(m.currentValues()) === JSON.stringify(m.scenarioValues('Base')) && m.scenario === 'Base', 'after all tools, the visitor\'s inputs are exactly as they were');
done('planner and sensitivity UI');
