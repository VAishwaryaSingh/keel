// With browser storage blocked, the planner still compares presets; only saving is unavailable (acceptance 5.9 #7).
import { makeDom, Chart, bytes, names, ExcelJS, $, $$, waitFor, reporter } from './_ui_setup.mjs';
makeDom({ blockStorage: true });
const { startApp } = await import('../app.js');
const { ok, done } = reporter();
const S = await startApp({ bytes: bytes(), names: names(), Chart, ExcelJS });
ok($('#pl-save').disabled && $('#planner').textContent.includes('Saving is unavailable'), 'storage blocked: Save is disabled with an explanation');
ok($$('.picker input[type=checkbox]').length === 4, 'presets and live inputs are still offered');
$('#pl-run').click();
await waitFor(() => $('#pl-results table'));
ok($$('#pl-results table:first-of-type tbody tr').length === 10, 'comparison still works without storage');
ok(S.model.get('out_master_check') === 1, 'model unaffected');
done('storage-blocked');
