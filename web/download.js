// download.js: writes the visitor's current inputs into the original workbook with ExcelJS.
// Every formula stays intact; Excel recalculates on open. A Run_Log tab records what was changed.
// Needs an ExcelJS constructor (globalThis.ExcelJS in the browser; passed in by tests).

export async function buildDownload(ExcelJS, templateBytes, model, now = new Date()) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(templateBytes);
  const drv = wb.getWorksheet('Inputs_Drivers');
  const scn = wb.getWorksheet('Inputs_Scenarios');
  if (!drv || !scn) throw new Error('Template is missing the input tabs.');

  // 1. scenario selector
  const sel = model.addr.inp_scenario;
  drv.getCell(sel.r0 + 1, sel.c0 + 1).value = model.scenario;

  // 2. Custom scenario column, written for every scenario-driven input (so a saved Custom case survives)
  for (const r of model.registry.filter(x => x.kind === 'scenario')) {
    scn.getCell(`F${r.scenarioRow}`).value = model.cell('Inputs_Scenarios', `F${r.scenarioRow}`);
  }
  // 3. static inputs, straight into their named cells
  for (const r of model.registry.filter(x => x.kind === 'static')) {
    const a = model.addr[r.name];
    drv.getCell(a.r0 + 1, a.c0 + 1).value = model.value(r.name);
  }

  // 4. recalculate on open (cached values are not written)
  wb.calcProperties = { ...(wb.calcProperties || {}), fullCalcOnLoad: true };

  // 5. Run_Log
  let log = wb.getWorksheet('Run_Log');
  if (log) wb.removeWorksheet(log.id);
  log = wb.addWorksheet('Run_Log');
  const changed = model.changedInputs();
  log.addRow(['Keel run log']).font = { bold: true, size: 14 };
  log.addRow(['Generated', now.toISOString()]);
  log.addRow(['Scenario selected', model.scenario]);
  log.addRow(['Inputs changed from Base', changed.length]);
  log.addRow(['Master check at download', model.get('out_master_check') === 1 ? 'ALL CHECKS PASS' : 'CHECK FAILURE']);
  const failing = model.failingChecks();
  if (failing.length) log.addRow(['Failing checks', failing.map(f => f.name).join('; ')]);
  log.addRow(['Note', 'Synthetic data. Formulas are the original workbook formulas; values recalculate when the file is opened. Charts in the original workbook are not carried over by this download.']);
  log.addRow([]);
  const head = log.addRow(['Input (named range)', 'Label', 'Base value', 'Value in this file']);
  head.font = { bold: true };
  for (const c of changed) log.addRow([c.name, c.label, c.base, c.now]);
  log.columns = [{ width: 30 }, { width: 56 }, { width: 16 }, { width: 18 }];
  return new Uint8Array(await wb.xlsx.writeBuffer());
}
