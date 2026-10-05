function getRecomendacion(perfilId, fila) {
  const cfg = getConfig(perfilId);
  const def = String(cfg.recomendacion || '').trim();
  if (!def) return { def: '', tipo: '', por: '' };
  const sh = getSheet(cfg);
  if (sh.getMaxColumns() < 48) return { def: def, tipo: '', por: '' };
  const v = sh.getRange(fila, 47, 1, 2).getValues()[0];
  return { def: def, tipo: String(v[0] || ''), por: String(v[1] || '') };
}

function guardarRecomendacion(perfilId, fila, tipo, por) {
  const sh = getSheet(getConfig(perfilId));
  if (sh.getMaxColumns() < 48) sh.insertColumnsAfter(sh.getMaxColumns(), 48 - sh.getMaxColumns());
  const h = sh.getRange(1, 47, 1, 2);
  if (!String(h.getValue()).trim()) h.setValues([['Recomendó', 'Recomendado por']]).setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
  sh.getRange(fila, 47, 1, 2).clearDataValidations().setValues([[tipo || '', por || '']]);
  return true;
}

function instalarRecomendacionPAMI() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const config = ss.getSheetByName('Config');
  const data = config.getDataRange().getValues();
  const headers = data[0];
  let c = headers.indexOf('recomendacion') + 1;
  if (!c) { c = headers.length + 1; config.getRange(1, c).setValue('recomendacion'); }
  const i = data.findIndex((r, k) => k > 0 && /pami|enfermer/i.test([r[0], r[1], r[3]].join(' ')));
  if (i < 1) throw new Error('No encontré el perfil PAMI en Config');
  config.getRange(i + 1, c).setValue('Sindicato:NORTE;UTI;UTERA;APPAMIA;ATE;UPCN;SUTEPA|Director|Empleado');
  Logger.log('✅ Config.recomendacion cargada para ' + data[i][0]);
}
