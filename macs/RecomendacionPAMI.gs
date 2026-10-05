function getRecomendacion(perfilId, fila) {
  const sh = getSheet(getConfig(perfilId));
  if (sh.getMaxColumns() < 48) return { tipo: '', por: '' };
  const v = sh.getRange(fila, 47, 1, 2).getValues()[0];
  return { tipo: String(v[0] || ''), por: String(v[1] || '') };
}

function guardarRecomendacion(perfilId, fila, tipo, por) {
  const sh = getSheet(getConfig(perfilId));
  if (sh.getMaxColumns() < 48) sh.insertColumnsAfter(sh.getMaxColumns(), 48 - sh.getMaxColumns());
  const h = sh.getRange(1, 47, 1, 2);
  if (!String(h.getValue()).trim()) h.setValues([['Recomendó', 'Recomendado por']]).setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
  sh.getRange(fila, 47, 1, 2).clearDataValidations().setValues([[tipo || '', por || '']]);
  return true;
}
