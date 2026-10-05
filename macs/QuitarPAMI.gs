function quitarPAMIdeMACS() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const config = ss.getSheetByName('Config');
  const data = config.getDataRange().getValues();
  const headers = data[0];
  const i = data.findIndex((r, k) => k > 0 && /pami|enfermer/i.test([r[0], r[1], r[3]].join(' ')));
  if (i < 1) { Logger.log('No hay perfil PAMI en Config. Nada que borrar.'); return; }
  const id = String(data[i][0]);
  const hojaNombre = String(data[i][3]);

  ['Perfiles_Hero', 'Perfiles_Screening_Eval', 'Perfiles_Screening_Val', 'Perfiles_Competencias', 'Postulacion_Preguntas', 'Mail_Routing'].forEach(n => {
    const sh = ss.getSheetByName(n);
    if (!sh) return;
    const d = sh.getDataRange().getValues();
    let borradas = 0;
    for (let r = d.length - 1; r >= 1; r--) if (d[r][0] === id) { sh.deleteRow(r + 1); borradas++; }
    Logger.log(n + ': ' + borradas + ' filas borradas');
  });

  const hero = ss.getSheetByName('Perfiles_Hero');
  const hc = hero.getRange(1, 1, 1, hero.getLastColumn()).getValues()[0].indexOf('form_conocimientos');
  if (hc > -1) hero.deleteColumn(hc + 1);

  config.deleteRow(i + 1);
  const rc = headers.indexOf('recomendacion');
  if (rc > -1) config.deleteColumn(rc + 1);
  Logger.log('Config: perfil ' + id + ' borrado');

  const hoja = ss.getSheetByName(hojaNombre);
  if (hoja) {
    const hay = hoja.getLastRow() > 1 && hoja.getRange(2, 2, hoja.getLastRow() - 1, 1).getValues().some(r => String(r[0]).trim());
    if (hay) Logger.log('La pestaña "' + hojaNombre + '" tiene candidatos. NO la borré.');
    else { ss.deleteSheet(hoja); Logger.log('Pestaña "' + hojaNombre + '" borrada'); }
  }
  Logger.log('LISTO');
}
