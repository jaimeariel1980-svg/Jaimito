var ID_BASE = '1MErbdbNaRVZULcVQmWdyD9SRT4iVwBK7BckfTEqNgLQ';
var CONFIG_DEF = [
  ['mes', 10, 'Mes a liquidar (1-12)'],
  ['anio', 2026, 'Año'],
  ['id_horarios', '1mxpaPAUGhe-QEtc_asX_xX4bLaq9ujFrgXBlMDmSpOE', 'Planilla de horarios del mes (link o ID)'],
  ['id_fichadas', '', 'Fichadas del reloj del mes: una pestaña por día "1".."31" (link o ID)'],
  ['id_grh', '', 'Excel GRH con las autorizaciones del mes (link o ID). Vacío = sin consolidado GRH'],
  ['feriados_extra', '', 'Días feriado que no estén marcados FDO en la planilla, separados por coma (ej: 12, 13)'],
  ['no_laborables', '', 'Días no laborables (se tratan como sábado), separados por coma'],
  ['tope_extra', 30, 'Tope mensual de horas extra por agente (50% + 100%)'],
  ['hojas_excluidas', 'LISTAS, SUP, SUP SADOFE', 'Pestañas de la planilla que NO se liquidan (supervisores y listas)']
];

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Liquidaciones')
    .addItem('1) Preparar CONFIG', 'prepararConfig')
    .addItem('2) Liquidar horas extra', 'liquidarExtras')
    .addToUi();
}

function avisar_(msg) {
  Logger.log(msg);
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { }
}

function idDe_(v) {
  var s = String(v || '').trim(), m = s.match(/\/d\/([A-Za-z0-9_-]+)/);
  return m ? m[1] : s;
}

function prepararConfig() {
  var ss = SpreadsheetApp.openById(ID_BASE), h = ss.getSheetByName('CONFIG');
  if (!h) h = ss.insertSheet('CONFIG');
  var ex = {};
  h.getDataRange().getValues().forEach(function (r) { if (r[0]) ex[String(r[0]).trim()] = r[1]; });
  var filas = CONFIG_DEF.map(function (r) { return [r[0], ex[r[0]] !== undefined && ex[r[0]] !== '' ? ex[r[0]] : r[1], r[2]]; });
  h.clear();
  h.getRange(1, 1, 1, 3).setValues([['Clave', 'Valor', 'Qué es']]).setFontWeight('bold').setBackground('#434343').setFontColor('white');
  h.getRange(2, 1, filas.length, 3).setNumberFormat('@').setValues(filas);
  h.setFrozenRows(1); h.setColumnWidth(1, 150); h.setColumnWidth(2, 380); h.setColumnWidth(3, 520);
  avisar_('Hoja CONFIG lista. Completá lo que falte (columna Valor) y corré "Liquidar horas extra".');
}

function leerConfig_() {
  var h = SpreadsheetApp.openById(ID_BASE).getSheetByName('CONFIG');
  if (!h) { prepararConfig(); throw new Error('Se creó la hoja CONFIG: completala y volvé a correr.'); }
  var o = {};
  h.getDataRange().getDisplayValues().forEach(function (r) { if (r[0]) o[String(r[0]).trim()] = String(r[1]).trim(); });
  var dias = function (s) { var k = {}; String(s || '').split(/[,;\s]+/).forEach(function (x) { if (+x) k[+x] = 1; }); return k; };
  var c = {
    mes: +o.mes, anio: +o.anio, idHorarios: idDe_(o.id_horarios), idFichadas: idDe_(o.id_fichadas), idGrh: idDe_(o.id_grh),
    feriadosExtra: dias(o.feriados_extra), noLab: dias(o.no_laborables), tope: +o.tope_extra || 30,
    excluidas: String(o.hojas_excluidas || '').split(',').map(function (x) { return x.trim().toUpperCase(); }).filter(Boolean)
  };
  if (!(c.mes >= 1 && c.mes <= 12) || !(c.anio >= 2020)) throw new Error('CONFIG: revisá mes y año.');
  if (!c.idHorarios) throw new Error('CONFIG: falta id_horarios.');
  if (!c.idFichadas) throw new Error('CONFIG: falta id_fichadas (fichadas del reloj del mes).');
  return c;
}
