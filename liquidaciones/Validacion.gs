var CAB_CTRL_ = ['Liquidación 50%', 'Liquidación 100%', 'Diferencia (h)', 'Fichadas del día (tarjeta reloj)', 'Qué se encontró', 'Decisión del jefe', 'Comentario del jefe (n.º de nota si corresponde)'];
var DEC_DIF_ = ['Conforme con la liquidación', 'Omisión justificada: pagar la franja', 'Corregir la planilla de horarios'];
var DEC_NOINF_ = ['Autorizo la extra no informada', 'No autorizo'];
var CONF_TIT_ = 'CONFORMIDAD DEL JEFE DE ENFERMERÍA';

function abrirValidacion_(cfg, filas) {
  var hoja = SpreadsheetApp.openById(cfg.idHorarios).getSheetByName('Reporte Extras Global');
  if (!hoja) return null;
  var datos = hoja.getDataRange().getDisplayValues(), x = compararEnfermeria_(datos, filas);
  if (!x) return null;
  var v = { hoja: hoja, datos: datos, dec: leerValJefe_(), conf: null };
  var hdr = datos[x.h], iv = hdr.indexOf(CAB_CTRL_[5]), ic = hdr.indexOf(CAB_CTRL_[6]);
  if (iv >= 0) {
    var leer = function (i, k) { var r = datos[i], d = [r[iv] || '', ic >= 0 ? r[ic] || '' : '']; if (d[0] || d[1]) v.dec[k] = d; else delete v.dec[k]; };
    for (var i in x.porFila) leer(i, x.porFila[i].k);
    for (var j in x.marca) leer(j, x.marca[j]);
  }
  var cb = colCtrl_(hdr) + CAB_CTRL_.length + 1;
  if (datos[0][cb] === CONF_TIT_ && datos[2] && datos[3]) v.dec.CONFORMIDAD = [datos[3][cb + 1] || '', datos[2][cb + 1] || '', (v.dec.CONFORMIDAD || [])[2] || ''];
  aplicarDecisiones_(filas, v.dec, x.E);
  return v;
}

function aplicarDecisiones_(filas, dec, E) {
  var hechos = {};
  filas.forEach(function (r) {
    var k = r[1] + '|' + r[3], d = dec[k];
    if (!d || !d[0]) return;
    var nota = d[1] ? ' (' + d[1] + ')' : '';
    if (d[0] === DEC_NOINF_[0] && /NO INFORMADA/.test(r[12]) && r[11] !== 'AUTORIZADA') {
      r[11] = 'AUTORIZADA'; r[12] = 'EXTRA NO INFORMADA · AUTORIZADA por el jefe de Enfermería' + nota;
    } else if (d[0] === DEC_DIF_[1] && E[k] && !hechos[k] && r[11] !== 'AUTORIZADA') {
      hechos[k] = 1; r[9] = E[k].h50; r[10] = E[k].h100; r[11] = 'AUTORIZADA';
      r[12] = 'OMISIÓN JUSTIFICADA por el jefe de Enfermería: se paga la franja ' + E[k].fr.join(' + ') + nota;
    }
  });
}

function cerrarValidacion_(cfg, filas, v) {
  var x = compararEnfermeria_(v.datos, filas, function (leg, d) {
    var m = marcasDia_(leg, d, cfg).map(function (z) { return z.t; });
    return m.length ? m.join(' · ') : 'sin fichadas';
  });
  var pend = pendientesJefe_(x, v.dec), nota = '';
  var conf = v.dec.CONFORMIDAD || ['', '', ''];
  if (/^s/i.test(conf[0]) && conf[1] && !pend.length && !conf[2]) conf[2] = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
  if (!/^s/i.test(conf[0]) || pend.length) conf[2] = '';
  v.dec.CONFORMIDAD = conf;
  try { marcarReporteEnf_(v.hoja, v.datos, x, v.dec, pend.length); nota = '\nLas diferencias están en «Reporte Extras Global» de la planilla de horarios para que el jefe de Enfermería decida.'; }
  catch (e) { nota = '\nNo pude escribir en la planilla de horarios (' + e.message + '): revisá que esta cuenta pueda editarla.'; }
  guardarValJefe_(v.dec);
  escribirActa_(cfg, x, v.dec, pend);
  var viejo = SpreadsheetApp.openById(ID_BASE).getSheetByName('CONTROL_ENFERMERIA');
  if (viejo) SpreadsheetApp.openById(ID_BASE).deleteSheet(viejo);
  var t = x.tot.map(function (n) { return Math.round(n * 10) / 10; });
  var nd = 0; for (var i in x.porFila) if (x.porFila[i].dif !== '') nd++;
  return '\nControl con Enfermería: ' + nd + ' días distintos del reporte y ' + x.solo.length + ' casos que no están en el reporte (extras no informadas). Enfermería ' + t[0] + ' h al 50 y ' + t[1] + ' al 100; liquidación ' + t[2] + ' y ' + t[3] + '.' +
    (conf[2] ? '\n✅ Conformidad del jefe dada (' + conf[1] + ', ' + conf[2] + ').' : '\n⏳ Falta la conformidad del jefe (' + pend.length + ' decisiones pendientes). Ver ACTA_CONFORMIDAD.') + nota;
}

function pendientesJefe_(x, dec) {
  var p = [];
  for (var i in x.porFila) if (x.porFila[i].dif !== '' && !(dec[x.porFila[i].k] || [])[0]) p.push(x.porFila[i].k);
  x.solo.forEach(function (s) { if (!(dec[s.k] || [])[0]) p.push(s.k); });
  return p;
}

function colCtrl_(hdr) {
  var c0 = hdr.indexOf(CAB_CTRL_[0]);
  if (c0 < 0) { c0 = hdr.length; while (c0 > 0 && !String(hdr[c0 - 1]).trim()) c0--; }
  return c0;
}

function leerValJefe_() {
  var h = SpreadsheetApp.openById(ID_BASE).getSheetByName('_VALID_JEFE'), val = {};
  if (h) h.getDataRange().getDisplayValues().forEach(function (r) { if (r[0] && r[0] !== 'clave') val[r[0]] = [r[1], r[2], r[3] || '']; });
  return val;
}

function guardarValJefe_(val) {
  var ss = SpreadsheetApp.openById(ID_BASE), h = ss.getSheetByName('_VALID_JEFE') || ss.insertSheet('_VALID_JEFE');
  var rows = [['clave', 'decisión', 'comentario', 'fecha']].concat(Object.keys(val).map(function (k) { return [k, val[k][0], val[k][1], val[k][2] || '']; }));
  h.clear(); h.getRange(1, 1, rows.length, 4).setNumberFormat('@').setValues(rows); h.hideSheet();
}

function marcarReporteEnf_(hoja, datos, x, dec, nPend) {
  var c0 = colCtrl_(datos[x.h]), fin = datos.length, NC = CAB_CTRL_.length, cb = c0 + NC + 1;
  while (fin > x.h + 1 && (!String(datos[fin - 1][x.c.leg]).trim() || String(datos[fin - 1][0]).trim() === MARCA_NO_REPORTE || /^EXTRAS NO INFORMADAS/.test(datos[fin - 1][0]))) fin--;
  if (hoja.getMaxColumns() < cb + 2) hoja.insertColumnsAfter(hoja.getMaxColumns(), cb + 2 - hoja.getMaxColumns());
  var maxR = hoja.getMaxRows();
  if (maxR > fin) hoja.getRange(fin + 1, 1, maxR - fin, hoja.getMaxColumns()).clearContent().clearFormat().clearDataValidations();
  if (maxR > x.h + 1) hoja.getRange(x.h + 2, c0 + 1, maxR - x.h - 1, NC).clearContent().clearFormat().clearDataValidations();
  hoja.getRange(x.h + 1, c0 + 1, 1, NC).setValues([CAB_CTRL_]).setFontWeight('bold').setBackground('#f4b183').setWrap(true).setVerticalAlignment('middle');
  var vals = [], fondos = [], listas = [];
  for (var i = x.h + 1; i < fin; i++) {
    var p = x.porFila[i];
    if (!p) { vals.push(new Array(NC).fill('')); fondos.push(new Array(NC).fill(null)); continue; }
    var d = p.dif === '' && !(dec[p.k] || [])[0] ? ['', ''] : (dec[p.k] || ['', '']);
    vals.push([p.n50, p.n100, p.dif, p.fich, p.mot, d[0], d[1]]);
    fondos.push(new Array(NC).fill(d[0] ? '#d9ead3' : p.dif === '' ? null : p.dif < 0 ? '#fbd9d9' : '#fff3b8'));
    if (p.dif !== '' || d[0]) listas.push([i + 1, DEC_DIF_]);
  }
  var n = fin - x.h - 1;
  if (n > 0) hoja.getRange(x.h + 2, c0 + 1, n, NC).setNumberFormat('@').setValues(vals).setBackgrounds(fondos).setWrap(true).setVerticalAlignment('top');
  if (x.solo.length) {
    var r0 = fin + 2;
    hoja.getRange(r0, 1).setValue('EXTRAS NO INFORMADAS Y DÍAS QUE NO ESTÁN EN ESTE REPORTE (las decide el jefe de Enfermería)').setFontWeight('bold').setFontColor('#9c2b00');
    var rows = x.solo.map(function (s) {
      var d = dec[s.k] || ['', ''], fila = new Array(c0 + NC).fill('');
      fila[0] = MARCA_NO_REPORTE; fila[x.c.leg] = s.leg; if (x.c.ape !== undefined) fila[x.c.ape] = s.ape; fila[x.c.dia] = s.dia;
      [s.n50, s.n100, s.dif === '' ? '' : s.dif, s.fich, s.mot, d[0], d[1]].forEach(function (z, j) { fila[c0 + j] = z; });
      return fila;
    });
    hoja.getRange(r0 + 1, 1, rows.length, c0 + NC).setNumberFormat('@').setValues(rows).setWrap(true).setVerticalAlignment('top')
      .setBackgrounds(x.solo.map(function (s) { return new Array(c0 + NC).fill((dec[s.k] || [])[0] ? '#d9ead3' : '#ffe8c7'); }));
    x.solo.forEach(function (s, j) { listas.push([r0 + 1 + j, s.noInf ? DEC_NOINF_ : DEC_DIF_]); });
  }
  listas.forEach(function (l) { hoja.getRange(l[0], c0 + 6).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(l[1], true).setAllowInvalid(false).build()); });
  var conf = dec.CONFORMIDAD || ['', '', ''];
  hoja.getRange(1, cb + 1, 6, 2).clearContent().clearDataValidations().setValues([
    [CONF_TIT_, ''],
    ['Decisiones pendientes', nPend ? nPend + ' (completá la columna «Decisión del jefe»)' : 'Ninguna'],
    ['Nombre y apellido del jefe', conf[1]],
    ['¿Das conformidad a la liquidación?', conf[0]],
    ['Conformidad registrada', conf[2] ? conf[2] + ' (al liquidar)' : 'Todavía no'],
    ['', 'Las filas en rojo se pagan menos que en este reporte; en amarillo, más; en naranja, extras que no están en la planilla. Verde = ya decidido.']
  ]).setWrap(true).setVerticalAlignment('top').setBackground('#eef3fb');
  hoja.getRange(1, cb + 1, 1, 2).setFontWeight('bold').setBackground('#1f4e79').setFontColor('white');
  hoja.getRange(4, cb + 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Sí', 'No'], true).build());
  hoja.getRange(2, cb + 2).setFontColor(nPend ? '#c62828' : '#2e7d32').setFontWeight('bold');
  hoja.setColumnWidth(c0 + 4, 230); hoja.setColumnWidth(c0 + 5, 360); hoja.setColumnWidth(c0 + 6, 220); hoja.setColumnWidth(c0 + 7, 230);
  hoja.setColumnWidth(cb + 1, 230); hoja.setColumnWidth(cb + 2, 300);
}

function escribirActa_(cfg, x, dec, pend) {
  var ss = SpreadsheetApp.openById(ID_BASE), h = ss.getSheetByName('ACTA_CONFORMIDAD') || ss.insertSheet('ACTA_CONFORMIDAD');
  var conf = dec.CONFORMIDAD || ['', '', ''], mes = MESES_[cfg.mes - 1][0];
  var items = [];
  for (var i in x.porFila) { var p = x.porFila[i]; if (p.dif !== '' || (dec[p.k] || [])[0]) items.push(p); }
  items = items.concat(x.solo);
  var t = x.tot.map(function (n) { return Math.round(n * 10) / 10; });
  var est = conf[2] ? 'CONFORMIDAD DADA por ' + conf[1] + ' el ' + conf[2] : pend.length ? 'PENDIENTE: faltan ' + pend.length + ' decisiones del jefe' : 'PENDIENTE: falta que el jefe marque «Sí» y su nombre';
  var cab = [['Acta de conformidad de horas extra · Enfermería · ' + mes.charAt(0).toUpperCase() + mes.slice(1) + ' ' + cfg.anio, '', '', '', '', '', '', ''],
    ['Estado', est, '', '', '', '', '', ''],
    ['Reporte de Enfermería', t[0] + ' h al 50 % y ' + t[1] + ' h al 100 %', '', '', '', '', '', ''],
    ['Liquidación por fichadas', t[2] + ' h al 50 % y ' + t[3] + ' h al 100 % (ya incluye lo que autorizó el jefe)', '', '', '', '', '', ''],
    ['', '', '', '', '', '', '', ''],
    ['Legajo', 'Apellido', 'Día', 'Fichadas del día', 'Qué se encontró', 'Horas que se pagan', 'Decisión del jefe', 'Comentario']];
  var rows = items.map(function (p) {
    var d = dec[p.k] || ['', ''], b = p.k.split('|');
    return [b[0], p.ape || (x.E[p.k] || {}).ape || '', b[1], p.fich, p.mot, p.n50 + p.n100, d[0] || 'PENDIENTE', d[1]];
  }).sort(function (a, b) { return String(a[1]).localeCompare(String(b[1])) || a[2] - b[2]; });
  var all = cab.concat(rows);
  h.clear();
  h.getRange(1, 1, all.length, 8).setNumberFormat('@').setValues(all).setWrap(true).setVerticalAlignment('top');
  h.getRange(1, 1, 1, 8).merge().setFontWeight('bold').setFontSize(14).setBackground('#1f4e79').setFontColor('white');
  h.getRange(2, 2, 1, 7).merge().setFontWeight('bold').setFontColor(conf[2] ? '#2e7d32' : '#c62828');
  h.getRange(3, 2, 1, 7).merge(); h.getRange(4, 2, 1, 7).merge();
  h.getRange(6, 1, 1, 8).setFontWeight('bold').setBackground('#434343').setFontColor('white');
  if (rows.length) h.getRange(7, 1, rows.length, 8).setBorder(true, true, true, true, true, true, '#c8c8c8', SpreadsheetApp.BorderStyle.SOLID)
    .setBackgrounds(rows.map(function (r) { return new Array(8).fill(r[6] === 'PENDIENTE' ? '#fbd9d9' : '#d9ead3'); }));
  [70, 140, 45, 220, 380, 80, 220, 220].forEach(function (w, i) { h.setColumnWidth(i + 1, w); });
}
