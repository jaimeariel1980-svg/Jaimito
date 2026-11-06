var CAB_CTRL_ = ['Liquidación 50%', 'Liquidación 100%', 'Diferencia (h)', 'Fichadas del día (tarjeta reloj)', 'Qué se encontró', 'Decisión del jefe', 'Comentario del jefe (n.º de nota si corresponde)'];
var DEC_DIF_ = ['Conforme con la liquidación', 'Omisión justificada: pagar la franja', 'Corregir la planilla de horarios', 'El reloj tiene el registro: pagar la franja'];
var DEC_NOINF_ = ['Autorizo la extra no informada', 'No autorizo'];
var CONF_TIT_ = 'CONFORMIDAD DEL JEFE DE ENFERMERÍA';

var HOJA_VAL_ = 'Extras - Validación del jefe';
var CAB_PLAN_ = ['Servicio', 'Legajo', 'Apellido', 'Día', 'Turno', 'Extra cargada en la planilla', 'Hs 50% (planilla)', 'Hs 100% (planilla)'];

function planillaExtras_(hojas, ctx) {
  var l = [];
  hojas.forEach(function (h) {
    h.ags.forEach(function (a) {
      for (var d in a.ext) {
        var cd = { anio: ctx.anio, mes: ctx.mes, dia: +d, feriados: ctx.feriados, noLab: ctx.noLab }, h50 = 0, h100 = 0;
        parseFranjas(a.ext[d]).forEach(function (w) { var r = calcularExtraTramos([[w.ini, w.fin]], cd); h50 += r.h50; h100 += r.h100; });
        l.push([h.nombre, a.leg, a.ape, +d, a.cam[d] || a.hab[d] || '', String(a.ext[d]), h50, h100]);
      }
    });
  });
  l.sort(function (a, b) { return String(a[0]).localeCompare(String(b[0])) || String(a[2]).localeCompare(String(b[2])) || a[3] - b[3]; });
  return [CAB_PLAN_].concat(l);
}

function abrirValidacion_(cfg, filas, hojas, ctx) {
  var ssH = SpreadsheetApp.openById(cfg.idHorarios), rep = planillaExtras_(hojas, ctx), x = compararEnfermeria_(rep, filas);
  if (!x) return null;
  var v = { ssH: ssH, rep: rep, dec: leerValJefe_() }, hoja = ssH.getSheetByName(HOJA_VAL_);
  if (hoja) {
    var datos = hoja.getDataRange().getDisplayValues(), hdr = datos[0] || [];
    var il = hdr.indexOf('Legajo'), id = hdr.indexOf('Día'), iv = hdr.indexOf(CAB_CTRL_[5]), ic = hdr.indexOf(CAB_CTRL_[6]);
    var fin = false;
    if (il >= 0 && id >= 0 && iv >= 0) datos.slice(1).forEach(function (r) {
      if (/^CORRECCIONES A LA PLANILLA/.test(r[0])) fin = true;
      if (fin || !/^\d+$/.test(String(r[il]).trim()) || !+r[id]) return;
      var k = String(r[il]).trim() + '|' + (+r[id]), d = [r[iv] || '', ic >= 0 ? r[ic] || '' : ''];
      if (d[0] || d[1]) v.dec[k] = d; else delete v.dec[k];
    });
    var cb = CAB_PLAN_.length + CAB_CTRL_.length + 1;
    if (datos[0] && datos[0][cb] === CONF_TIT_ && datos[2] && datos[3]) v.dec.CONFORMIDAD = [datos[3][cb + 1] || '', datos[2][cb + 1] || '', (v.dec.CONFORMIDAD || [])[2] || ''];
  }
  aplicarDecisiones_(filas, v.dec, x.E);
  return v;
}

function aplicarDecisiones_(filas, dec, E) {
  var hechos = {}, jefe = (dec.CONFORMIDAD || [])[1] || '';
  filas.forEach(function (r) {
    var k = r[1] + '|' + r[3], d = dec[k];
    if (!d || !d[0]) return;
    var info = ' · decisión: «' + d[0] + '»' + (d[1] ? ' · comentario: «' + d[1] + '»' : '') + (jefe ? ' · jefe: ' + jefe : '') + ' · antes: ' + r[12];
    if (d[0] === DEC_NOINF_[0] && /NO INFORMADA/.test(r[12]) && r[11] !== 'AUTORIZADA') {
      r[11] = 'AUTORIZADA'; r[12] = 'AUTORIZADA por el jefe de Enfermería' + info;
    } else if ((d[0] === DEC_DIF_[1] || d[0] === DEC_DIF_[3]) && E[k] && !hechos[k] && r[11] !== 'AUTORIZADA') {
      hechos[k] = 1; r[9] = E[k].h50; r[10] = E[k].h100; r[11] = 'AUTORIZADA';
      r[12] = 'AUTORIZADA por el jefe de Enfermería: se paga la franja ' + E[k].fr.join(' + ') + info;
    } else if (!/DECISIÓN DEL JEFE/.test(r[12])) {
      r[12] += ' · DECISIÓN DEL JEFE: «' + d[0] + '»' + (d[1] ? ' (' + d[1] + ')' : '') + (jefe ? ' — ' + jefe : '');
    }
  });
}

function cerrarValidacion_(cfg, filas, v) {
  var x = compararEnfermeria_(v.rep, filas, function (leg, d) {
    var m = marcasDia_(leg, d, cfg).map(function (z) { return z.t; }), x = fichDia_(d, cfg)[leg];
    return (m.length ? m.join(' · ') : 'sin fichadas') + (x && x.novedad ? ' · Novedad del reloj: ' + x.novedad : '');
  });
  var pend = pendientesJefe_(x, v.dec), nota = '';
  var conf = v.dec.CONFORMIDAD || ['', '', ''];
  if (/^s/i.test(conf[0]) && conf[1] && !pend.length && !conf[2]) conf[2] = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
  if (!/^s/i.test(conf[0]) || pend.length) conf[2] = '';
  v.dec.CONFORMIDAD = conf;
  try { escribirValidacion_(v.ssH, v.rep, x, v.dec, pend.length); try { limpiarReporteViejo_(v.ssH); } catch (e2) { } nota = '\nEl jefe de Enfermería decide en la pestaña «' + HOJA_VAL_ + '» de la planilla de horarios.'; }
  catch (e) { nota = '\nNo pude escribir en la planilla de horarios (' + e.message + '): revisá que esta cuenta pueda editarla.'; }
  guardarValJefe_(v.dec);
  try {
    var ssB = SpreadsheetApp.openById(ID_BASE), acta = ssB.getSheetByName('ACTA_CONFORMIDAD'), rs = ssB.getSheetByName('RESUMEN');
    if (acta) ssB.deleteSheet(acta);
    if (rs) rs.getRange(1, RESUMEN_CAB_.length + 2).setValue(conf[2] ? 'Conformidad del jefe: DADA por ' + conf[1] + ' el ' + conf[2] : 'Conformidad del jefe: PENDIENTE' + (pend.length ? ' (' + pend.length + ' decisiones)' : ' (falta «Sí» y su nombre)'))
      .setFontWeight('bold').setFontColor('white').setBackground(conf[2] ? '#2e7d32' : '#c62828');
  } catch (e3) { Logger.log('conformidad: ' + e3.message); }
  var viejo = SpreadsheetApp.openById(ID_BASE).getSheetByName('CONTROL_ENFERMERIA');
  if (viejo) SpreadsheetApp.openById(ID_BASE).deleteSheet(viejo);
  var t = x.tot.map(function (n) { return Math.round(n * 10) / 10; });
  var nd = 0; for (var i in x.porFila) if (x.porFila[i].dif !== '') nd++;
  return '\nValidación del jefe: ' + nd + ' extras cargadas con diferencia y ' + x.solo.length + ' extras no informadas. Planilla de horarios ' + t[0] + ' h al 50 y ' + t[1] + ' al 100; liquidación ' + t[2] + ' y ' + t[3] + '.' +
    (conf[2] ? '\n✅ Conformidad del jefe dada (' + conf[1] + ', ' + conf[2] + ').' : '\n⏳ Falta la conformidad del jefe (' + pend.length + ' decisiones pendientes).') + nota;
}

function pendientesJefe_(x, dec) {
  var p = [];
  for (var i in x.porFila) if (x.porFila[i].dif !== '' && !(dec[x.porFila[i].k] || [])[0]) p.push(x.porFila[i].k);
  x.solo.forEach(function (s) { if (!(dec[s.k] || [])[0]) p.push(s.k); });
  return p;
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

function limpiarReporteViejo_(ssH) {
  var h = ssH.getSheetByName('Reporte Extras Global');
  if (!h) return;
  var hdr = h.getRange(1, 1, 1, h.getMaxColumns()).getDisplayValues()[0], c0 = hdr.indexOf(CAB_CTRL_[0]);
  if (c0 >= 0) h.getRange(1, c0 + 1, h.getMaxRows(), h.getMaxColumns() - c0).clearContent().clearFormat().clearDataValidations();
}

function escribirValidacion_(ssH, rep, x, dec, nPend) {
  var h = ssH.getSheetByName(HOJA_VAL_) || ssH.insertSheet(HOJA_VAL_, 0), NP = CAB_PLAN_.length, NC = CAB_CTRL_.length, NT = NP + NC, cb = NT + 1;
  var filas = [CAB_PLAN_.concat(CAB_CTRL_)], fondos = [new Array(NT).fill(null)], listas = [];
  rep.slice(1).forEach(function (r, i) {
    var p = x.porFila[i + 1], d = p && (p.dif !== '' || (dec[p.k] || [])[0]) ? (dec[p.k] || ['', '']) : ['', ''];
    filas.push(r.concat(p ? [p.n50, p.n100, p.dif, p.fich, p.mot, d[0], d[1]] : ['', '', '', '', '', '', '']));
    fondos.push(new Array(NT).fill(!p || (p.dif === '' && !d[0]) ? null : d[0] ? '#d9ead3' : p.dif < 0 ? '#fbd9d9' : '#fff3b8'));
    if (p && (p.dif !== '' || d[0])) listas.push([filas.length, DEC_DIF_]);
  });
  if (x.solo.length) {
    filas.push(new Array(NT).fill('')); fondos.push(new Array(NT).fill(null));
    var t = new Array(NT).fill(''); t[0] = 'EXTRAS NO INFORMADAS (fichó de más y no está cargado en la planilla): las decide el jefe de Enfermería';
    filas.push(t); fondos.push(new Array(NT).fill('#9c2b00'));
    x.solo.forEach(function (s) {
      var d = dec[s.k] || ['', ''];
      filas.push([s.serv, s.leg, s.ape, s.dia, s.turno, '(no cargada)', '', ''].concat([s.n50, s.n100, s.dif === '' ? '' : s.dif, s.fich, s.mot, d[0], d[1]]));
      fondos.push(new Array(NT).fill(d[0] ? '#d9ead3' : '#ffe8c7'));
      listas.push([filas.length, s.noInf ? DEC_NOINF_ : DEC_DIF_]);
    });
  }
  h.clear(); h.getRange(1, 1, h.getMaxRows(), h.getMaxColumns()).clearDataValidations();
  if (h.getMaxColumns() < cb + 2) h.insertColumnsAfter(h.getMaxColumns(), cb + 2 - h.getMaxColumns());
  h.getRange(1, 1, filas.length, NT).setNumberFormat('@').setValues(filas).setBackgrounds(fondos).setWrap(true).setVerticalAlignment('top');
  h.getRange(1, 1, 1, NP).setFontWeight('bold').setBackground('#434343').setFontColor('white');
  h.getRange(1, NP + 1, 1, NC).setFontWeight('bold').setBackground('#f4b183');
  if (x.solo.length) h.getRange(filas.length - x.solo.length, 1, 1, NT).setFontWeight('bold').setFontColor('white');
  h.setFrozenRows(1);
  listas.forEach(function (l) { h.getRange(l[0], NP + 6).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(l[1], true).setAllowInvalid(false).build()); });
  var conf = dec.CONFORMIDAD || ['', '', ''];
  h.getRange(1, cb + 1, 7, 2).setValues([
    [CONF_TIT_, ''],
    ['Decisiones pendientes', nPend ? nPend + ' (completá la columna «Decisión del jefe» en las filas de color)' : 'Ninguna'],
    ['Nombre y apellido del jefe', conf[1]],
    ['¿Das conformidad a la liquidación?', conf[0]],
    ['Conformidad registrada', conf[2] ? conf[2] + ' (al liquidar)' : 'Todavía no'],
    ['Cómo leer los colores', 'Sin color: coincide con lo cargado. Rojo: se paga menos que lo cargado. Amarillo: se paga más. Naranja: extra que no está cargada en la planilla. Verde: ya decidido por vos.'],
    ['Cuándo se aplica', 'Lo que decidas se aplica la próxima vez que se liquide.']
  ]).setWrap(true).setVerticalAlignment('top').setBackground('#eef3fb').setNumberFormat('@');
  h.getRange(1, cb + 1, 1, 2).setFontWeight('bold').setBackground('#1f4e79').setFontColor('white');
  h.getRange(4, cb + 2).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Sí', 'No'], true).build());
  h.getRange(2, cb + 2).setFontColor(nPend ? '#c62828' : '#2e7d32').setFontWeight('bold');
  [70, 60, 130, 40, 70, 110, 70, 70, 70, 70, 70, 210, 340, 210, 220, 20, 220, 300].forEach(function (w, i) { h.setColumnWidth(i + 1, w); });
}

