function jornadaFichada_(a, d, cfg) {
  if (esCola_(a, d)) return null;
  var t = turnoCompleto_(a, d), prevFin = -Infinity, prox = Infinity;
  bloquesTurno_(a, d).forEach(function (b) {
    var ini = t ? t.ini : 0, fin = t ? t.fin : 1440;
    if (b.fin <= ini && b.fin > prevFin) prevFin = b.fin;
    if (b.ini >= fin && b.ini < prox) prox = b.ini;
  });
  if (!t) {
    var S = marcasAbs_(a.leg, d, Math.max(prevFin + 1, 0), Math.min(prox - 240, 2160), cfg).filter(function (x, i, arr) { return i > 0 || x.m < 1440; });
    if (!S.length || S[0].m >= 1440) return null;
    var rs = { t: null, ent: S[0].t };
    if (S.length === 1) { rs.estado = 'omision'; rs.nota = 'falta la salida'; return rs; }
    rs.estado = 'ok'; rs.sal = S[S.length - 1].t; rs.ri = S[0].m; rs.rf = S[S.length - 1].m;
    return rs;
  }
  var M = marcasAbs_(a.leg, d, Math.max(t.ini - 480, prevFin + 1), Math.min(t.fin + 480, prox - 240), cfg);
  var r = { t: t };
  if (!M.length) { r.estado = 'sin'; return r; }
  if (M[0].m > t.ini + 180) {
    var cerca = M.filter(function (x) { return x.m >= t.fin - 180; });
    if (!cerca.length) { r.estado = 'sin'; return r; }
    r.estado = 'omision'; r.sal = cerca[cerca.length - 1].t; r.nota = 'falta la entrada';
    return r;
  }
  r.ent = M[0].t;
  if (M.length === 1) { r.estado = 'omision'; r.nota = 'falta la salida'; return r; }
  r.sal = M[M.length - 1].t; r.estado = 'ok'; r.m0 = M[0].m; r.m1 = M[M.length - 1].m;
  r.ri = Math.max(t.ini, r.m0); r.rf = Math.min(t.fin, r.m1);
  return r;
}

function nocheAgenteDia_(a, d, cfg) {
  if (esCola_(a, d)) return null;
  var t = turnoCompleto_(a, d);
  if (!t || minNoche_(t.ini, t.fin) <= 0) return null;
  var r = jornadaFichada_(a, d, cfg);
  r.noche = minNoche_(t.ini, t.fin);
  if (r.estado === 'ok') r.min = r.rf > r.ri ? minNoche_(r.ri, r.rf) : 0;
  return r;
}

function procesarNocheAgente_(a, serv, cfg, filas, acum) {
  for (var d = 1; d <= _diasMes; d++) {
    var x = nocheAgenteDia_(a, d, cfg);
    if (!x || x.estado === 'sin') continue;
    var g = acum[a.leg] || (acum[a.leg] = { ape: a.ape, serv: {}, min: 0, jor: 0, omis: [] });
    g.serv[serv] = 1;
    var turno = a.cam[d] || a.hab[d] || '-', franja = minToHora(x.t.ini) + ' a ' + minToHora(x.t.fin);
    if (x.estado === 'omision') {
      g.omis.push(d);
      filas.push([serv, a.leg, a.ape, d, turno, x.ent || '', x.sal || '', franja, '', 0, 'OMISIÓN: ' + x.nota + ' (no se paga)']);
    } else {
      g.min += x.min; g.jor++;
      filas.push([serv, a.leg, a.ape, d, turno, x.ent, x.sal, franja, x.min ? horaAbs_(x.ri) + ' a ' + horaAbs_(x.rf) : '', Math.round(x.min / 6) / 10,
        x.min < x.noche ? 'OK · cumplió ' + Math.round(x.min / 6) / 10 + ' de ' + Math.round(x.noche / 6) / 10 + ' h nocturnas del turno' : 'OK']);
    }
    agregarFichas_(filas, filas.length - 1, a.leg, d, cfg);
  }
}

function liquidarNocturnidad() {
  var cfg = leerConfig_(), H = cargarHorarios_(cfg, true), filas = [], acum = {};
  H.hojas.forEach(function (h) { h.ags.forEach(function (a) { procesarNocheAgente_(a, h.nombre, cfg, filas, acum); }); });
  filas.sort(function (a, b) { return String(a[2]).localeCompare(String(b[2])) || a[3] - b[3]; });

  var ss = SpreadsheetApp.openById(ID_BASE);
  var cabD = ['Servicio', 'Legajo', 'Apellido', 'Día', 'Turno', 'Entrada', 'Salida', 'Turno (horario)', 'Franja nocturna contada', 'Hs nocturnas', 'Estado', 'Fichadas del día', 'Horario en tarjeta reloj'];
  var hd = escribir_(ss, 'NOCTURNIDAD_DETALLE', cabD, filas, 5, 9);
  if (filas.length) {
    hd.getRange(2, 1, filas.length, cabD.length).setBorder(true, true, true, true, true, true, '#c8c8c8', SpreadsheetApp.BorderStyle.SOLID)
      .setBackgrounds(filas.map(function (r) { return new Array(cabD.length).fill(/^OMISIÓN/.test(r[10]) ? '#fbd9d9' : '#ffffff'); }));
    hd.getRange(2, 10, filas.length, 1).setNumberFormat('0.0');
  }

  var res = Object.keys(acum).map(function (l) {
    var g = acum[l], hs = Math.round(g.min / 6) / 10;
    return [l, g.ape, Object.keys(g.serv).join(', '), hs, g.jor, g.omis.length, g.omis.join(', '), g.omis.length ? 'omisión en ' + g.omis.length + ' jornada(s): no se pagan' : 'OK'];
  }).sort(function (a, b) { return String(a[1]).localeCompare(String(b[1])); });
  var cab = ['Legajo', 'Apellido', 'Servicio', 'Hs nocturnas a pagar', 'Jornadas pagadas', 'Omisiones', 'Días con omisión', 'Alerta'];
  var h = escribir_(ss, 'NOCTURNIDAD', cab, res);
  if (res.length) {
    h.getRange(2, 1, res.length, cab.length).setBorder(true, true, true, true, true, true, '#c8c8c8', SpreadsheetApp.BorderStyle.SOLID)
      .setBackgrounds(res.map(function (r) { return new Array(cab.length).fill(r[5] ? '#fbd9d9' : '#ffffff'); }));
    h.getRange(2, 4, res.length, 1).setNumberFormat('0.0').setFontWeight('bold');
  }
  var tot = res.reduce(function (s, r) { return s + r[3]; }, 0), om = res.reduce(function (s, r) { return s + r[5]; }, 0);
  avisar_('NOCTURNIDAD lista: ' + res.length + ' agentes · ' + Math.round(tot * 10) / 10 + ' hs nocturnas' + (om ? ' · ' + om + ' omisiones (no se pagan, en rojo)' : '') +
    '\nDetalle por jornada en NOCTURNIDAD_DETALLE.' + (cfg.fichadasAuto ? '\nFichadas: ' + cfg.fichadasAuto : ''));
}
