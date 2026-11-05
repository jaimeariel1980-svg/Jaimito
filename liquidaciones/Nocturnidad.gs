function jornadaFichada_(a, d, cfg, sinPrev) {
  if (esCola_(a, d)) return null;
  var t = turnoCompleto_(a, d), prevFin = -Infinity, prox = Infinity;
  bloquesTurno_(a, d).forEach(function (b) {
    var ini = t ? t.ini : 0, fin = t ? t.fin : 1440;
    if (b.fin <= ini && b.fin > prevFin) prevFin = b.fin;
    if (b.ini >= fin && b.ini < prox) prox = b.ini;
  });
  if (!t) {
    var vec = bloquesTurno_(a, d).concat([[-1440, d - 1], [1440, d + 1]].reduce(function (s, o) { return s.concat(parseFranjas(a.ext[o[1]]).map(function (w) { return { ini: w.ini + o[0], fin: w.fin + o[0] }; })); }, []));
    prevFin = -Infinity; prox = Infinity;
    vec.forEach(function (b) { if (b.ini < 0 && b.fin > prevFin) prevFin = b.fin; if (b.ini > 0 && b.ini < prox) prox = b.ini; });
    if (!sinPrev && !turnoCompleto_(a, d - 1)) {
      var jp = jornadaFichada_(a, d - 1, cfg, true);
      if (jp && jp.estado === 'ok' && jp.rf >= 1440) prevFin = Math.max(prevFin, jp.rf - 1440 - 120);
    }
    var G = [], g = [];
    marcasAbs_(a.leg, d, Math.max(prevFin + 180, 0), Math.min(prox - 60, 2160), cfg).forEach(function (x) {
      if (g.length && x.m - g[g.length - 1].m > 840) { G.push(g); g = []; }
      g.push(x);
    });
    if (g.length) G.push(g);
    G = G.filter(function (x) { return x[0].m < 1440; });
    if (G.length > 1 && G[0].length === 1) G.shift();
    var S = G[0] || [];
    if (!S.length || S[0].m >= 1440) return null;
    var rs = { t: null, ent: S[0].t };
    if (S.length === 1) { rs.estado = 'omision'; rs.nota = 'falta la salida'; return rs; }
    rs.estado = 'ok'; rs.sal = S[S.length - 1].t; rs.ri = S[0].m; rs.rf = S[S.length - 1].m;
    return rs;
  }
  var M = marcasAbs_(a.leg, d, Math.max(t.ini - 480, prevFin + 1), Math.min(t.fin + 480, prox - 240), cfg);
  var r = { t: t };
  r.na = faltaReloj_(a.leg, d, t.ini, t.fin, cfg);
  if (!M.length) { r.estado = 'sin'; return r; }
  if (M[0].m > t.ini + 180 && !(M.length > 1 && M[M.length - 1].m > t.fin + 60)) {
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

function horasReales_(a, d, j) {
  var t = j.t, ex = [[-1440, d - 1], [0, d], [1440, d + 1]].reduce(function (s, o) { return s.concat(parseFranjas(a.ext[o[1]]).map(function (w) { return { ini: w.ini + o[0], fin: w.fin + o[0] }; })); }, []);
  if (!t) return { distinto: false, pz: [{ ini: j.ri, fin: j.rf }] };
  var e0 = t.ini, e1 = t.fin;
  ex.forEach(function (w) { if (w.ini < e0 && w.fin >= e0 - 30) e0 = w.ini; if (w.fin > e1 && w.ini <= e1 + 30) e1 = w.fin; });
  if (j.m0 < e0 - 60 || j.m1 > e1 + 60) return { distinto: true, pz: restar_({ ini: j.m0, fin: j.m1 }, ex.map(function (w) { return { ini: w.ini - TOLERANCIA, fin: w.fin + TOLERANCIA }; })).filter(function (p) { return p.fin - p.ini > TOLERANCIA; }) };
  return { distinto: false, pz: j.rf > j.ri ? [{ ini: j.ri, fin: j.rf }] : [] };
}

function nocheAgenteDia_(a, d, cfg) {
  if (esCola_(a, d)) return null;
  var t = turnoCompleto_(a, d);
  if (!t || minNoche_(t.ini, t.fin) <= 0) return null;
  var r = jornadaFichada_(a, d, cfg);
  r.noche = minNoche_(t.ini, t.fin);
  if (r.estado !== 'ok') return r;
  var hr = horasReales_(a, d, r);
  r.distinto = hr.distinto; r.min = hr.pz.reduce(function (s, p) { return s + minNoche_(p.ini, p.fin); }, 0);
  if (r.distinto) { r.ri = r.m0; r.rf = r.m1; }
  return r;
}

function procesarNocheAgente_(a, serv, cfg, filas, acum) {
  if (!enReloj_(a.leg, cfg)) {
    var dn = []; for (var x0 = 1; x0 <= _diasMes; x0++) { var tn = turnoCompleto_(a, x0); if (tn && !esCola_(a, x0) && minNoche_(tn.ini, tn.fin) > 0) dn.push(x0); }
    if (!dn.length) return;
    acum[a.leg] = { ape: a.ape, serv: {}, min: 0, jor: 0, omis: [], dist: [], na: dn };
    acum[a.leg].serv[serv] = 1;
    filas.push([serv, a.leg, a.ape, dn[0], a.hab[dn[0]] || '', '', '', '', '', 0, NO_RELOJ_ + ' (' + dn.length + ' turnos nocturnos en la planilla)', '[sin datos]', '']);
    return;
  }
  for (var d = 1; d <= _diasMes; d++) {
    var x = nocheAgenteDia_(a, d, cfg);
    var na = x && x.estado !== 'ok' && x.na && x.na.length;
    if (!x || (x.estado === 'sin' && !na)) continue;
    var g = acum[a.leg] || (acum[a.leg] = { ape: a.ape, serv: {}, min: 0, jor: 0, omis: [], dist: [], na: [] });
    g.serv[serv] = 1;
    var turno = a.cam[d] || a.hab[d] || '-', franja = minToHora(x.t.ini) + ' a ' + minToHora(x.t.fin);
    if (na) {
      g.na.push(d);
      filas.push([serv, a.leg, a.ape, d, turno, x.ent || '', x.sal || '', franja, '', 0, textoNA_(x.na)]);
    } else if (x.estado === 'omision') {
      g.omis.push(d);
      filas.push([serv, a.leg, a.ape, d, turno, x.ent || '', x.sal || '', franja, '', 0, 'OMISIÓN: ' + x.nota + ' (no se paga)']);
    } else {
      g.min += x.min; if (x.min > 0) g.jor++;
      if (x.distinto) g.dist.push(d);
      filas.push([serv, a.leg, a.ape, d, turno, x.ent, x.sal, franja, x.min ? horaAbs_(x.ri) + ' a ' + horaAbs_(x.rf) : '', Math.round(x.min / 6) / 10,
        x.distinto ? 'HORARIO DISTINTO AL PROGRAMADO (sin cambio cargado): se pagan ' + Math.round(x.min / 6) / 10 + ' h nocturnas de lo fichado' : x.min < x.noche ? 'OK · cumplió ' + Math.round(x.min / 6) / 10 + ' de ' + Math.round(x.noche / 6) / 10 + ' h nocturnas del turno' : 'OK']);
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
      .setBackgrounds(filas.map(function (r) { return new Array(cabD.length).fill(/^SIN DATOS/.test(r[10]) ? '#d0e0e3' : /^OMISIÓN/.test(r[10]) ? '#fbd9d9' : /^HORARIO DISTINTO/.test(r[10]) ? '#fff3b8' : '#ffffff'); }));
    hd.getRange(2, 10, filas.length, 1).setNumberFormat('0.0');
  }

  var res = Object.keys(acum).map(function (l) {
    var g = acum[l], hs = Math.round(g.min / 6) / 10;
    var al = [];
    if (g.omis.length) al.push('omisión en ' + g.omis.length + ' jornada(s): no se pagan');
    if (g.na.length) al.push(g.jor || g.omis.length ? 'SIN DATOS DEL RELOJ (#N/A en la tarjeta) los días ' + g.na.join(', ') + ': no se pagan hasta tener el registro' : 'SIN DATOS DEL RELOJ: no figura en la tarjeta reloj en todo el mes (revisar el legajo)');
    if (g.dist.length) al.push('horario distinto al programado los días ' + g.dist.join(', ') + ' (cargar el cambio en la planilla)');
    return [l, g.ape, Object.keys(g.serv).join(', '), hs, g.jor, g.omis.length, g.omis.join(', '), al.join(' · ') || 'OK'];
  }).sort(function (a, b) { return String(a[1]).localeCompare(String(b[1])); });
  var cab = ['Legajo', 'Apellido', 'Servicio', 'Hs nocturnas a pagar', 'Jornadas pagadas', 'Omisiones', 'Días con omisión', 'Alerta'];
  var h = escribir_(ss, 'NOCTURNIDAD', cab, res);
  if (res.length) {
    h.getRange(2, 1, res.length, cab.length).setBorder(true, true, true, true, true, true, '#c8c8c8', SpreadsheetApp.BorderStyle.SOLID)
      .setBackgrounds(res.map(function (r) { return new Array(cab.length).fill(/SIN DATOS/.test(r[7]) ? '#d0e0e3' : r[5] ? '#fbd9d9' : /horario distinto/.test(r[7]) ? '#fff3b8' : '#ffffff'); }));
    h.getRange(2, 4, res.length, 1).setNumberFormat('0.0').setFontWeight('bold');
  }
  ordenarPestanas_(ss);
  var tot = res.reduce(function (s, r) { return s + r[3]; }, 0), om = res.reduce(function (s, r) { return s + r[5]; }, 0);
  avisar_('NOCTURNIDAD lista: ' + res.length + ' agentes · ' + Math.round(tot * 10) / 10 + ' hs nocturnas' + (om ? ' · ' + om + ' omisiones (no se pagan, en rojo)' : '') +
    '\nDetalle por jornada en NOCTURNIDAD_DETALLE.' + (cfg.fichadasAuto ? '\nFichadas: ' + cfg.fichadasAuto : '') + notaMeses_());
}
