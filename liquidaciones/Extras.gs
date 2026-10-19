var _fich = {}, _ssFich = null, _diasMes = 31;

function fichDia_(n, cfg) {
  if (n < 1 || n > _diasMes) return {};
  if (_fich[n]) return _fich[n];
  if (!_ssFich) _ssFich = SpreadsheetApp.openById(cfg.idFichadas);
  var out = {}, hoja = _ssFich.getSheetByName(String(n));
  if (hoja) {
    var d = hoja.getDataRange().getDisplayValues();
    for (var r = 1; r < d.length; r++) {
      var leg = String(d[r][0]).trim();
      if (!/^\d+$/.test(leg)) continue;
      var mk = [];
      for (var c = 7; c <= 12; c++) { var t = String(d[r][c]).trim(), m = horaToMin(t); if (m !== null) mk.push({ m: m, t: t }); }
      mk.sort(function (a, b) { return a.m - b.m; });
      if (!out[leg] || mk.length > out[leg].marcas.length) out[leg] = { marcas: mk.map(function (x) { return x.m; }), txt: mk.map(function (x) { return x.t; }), nombre: String(d[r][1]).trim() };
    }
  }
  return (_fich[n] = out);
}

function leerHoja_(hoja) {
  var datos = hoja.getDataRange().getDisplayValues(), fh = -1, mejor = 0, r, c;
  for (r = 0; r < Math.min(datos.length, 15); r++) {
    var n = 0;
    for (c = 3; c < datos[r].length; c++) if (/^\d{1,2}$/.test(datos[r][c].trim()) && +datos[r][c] >= 1 && +datos[r][c] <= 31) n++;
    if (n > mejor) { mejor = n; fh = r; }
  }
  if (fh < 0) return null;
  var cols = {}, tc = 2;
  for (c = 0; c < datos[fh].length; c++) {
    var v = datos[fh][c].trim();
    if (/^tipo$/i.test(v)) tc = c;
    else if (c >= 3 && /^\d{1,2}$/.test(v) && +v >= 1 && +v <= 31) cols[+v] = c;
  }
  var ags = [], fdo = {}, cur = null, destino = { habitual: 'hab', cambio: 'cam', extra: 'ext', licencias: 'lic', licencia: 'lic' };
  for (r = fh + 1; r < datos.length; r++) {
    var tipo = String(datos[r][tc]).trim().toLowerCase();
    var key = destino[tipo];
    if (!key) continue;
    if (key === 'hab') {
      var leg = String(datos[r][0]).trim();
      if (!/^\d+$/.test(leg)) { cur = null; continue; }
      cur = { leg: leg, ape: String(datos[r][1]).trim(), hab: {}, cam: {}, ext: {}, lic: {} };
      ags.push(cur);
    }
    if (!cur) continue;
    for (var d in cols) {
      var t = String(datos[r][cols[d]]).trim();
      if (!t) continue;
      cur[key][d] = t;
      if (key === 'hab' && /^FDO$/i.test(t)) fdo[d] = 1;
    }
  }
  return { agentes: ags, fdo: fdo };
}

function leerGRH_(id) {
  var h = SpreadsheetApp.openById(id).getSheetByName('Planillas Sectores');
  if (!h) throw new Error('El Excel GRH no tiene la hoja "Planillas Sectores".');
  var datos = h.getDataRange().getValues(), out = {};
  datos.forEach(function (r) {
    var leg = String(r[0]).trim();
    if (/^\d+$/.test(leg) && String(r[2]).toUpperCase().indexOf('ENFERMER') >= 0) out[leg] = { a50: parseFloat(r[7]) || 0, a100: parseFloat(r[8]) || 0 };
  });
  return out;
}

function turnoDia_(a, d) {
  if (a.lic[d]) return null;
  var c = a.cam[d], t = c ? parseTurno(c) : null;
  if (c && !t) return null;
  return t || parseTurno(a.hab[d]);
}

function turnoCompleto_(a, d) {
  var t = turnoDia_(a, d);
  if (t && t.fin === 1440) { var n = turnoDia_(a, d + 1); if (n && n.ini === 0) t = { ini: t.ini, fin: 1440 + n.fin }; }
  return t;
}

function esCola_(a, d) {
  var t = turnoDia_(a, d), p = turnoDia_(a, d - 1);
  return !!(t && t.ini === 0 && p && p.fin === 1440);
}

function nombreFich_(leg, d, cfg) {
  var x = fichDia_(d, cfg)[leg];
  return x ? x.nombre : '';
}

function marcasDia_(leg, d, cfg) {
  var x = fichDia_(d, cfg)[leg];
  return x ? x.marcas.map(function (m, i) { return { m: m, t: x.txt ? x.txt[i] : minToHora(m) }; }) : [];
}

function marcasAbs_(leg, d, desde, hasta, cfg) {
  var out = [];
  [-1, 0, 1].forEach(function (o) {
    marcasDia_(leg, d + o, cfg).forEach(function (x) {
      var abs = x.m + o * 1440;
      if (abs >= desde && abs <= hasta) out.push({ m: abs, t: x.t + (o < 0 ? '(-1)' : o > 0 ? '(+1)' : '') });
    });
  });
  return out.sort(function (x, y) { return x.m - y.m; });
}

function cruzaMedianoche_(r) {
  var w = parseTurno(r[4]);
  if (w && w.fin >= 1440 || /\(\+1\)/.test(r[6])) return true;
  return [r[7], r[8]].some(function (x) {
    var m = String(x || '').match(/(\d\d:\d\d) a (\d\d:\d\d)/);
    return !!m && horaToMin(m[2]) <= horaToMin(m[1]);
  });
}

function agregarFichas_(filas, desde, leg, d, cfg) {
  var t = function (arr) { return arr.map(function (x) { return x.t; }); };
  var hoy = t(marcasDia_(leg, d, cfg));
  var ayer = t(marcasAbs_(leg, d, -720, -1, cfg)), man = t(marcasAbs_(leg, d, 1440, 2159, cfg));
  for (var i = desde; i < filas.length; i++) {
    var l = (/\(-1\)/.test(filas[i][5]) ? ayer : []).concat(hoy, cruzaMedianoche_(filas[i]) ? man : []);
    filas[i].push('[' + (l.length ? l.join(' · ') : 'sin fichadas') + ']');
  }
}

function bloquesTurno_(a, d) {
  return [[0, d], [-1440, d - 1], [1440, d + 1]].map(function (x) {
    var t = turnoCompleto_(a, x[1]);
    return t && { ini: t.ini + x[0], fin: t.fin + x[0] };
  }).filter(Boolean);
}

function tramosPares_(M, ini, fin) {
  var out = [];
  for (var k = 0; k + 1 < M.length; k += 2) {
    var i0 = Math.max(ini, M[k].m), f0 = Math.min(fin, M[k + 1].m);
    if (f0 > i0) out.push([i0, f0]);
  }
  return out;
}

function presencia_(a, d, w, cfg) {
  var bIni = w.ini, bFin = w.fin, contiguo = false;
  bloquesTurno_(a, d).forEach(function (t) {
    if (Math.abs(t.fin - w.ini) <= 30) { bIni = Math.min(bIni, t.ini); contiguo = true; }
    if (Math.abs(t.ini - w.fin) <= 30) { bFin = Math.max(bFin, t.fin); contiguo = true; }
  });
  var eDesde = contiguo ? bIni - 240 : w.ini - 180, eHasta = contiguo ? bIni + 180 : w.ini + 180;
  var M = marcasAbs_(a.leg, d, eDesde, (contiguo ? bFin : w.fin) + 240, cfg);
  if (!M.length || M[0].m > eHasta) return { estado: 'sin' };
  var res = { ent: M[0].t, sal: M.length > 1 ? M[M.length - 1].t : '', n: M.length };
  if (M.length === 1) { res.estado = 'omision'; return res; }
  if (M.length % 2) { res.estado = 'impar'; res.tramos = tramosPares_([M[0], M[M.length - 1]], w.ini, w.fin); return res; }
  res.estado = 'ok'; res.tramos = tramosPares_(M, w.ini, w.fin);
  return res;
}

function textoTramos_(tr) {
  return tr.map(function (x) { return horaAbs_(x[0]) + ' a ' + horaAbs_(x[1]); }).join(' + ');
}

function horas_(tr) {
  return Math.round(tr.reduce(function (s, x) { return s + x[1] - x[0]; }, 0) / 6) / 10;
}

function procesarAgente_(a, serv, cfg, ctx, filas) {
  var n0 = filas.length;
  for (var d = 1; d <= _diasMes; d++) {
    if (d > 1) agregarFichas_(filas, n0, a.leg, d - 1, cfg);
    n0 = filas.length;
    var cd = { anio: ctx.anio, mes: ctx.mes, dia: d, feriados: ctx.feriados, noLab: ctx.noLab };
    var fer = !!ctx.feriados[claveFecha_(new Date(ctx.anio, ctx.mes - 1, d))];
    var turnoTxt = a.cam[d] || a.hab[d] || '-', ext = a.ext[d], base = [serv, a.leg, a.ape, d, turnoTxt];
    var nf = nombreFich_(a.leg, d, cfg);
    var ident = function () { return nf && !apellidoCoincide(a.ape, nf); };

    if (ext) {
      var fr = parseFranjas(ext);
      if (ident()) { filas.push(base.concat(['', '', ext, '', '', '', 'REVISAR IDENTIDAD', 'fichada: ' + nf])); continue; }
      if (!fr.length) { filas.push(base.concat(['', '', ext, '', '', '', 'A REVISAR', 'franja ilegible'])); continue; }
      var bloques = bloquesTurno_(a, d).concat(parseFranjas(a.ext[d - 1]).map(function (w) { return { ini: w.ini - 1440, fin: w.fin - 1440 }; }));
      var hoy = marcasDia_(a.leg, d, cfg), hE = hoy.length ? hoy[0].t : '', hS = hoy.length > 1 ? hoy[hoy.length - 1].t : '';
      fr.forEach(function (w0) {
        var txt = minToHora(w0.ini) + ' a ' + minToHora(w0.fin), piezas = restar_(w0, bloques);
        if (!piezas.length) {
          filas.push(base.concat([hE, hS, txt, '', 0, 0, 'A REVISAR', 'la franja declarada cae dentro del turno o de otra extra (no es extra)']));
          return;
        }
        piezas.forEach(function (w) {
          var cond = fer ? 'feriado' : '', pr = presencia_(a, d, w, cfg);
          if (w.ini !== w0.ini || w.fin !== w0.fin) cond += (cond ? ' + ' : '') + 'se descontó lo que cae dentro del turno';
          if (pr.estado === 'sin') { filas.push(base.concat([hE, hS, txt, '', 0, 0, 'A REVISAR', 'no fichó la extra (no se paga)' + (cond ? ' · ' + cond : '')])); return; }
          if (pr.estado === 'omision') { filas.push(base.concat([pr.ent, '', txt, '', 0, 0, 'A REVISAR', 'omisión: una sola fichada, no corresponde extra'])); return; }
          if (pr.estado === 'impar') { filas.push(base.concat([pr.ent, pr.sal, txt, '', 0, 0, 'A REVISAR', pr.n + ' fichadas (impar): falta una. Entre la primera y la última serían ' + horas_(pr.tramos) + ' h, revisar a mano'])); return; }
          if (!pr.tramos.length) { filas.push(base.concat([pr.ent, pr.sal, txt, '', 0, 0, 'A REVISAR', 'las fichadas no cubren la franja (no se paga)'])); return; }
          var r = calcularExtraTramos(pr.tramos, cd), fz = textoTramos_(pr.tramos);
          if (r.total < w.fin - w.ini) cond += (cond ? ' + ' : '') + 'realizó menos que lo autorizado';
          if (r.h50 + r.h100 <= 0) { filas.push(base.concat([pr.ent, pr.sal, txt, fz, 0, 0, 'A REVISAR', 'menos de 1 h (no se paga)'])); return; }
          filas.push(base.concat([pr.ent, pr.sal, txt, fz, r.h50, r.h100, 'AUTORIZADA', cond]));
        });
      });
      continue;
    }

    if (fer || esCola_(a, d)) continue;
    var t = turnoCompleto_(a, d);
    if (!t) continue;
    var M = marcasAbs_(a.leg, d, t.ini - 240, t.fin + 480, cfg);
    if (M.length < 2 || M[0].m > t.ini + 180) continue;
    var ult = M[M.length - 1].m, impar = M.length % 2 === 1;
    var tr = tramosPares_(impar ? [M[0], M[M.length - 1]] : M, t.fin, Infinity);
    if (!tr.length || ult - t.fin < 60) continue;
    var decl = parseFranjas(a.ext[d + 1]).map(function (w) { return { ini: w.ini + 1440, fin: w.fin + 1440 }; });
    if (decl.some(function (w) { return w.ini < ult && w.fin > t.fin; })) continue;
    if (ident()) { filas.push(base.concat(['', '', '', '', '', '', 'REVISAR IDENTIDAD', 'fichada: ' + nf])); continue; }
    var rn = calcularExtraTramos(tr, cd);
    if (rn.h50 + rn.h100 <= 0) continue;
    filas.push(base.concat([M[0].t, M[M.length - 1].t, '', textoTramos_(tr), impar ? 0 : rn.h50, impar ? 0 : rn.h100, 'A REVISAR',
      'extra NO autorizada (no hay franja declarada)' + (impar ? ' · ' + M.length + ' fichadas (impar): serían ' + horas_(tr) + ' h' : '')]));
  }
  agregarFichas_(filas, n0, a.leg, _diasMes, cfg);
}

function liquidarExtras() {
  var cfg = leerConfig_();
  _fich = {}; _ssFich = null; _diasMes = new Date(cfg.anio, cfg.mes, 0).getDate();
  var excl = cfg.excluidas.concat(['LIQUIDACION', 'NOCTURNIDAD', 'CONFIG']);
  var hojas = [], fdo = {};
  SpreadsheetApp.openById(cfg.idHorarios).getSheets().forEach(function (h) {
    if (excl.indexOf(h.getName().trim().toUpperCase()) >= 0) return;
    var x = leerHoja_(h);
    if (!x) return;
    hojas.push({ nombre: h.getName(), ags: x.agentes });
    for (var d in x.fdo) fdo[d] = 1;
  });
  for (var e in cfg.feriadosExtra) fdo[e] = 1;
  var ctx = { anio: cfg.anio, mes: cfg.mes, feriados: {}, noLab: {} };
  for (var f in fdo) ctx.feriados[cfg.mes + '-' + f] = 1;
  for (var n in cfg.noLab) ctx.noLab[cfg.mes + '-' + n] = 1;

  var filas = [];
  hojas.forEach(function (h) { h.ags.forEach(function (a) { procesarAgente_(a, h.nombre, cfg, ctx, filas); }); });

  var grh = cfg.idGrh ? leerGRH_(cfg.idGrh) : null, mens = {};
  filas.forEach(function (r) {
    if (r[11] !== 'AUTORIZADA') return;
    var m = mens[r[1]] || (mens[r[1]] = { r50: 0, r100: 0 });
    m.r50 += +r[9] || 0; m.r100 += +r[10] || 0;
  });
  for (var leg in mens) {
    var g = grh && grh[leg];
    var c = consolidarGRH_(mens[leg].r50, mens[leg].r100, g ? g.a50 : 0, g ? g.a100 : 0, !!g, cfg.tope);
    mens[leg].g = g; mens[leg].c = c;
  }
  var out = filas.map(function (r) {
    var m = mens[r[1]];
    return r.concat(m ? [m.g ? m.g.a50 : '', m.g ? m.g.a100 : '', m.c.total, m.c.pagar, m.c.alerta] : ['', '', '', '', '']);
  });
  out.sort(function (a, b) { return String(a[2]).localeCompare(String(b[2])) || a[3] - b[3]; });

  var cab = ['Servicio', 'Legajo', 'Apellido', 'Día', 'Turno', 'Entrada', 'Salida', 'Aut. Jefatura', 'Franja realizada', 'Hs 50%', 'Hs 100%', 'Estado', 'Condición', 'Fichadas del día',
             'Aut GRH 50%', 'Aut GRH 100%', 'Total real mes', 'A pagar (tope ' + cfg.tope + ')', 'Alerta mensual'];
  escribir_(SpreadsheetApp.openById(ID_BASE), 'LIQUIDACION', cab, out, 5, 9);
  var aut = filas.filter(function (r) { return r[11] === 'AUTORIZADA'; }).length;
  avisar_('LIQUIDACION lista: ' + out.length + ' filas (' + aut + ' autorizadas) · ' + Object.keys(mens).length + ' agentes' + (cfg.fichadasAuto ? '\nFichadas: ' + cfg.fichadasAuto : '') + (grh ? '' : '\nSin Excel GRH en CONFIG: no hay consolidado de autorizaciones.'));
}

function escribir_(ss, nombre, cab, filas, textoDesde, textoHasta) {
  var h = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  h.clear();
  if (textoDesde && filas.length) h.getRange(2, textoDesde, filas.length, textoHasta - textoDesde + 1).setNumberFormat('@');
  h.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#434343').setFontColor('white');
  if (filas.length) h.getRange(2, 1, filas.length, cab.length).setValues(filas);
  h.setFrozenRows(1);
  h.autoResizeColumns(1, cab.length);
}
