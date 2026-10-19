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
      for (var c = 7; c <= 12; c++) { var m = horaToMin(d[r][c]); if (m !== null) mk.push(m); }
      mk.sort(function (a, b) { return a - b; });
      if (!out[leg] || mk.length > out[leg].marcas.length) out[leg] = { marcas: mk, nombre: String(d[r][1]).trim() };
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

function corridas_(leg, d, desde, hasta, cfg) {
  var mk = [];
  [-1, 0, 1].forEach(function (o) {
    var x = fichDia_(d + o, cfg)[leg];
    if (x) x.marcas.forEach(function (m) { var abs = m + o * 1440; if (abs >= desde && abs <= hasta) mk.push(abs); });
  });
  mk.sort(function (a, b) { return a - b; });
  var out = [], cur = [];
  mk.forEach(function (m) {
    if (cur.length && m - cur[cur.length - 1] > 780) { out.push(cur); cur = []; }
    cur.push(m);
  });
  if (cur.length) out.push(cur);
  return out;
}

function nombreFich_(leg, d, cfg) {
  var x = fichDia_(d, cfg)[leg];
  return x ? x.nombre : '';
}

function marcasDia_(leg, d, cfg) {
  var x = fichDia_(d, cfg)[leg];
  return x ? x.marcas : [];
}

function cruzaMedianoche_(r) {
  var w = parseTurno(r[4]);
  if (w && w.fin >= 1440) return true;
  return [r[7], r[8]].some(function (x) {
    var m = String(x || '').match(/(\d\d:\d\d) a (\d\d:\d\d)/);
    return !!m && horaToMin(m[2]) <= horaToMin(m[1]);
  });
}

function agregarFichas_(filas, desde, leg, d, cfg) {
  var h = function (arr, suf) { return arr.map(function (m) { return minToHora(m) + suf; }); };
  var hoy = h(marcasDia_(leg, d, cfg), ''), man = h(marcasDia_(leg, d + 1, cfg).filter(function (m) { return m < 720; }), '(+1)');
  for (var i = desde; i < filas.length; i++) {
    var t = cruzaMedianoche_(filas[i]) ? hoy.concat(man) : hoy;
    filas[i].push('[' + (t.length ? t.join(' · ') : 'sin fichadas') + ']');
  }
}

function marcasAbs_(leg, d, desde, hasta, cfg) {
  var out = [];
  [-1, 0, 1].forEach(function (o) {
    marcasDia_(leg, d + o, cfg).forEach(function (m) { var abs = m + o * 1440; if (abs >= desde && abs <= hasta) out.push(abs); });
  });
  return out.sort(function (x, y) { return x - y; });
}

function bloquesTurno_(a, d) {
  return [[0, d], [-1440, d - 1], [1440, d + 1]].map(function (x) {
    var t = turnoCompleto_(a, x[1]);
    return t && { ini: t.ini + x[0], fin: t.fin + x[0] };
  }).filter(Boolean);
}

function presencia_(a, d, w, cfg) {
  var bIni = w.ini, bFin = w.fin, contiguo = false;
  bloquesTurno_(a, d).forEach(function (t) {
    if (Math.abs(t.fin - w.ini) <= 30) { bIni = Math.min(bIni, t.ini); contiguo = true; }
    if (Math.abs(t.ini - w.fin) <= 30) { bFin = Math.max(bFin, t.fin); contiguo = true; }
  });
  var eDesde = contiguo ? bIni - 240 : w.ini - 180, eHasta = contiguo ? bIni + 180 : w.ini + 180;
  var M = marcasAbs_(a.leg, d, eDesde, (contiguo ? bFin : w.fin) + 240, cfg);
  if (!M.length || M[0] > eHasta) return { estado: 'sin' };
  var res = { ent: M[0], sal: M.length > 1 ? M[M.length - 1] : null };
  if (M.length % 2) { res.estado = 'irregular'; res.nota = 'cantidad impar de fichadas (' + M.length + '): falta una'; return res; }
  for (var k = 1; k + 1 < M.length; k += 2) {
    if (Math.min(w.fin, M[k + 1]) - Math.max(w.ini, M[k]) > 30) { res.estado = 'irregular'; res.nota = 'según fichadas estuvo ausente de ' + horaAbs_(M[k]) + ' a ' + horaAbs_(M[k + 1]); return res; }
  }
  res.ri = Math.max(w.ini, M[0]); res.rf = Math.min(w.fin, M[M.length - 1]);
  res.dur = Math.max(0, res.rf - res.ri); res.estado = 'ok';
  return res;
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
      var hoy = marcasDia_(a.leg, d, cfg);
      fr.forEach(function (w0) {
        var txt = minToHora(w0.ini) + ' a ' + minToHora(w0.fin), piezas = restar_(w0, bloques);
        if (!piezas.length) {
          filas.push(base.concat([hoy.length ? minToHora(hoy[0]) : '', hoy.length > 1 ? minToHora(hoy[hoy.length - 1]) : '', txt, '', 0, 0, 'A REVISAR', 'la franja declarada cae dentro del turno o de otra extra (no es extra)']));
          return;
        }
        piezas.forEach(function (w) {
          var cond = fer ? 'feriado' : '', pr = presencia_(a, d, w, cfg);
          if (w.ini !== w0.ini || w.fin !== w0.fin) cond += (cond ? ' + ' : '') + 'se descontó lo que cae dentro del turno';
          if (pr.estado === 'sin') {
            filas.push(base.concat([hoy.length ? minToHora(hoy[0]) : '', hoy.length > 1 ? minToHora(hoy[hoy.length - 1]) : '', txt, '', 0, 0, 'A REVISAR', 'no fichó la extra (no se paga)' + (cond ? ' · ' + cond : '')]));
            return;
          }
          if (pr.estado === 'irregular') {
            filas.push(base.concat([horaAbs_(pr.ent), horaAbs_(pr.sal), txt, '', 0, 0, 'A REVISAR', 'fichadas irregulares, revisar a mano: ' + pr.nota]));
            return;
          }
          if (!pr.dur) { filas.push(base.concat([horaAbs_(pr.ent), horaAbs_(pr.sal), txt, '', 0, 0, 'A REVISAR', 'las fichadas no cubren la franja (no se paga)'])); return; }
          var r = calcularExtra(pr.dur, pr.ri, cd), fz = horaAbs_(pr.ri) + ' a ' + horaAbs_(pr.rf);
          if (pr.ri > w.ini || pr.rf < w.fin) cond += (cond ? ' + ' : '') + 'realizó menos que lo autorizado';
          if (r.h50 + r.h100 <= 0) { filas.push(base.concat([horaAbs_(pr.ent), horaAbs_(pr.sal), txt, fz, 0, 0, 'A REVISAR', 'menos de 1 h (no se paga)'])); return; }
          filas.push(base.concat([horaAbs_(pr.ent), horaAbs_(pr.sal), txt, fz, r.h50, r.h100, 'AUTORIZADA', cond]));
        });
      });
      continue;
    }

    if (fer || esCola_(a, d)) continue;
    var t = turnoCompleto_(a, d);
    if (!t) continue;
    var salida = null, entrada = null;
    corridas_(a.leg, d, t.ini - 240, t.fin + 480, cfg).forEach(function (c) {
      if (c.length >= 2 && c.some(function (m) { return m >= t.ini - 240 && m <= t.ini + 180; })) salida = c[c.length - 1]; entrada = c[0];
    });
    if (salida === null || salida - t.fin < 60) continue;
    var decl = parseFranjas(a.ext[d + 1]).map(function (w) { return { ini: w.ini + 1440, fin: w.fin + 1440 }; });
    if (decl.some(function (w) { return w.ini < salida && w.fin > t.fin; })) continue;
    if (ident()) { filas.push(base.concat(['', '', '', '', '', '', 'REVISAR IDENTIDAD', 'fichada: ' + nf])); continue; }
    var rn = calcularExtra(salida - t.fin, t.fin, cd);
    if (rn.h50 + rn.h100 <= 0) continue;
    filas.push(base.concat([horaAbs_(entrada), horaAbs_(salida), '', horaAbs_(t.fin) + ' a ' + horaAbs_(salida), rn.h50, rn.h100, 'A REVISAR', 'extra NO autorizada (no hay franja declarada)']));
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
