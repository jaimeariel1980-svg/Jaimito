var _fich = {}, _ssFich = null, _diasMes = 31, _mes = 1, _anio = 2026;

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
      if (!out[leg] || mk.length > out[leg].marcas.length) out[leg] = { marcas: mk.map(function (x) { return x.m; }), txt: mk.map(function (x) { return x.t; }), nombre: String(d[r][1]).trim(), horario: String(d[r][5]).trim() };
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

function licVigente_(txt, d) {
  var m = String(txt || '').match(/al\s+(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return !m || new Date(+m[3], +m[2] - 1, +m[1]) >= new Date(_anio, _mes - 1, d);
}

function turnoDia_(a, d) {
  if (a.lic[d] && licVigente_(a.lic[d], d)) return null;
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
    filas[i].push('[' + (l.length ? l.join(' · ') : 'sin fichadas') + ']', horarioReloj_(leg, d, filas[i][4], cfg));
  }
}

function horarioReloj_(leg, d, turnoTxt, cfg) {
  var x = fichDia_(d, cfg)[leg], h = x ? x.horario : '';
  var m = String(h).match(/(\d{1,2}):(\d{2})\s*a\s*(\d{1,2}):(\d{2})/);
  if (!m) return h;
  var r = parseTurno(m[1] + ':' + m[2] + '/' + m[3] + ':' + m[4]), t = parseTurno(turnoTxt);
  if (!t) return h;
  return Math.abs(r.ini - t.ini) > 30 || Math.abs(r.fin - t.fin) > 30 ? h + ' ≠ planilla ' + turnoTxt : h;
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
  var eDesde = contiguo ? bIni - 240 : w.ini - 180, eHasta = contiguo ? bIni + 180 : w.ini + 180, fin = contiguo ? bFin : w.fin, prox = Infinity;
  bloquesTurno_(a, d).forEach(function (t) { if (t.ini >= fin + 30) prox = Math.min(prox, t.ini); });
  var M = marcasAbs_(a.leg, d, eDesde, Math.min(fin + 480, prox - 240), cfg);
  if (!M.length || M[0].m > eHasta) return { estado: 'sin' };
  var res = { ent: M[0].t, sal: M.length > 1 ? M[M.length - 1].t : '', n: M.length };
  if (M.length === 1) { res.estado = 'omision'; return res; }
  res.estado = 'ok'; res.intermedias = M.length - 2;
  res.tramos = tramosPares_([M[0], M[M.length - 1]], w.ini, w.fin);
  return res;
}

function textoTramos_(tr) {
  return tr.map(function (x) { return horaAbs_(x[0]) + ' a ' + horaAbs_(x[1]); }).join(' + ');
}

function horas_(tr) {
  return Math.round(tr.reduce(function (s, x) { return s + x[1] - x[0]; }, 0) / 6) / 10;
}

function procesarAgente_(a, serv, cfg, ctx, filas) {
  _mes = ctx.mes; _anio = ctx.anio;
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
          if (!pr.tramos.length) { filas.push(base.concat([pr.ent, pr.sal, txt, '', 0, 0, 'A REVISAR', 'las fichadas no cubren la franja (no se paga)'])); return; }
          var r = calcularExtraTramos(pr.tramos, cd), fz = textoTramos_(pr.tramos);
          if (pr.intermedias > 0) cond += (cond ? ' + ' : '') + pr.intermedias + ' fichada(s) intermedia(s), no cuentan: jornada completa';
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
    var ult = M[M.length - 1].m;
    var tr = tramosPares_([M[0], M[M.length - 1]], t.fin, Infinity);
    if (!tr.length || ult - t.fin < 60) continue;
    var decl = parseFranjas(a.ext[d + 1]).map(function (w) { return { ini: w.ini + 1440, fin: w.fin + 1440 }; });
    if (decl.some(function (w) { return w.ini < ult && w.fin > t.fin; })) continue;
    if (ident()) { filas.push(base.concat(['', '', '', '', '', '', 'REVISAR IDENTIDAD', 'fichada: ' + nf])); continue; }
    var rn = calcularExtraTramos(tr, cd);
    if (rn.h50 + rn.h100 <= 0) continue;
    filas.push(base.concat([M[0].t, M[M.length - 1].t, '', textoTramos_(tr), rn.h50, rn.h100, 'A REVISAR',
      'extra NO autorizada (no hay franja declarada)']));
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
  out.sort(function (a, b) { return tipoFila_(a) - tipoFila_(b) || String(a[2]).localeCompare(String(b[2])) || a[3] - b[3]; });

  var cab = ['Servicio', 'Legajo', 'Apellido', 'Día', 'Turno', 'Entrada', 'Salida', 'Aut. Jefatura', 'Franja realizada', 'Hs 50%', 'Hs 100%', 'Estado', 'Condición', 'Fichadas del día', 'Horario en tarjeta reloj',
             'Aut GRH 50%', 'Aut GRH 100%', 'Total real mes', 'A pagar (tope ' + cfg.tope + ')', 'Alerta mensual'];
  var h = escribir_(SpreadsheetApp.openById(ID_BASE), 'LIQUIDACION', cab, out, 5, 9);
  pintarLiquidacion_(h, out, cab.length);
  var cambios = registrarCambios_(hojas, cfg);
  var aut = filas.filter(function (r) { return r[11] === 'AUTORIZADA'; }).length;
  avisar_('LIQUIDACION lista: ' + out.length + ' filas (' + aut + ' autorizadas) · ' + Object.keys(mens).length + ' agentes' + (cfg.fichadasAuto ? '\nFichadas: ' + cfg.fichadasAuto : '') + (grh ? '' : '\nSin Excel GRH en CONFIG: no hay consolidado de autorizaciones.') + (cambios ? '\n' + cambios : ''));
}

function escribir_(ss, nombre, cab, filas, textoDesde, textoHasta) {
  var h = ss.getSheetByName(nombre) || ss.insertSheet(nombre);
  h.clear();
  if (textoDesde && filas.length) h.getRange(2, textoDesde, filas.length, textoHasta - textoDesde + 1).setNumberFormat('@');
  h.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#434343').setFontColor('white');
  if (filas.length) h.getRange(2, 1, filas.length, cab.length).setValues(filas);
  h.setFrozenRows(1);
  h.autoResizeColumns(1, cab.length);
  return h;
}

function pintarLiquidacion_(h, filas, nc) {
  h.setRowHeight(1, 32);
  h.getRange(1, 1, 1, nc).setVerticalAlignment('middle').setHorizontalAlignment('center').setWrap(true);
  if (!filas.length) return;
  var rg = h.getRange(2, 1, filas.length, nc), fondos = [], est = [];
  filas.forEach(function (r) {
    var t = tipoFila_(r), c = TIPOS_FILA[t][1];
    fondos.push(new Array(nc).fill(c));
    est.push([t === 0 ? '#c9ebcf' : c]);
  });
  rg.setBackgrounds(fondos).setFontSize(10).setVerticalAlignment('middle')
    .setBorder(true, true, true, true, true, true, '#c8c8c8', SpreadsheetApp.BorderStyle.SOLID);
  h.getRange(2, 12, filas.length, 1).setBackgrounds(est).setFontWeight('bold').setHorizontalAlignment('center');
  h.getRange(2, 4, filas.length, 8).setHorizontalAlignment('center');
  h.getRange(2, 10, filas.length, 2).setNumberFormat('0.0');
  var hr = h.getRange(2, 15, filas.length, 1), fc = hr.getValues().map(function (v) { return [/≠/.test(v[0]) ? '#c62828' : '#000000']; });
  hr.setFontColors(fc);
  for (var i = 1; i < filas.length; i++) {
    if (tipoFila_(filas[i]) !== tipoFila_(filas[i - 1])) h.getRange(i + 1, 1, 1, nc).setBorder(null, null, true, null, null, null, '#555555', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  }
  var ley = TIPOS_FILA.map(function (x) { return [x[0]]; }), col = nc + 2;
  h.getRange(1, col).setValue('Referencias').setFontWeight('bold');
  h.getRange(2, col, ley.length, 1).setValues(ley).setBackgrounds(TIPOS_FILA.map(function (x) { return [x[1]]; }))
    .setBorder(true, true, true, true, true, true, '#c8c8c8', SpreadsheetApp.BorderStyle.SOLID);
  h.setColumnWidth(col, 330);
}

function registrarCambios_(hojas, cfg) {
  var ss = SpreadsheetApp.openById(ID_BASE), per = cfg.mes + '/' + cfg.anio, ahora = {}, info = {};
  hojas.forEach(function (h) {
    h.ags.forEach(function (a) {
      for (var d in a.ext) { var k = h.nombre + '|' + a.leg + '|' + d; ahora[k] = String(a.ext[d]).trim(); info[k] = a.ape; }
    });
  });
  var foto = ss.getSheetByName('_FOTO_EXTRAS'), antes = null;
  if (foto && foto.getLastRow() > 1) {
    var v = foto.getDataRange().getValues();
    if (String(v[0][1]) === per) { antes = {}; v.slice(1).forEach(function (r) { antes[r[0]] = String(r[1]); info[r[0]] = info[r[0]] || r[2]; }); }
  }
  if (!foto) { foto = ss.insertSheet('_FOTO_EXTRAS'); foto.hideSheet(); }
  foto.clear();
  var filasFoto = Object.keys(ahora).map(function (k) { return [k, ahora[k], info[k]]; });
  foto.getRange(1, 1, 1, 3).setValues([['periodo', per, '']]);
  if (filasFoto.length) foto.getRange(2, 1, filasFoto.length, 3).setNumberFormat('@').setValues(filasFoto);
  if (!antes) return 'Primera foto de extras de ' + per + ' guardada: los cambios se van a ver desde la próxima corrida.';
  var dif = compararFotos_(antes, ahora);
  if (!dif.length) return 'Sin cambios en las extras cargadas desde la corrida anterior.';
  var hoy = new Date(); hoy.setHours(0, 0, 0, 0);
  var cab = ['Detectado', 'Servicio', 'Legajo', 'Apellido', 'Día', 'Antes', 'Ahora', 'Cambio', '¿Día ya pasado?'];
  var hc = ss.getSheetByName('CAMBIOS');
  if (!hc) {
    hc = ss.insertSheet('CAMBIOS');
    hc.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold').setBackground('#434343').setFontColor('white');
    hc.setFrozenRows(1);
  }
  var pasados = 0, cuando = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
  var filas = dif.map(function (c) {
    var p = c.k.split('|'), pasado = new Date(cfg.anio, cfg.mes - 1, +p[2]) < hoy;
    if (pasado) pasados++;
    return [cuando, p[0], p[1], info[c.k] || '', +p[2], c.antes, c.ahora, c.tipo, pasado ? 'SÍ' : 'no'];
  });
  var r0 = hc.getLastRow() + 1, rg = hc.getRange(r0, 1, filas.length, cab.length);
  hc.getRange(r0, 6, filas.length, 2).setNumberFormat('@');
  rg.setValues(filas).setBorder(true, true, true, true, true, true, '#c8c8c8', SpreadsheetApp.BorderStyle.SOLID)
    .setBackgrounds(filas.map(function (f) { return new Array(cab.length).fill(f[8] === 'SÍ' ? '#f8caca' : '#ffffff'); }));
  hc.autoResizeColumns(1, cab.length);
  return dif.length + ' cambio(s) en las extras desde la corrida anterior' + (pasados ? ', ' + pasados + ' en días YA PASADOS (ver CAMBIOS, en rojo)' : '') + '.';
}
