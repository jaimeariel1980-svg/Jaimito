function esFeriadoDia_(ctx, d) { return !!ctx.feriados[claveFecha_(new Date(ctx.anio, ctx.mes - 1, d))]; }

function feriadoAgente_(a, serv, cfg, ctx, filas, acum) {
  if (esSadofe_(a, ctx)) return;
  for (var F = 1; F <= _diasMes; F++) {
    if (!esFeriadoDia_(ctx, F)) continue;
    [[F - 1, 1440], [F, 0]].forEach(function (par) {
      var d = par[0], off = par[1];
      if (d < 1) return;
      var j = jornadaFichada_(a, d, cfg);
      if (!j || j.estado === 'sin') return;
      var ini = j.t ? j.t.ini : (j.ri !== undefined ? j.ri : 0), fin = j.t ? j.t.fin : (j.rf !== undefined ? j.rf : 0);
      if (Math.min(fin, off + 1440) <= Math.max(ini, off)) return;
      var g = acum[a.leg] || (acum[a.leg] = { ape: a.ape, serv: {}, min: 0, jor: 0, omis: [], rev: 0 });
      g.serv[serv] = 1;
      var turno = a.cam[d] || a.hab[d] || '-', franja = j.t ? minToHora(j.t.ini) + ' a ' + minToHora(j.t.fin) : 'sin turno (' + turno + ')';
      var base = [serv, a.leg, a.ape, F, turno + (d !== F ? ' (del ' + d + ')' : ''), j.ent || '', j.sal || '', franja];
      if (j.estado === 'omision') {
        g.omis.push(F);
        filas.push(base.concat(['', 0, 'OMISIÓN: ' + j.nota + ' (no se paga)']));
      } else {
        var ri = Math.max(j.ri, off), rf = Math.min(j.rf, off + 1440), min = Math.max(0, rf - ri), est = 'OK';
        if (!j.t && a.ext[F]) { est = 'REVISAR: sin turno ese día y con extra cargada (que no se pague dos veces)'; g.rev++; }
        else if (!j.t) est = 'OK · trabajó con ' + turno + ' (franco por feriado)';
        if (est.indexOf('REVISAR') !== 0) { g.min += min; g.jor++; }
        filas.push(base.concat([min ? horaAbs_(ri - off) + ' a ' + horaAbs_(rf - off) : '', Math.round(min / 6) / 10, est]));
      }
      agregarFichas_(filas, filas.length - 1, a.leg, d, cfg);
    });
  }
}

function liquidarFeriados() {
  var cfg = leerConfig_(), H = cargarHorarios_(cfg), ctx = H.ctx, filas = [], acum = {};
  var dias = [];
  for (var F = 1; F <= _diasMes; F++) if (esFeriadoDia_(ctx, F)) dias.push(F);
  if (!dias.length) { avisar_('No hay feriados en ' + cfg.mes + '/' + cfg.anio + ' (ni FDO en la planilla ni en CONFIG feriados_extra).'); return; }
  H.hojas.forEach(function (h) { h.ags.forEach(function (a) { feriadoAgente_(a, h.nombre, cfg, ctx, filas, acum); }); });
  filas.sort(function (a, b) { return String(a[2]).localeCompare(String(b[2])) || a[3] - b[3]; });
  var ss = SpreadsheetApp.openById(ID_BASE);
  var cabD = ['Servicio', 'Legajo', 'Apellido', 'Día feriado', 'Turno', 'Entrada', 'Salida', 'Turno (horario)', 'Horas contadas del feriado', 'Hs feriado', 'Estado', 'Fichadas del día', 'Horario en tarjeta reloj'];
  var hd = escribir_(ss, 'FERIADOS_DETALLE', cabD, filas, 5, 9);
  if (filas.length) {
    hd.getRange(2, 1, filas.length, cabD.length).setBorder(true, true, true, true, true, true, '#c8c8c8', SpreadsheetApp.BorderStyle.SOLID)
      .setBackgrounds(filas.map(function (r) { return new Array(cabD.length).fill(/^OMISIÓN/.test(r[10]) ? '#fbd9d9' : /^REVISAR/.test(r[10]) ? '#fff3b8' : '#ffffff'); }));
    hd.getRange(2, 10, filas.length, 1).setNumberFormat('0.0');
  }
  var res = Object.keys(acum).map(function (l) {
    var g = acum[l];
    return [l, g.ape, Object.keys(g.serv).join(', '), Math.round(g.min / 6) / 10, g.jor, g.omis.length, g.omis.join(', '), g.rev ? g.rev + ' a revisar' : (g.omis.length ? 'omisión: no se paga' : 'OK')];
  }).sort(function (a, b) { return String(a[1]).localeCompare(String(b[1])); });
  var cab = ['Legajo', 'Apellido', 'Servicio', 'Hs feriado a pagar', 'Jornadas pagadas', 'Omisiones', 'Días con omisión', 'Alerta'];
  var h = escribir_(ss, 'FERIADOS', cab, res);
  if (res.length) {
    h.getRange(2, 1, res.length, cab.length).setBorder(true, true, true, true, true, true, '#c8c8c8', SpreadsheetApp.BorderStyle.SOLID)
      .setBackgrounds(res.map(function (r) { return new Array(cab.length).fill(/revisar/.test(r[7]) ? '#fff3b8' : r[5] ? '#fbd9d9' : '#ffffff'); }));
    h.getRange(2, 4, res.length, 1).setNumberFormat('0.0').setFontWeight('bold');
  }
  var tot = res.reduce(function (s, r) { return s + r[3]; }, 0);
  avisar_('FERIADOS lista (días ' + dias.join(', ') + '): ' + res.length + ' agentes · ' + Math.round(tot * 10) / 10 + ' hs de feriado.\nDetalle por jornada en FERIADOS_DETALLE.');
}
