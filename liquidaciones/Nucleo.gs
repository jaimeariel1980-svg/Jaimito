var RX_FRANJA = /(\d{1,2})(?:[:.](\d{2}))?\s*\/\s*(\d{1,2})(?:[:.](\d{2}))?/g;

function ampm_(s) {
  var h = function (x, ap) { x = +x % 12; return /p/i.test(ap) ? x + 12 : x; };
  return String(s || '').replace(/(\d{1,2})(?:[:.](\d{2}))?\s*([ap])\.?\s*m\.?\s*a\s*(\d{1,2})(?:[:.](\d{2}))?\s*([ap])\.?\s*m\.?/i, function (x, h1, m1, a1, h2, m2, a2) {
    return h(h1, a1) + ':' + (m1 || '00') + '/' + h(h2, a2) + ':' + (m2 || '00');
  }).replace(/^\s*C\/H\s*/i, '');
}

function parseTurno(s) {
  var m = ampm_(String(s || '').replace(/^'/, '').trim()).trim().match(/^(\d{1,2})(?:[:.](\d{2}))?\s*\/\s*(\d{1,2})(?:[:.](\d{2}))?$/);
  if (!m || +m[1] > 24 || +m[3] > 24) return null;
  var ini = +m[1] * 60 + (+m[2] || 0), fin = +m[3] * 60 + (+m[4] || 0);
  if (fin <= ini) fin += 1440;
  return { ini: ini, fin: fin };
}

function parseFranjas(s) {
  var out = [], m, t = String(s || '').replace(/'/g, '');
  RX_FRANJA.lastIndex = 0;
  while ((m = RX_FRANJA.exec(t))) { var w = parseTurno(m[0].replace(/\s+/g, '')); if (w) out.push(w); }
  return out;
}

function horaToMin(s) {
  var m = String(s || '').trim().match(/^(\d{1,2}):(\d{2})/);
  return m ? +m[1] * 60 + +m[2] : null;
}

function minToHora(m) {
  if (m === null || m === '' || m === undefined) return '';
  m = ((m % 1440) + 1440) % 1440;
  var h = Math.floor(m / 60), mm = m % 60;
  return (h < 10 ? '0' : '') + h + ':' + (mm < 10 ? '0' : '') + mm;
}

function claveFecha_(d) { return (d.getMonth() + 1) + '-' + d.getDate(); }

function esHoraAl100(fecha, hora, ctx) {
  var k = claveFecha_(fecha);
  if (ctx.feriados[k]) return true;
  var dow = ctx.noLab[k] ? 6 : fecha.getDay();
  return dow === 0 || hora >= 22 || hora < 6 || (dow === 6 && hora >= 13);
}

function calcularExtraTramos(tramos, ctx) {
  var total = tramos.reduce(function (s, t) { return s + Math.max(0, t[1] - t[0]); }, 0);
  var liquidar = Math.floor(total / 30) * 30;
  if (liquidar < 60) return { h50: 0, h100: 0, total: total };
  var min50 = 0, resto = liquidar, fechas = {};
  tramos.forEach(function (t) {
    for (var abs = t[0]; abs < t[1] && resto > 0; abs++, resto--) {
      var off = Math.floor(abs / 1440);
      var f = fechas[off] || (fechas[off] = new Date(ctx.anio, ctx.mes - 1, ctx.dia + off));
      if (!esHoraAl100(f, Math.floor((((abs % 1440) + 1440) % 1440) / 60), ctx)) min50++;
    }
  });
  var b50 = Math.round(min50 / 30) * 30;
  return { h50: b50 / 60, h100: (liquidar - b50) / 60, total: total };
}

function calcularExtra(exc, inicioAbs, ctx) {
  return calcularExtraTramos([[inicioAbs, inicioAbs + Math.max(0, exc)]], ctx);
}

function normaliza_(s) {
  return String(s || '').toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^A-Z ]/g, ' ').replace(/\s+/g, ' ').trim();
}

function bigramas_(s) {
  var b = []; s = ' ' + s + ' ';
  for (var i = 0; i < s.length - 1; i++) b.push(s.substr(i, 2));
  return b;
}

function dice(a, b) {
  var A = bigramas_(a), B = bigramas_(b).slice(), n = 0;
  A.forEach(function (x) { var i = B.indexOf(x); if (i >= 0) { n++; B.splice(i, 1); } });
  return A.length + B.length + n ? 2 * n / (A.length + B.length + n) : 0;
}

function apellidoCoincide(apeGrilla, nomFichada) {
  var ta = normaliza_(apeGrilla).split(' '), tf = normaliza_(nomFichada).split(' ');
  for (var i = 0; i < ta.length; i++) {
    if (ta[i].length < 3) continue;
    for (var j = 0; j < tf.length; j++) if (tf[j].length >= 3 && dice(ta[i], tf[j]) >= 0.6) return true;
  }
  return false;
}

function consolidarGRH_(real50, real100, aut50, aut100, hayGRH, tope) {
  var r1 = function (x) { return Math.round(x * 10) / 10; };
  real50 = r1(real50); real100 = r1(real100);
  var total = r1(real50 + real100), cap50 = aut50, cap100 = aut100, reconv = 0, dir = '';
  if (real50 > aut50 && real100 === 0 && aut100 > 0) { reconv = aut100; cap50 += aut100; cap100 = 0; dir = 'todo el 100 al 50'; }
  else if (real100 > aut100 && real50 === 0 && aut50 > 0) { reconv = aut50; cap100 += aut50; cap50 = 0; dir = 'todo el 50 al 100'; }
  var pagar = Math.min(total, tope), recorte = Math.max(0, r1(total - tope));
  var endoso = Math.max(0, r1(Math.max(0, real50 - cap50) + Math.max(0, real100 - cap100) - recorte));
  var al = [];
  if (!hayGRH) al.push('NO FIGURA EN EL EXPEDIENTE (' + total + ' h sin autorización)');
  if (reconv > 0) al.push('reconversión ' + reconv + 'h (' + dir + ')');
  if (hayGRH && endoso > 0) al.push('endoso ' + endoso + 'h (pedir a otro agente)');
  if (recorte > 0) al.push('EXCEDE ' + tope + ' (-' + recorte + 'h)');
  return { total: total, pagar: pagar, reconv: reconv, endoso: hayGRH ? endoso : 0, recorte: recorte, alerta: al.length ? al.join(' + ') : 'OK' };
}

function restar_(w, bloques) {
  var piezas = [{ ini: w.ini, fin: w.fin }];
  bloques.forEach(function (b) {
    var nuevas = [];
    piezas.forEach(function (p) {
      if (b.fin <= p.ini || b.ini >= p.fin) { nuevas.push(p); return; }
      if (b.ini > p.ini) nuevas.push({ ini: p.ini, fin: b.ini });
      if (b.fin < p.fin) nuevas.push({ ini: b.fin, fin: p.fin });
    });
    piezas = nuevas;
  });
  return piezas;
}

function horaAbs_(m) {
  if (m === null || m === undefined || m === '') return '';
  return minToHora(m) + (m < 0 ? '(-1)' : m >= 1440 ? '(+1)' : '');
}

var TIPOS_FILA = [
  ['AUTORIZADA (se paga)', '#ffffff', function (e, c) { return e === 'AUTORIZADA' && !/HORARIO DISTINTO|por el jefe/.test(c); }],
  ['AUTORIZADA por el jefe de Enfermería (extra no informada u omisión justificada)', '#d9ead3', function (e, c) { return e === 'AUTORIZADA' && /por el jefe/.test(c); }],
  ['AUTORIZADA con horario distinto al programado (se paga; cargar el cambio)', '#fde2a7', function (e, c) { return e === 'AUTORIZADA' && /HORARIO DISTINTO/.test(c); }],
  ['Extra no informada (fichó de más sin franja en la planilla): la decide el jefe', '#ffe8c7', function (e, c) { return /NO autorizada|NO INFORMADA/.test(c); }],
  ['Sin datos del reloj (#N/A en la tarjeta): no es omisión, pedir el registro', '#d0e0e3', function (e, c) { return /SIN DATOS DEL RELOJ/.test(c); }],
  ['Omisión / no fichó la extra', '#fbd9d9', function (e, c) { return /omisión|no fichó|no cubren|ilegible/.test(c); }],
  ['Fichadas dudosas (impares o con hueco): revisar a mano', '#fff3b8', function (e, c) { return /impar|revisar a mano/i.test(c); }],
  ['Franja dentro del turno (no es extra)', '#e9e0f7', function (e, c) { return /dentro del turno/.test(c); }],
  ['Menos de 1 h (no se paga)', '#ececec', function (e, c) { return /menos de 1 h/.test(c); }],
  ['Revisar identidad', '#d7e6fb', function (e) { return e === 'REVISAR IDENTIDAD'; }],
  ['Otro: revisar', '#f4f4f4', function () { return true; }]
];

function tipoFila_(r) {
  for (var i = 0; i < TIPOS_FILA.length; i++) if (TIPOS_FILA[i][2](r[11], String(r[12] || ''))) return i;
  return TIPOS_FILA.length - 1;
}

function compararFotos_(antes, ahora) {
  var out = [], k;
  for (k in ahora) if (!(k in antes)) out.push({ k: k, antes: '', ahora: ahora[k], tipo: 'Agregada' });
  else if (antes[k] !== ahora[k]) out.push({ k: k, antes: antes[k], ahora: ahora[k], tipo: 'Modificada' });
  for (k in antes) if (!(k in ahora)) out.push({ k: k, antes: antes[k], ahora: '', tipo: 'Quitada' });
  return out;
}

function parsearGRH_(datos) {
  var col = { leg: 0, area: 2, a50: 7, a100: 8 }, out = {};
  for (var r = 0; r < Math.min(datos.length, 150); r++) {
    var fila = datos[r].map(function (x) { return normaliza_(x); });
    var iL = fila.indexOf('LEGAJO');
    if (iL < 0) continue;
    col.leg = iL;
    fila.forEach(function (x, c) {
      if (/^AREA|^SECTOR/.test(x)) col.area = c;
      if (/50/.test(String(datos[r][c]))) col.a50 = c;
      if (/100/.test(String(datos[r][c]))) col.a100 = c;
    });
    break;
  }
  datos.forEach(function (r) {
    var leg = String(r[col.leg]).trim();
    if (/^\d+$/.test(leg) && normaliza_(r[col.area]).indexOf('ENFERMER') >= 0) out[leg] = { a50: parseFloat(r[col.a50]) || 0, a100: parseFloat(r[col.a100]) || 0, ape: String(r[col.leg + 1] || '').trim() };
  });
  return out;
}

function colorAlerta_(a) {
  a = String(a || '');
  if (/EXCEDE/.test(a)) return '#f8b4b4';
  if (/NO FIGURA|Figura en el expediente/.test(a)) return '#fbd9d9';
  if (/endoso/.test(a)) return '#ffe8c7';
  if (/reconversión/.test(a)) return '#fff3b8';
  return '#ffffff';
}

function minNoche_(a, b) {
  var n = 0;
  for (var m = a; m < b; m++) { var h = Math.floor((((m % 1440) + 1440) % 1440) / 60); if (h >= 21 || h < 6) n++; }
  return n;
}

var NO_LAB_FIJOS = { 2026: ['11-13', '12-7'] };

function noLabFijos_(anio, mes) {
  return (NO_LAB_FIJOS[anio] || []).filter(function (k) { return +k.split('-')[0] === mes; });
}

function motivoClaro_(rows, e) {
  if (!rows.length) return 'No figura como extra en la liquidación (no hay franja cargada en la planilla o no hay fichadas ese día).';
  return rows.map(function (r) {
    var c = String(r[12] || ''), fz = r[7] || 'la extra', fich = '';
    var hs = (+r[9] || 0) + (+r[10] || 0), fichado = r[5] && r[6] ? r[5] + ' a ' + r[6] : '';
    if (r[11] === 'AUTORIZADA') {
      if (/por el jefe/.test(c)) return c.replace(/^EXTRA NO INFORMADA · /, 'Extra no informada ') + ': se pagan ' + hs + ' h.';
      if (/HORARIO DISTINTO/.test(c)) return 'Trabajó en otro horario que el de la planilla (fichó ' + fichado + '). Se pagan ' + hs + ' h por lo fichado. Falta informar el cambio de turno.';
      if (/menos que lo autorizado/.test(c)) return 'Hizo menos horas que las autorizadas: la extra era ' + fz + ' y fichó ' + fichado + '. Se pagan ' + hs + ' h.';
      if (/:30/.test(fz)) {
        var w0 = parseTurno(fz.replace(' a ', '/')), m0 = horaToMin(r[5]), antes = w0 && m0 !== null && m0 <= w0.ini - 10 && m0 > w0.ini - 180;
        return 'La extra autorizada es ' + fz + ': son ' + hs + ' h' + (antes ? ' (entró antes, a las ' + r[5] + ', pero se paga solo lo autorizado)' : '') + '. La planilla de horarios dice ' + (e.h50 + e.h100) + ' h.';
      }
      return 'Se pagan ' + hs + ' h según las fichadas (' + fichado + ').';
    }
    if (/SIN DATOS DEL RELOJ/.test(c)) return c.replace('SIN DATOS DEL RELOJ: el', 'Sin datos del reloj: el') + '.';
    if (/no cubren/.test(c)) return 'Las fichadas (' + fichado + ') no llegan al horario de la extra ' + fz + ': no se paga.' + fich;
    if (/no fichó la extra/.test(c)) return 'No hay fichadas en el horario de la extra ' + fz + ': no se paga.' + fich;
    if (/omisión|una sola fichada/.test(c)) return 'Falta la entrada o la salida en el reloj: con una sola fichada no se paga.' + fich;
    if (/NO autorizada|NO INFORMADA/.test(c)) return 'Extra no informada: ' + c.replace(/^.*?NO INFORMADA · /, '').replace(/ \(no hay franja.*$/, '') + ' (fichó ' + fichado + '). No está cargada en la planilla: si la autorizás se pagan ' + (+r[9] || 0) + ' h al 50 % y ' + (+r[10] || 0) + ' h al 100 %.';
    if (/dentro del turno/.test(c)) return 'La extra ' + fz + ' cae dentro de su propio turno: no es extra.';
    if (/ilegible/.test(c)) return 'No se entiende la franja escrita en la planilla.';
    if (r[11] === 'REVISAR IDENTIDAD') return 'El nombre de la fichada no coincide con el agente: revisar.';
    return c || r[11];
  }).join(' · ');
}

var MARCA_NO_REPORTE = 'NO ESTÁ EN EL REPORTE';

function compararEnfermeria_(rep, filas, fichDe) {
  var num = function (x) { return parseFloat(String(x || '0').replace(',', '.')) || 0; }, h = -1, c = {};
  for (var i = 0; i < Math.min(rep.length, 20) && h < 0; i++) {
    var f = rep[i].map(function (x) { return normaliza_(x); });
    if (f.indexOf('LEGAJO') >= 0 && f.indexOf('DIA') >= 0) {
      h = i;
      f.forEach(function (x, j) {
        if (/^LIQUIDACION|^DIFERENCIA|^QUE SE|^VALIDACION|^COMENTARIO|^FICHADAS|^DECISION|^CONFORMIDAD/.test(x)) return; if (x === 'LEGAJO') c.leg = j; else if (x === 'APELLIDO') c.ape = j; else if (x === 'DIA') c.dia = j;
        else if (/EXTRA/.test(x)) c.fr = j; else if (/50/.test(rep[i][j])) c.a50 = j; else if (/100/.test(rep[i][j])) c.a100 = j;
      });
    }
  }
  if (h < 0) return null;
  var E = {}, N = {}, k, fila = {}, marca = {};
  rep.slice(h + 1).forEach(function (r, i) {
    var leg = String(r[c.leg]).trim();
    if (String(r[0]).trim() === MARCA_NO_REPORTE) { if (/^\d+$/.test(leg)) marca[h + 1 + i] = leg + '|' + (+r[c.dia]); return; }
    if (!/^\d+$/.test(leg) || !+r[c.dia]) return;
    k = leg + '|' + (+r[c.dia]);
    if (!(k in fila)) fila[k] = h + 1 + i;
    var e = E[k] || (E[k] = { leg: leg, ape: r[c.ape], dia: +r[c.dia], fr: [], h50: 0, h100: 0 });
    e.fr.push(r[c.fr]); e.h50 += num(r[c.a50]); e.h100 += num(r[c.a100]);
  });
  filas.forEach(function (r) {
    k = r[1] + '|' + r[3];
    var n = N[k] || (N[k] = { leg: r[1], ape: r[2], dia: r[3], h50: 0, h100: 0, rows: [] });
    if (r[11] === 'AUTORIZADA') { n.h50 += +r[9] || 0; n.h100 += +r[10] || 0; }
    n.rows.push(r);
  });
  var out = [], tot = [0, 0, 0, 0], porFila = {}, solo = [];
  Object.keys(E).concat(Object.keys(N).filter(function (x) { return !E[x]; })).forEach(function (k) {
    var e = E[k] || { h50: 0, h100: 0, fr: [] }, n = N[k] || { h50: 0, h100: 0, rows: [] }, b = E[k] || N[k];
    tot[0] += e.h50; tot[1] += e.h100; tot[2] += n.h50; tot[3] += n.h100;
    var igual = Math.abs(e.h50 - n.h50) < 0.01 && Math.abs(e.h100 - n.h100) < 0.01;
    var jefe = n.rows.some(function (r) { return /por el jefe/.test(r[12]); }), ver = !igual || jefe;
    var dif = Math.round((n.h50 + n.h100 - e.h50 - e.h100) * 10) / 10, mot = ver ? motivoClaro_(n.rows, e) : '', fi = '';
    if (ver) fi = n.rows.length && n.rows[0][13] ? String(n.rows[0][13]).replace(/^\[|\]$/g, '') + (/Novedad del reloj: /.test(n.rows[0][14]) ? ' · ' + String(n.rows[0][14]).replace(/^.*?Novedad del reloj: /, 'Novedad del reloj: ') : '') : fichDe ? fichDe(b.leg, b.dia) : '';
    if (k in fila) porFila[fila[k]] = { k: k, n50: n.h50, n100: n.h100, dif: igual ? '' : dif, mot: ver ? mot : 'Coincide', fich: fi };
    else if (ver || n.rows.some(function (r) { return /NO INFORMADA/.test(r[12]); })) solo.push({ k: k, leg: b.leg, ape: b.ape, dia: b.dia, serv: n.rows.length ? n.rows[0][0] : '', turno: n.rows.length ? n.rows[0][4] : '', n50: n.h50, n100: n.h100, dif: dif, mot: mot || motivoClaro_(n.rows, e), fich: fi || (n.rows.length && n.rows[0][13] ? String(n.rows[0][13]).replace(/^\[|\]$/g, '') : ''), noInf: n.rows.some(function (r) { return /NO INFORMADA/.test(r[12]); }) });
    if (!igual) out.push([b.leg, b.ape, b.dia, e.fr.join(' + '), e.h50, e.h100, n.h50, n.h100, dif, fi, mot, k]);
  });
  out.sort(function (a, b) { return String(a[1]).localeCompare(String(b[1])) || a[2] - b[2]; });
  return { filas: out, tot: tot, h: h, c: c, porFila: porFila, solo: solo, marca: marca, E: E };
}
