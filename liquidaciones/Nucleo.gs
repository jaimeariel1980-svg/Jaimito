var RX_FRANJA = /(\d{1,2})(?::(\d{2}))?\s*\/\s*(\d{1,2})(?::(\d{2}))?/g;

function parseTurno(s) {
  var m = String(s || '').replace(/^'/, '').trim().match(/^(\d{1,2})(?::(\d{2}))?\s*\/\s*(\d{1,2})(?::(\d{2}))?$/);
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
  real50 = Math.round(real50 * 10) / 10; real100 = Math.round(real100 * 10) / 10;
  var total = Math.round((real50 + real100) * 10) / 10, totalAut = aut50 + aut100;
  var reconv = Math.round(Math.min(Math.max(0, real50 - aut50), Math.max(0, aut100 - real100)) * 10) / 10;
  var pagar = Math.min(total, tope);
  var endoso = Math.max(0, Math.round((pagar - totalAut) * 10) / 10);
  var recorte = Math.max(0, Math.round((total - tope) * 10) / 10);
  var al = [];
  if (!hayGRH) al.push('NO FIGURA EN EL EXPEDIENTE (' + total + ' h sin autorización)');
  if (reconv > 0) al.push('reconversión ' + reconv + 'h');
  if (hayGRH && endoso > 0) al.push('endoso ' + endoso + 'h');
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
  ['AUTORIZADA (se paga)', '#ffffff', function (e) { return e === 'AUTORIZADA'; }],
  ['Extra NO autorizada (fichó de más sin franja cargada)', '#ffe8c7', function (e, c) { return /NO autorizada/.test(c); }],
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
