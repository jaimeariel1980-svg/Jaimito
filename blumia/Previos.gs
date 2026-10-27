/* «Ya se postuló antes»: busca al mismo candidato (DNI, mail, teléfono o nombre) en las otras búsquedas de Blumia. */
function claveNombre_(s) {
  const n = String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z ]/g, ' ').replace(/\s+/g, ' ').trim();
  return n.split(' ').length >= 2 && n.length >= 7 ? 'n:' + n.split(' ').sort().join(' ') : '';
}

function clavesCand_(nombre, tel, dni, mail) {
  const k = [];
  const d = String(dni || '').replace(/\D/g, ''); if (d.length >= 6) k.push('d:' + d);
  const m = String(mail || '').toLowerCase().trim(); if (m.indexOf('@') > 0) k.push('m:' + m);
  const t = String(tel || '').replace(/\D/g, ''); if (t.length >= 8) k.push('t:' + t.slice(-8));
  const n = claveNombre_(nombre); if (n) k.push(n);
  return k;
}

// Valores grandes en CacheService (cada clave admite hasta 100 KB): se parten en trozos de 30.000 caracteres
function cacheGrandeLeer_(k) {
  const c = CacheService.getScriptCache();
  const n = Number(c.get(k + '_n') || 0);
  if (!n) return null;
  const ks = []; for (let i = 0; i < n; i++) ks.push(k + '_' + i);
  const got = c.getAll(ks);
  let txt = '';
  for (let i = 0; i < n; i++) { if (got[k + '_' + i] == null) return null; txt += got[k + '_' + i]; }
  return txt;
}
function cacheGrandePoner_(k, txt, seg) {
  const o = {}; let n = 0;
  for (let i = 0; i < txt.length; i += 30000) { o[k + '_' + n] = txt.slice(i, i + 30000); n++; }
  o[k + '_n'] = String(n);
  CacheService.getScriptCache().putAll(o, seg);
}

// clave de candidato → en qué búsquedas está (de TODAS las búsquedas; cache de 10 minutos: leer las pestañas de todas tarda ~2 s)
function mapaPrevios_() {
  try { const t = cacheGrandeLeer_('prev_v2'); if (t) return JSON.parse(t); } catch (e) { }
  const mapa = {};
  listarPerfiles().forEach(p => {
    let cfg, sh;
    try { cfg = getConfig(p.id); sh = getSheet(cfg); } catch (e) { return; }
    if (!sh || sh.getLastRow() < 2) return;
    sh.getRange(2, 1, sh.getLastRow() - 1, 22).getValues().forEach(r => {
      const estado = String(r[8] || '').trim();
      if (!String(r[1]).trim() || /^Eliminado\|/.test(estado)) return;
      clavesCand_(r[1], r[2], r[3], r[21]).forEach(k => { (mapa[k] = mapa[k] || []).push({ p: p.id, busq: String(cfg.nombre || p.nombre || p.id), estado: estado }); });
    });
  });
  try { cacheGrandePoner_('prev_v2', JSON.stringify(mapa), 600); } catch (e) { }
  return mapa;
}

// devuelve un arreglo paralelo a las filas de datos (sin encabezado): [{busq, estado}, …]
function previosDeFilas_(perfilId, datos) {
  const out = datos.map(() => []);
  try {
    const mapa = mapaPrevios_();
    datos.forEach((r, i) => {
      const vistos = {};
      clavesCand_(r[1], r[2], r[3], r[21]).forEach(k => (mapa[k] || []).forEach(x => {
        if (x.p === perfilId) return;
        const id = x.busq + '|' + x.estado;
        if (!vistos[id] && out[i].length < 2) { vistos[id] = 1; out[i].push({ busq: x.busq, estado: x.estado }); }
      }));
    });
  } catch (e) { }
  return out;
}
