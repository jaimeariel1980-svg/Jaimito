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

// devuelve un arreglo paralelo a las filas de datos (sin encabezado): [{busq, estado}, …]
function previosDeFilas_(perfilId, datos) {
  const out = datos.map(() => []);
  try {
    const mapa = {};
    listarPerfiles().forEach(p => {
      if (p.id === perfilId) return;
      let cfg, sh;
      try { cfg = getConfig(p.id); sh = getSheet(cfg); } catch (e) { return; }
      if (!sh || sh.getLastRow() < 2) return;
      sh.getRange(2, 1, sh.getLastRow() - 1, 22).getValues().forEach(r => {
        const estado = String(r[8] || '').trim();
        if (!String(r[1]).trim() || /^Eliminado\|/.test(estado)) return;
        clavesCand_(r[1], r[2], r[3], r[21]).forEach(k => { (mapa[k] = mapa[k] || []).push({ busq: String(cfg.nombre || p.nombre || p.id), estado: estado }); });
      });
    });
    datos.forEach((r, i) => {
      const vistos = {};
      clavesCand_(r[1], r[2], r[3], r[21]).forEach(k => (mapa[k] || []).forEach(x => { const id = x.busq + '|' + x.estado; if (!vistos[id] && out[i].length < 2) { vistos[id] = 1; out[i].push(x); } }));
    });
  } catch (e) { }
  return out;
}
