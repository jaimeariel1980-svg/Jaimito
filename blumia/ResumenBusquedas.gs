// Misma clasificación que el panel 📊 Resumen del kanban (ResumenPanel.html)
function resEstado_(estado) {
  const l = String(estado || '').trim().toLowerCase();
  if (/^eliminado/.test(l)) return null;
  if (l === 'contratado') return { lv: 5, t: 'cont' };
  if (/^descartad/.test(l)) {
    const lv = /post screening/.test(l) ? 1 : /en competencias/.test(l) ? 2 : /screening empresa/.test(l) ? 3 : /presencial empresa|en empresa/.test(l) ? 4 : 0;
    return { lv: lv, t: 'desc' };
  }
  if (/reserva/.test(l)) return { lv: /competencias/.test(l) ? 2 : /screening empresa/.test(l) ? 3 : /presencial empresa/.test(l) ? 4 : 1, t: 'res' };
  const lv = /competencias/.test(l) ? 2 : /screening empresa/.test(l) ? 3 : /presencial empresa|avanza a empresa|entrevista empresa|en cliente/.test(l) ? 4 : /screening/.test(l) ? 1 : 0;
  return { lv: lv, t: 'act' };
}

// una fila por búsqueda de la misma empresa que perfilId
function resumenBusquedas(perfilId) {
  const emp = String(getConfig(perfilId).empresa || '');
  const out = [];
  listarPerfiles().forEach(p => {
    let cfg, sh;
    try { cfg = getConfig(p.id); sh = getSheet(cfg); } catch (e) { return; }
    if (String(cfg.empresa || '') !== emp || !sh || sh.getLastRow() < 2) return;
    const r = { id: p.id, nombre: String(cfg.nombre || p.nombre || p.id), total: 0, activos: 0, l1: 0, l2: 0, l3: 0, l4: 0, contratados: 0, descartados: 0 };
    sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues().forEach(f => {
      if (!String(f[1]).trim()) return;
      const e = resEstado_(f[8]);
      if (!e) return;
      r.total++;
      if (e.t === 'act') r.activos++;
      if (e.t === 'cont') r.contratados++;
      if (e.t === 'desc') r.descartados++;
      for (let k = 1; k <= 4; k++) if (e.lv >= k) r['l' + k]++;
    });
    out.push(r);
  });
  return out;
}
