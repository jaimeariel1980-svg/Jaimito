const HOJA_LOG_FORM_ = 'Historial_Formularios';

function limpiarTxt_(v) { return String(v == null ? '' : v).replace(/[<>]/g, '').trim(); } // las celdas se escriben con formato texto: no se interpretan fórmulas
function numEd_(v, def, min, max) { const n = Number(String(v).replace(',', '.')); return isFinite(n) && String(v).trim() !== '' ? Math.min(max, Math.max(min, n)) : def; }

// id nuevo que no choque con los existentes (e1, e2… / v1… / c1…)
function ids_(items, pref) {
  const usados = {};
  items.forEach(i => { const id = limpiarTxt_(i.id); if (id) usados[id] = true; });
  let n = 0;
  return items.map(i => {
    const id = limpiarTxt_(i.id);
    if (id) return id;
    do { n++; } while (usados[pref + n]);
    usados[pref + n] = true;
    return pref + n;
  });
}

function logForm_(ss, quien, perfilId, hoja, antes) {
  let sh = ss.getSheetByName(HOJA_LOG_FORM_);
  if (!sh) {
    sh = ss.insertSheet(HOJA_LOG_FORM_);
    sh.getRange(1, 1, 1, 5).setValues([['fecha', 'quien', 'perfil', 'hoja', 'antes (JSON)']]).setFontWeight('bold').setBackground('#1A6645').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  sh.appendRow([new Date(), quien, perfilId, hoja, JSON.stringify(antes).slice(0, 49000)]);
}

// reemplaza las filas de un perfil en una hoja Perfiles_*; guarda lo anterior en Historial_Formularios
function reemplazarFilasPerfil_(ss, hoja, perfilId, filas, quien) {
  const sh = ss.getSheetByName(hoja);
  if (!sh) throw new Error('Falta la pestaña ' + hoja);
  const data = sh.getDataRange().getValues();
  const idx = [];
  for (let i = 1; i < data.length; i++) if (data[i][0] === perfilId) idx.push(i + 1);
  logForm_(ss, quien, perfilId, hoja, idx.map(r => data[r - 1]));
  for (let k = idx.length - 1; k >= 0; k--) sh.deleteRow(idx[k]);
  if (filas.length) sh.getRange(sh.getLastRow() + 1, 1, filas.length, filas[0].length).setNumberFormat('@').setValues(filas);
}

function guardarFormularioPerfil(pin, perfilId, p, quien) {
  pinValido_(pin);
  getConfig(perfilId);
  quien = limpiarTxt_(quien) || 'sin nombre';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);

    if (p.eval) {
      if (!p.eval.length) throw new Error('El screening necesita al menos un criterio.');
      const ids = ids_(p.eval, 'e');
      reemplazarFilasPerfil_(ss, 'Perfiles_Screening_Eval', perfilId, p.eval.map((e, k) => {
        const label = limpiarTxt_(e.label);
        if (!label) throw new Error('Hay un criterio sin texto.');
        return [perfilId, ids[k], label, numEd_(e.peso, 1, 0, 10)];
      }), quien);
    }
    if (p.val) {
      const ids = ids_(p.val, 'v');
      reemplazarFilasPerfil_(ss, 'Perfiles_Screening_Val', perfilId, p.val.map((v, k) => {
        const label = limpiarTxt_(v.label);
        const opts = (v.opts || []).map(o => limpiarTxt_(o).replace(/\|/g, '/')).filter(Boolean);
        if (!label) throw new Error('Hay una condición sin texto.');
        if (!opts.length) throw new Error('La condición «' + label + '» necesita opciones.');
        return [perfilId, ids[k], label, opts.join('|')];
      }), quien);
    }
    if (p.comps) {
      if (!p.comps.length) throw new Error('Hace falta al menos una competencia.');
      const ids = ids_(p.comps, 'c');
      reemplazarFilasPerfil_(ss, 'Perfiles_Competencias', perfilId, p.comps.map((c, k) => {
        const nombre = limpiarTxt_(c.nombre);
        if (!nombre) throw new Error('Hay una competencia sin nombre.');
        return [perfilId, ids[k], nombre, limpiarTxt_(c.descripcion), numEd_(c.peso, 1, 0, 10),
          (c.preguntas || []).map(q => limpiarTxt_(q).replace(/\|/g, '/')).filter(Boolean).join('|')];
      }), quien);
    }
    if (p.hero) {
      const sh = ss.getSheetByName('Perfiles_Hero');
      let data = sh.getDataRange().getValues();
      const fila = data.findIndex((r, i) => i > 0 && r[0] === perfilId) + 1;
      if (!fila) throw new Error('No encuentro el perfil en Perfiles_Hero.');
      const h = p.hero;
      if (h.criterio_avanza !== undefined) {
        const a = numEd_(h.criterio_avanza, 8.5, 0, 10), b = numEd_(h.criterio_bplus, 8, 0, 10), r = numEd_(h.criterio_reserva, 6.5, 0, 10);
        if (!(a >= b && b >= r)) throw new Error('Los umbrales tienen que ir de mayor a menor: A ≥ B+ ≥ B.');
        h.criterio_avanza = a; h.criterio_bplus = b; h.criterio_reserva = r;
      }
      const antes = {};
      Object.keys(h).forEach(k => {
        let col = data[0].indexOf(k) + 1;
        if (!col) { col = sh.getLastColumn() + 1; sh.getRange(1, col).setValue(k); data[0][col - 1] = k; }
        antes[k] = sh.getRange(fila, col).getValue();
        const v = (k.indexOf('criterio_') === 0) ? h[k] : limpiarTxt_(h[k]);
        const cel = sh.getRange(fila, col);
        if (typeof v === 'number') cel.setNumberFormat('0.0#').setValue(v); else cel.setNumberFormat('@').setValue(v);
      });
      logForm_(ss, quien, perfilId, 'Perfiles_Hero', antes);
    }
  } finally { lock.releaseLock(); }
  return { ok: true };
}
