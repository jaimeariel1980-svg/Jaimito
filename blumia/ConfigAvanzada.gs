/* ============================================================
   ⚙ Configurar → «Datos de la búsqueda»: puesto y proceso, estado abierta/cerrada,
   lectura de CVs, palabras del asunto del mail y enlaces. Todo lee y escribe las mismas
   pestañas de la planilla (Config, Perfiles_Hero, Mail_Routing), así que lo que se cambia acá
   se ve en el sheet y al revés. Si el sheet cambió mientras se editaba, se avisa en vez de pisar.
   ============================================================ */

const HERO_EDIT_ = ['posicion', 'ubicacion', 'instancia', 'entrevistadoras'];
const CV_IA_OPC_ = ['', 'claude', 'openai'];

function estadoBusqueda_(cfg) { return String((cfg && cfg.estado_busqueda) || '').toUpperCase().trim() === 'CERRADA' ? 'CERRADA' : 'ABIERTA'; }

function mailPalabras_(ss, perfilId) {
  const sh = ss.getSheetByName('Mail_Routing');
  if (!sh) return '';
  const f = sh.getDataRange().getValues().slice(1).find(r => r[0] === perfilId);
  return f ? String(f[1] || '').trim() : '';
}

function enlacesBusqueda_(cfg) {
  const enlaces = {};
  try { enlaces.carpeta = DriveApp.getFolderById(cfg.folder_madre).getParents().next().getUrl(); } catch (e) { }
  try {
    const sh = SpreadsheetApp.openById(cfg.sheet_id).getSheetByName(cfg.hoja_nombre);
    enlaces.hoja = 'https://docs.google.com/spreadsheets/d/' + cfg.sheet_id + '/edit#gid=' + sh.getSheetId();
  } catch (e) { }
  return enlaces;
}

function configExtraLeer_(ss, cfg, perfilId) {
  const hero = {};
  HERO_EDIT_.forEach(k => hero[k] = '');
  const shH = ss.getSheetByName('Perfiles_Hero');
  if (shH) {
    const d = shH.getDataRange().getValues();
    const f = d.slice(1).find(r => r[0] === perfilId);
    if (f) HERO_EDIT_.forEach(k => { const c = d[0].indexOf(k); if (c >= 0) hero[k] = String(f[c] == null ? '' : f[c]); });
  }
  const enlaces = {}; // se completan aparte (configLazy_) para no frenar la apertura
  return { hero: hero, estado: estadoBusqueda_(cfg), cv_ia: CV_IA_OPC_.indexOf(String(cfg.cv_ia || '').toLowerCase().trim()) >= 0 ? String(cfg.cv_ia || '').toLowerCase().trim() : '', mail: mailPalabras_(ss, perfilId), enlaces: enlaces };
}

function choqueEdicion_(nombre, actual, nuevo, base) {
  if (base === undefined || base === null) return;
  if (String(actual) !== String(base) && String(nuevo) !== String(actual))
    throw new Error('«' + nombre + '» se cambió en la planilla mientras lo editabas (ahora dice: «' + String(actual).slice(0, 80) + '»). Cerrá Configurar y volvé a abrirlo para no pisar ese cambio.');
}

// x = {hero, estado, cv_ia, mail}; base = valores que se veían al abrir
function configExtraGuardar_(ss, perfilId, x, base, quien) {
  base = base || {};
  if (x.estado !== undefined || x.cv_ia !== undefined) {
    const sh = ss.getSheetByName('Config');
    const data = sh.getDataRange().getValues();
    const fila = data.findIndex((r, i) => i > 0 && r[0] === perfilId) + 1;
    const antes = {};
    const poner = (k, v, b) => {
      let col = data[0].indexOf(k) + 1;
      if (!col) { col = sh.getLastColumn() + 1; sh.getRange(1, col).setValue(k); data[0][col - 1] = k; }
      const cel = sh.getRange(fila, col), cur = cel.getValue();
      choqueEdicion_(k, cur, v, b);
      if (String(cur) === String(v)) return;
      antes[k] = cur;
      cel.setNumberFormat('@').setValue(v);
    };
    if (x.estado !== undefined) {
      const e = String(x.estado).toUpperCase();
      if (e !== 'ABIERTA' && e !== 'CERRADA') throw new Error('El estado de la búsqueda tiene que ser ABIERTA o CERRADA.');
      const c = data[0].indexOf('estado_busqueda');
      const cur = c >= 0 ? String(data[fila - 1][c] || '').toUpperCase() || 'ABIERTA' : 'ABIERTA';
      choqueEdicion_('estado de la búsqueda', cur, e, base.estado);
      if (cur !== e) poner('estado_busqueda', e);
    }
    if (x.cv_ia !== undefined) {
      const v = String(x.cv_ia).toLowerCase().trim();
      if (CV_IA_OPC_.indexOf(v) < 0) throw new Error('La lectura de CVs tiene que ser automática, Claude o ChatGPT.');
      poner('cv_ia', v, base.cv_ia);
    }
    if (Object.keys(antes).length) logForm_(ss, quien, perfilId, 'Config', antes);
  }

  if (x.hero) {
    const sh = ss.getSheetByName('Perfiles_Hero');
    const data = sh.getDataRange().getValues();
    const fila = data.findIndex((r, i) => i > 0 && r[0] === perfilId) + 1;
    if (!fila) throw new Error('No encuentro la búsqueda en Perfiles_Hero.');
    const antes = {};
    HERO_EDIT_.forEach(k => {
      if (x.hero[k] === undefined) return;
      let v = limpiarTxt_(x.hero[k]).slice(0, 200);
      if (k === 'posicion' && !v) throw new Error('El nombre del puesto no puede quedar vacío.');
      if (k === 'entrevistadoras') {
        const L = v.split(',').map(s => s.trim()).filter(Boolean);
        if (!L.length) throw new Error('Cargá al menos un entrevistador.');
        if (L.length > 12) throw new Error('Son demasiados entrevistadores (máximo 12).');
        v = L.join(', ');
      }
      let col = data[0].indexOf(k) + 1;
      if (!col) { col = sh.getLastColumn() + 1; sh.getRange(1, col).setValue(k); data[0][col - 1] = k; }
      const cel = sh.getRange(fila, col), cur = cel.getValue();
      choqueEdicion_(k, cur, v, (base.hero || {})[k]);
      if (String(cur) === v) return;
      antes[k] = cur;
      cel.setNumberFormat('@').setValue(v);
    });
    if (Object.keys(antes).length) logForm_(ss, quien, perfilId, 'Perfiles_Hero', antes);
  }

  if (x.mail !== undefined) {
    const claves = String(x.mail).split('|').map(s => limpiarTxt_(s)).filter(Boolean);
    claves.forEach(k => { if (k.length < 2) throw new Error('La palabra «' + k + '» es muy corta: coincidiría con casi cualquier asunto.'); });
    const v = claves.join('|');
    let sh = ss.getSheetByName('Mail_Routing');
    if (!sh) { sh = ss.insertSheet('Mail_Routing'); sh.getRange(1, 1, 1, 2).setValues([['perfil_id', 'palabras_clave_asunto (separadas por |)']]).setFontWeight('bold'); }
    const data = sh.getDataRange().getValues();
    const i = data.findIndex((r, n) => n > 0 && r[0] === perfilId);
    const cur = i >= 0 ? String(data[i][1] || '').trim() : '';
    choqueEdicion_('palabras del asunto', cur, v, base.mail);
    if (cur !== v) {
      logForm_(ss, quien, perfilId, 'Mail_Routing', { palabras: cur });
      if (i >= 0) sh.getRange(i + 1, 2).setNumberFormat('@').setValue(v);
      else if (v) sh.getRange(sh.getLastRow() + 1, 1, 1, 2).setNumberFormat('@').setValues([[perfilId, v]]);
    }
  }
}
