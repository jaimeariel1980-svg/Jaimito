const CFG_CAMPOS_ = ['nombre', 'nombre_perfil', 'empresa', 'empresa_candidatos', 'logo_cliente', 'preguntas_base', 'banda_salarial', 'link_screening', 'link_competencias', 'link_empresa_virtual', 'link_empresa_presencial', 'link_flyer', 'link_perfil'];

function getConfigBusqueda_(pin, perfilId) {
  pinValido_(pin);
  const cfg = getConfig(perfilId);
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const datos = {};
  CFG_CAMPOS_.forEach(k => { datos[k] = typeof cfg[k] === 'string' ? cfg[k] : (cfg[k] == null ? '' : String(cfg[k])); });
  return { perfil: perfilPuestoInfo_(cfg), extra: configExtraLeer_(ss, cfg, perfilId), datos: datos, prompt_extra: String(cfg.prompt_extra || ''), preguntas: getPreguntasPostulacion(perfilId), plantillas: listarPlantillas_(ss, perfilId) };
}

function guardarConfigBusqueda_(pin, perfilId, p, quien) {
  pinValido_(pin);
  getConfig(perfilId);
  quien = limpiarTxt_(quien) || 'sin nombre';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);

    const base = p.base || {};
    if (p.datos || p.prompt_extra !== undefined) {
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
      if (p.datos) CFG_CAMPOS_.forEach(k => {
        if (p.datos[k] === undefined) return;
        const v = String(p.datos[k]).replace(/[<>]/g, '').trim();
        if ((k === 'nombre' || k === 'nombre_perfil' || k === 'empresa') && !v) throw new Error('«' + k + '» no puede quedar vacío.');
        if ((/^link_/.test(k) || k === 'logo_cliente') && v && !/^https?:\/\//i.test(v)) throw new Error('El link «' + k + '» tiene que empezar con https://');
        if (k === 'preguntas_base') {
          const L = v.toLowerCase().split(',').map(x => x.trim()).filter(Boolean);
          if (L.some(x => x !== 'ninguna' && PREG_BASE_.indexOf(x) < 0)) throw new Error('Alguna pregunta fija no es válida.');
          if (!v) return; // sin dato = la búsqueda sigue mostrando todas las preguntas fijas
          poner(k, L.join(','), (base.datos || {})[k]); return;
        }
        if (k === 'banda_salarial') { const num = String(v).replace(/\D/g, ''); if (v && Number(num) < 1000) throw new Error('La banda salarial va en pesos, por ejemplo 13000000.'); poner(k, num, (base.datos || {})[k]); return; }
        poner(k, v, (base.datos || {})[k]);
      });
      if (p.prompt_extra !== undefined) {
        const t = String(p.prompt_extra).replace(/[<>]/g, '').trim();
        if (t.length < 40) throw new Error('Los requisitos para la IA son muy cortos.');
        if (t.length > 45000) throw new Error('Los requisitos para la IA son demasiado largos.');
        poner('prompt_extra', t, base.prompt_extra);
      }
      logForm_(ss, quien, perfilId, 'Config', antes);
    }
    if (p.datos && p.datos.empresa !== undefined) {
      const shH = ss.getSheetByName('Perfiles_Hero');
      if (shH) {
        const dh = shH.getDataRange().getValues(), fh = dh.findIndex((r, i) => i > 0 && r[0] === perfilId) + 1, ch = dh[0].indexOf('empresa') + 1;
        const ne = String(p.datos.empresa).replace(/[<>]/g, '').trim();
        if (fh && ch && ne && String(dh[fh - 1][ch - 1]) !== ne) shH.getRange(fh, ch).setNumberFormat('@').setValue(ne);
      }
    }
    if (p.extra) configExtraGuardar_(ss, perfilId, p.extra, base, quien);

    if (p.preguntas) {
      const ids = ids_(p.preguntas, 'p');
      reemplazarFilasPerfil_(ss, 'Postulacion_Preguntas', perfilId, p.preguntas.map((q, k) => {
        const et = limpiarTxt_(q.etiqueta);
        if (!et) throw new Error('Hay una pregunta sin texto.');
        const tipo = String(q.tipo).toLowerCase() === 'select' ? 'select' : 'texto';
        const ops = tipo === 'select' ? String(q.opciones || '').split('|').map(o => limpiarTxt_(o)).filter(Boolean) : [];
        if (tipo === 'select' && !ops.length) throw new Error('La pregunta «' + et + '» necesita opciones.');
        return [perfilId, ids[k], et, tipo, ops.join('|'), q.requerido === 'NO' ? 'NO' : 'SI'];
      }), quien);
    }

    if (p.plantillas) guardarPlantillas_(ss, perfilId, p.plantillas, quien);
  } finally { lock.releaseLock(); }
  return { ok: true };
}

// ---- Plantillas de mensajes: texto general (columna perfil vacía) y texto propio de una búsqueda (columna perfil) ----
function colPerfilPlantillas_(sh) {
  const h = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 4)).getValues()[0];
  let c = h.indexOf('perfil') + 1;
  if (!c) { c = Math.max(sh.getLastColumn(), 4) + 1; sh.getRange(1, c).setValue('perfil').setFontWeight('bold'); }
  return c;
}

// «Soy <nombre>» y la firma con el nombre de un usuario pasan a [ENTREVISTADOR] (así firma quien envía)
function firmaConPlaceholder_(t) {
  const nombres = entrevistadoresDefault_().map(n => n.replace(/[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ]/g, '')).filter(Boolean).join('|');
  if (!nombres) return String(t || '');
  return String(t || '').replace(new RegExp('\\bSoy (' + nombres + ')\\b', 'g'), 'Soy [ENTREVISTADOR]').replace(new RegExp('(Saludos,?[ \\t]*\\r?\\n)(' + nombres + ')\\b', 'g'), '$1[ENTREVISTADOR]');
}

// devuelve {texto, propia}: el texto propio de la búsqueda si existe, si no el general
function plantillaTexto_(clave, perfilId) {
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Plantillas');
  if (!sh || sh.getLastRow() < 2) return { texto: '', propia: false };
  const data = sh.getDataRange().getValues();
  const c = data[0].indexOf('perfil');
  let general = '', propia = null;
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] !== clave) continue;
    const pf = c >= 0 ? String(data[i][c] || '').trim() : '';
    if (!pf) general = data[i][3];
    else if (pf === perfilId) propia = data[i][3];
  }
  return propia ? { texto: propia, propia: true } : { texto: general, propia: false };
}

function listarPlantillas_(ss, perfilId) {
  const sh = ss.getSheetByName('Plantillas');
  if (!sh || sh.getLastRow() < 2) return [];
  const data = sh.getDataRange().getValues();
  const c = data[0].indexOf('perfil');
  const out = [];
  data.slice(1).forEach(r => {
    if (!r[0] || (c >= 0 && String(r[c] || '').trim())) return;
    const ef = plantillaTexto_(String(r[0]), perfilId);
    out.push({ clave: String(r[0]), canal: String(r[1]), texto: firmaConPlaceholder_(ef.texto), propia: ef.propia, general: firmaConPlaceholder_(r[3]) });
  });
  return out;
}

function guardarPlantillas_(ss, perfilId, lista, quien) {
  const sh = ss.getSheetByName('Plantillas');
  const cp = colPerfilPlantillas_(sh);
  const antes = [];
  lista.forEach(t => {
    const data = sh.getDataRange().getValues();
    let iGen = -1, iPro = -1;
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] !== t.clave) continue;
      const pf = String(data[i][cp - 1] || '').trim();
      if (!pf) iGen = i; else if (pf === perfilId) iPro = i;
    }
    if (iGen < 0) return;
    if (t.volver) {
      if (iPro >= 0) { antes.push({ clave: t.clave, perfil: perfilId, texto: data[iPro][3] }); sh.deleteRow(iPro + 1); }
      return;
    }
    const txt = String(t.texto || '').replace(/[<>]/g, '').trim();
    if (txt.length < 10) throw new Error('El mensaje «' + t.clave + '» quedó vacío.');
    if (t.alcance === 'todas') {
      if (txt !== String(data[iGen][3])) { antes.push({ clave: t.clave, texto: data[iGen][3] }); sh.getRange(iGen + 1, 4).setNumberFormat('@').setValue(txt); }
    } else if (iPro >= 0) {
      if (txt !== String(data[iPro][3])) { antes.push({ clave: t.clave, perfil: perfilId, texto: data[iPro][3] }); sh.getRange(iPro + 1, 4).setNumberFormat('@').setValue(txt); }
    } else if (txt !== String(data[iGen][3])) {
      const fila = [t.clave, data[iGen][1], '', txt]; fila[cp - 1] = perfilId;
      sh.getRange(sh.getLastRow() + 1, 1, 1, Math.max(cp, 4)).setNumberFormat('@').setValues([fila.concat(new Array(Math.max(cp, 4) - fila.length).fill(''))]);
    }
  });
  if (antes.length) logForm_(ss, quien, perfilId, 'Plantillas', antes);
}
