const CFG_CAMPOS_ = ['nombre', 'nombre_perfil', 'empresa', 'empresa_candidatos', 'link_screening', 'link_competencias', 'link_empresa_virtual', 'link_empresa_presencial', 'link_flyer', 'link_perfil'];

function getConfigBusqueda(pin, perfilId) {
  pinValido_(pin);
  const cfg = getConfig(perfilId);
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const datos = {};
  CFG_CAMPOS_.forEach(k => { datos[k] = typeof cfg[k] === 'string' ? cfg[k] : (cfg[k] == null ? '' : String(cfg[k])); });
  const plSh = ss.getSheetByName('Plantillas');
  const plantillas = plSh && plSh.getLastRow() > 1
    ? plSh.getRange(2, 1, plSh.getLastRow() - 1, 4).getValues().filter(r => r[0]).map(r => ({ clave: String(r[0]), canal: String(r[1]), texto: String(r[3] || '') }))
    : [];
  return { datos: datos, prompt_extra: String(cfg.prompt_extra || ''), preguntas: getPreguntasPostulacion(perfilId), plantillas: plantillas };
}

function guardarConfigBusqueda(pin, perfilId, p, quien) {
  pinValido_(pin);
  getConfig(perfilId);
  quien = limpiarTxt_(quien) || 'sin nombre';
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);

    if (p.datos || p.prompt_extra !== undefined) {
      const sh = ss.getSheetByName('Config');
      const data = sh.getDataRange().getValues();
      const fila = data.findIndex((r, i) => i > 0 && r[0] === perfilId) + 1;
      const antes = {};
      const poner = (k, v) => {
        let col = data[0].indexOf(k) + 1;
        if (!col) { col = sh.getLastColumn() + 1; sh.getRange(1, col).setValue(k); data[0][col - 1] = k; }
        const cel = sh.getRange(fila, col);
        antes[k] = cel.getValue();
        cel.setNumberFormat('@').setValue(v);
      };
      if (p.datos) CFG_CAMPOS_.forEach(k => {
        if (p.datos[k] === undefined) return;
        const v = String(p.datos[k]).replace(/[<>]/g, '').trim();
        if ((k === 'nombre' || k === 'nombre_perfil' || k === 'empresa') && !v) throw new Error('«' + k + '» no puede quedar vacío.');
        if (/^link_/.test(k) && v && !/^https?:\/\//i.test(v)) throw new Error('El link «' + k + '» tiene que empezar con https://');
        poner(k, v);
      });
      if (p.prompt_extra !== undefined) {
        const t = String(p.prompt_extra).replace(/[<>]/g, '').trim();
        if (t.length < 40) throw new Error('Los requisitos para la IA son muy cortos.');
        if (t.length > 45000) throw new Error('Los requisitos para la IA son demasiado largos.');
        poner('prompt_extra', t);
      }
      logForm_(ss, quien, perfilId, 'Config', antes);
    }

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

    if (p.plantillas) {
      const sh = ss.getSheetByName('Plantillas');
      const data = sh.getDataRange().getValues();
      const antes = [];
      p.plantillas.forEach(t => {
        const i = data.findIndex((r, j) => j > 0 && r[0] === t.clave);
        if (i < 1) return;
        const txt = String(t.texto || '').replace(/[<>]/g, '').trim();
        if (txt.length < 10) throw new Error('El mensaje «' + t.clave + '» quedó vacío.');
        if (txt !== String(data[i][3])) {
          antes.push({ clave: t.clave, texto: data[i][3] });
          sh.getRange(i + 1, 4).setNumberFormat('@').setValue(txt);
        }
      });
      if (antes.length) logForm_(ss, quien, perfilId, 'Plantillas', antes);
    }
  } finally { lock.releaseLock(); }
  return { ok: true };
}
