/* Configuración de cada búsqueda desde el tablero (solo administrador): datos y links, requisitos para la IA, preguntas de postulación,
   criterios del screening, competencias, mensajes y ajustes generales. Todo se escribe en las mismas pestañas que antes se editaban a mano;
   lo anterior queda en la pestaña «Historial_Config». */
const HOJA_CFG_HIST_ = 'Historial_Config';
const AJUSTES_EDIT_ = [
  ['nombre_cliente', 'Nombre del cliente (aparece en los formularios y mensajes)', 'txt'],
  ['firma', 'Firma de los mails', 'txt'],
  ['asunto_mail', 'Asunto de los mails de descarte', 'txt'],
  ['umbral_a', 'Puntaje mínimo para nota A (0 a 10)', 'num'],
  ['umbral_bplus', 'Puntaje mínimo para nota B+ (0 a 10)', 'num'],
  ['umbral_b', 'Puntaje mínimo para nota B (0 a 10)', 'num'],
  ['dias_alerta', 'Días sin movimiento para avisar', 'int'],
  ['duracion_screening', 'Duración del screening (minutos)', 'int'],
  ['duracion_competencias', 'Duración de competencias (minutos)', 'int'],
  ['duracion_pami2', 'Duración de la entrevista técnica (minutos)', 'int'],
  ['link_agenda', 'Link de agenda (general)', 'url'],
  ['link_agenda_screening', 'Link de agenda · screening', 'url'],
  ['link_agenda_competencias', 'Link de agenda · competencias', 'url'],
  ['link_agenda_tecnica', 'Link de agenda · entrevista técnica', 'url'],
  ['link_agenda_screening_poli2', 'Link de agenda · screening (Policlínico 2)', 'url'],
  ['link_agenda_competencias_poli2', 'Link de agenda · competencias (Policlínico 2)', 'url'],
  ['link_agenda_tecnica_poli2', 'Link de agenda · entrevista técnica (Policlínico 2)', 'url']
];

function cfgTxt_(v, max) { return String(v == null ? '' : v).replace(/[<>]/g, '').replace(/\r/g, '').trim().replace(/^=+/, '').slice(0, max || 2000); }

function cfgHist_(tipo, bid, antes) {
  try {
    const ss = ss_();
    let sh = ss.getSheetByName(HOJA_CFG_HIST_);
    if (!sh) {
      sh = ss.insertSheet(HOJA_CFG_HIST_);
      sh.getRange(1, 1, 1, 5).setValues([['Fecha', 'Quién', 'Búsqueda', 'Qué', 'Valor anterior']]).setFontWeight('bold').setBackground('#034EA2').setFontColor('#FFFFFF');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), QUIEN_ || '', bid || '', tipo, JSON.stringify(antes).slice(0, 45000)]);
  } catch (e) { }
}

function cfgFilas_(nombre, bid) {
  const sh = hoja_(nombre);
  return sh ? tablaBusq_(nombre, bid) : [];
}

function getConfigBusqueda(k, bid) {
  soloAdmin_(k);
  const b = busquedas_().find(x => x.id === bid);
  if (!b) throw new Error('No encontré la búsqueda');
  const aj = ajustes_();
  const pl = hoja_('Plantillas').getDataRange().getValues().slice(1).filter(r => r[0] && !/-anterior$/.test(String(r[0])));
  return {
    busqueda: b,
    preguntas: cfgFilas_('Preguntas', bid).map(r => ({ id: String(r[1]), etiqueta: String(r[2] || ''), tipo: String(r[3] || 'texto').toLowerCase(), opciones: String(r[4] || ''), requerido: String(r[5] || '').toUpperCase() === 'NO' ? 'NO' : 'SI' })),
    eval: cfgFilas_('Screening_Eval', bid).map(r => ({ id: String(r[1]), etiqueta: String(r[2] || ''), peso: Number(r[3]) || 1 })),
    val: cfgFilas_('Screening_Val', bid).map(r => ({ id: String(r[1]), etiqueta: String(r[2] || ''), opciones: String(r[3] || '') })),
    comps: cfgFilas_('Competencias', bid).map(r => ({ id: String(r[1]), nombre: String(r[2] || ''), descripcion: String(r[3] || ''), peso: Number(r[4]) || 1, preguntas: String(r[5] || '') })),
    tecnica: cfgFilas_('Tecnica', bid).map(r => ({ id: String(r[1]), nombre: String(r[2] || ''), descripcion: String(r[3] || ''), peso: Number(r[4]) || 1, preguntas: String(r[5] || '') })),
    hayTecnica: !!ss_().getSheetByName('Tecnica'),
    plantillas: pl.map(r => ({ clave: String(r[0]), canal: String(r[1] || ''), texto: String(r[2] || '') })),
    ajustes: AJUSTES_EDIT_.map(a => ({ k: a[0], l: a[1], t: a[2], v: String(aj[a[0]] == null ? '' : aj[a[0]]) }))
  };
}

// reemplaza todas las filas de una búsqueda en una pestaña por las nuevas (deja intactas las de otras búsquedas)
function cfgReemplazar_(nombre, bid, filas, ncols) {
  const sh = hoja_(nombre);
  const d = sh.getDataRange().getValues();
  for (let i = d.length - 1; i >= 1; i--) if (d[i][0] === bid) sh.deleteRow(i + 1);
  if (filas.length) {
    const n = sh.getLastRow() + 1;
    if (sh.getMaxRows() < n + filas.length) sh.insertRowsAfter(sh.getMaxRows(), n + filas.length - sh.getMaxRows());
    sh.getRange(n, 1, filas.length, ncols).setValues(filas);
  }
  return d.slice(1).filter(r => r[0] === bid);
}

function cfgIds_(items, pref) {
  const usados = {};
  items.forEach(x => { if (x.id) usados[x.id] = true; });
  return items.map(x => {
    if (x.id) return String(x.id);
    let id; do { id = pref + nuevoId_().slice(0, 6); } while (usados[id]);
    usados[id] = true; return id;
  });
}

function cfgPeso_(v, et) {
  const n = Number(String(v).replace(',', '.'));
  if (!isFinite(n) || n <= 0 || n > 10) throw new Error('El peso de «' + et + '» tiene que ser un número entre 0,5 y 10.');
  return n;
}

function cfgComps_(items, bid, hoja) {
  if (!items.length) throw new Error('Tiene que quedar al menos una competencia.');
  const ids = cfgIds_(items, 'c');
  return items.map((c, i) => {
    const nom = cfgTxt_(c.nombre, 120);
    if (!nom) throw new Error('Hay una competencia sin nombre.');
    const pr = String(c.preguntas || '').split(/\n|\|/).map(x => cfgTxt_(x, 400)).filter(Boolean);
    return [bid, ids[i], nom, cfgTxt_(c.descripcion, 1500), cfgPeso_(c.peso, nom), pr.join('|')];
  });
}

function guardarConfigBusqueda(k, bid, p) {
  soloAdmin_(k);
  const b = busquedas_().find(x => x.id === bid);
  if (!b) throw new Error('No encontré la búsqueda');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (p.busqueda) {
      const sh = hoja_('Busquedas');
      const d = sh.getDataRange().getValues();
      const fila = d.findIndex((r, i) => i > 0 && r[0] === bid) + 1;
      const antes = {};
      const poner = (campo, v) => {
        const c = HEAD_BUSQ.indexOf(campo) + 1;
        antes[campo] = d[fila - 1][c - 1];
        sh.getRange(fila, c).setNumberFormat('@').setValue(v);
      };
      const q = p.busqueda;
      const tit = cfgTxt_(q.titulo, 150); if (!tit) throw new Error('El título no puede quedar vacío.');
      const est = String(q.estado || '').toUpperCase(); if (['ABIERTA', 'CERRADA'].indexOf(est) === -1) throw new Error('El estado tiene que ser ABIERTA o CERRADA.');
      ['link_perfil', 'link_flyer'].forEach(l => { const v = cfgTxt_(q[l], 600); if (v && !/^https?:\/\//i.test(v)) throw new Error('El link «' + l + '» tiene que empezar con https://'); });
      poner('titulo', tit); poner('ubicacion', cfgTxt_(q.ubicacion, 150)); poner('modalidad', cfgTxt_(q.modalidad, 100)); poner('estado', est);
      poner('descripcion', cfgTxt_(q.descripcion, 8000)); poner('requisitos', cfgTxt_(q.requisitos, 8000));
      poner('recomendacion', cfgTxt_(q.recomendacion, 2000)); poner('palabras_clave', cfgTxt_(q.palabras_clave, 1000));
      poner('link_perfil', cfgTxt_(q.link_perfil, 600)); poner('link_flyer', cfgTxt_(q.link_flyer, 600));
      cfgHist_('Búsqueda', bid, antes);
    }
    if (p.prompt_ia !== undefined) {
      const t = String(p.prompt_ia).replace(/[<>]/g, '').replace(/\r/g, '').trim();
      if (t.length < 40) throw new Error('Los requisitos para la IA son muy cortos.');
      if (t.length > 45000) throw new Error('Los requisitos para la IA son demasiado largos.');
      const sh = hoja_('Busquedas');
      const fila = sh.getDataRange().getValues().findIndex((r, i) => i > 0 && r[0] === bid) + 1;
      const c = HEAD_BUSQ.indexOf('prompt_ia') + 1;
      cfgHist_('Requisitos para la IA', bid, { prompt_ia: sh.getRange(fila, c).getValue() });
      sh.getRange(fila, c).setNumberFormat('@').setValue(t);
    }
    if (p.preguntas) {
      const ids = cfgIds_(p.preguntas, 'p');
      const filas = p.preguntas.map((q, i) => {
        const et = cfgTxt_(q.etiqueta, 400); if (!et) throw new Error('Hay una pregunta sin texto.');
        const tipo = String(q.tipo).toLowerCase() === 'select' ? 'select' : 'texto';
        const ops = tipo === 'select' ? String(q.opciones || '').split('|').map(o => cfgTxt_(o, 200)).filter(Boolean) : [];
        if (tipo === 'select' && ops.length < 2) throw new Error('La pregunta «' + et + '» necesita al menos dos opciones separadas con |');
        return [bid, ids[i], et, tipo, ops.join('|'), q.requerido === 'NO' ? 'NO' : 'SI'];
      });
      cfgHist_('Preguntas de postulación', bid, cfgReemplazar_('Preguntas', bid, filas, 6));
    }
    if (p.eval || p.val) {
      if (p.eval) {
        if (!p.eval.length) throw new Error('Tiene que quedar al menos un criterio.');
        const ids = cfgIds_(p.eval, 'e');
        const filas = p.eval.map((x, i) => { const et = cfgTxt_(x.etiqueta, 300); if (!et) throw new Error('Hay un criterio sin texto.'); return [bid, ids[i], et, cfgPeso_(x.peso, et)]; });
        cfgHist_('Screening · criterios', bid, cfgReemplazar_('Screening_Eval', bid, filas, 4));
      }
      if (p.val) {
        const ids = cfgIds_(p.val, 'v');
        const filas = p.val.map((x, i) => {
          const et = cfgTxt_(x.etiqueta, 300); if (!et) throw new Error('Hay una condición sin texto.');
          const ops = String(x.opciones || '').split('|').map(o => cfgTxt_(o, 100)).filter(Boolean);
          if (ops.length < 2) throw new Error('La condición «' + et + '» necesita al menos dos opciones separadas con |');
          return [bid, ids[i], et, ops.join('|')];
        });
        cfgHist_('Screening · condiciones', bid, cfgReemplazar_('Screening_Val', bid, filas, 4));
      }
    }
    if (p.comps) cfgHist_('Competencias', bid, cfgReemplazar_('Competencias', bid, cfgComps_(p.comps, bid), 6));
    if (p.tecnica) {
      if (!ss_().getSheetByName('Tecnica')) throw new Error('Esta instalación no tiene la pestaña Tecnica.');
      cfgHist_('Entrevista técnica', bid, cfgReemplazar_('Tecnica', bid, cfgComps_(p.tecnica, bid), 6));
    }
    if (p.plantillas) {
      const sh = hoja_('Plantillas');
      const d = sh.getDataRange().getValues();
      const antes = [];
      p.plantillas.forEach(t => {
        const i = d.findIndex((r, j) => j > 0 && r[0] === t.clave);
        if (i < 1) throw new Error('No existe el mensaje «' + t.clave + '».');
        const txt = String(t.texto == null ? '' : t.texto).replace(/\r/g, '').trim();
        if (txt.length < 10) throw new Error('El mensaje «' + t.clave + '» está vacío o es muy corto.');
        if (txt.length > 20000) throw new Error('El mensaje «' + t.clave + '» es demasiado largo.');
        if (txt !== String(d[i][2])) { antes.push({ clave: t.clave, texto: d[i][2] }); sh.getRange(i + 1, 3).setNumberFormat('@').setValue(txt).setWrap(true); }
      });
      if (antes.length) cfgHist_('Mensajes', '', antes);
    }
    if (p.ajustes) {
      const sh = hoja_('Ajustes');
      const d = sh.getDataRange().getValues();
      const upd = {};
      Object.keys(p.ajustes).forEach(clave => {
        const def = AJUSTES_EDIT_.find(a => a[0] === clave);
        if (!def) throw new Error('Ajuste no editable: ' + clave);
        let v = cfgTxt_(p.ajustes[clave], 600);
        if (def[2] === 'url' && v && !/^https?:\/\//i.test(v)) throw new Error('«' + def[1] + '» tiene que empezar con https://');
        if (def[2] === 'int' && !/^\d+$/.test(v)) throw new Error('«' + def[1] + '» tiene que ser un número entero.');
        if (def[2] === 'num') { v = v.replace(',', '.'); const n = Number(v); if (v === '' || !isFinite(n) || n < 0 || n > 10) throw new Error('«' + def[1] + '» tiene que ser un número entre 0 y 10.'); }
        if (clave === 'nombre_cliente' && !v) throw new Error('El nombre del cliente no puede quedar vacío.');
        upd[clave] = v;
      });
      const m = Object.assign({}, ajustes_(), upd);
      if (!(Number(m.umbral_a) >= Number(m.umbral_bplus) && Number(m.umbral_bplus) >= Number(m.umbral_b))) throw new Error('Los puntajes tienen que ir de mayor a menor: A, B+ y B.');
      const antes = {};
      Object.keys(upd).forEach(clave => {
        const i = d.findIndex((r, j) => j > 0 && String(r[0]).trim() === clave);
        if (i < 1) { sh.appendRow([clave, upd[clave]]); antes[clave] = ''; }
        else { antes[clave] = d[i][1]; if (String(d[i][1]) !== upd[clave]) sh.getRange(i + 1, 2).setNumberFormat('@').setValue(upd[clave]); }
      });
      cfgHist_('Ajustes generales', '', antes);
    }
  } finally { lock.releaseLock(); }
  return { ok: true };
}
