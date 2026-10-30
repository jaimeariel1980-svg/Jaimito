/* Bloomy: asistente de chat del tablero (como Azcu en Azcubot, pero chico). Consulta candidatos de UNA búsqueda, da el link del CV,
   arma la devolución del screening y propone agendar. Todo lo que cambia datos pasa por una tarjeta con «Confirmar» (bloomyEjecutar_).
   Nunca ve edad, género ni situación familiar. Modelo: tarea «bloomy» (Haiku; MODELO_BLOOMY lo cambia sin publicar). */

const BLOOMY_PRIVADO_ = /\b(edad|g[eé]nero|situaci[oó]n familiar|estado civil|hij[oa]s?|casad[oa]|solter[oa]|divorciad[oa]|a[nñ]os de edad)\b/i;

function bloomyLimpio_(txt, max) {
  return String(txt || '').split(/\r?\n/).filter(l => !BLOOMY_PRIVADO_.test(l)).join('\n').replace(/\s+\n/g, '\n').trim().slice(0, max);
}

function bloomyTexto_(s) { return String(s == null ? '' : s).replace(/[<>]/g, '').trim(); }

function bloomyLinkCV_(sh, fila) {
  try {
    const cel = sh.getRange(fila, 8);
    let u = '';
    const f = cel.getFormula().match(/HYPERLINK\("([^"]+)"/i);
    if (f) u = f[1];
    if (!u) { const rt = cel.getRichTextValue(); if (rt && rt.getLinkUrl()) u = rt.getLinkUrl(); }
    if (!u) u = String(cel.getValue() || '');
    const id = (u.match(/[?&]id=([-\w]{25,})/) || [])[1] || (u.match(/[-\w]{25,}/) || [''])[0];
    return id ? 'https://drive.google.com/file/d/' + id + '/view' : '';
  } catch (e) { return ''; }
}

function bloomyScreeningTexto_(perfilId, nombre) {
  try {
    const p = getEvaluacionPrevia_(perfilId, nombre, 'screening');
    if (!p || !p.estado) return '';
    let o = JSON.parse(p.estado);
    const st = o && (o.state || o.estado || o);
    let t = typeof st === 'string' ? st : JSON.stringify(st);
    t = t.replace(/<[^>]+>/g, ' ').replace(/\\n/g, '\n').replace(/\s{3,}/g, '  ');
    return bloomyLimpio_(t, 3500);
  } catch (e) { return ''; }
}

function bloomyChat_(perfilId, mensajes, quien) {
  QUIEN_ = quien || '';
  const cfg = getConfig(perfilId);
  const sh = getSheet(cfg);
  const d = sh.getDataRange().getValues();
  const msgs = (Array.isArray(mensajes) ? mensajes : []).slice(-12).map(m => ({ rol: m && m.rol === 'bot' ? 'bot' : 'user', texto: bloomyTexto_(m && m.texto).slice(0, 1500) })).filter(m => m.texto);
  if (!msgs.length || msgs[msgs.length - 1].rol !== 'user') throw new Error('No hay mensaje para responder.');
  const norm = evalNorm_;
  const entEval = entrevistadoresEval_(perfilId);
  const cands = [];
  for (let i = 1; i < d.length && cands.length < 400; i++) {
    const nombre = String(d[i][1] || '').trim();
    const estado = String(d[i][8] || '').trim();
    if (!nombre || /^Eliminado\|/.test(estado)) continue;
    cands.push({ fila: i + 1, nombre: nombre, estado: estado, cv: d[i][4] || '', scr: d[i][11] || '', comp: d[i][15] || '', zona: d[i][23] || '', cita: d[i][36] || '', ent: String(d[i][45] || ''), via: String(d[i][9] || '').trim(), eS: entEval[norm(nombre) + '|screening'] || '', eC: entEval[norm(nombre) + '|competencias'] || '' });
  }
  const resumen = bloomyResumen_(cfg, sh, d, cands);
  const lista = cands.map(c => c.fila + ' | ' + c.nombre + ' | ' + (c.estado || '—') + ' | CV ' + (c.cv || '-') + ' · SCR ' + (c.scr || '-') + ' · COMP ' + (c.comp || '-') + (c.cita ? ' | cita ' + c.cita + (c.ent ? ' con ' + c.ent : '') : '') + (c.eS || c.eC ? ' | entrevistó' + (c.eS ? ' screening ' + c.eS : '') + (c.eC ? ' competencias ' + c.eC : '') : '') + (c.via ? ' | llegó por ' + c.via : '')).join('\n');
  // fichas completas de los candidatos nombrados en la conversación (hasta 2): resumen sin datos sensibles, contacto y screening cargado
  const conv = norm(msgs.slice(-4).map(m => m.texto).join(' '));
  const nombrados = cands.filter(c => {
    const t = norm(c.nombre).split(' ').filter(x => x.length > 2);
    return t.length >= 2 && t.filter(x => conv.indexOf(x) >= 0).length >= 2;
  }).slice(0, 2);
  const fichas = nombrados.map(c => {
    const r = d[c.fila - 1];
    let hist = '';
    try { hist = leerHistorial_(sh.getName(), c.nombre, 20).map(h => h.fecha + ' · ' + h.tipo + ' · ' + bloomyLimpio_(h.detalle, 160) + (h.quien ? ' (' + h.quien + ')' : '')).join('\n'); } catch (e) { }
    return '--- FICHA fila ' + c.fila + ': ' + c.nombre + '\nEstado: ' + c.estado + '\nRecibido: ' + (r[0] ? Utilities.formatDate(new Date(r[0]), 'GMT-3', 'dd/MM/yyyy') : '-') + ' · Llegó por: ' + (c.via || '-') + '\nEntrevistó screening: ' + (c.eS || '-') + ' · competencias: ' + (c.eC || '-') + ' · Cita: ' + (c.cita || '-') + (c.ent ? ' con ' + c.ent : '') + '\nCitas (calendario):\n' + String(r[37] || '-').slice(0, 600) + '\nObservaciones: ' + bloomyLimpio_(r[d[0].indexOf('Observaciones')] || '', 500) + '\nHistorial de movimientos (el más reciente primero):\n' + (hist || '-') + '\nTeléfono: ' + (r[2] || '-') + ' · Mail: ' + (r[21] || '-') + ' · Zona: ' + (r[23] || '-') +
      '\nAnálisis del CV:\n' + bloomyLimpio_(r[6], 1800) + '\nRespuestas del screening cargado:\n' + (bloomyScreeningTexto_(perfilId, c.nombre) || '(todavía no hay screening cargado)');
  }).join('\n\n');
  let entrevistadores = [];
  try { entrevistadores = leerUsuarios_().filter(u => u.activo === 'SI' && (usuHaceEn_(u, 'screening', perfilId) || usuHaceEn_(u, 'competencias', perfilId))).map(u => u.nombre); } catch (e) { }
  const hoy = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', "EEEE dd/MM/yyyy HH:mm");
  const sistema = 'Sos Bloomy, el asistente del tablero de reclutamiento de Blumia Consulting. Hablás en español rioplatense, corto y directo. Trabajás solo con la búsqueda «' + (cfg.nombre_perfil || cfg.nombre || perfilId) + '». Hoy es ' + hoy + ' (hora de Argentina).\n' +
    'REGLAS: 1) Usá solo los datos de abajo; si algo no está, decilo, no inventes. 2) Nunca menciones edad, género ni situación familiar. 3) Si hay dos candidatos que pueden ser a quien se refieren, preguntá. 4) Para preguntas de proceso (quién lo entrevistó, cuándo, en qué estado está, quién mandó el CV, por dónde llegó) respondé con los datos de la lista y de la FICHA, citando fecha y persona. 5) Podés responder cuántos candidatos hay por etapa o nota, motivos de descarte, de dónde llegan, entrevistas próximas, quién entrevistó, qué falta cargar y datos de la búsqueda, con el RESUMEN y la lista; hacé las cuentas con cuidado. 5b) No escribas datos de contacto salvo que los pidan y estén en la FICHA. 6) Nunca digas que hiciste algo que cambia datos: solo proponés una acción y la persona la confirma con un botón.\n' +
    'ACCIONES (campo "accion", una sola por respuesta, o null): ' +
    '{"tipo":"cv","fila":N} para dar el link del CV; ' +
    '{"tipo":"screening","fila":N} cuando quieren CARGAR o completar el screening de alguien (yo les voy preguntando lo mismo que pide el formulario: no lo cargues vos); '+
    '{"tipo":"agendar","fila":N,"etapa":"screening|competencias","fecha":"AAAA-MM-DD","hora":"HH:MM","entrevistador":"nombre o vacío"} para agendar (necesita fecha y hora claras; si falta algo, preguntá; entrevistadores posibles: ' + (entrevistadores.join(', ') || 'ninguno cargado') + '); ' +
    '{"tipo":"abrir","fila":N} para abrir la ficha.\n' +
    'FORMATO: respondé SOLO un JSON {"texto":"respuesta breve","accion":null o la acción}. "fila" es el número de la primera columna de la lista.\n\n' +
    'RESUMEN DE LA BÚSQUEDA (datos reales, usalos para contestar cuántos, cómo viene, etc.):\n' + resumen + '\n\nCANDIDATOS (fila | nombre | estado | notas | cita | quién entrevistó | llegó por):\n' + lista + (fichas ? '\n\n' + fichas : '');
  const usuario = msgs.map(m => (m.rol === 'bot' ? 'Bloomy: ' : 'Usuario: ') + m.texto).join('\n');
  const crudo = procesarConClaude(sistema, usuario, 'bloomy');
  let o = null;
  try { const m = String(crudo).match(/\{[\s\S]*\}/); o = JSON.parse(m ? m[0] : crudo); } catch (e) { o = { texto: String(crudo).slice(0, 800), accion: null }; }
  const texto = bloomyTexto_(o && o.texto).slice(0, 1500) || 'No pude armar una respuesta. ¿Me lo repetís?';
  return { texto: texto, accion: bloomyAccion_(sh, d, o && o.accion, entrevistadores) };
}

// valida lo que propone el modelo contra la planilla (fila y nombre reales) y arma la tarjeta
function bloomyAccion_(sh, d, a, entrevistadores) {
  if (!a || typeof a !== 'object') return null;
  const fila = Number(a.fila);
  if (!(fila >= 2 && fila <= d.length) || !String(d[fila - 1][1] || '').trim()) return null;
  const r = d[fila - 1];
  const base = { tipo: String(a.tipo), fila: fila, nombre: String(r[1]).trim() };
  if (a.tipo === 'abrir') return base;
  if (a.tipo === 'cv') { const u = bloomyLinkCV_(sh, fila); return Object.assign(base, { url: u }); }
  if (a.tipo === 'screening') return base;
  if (a.tipo === 'agendar') {
    const etapa = a.etapa === 'competencias' ? 'competencias' : (a.etapa === 'screening' ? 'screening' : '');
    if (!etapa || !/^\d{4}-\d{2}-\d{2}$/.test(String(a.fecha)) || !/^\d{2}:\d{2}$/.test(String(a.hora))) return null;
    let ent = bloomyTexto_(a.entrevistador);
    if (ent) { const k = usuClave_(ent); const u = entrevistadores.find(n => usuClave_(n) === k) || entrevistadores.find(n => usuClave_(n).indexOf(k) === 0); ent = u || ''; }
    return Object.assign(base, { etapa: etapa, fecha: String(a.fecha), hora: String(a.hora), entrevistador: ent });
  }
  return null;
}

function bloomyEjecutar_(perfilId, accion, quien) {
  QUIEN_ = quien || '';
  if (!accion || accion.tipo !== 'agendar') throw new Error('Acción no disponible.');
  const sh = getSheet(getConfig(perfilId));
  const fila = filaCandidato_(sh, Number(accion.fila));
  if (String(sh.getRange(fila, 2).getValue() || '').trim() !== String(accion.nombre || '').trim()) throw new Error('El candidato cambió de lugar en la planilla. Pedímelo de nuevo.');
  return agendarCitaTablero_(perfilId, fila, accion.etapa === 'competencias' ? 'competencias' : 'screening', String(accion.fecha), String(accion.hora), 30, '', !!accion.forzar, quien, accion.entrevistador || '');
}

// resumen numérico de la búsqueda para el modelo (las cuentas las hace el código, no la IA)
function bloomyResumen_(cfg, sh, d, cands) {
  const cuenta = (arr) => { const o = {}; arr.forEach(x => { x = String(x || '').trim() || '—'; o[x] = (o[x] || 0) + 1; }); return Object.keys(o).sort((a, b) => o[b] - o[a]).map(k => k + ': ' + o[k]).join(' · '); };
  const H = d[0], cMot = H.indexOf('Motivo de descarte');
  const activos = cands.filter(c => !/^Descartad/i.test(c.estado));
  const desc = cands.filter(c => /^Descartad/i.test(c.estado));
  const motivos = [];
  if (cMot >= 0) desc.forEach(c => String(d[c.fila - 1][cMot] || '').split(/;\s*/).forEach(m => { if (m.trim()) motivos.push(m.trim()); }));
  const citas = cands.filter(c => c.cita).map(c => c.nombre + ' ' + c.cita + (c.ent ? ' (con ' + c.ent + ')' : '')).slice(0, 15);
  const lineas = [
    'Búsqueda: ' + (cfg.nombre_perfil || cfg.nombre || '') + ' · Estado: ' + (String(cfg.estado_busqueda || '').toUpperCase() === 'CERRADA' ? 'CERRADA' : 'ABIERTA') + ' · Banda salarial de referencia: ' + (cfg.banda_salarial || '—'),
    'Total de CVs: ' + cands.length + ' · En proceso (no descartados): ' + activos.length + ' · Descartados: ' + desc.length,
    'Por estado → ' + cuenta(cands.map(c => c.estado)),
    'Notas de CV → ' + cuenta(cands.map(c => c.cv)),
    'Notas de screening → ' + cuenta(cands.filter(c => c.scr).map(c => c.scr)),
    'Notas de competencias → ' + cuenta(cands.filter(c => c.comp).map(c => c.comp)),
    'Motivos de descarte → ' + (motivos.length ? cuenta(motivos) : 'ninguno cargado') + ' · Descartados sin motivo: ' + desc.filter(c => !cMot >= 0 || !String(d[c.fila - 1][cMot] || '').trim()).length,
    'De dónde llegan → ' + cuenta(cands.map(c => c.via)),
    'Citas agendadas → ' + (citas.length ? citas.join(' | ') : 'ninguna')
  ];
  return lineas.join('\n');
}

// ---- Cargar el screening conversando: Bloomy pregunta lo del formulario y deja un borrador que se abre ya completo ----
function bloomyScreeningDef_(perfilId, fila) {
  const sh = getSheet(getConfig(perfilId));
  fila = filaCandidato_(sh, fila);
  const D = getDatosPerfil_(perfilId);
  return {
    fila: fila,
    nombre: String(sh.getRange(fila, 2).getValue() || '').trim(),
    eval: (D.eval_items || []).map(x => ({ id: String(x.id), label: String(x.label || '') })),
    val: (D.val_items || []).map(x => ({ id: String(x.id), label: String(x.label || ''), opts: (x.opts || []).map(String) }))
  };
}

function bloomyScreeningBorrador_(perfilId, fila, resp, quien) {
  const def = bloomyScreeningDef_(perfilId, fila);
  const evalOk = {}, valOk = {};
  const E = ['bajo', 'medio', 'alto'];
  def.eval.forEach(x => { const v = resp && resp.eval ? String(resp.eval[x.id] || '') : ''; if (E.indexOf(v) >= 0) evalOk[x.id] = v; });
  def.val.forEach(x => { const v = resp && resp.val ? String(resp.val[x.id] || '') : ''; if (x.opts.indexOf(v) >= 0) valOk[x.id] = v; });
  const borrador = { eval: evalOk, val: valOk, comentario: bloomyTexto_(resp && resp.comentario).slice(0, 4000), nombre: def.nombre, entrevistador: bloomyTexto_(quien).slice(0, 60) };
  const id = Utilities.getUuid();
  CacheService.getScriptCache().put('bsb_' + id, JSON.stringify(borrador), 7200);
  const url = URL_BASE() + '?tipo=screening&perfil=' + encodeURIComponent(perfilId) + '&nombre=' + encodeURIComponent(def.nombre) + '&entrevistador=' + encodeURIComponent(borrador.entrevistador) + '&borrador=' + id;
  try { registrarHistorial_(getSheet(getConfig(perfilId)).getName(), def.nombre, 'Screening', 'Preparado con Bloomy (borrador sin guardar)', quien); } catch (e) { }
  return { id: id, url: url };
}

function bloomyBorradorLeer_(id) {
  const t = CacheService.getScriptCache().get('bsb_' + String(id || '').replace(/[^\w-]/g, ''));
  return t ? JSON.parse(t) : null;
}
