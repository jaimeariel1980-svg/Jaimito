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
  const cands = [];
  for (let i = 1; i < d.length && cands.length < 400; i++) {
    const nombre = String(d[i][1] || '').trim();
    const estado = String(d[i][8] || '').trim();
    if (!nombre || /^Eliminado\|/.test(estado)) continue;
    cands.push({ fila: i + 1, nombre: nombre, estado: estado, cv: d[i][4] || '', scr: d[i][11] || '', comp: d[i][15] || '', zona: d[i][23] || '', cita: d[i][36] || '', ent: String(d[i][45] || '') });
  }
  const lista = cands.map(c => c.fila + ' | ' + c.nombre + ' | ' + (c.estado || '—') + ' | CV ' + (c.cv || '-') + ' · SCR ' + (c.scr || '-') + ' · COMP ' + (c.comp || '-') + (c.cita ? ' | cita ' + c.cita + (c.ent ? ' con ' + c.ent : '') : '')).join('\n');
  // fichas completas de los candidatos nombrados en la conversación (hasta 2): resumen sin datos sensibles, contacto y screening cargado
  const conv = norm(msgs.slice(-4).map(m => m.texto).join(' '));
  const nombrados = cands.filter(c => {
    const t = norm(c.nombre).split(' ').filter(x => x.length > 2);
    return t.length >= 2 && t.filter(x => conv.indexOf(x) >= 0).length >= 2;
  }).slice(0, 2);
  const fichas = nombrados.map(c => {
    const r = d[c.fila - 1];
    return '--- FICHA fila ' + c.fila + ': ' + c.nombre + '\nEstado: ' + c.estado + '\nTeléfono: ' + (r[2] || '-') + ' · Mail: ' + (r[21] || '-') + ' · Zona: ' + (r[23] || '-') +
      '\nAnálisis del CV:\n' + bloomyLimpio_(r[6], 1800) + '\nRespuestas del screening cargado:\n' + (bloomyScreeningTexto_(perfilId, c.nombre) || '(todavía no hay screening cargado)');
  }).join('\n\n');
  let entrevistadores = [];
  try { entrevistadores = leerUsuarios_().filter(u => u.activo === 'SI' && (usuHace_(u, 'screening') || usuHace_(u, 'competencias'))).map(u => u.nombre); } catch (e) { }
  const hoy = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', "EEEE dd/MM/yyyy HH:mm");
  const sistema = 'Sos Bloomy, el asistente del tablero de reclutamiento de Blumia Consulting. Hablás en español rioplatense, corto y directo. Trabajás solo con la búsqueda «' + (cfg.nombre_perfil || cfg.nombre || perfilId) + '». Hoy es ' + hoy + ' (hora de Argentina).\n' +
    'REGLAS: 1) Usá solo los datos de abajo; si algo no está, decilo, no inventes. 2) Nunca menciones edad, género ni situación familiar. 3) Si hay dos candidatos que pueden ser a quien se refieren, preguntá. 4) No escribas datos de contacto salvo que los pidan y estén en la FICHA. 5) Nunca digas que hiciste algo que cambia datos: solo proponés una acción y la persona la confirma con un botón.\n' +
    'ACCIONES (campo "accion", una sola por respuesta, o null): ' +
    '{"tipo":"cv","fila":N} para dar el link del CV; ' +
    '{"tipo":"devolucion","fila":N,"mensaje":"..."} para redactar el mensaje de devolución del screening al candidato (cálido, profesional, 4-6 líneas, sin prometer contratación, sin nota ni puntajes, firmado «Equipo Blumia Consulting»; si no hay screening cargado, avisá y no lo armes); ' +
    '{"tipo":"agendar","fila":N,"etapa":"screening|competencias","fecha":"AAAA-MM-DD","hora":"HH:MM","entrevistador":"nombre o vacío"} para agendar (necesita fecha y hora claras; si falta algo, preguntá; entrevistadores posibles: ' + (entrevistadores.join(', ') || 'ninguno cargado') + '); ' +
    '{"tipo":"abrir","fila":N} para abrir la ficha.\n' +
    'FORMATO: respondé SOLO un JSON {"texto":"respuesta breve","accion":null o la acción}. "fila" es el número de la primera columna de la lista.\n\n' +
    'CANDIDATOS (fila | nombre | estado | notas | cita):\n' + lista + (fichas ? '\n\n' + fichas : '');
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
  if (a.tipo === 'devolucion') {
    const m = bloomyTexto_(a.mensaje).slice(0, 1800);
    if (!m) return null;
    return Object.assign(base, { mensaje: m, tel: String(r[2] || '').replace(/\D/g, '') });
  }
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
