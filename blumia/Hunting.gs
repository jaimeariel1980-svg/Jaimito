/* ============================================================
   HUNTING (30/10, pedido de Ariel para la búsqueda del Director Comercial): personas a las que se aborda directamente, antes de que sean candidatos.
   Viven en la pestaña «Hunting» de la planilla (una fila por persona y búsqueda) y se manejan desde la pestaña «Hunting» del tablero (HuntingPanel.html).
   Etapas: Identificado → Contactado → Respondió → Interesado → En proceso. Cierres: No le interesa · Sin respuesta.
   Cuando dice «me interesa» se sube su perfil de LinkedIn (PDF): entra al tablero como candidato con «Llegó por: Hunting (LinkedIn)» y sigue el flujo normal.
   El embudo (identificados → contactados → respondieron → interesados → en proceso) sale en el Resumen y en el informe de la búsqueda.
   Estas personas NO están en la hoja de candidatos hasta que se interesan: no ensucian los números del proceso.
   ============================================================ */

const HOJA_HUNT_ = 'Hunting';
const HUNT_COLS_ = ['id', 'perfil', 'nombre', 'cargo', 'empresa', 'linkedin', 'prioridad', 'estado', 'nivel', 'creado', 'contactado', 'ultimo_mov', 'seguimiento', 'quien', 'notas', 'mensaje', 'candidato', 'seguimientos', 'pdf', 'datos'];
const HUNT_NIVEL_ = { 'Identificado': 0, 'Contactado': 1, 'Sin respuesta': 1, 'Respondió': 2, 'No le interesa': 2, 'Interesado': 3, 'En proceso': 4 };
const HUNT_FUENTE_ = 'Hunting (LinkedIn)';

function huntHoy_() { return Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'); }
function huntSumarDias_(n) { return Utilities.formatDate(new Date(Date.now() + n * 864e5), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd'); }
function huntNorm_(t) { return String(t || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, ' ').trim(); }

// Solo links de LinkedIn, sin parámetros de seguimiento
function huntUrl_(u) {
  u = String(u || '').trim();
  if (!u) return '';
  if (!/^https?:\/\//i.test(u)) u = 'https://' + u.replace(/^\/+/, '');
  const m = u.match(/^https:\/\/([a-z0-9-]+\.)?linkedin\.com\/[^\s?#]*/i);
  return m ? m[0].replace(/\/+$/, '') : '';
}

function huntHoja_() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let sh = ss.getSheetByName(HOJA_HUNT_);
  if (!sh) {
    sh = ss.insertSheet(HOJA_HUNT_);
    sh.getRange(1, 1, 1, HUNT_COLS_.length).setValues([HUNT_COLS_]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, 1000, HUNT_COLS_.length).setNumberFormat('@'); // todo como texto: las fechas no se convierten solas
  } else {
    const n = sh.getLastColumn();
    if (n < HUNT_COLS_.length) { // columnas nuevas (pdf, datos) en una hoja que ya existía
      sh.getRange(1, n + 1, 1, HUNT_COLS_.length - n).setValues([HUNT_COLS_.slice(n)]).setFontWeight('bold');
      sh.getRange(2, n + 1, 1000, HUNT_COLS_.length - n).setNumberFormat('@');
    }
  }
  return sh;
}

function huntLeerTodo_() {
  const sh = huntHoja_();
  if (sh.getLastRow() < 2) return [];
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, HUNT_COLS_.length).getValues();
  const out = [];
  v.forEach((r, i) => {
    if (!String(r[0]).trim()) return;
    const o = { fila: i + 2 };
    HUNT_COLS_.forEach((c, k) => { o[c] = r[k] instanceof Date ? Utilities.formatDate(r[k], 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd') : String(r[k] == null ? '' : r[k]); });
    o.nivel = Number(o.nivel) || 0; o.seguimientos = Number(o.seguimientos) || 0;
    out.push(o);
  });
  return out;
}

// Lo que viaja al tablero: sin el JSON del PDF (pesa). «tienePdf» habilita «Pasar a la búsqueda»; «nota» es la nota que la IA le puso al leer el PDF.
function huntItem_(o) {
  const x = Object.assign({}, o); delete x.fila;
  x.tienePdf = !!(o.pdf && o.datos);
  let nota = ''; try { nota = JSON.parse(o.datos || '{}').nivel_ajuste || ''; } catch (e) { }
  x.nota = nota; delete x.datos;
  return x;
}

function huntStats_(items) {
  const s = { identificados: 0, contactados: 0, respondieron: 0, interesados: 0, enProceso: 0, noInteresa: 0, sinRespuesta: 0, vencidos: 0 };
  const hoy = huntHoy_();
  items.forEach(o => {
    s.identificados++;
    if (o.nivel >= 1) s.contactados++;
    if (o.nivel >= 2) s.respondieron++;
    if (o.nivel >= 3) s.interesados++;
    if (o.nivel >= 4) s.enProceso++;
    if (o.estado === 'No le interesa') s.noInteresa++;
    if (o.estado === 'Sin respuesta') s.sinRespuesta++;
    if (o.estado === 'Contactado' && o.seguimiento && o.seguimiento <= hoy) s.vencidos++;
  });
  return s;
}

// Números del embudo de hunting de una búsqueda (Resumen e informe)
function huntingStats_(perfilId) {
  try { return huntStats_(huntLeerTodo_().filter(o => o.perfil === perfilId)); } catch (e) { return huntStats_([]); }
}

function huntingListar_(perfilId) {
  const items = huntLeerTodo_().filter(o => o.perfil === perfilId);
  return { items: items.map(huntItem_), stats: huntStats_(items), hoy: huntHoy_() };
}

function huntEscribir_(sh, o) {
  sh.getRange(o.fila, 1, 1, HUNT_COLS_.length).setNumberFormat('@').setValues([HUNT_COLS_.map(c => o[c] == null ? '' : o[c])]);
}

function huntHist_(perfilId, o, detalle, quien) {
  try { registrarHistorial_(getConfig(perfilId).hoja_nombre, o.nombre, 'Hunting', detalle, quien); } catch (e) { }
}

/* Agrega personas a la lista: [{nombre, cargo, empresa, linkedin, prioridad, notas}]. No duplica (mismo LinkedIn o mismo nombre y empresa). */
function huntingAgregar_(perfilId, personasJson, quien) {
  getConfig(perfilId);
  let lista;
  try { lista = JSON.parse(personasJson); } catch (e) { throw new Error('Datos inválidos.'); }
  if (!Array.isArray(lista) || !lista.length) throw new Error('No hay personas para agregar.');
  if (lista.length > 200) throw new Error('Son demasiadas de una vez (máximo 200).');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = huntHoja_();
    const existentes = huntLeerTodo_().filter(o => o.perfil === perfilId);
    const urls = {}, claves = {};
    existentes.forEach(o => { if (o.linkedin) urls[o.linkedin.toLowerCase()] = 1; claves[huntNorm_(o.nombre) + '|' + huntNorm_(o.empresa)] = 1; });
    const cut = (v, n) => limpiarTxt_(v).slice(0, n);
    const filas = []; let repetidos = 0, sinNombre = 0;
    lista.forEach(p => {
      const nombre = cut(p && p.nombre, 90);
      if (nombre.length < 2) { sinNombre++; return; }
      const empresa = cut(p && p.empresa, 90), url = huntUrl_(p && p.linkedin);
      const k = huntNorm_(nombre) + '|' + huntNorm_(empresa);
      if ((url && urls[url.toLowerCase()]) || claves[k]) { repetidos++; return; }
      if (url) urls[url.toLowerCase()] = 1; claves[k] = 1;
      const pr = String(p && p.prioridad || '').toUpperCase() === 'B' ? 'B' : 'A';
      filas.push(['h' + Utilities.getUuid().replace(/-/g, '').slice(0, 8), perfilId, nombre, cut(p && p.cargo, 120), empresa, url, pr, 'Identificado', 0, huntHoy_(), '', huntHoy_(), '', quien || '', cut(p && p.notas, 500), '', '', 0, '', '']);
    });
    if (filas.length) {
      const ini = Math.max(sh.getLastRow(), 1) + 1;
      sh.getRange(ini, 1, filas.length, HUNT_COLS_.length).setNumberFormat('@').setValues(filas);
      filas.forEach(f => huntHist_(perfilId, { nombre: f[2] }, 'Identificado' + (f[4] ? ' · ' + f[4] : ''), quien));
    }
    return { agregadas: filas.length, repetidas: repetidos, sinNombre: sinNombre };
  } finally { lock.releaseLock(); }
}

function huntBuscar_(perfilId, id) {
  const o = huntLeerTodo_().find(x => x.id === id && x.perfil === perfilId);
  if (!o) throw new Error('No encontré a esa persona (¿otra persona la borró?).');
  return o;
}

function huntingActualizar_(perfilId, id, camposJson, quien) {
  let c; try { c = JSON.parse(camposJson); } catch (e) { throw new Error('Datos inválidos.'); }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const o = huntBuscar_(perfilId, id), sh = huntHoja_();
    const cut = (v, n) => limpiarTxt_(v).slice(0, n);
    if ('nombre' in c) { const n = cut(c.nombre, 90); if (n.length < 2) throw new Error('Falta el nombre.'); o.nombre = n; }
    if ('cargo' in c) o.cargo = cut(c.cargo, 120);
    if ('empresa' in c) o.empresa = cut(c.empresa, 90);
    if ('linkedin' in c) o.linkedin = huntUrl_(c.linkedin);
    if ('prioridad' in c) o.prioridad = String(c.prioridad).toUpperCase() === 'B' ? 'B' : 'A';
    if ('notas' in c) o.notas = cut(c.notas, 500);
    if ('seguimiento' in c) o.seguimiento = /^\d{4}-\d{2}-\d{2}$/.test(String(c.seguimiento)) ? c.seguimiento : '';
    o.ultimo_mov = huntHoy_();
    huntEscribir_(sh, o);
    return huntItem_(o);
  } finally { lock.releaseLock(); }
}

/* Cambia la etapa. «Contactado» otra vez = se mandó un seguimiento (los avisos son a los 4, 5 y 7 días; después se deja de insistir). */
function huntingMover_(perfilId, id, estado, quien) {
  if (!(estado in HUNT_NIVEL_)) throw new Error('Etapa inválida.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const o = huntBuscar_(perfilId, id), sh = huntHoja_(), antes = o.estado;
    o.estado = estado;
    o.nivel = Math.max(o.nivel, HUNT_NIVEL_[estado]);
    o.ultimo_mov = huntHoy_();
    if (estado === 'Contactado') {
      if (!o.contactado) o.contactado = huntHoy_();
      if (antes === 'Contactado') o.seguimientos++;
      o.seguimiento = o.seguimientos >= 3 ? '' : huntSumarDias_(o.seguimientos === 0 ? 4 : o.seguimientos === 1 ? 5 : 7);
    } else o.seguimiento = '';
    if (!o.quien) o.quien = quien || '';
    huntEscribir_(sh, o);
    huntHist_(perfilId, o, antes === estado ? 'Seguimiento enviado (' + o.seguimientos + ')' : antes + ' → ' + estado, quien);
    return huntItem_(o);
  } finally { lock.releaseLock(); }
}

// Borra a la persona (por si pide que no la contactemos más): no queda ningún dato suyo en la lista
function huntingBorrar_(perfilId, id, quien) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const o = huntBuscar_(perfilId, id);
    huntHoja_().deleteRow(o.fila);
    huntHist_(perfilId, { nombre: o.nombre }, 'Borrado de la lista de hunting', quien);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/* Mensaje personalizado con IA: nota de conexión (hasta 300 caracteres) y mensaje directo. tipo: primero | seg1 | seg2.
   No inventa datos de su trayectoria, no habla de sueldo ni de criterios internos y respeta la búsqueda ciega. */
// Mail modelo de Blumia para el primer acercamiento (texto de Flavia, 2/11): se usa tal cual, sin IA. [Nombre] y [Firma] se completan solos.
const HUNT_MAIL_ = {
  'director-comercial': {
    asunto: 'Oportunidad ejecutiva | Dirección Comercial – Consultoría de Negocios',
    cuerpo: 'Estimado/a [Nombre]:\n\nEspero que estés muy bien.\n\nDesde Blumia Consulting estamos liderando una búsqueda confidencial para incorporar un/a Director/a Comercial de Proyectos de Consultoría para el sector de Energía y Minería, en una firma de reconocida trayectoria en consultoría de negocios.\n\nSe trata de una posición de nueva creación, con reporte directo al CEO, cuyo principal desafío será impulsar el desarrollo de nuevas soluciones de alto valor para la toma de decisiones estratégicas para el desarrollo de negocios.\n\nBuscamos un perfil con sólida trayectoria en desarrollo comercial B2B, experiencia en ventas consultivas de soluciones complejas, conocimiento del ecosistema energético y minero y una red activa de relacionamiento a nivel C-Level.\n\nLa posición ofrece la oportunidad de asumir un rol estratégico en el crecimiento de la firma, con autonomía, exposición a la alta dirección y participación directa en el desarrollo de negocios.\n\nLa posición está basada en CABA con modalidad híbrida y se requiere disponibilidad para viajes comerciales nacionales e internacionales.\n\nMe gustaría conocer si esta oportunidad podría resultar de tu interés y, en ese caso, coordinar una breve conversación para compartirte más detalles.\n\nDesde ya, muchas gracias por tu tiempo.\n\nSaludos cordiales,\n[Firma]',
    firma: 'Flavia Jaime\nSocia Fundadora\nBlumia Consulting\nExecutive Search\nempleos@blumiaconsulting.com',
    nota: 'Hola [Nombre], desde Blumia Consulting lideramos una búsqueda confidencial de Director/a Comercial en consultoría (Energía y Minería, CABA). Tu perfil podría encajar. ¿Conectamos? Flavia'
  }
};
// La firma es siempre la de Flavia (Ariel, 2/11), no la de quien usa el tablero.
function huntFirma_(perfilId) { const m = HUNT_MAIL_[perfilId]; return (m && m.firma) || 'Flavia Jaime\nBlumia Consulting\nExecutive Search\nempleos@blumiaconsulting.com'; }
function huntMailModelo_(perfilId, o, quien) {
  const t = huntTextos_(perfilId); if (!t) return null;
  const nombre = String(o.nombre || '').trim().split(/\s+/)[0] || '';
  return { asunto: t.asunto, mail: t.cuerpo.replace(/\[Nombre\]/g, nombre) };
}
/* El primer contacto de hunting vive en ⚙ Configurar → Mensajes («hunting-primer-contacto»: primera línea «Asunto: …», después el cuerpo) y la nota corta de conexión en «hunting-nota-conexion».
   Se pueden cambiar sin tocar código (solo esta búsqueda o todas) y con «✨ Redactar con IA». Si la hoja no los tiene, se siembran desde HUNT_MAIL_ (migración de una sola vez). Devuelve {asunto, cuerpo, nota} con [Nombre] para que el tablero lo reemplace. */
function huntTextos_(perfilId) {
  try { asegurarPlantillasHunting_(perfilId); } catch (e) { Logger.log('plantillas de hunting: ' + e); }
  const cfg = getConfig(perfilId);
  const puesto = (typeof cfg.nombre_perfil === 'string' && cfg.nombre_perfil) ? cfg.nombre_perfil : perfilId;
  const arma = t => String(t || '').replace(/\[NOMBRE\]/g, '[Nombre]').replace(/\[PUESTO\]/g, puesto).replace(/\[EMPRESA\]/g, empresaParaCandidatos_(cfg));
  let principal = plantillaTexto_('hunting-primer-contacto', perfilId).texto;
  const nota = arma(plantillaTexto_('hunting-nota-conexion', perfilId).texto);
  const m = HUNT_MAIL_[perfilId];
  if (!String(principal || '').trim() && !m) return null;
  if (!String(principal || '').trim()) return { asunto: m.asunto, cuerpo: m.cuerpo.replace('[Firma]', huntFirma_(perfilId)), nota: m.nota || '' };
  principal = arma(principal);
  const mm = principal.match(/^\s*Asunto:\s*(.+)\n+/i);
  return { asunto: mm ? mm[1].trim() : (m ? m.asunto : 'Oportunidad ejecutiva'), cuerpo: (mm ? principal.slice(mm[0].length) : principal).trim(), nota: nota || (m && m.nota) || '' };
}
const HUNT_GENERAL_ = {
  'hunting-primer-contacto': 'Asunto: Oportunidad ejecutiva | [PUESTO]\n\nEstimado/a [NOMBRE]:\n\nEspero que estés muy bien.\n\nDesde Blumia Consulting estamos liderando una búsqueda confidencial para incorporar un/a [PUESTO] en [EMPRESA].\n\n(Contá acá en 2 o 3 párrafos qué se busca, por qué es una buena oportunidad y dónde está basada la posición. Podés usar «✨ Redactar con IA» con tus instrucciones.)\n\nMe gustaría conocer si esta oportunidad podría resultar de tu interés y, en ese caso, coordinar una breve conversación para compartirte más detalles.\n\nDesde ya, muchas gracias por tu tiempo.\n\nSaludos cordiales,\nFlavia Jaime\nSocia Fundadora\nBlumia Consulting\nExecutive Search\nempleos@blumiaconsulting.com',
  'hunting-nota-conexion': 'Hola [NOMBRE], desde Blumia Consulting lideramos una búsqueda confidencial de [PUESTO]. Tu perfil podría encajar. ¿Conectamos? Flavia'
};
function asegurarPlantillasHunting_(perfilId) {
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Plantillas');
  if (!sh) return;
  const cp = colPerfilPlantillas_(sh);
  const data = sh.getDataRange().getValues();
  const hay = (clave, propia) => data.slice(1).some(r => r[0] === clave && (propia ? String(r[cp - 1] || '').trim() === perfilId : !String(r[cp - 1] || '').trim()));
  const nuevas = [];
  Object.keys(HUNT_GENERAL_).forEach(k => { if (!hay(k, false)) nuevas.push([k, 'Hunting · mail, LinkedIn y WhatsApp', '', HUNT_GENERAL_[k], '']); });
  const m = HUNT_MAIL_[perfilId];
  if (m) {
    if (!hay('hunting-primer-contacto', true)) nuevas.push(['hunting-primer-contacto', 'Hunting · mail, LinkedIn y WhatsApp', '', 'Asunto: ' + m.asunto + '\n\n' + m.cuerpo.replace('[Firma]', huntFirma_(perfilId)).replace(/\[Nombre\]/g, '[NOMBRE]'), perfilId]);
    if (m.nota && !hay('hunting-nota-conexion', true)) nuevas.push(['hunting-nota-conexion', 'Hunting · mail, LinkedIn y WhatsApp', '', m.nota.replace(/\[Nombre\]/g, '[NOMBRE]'), perfilId]);
  }
  if (!nuevas.length) return;
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    const ancho = Math.max(cp, 4);
    nuevas.forEach(f => { const fila = [f[0], f[1], f[2], f[3]]; while (fila.length < ancho) fila.push(''); fila[cp - 1] = f[4]; sh.getRange(sh.getLastRow() + 1, 1, 1, ancho).setNumberFormat('@').setValues([fila]); });
  } finally { lock.releaseLock(); }
}
// Mensaje modelo (InMail / mensaje de LinkedIn) y nota de conexión, sin IA. Quedan [Nombre] para que el tablero lo reemplace al escribir el nombre.
function huntingMailModelo_(perfilId, quien) {
  const t = huntTextos_(perfilId);
  if (!t) throw new Error('Esta búsqueda todavía no tiene un mensaje modelo cargado (⚙ Configurar → Mensajes).');
  return { asunto: t.asunto, mail: t.cuerpo, nota: t.nota };
}

function huntingMensaje_(perfilId, id, tipo, nombrar, quien) {
  const o = huntBuscar_(perfilId, id), cfg = getConfig(perfilId);
  if (tipo === 'mail') {
    const mm = huntMailModelo_(perfilId, o, quien);
    if (!mm) throw new Error('Esta búsqueda todavía no tiene un mail modelo cargado.');
    return { asunto: mm.asunto, mail: mm.mail, conexion: '', mensaje: '' };
  }
  const hero = getHeroPostulacion(perfilId);
  let ubic = ''; try {
    const rows = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Perfiles_Hero').getDataRange().getValues();
    const f = rows.slice(1).find(r => r[0] === perfilId); const k = rows[0].indexOf('ubicacion'); ubic = f && k >= 0 ? String(f[k] || '') : '';
  } catch (e) { }
  const ciega = busquedaCiega_(cfg);
  const empresa = (nombrar && !ciega) ? empresaParaCandidatos_(cfg) : (ciega ? 'una empresa cliente' : 'una empresa cliente de primer nivel');
  const yo = String(quien || '').trim().split(/\s+/)[0] || 'Blumia Consulting';
  const t = tipo === 'seg2' ? 'seg2' : tipo === 'seg1' ? 'seg1' : 'primero';
  const guia = t === 'primero' ? 'Es el PRIMER contacto: presentate, decí por qué la/lo buscás (su perfil encaja con el rol) y proponé una llamada de 15 minutos.'
    : t === 'seg1' ? 'Es un SEGUIMIENTO breve (ya le escribimos y no contestó): recordá el motivo en una frase, sumá un dato nuevo sobre la oportunidad y reiterá la llamada de 15 minutos.'
    : 'Es el ÚLTIMO seguimiento: cortés y sin presión, dejá la puerta abierta y pedí, si no le interesa, que nos lo diga o nos recomiende a alguien.';
  const sistema = 'Sos consultor/a senior de búsqueda ejecutiva en Blumia Consulting (Argentina). Escribís mensajes de LinkedIn en español rioplatense, con «vos», profesionales y cálidos, de parte de «' + yo + '», a una persona que NO se postuló y a la que contactamos directamente. ' + guia + ' ' +
    'REGLAS: usá solo los datos que te paso (no inventes logros, empresas ni fechas de su trayectoria); no menciones sueldo, la banda ni criterios internos de evaluación; ' + (ciega ? 'la búsqueda es CIEGA: no nombres a la empresa cliente ni insinúes cuál es; ' : '') +
    'ofrecé confidencialidad; sin emojis ni frases de plantilla vacías; sin adjetivos exagerados. Devolvé SOLO un JSON válido: {"conexion":"nota para la invitación de conexión, máximo 280 caracteres","mensaje":"mensaje directo, máximo 700 caracteres, con saludo y cierre"}.';
  const usuario = 'Rol: ' + (hero.pos || cfg.nombre) + '. Cliente: ' + empresa + '. Ubicación: ' + (ubic || 'a confirmar') + '. Persona: ' + o.nombre + (o.cargo ? ', ' + o.cargo : '') + (o.empresa ? ' en ' + o.empresa : '') + '. ' + (o.notas ? 'Notas para personalizar: ' + o.notas + '.' : '') + (o.mensaje && t !== 'primero' ? ' Mensaje anterior: ' + o.mensaje : '');
  const ficha = perfilFicha_(cfg, 'candidato').slice(0, 3500);
  const r = extraerJSON_(procesarConClaude(sistema + (ficha ? ' Podés apoyarte en la ficha del puesto acordada con el cliente para decir de qué se trata el rol (misión, desafíos, reporte, modalidad), sin copiar cifras internas.' : ''), usuario + (ficha ? '\n\nFICHA DEL PUESTO:\n' + ficha : ''), 'redactar'));
  const limpio = (v, n) => String(v == null ? '' : v).replace(/[<>]/g, '').trim().slice(0, n);
  const res = { conexion: limpio(r.conexion, 300), mensaje: limpio(r.mensaje, 1900) };
  if (!res.mensaje) throw new Error('La IA no pudo armar el mensaje. Probá de nuevo.');
  if (t === 'primero') {
    const lock = LockService.getScriptLock(); lock.waitLock(20000);
    try { const a = huntBuscar_(perfilId, id); a.mensaje = res.mensaje.slice(0, 450); huntEscribir_(huntHoja_(), a); } finally { lock.releaseLock(); }
  }
  return res;
}

function huntCarpeta_(cfg) {
  const madre = DriveApp.getFolderById(cfg.folder_madre);
  const it = madre.getFoldersByName('Hunting');
  return it.hasNext() ? it.next() : madre.createFolder('Hunting');
}

// Si hay una persona de la lista con ese nombre, la pasa a «En proceso» y la vincula con el candidato
function huntingVincular_(perfilId, nombreCandidato) {
  try {
    const nn = huntNorm_(nombreCandidato); if (!nn) return;
    const o = huntLeerTodo_().find(x => x.perfil === perfilId && x.estado !== 'En proceso' && huntNorm_(x.nombre) === nn);
    if (!o) return;
    o.estado = 'En proceso'; o.nivel = 4; o.candidato = nombreCandidato; o.seguimiento = ''; o.ultimo_mov = huntHoy_();
    huntEscribir_(huntHoja_(), o);
    huntHist_(perfilId, o, 'Pasó al proceso de selección', 'sistema');
  } catch (e) { Logger.log('hunting vincular: ' + e); }
}

/* Sube el PDF del perfil de LinkedIn («Más → Guardar como PDF»): la IA lo lee como a cualquier CV y entra al tablero con «Llegó por: Hunting (LinkedIn)». */
function huntingSubirPdf_(perfilId, id, b64, nombreArchivo, quien) {
  QUIEN_ = quien || '';
  if (!b64 || String(b64).length > 14000000) throw new Error('El archivo está vacío o es muy pesado (máximo 10 MB).');
  const o = huntBuscar_(perfilId, id), cfg = getConfig(perfilId);
  const bytes = Utilities.base64Decode(b64);
  if (bytes.length < 5 || String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]) !== '%PDF') throw new Error('Tiene que ser el PDF del perfil de LinkedIn (Más → Guardar como PDF).');
  const nom = 'LinkedIn - ' + (o.nombre || 'perfil').replace(/[^\w\sáéíóúñÁÉÍÓÚÑ.-]/g, '').slice(0, 60) + '.pdf';
  const file = huntCarpeta_(cfg).createFile(Utilities.newBlob(bytes, 'application/pdf', nom));
  const sheet = getSheet(cfg), recibidos = DriveApp.getFolderById(cfg.folder_recibidos);
  const datos = procesarConChatGPT(file, cfg);
  if (!datos || datos.nombre === 'Error') { try { file.setTrashed(true); } catch (e) { } throw new Error('No pude leer ese PDF. Probá bajarlo de nuevo desde LinkedIn (Más → Guardar como PDF).'); }
  const repetido = esDuplicado(datos, sheet);
  if (!repetido) insertarFilaLibre(sheet, datos, file.getUrl(), new Date(), HUNT_FUENTE_);
  file.moveTo(recibidos);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const a = huntBuscar_(perfilId, id);
    a.estado = 'En proceso'; a.nivel = 4; a.candidato = datos.nombre || a.nombre; a.seguimiento = ''; a.ultimo_mov = huntHoy_();
    huntEscribir_(huntHoja_(), a);
    huntHist_(perfilId, a, repetido ? 'Ya estaba en el tablero: se vinculó' : 'Pasó al proceso de selección (PDF de LinkedIn)', quien);
  } finally { lock.releaseLock(); }
  return { ok: true, candidato: datos.nombre || o.nombre, nota: datos.nivel_ajuste || '', repetido: !!repetido };
}


function huntCarpetaArchivos_(cfg) {
  const madre = DriveApp.getFolderById(cfg.folder_madre);
  const it = madre.getFoldersByName('Hunting - archivos'); // distinta de «Hunting»: esa la barre procesarDriveManual y metería a todos en el tablero
  return it.hasNext() ? it.next() : madre.createFolder('Hunting - archivos');
}

/* Carga a una persona desde su PDF de LinkedIn («Más → Guardar como PDF»): la IA lo lee como a un CV y queda en la lista como «Identificado»
   con nombre, cargo, empresa, link y nota. NO entra al tablero de candidatos hasta que se toca «Pasar a la búsqueda» (huntingPasar_).
   Si ya estaba en la lista (mismo link o mismo nombre), le agrega el PDF a esa ficha. */
function huntingCargarPdf_(perfilId, b64, nombreArchivo, quien) {
  QUIEN_ = quien || '';
  if (!b64 || String(b64).length > 14000000) throw new Error('El archivo está vacío o es muy pesado (máximo 10 MB).');
  const cfg = getConfig(perfilId);
  const bytes = Utilities.base64Decode(b64);
  if (bytes.length < 5 || String.fromCharCode(bytes[0], bytes[1], bytes[2], bytes[3]) !== '%PDF') throw new Error('«' + (nombreArchivo || 'El archivo') + '» no es un PDF. Tiene que ser el perfil de LinkedIn (Más → Guardar como PDF).');
  const nom = 'LinkedIn - ' + String(nombreArchivo || 'perfil').replace(/\.pdf$/i, '').replace(/[^\w\sáéíóúñÁÉÍÓÚÑ.-]/g, '').slice(0, 60) + '.pdf';
  const file = huntCarpetaArchivos_(cfg).createFile(Utilities.newBlob(bytes, 'application/pdf', nom));
  const datos = procesarConChatGPT(file, cfg);
  if (!datos || datos.nombre === 'Error' || datos.nombre === 'No detectado') {
    try { file.setTrashed(true); } catch (e) { }
    throw new Error('No pude leer «' + (nombreArchivo || 'ese PDF') + '». Probá bajarlo de nuevo desde LinkedIn (Más → Guardar como PDF).');
  }
  let json = JSON.stringify(datos); if (json.length > 45000) { datos.resumen = String(datos.resumen || '').slice(0, 800); json = JSON.stringify(datos); }
  const yaEnTablero = !!esDuplicado(datos, getSheet(cfg));
  const cut = (v, n) => limpiarTxt_(v).slice(0, n);
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = huntHoja_();
    const todos = huntLeerTodo_().filter(o => o.perfil === perfilId);
    const url = huntUrl_(datos.linkedin), nn = huntNorm_(datos.nombre);
    const ex = todos.find(o => (url && o.linkedin && o.linkedin.toLowerCase() === url.toLowerCase()) || huntNorm_(o.nombre) === nn);
    if (ex) {
      ex.pdf = file.getUrl(); ex.datos = json; ex.ultimo_mov = huntHoy_();
      if (!ex.cargo) ex.cargo = cut(datos.cargo, 120);
      if (!ex.empresa) ex.empresa = cut(datos.empresa, 90);
      if (!ex.linkedin) ex.linkedin = url;
      huntEscribir_(sh, ex);
      huntHist_(perfilId, ex, 'Se agregó su PDF de LinkedIn', quien);
      return { id: ex.id, nombre: ex.nombre, existente: true, yaEnTablero: yaEnTablero, nota: datos.nivel_ajuste || '' };
    }
    const id = 'h' + Utilities.getUuid().replace(/-/g, '').slice(0, 8);
    const fila = [id, perfilId, cut(datos.nombre, 90), cut(datos.cargo, 120), cut(datos.empresa, 90), url, 'A', 'Identificado', 0, huntHoy_(), '', huntHoy_(), '', quien || '', '', '', '', 0, file.getUrl(), json];
    const ini = Math.max(sh.getLastRow(), 1) + 1;
    sh.getRange(ini, 1, 1, HUNT_COLS_.length).setNumberFormat('@').setValues([fila]);
    huntHist_(perfilId, { nombre: fila[2] }, 'Identificado (PDF de LinkedIn)' + (fila[4] ? ' · ' + fila[4] : ''), quien);
    return { id: id, nombre: fila[2], existente: false, yaEnTablero: yaEnTablero, nota: datos.nivel_ajuste || '' };
  } finally { lock.releaseLock(); }
}

/* «Pasar a la búsqueda»: la persona (que ya tiene su PDF cargado) entra al tablero de candidatos con «Llegó por: Hunting (LinkedIn)», sin volver a leer nada. */
function huntingPasar_(perfilId, id, quien) {
  QUIEN_ = quien || '';
  const o = huntBuscar_(perfilId, id), cfg = getConfig(perfilId);
  if (o.estado === 'En proceso') throw new Error('Ya está en el tablero de candidatos.');
  if (!o.pdf || !o.datos) throw new Error('A esta persona le falta el PDF de LinkedIn. Subilo con «Subir PDF».');
  let datos; try { datos = JSON.parse(o.datos); } catch (e) { throw new Error('Los datos del PDF están dañados. Subí el PDF de nuevo.'); }
  const sheet = getSheet(cfg);
  const repetido = esDuplicado(datos, sheet);
  if (!repetido) insertarFilaLibre(sheet, datos, o.pdf, new Date(), HUNT_FUENTE_);
  try {
    const m = String(o.pdf).match(/[-\w]{25,}/);
    if (m) DriveApp.getFileById(m[0]).moveTo(DriveApp.getFolderById(cfg.folder_recibidos));
  } catch (e) { Logger.log('hunting pasar (mover pdf): ' + e); }
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const a = huntBuscar_(perfilId, id);
    a.estado = 'En proceso'; a.nivel = 4; a.candidato = datos.nombre || a.nombre; a.seguimiento = ''; a.ultimo_mov = huntHoy_();
    huntEscribir_(huntHoja_(), a);
    huntHist_(perfilId, a, repetido ? 'Ya estaba en el tablero: se vinculó' : 'Pasó a la búsqueda (PDF de LinkedIn)', quien);
  } finally { lock.releaseLock(); }
  return { ok: true, candidato: datos.nombre || o.nombre, nota: datos.nivel_ajuste || '', repetido: !!repetido };
}

/* Ficha de una persona de la lista: lo que la IA leyó de su PDF (nota, por qué, resumen, formación, zona, contacto) y el link al PDF.
   No incluye «requisitos» (ahí van edad y género, que no se muestran). */
function huntingFicha_(perfilId, id) {
  const o = huntBuscar_(perfilId, id);
  let d = {}; try { d = JSON.parse(o.datos || '{}'); } catch (e) { }
  const v = x => { x = String(x == null ? '' : x).trim(); return /^sin dato$/i.test(x) ? '' : x; };
  return {
    id: o.id, nombre: o.nombre, cargo: o.cargo, empresa: o.empresa, linkedin: o.linkedin, estado: o.estado, notas: o.notas,
    pdf: o.pdf || '', tienePdf: !!(o.pdf && o.datos),
    nota: v(d.nivel_ajuste), justificacion: v(d.justificacion), resumen: v(d.resumen),
    formacion: v(d.formacion), zona: v(d.zona), email: v(d.email), telefono: v(d.telefono)
  };
}


/* ============================================================
   EMPRESAS OBJETIVO (3/11, Ariel: «esta búsqueda no tenemos ni un CV»): el mapa de empresas donde buscar gente.
   Se pega la tabla del Excel de hunting (con encabezados) y se trabaja empresa por empresa: estado, personas cargadas y búsqueda en LinkedIn.
   Pestaña «Hunting_Empresas» (una fila por empresa y búsqueda). Las personas siguen en «Hunting» y se atan a la empresa por el nombre.
   ============================================================ */
const HOJA_HUNTE_ = 'Hunting_Empresas';
const HUNTE_COLS_ = ['id', 'perfil', 'empresa', 'categoria', 'sectores', 'competidora', 'prioridad', 'evidencia', 'nota', 'estado', 'ultimo_mov', 'quien'];
const HUNTE_ESTADOS_ = ['Sin revisar', 'En búsqueda', 'Revisada', 'Descartada'];

function huntEHoja_() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let sh = ss.getSheetByName(HOJA_HUNTE_);
  if (!sh) {
    sh = ss.insertSheet(HOJA_HUNTE_);
    sh.getRange(1, 1, 1, HUNTE_COLS_.length).setValues([HUNTE_COLS_]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.getRange(2, 1, 1000, HUNTE_COLS_.length).setNumberFormat('@');
  }
  return sh;
}

function huntELeer_() {
  const sh = huntEHoja_();
  if (sh.getLastRow() < 2) return [];
  const v = sh.getRange(2, 1, sh.getLastRow() - 1, HUNTE_COLS_.length).getValues();
  const out = [];
  v.forEach((r, i) => {
    if (!String(r[0]).trim()) return;
    const o = { fila: i + 2 };
    HUNTE_COLS_.forEach((c, k) => { o[c] = r[k] instanceof Date ? Utilities.formatDate(r[k], 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd') : String(r[k] == null ? '' : r[k]); });
    out.push(o);
  });
  return out;
}

// Una persona pertenece a una empresa si los nombres coinciden o uno contiene al otro («NTT Data» ~ «NTT Data (ex Everis)»)
function huntMismaEmpresa_(a, b) {
  const x = huntNorm_(a), y = huntNorm_(b);
  if (!x || !y) return false;
  return x === y || (x.length >= 4 && y.length >= 4 && (x.indexOf(y) >= 0 || y.indexOf(x) >= 0));
}

function huntingEmpresasListar_(perfilId) {
  const emp = huntELeer_().filter(e => e.perfil === perfilId);
  const pers = huntLeerTodo_().filter(o => o.perfil === perfilId);
  const items = emp.map(e => {
    const ps = pers.filter(p => huntMismaEmpresa_(p.empresa, e.empresa));
    const x = Object.assign({}, e); delete x.fila;
    x.personas = ps.length; x.contactadas = ps.filter(p => p.nivel >= 1).length; x.interesadas = ps.filter(p => p.nivel >= 3).length;
    return x;
  });
  const alta = items.filter(e => e.prioridad === 'Alta');
  const stats = {
    total: items.length, alta: alta.length, altaRevisadas: alta.filter(e => e.estado === 'Revisada').length,
    enBusqueda: items.filter(e => e.estado === 'En búsqueda').length, revisadas: items.filter(e => e.estado === 'Revisada').length,
    personas: items.reduce((n, e) => n + e.personas, 0), conPersonas: items.filter(e => e.personas > 0).length
  };
  return { items: items, stats: stats };
}

// Separa texto copiado de Excel/Sheets (tabuladores, con comillas para las celdas con saltos de línea)
function huntParseTsv_(txt) {
  const filas = []; let fila = [], cel = '', q = false;
  const t = String(txt || '').replace(/\r\n?/g, '\n');
  for (let i = 0; i < t.length; i++) {
    const c = t[i];
    if (q) { if (c === '"') { if (t[i + 1] === '"') { cel += '"'; i++; } else q = false; } else cel += c; }
    else if (c === '"' && cel === '') q = true;
    else if (c === '\t') { fila.push(cel); cel = ''; }
    else if (c === '\n') { fila.push(cel); filas.push(fila); fila = []; cel = ''; }
    else cel += c;
  }
  fila.push(cel); filas.push(fila);
  return filas.filter(f => f.some(x => String(x).trim()));
}

/* Carga empresas pegando la tabla (con encabezados: Empresa, Categoría, Sectores que atiende, Competidor directo…, Prioridad…, Evidencia, Nota; sin encabezados:
   Empresa | Categoría | Prioridad | Competidora | Nota). Si la tabla trae «Candidato identificado» / «Cargo actual» / «Perfil de LinkedIn», esas personas pasan a la lista de personas. */
function huntingEmpresasAgregar_(perfilId, texto, quien) {
  getConfig(perfilId);
  const filas = huntParseTsv_(texto);
  if (!filas.length) throw new Error('No hay nada para cargar. Copiá la tabla del Excel (con la fila de encabezados) y pegala.');
  if (filas.length > 400) throw new Error('Son demasiadas filas de una vez (máximo 400).');
  const n = h => huntNorm_(h);
  const cab = filas[0].map(n);
  const tieneCab = cab.some(h => h === 'empresa');
  const col = {};
  if (tieneCab) {
    cab.forEach((h, i) => {
      if (h === 'empresa' && col.empresa == null) col.empresa = i;
      else if (h.indexOf('categor') === 0 && col.categoria == null) col.categoria = i;
      else if (h.indexOf('sectores') === 0 && col.sectores == null) col.sectores = i;
      else if (h.indexOf('competidor') === 0 && col.competidora == null) col.competidora = i;
      else if (h.indexOf('prioridad') === 0 && col.prioridad == null) col.prioridad = i;
      else if (h.indexOf('evidencia') === 0 && col.evidencia == null) col.evidencia = i;
      else if (h === 'nota' && col.nota == null) col.nota = i;
      else if (h.indexOf('candidato identificado') === 0) col.cand = i;
      else if (h.indexOf('cargo actual') === 0) col.cargo = i;
      else if (h.indexOf('perfil de linkedin') === 0) col.li = i;
      else if (h === 'estado' && col.estadoX == null) col.estadoX = i;
    });
  } else { col.empresa = 0; col.categoria = 1; col.prioridad = 2; col.competidora = 3; col.nota = 4; }
  if (col.empresa == null) throw new Error('No encuentro la columna «Empresa».');
  const cut = (v, k) => limpiarTxt_(v).slice(0, k);
  const get = (f, k) => (col[k] == null ? '' : String(f[col[k]] == null ? '' : f[col[k]]));
  const prio = v => { const x = huntNorm_(v); return x.indexOf('alta') === 0 ? 'Alta' : x.indexOf('media') === 0 ? 'Media' : x.indexOf('baja') === 0 ? 'Baja' : ''; };
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  let agregadas = 0, repetidas = 0; const personas = [];
  try {
    const sh = huntEHoja_();
    const ex = huntELeer_().filter(e => e.perfil === perfilId);
    const vistos = {}; ex.forEach(e => { vistos[huntNorm_(e.empresa)] = 1; });
    const nuevas = [];
    filas.slice(tieneCab ? 1 : 0).forEach(f => {
      const nombre = cut(get(f, 'empresa'), 120);
      if (nombre.length < 2) return;
      const k = huntNorm_(nombre);
      const cand = cut(get(f, 'cand'), 90);
      if (cand) personas.push({ nombre: cand, cargo: cut(get(f, 'cargo'), 120), empresa: nombre, linkedin: get(f, 'li'), prioridad: 'A' });
      if (vistos[k]) { repetidas++; return; }
      vistos[k] = 1;
      const comp = /^(si|sí|s)$/i.test(String(get(f, 'competidora')).trim()) ? 'Sí' : '';
      const nota = [cut(get(f, 'nota'), 300), cut(get(f, 'evidencia'), 120)].filter(x => x && !/^fuente web$/i.test(x)).join(' · ').slice(0, 400);
      const est = HUNTE_ESTADOS_.indexOf(get(f, 'estadoX')) >= 0 ? get(f, 'estadoX') : 'Sin revisar';
      nuevas.push(['e' + Utilities.getUuid().replace(/-/g, '').slice(0, 8), perfilId, nombre, cut(get(f, 'categoria'), 120), cut(get(f, 'sectores'), 150), comp, prio(get(f, 'prioridad')), '', nota, est, huntHoy_(), quien || '']);
    });
    if (nuevas.length) {
      const ini = Math.max(sh.getLastRow(), 1) + 1;
      sh.getRange(ini, 1, nuevas.length, HUNTE_COLS_.length).setNumberFormat('@').setValues(nuevas);
      agregadas = nuevas.length;
      try { registrarHistorial_(getConfig(perfilId).hoja_nombre, 'Hunting', 'Empresas objetivo', agregadas + ' empresas cargadas', quien); } catch (e) { }
    }
  } finally { lock.releaseLock(); }
  let pers = null;
  if (personas.length) { try { pers = huntingAgregar_(perfilId, JSON.stringify(personas), quien); } catch (e) { Logger.log('empresas → personas: ' + e); } }
  return { agregadas: agregadas, repetidas: repetidas, personas: pers ? pers.agregadas : 0 };
}

function huntingEmpresaEstado_(perfilId, id, estado, quien) {
  if (HUNTE_ESTADOS_.indexOf(estado) < 0) throw new Error('Estado inválido.');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const e = huntELeer_().find(x => x.id === id && x.perfil === perfilId);
    if (!e) throw new Error('No encontré esa empresa (¿otra persona la borró?).');
    e.estado = estado; e.ultimo_mov = huntHoy_(); if (!e.quien) e.quien = quien || '';
    huntEHoja_().getRange(e.fila, 1, 1, HUNTE_COLS_.length).setNumberFormat('@').setValues([HUNTE_COLS_.map(c => e[c] == null ? '' : e[c])]);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function huntingEmpresaBorrar_(perfilId, id, quien) {
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const e = huntELeer_().find(x => x.id === id && x.perfil === perfilId);
    if (!e) throw new Error('No encontré esa empresa.');
    huntEHoja_().deleteRow(e.fila);
    return { ok: true };
  } finally { lock.releaseLock(); }
}
