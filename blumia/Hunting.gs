/* ============================================================
   HUNTING (30/10, pedido de Ariel para la búsqueda del Director Comercial): personas a las que se aborda directamente, antes de que sean candidatos.
   Viven en la pestaña «Hunting» de la planilla (una fila por persona y búsqueda) y se manejan desde la pestaña «Hunting» del tablero (HuntingPanel.html).
   Etapas: Identificado → Contactado → Respondió → Interesado → En proceso. Cierres: No le interesa · Sin respuesta.
   Cuando dice «me interesa» se sube su perfil de LinkedIn (PDF): entra al tablero como candidato con «Llegó por: Hunting (LinkedIn)» y sigue el flujo normal.
   El embudo (identificados → contactados → respondieron → interesados → en proceso) sale en el Resumen y en el informe de la búsqueda.
   Estas personas NO están en la hoja de candidatos hasta que se interesan: no ensucian los números del proceso.
   ============================================================ */

const HOJA_HUNT_ = 'Hunting';
const HUNT_COLS_ = ['id', 'perfil', 'nombre', 'cargo', 'empresa', 'linkedin', 'prioridad', 'estado', 'nivel', 'creado', 'contactado', 'ultimo_mov', 'seguimiento', 'quien', 'notas', 'mensaje', 'candidato', 'seguimientos'];
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

function huntItem_(o) { const x = Object.assign({}, o); delete x.fila; return x; }

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
      filas.push(['h' + Utilities.getUuid().replace(/-/g, '').slice(0, 8), perfilId, nombre, cut(p && p.cargo, 120), empresa, url, pr, 'Identificado', 0, huntHoy_(), '', huntHoy_(), '', quien || '', cut(p && p.notas, 500), '', '', 0]);
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
function huntingMensaje_(perfilId, id, tipo, nombrar, quien) {
  const o = huntBuscar_(perfilId, id), cfg = getConfig(perfilId);
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
  if (!datos) { try { file.setTrashed(true); } catch (e) { } throw new Error('No pude leer ese PDF. Probá bajarlo de nuevo desde LinkedIn (Más → Guardar como PDF).'); }
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
