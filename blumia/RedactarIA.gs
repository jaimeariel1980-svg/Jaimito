/* Redactar con IA, mensaje por mensaje (4/11, Ariel: «que sean independientes la edición del texto y las instrucciones de cada uno» y «que le dé bola a las instrucciones»).
   Cada mensaje (y cada texto de Difusión) tiene SUS instrucciones, que se guardan en la pestaña `Instrucciones_IA` (perfil, clave, texto). La IA devuelve el texto nuevo para REVISAR (no guarda nada)
   y una lista corta de qué hizo con cada instrucción. Las instrucciones van primero en el pedido, se marcan como obligatorias y mandan sobre el estilo y sobre los límites de largo. */
const MENSAJE_MARCAS_ = ['[NOMBRE]', '[PUESTO]', '[EMPRESA]', '[ENTREVISTADOR]', '[LINK]'];
const HOJA_INSTR_IA_ = 'Instrucciones_IA';

function instrIAHoja_(crear) {
  const ss = ssAbrir_(SHEET_MASTER_ID);
  let sh = ss.getSheetByName(HOJA_INSTR_IA_);
  if (!sh && crear) { sh = ss.insertSheet(HOJA_INSTR_IA_); sh.getRange(1, 1, 1, 5).setValues([['perfil', 'clave', 'texto', 'quien', 'fecha']]).setFontWeight('bold'); sh.setFrozenRows(1); }
  return sh;
}
// { clave: texto } de una búsqueda
function instrIALeer_(perfilId) {
  const out = {};
  try {
    const sh = instrIAHoja_(false);
    if (!sh || sh.getLastRow() < 2) return out;
    sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues().forEach(r => { if (r[0] === perfilId && r[1]) out[String(r[1])] = String(r[2] || ''); });
  } catch (e) { }
  return out;
}
function instrIAGuardar_(perfilId, clave, texto) {
  getConfig(perfilId);
  clave = String(clave || '').replace(/[^\w\-]/g, '').slice(0, 60);
  if (!clave) throw new Error('Falta indicar a qué mensaje corresponden las instrucciones.');
  texto = limpiarTxt_(texto).slice(0, 1500);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = instrIAHoja_(true), n = sh.getLastRow();
    let fila = 0;
    if (n >= 2) sh.getRange(2, 1, n - 1, 2).getValues().forEach((r, i) => { if (!fila && r[0] === perfilId && r[1] === clave) fila = i + 2; });
    const quien = (typeof SES_ !== 'undefined' && SES_ && SES_.nombre) ? SES_.nombre : '';
    if (fila) sh.getRange(fila, 3, 1, 3).setNumberFormat('@').setValues([[texto, quien, new Date().toISOString()]]);
    else sh.getRange(n + 1, 1, 1, 5).setNumberFormat('@').setValues([[perfilId, clave, texto, quien, new Date().toISOString()]]);
  } finally { lock.releaseLock(); }
  return { ok: true };
}

// Bloque común de reglas de obediencia (mensajes y difusión)
function reglasInstrIA_() {
  return 'INSTRUCCIONES DEL EQUIPO: son la parte MÁS IMPORTANTE del pedido. Aplicá CADA una al pie de la letra y de forma bien visible en el texto: si piden cambiar el tono, el largo, el orden o el contenido, el texto nuevo tiene que verse claramente distinto del actual en eso (no devuelvas el mismo texto con retoques mínimos). ' +
    'Mandan sobre el estilo por defecto y sobre cualquier límite de largo que figure abajo. Solo dejá de aplicar una instrucción si te obliga a inventar datos, a nombrar a la empresa cliente cuando no corresponde o a tocar datos personales protegidos: en ese caso aplicá todas las demás y explicalo en «aviso». ' +
    'Antes de responder, revisá una por una que cada instrucción esté cumplida.';
}

function redactarMensajeIA_(perfilId, clave, textoActual, instrucciones, canal) {
  const cfg = getConfig(perfilId);
  clave = String(clave || '').slice(0, 60);
  const actual = String(textoActual == null ? '' : textoActual).replace(/[<>]/g, '').trim().slice(0, 4000);
  const instr = limpiarTxt_(instrucciones).slice(0, 1500);
  const esMail = /mail/i.test(String(canal || ''));
  const puesto = (typeof cfg.nombre_perfil === 'string' && cfg.nombre_perfil) ? cfg.nombre_perfil : perfilId;
  const sistema = 'Sos redactor/a de mensajes de selección de personal de Blumia Consulting (Argentina). Escribís el texto de una plantilla que el equipo le manda a candidatos. ' +
    (esMail ? 'Es un MAIL: saludo, 2 o 3 párrafos cortos y cierre; por defecto sin asunto (el asunto se arma aparte). ' : 'Es un WHATSAPP: tono cercano y profesional, por defecto breve (máx. 600 caracteres) y con pocos emojis. ') +
    (/^hunting/.test(clave) ? 'Este mensaje es el PRIMER CONTACTO DE HUNTING (mail, mensaje de LinkedIn o WhatsApp a una persona que todavía no se postuló): conservá la primera línea «Asunto: …» y la firma del texto actual, tono ejecutivo y cordial, sin sueldos y sin nombrar a la empresa cliente. En [NOMBRE] va el nombre de la persona. ' : '') +
    'Español rioplatense (vos). ' +
    'Marcas que se completan solas al enviar y que tenés que escribir TAL CUAL, entre corchetes: [NOMBRE] (nombre de la persona), [PUESTO], [EMPRESA], [ENTREVISTADOR] (quien envía y firma) y [LINK] (link de agenda). ' +
    'Conservá todas las marcas que ya tiene el texto actual (si tiene [LINK], el mensaje nuevo también). Nunca escribas el nombre real de la empresa: usá [EMPRESA]. ' +
    'No inventes datos (fechas, horarios, sueldos, condiciones) que no estén en el texto actual o en las instrucciones. No preguntes ni insinúes nada sobre edad, género, estado civil, hijos ni otros datos personales protegidos. ' +
    (instr ? reglasInstrIA_() + ' ' : 'Como no hay instrucciones, mejorá la claridad y el tono sin cambiar el sentido. ') +
    'Devolvé SOLO un JSON válido: {"texto":"el mensaje completo, con saltos de línea \\n","cambios":["una línea por cada instrucción: qué hiciste"],"aviso":"solo si no pudiste cumplir algo y por qué; si no, vacío"}.';
  const usuario = (instr ? 'INSTRUCCIONES (OBLIGATORIAS, cumplilas todas):\n' + instr + '\n\n' : '') +
    'MENSAJE: ' + clave + ' (' + (esMail ? 'mail' : 'WhatsApp') + ')\nBÚSQUEDA: ' + puesto +
    '\n\nTEXTO ACTUAL:\n' + (actual || '(vacío: armalo de cero según el nombre del mensaje)');
  const r = extraerJSON_(procesarConClaude(sistema, usuario, 'analisis'));
  const texto = String(r.texto == null ? '' : r.texto).replace(/[<>]/g, '').trim().slice(0, 3000);
  if (texto.length < 10) throw new Error('La IA no devolvió un texto. Probá de nuevo.');
  const faltan = MENSAJE_MARCAS_.filter(m => actual.indexOf(m) >= 0 && texto.indexOf(m) < 0);
  const avisos = [];
  if (faltan.length) avisos.push('Ojo: al texto nuevo le faltan ' + faltan.join(', ') + ', que tenía el anterior. Agregalo(s) antes de guardar o tocá «Deshacer».');
  if (r.aviso) avisos.push(String(r.aviso).slice(0, 400));
  return { texto: texto, cambios: (Array.isArray(r.cambios) ? r.cambios : []).map(c => limpiarTxt_(c).slice(0, 200)).filter(Boolean).slice(0, 8), aviso: avisos.join(' ') };
}
