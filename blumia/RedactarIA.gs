/* Redactar los mensajes a candidatos con IA (4/11, Ariel: «metele IA a la parte de mensajes para redactarlos y para darle instrucciones»).
   Toma el texto actual de una plantilla (o la deja armar de cero) y las instrucciones de la persona, y devuelve el texto nuevo para REVISAR: no guarda nada, se guarda con el botón de siempre. */
const MENSAJE_MARCAS_ = ['[NOMBRE]', '[PUESTO]', '[EMPRESA]', '[ENTREVISTADOR]', '[LINK]'];

function redactarMensajeIA_(perfilId, clave, textoActual, instrucciones, canal) {
  const cfg = getConfig(perfilId);
  clave = String(clave || '').slice(0, 60);
  const actual = String(textoActual == null ? '' : textoActual).replace(/[<>]/g, '').trim().slice(0, 4000);
  const instr = limpiarTxt_(instrucciones).slice(0, 1500);
  const esMail = /mail/i.test(String(canal || ''));
  const puesto = (typeof cfg.nombre_perfil === 'string' && cfg.nombre_perfil) ? cfg.nombre_perfil : perfilId;
  const sistema = 'Sos redactor/a de mensajes de selección de personal de Blumia Consulting (Argentina). Escribís el texto de una plantilla que el equipo le manda a candidatos. ' +
    (esMail ? 'Es un MAIL: saludo, 2 o 3 párrafos cortos y cierre; sin asunto (el asunto se arma aparte); ' : 'Es un WHATSAPP: tono cercano y profesional, breve (máx. 600 caracteres salvo que las instrucciones pidan otra cosa), pocos emojis; ') +
    'español rioplatense (vos). ' +
    'Marcas que se completan solas al enviar y que tenés que escribir TAL CUAL, entre corchetes: [NOMBRE] (nombre de la persona), [PUESTO], [EMPRESA], [ENTREVISTADOR] (quien envía y firma) y [LINK] (link de agenda). ' +
    'Conservá todas las marcas que ya tiene el texto actual (si tiene [LINK], el mensaje nuevo también). Nunca escribas el nombre real de la empresa: usá [EMPRESA]. ' +
    'No inventes datos (fechas, horarios, sueldos, condiciones) que no estén en el texto actual o en las instrucciones. No preguntes ni insinúes nada sobre edad, género, estado civil, hijos ni otros datos personales protegidos. ' +
    'Si hay instrucciones del equipo, aplicalas con prioridad sobre el estilo por defecto (pero respetando todo lo anterior). Si no hay instrucciones, mejorá la claridad y el tono sin cambiar el sentido. ' +
    'Devolvé SOLO un JSON válido: {"texto":"el mensaje completo, con saltos de línea \\n"}.';
  const usuario = 'MENSAJE: ' + clave + ' (' + (esMail ? 'mail' : 'WhatsApp') + ')\nBÚSQUEDA: ' + puesto +
    '\n\nTEXTO ACTUAL:\n' + (actual || '(vacío: armalo de cero según el nombre del mensaje)') +
    '\n\nINSTRUCCIONES DEL EQUIPO:\n' + (instr || '(ninguna)');
  const r = extraerJSON_(procesarConClaude(sistema, usuario, 'redactar'));
  const texto = String(r.texto == null ? '' : r.texto).replace(/[<>]/g, '').trim().slice(0, 3000);
  if (texto.length < 10) throw new Error('La IA no devolvió un texto. Probá de nuevo.');
  const faltan = MENSAJE_MARCAS_.filter(m => actual.indexOf(m) >= 0 && texto.indexOf(m) < 0);
  return { texto: texto, aviso: faltan.length ? 'Ojo: al texto nuevo le faltan ' + faltan.join(', ') + ', que tenía el anterior. Agregalo(s) antes de guardar o tocá «Deshacer».' : '' };
}
