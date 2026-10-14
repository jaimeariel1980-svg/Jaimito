let QUIEN_ = '';

function usuarios_() {
  const cache = CacheService.getScriptCache();
  const c = cache.get('USUARIOS');
  if (c) return JSON.parse(c);
  let lista = [];
  try {
    const sh = hoja_('Usuarios');
    if (sh) lista = sh.getDataRange().getValues().slice(1).filter(r => String(r[0]).trim()).map(r => ({ nombre: String(r[0]).trim(), pin: String(r[1] || '').trim(), activo: !/^no$/i.test(String(r[2] || '').trim()), rol: /^admin/i.test(String(r[4] || '').trim()) ? 'Admin' : (/^(lectura|solo)/i.test(String(r[4] || '').trim()) ? 'Lectura' : 'Entrevistador') }));
  } catch (e) { }
  if (lista.length && !lista.some(u => u.rol === 'Admin')) lista[0].rol = 'Admin';
  lista.forEach(u => { u.admin = u.rol === 'Admin'; });
  const prop = lista.find(u => u.rol === 'Admin');
  if (prop) prop.propietario = true;
  cache.put('USUARIOS', JSON.stringify(lista), 30);
  return lista;
}

function exigePin_() { return usuarios_().some(u => u.pin); }

function accesoLog_(usuario, accion, detalle) {
  try {
    const ss = ss_();
    let sh = ss.getSheetByName('Accesos');
    if (!sh) { sh = ss.insertSheet('Accesos'); asegurarEncabezados_(sh, ['Fecha', 'Usuario', 'Acción', 'Detalle']); }
    sh.appendRow([new Date(), usuario || '', accion, detalle || '']);
  } catch (e) { }
}

function verificarPin_(u, pin) {
  const cache = CacheService.getScriptCache();
  const key = 'FALLOS_' + u.nombre.toLowerCase().replace(/\s+/g, '_');
  const n = Number(cache.get(key) || 0);
  if (n >= 5) throw new Error('LOGIN: Demasiados intentos. Esperá 15 minutos.');
  if (u.pin !== String(pin || '').trim()) {
    cache.put(key, String(n + 1), 900);
    accesoLog_(u.nombre, 'Intento fallido', 'PIN incorrecto (' + (n + 1) + ')');
    if (n + 1 === 5) {
      accesoLog_(u.nombre, 'Bloqueo', 'Usuario bloqueado 15 minutos por intentos fallidos');
      try { const to = ajustes_().correo_aviso; if (to) MailApp.sendEmail(to, 'Alerta de seguridad · Atracción y Elección', 'Hubo 5 intentos fallidos de PIN para el usuario "' + u.nombre + '". Quedó bloqueado 15 minutos. Si no fue esa persona, reiniciá su PIN desde el tablero.'); } catch (e) { }
    }
    throw new Error('LOGIN: PIN incorrecto');
  }
  if (n) cache.remove(key);
}

let ROL_ADMIN_ = false;
let ROL_ = 'Entrevistador';

function admin_(k) {
  const p = String(k).split('|');
  if (p[0] !== prop_('ADMIN_KEY')) throw new Error('No autorizado');
  const nombre = String(p[1] || '').replace(/[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9 ]/g, '').trim().slice(0, 30);
  if (exigePin_()) {
    const u = usuarios_().find(x => x.activo && x.nombre.toLowerCase() === nombre.toLowerCase());
    if (!u) { if (nombre) accesoLog_(nombre, 'Intento fallido', 'Usuario inexistente o bloqueado'); throw new Error('LOGIN: Usuario no habilitado. Pedíselo al administrador.'); }
    verificarPin_(u, p[2]);
    QUIEN_ = u.nombre;
    ROL_ = u.rol;
    ROL_ADMIN_ = !!u.admin;
  } else {
    QUIEN_ = nombre;
    ROL_ = 'Entrevistador';
    ROL_ADMIN_ = false;
  }
}

function escritor_(k) {
  admin_(k);
  if (ROL_ === 'Lectura') throw new Error('Tu usuario es de solo lectura');
}

function soloAdmin_(k) {
  admin_(k);
  if (!exigePin_() || !ROL_ADMIN_) throw new Error('Solo el administrador puede hacer esto');
}

function nombreUsuario_(n) { return String(n || '').replace(/[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9 ]/g, '').replace(/\s+/g, ' ').trim().slice(0, 30); }

function protegePropietario_(nombre) {
  const t = usuarios_().find(u => u.nombre.toLowerCase() === String(nombre).toLowerCase());
  if (t && t.propietario && t.nombre.toLowerCase() !== QUIEN_.toLowerCase()) throw new Error('Solo el propietario puede modificar su propio usuario');
}

function pinAzar_() { return String(1000 + Math.floor(Math.random() * 9000)); }

function hojaUsuarios_() {
  const ss = ss_();
  let sh = ss.getSheetByName('Usuarios');
  if (!sh) sh = ss.insertSheet('Usuarios');
  asegurarEncabezados_(sh, ['Nombre', 'PIN', 'Activo', 'Notas', 'Rol']);
  return sh;
}

function filaUsuario_(sh, nombre) {
  const d = sh.getDataRange().getValues();
  const i = d.findIndex((r, j) => j > 0 && String(r[0]).trim().toLowerCase() === String(nombre).trim().toLowerCase());
  return i === -1 ? 0 : i + 1;
}

function getLogin(clave) {
  if (String(clave).split('|')[0] !== prop_('ADMIN_KEY')) throw new Error('No autorizado');
  return { pin: exigePin_() };
}

function registrarUsuario(clave, nombre, pin) {
  if (String(clave).split('|')[0] !== prop_('ADMIN_KEY')) throw new Error('No autorizado');
  nombre = nombreUsuario_(nombre);
  pin = String(pin || '').trim();
  if (nombre.length < 2) throw new Error('LOGIN: Escribí tu nombre');
  if (!/^\d{4,6}$/.test(pin)) throw new Error('LOGIN: El PIN tiene que ser de 4 a 6 números');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    CacheService.getScriptCache().remove('USUARIOS');
    if (usuarios_().length) throw new Error('LOGIN: Solo el administrador puede crear usuarios. Pedile el tuyo.');
    hojaUsuarios_().appendRow([nombre, pin, 'SI', '', 'Admin']);
    CacheService.getScriptCache().remove('USUARIOS');
  } finally { lock.releaseLock(); }
  registrarHistorial_('', nombre, 'Usuario', 'Creó el usuario administrador', nombre);
  return true;
}

function listarUsuarios(k) {
  soloAdmin_(k);
  return usuarios_().map(u => ({ nombre: u.nombre, activo: u.activo, admin: u.admin, rol: u.rol, propietario: !!u.propietario }));
}

function rolValido_(r) { r = String(r || 'Entrevistador'); return ['Admin', 'Entrevistador', 'Lectura'].indexOf(r) > -1 ? r : 'Entrevistador'; }

function crearUsuario(k, nombre, rol) {
  soloAdmin_(k);
  nombre = nombreUsuario_(nombre);
  if (nombre.length < 2) throw new Error('Escribí el nombre');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = hojaUsuarios_();
    if (filaUsuario_(sh, nombre)) throw new Error('Ese usuario ya existe');
    const pin = pinAzar_();
    sh.appendRow([nombre, pin, 'SI', '', rolValido_(rol)]);
    CacheService.getScriptCache().remove('USUARIOS');
    registrarHistorial_('', nombre, 'Usuario', 'Usuario creado por ' + QUIEN_ + ' (' + rolValido_(rol) + ')', QUIEN_);
    return { nombre: nombre, pin: pin };
  } finally { lock.releaseLock(); }
}

function cambiarRol(k, nombre, rol) {
  soloAdmin_(k);
  if (String(nombre).toLowerCase() === QUIEN_.toLowerCase()) throw new Error('No podés cambiarte el rol a vos mismo');
  protegePropietario_(nombre);
  const sh = hojaUsuarios_();
  const f = filaUsuario_(sh, nombre);
  if (!f) throw new Error('Usuario no encontrado');
  sh.getRange(f, 5).setValue(rolValido_(rol));
  CacheService.getScriptCache().remove('USUARIOS');
  registrarHistorial_('', nombre, 'Usuario', 'Rol cambiado a ' + rolValido_(rol) + ' por ' + QUIEN_, QUIEN_);
  return true;
}

function reiniciarPin(k, nombre) {
  soloAdmin_(k);
  protegePropietario_(nombre);
  const sh = hojaUsuarios_();
  const f = filaUsuario_(sh, nombre);
  if (!f) throw new Error('Usuario no encontrado');
  const pin = pinAzar_();
  sh.getRange(f, 2).setValue(pin);
  CacheService.getScriptCache().remove('USUARIOS');
  registrarHistorial_('', nombre, 'Usuario', 'PIN reiniciado por ' + QUIEN_, QUIEN_);
  return { nombre: nombre, pin: pin };
}

function activarUsuario(k, nombre, activo) {
  soloAdmin_(k);
  if (!activo && String(nombre).toLowerCase() === QUIEN_.toLowerCase()) throw new Error('No podés bloquearte a vos mismo');
  protegePropietario_(nombre);
  const sh = hojaUsuarios_();
  const f = filaUsuario_(sh, nombre);
  if (!f) throw new Error('Usuario no encontrado');
  sh.getRange(f, 3).setValue(activo ? 'SI' : 'NO');
  CacheService.getScriptCache().remove('USUARIOS');
  registrarHistorial_('', nombre, 'Usuario', (activo ? 'Habilitado' : 'Bloqueado') + ' por ' + QUIEN_, QUIEN_);
  return true;
}

function cambiarPin(k, nuevo) {
  admin_(k);
  nuevo = String(nuevo || '').trim();
  if (!/^\d{4,6}$/.test(nuevo)) throw new Error('El PIN tiene que ser de 4 a 6 números');
  if (!QUIEN_) throw new Error('No hay usuario');
  const sh = hoja_('Usuarios');
  const d = sh.getDataRange().getValues();
  const i = d.findIndex((r, j) => j > 0 && String(r[0]).trim().toLowerCase() === QUIEN_.toLowerCase());
  if (i === -1) throw new Error('Usuario no encontrado');
  sh.getRange(i + 1, 2).setValue(nuevo);
  CacheService.getScriptCache().remove('USUARIOS');
  registrarHistorial_('', QUIEN_, 'Usuario', 'Cambió su PIN', QUIEN_);
  return true;
}

function fmt_(d, f) { return Utilities.formatDate(d, 'GMT-3', f); }

function filaPorId_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const ids = sh.getRange(2, COL.id, n, 1).getValues();
  const i = ids.findIndex(r => String(r[0]) === String(id));
  return i === -1 ? 0 : i + 2;
}

function plantillasIniciales_() {
  return [
    ['contacto-screening', 'WhatsApp', 'Hola [NOMBRE], ¿cómo estás? Soy [ENTREVISTADOR] de Recursos humanos del [POLI].\nTe contacto por la búsqueda de [PUESTO].\nMe gustaría coordinar una breve charla telefónica con vos para conocerte.\nSi te interesa, elegí el día y horario que mejor te quede desde este link y queda agendado:\n[LINK]\nSaludos!'],
    ['contacto-competencias', 'WhatsApp', 'Hola [NOMBRE], ¿cómo estás? Soy [ENTREVISTADOR] del equipo de Atracción y Elección.\nTe escribo por la búsqueda de [PUESTO] en [EMPRESA].\nNos gustaría avanzar con tu perfil a la siguiente etapa del proceso. Se trata de una entrevista de aproximadamente una hora, en la que vamos a profundizar sobre tu experiencia y tu recorrido profesional.\nSi te interesa continuar, elegí el día y horario que mejor te quede desde este link y queda agendada:\n[LINK]\n¡Nos vemos pronto!\nSaludos,\n[FIRMA]'],
    ['contacto-tecnica', 'WhatsApp', 'Hola [NOMBRE], ¿cómo estás?\nQueríamos contarte que nos gustaría avanzar con tu perfil a la siguiente etapa del proceso para la posición de [PUESTO] en [POLICLINICO].\nLa próxima instancia consiste en una entrevista técnica con la División de Enfermería, de aproximadamente una hora. Será una conversación para profundizar sobre tu experiencia y conocer los desafíos de la posición.\nSi te interesa continuar, elegí el día y horario que mejor te quede desde este link y queda agendada:\n[LINK]\nSaludos,\n[FIRMA]'],
    ['confirmacion-cita', 'WhatsApp', 'Hola [NOMBRE], te confirmamos tu entrevista de [TIPO] para la búsqueda de [PUESTO]:\n📅 [FECHA] a las [HORA]\n📍 [LUGAR]\nSi no podés asistir, avisanos con anticipación para reprogramarla. ¡Gracias!\n[FIRMA]'],
    ['descarte-sin', 'Mail', 'Hola [NOMBRE],\nQueremos agradecerte sinceramente por tu interés en la posición de [PUESTO] y por el tiempo dedicado a postularte.\nLuego de revisar cuidadosamente los currículums recibidos, en esta oportunidad decidimos avanzar con candidatos cuyos perfiles se encuentran más alineados con los requerimientos específicos de la posición.\nValoramos mucho tu trayectoria profesional. Conservaremos tu perfil para futuras oportunidades que puedan resultar acordes a tu experiencia.\nTe deseamos muchos éxitos en tus próximos desafíos profesionales.\nSaludos cordiales,\n[FIRMA]'],
    ['descarte-post', 'Mail', 'Hola [NOMBRE],\nMuchas gracias por el tiempo compartido durante el proceso de selección para la posición de [PUESTO].\nQueremos agradecer especialmente tu apertura y predisposición durante la conversación. Luego de analizar toda la información recabada, hemos decidido avanzar con otros perfiles que presentan un ajuste más cercano a las necesidades actuales de la posición.\nValoramos mucho tu experiencia y el intercambio mantenido, y esperamos poder considerar tu perfil para futuras oportunidades.\nTe deseamos muchos éxitos en tus próximos proyectos.\nSaludos cordiales,\n[FIRMA]'],
    ['descarte-competencias', 'WhatsApp', 'Hola [NOMBRE],\nAntes que nada, queremos agradecerte el tiempo que compartiste con nosotros durante las entrevistas para la posición de [PUESTO].\nValoramos mucho tu predisposición y el interés demostrado a lo largo del proceso.\nLuego de evaluar integralmente los perfiles, en esta oportunidad hemos decidido continuar con candidatos que presentan un ajuste más cercano a los desafíos específicos de la posición.\nEsperamos poder considerar tu perfil en futuras oportunidades. Te deseamos muchos éxitos.\nSaludos cordiales,\n[FIRMA]'],
    ['descarte-pami', 'WhatsApp', 'Hola [NOMBRE],\nAntes que nada, muchas gracias por el tiempo, el compromiso y la buena predisposición que demostraste durante todo el proceso de selección para la posición de [PUESTO].\nLuego de la entrevista con [EMPRESA], decidimos avanzar con otro perfil, principalmente porque [MOTIVO TÉCNICO].\nNos resultó muy interesante conocerte y tu perfil nos dejó una muy buena impresión. Por ese motivo, con tu autorización, mantendremos tus antecedentes para tenerte en cuenta en futuras búsquedas.\nTe agradecemos nuevamente y te deseamos muchos éxitos.\nSaludos,\n[FIRMA]']
  ];
}


function colorNota_(n) {
  return { A: ['#1E8E5A', '#FFFFFF'], 'B+': ['#8BD3A8', '#0F3A24'], B: ['#F2B705', '#4A3A00'], C: ['#D9534F', '#FFFFFF'] }[n] || null;
}

function estadoPorNota_(etapa, nota) {
  const C = { cv: 'Descartado sin screening', screening: 'Descartado post screening', competencias: 'Descartado en competencias', pami2: 'Descartado en entrevista técnica' };
  const B = { cv: 'Avanza a screening', screening: 'Avanza a competencias', competencias: 'Reserva competencias', pami2: 'Reserva entrevista técnica' };
  const A = { cv: 'Avanza a screening', screening: 'Avanza a competencias', competencias: 'Avanza a entrevista técnica', pami2: 'Contratado' };
  return (nota === 'C' ? C : nota === 'B' ? B : A)[etapa];
}

function etapaDeEstado_(estado) {
  const e = String(estado || '').trim();
  for (let i = 0; i < ESTADOS_COL.length; i++) if (ESTADOS_COL[i][1].indexOf(e) > -1) return i === ESTADOS_COL.length - 1 ? -1 : i;
  return -1;
}

function reactivar_(sh, fila, antes, estado) {
  if (!/^Descartad/i.test(String(antes || '')) || /^Descartad/i.test(String(estado || ''))) return;
  sh.getRange(fila, COL.descPor).clearContent();
  const mot = String(sh.getRange(fila, COL.motivo).getValue() || '');
  if (!mot && !String(sh.getRange(fila, COL.descarte).getValue() || '').trim()) return;
  sh.getRange(fila, COL.motivo).clearContent();
  sh.getRange(fila, COL.descarte).clearContent();
  registrarHistorial_(sh.getRange(fila, COL.id).getValue(), sh.getRange(fila, COL.nombre).getValue(), 'Reactivado', 'Sale de Descartados' + (mot ? ' (se quitó el motivo: ' + mot + ')' : ''), QUIEN_ ? 'tablero' : 'sistema');
}

function setEstado_(sh, fila, estado, nota, quien) {
  const c = sh.getRange(fila, COL.estado);
  const antes = String(c.getValue() || '');
  const regla = c.getDataValidation();
  c.setDataValidation(null).setValue(estado);
  if (regla) c.setDataValidation(regla);
  const col = nota ? colorNota_(nota) : null;
  if (col) c.setBackground(col[0]).setFontColor(col[1]).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(fila, COL.mov).setValue(new Date());
  reactivar_(sh, fila, antes, estado);
  if (/^Descartad/i.test(estado) && !/^Descartad/i.test(antes)) {
    const q = QUIEN_ && (!quien || quien === 'tablero' || quien === 'formulario') ? QUIEN_ : (quien || '');
    sh.getRange(fila, COL.descPor).setValue(q);
  }
}

function registrarHistorial_(id, nombre, tipo, detalle, quien) {
  const q = QUIEN_ && (!quien || quien === 'tablero' || quien === 'formulario') ? QUIEN_ : (quien || '');
  try { hoja_('Historial').appendRow([new Date(), id, nombre, tipo, detalle, q]); } catch (e) { }
}

function aplicarNota_(sh, fila, etapa, nota, quien) {
  let v = String(nota || '').trim().toUpperCase();
  const m = { ALTO: 'A', MEDIO: 'B', BAJO: 'C' };
  if (m[v]) v = m[v];
  if (['A', 'B+', 'B', 'C'].indexOf(v) === -1 || !NOTA_COL[etapa]) return null;
  const r = sh.getRange(fila, 1, 1, HEAD_CAND.length).getValues()[0];
  const celda = sh.getRange(fila, NOTA_COL[etapa]);
  celda.clearDataValidations().setValue(v);
  const col = colorNota_(v);
  celda.setBackground(col[0]).setFontColor(col[1]).setFontWeight('bold').setHorizontalAlignment('center');
  const actual = String(r[COL.estado - 1]).trim();
  const etq = { cv: 'CV', screening: 'Screening', competencias: 'Competencias', pami2: 'Entrevista técnica' }[etapa];
  if (actual === 'Contratado') {
    registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Nota ' + etq, v + ' (ya estaba Contratado, no cambia el estado)', quien);
    return actual;
  }
  const estado = estadoPorNota_(etapa, v);
  setEstado_(sh, fila, estado, v, quien);
  registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Nota ' + etq, v + ' → ' + estado, quien);
  return estado;
}

function alEditar(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== 'Candidatos') return;
    const fila = e.range.getRow(), col = e.range.getColumn();
    if (fila < 2 || e.range.getNumRows() > 1) return;
    let quien = 'planilla';
    try { const em = e.user && e.user.getEmail && e.user.getEmail(); if (em) quien = 'planilla · ' + em; } catch (x) { }
    const etapa = Object.keys(NOTA_COL).find(k => NOTA_COL[k] === col);
    if (etapa) { aplicarNota_(sh, fila, etapa, e.range.getValue(), quien); return; }
    if (col === COL.domicilio || col === COL.loc) sh.getRange(fila, COL.polSug).clearContent();
    if (col === COL.domicilio) sh.getRange(fila, COL.domRev).setValue('Manual');
    if (col === COL.estado) {
      sh.getRange(fila, COL.mov).setValue(new Date());
      reactivar_(sh, fila, e.oldValue, e.range.getValue());
      registrarHistorial_(sh.getRange(fila, COL.id).getValue(), sh.getRange(fila, COL.nombre).getValue(), 'Estado', 'Cambiado a mano a "' + e.range.getValue() + '"', quien);
    }
  } catch (err) { registrarError_('alEditar', err); }
}

function citaObj_(txt) {
  const m = String(txt || '').match(/^(\d{4}-\d{2}-\d{2} \d{2}:\d{2})\|(.*)$/);
  if (!m) return null;
  return { ts: m[1], tipo: m[2] };
}

function getTablero(k) {
  admin_(k);
  try { asegurarColumnas_(); } catch (e) { registrarError_('asegurar columnas', e); }
  const aj = ajustes_();
  const d = hoja_('Candidatos').getDataRange().getValues();
  const cands = [];
  const entPor = {};
  try {
    const ev = hoja_('Evaluaciones');
    if (ev && ev.getLastRow() > 1) ev.getRange(2, 1, ev.getLastRow() - 1, 5).getValues().forEach(x => {
      if (x[1] && x[3] && String(x[4] || '').trim()) (entPor[String(x[1])] = entPor[String(x[1])] || {})[String(x[3])] = String(x[4]).trim();
    });
  } catch (e) { registrarError_('entrevistadores por etapa', e); }
  const descPor = {};
  try {
    const hs = hoja_('Historial'), n = hs.getLastRow() - 1;
    if (n > 0) {
      const ini = Math.max(2, hs.getLastRow() - 4999);
      hs.getRange(ini, 1, hs.getLastRow() - ini + 1, 6).getValues().forEach(x => {
        if (x[1] && x[3] !== 'Reactivado' && /(→\s*|a ")Descartad/.test(String(x[4]))) descPor[String(x[1])] = (String(x[5] || '').trim() || 'sistema') + (x[0] instanceof Date ? ' · ' + fmt_(x[0], 'dd/MM') : '');
      });
    }
  } catch (e) { registrarError_('quién descartó', e); }
  for (let i = 1; i < d.length; i++) {
    const r = d[i];
    if (!r[COL.nombre - 1]) continue;
    let resp = [];
    try { resp = JSON.parse(r[COL.resp - 1] || '[]'); } catch (e) { }
    const dt = v => v instanceof Date ? fmt_(v, 'dd/MM/yyyy') : String(v || '');
    cands.push({
      fila: i + 1, id: String(r[COL.id - 1]), fecha: dt(r[0]), mov: r[COL.mov - 1] instanceof Date ? fmt_(r[COL.mov - 1], "yyyy-MM-dd'T'HH:mm") : '',
      nombre: r[COL.nombre - 1], tel: String(r[COL.tel - 1] || ''), dni: String(r[COL.dni - 1] || ''), email: r[COL.email - 1], loc: r[COL.loc - 1],
      busq: r[COL.busq - 1], pret: String(r[COL.pret - 1] || ''), ajuste: r[COL.ajuste - 1], just: r[COL.just - 1], resumen: r[COL.resumen - 1], req: r[COL.req - 1],
      cv: r[COL.cv - 1], estado: String(r[COL.estado - 1] || 'Nuevo').trim(), entrev: r[COL.entrev - 1], ent: entPor[String(r[COL.id - 1])] || {}, descPor: String(r[COL.descPor - 1] || '').trim() || descPor[String(r[COL.id - 1])] || '',
      notaScr: r[COL.notaScr - 1], notaComp: r[COL.notaComp - 1], notaP1: r[COL.notaP1 - 1], notaP2: r[COL.notaP2 - 1],
      recTipo: r[COL.recTipo - 1], recPor: r[COL.recPor - 1], coment: r[COL.coment - 1], resp: resp, proc: r[COL.proc - 1], fuente: r[COL.fuente - 1],
      cita: citaObj_(r[COL.cita - 1]), entrevistas: String(r[COL.entrevistas - 1] || ''), contactado: String(r[COL.contactado - 1] || ''),
      descarte: String(r[COL.descarte - 1] || ''), infScr: String(r[COL.infScr - 1] || ''), infComp: String(r[COL.infComp - 1] || ''),
      zona: String(r[COL.zona - 1] || ''), formacion: String(r[COL.formacion - 1] || ''), postulacion: String(r[COL.postulacion - 1] || ''), policlinico: String(r[COL.policlinico - 1] || ''), domicilio: String(r[COL.domicilio - 1] || ''), polSug: String(r[COL.polSug - 1] || ''), domRev: String(r[COL.domRev - 1] || ''), motivo: String(r[COL.motivo - 1] || ''), turno: String(r[COL.turno - 1] || ''),
      conoce: (resp.find(x => /conoc[eé]s a alguien/i.test(x.q) && x.a && !/^(no|ninguno|nadie|-)\.?$/i.test(String(x.a).trim())) || {}).a || ''
    });
  }
  return JSON.stringify({
    policlinicos: POLICLINICOS, motivos: MOTIVOS_DESCARTE, turnos: TURNOS, estadosCol: ESTADOS_COL, desc: ESTADOS_DESC, res: ESTADOS_RES, todos: todosEstados_(),
    entrevistadores: lista_(aj.entrevistadores, ','),
    busquedas: busquedas_().map(b => ({ id: b.id, titulo: b.titulo, recomendacion: b.recomendacion, link_perfil: b.link_perfil, link_flyer: b.link_flyer, ubicacion: b.ubicacion })),
    cfg: { sesionHoras: Number(aj.sesion_inactividad_horas) || 8, recordarDias: Number(aj.sesion_recordar_dias) || 30, cliente: aj.nombre_cliente || 'PAMI', diasAlerta: Number(aj.dias_alerta) || 7, linkAgenda: aj.link_agenda || '', duraciones: { screening: Number(aj.duracion_screening) || 30, competencias: Number(aj.duracion_competencias) || 60, pami1: Number(aj.duracion_pami1) || 30, pami2: Number(aj.duracion_pami2) || 60 } },
    pinActivo: exigePin_(), admin: ROL_ADMIN_, rol: ROL_, ahora: fmt_(new Date(), "yyyy-MM-dd'T'HH:mm"), cands: cands
  });
}

function getHistorial(k, id) {
  admin_(k);
  try {
    const cache = CacheService.getScriptCache(), ck = 'VISTA_' + QUIEN_ + '_' + id;
    if (!cache.get(ck)) {
      cache.put(ck, '1', 600);
      const sh = hoja_('Candidatos'), fl = filaPorId_(sh, id);
      accesoLog_(QUIEN_, 'Ficha vista', fl ? String(sh.getRange(fl, COL.nombre).getValue()) : String(id));
    }
  } catch (e) { }
  const d = hoja_('Historial').getDataRange().getValues().slice(1).filter(r => String(r[1]) === String(id));
  return d.reverse().slice(0, 80).map(r => ({ fecha: r[0] instanceof Date ? fmt_(r[0], 'dd/MM/yyyy HH:mm') : String(r[0]), tipo: r[3], detalle: r[4], quien: r[5] }));
}

function moverEstado(k, fila, estado) {
  escritor_(k);
  if (todosEstados_().indexOf(estado) === -1) throw new Error('Estado inválido');
  const sh = hoja_('Candidatos');
  fila = Number(fila);
  if (!(fila >= 2 && fila <= sh.getLastRow())) throw new Error('Fila inválida');
  const antes = String(sh.getRange(fila, COL.estado).getValue());
  if (antes === estado) return true;
  setEstado_(sh, fila, estado, null);
  registrarHistorial_(sh.getRange(fila, COL.id).getValue(), sh.getRange(fila, COL.nombre).getValue(), 'Estado', antes + ' → ' + estado, 'tablero');
  return true;
}

function ponerNota(k, fila, etapa, nota) {
  escritor_(k);
  const sh = hoja_('Candidatos');
  fila = Number(fila);
  if (!(fila >= 2 && fila <= sh.getLastRow())) throw new Error('Fila inválida');
  if (!NOTA_COL[etapa]) throw new Error('Etapa inválida');
  return aplicarNota_(sh, fila, etapa, nota, 'tablero');
}

function guardarCandidato(k, fila, c) {
  escritor_(k);
  const sh = hoja_('Candidatos');
  fila = Number(fila);
  if (!(fila >= 2 && fila <= sh.getLastRow())) throw new Error('Fila inválida');
  const corto = (s, n) => String(s == null ? '' : s).slice(0, n);
  const set = (col, v) => sh.getRange(fila, col).clearDataValidations().setValue(v);
  const id = sh.getRange(fila, COL.id).getValue(), nombre = sh.getRange(fila, COL.nombre).getValue();
  const cambio = (col, nuevo, etiqueta) => {
    const antes = String(sh.getRange(fila, col).getValue() || '');
    if (antes !== String(nuevo)) registrarHistorial_(id, nombre, etiqueta, (antes || '—') + ' → ' + (nuevo || '—'), 'tablero');
  };
  if (c.busq != null) {
    if (c.busq !== '' && !busquedas_().some(b => b.id === c.busq)) throw new Error('Búsqueda inválida');
    cambio(COL.busq, c.busq, 'Búsqueda');
    set(COL.busq, c.busq);
  }
  if (c.accion === 'descartar' || c.accion === 'reactivar') {
    const est = String(sh.getRange(fila, COL.estado).getValue()).trim();
    const desc = /^Descartad/i.test(est);
    if (c.accion === 'descartar' && !desc) {
      if (est === 'Contratado') throw new Error('Está contratado/a: no se puede descartar');
      if (!String(c.motivo || '').trim()) throw new Error('Elegí el motivo de descarte');
      const i = ESTADOS_COL.findIndex(e => e[1].indexOf(est) > -1);
      const etapa = /competencias/i.test(est) && /^Reserva/i.test(est) ? 'competencias' : /t[eé]cnica/i.test(est) && /^Reserva/i.test(est) ? 'pami2' : ETAPAS[i < 0 ? 0 : Math.min(i, 3)];
      moverEstado(k, fila, estadoPorNota_(etapa, 'C'));
    } else if (c.accion === 'reactivar' && desc) {
      const i = ['sin screening', 'post screening', 'en competencias', 'en entrevista técnica'].findIndex(x => est === 'Descartado ' + x);
      moverEstado(k, fila, ESTADOS_COL[i < 0 ? 0 : i][2]);
    }
  }
  if (c.policlinico != null) {
    if (c.policlinico !== '' && POLICLINICOS.indexOf(c.policlinico) === -1) throw new Error('Policlínico inválido');
    cambio(COL.policlinico, c.policlinico, 'Policlínico');
    set(COL.policlinico, c.policlinico);
  }
  if (c.motivo != null && /^Descartad/i.test(String(sh.getRange(fila, COL.estado).getValue()))) {
    c.motivo = motivoNorm_(c.motivo);
    const antes = String(sh.getRange(fila, COL.motivo).getValue() || '');
    set(COL.motivo, c.motivo);
    if (antes !== c.motivo) registrarHistorial_(id, nombre, 'Motivo de descarte', c.motivo || '(sin motivo)', 'tablero');
  }
  if (c.turno != null) { const t = turnoNorm_(c.turno); cambio(COL.turno, t, 'Turno preferido'); set(COL.turno, t); }
  if (c.entrev != null) { cambio(COL.entrev, corto(c.entrev, 60), 'Entrevistó'); set(COL.entrev, corto(c.entrev, 60)); }
  if (c.recTipo != null) set(COL.recTipo, corto(c.recTipo, 40));
  if (c.recPor != null) {
    set(COL.recPor, corto(c.recPor, 120));
    if (c.recTipo) registrarHistorial_(id, nombre, 'Recomendación', c.recTipo + (c.recPor ? ': ' + c.recPor : ''), 'tablero');
  }
  if (c.coment != null) {
    if (String(sh.getRange(fila, COL.coment).getValue() || '') !== corto(c.coment, 2000)) registrarHistorial_(id, nombre, 'Observaciones', 'Se editaron las observaciones', 'tablero');
    set(COL.coment, corto(c.coment, 2000));
  }
  return true;
}

function reanalizar(k, fila) {
  escritor_(k);
  const sh = hoja_('Candidatos');
  sh.getRange(Number(fila), COL.intentos).setValue(0);
  analizarFila_(Number(fila));
  return sh.getRange(Number(fila), COL.proc).getValue();
}

function rellenar_(txt, v) {
  return String(txt).replace(/\[([A-ZÁÉÍÓÚÑ ]+)\]/g, (m, k) => (v[k] != null ? v[k] : m));
}

function textoPlantilla_(clave) {
  const f = hoja_('Plantillas').getDataRange().getValues().find(r => r[0] === clave);
  if (f) return String(f[2]);
  const d = plantillasIniciales_().find(x => x[0] === clave);
  return d ? d[2] : '';
}

function poliRomano_(t) {
  return String(t || '').replace(/PAMI\s*1\b/, 'PAMI I').replace(/PAMI\s*2\b/, 'PAMI II');
}

function datosComunes_(r) {
  const aj = ajustes_();
  const b = busquedas_().find(x => x.id === r[COL.busq - 1]);
  const pol = poliRomano_(r[COL.policlinico - 1] || String(r[COL.polSug - 1] || '').split(' · ')[0]);
  const hayPol = pol.indexOf('Policlínico') === 0;
  const quien = String(QUIEN_ || '').trim().split(/\s+/)[0];
  return {
    NOMBRE: String(r[COL.nombre - 1]).trim().split(/\s+/)[0], PUESTO: b ? b.titulo : '', EMPRESA: aj.nombre_cliente || 'PAMI',
    ENTREVISTADOR: quien || String(r[COL.entrev - 1] || lista_(aj.entrevistadores, ',')[0] || ''), LINK: aj.link_agenda || '', FIRMA: aj.firma || '',
    POLICLINICO: hayPol ? 'el ' + pol : 'los Policlínicos PAMI I y II', POLI: hayPol ? pol : 'Policlínico PAMI', aj: aj
  };
}

function telWa_(t) {
  let d = String(t || '').replace(/\D/g, '');
  if (d.indexOf('54') === 0) d = d.slice(2);
  d = d.replace(/^0+/, '');
  if (d.charAt(0) === '9' && d.length >= 11) d = d.slice(1);
  return '549' + d;
}

function armarComunicacion(k, fila, tipo) {
  escritor_(k);
  const sh = hoja_('Candidatos');
  fila = Number(fila);
  const r = sh.getRange(fila, 1, 1, HEAD_CAND.length).getValues()[0];
  const v = datosComunes_(r);
  const claveLink = { 'contacto-screening': 'link_agenda_screening', 'contacto-competencias': 'link_agenda_competencias', 'contacto-tecnica': 'link_agenda_pami2' }[tipo];
  if (claveLink && v.aj[claveLink]) v.LINK = v.aj[claveLink];
  const polEf = String(r[COL.policlinico - 1] || String(r[COL.polSug - 1] || '').split(' · ')[0]);
  if (claveLink && /2$/.test(polEf) && v.aj[claveLink + '_poli2']) v.LINK = v.aj[claveLink + '_poli2'];
  let texto = textoPlantilla_(tipo);
  if (!texto) return { ok: false, error: 'No encontré la plantilla ' + tipo + ' en la pestaña Plantillas' };
  if (!v.LINK) texto = texto.replace(/[^\n]*(link|enlace)[^\n]*\n\[LINK\]/i, 'Si te interesa, contanos qué día y horario te quedan mejor para agendarla.').replace(/ ?\[LINK\]/g, '');
  texto = rellenar_(texto, v);
  if (!String(r[COL.tel - 1]).replace(/\D/g, '')) return { ok: false, error: 'El candidato no tiene teléfono' };
  const etq = { 'contacto-screening': 'screening', 'contacto-competencias': 'competencias', 'contacto-tecnica': 'entrevista técnica', 'descarte-competencias': 'descarte comp.', 'descarte-pami': 'descarte PAMI', 'confirmacion-cita': 'confirmación' }[tipo] || '';
  if (etq) {
    sh.getRange(fila, COL.contactado).setValue(etq + ' · ' + fmt_(new Date(), 'dd/MM/yyyy'));
    sh.getRange(fila, COL.mov).setValue(new Date());
    registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Contacto', 'WhatsApp: ' + etq, 'tablero');
  }
  return { ok: true, url: 'https://wa.me/' + telWa_(r[COL.tel - 1]) + '?text=' + encodeURIComponent(texto), texto: texto };
}

function borrarMarcaContacto(k, fila) {
  escritor_(k);
  hoja_('Candidatos').getRange(Number(fila), COL.contactado).clearContent();
  return true;
}

function mailValido_(m) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(m || '').trim()); }

function mailDescarte_(sh, fila, textoSin, textoPost) {
  const r = sh.getRange(fila, 1, 1, HEAD_CAND.length).getValues()[0];
  const mail = String(r[COL.email - 1]).trim();
  if (!mailValido_(mail)) return { ok: false, error: 'Sin mail válido' };
  const v = datosComunes_(r);
  const sin = /sin screening/i.test(String(r[COL.estado - 1]));
  const base = sin ? (textoSin || textoPlantilla_('descarte-sin')) : (textoPost || textoPlantilla_('descarte-post'));
  if (!base) return { ok: false, error: 'No hay plantilla de descarte' };
  GmailApp.sendEmail(mail, v.aj.asunto_mail || 'Proceso de selección', rellenar_(base, v));
  sh.getRange(fila, COL.descarte).setValue('Enviado ' + fmt_(new Date(), 'dd/MM/yyyy'));
  registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Mail', 'Descarte enviado por mail', 'tablero');
  return { ok: true };
}

function enviarMailDescarte(k, fila) {
  escritor_(k);
  const sh = hoja_('Candidatos');
  try { return mailDescarte_(sh, Number(fila)); } catch (e) { return { ok: false, error: String(e) }; }
}

function previaDescartes(k, busq) {
  admin_(k);
  const d = hoja_('Candidatos').getDataRange().getValues();
  let sin = 0, post = 0, sinMail = 0;
  for (let i = 1; i < d.length; i++) {
    if (!d[i][COL.nombre - 1] || (busq && d[i][COL.busq - 1] !== busq)) continue;
    const e = String(d[i][COL.estado - 1]).trim();
    if (!/^Descartado (sin screening|post screening)$/i.test(e) || String(d[i][COL.descarte - 1]).trim()) continue;
    if (!mailValido_(d[i][COL.email - 1])) { sinMail++; continue; }
    if (/sin screening/i.test(e)) sin++; else post++;
  }
  return { sin: sin, post: post, sinMail: sinMail, textoSin: textoPlantilla_('descarte-sin'), textoPost: textoPlantilla_('descarte-post') };
}

function enviarDescartesMasivo(k, busq, textoSin, textoPost) {
  escritor_(k);
  const sh = hoja_('Candidatos');
  const d = sh.getDataRange().getValues();
  let enviados = 0, sinMail = 0, pendientes = 0;
  for (let i = 1; i < d.length; i++) {
    if (!d[i][COL.nombre - 1] || (busq && d[i][COL.busq - 1] !== busq)) continue;
    if (!/^Descartado (sin screening|post screening)$/i.test(String(d[i][COL.estado - 1]).trim()) || String(d[i][COL.descarte - 1]).trim()) continue;
    if (!mailValido_(d[i][COL.email - 1])) { sinMail++; continue; }
    if (enviados >= 40) { pendientes++; continue; }
    try { const res = mailDescarte_(sh, i + 1, textoSin, textoPost); if (res.ok) enviados++; } catch (e) { registrarError_('mail descarte fila ' + (i + 1), e); }
  }
  return { enviados: enviados, sinMail: sinMail, pendientes: pendientes };
}

function getAgenda(k) {
  admin_(k);
  const d = hoja_('Candidatos').getDataRange().getValues();
  const desde = fmt_(new Date(Date.now() - 3600000), 'yyyy-MM-dd HH:mm');
  const citas = [];
  let evs = [];
  try { evs = eventosCal_(new Date(), new Date(Date.now() + 60 * 86400000)); } catch (e) { registrarError_('leer calendarios agenda', e); }
  for (let i = 1; i < d.length; i++) {
    const c = citaObj_(d[i][COL.cita - 1]);
    if (!c || c.ts < desde || !d[i][COL.nombre - 1]) continue;
    const cand = candDe_(d[i]);
    const ev = evs.find(e => fmt_(e.inicio, 'yyyy-MM-dd HH:mm') === c.ts && eventoEsDe_(e, cand));
    const origen = ev ? ev.poli : '', asignado = String(d[i][COL.policlinico - 1] || '').trim();
    citas.push({ fila: i + 1, nombre: d[i][COL.nombre - 1], ts: c.ts, tipo: c.tipo, busq: d[i][COL.busq - 1], tel: String(d[i][COL.tel - 1] || ''), origen: origen, desfase: !!(origen && asignado && origen !== asignado) });
  }
  citas.sort((a, b) => a.ts < b.ts ? -1 : 1);
  let sin = [];
  try { sin = citasSinCandidato_(evs); } catch (e) { registrarError_('citas sin candidato', e); }
  let avisos = [];
  try { avisos = estadoCalendarios_(); } catch (e) { registrarError_('estado calendarios', e); }
  return { citas: citas, sinAsignar: sin, avisos: avisos };
}

function postulacionPdf(k, fila) {
  admin_(k);
  const sh = hoja_('Candidatos');
  fila = Number(fila);
  if (!(fila >= 2 && fila <= sh.getLastRow())) throw new Error('Fila inválida');
  const existente = String(sh.getRange(fila, COL.postulacion).getValue() || '');
  if (existente) return existente;
  const r = sh.getRange(fila, 1, 1, HEAD_CAND.length).getValues()[0];
  let resp = [];
  try { resp = JSON.parse(r[COL.resp - 1] || '[]'); } catch (e) { }
  const b = busquedas_().find(x => x.id === r[COL.busq - 1]) || { titulo: '' };
  const fila2 = (a, v) => v ? '<tr><td style="padding:5px 8px;color:#5D6B82;width:38%;font-size:12px">' + h_(a) + '</td><td style="padding:5px 8px;font-size:13px;font-weight:bold">' + h_(v) + '</td></tr>' : '';
  const cuerpo = seccion_('Datos de contacto') + '<table style="width:100%;border-collapse:collapse">' + fila2('Nombre', r[COL.nombre - 1]) + fila2('Email', r[COL.email - 1]) + fila2('Teléfono', r[COL.tel - 1]) + fila2('DNI', r[COL.dni - 1]) + fila2('Localidad', r[COL.loc - 1]) + fila2('Domicilio', r[COL.domicilio - 1]) + fila2('Remuneración pretendida', r[COL.pret - 1]) + fila2('Llegó por', r[COL.fuente - 1]) + '</table>' +
    (resp.length ? seccion_('Respuestas del formulario') + '<table style="width:100%;border-collapse:collapse">' + resp.map(x => fila2(x.q, x.a)).join('') + '</table>' : '');
  const html = htmlInforme_('Postulación', { nombre: r[COL.nombre - 1], puesto: b.titulo }, [['Recibida', r[0] instanceof Date ? fmt_(r[0], 'dd/MM/yyyy HH:mm') : String(r[0])]], cuerpo, r[COL.ajuste - 1] || '—', null, 'Nivel de ajuste sugerido por la IA al CV');
  const pdf = DriveApp.getFolderById(ajustes_().carpeta_informes).createFile(Utilities.newBlob(html, 'text/html', 'x.html').getAs('application/pdf').setName('Postulacion_' + String(r[COL.nombre - 1]).replace(/\s+/g, '_') + '.pdf'));
  sh.getRange(fila, COL.postulacion).setValue(pdf.getUrl());
  return pdf.getUrl();
}


function registrarAcceso(k, accion) {
  admin_(k);
  if (['Ingreso', 'Cierre de sesión'].indexOf(accion) === -1) return false;
  accesoLog_(QUIEN_, accion, '');
  return true;
}

function leerUltimas_(sh, cols, max) {
  const n = sh.getLastRow();
  if (n < 2) return [];
  const ini = Math.max(2, n - max + 1);
  return sh.getRange(ini, 1, n - ini + 1, cols).getValues().reverse();
}

function getAuditoria(k, tipo) {
  soloAdmin_(k);
  const f = v => v instanceof Date ? fmt_(v, 'dd/MM/yyyy HH:mm') : String(v || '');
  if (tipo === 'accesos') {
    const sh = ss_().getSheetByName('Accesos');
    return { filas: sh ? leerUltimas_(sh, 4, 400).map(r => ({ fecha: f(r[0]), usuario: String(r[1] || ''), accion: String(r[2] || ''), detalle: String(r[3] || '') })) : [] };
  }
  if (tipo === 'copias') {
    return { ultima: PropertiesService.getScriptProperties().getProperty('ULTIMA_COPIA') || '' };
  }
  return { filas: leerUltimas_(hoja_('Historial'), 6, 400).map(r => ({ fecha: f(r[0]), candidato: String(r[2] || ''), tipo: String(r[3] || ''), detalle: String(r[4] || ''), usuario: String(r[5] || '') })) };
}

function hacerCopia_() {
  const cfg = ajustes_();
  const padre = DriveApp.getFolderById(cfg.carpeta_cvs);
  const it = padre.getFoldersByName('Copias de seguridad');
  const carpeta = it.hasNext() ? it.next() : padre.createFolder('Copias de seguridad');
  const nombre = 'Copia Empleos PAMI ' + fmt_(new Date(), 'yyyy-MM-dd HH-mm');
  DriveApp.getFileById(prop_('SHEET_ID')).makeCopy(nombre, carpeta);
  const limite = Date.now() - 30 * 86400000;
  const fs = carpeta.getFiles();
  while (fs.hasNext()) { const x = fs.next(); if (x.getDateCreated().getTime() < limite) x.setTrashed(true); }
  PropertiesService.getScriptProperties().setProperty('ULTIMA_COPIA', fmt_(new Date(), 'dd/MM/yyyy HH:mm') + ' · ' + nombre);
  return nombre;
}

function copiaDiaria() {
  try { hacerCopia_(); accesoLog_('sistema', 'Copia de seguridad', 'Automática'); } catch (e) { registrarError_('copiaDiaria', e); }
}

function copiaAhora(k) {
  soloAdmin_(k);
  const n = hacerCopia_();
  accesoLog_(QUIEN_, 'Copia de seguridad', 'Manual: ' + n);
  return n;
}
