/* Usuarios y sesiones del kanban (como PAMI): se ingresa con nombre y apellido + PIN. Dos roles: Administrador y Usuario.
   - Pestaña «Usuarios»: nombre, rol, activo, sal, hash (el PIN nunca se guarda en claro), pin_temporal (se borra cuando la persona elige el suyo), cambio, ultimo_ingreso.
   - Para reiniciar un PIN a mano: borrar la celda «hash» de esa persona; al ingresar le pide elegir uno nuevo.
   - Sesión: token firmado (HMAC) con el usuario y el vencimiento (30 días). Todas las pantallas llaman al servidor por `api` (ver más abajo), que verifica el token en cada pedido.
   - Las funciones que antes se llamaban directo desde la página ahora terminan en «_» (privadas) y solo se pueden llegar por `api`. */
const HOJA_USU_ = 'Usuarios';
const HOJA_ACC_ = 'Accesos';
const SES_DIAS_ = 30;
const USAR_PIN_ = false; // por ahora se ingresa solo con nombre y apellido; con true se vuelve a pedir PIN (los usuarios sin PIN eligen uno en su primer ingreso)
const ROLES_ = ['Administrador', 'Usuario'];
const ADMINS_INICIALES_ = ['Ariel Jaime', 'Flavia Jaime'];
var SES_ = null; // usuario de esta ejecución (lo pone api)

function usuNorm_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9 ]/g, '').replace(/\s+/g, ' ').trim(); }
function usuClave_(s) { return usuNorm_(s).replace(/ /g, '_'); }
function usuNombre_(s) { return String(s || '').replace(/[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ0-9 ]/g, '').replace(/\s+/g, ' ').trim().slice(0, 40); }
function pinAzar_() { return String(1000 + Math.floor(Math.random() * 9000)); }

function secretoSes_() {
  const p = PropertiesService.getScriptProperties();
  let s = p.getProperty('BLUMIA_SECRET');
  if (!s) { s = Utilities.getUuid() + Utilities.getUuid(); p.setProperty('BLUMIA_SECRET', s); }
  return s;
}
function hashPin_(sal, pin) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, sal + '|' + pin + '|' + secretoSes_())); }
function firmar_(txt) { return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(txt, secretoSes_())).replace(/=+$/, ''); }

function accesoLog_(usuario, evento, detalle) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
    let sh = ss.getSheetByName(HOJA_ACC_);
    if (!sh) {
      sh = ss.insertSheet(HOJA_ACC_);
      sh.getRange(1, 1, 1, 4).setValues([['fecha', 'usuario', 'evento', 'detalle']]).setFontWeight('bold').setBackground('#0F3F27').setFontColor('#fff');
      sh.setFrozenRows(1);
    }
    sh.appendRow([new Date(), String(usuario || ''), String(evento), String(detalle || '').slice(0, 300)]);
  } catch (e) { }
}

function usuariosHoja_() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let sh = ss.getSheetByName(HOJA_USU_);
  if (sh) return sh;
  sh = ss.insertSheet(HOJA_USU_);
  sh.getRange(1, 1, 1, 8).setValues([['nombre', 'rol', 'activo', 'sal', 'hash', 'pin_temporal', 'cambio', 'ultimo_ingreso']]).setFontWeight('bold').setBackground('#0F3F27').setFontColor('#fff');
  sh.setFrozenRows(1);
  ADMINS_INICIALES_.forEach(n => nuevaFilaUsuario_(sh, n, 'Administrador'));
  return sh;
}

function nuevaFilaUsuario_(sh, nombre, rol) {
  if (!USAR_PIN_) { sh.appendRow([nombre, rol, 'SI', '', '', '', new Date().getTime(), '']); usuInvalidar_(); return ''; }
  const sal = 's' + Utilities.getUuid().slice(0, 7), pin = pinAzar_();
  sh.appendRow([nombre, rol, 'SI', sal, hashPin_(sal, pin), pin, new Date().getTime(), '']);
  sh.getRange(sh.getLastRow(), 4, 1, 3).setNumberFormat('@');
  usuInvalidar_();
  return pin;
}

function usuInvalidar_() { try { CacheService.getScriptCache().remove('usuarios_v1'); } catch (e) { } }

function leerUsuarios_() {
  const cache = CacheService.getScriptCache();
  const c = cache.get('usuarios_v1');
  if (c) { try { return JSON.parse(c); } catch (e) { } }
  const lista = leerUsuariosHoja_();
  try { cache.put('usuarios_v1', JSON.stringify(lista), 60); } catch (e) { }
  return lista;
}

function leerUsuariosHoja_() {
  const sh = usuariosHoja_();
  const n = sh.getLastRow();
  if (n < 2) return [];
  return sh.getRange(2, 1, n - 1, Math.max(8, Math.min(sh.getLastColumn(), 12))).getValues().map((r, i) => ({
    fila: i + 2, nombre: String(r[0]).trim(), clave: usuClave_(r[0]), rol: String(r[1]).trim() === 'Administrador' ? 'Administrador' : 'Usuario',
    activo: String(r[2]).trim().toUpperCase() === 'NO' ? 'NO' : 'SI', sal: String(r[3] || ''), hash: String(r[4] || ''), temporal: String(r[5] || '').trim(), cambio: Number(r[6]) || 0,
    mail: String(r[8] || '').trim(), entrevista: String(r[9] || '').trim().toLowerCase(), linkScr: String(r[10] || '').trim(), linkComp: String(r[11] || '').trim()
  })).filter(u => u.nombre);
}

// Primer nombre de cada usuario activo de la pestaña «Usuarios»: es la lista de entrevistadores por defecto (no hay nombres escritos en el código)
function entrevistadoresDefault_() {
  let L = [];
  try { L = leerUsuarios_().filter(u => u.activo === 'SI').map(u => String(u.nombre).split(' ')[0]); } catch (e) { }
  L = L.filter((n, i) => n && L.indexOf(n) === i);
  return L.length ? L : ['Entrevistador'];
}

function emitirToken_(u) {
  const iat = Date.now(), cuerpo = u.clave + '.' + iat + '.' + (iat + SES_DIAS_ * 864e5);
  return cuerpo + '.' + firmar_(cuerpo);
}

function sesionValida_(token) {
  const p = String(token || '').split('.');
  if (p.length !== 4) throw new Error('SESION: Ingresá de nuevo.');
  if (firmar_(p[0] + '.' + p[1] + '.' + p[2]) !== p[3]) throw new Error('SESION: Ingresá de nuevo.');
  if (Date.now() > Number(p[2])) throw new Error('SESION: Tu sesión venció. Ingresá de nuevo.');
  const u = leerUsuarios_().find(x => x.clave === p[0]);
  if (!u || u.activo !== 'SI') throw new Error('SESION: Tu usuario no está habilitado. Pedíselo a un administrador.');
  if (Number(p[1]) < u.cambio) throw new Error('SESION: Cambió tu PIN. Ingresá de nuevo.');
  return { nombre: u.nombre, rol: u.rol, admin: u.rol === 'Administrador', fila: u.fila };
}

// ---- Públicas (no pasan por api): ingresar, validar la sesión guardada y elegir el PIN propio ----
function login(nombre, pin) {
  const clave = usuClave_(nombre);
  if (!clave) throw new Error('Escribí tu nombre y apellido.');
  const cache = CacheService.getScriptCache(), kf = 'lg_' + clave, f = Number(cache.get(kf) || 0);
  if (f >= 5) throw new Error('Demasiados intentos. Probá de nuevo en 15 minutos.');
  const u = leerUsuarios_().find(x => x.clave === clave);
  const mal = () => { cache.put(kf, String(f + 1), 900); accesoLog_(nombre, 'Ingreso fallido', ''); throw new Error(USAR_PIN_ ? 'Nombre o PIN incorrectos.' : 'No encontré ese usuario. Escribí tu nombre y apellido como te dieron de alta, o pedile a un administrador que te agregue.'); };
  if (!u || u.activo !== 'SI') mal();
  if (USAR_PIN_ && u.hash && hashPin_(u.sal, String(pin == null ? '' : pin).trim()) !== u.hash) mal();
  cache.remove(kf);
  usuariosHoja_().getRange(u.fila, 8).setValue(new Date());
  accesoLog_(u.nombre, 'Ingreso', '');
  return { token: emitirToken_(u), nombre: u.nombre, rol: u.rol, debeCambiar: USAR_PIN_ && (!u.hash || !!u.temporal) };
}

function validarSesion(token) {
  const u = sesionValida_(token);
  return { nombre: u.nombre, rol: u.rol };
}

function cambiarMiPin(token, nuevo) {
  if (!USAR_PIN_) throw new Error('Por ahora se ingresa sin PIN.');
  const s = sesionValida_(token);
  const pin = String(nuevo == null ? '' : nuevo).trim();
  if (!/^\d{4,6}$/.test(pin)) throw new Error('El PIN tiene que tener de 4 a 6 números.');
  const sh = usuariosHoja_(), sal = 's' + Utilities.getUuid().slice(0, 7), ahora = Date.now();
  sh.getRange(s.fila, 4, 1, 4).setNumberFormat('@').setValues([[sal, hashPin_(sal, pin), '', ahora]]);
  usuInvalidar_();
  accesoLog_(s.nombre, 'PIN cambiado', '');
  const u = leerUsuarios_().find(x => x.fila === s.fila);
  return { token: emitirToken_(u), nombre: u.nombre, rol: u.rol };
}

// ---- Páginas: todo lo que muestra datos se pide después de ingresar (la página no trae nada incrustado) ----
function doGet(e) {
  const t = e && e.parameter && e.parameter.tipo;
  if (t === 'postulacion' || t === 'perfilpuesto') return doGetPagina_(e); // pública / con acceso temporal propio
  try {
    const cache = CacheService.getScriptCache();
    if (!cache.get('usu_ok')) { usuariosHoja_(); cache.put('usu_ok', '1', 3600); } // la primera vez crea la pestaña «Usuarios» con los administradores y sus PIN temporales (después no se vuelve a leer en cada carga)
  } catch (err) { }
  const p = {};
  Object.keys((e && e.parameter) || {}).forEach(k => { p[k] = String(e.parameter[k]); });
  const tpl = HtmlService.createTemplateFromFile('Shell');
  tpl.params = JSON.stringify(p).replace(/</g, '\\u003c');
  tpl.usaPin = USAR_PIN_ ? 'true' : 'false';
  return tpl.evaluate().setTitle('Blumia').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL).addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function pagina_(params) {
  const html = doGetPagina_({ parameter: params || {} }).getContent();
  return html.split('google.script.run').join('blRun'); // la página llama al servidor por la sesión (blRun la arma Shell.html)
}

// ---- Usuarios (solo administradores) ----
function yo_() { return { nombre: SES_.nombre, rol: SES_.rol }; }

function listarUsuarios_() {
  return leerUsuarios_().map(u => ({ nombre: u.nombre, rol: u.rol, activo: u.activo, pendiente: USAR_PIN_ && (!u.hash || !!u.temporal), mail: u.mail, screening: usuHace_(u, 'screening'), competencias: usuHace_(u, 'competencias'), linkScr: u.linkScr, linkComp: u.linkComp }));
}

/* Quién entrevista (29/10): en «Usuarios», columna «mail» (su Google, para invitarlo al evento) y «entrevista» (qué etapas hace: vacío = screening y competencias,
   «screening», «competencias», «screening,competencias» o «ninguna»). */
function usuHace_(u, etapa) {
  const e = String(u.entrevista || '');
  if (!e) return true;
  return e.split(',').map(x => x.trim()).indexOf(etapa) >= 0;
}

function guardarEntrevistaUsuario_(nombre, mail, screening, competencias, linkScr, linkComp) {
  const u = usuarioPorNombre_(nombre);
  mail = String(mail || '').trim().toLowerCase();
  if (mail && !/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(mail)) throw new Error('Ese mail no es válido.');
  const sh = usuariosHoja_();
  const url = v => { v = String(v || '').trim(); if (v && !/^https:\/\/[^\s<>"']{8,400}$/.test(v)) throw new Error('El link de agenda tiene que empezar con https:// y no puede llevar espacios.'); return v; };
  linkScr = url(linkScr || linkComp); linkComp = '';
  if (sh.getMaxColumns() < 12) sh.insertColumnsAfter(sh.getMaxColumns(), 12 - sh.getMaxColumns());
  if (!sh.getRange(1, 9).getValue()) { sh.getRange(1, 9, 1, 4).setValues([['mail', 'entrevista', 'link_screening', 'link_competencias']]).setFontWeight('bold').setBackground('#0F3F27').setFontColor('#fff'); }
  else if (!sh.getRange(1, 11).getValue()) { sh.getRange(1, 11, 1, 2).setValues([['link_screening', 'link_competencias']]).setFontWeight('bold').setBackground('#0F3F27').setFontColor('#fff'); }
  const et = [screening ? 'screening' : '', competencias ? 'competencias' : ''].filter(Boolean).join(',') || 'ninguna';
  sh.getRange(u.fila, 9, 1, 4).setNumberFormat('@').setValues([[mail, et === 'screening,competencias' ? '' : et, linkScr, linkComp]]);
  usuInvalidar_();
  accesoLog_(SES_ ? SES_.nombre : '', 'Entrevistador', u.nombre + ': ' + (et === 'ninguna' ? 'no entrevista' : et) + (mail ? ' · ' + mail : '') + (linkScr || linkComp ? ' · links de agenda cargados' : ''));
  return { ok: true };
}

// Para el panel «Agendar» de la ficha: quién puede entrevistar cada etapa (sin mails) y quién es el usuario actual
function entrevistadoresAgenda_() {
  return { yo: SES_ ? SES_.nombre : '', lista: leerUsuarios_().filter(u => u.activo === 'SI').map(u => ({ nombre: u.nombre, screening: usuHace_(u, 'screening'), competencias: usuHace_(u, 'competencias'), linkScr: !!u.linkScr, linkComp: !!u.linkComp })) };
}

function usuarioPorNombre_(nombre) {
  const u = leerUsuarios_().find(x => x.clave === usuClave_(nombre));
  if (!u) throw new Error('No encontré a ese usuario.');
  return u;
}

function cantAdmins_() { return leerUsuarios_().filter(u => u.rol === 'Administrador' && u.activo === 'SI').length; }

function crearUsuario_(nombre, rol) {
  const n = usuNombre_(nombre);
  if (n.split(' ').length < 2) throw new Error('Escribí nombre y apellido.');
  if (leerUsuarios_().some(u => u.clave === usuClave_(n))) throw new Error('Ya existe un usuario con ese nombre.');
  const r = ROLES_.indexOf(rol) > -1 ? rol : 'Usuario';
  const pin = nuevaFilaUsuario_(usuariosHoja_(), n, r);
  accesoLog_(SES_.nombre, 'Usuario creado', n + ' (' + r + ')');
  return { nombre: n, pin: pin };
}

function reiniciarPinUsuario_(nombre) {
  if (!USAR_PIN_) throw new Error('Por ahora se ingresa sin PIN.');
  const u = usuarioPorNombre_(nombre), sal = 's' + Utilities.getUuid().slice(0, 7), pin = pinAzar_();
  usuariosHoja_().getRange(u.fila, 4, 1, 4).setNumberFormat('@').setValues([[sal, hashPin_(sal, pin), pin, Date.now()]]);
  usuInvalidar_();
  accesoLog_(SES_.nombre, 'PIN reiniciado', u.nombre);
  return { nombre: u.nombre, pin: pin };
}

function activarUsuario_(nombre, activo) {
  const u = usuarioPorNombre_(nombre);
  if (!activo) {
    if (u.fila === SES_.fila) throw new Error('No podés bloquear tu propio usuario.');
    if (u.rol === 'Administrador' && cantAdmins_() <= 1) throw new Error('Tiene que quedar al menos un administrador.');
  }
  usuariosHoja_().getRange(u.fila, 3).setValue(activo ? 'SI' : 'NO');
  if (!activo) usuariosHoja_().getRange(u.fila, 7).setValue(Date.now());
  usuInvalidar_();
  accesoLog_(SES_.nombre, activo ? 'Usuario habilitado' : 'Usuario bloqueado', u.nombre);
  return { ok: true };
}

function cambiarRolUsuario_(nombre, rol) {
  if (ROLES_.indexOf(rol) === -1) throw new Error('Rol inválido.');
  const u = usuarioPorNombre_(nombre);
  if (u.rol === 'Administrador' && rol !== 'Administrador' && cantAdmins_() <= 1) throw new Error('Tiene que quedar al menos un administrador.');
  usuariosHoja_().getRange(u.fila, 2).setValue(rol);
  usuInvalidar_();
  accesoLog_(SES_.nombre, 'Rol cambiado', u.nombre + ' → ' + rol);
  return { ok: true };
}

// ---- Entrada única desde las pantallas: verifica la sesión y el rol, y llama a la función de siempre ----
function registroApi_() {
  return {
    pagina: { f: pagina_ },
    yo: { f: yo_ },
    listarUsuarios: { f: listarUsuarios_, admin: true },
    crearUsuario: { f: crearUsuario_, admin: true },
    reiniciarPinUsuario: { f: reiniciarPinUsuario_, admin: true },
    activarUsuario: { f: activarUsuario_, admin: true },
    cambiarRolUsuario: { f: cambiarRolUsuario_, admin: true },
    guardarEntrevistaUsuario: { f: guardarEntrevistaUsuario_, admin: true },
    entrevistadoresAgenda: { f: entrevistadoresAgenda_ },
    agendarCitaTablero: { f: agendarCitaTablero_ },
    analizarEntrevista: { f: analizarEntrevista_ },
    anularEtapa: { f: anularEtapa_ },
    armarComunicacion: { f: armarComunicacion_ },
    borrarContactoRegistro: { f: borrarContactoRegistro_ },
    borrarDelBanco: { f: borrarDelBanco_ },
    borrarMarcaContacto: { f: borrarMarcaContacto_ },
    copiaAhoraBlumia: { f: copiaAhoraBlumia_, admin: true },
    doGetKanban: { f: doGetKanban_ },
    eliminarCandidato: { f: eliminarCandidato_ },
    enviarABanco: { f: enviarABanco_ },
    enviarABusqueda: { f: enviarABusqueda_ },
    enviarDescartesMasivo: { f: enviarDescartesMasivo_ },
    exportarContactosCSV: { f: exportarContactosCSV_ },
    generarInformeCandidato: { f: generarInformeCandidato_ },
    getConfigBusqueda: { f: getConfigBusqueda_, admin: true },
    configLazy: { f: configLazy_, admin: true },
    getContactosRegistro: { f: getContactosRegistro_ },
    getDatosPerfil: { f: getDatosPerfil_ },
    getEvaluacionPrevia: { f: getEvaluacionPrevia_ },
    getHistorialBusqueda: { f: getHistorialBusqueda_ },
    getHistorialCandidato: { f: getHistorialCandidato_ },
    getLinkScreening: { f: getLinkScreening_ },
    guardarConfigBusqueda: { f: guardarConfigBusqueda_, admin: true },
    guardarHabituales: { f: guardarHabituales_, admin: true },
    proponerBusquedaIA: { f: proponerBusquedaIA_, admin: true },
    crearBusquedaIA: { f: crearBusquedaIA_, admin: true },
    activarPerfilPuesto: { f: activarPerfilPuesto_, admin: true },
    copiaInfo: { f: copiaInfo_, admin: true },
    cambiarEstadoBusqueda: { f: cambiarEstadoBusqueda_, admin: true },
    diagnosticoVelocidad: { f: diagnosticoVelocidad_, admin: true },
    listarAccesos: { f: listarAccesos_, admin: true },
    guardarDatosCandidato: { f: guardarDatosCandidato_ },
    guardarEnPlanilla: { f: guardarEnPlanilla_ },
    flyerGuardar: { f: flyerGuardar_ },
    ibAnalizar: { f: ibAnalizar_ },
    ibRedactar: { f: ibRedactar_ },
    informePDF: { f: informePDF_ },
    informeEnviar: { f: informeEnviar_ },
    rankingGuardar: { f: rankingGuardar_ },
    vaciarBusquedaInfo: { f: vaciarBusquedaInfo_, admin: true },
    vaciarBusqueda: { f: vaciarBusqueda_, admin: true },
    flyerTextosIA: { f: flyerTextosIA_ },
    flyerSocial: { f: flyerSocial_ },
    difusionGuardar: { f: difusionGuardar_, admin: true },
    difusionTextos: { f: difusionTextosIA_, admin: true },
    flyerImagen: { f: flyerImagen_ },
    flyerRevisar: { f: flyerRevisarIA_, admin: true },
    guardarFormularioPerfil: { f: guardarFormularioPerfil_, admin: true },
    guardarInformeEditado: { f: guardarInformeEditado_ },
    guardarMotivoDescarte: { f: guardarMotivoDescarte_ },
    guardarObservaciones: { f: guardarObservaciones_ },
    linkPerfilConPin: { f: linkPerfilConPin_, admin: true },
    moverEstado: { f: moverEstado_ },
    notaEmpresa: { f: notaEmpresa_ },
    notaManual: { f: notaManual_ },
    previaDescartes: { f: previaDescartes_ },
    procesarScreening: { f: procesarScreening_ },
    recuperarEntrevista: { f: recuperarEntrevista_ },
    redactarIA: { f: redactarIA_ },
    registrarContactoManual: { f: registrarContactoManual_ },
    restaurarCandidato: { f: restaurarCandidato_ },
    transcribirAudio: { f: transcribirAudio_ }
  };
}

function api(token, nombre, args) {
  const u = sesionValida_(token);
  const def = registroApi_()[nombre];
  if (!def || typeof def.f !== 'function' || Object.prototype.hasOwnProperty.call(Object.prototype, nombre)) throw new Error('Función no disponible: ' + nombre);
  if (def.admin && !u.admin) throw new Error('Solo un administrador puede hacer esto.');
  SES_ = u; QUIEN_ = u.nombre;
  return def.f.apply(null, Array.isArray(args) ? args : []);
}

function enlacesAgendaPersonas_() {
  try { return leerUsuarios_().filter(u => u.activo === 'SI' && (u.linkScr || u.linkComp)).map(u => ['Agenda de ' + u.nombre, u.linkScr || u.linkComp]); } catch (e) { return []; }
}
