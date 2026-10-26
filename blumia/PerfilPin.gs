/* Perfil del puesto protegido con PIN (solo selectores).
   El PIN vive en la propiedad del script BLUMIA_PIN (no está en el repo).
   Un perfil queda protegido con Config.perfil_protegido = SI. */
/* Verifica el PIN de los selectores (propiedad BLUMIA_PIN). 5 errores bloquean 10 minutos. Lo usan el perfil del puesto y la IA. */
function pinValido_(pin) {
  const cache = CacheService.getScriptCache();
  const fallos = Number(cache.get('pin_fallos') || 0);
  if (fallos >= 5) throw new Error('Demasiados intentos. Probá de nuevo en 10 minutos.');
  const real = PropertiesService.getScriptProperties().getProperty('BLUMIA_PIN');
  if (!real) throw new Error('Falta configurar el PIN.');
  if (String(pin).trim() !== String(real).trim()) {
    cache.put('pin_fallos', String(fallos + 1), 600);
    throw new Error('PIN incorrecto.');
  }
  cache.remove('pin_fallos');
}

function linkPerfilConPin(perfilId, pin) {
  pinValido_(pin);
  const cache = CacheService.getScriptCache();
  const cfg = getConfig(perfilId);
  if (cfg.perfil_html) {
    const t = Utilities.getUuid();
    cache.put('pp_' + t, String(perfilId), 21600);
    return URL_DEPLOY + '?tipo=perfilpuesto&perfil=' + encodeURIComponent(perfilId) + '&t=' + t;
  }
  if (!cfg.link_perfil) throw new Error('Este perfil todavía no tiene el documento vinculado.');
  return String(cfg.link_perfil);
}

/* Sirve el HTML editable del perfil del puesto (Config.perfil_html) solo con el acceso temporal que da linkPerfilConPin.
   Si hay versiones guardadas con guardarPerfilDatos, carga la última sobre el formulario. */
function perfilPuestoHTML(perfilId, tok) {
  const msg = txt => HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:40px;text-align:center">' + txt + '</p>');
  if (!tok || CacheService.getScriptCache().get('pp_' + tok) !== String(perfilId)) return msg('Acceso vencido. Abrí el botón «Perfil del puesto» desde el kanban e ingresá el PIN.');
  const cfg = getConfig(perfilId);
  if (!cfg.perfil_html) return msg('Este perfil no tiene documento vinculado.');
  const plantilla = HtmlService.createTemplateFromFile(String(cfg.perfil_html));
  const ult = ultimoPerfilDatos_(cfg);
  plantilla.perfilId = String(perfilId);
  plantilla.token = String(tok);
  plantilla.datos = ult ? ult.texto.replace(/</g, '\\u003c') : 'null';
  plantilla.fecha = ult ? ult.fecha : '';
  return plantilla.evaluate()
    .setTitle('Perfil del puesto | Blumia Consulting')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function carpetaPerfil_(cfg) { return DriveApp.getFolderById(cfg.folder_madre).getParents().next(); }

function ultimoPerfilDatos_(cfg) {
  const it = carpetaPerfil_(cfg).getFilesByType(MimeType.PLAIN_TEXT);
  let mejor = null;
  while (it.hasNext()) { const f = it.next(); if (f.getName().indexOf('Perfil del puesto - datos ') === 0 && (!mejor || f.getDateCreated() > mejor.getDateCreated())) mejor = f; }
  if (!mejor) return null;
  return { texto: mejor.getBlob().getDataAsString(), fecha: Utilities.formatDate(mejor.getDateCreated(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm') };
}

/* Guarda una versión con fecha y hora del perfil del puesto (los valores del formulario) en la carpeta del perfil en Drive.
   Solo se puede con el acceso temporal que da linkPerfilConPin (o sea, después del PIN). */
function guardarPerfilDatos(perfilId, tok, datosJson) {
  if (!tok || CacheService.getScriptCache().get('pp_' + tok) !== String(perfilId)) throw new Error('Acceso vencido. Volvé a abrir el perfil desde el kanban con el PIN.');
  const txt = String(datosJson || '');
  if (txt.length > 300000) throw new Error('El perfil es demasiado grande para guardarlo.');
  const d = JSON.parse(txt);
  if (!d || !Array.isArray(d.v)) throw new Error('Datos inválidos.');
  const cfg = getConfig(perfilId);
  const ahora = new Date();
  const sello = Utilities.formatDate(ahora, 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH-mm-ss');
  carpetaPerfil_(cfg).createFile('Perfil del puesto - datos ' + sello + '.json', txt, MimeType.PLAIN_TEXT);
  return Utilities.formatDate(ahora, 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm');
}

/* Diagnóstico: muestra cómo está configurado el perfil del puesto y si el HTML se puede cargar. */
function diagnosticoPerfilPuesto() {
  const cfg = getConfig('director-comercial');
  Logger.log('perfil_protegido: ' + cfg.perfil_protegido + ' | perfil_html: ' + cfg.perfil_html + ' | link_perfil: ' + String(cfg.link_perfil || '').slice(0, 60));
  Logger.log('BLUMIA_PIN cargado: ' + (PropertiesService.getScriptProperties().getProperty('BLUMIA_PIN') ? 'sí' : 'NO'));
  if (cfg.perfil_html) {
    const c = HtmlService.createTemplateFromFile(String(cfg.perfil_html)).evaluate().getContent();
    Logger.log('HTML cargado: ' + c.length + ' caracteres, contiene formulario: ' + (c.indexOf('Perfil del puesto') > -1 ? 'sí' : 'NO'));
  }
}
