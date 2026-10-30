/* Perfil del puesto protegido con PIN (solo selectores).
   El PIN vive en la propiedad del script BLUMIA_PIN (no está en el repo).
   Un perfil queda protegido con Config.perfil_protegido = SI. */
/* Verifica el PIN de los selectores (propiedad BLUMIA_PIN). 5 errores bloquean 10 minutos. Lo usan el perfil del puesto y la IA. */
function pinValido_(pin) {
  if (typeof SES_ !== 'undefined' && SES_) return; // ya ingresó con su usuario (el rol se controla en api)
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

function linkPerfilConPin_(perfilId, pin, modo) {
  pinValido_(pin);
  const cache = CacheService.getScriptCache();
  const cfg = getConfig(perfilId);
  if (cfg.perfil_html) {
    const t = Utilities.getUuid();
    cache.put('pp_' + t, String(perfilId), 21600);
    return URL_DEPLOY + '?tipo=perfilpuesto&perfil=' + encodeURIComponent(perfilId) + '&t=' + t + (modo === 'editar' ? '&m=editar' : '&m=ver');
  }
  if (!cfg.link_perfil) throw new Error('Este perfil todavía no tiene el documento vinculado.');
  return String(cfg.link_perfil);
}

/* Sirve el HTML editable del perfil del puesto (Config.perfil_html) solo con el acceso temporal que da linkPerfilConPin.
   Si hay versiones guardadas con guardarPerfilDatos, carga la última sobre el formulario. */
function perfilPuestoHTML(perfilId, tok, modo) {
  const msg = txt => HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:40px;text-align:center">' + txt + '</p>');
  if (!tok || CacheService.getScriptCache().get('pp_' + tok) !== String(perfilId)) return msg('Acceso vencido. Abrí el botón «Perfil del puesto» desde el tablero e ingresá el PIN.');
  const cfg = getConfig(perfilId);
  if (!cfg.perfil_html) return msg('Este perfil no tiene documento vinculado.');
  const plantilla = HtmlService.createTemplateFromFile(String(cfg.perfil_html));
  const ult = ultimoPerfilDatos_(cfg);
  plantilla.perfilId = String(perfilId);
  plantilla.token = String(tok);
  plantilla.datos = ult ? ult.texto.replace(/</g, '\\u003c') : 'null';
  plantilla.fecha = ult ? ult.fecha : '';
  const out = plantilla.evaluate()
    .setTitle('Perfil del puesto | Blumia Consulting')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  // modo «ver»: la versión definitiva guardada, solo para consultar (no se edita ni se guarda; se puede exportar a PDF). Para editar: ⚙ Configurar → Perfil del puesto.
  if (modo !== 'editar') out.append('<style>.pill,.opt,label.pill{pointer-events:none}#btn-sis{display:none!important}</style><script>window.addEventListener("load",function(){document.querySelectorAll("input,textarea,select").forEach(function(e){e.disabled=true;});var b=document.createElement("div");b.style.cssText="position:fixed;left:50%;transform:translateX(-50%);bottom:12px;z-index:9999;background:#0F3F27;color:#fff;padding:8px 18px;border-radius:99px;font:600 12px Arial,sans-serif;box-shadow:0 4px 14px rgba(0,0,0,.3)";b.textContent="Versión guardada' + (ult ? ' el ' + ult.fecha : '') + ' · solo lectura";document.body.appendChild(b);var g=document.createElement("button");g.type="button";g.title="Editar este perfil";g.setAttribute("aria-label","Editar este perfil");g.innerHTML="&#9881; Editar";g.style.cssText="position:fixed;bottom:16px;right:16px;z-index:9999;height:40px;padding:0 16px;border:0;border-radius:99px;background:#FEE795;color:#0F3F27;font:700 13px Arial,sans-serif;cursor:pointer;box-shadow:0 4px 14px rgba(0,0,0,.3)";g.onclick=function(){location.href=location.href.replace("m=ver","m=editar");};document.body.appendChild(g);});</script>');
  else {
    out.append(HtmlService.createHtmlOutputFromFile('PerfilIAForm').getContent()); // modo «editar»: botón «Completar con IA desde un Word»
    // botón fijo «Guardar cambios» (hace lo mismo que «Guardar en el sistema») para tenerlo a mano sin bajar hasta el final
    out.append('<script>window.addEventListener("load",function(){var s=document.getElementById("btn-sis");if(!s)return;var g=document.createElement("button");g.type="button";g.textContent="Guardar cambios";g.style.cssText="position:fixed;bottom:16px;right:16px;z-index:9999;height:40px;padding:0 18px;border:0;border-radius:99px;background:#FEE795;color:#0F3F27;font:700 13px Arial,sans-serif;cursor:pointer;box-shadow:0 4px 14px rgba(0,0,0,.3)";g.onclick=function(){s.click();};document.body.appendChild(g);});</script>');
  }
  return out;
}

// fecha de la última versión guardada del perfil del puesto (para ⚙ Configurar)
function perfilPuestoInfo_(cfg) {
  if (!cfg.perfil_html) return { propio: false, fecha: '' };
  try { const u = ultimoPerfilDatos_(cfg); return { propio: true, fecha: u ? u.fecha : '' }; } catch (e) { return { propio: true, fecha: '' }; }
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
   Solo se puede con el acceso temporal que da linkPerfilConPin_ (o sea, después del PIN). */
function guardarPerfilDatos(perfilId, tok, datosJson, forzar) {
  if (!tok || CacheService.getScriptCache().get('pp_' + tok) !== String(perfilId)) throw new Error('Acceso vencido. Volvé a abrir el perfil desde el tablero con el PIN.');
  const txt = String(datosJson || '');
  if (txt.length > 300000) throw new Error('El perfil es demasiado grande para guardarlo.');
  const d = JSON.parse(txt);
  if (!d || !Array.isArray(d.v)) throw new Error('Datos inválidos.');
  const cfg = getConfig(perfilId);
  // Red de seguridad (29/10: se guardó una versión casi vacía encima de una completa): si la nueva tiene menos de la mitad de datos que la última, pide confirmar
  if (!forzar) {
    const llenos = v => (Array.isArray(v) ? v : []).filter(x => { const t = String(x == null ? '' : x).trim(); return t !== '' && t !== '0' && t !== '-1'; }).length;
    let antes = 0;
    try { const u = ultimoPerfilDatos_(cfg); if (u) antes = llenos((JSON.parse(u.texto) || {}).v); } catch (e) { }
    const ahoraN = llenos(d.v);
    if (antes >= 15 && ahoraN < antes / 2) throw new Error('POCOS_DATOS|La versión que vas a guardar tiene ' + ahoraN + ' datos cargados y la última guardada tiene ' + antes + '. Si seguís, la última pasa a ser esta (la anterior queda en Drive). ¿Guardar igual?');
  }
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
