/* Perfil del puesto protegido con PIN (solo selectores).
   El PIN vive en la propiedad del script BLUMIA_PIN (no está en el repo).
   Un perfil queda protegido con Config.perfil_protegido = SI. */
function linkPerfilConPin(perfilId, pin) {
  const cache = CacheService.getScriptCache();
  const fallos = Number(cache.get('pin_fallos') || 0);
  if (fallos >= 5) throw new Error('Demasiados intentos. Probá de nuevo en 10 minutos.');
  const real = PropertiesService.getScriptProperties().getProperty('BLUMIA_PIN');
  if (!real) throw new Error('Falta configurar el PIN del perfil del puesto.');
  if (String(pin).trim() !== String(real).trim()) {
    cache.put('pin_fallos', String(fallos + 1), 600);
    throw new Error('PIN incorrecto.');
  }
  cache.remove('pin_fallos');
  const cfg = getConfig(perfilId);
  if (cfg.perfil_html) {
    const t = Utilities.getUuid();
    cache.put('pp_' + t, String(perfilId), 21600);
    return URL_DEPLOY + '?tipo=perfilpuesto&perfil=' + encodeURIComponent(perfilId) + '&t=' + t;
  }
  if (!cfg.link_perfil) throw new Error('Este perfil todavía no tiene el documento vinculado.');
  return String(cfg.link_perfil);
}

/* Sirve el HTML editable del perfil del puesto (Config.perfil_html) solo con el acceso temporal que da linkPerfilConPin. */
function perfilPuestoHTML(perfilId, t) {
  const msg = txt => HtmlService.createHtmlOutput('<p style="font-family:sans-serif;padding:40px;text-align:center">' + txt + '</p>');
  if (!t || CacheService.getScriptCache().get('pp_' + t) !== String(perfilId)) return msg('Acceso vencido. Abrí el botón «Perfil del puesto» desde el kanban e ingresá el PIN.');
  const cfg = getConfig(perfilId);
  if (!cfg.perfil_html) return msg('Este perfil no tiene documento vinculado.');
  return HtmlService.createHtmlOutputFromFile(String(cfg.perfil_html))
    .setTitle('Perfil del puesto | Blumia Consulting')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
