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
  if (!cfg.link_perfil) throw new Error('Este perfil todavía no tiene el documento vinculado.');
  return String(cfg.link_perfil);
}
