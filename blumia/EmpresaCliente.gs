/* Nombre de la empresa cliente según el perfil (Config.empresa).
   - empresaDefault_: empresa de las búsquedas que no tienen el dato. Se cambia sin publicar con la propiedad del script `EMPRESA_DEFAULT` (sin propiedad = MACS, como siempre).
   - EMPRESAS_CONOCIDAS_: datos propios de una empresa (nombre legal en los informes, logo claro de la postulación, nombre de la hoja del informe de proceso).
     Una empresa que no está en la lista usa su nombre tal cual y no tiene logo hasta que se cargue Config.logo_cliente.
   - empresaInforme_: lo que ve el cliente en los informes.
   - empresaParaCandidatos_: lo que leen los candidatos en mails y WhatsApp. Si Config tiene la columna opcional
     `empresa_candidatos`, se usa tal cual (ej.: «una empresa cliente del sector Energía y Minería» o el nombre real);
     si no, para una empresa que no figura en EMPRESAS_CONOCIDAS_ no se revela el nombre: «una empresa cliente del sector <empresa>». */
const EMPRESAS_CONOCIDAS_ = {
  'MACS': { legal: 'MACS S.A.', logo: 'https://azcubot.pages.dev/macs-logo-blanco.png', hoja: 'Informe_Proceso' }
};

function empresaDefault_() {
  try { return PropertiesService.getScriptProperties().getProperty('EMPRESA_DEFAULT') || 'MACS'; } catch (e) { return 'MACS'; }
}

/* Clave de la empresa de un perfil (Config.empresa; sin dato = empresa por defecto). */
function empresaKey_(cfg) {
  const e = (cfg && typeof cfg.empresa === 'string') ? cfg.empresa.trim() : '';
  return e || empresaDefault_();
}

function empresaInforme_(cfg) {
  const e = empresaKey_(cfg);
  return EMPRESAS_CONOCIDAS_[e] ? EMPRESAS_CONOCIDAS_[e].legal : e;
}

function empresaParaCandidatos_(cfg) {
  const x = (cfg && typeof cfg.empresa_candidatos === 'string') ? cfg.empresa_candidatos.trim() : '';
  if (x) return x;
  const e = empresaKey_(cfg);
  return EMPRESAS_CONOCIDAS_[e] ? e : 'una empresa cliente del sector ' + e;
}

function logoConocido_(empresa) {
  const k = EMPRESAS_CONOCIDAS_[empresa];
  return k && k.logo ? k.logo : '';
}

/* Nombre de la hoja del informe de proceso de una empresa. */
function hojaProceso_(empresa) {
  const e = empresa || empresaDefault_(), k = EMPRESAS_CONOCIDAS_[e];
  return k && k.hoja ? k.hoja : 'Informe_Proceso_' + e;
}
