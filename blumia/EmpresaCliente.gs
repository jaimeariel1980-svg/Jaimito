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

/* Búsqueda ciega (Config.busqueda_ciega = SI): a los candidatos no se les muestra el nombre ni el logo de la empresa cliente
   (postulación, mensajes, mails, flyer y textos para redes). Se activa en ⚙ Configurar → Datos de la búsqueda. */
function busquedaCiega_(cfg) { return !!cfg && String(cfg.busqueda_ciega || '').toUpperCase().trim() === 'SI'; }

function empresaParaCandidatos_(cfg) {
  const x = (cfg && typeof cfg.empresa_candidatos === 'string') ? cfg.empresa_candidatos.trim() : '';
  if (busquedaCiega_(cfg)) return x && x.toLowerCase() !== empresaKey_(cfg).toLowerCase() ? x : 'una empresa cliente';
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

/* Mails que ve el candidato o el cliente (29/10, Ariel: «tiene que salir de Blumia Consulting»): salen con el nombre «Blumia Consulting» y desde la dirección
   de la empresa (propiedad del script MAIL_ENVIO; si no está, empleos@blumiaconsulting.com). Para que el remitente sea esa dirección, tiene que estar dada de alta en
   Gmail de la cuenta del sistema (Configuración → Cuentas → «Enviar como»); si no lo está, el mail igual sale (de la cuenta del sistema, con el nombre de Blumia) y las
   respuestas van a la dirección de la empresa. Devuelve { alias, desde }. */
function remitenteBlumia_() {
  const p = PropertiesService.getScriptProperties().getProperty('MAIL_ENVIO');
  return String(p || 'empleos@blumiaconsulting.com').trim();
}
function enviarMailBlumia_(para, asunto, cuerpo, opciones) {
  const o = Object.assign({ name: 'Blumia Consulting' }, opciones || {});
  const desde = remitenteBlumia_();
  let alias = false;
  try { alias = GmailApp.getAliases().some(a => String(a).toLowerCase() === desde.toLowerCase()); } catch (e) { }
  if (alias) o.from = desde; else if (!o.replyTo) o.replyTo = desde;
  GmailApp.sendEmail(para, asunto, cuerpo, o);
  return { alias: alias, desde: alias ? desde : Session.getEffectiveUser().getEmail() };
}
