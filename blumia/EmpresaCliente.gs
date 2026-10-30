/* Nombre de la empresa cliente según el perfil (Config.empresa).
   - empresaDefault_: empresa de las búsquedas que no tienen el dato. Se cambia sin publicar con la propiedad del script `EMPRESA_DEFAULT` (sin propiedad = MACS, como siempre).
   - EMPRESAS_CONOCIDAS_: datos propios de una empresa (nombre legal en los informes, logo claro de la postulación, nombre de la hoja del informe de proceso).
     Una empresa que no está en la lista usa su nombre tal cual y no tiene logo hasta que se cargue Config.logo_cliente.
   - empresaInforme_: lo que ve el cliente en los informes.
   - empresaParaCandidatos_: lo que leen los candidatos en mails y WhatsApp. Si Config tiene la columna opcional
     `empresa_candidatos`, se usa tal cual (ej.: «una empresa cliente del sector Energía y Minería» o el nombre real);
     si no, para una empresa que no figura en EMPRESAS_CONOCIDAS_ no se revela el nombre: «una empresa cliente del sector <empresa>». */
const EMPRESAS_CONOCIDAS_ = {
  'MACS': { legal: 'MACS S.A.', logo: 'https://azcubot.pages.dev/macs-logo-blanco.png', hoja: 'Informe_Proceso' },
  'ABECEB': { legal: 'ABECEB', logo: 'https://azcubot.pages.dev/abeceb-logo-blanco.png', logoColor: 'https://azcubot.pages.dev/abeceb-logo.png' }
};

// busca la empresa sin importar mayúsculas ni tildes («Abeceb», «ABECEB »)
function empresaConocida_(e) {
  const n = s => String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').trim();
  const k = Object.keys(EMPRESAS_CONOCIDAS_).find(x => n(x) === n(e));
  return k ? EMPRESAS_CONOCIDAS_[k] : null;
}

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
  const k = empresaConocida_(e);
  return k ? k.legal : e;
}

/* Búsqueda ciega (Config.busqueda_ciega = SI): a los candidatos no se les muestra el nombre ni el logo de la empresa cliente
   (postulación, mensajes, mails, flyer y textos para redes). Se activa en ⚙ Configurar → Datos de la búsqueda. */
function busquedaCiega_(cfg) { return !!cfg && String(cfg.busqueda_ciega || '').toUpperCase().trim() === 'SI'; }

function empresaParaCandidatos_(cfg) {
  const x = (cfg && typeof cfg.empresa_candidatos === 'string') ? cfg.empresa_candidatos.trim() : '';
  if (busquedaCiega_(cfg)) return x && x.toLowerCase() !== empresaKey_(cfg).toLowerCase() ? x : 'una empresa cliente';
  if (x) return x;
  const e = empresaKey_(cfg);
  return empresaConocida_(e) ? e : 'una empresa cliente del sector ' + e;
}

function logoConocido_(empresa) {
  const k = empresaConocida_(empresa);
  return k && k.logo ? k.logo : '';
}

/* Logo a color de la empresa cliente (para fondos claros: informe de la búsqueda, flyer); si no hay uno a color, el claro de siempre. */
function logoClienteColor_(cfg) {
  const c = (cfg && typeof cfg.logo_cliente_color === 'string') ? cfg.logo_cliente_color.trim() : '';
  if (/^https:\/\//i.test(c) || /^data:image\/png;base64,/i.test(c)) return c;
  const k = empresaConocida_(empresaKey_(cfg));
  return k ? (k.logoColor || k.logo || '') : '';
}

/* Logo del cliente en el encabezado de un informe para el cliente: nada si la búsqueda es ciega o la empresa no tiene logo. claro = fondo oscuro (logo blanco). */
function heroLogoCli_(cfg, claro) {
  if (busquedaCiega_(cfg)) return '';
  const u = claro ? logoCliente_(cfg) : logoClienteColor_(cfg);
  return u ? '<img class="hero-cli" src="' + u + '" alt="" style="float:right;height:36px;max-width:160px;object-fit:contain;margin:2px 0 0 12px">' : '';
}

/* Nombre de la hoja del informe de proceso de una empresa. */
function hojaProceso_(empresa) {
  const e = empresa || empresaDefault_(), k = empresaConocida_(e);
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

/* Logo claro del cliente como imagen incrustada (data:) para el flyer: así la exportación a imagen no depende de permisos entre sitios. Vacío si la búsqueda no tiene logo. */
function logoClienteDatos_(cfg) {
  try {
    const u = logoCliente_(cfg);
    if (!u) return '';
    if (/^data:/i.test(u)) return u;
    const cache = CacheService.getScriptCache(), k = 'logocli_' + Utilities.base64EncodeWebSafe(u).slice(-40);
    const c = cache.get(k);
    if (c) return c;
    const r = UrlFetchApp.fetch(u, { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) return '';
    const d = 'data:image/png;base64,' + Utilities.base64Encode(r.getContent());
    if (d.length < 90000) cache.put(k, d, 21600);
    return d;
  } catch (e) { return ''; }
}
