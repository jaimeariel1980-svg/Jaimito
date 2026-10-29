/* ============================================================
   FLYER Y TEXTOS PARA REDES de una búsqueda (menú 📚 Búsqueda → 📢 Difusión).
   - flyerHTML_: sirve Flyer.html con los datos de la búsqueda y el último diseño guardado.
   - flyerGuardar_: guarda el diseño (textos, colores, foto de fondo ya comprimida) como «Flyer - datos.json» en la carpeta de la búsqueda.
   - flyerTextosIA_: Claude redacta el flyer (bajada, desafíos, requisitos) y los posteos de LinkedIn / Instagram / WhatsApp.
   La foto cambia casi en cada búsqueda: se sube desde la pantalla y queda guardada con el diseño.
   ============================================================ */

const FLYER_ARCHIVO_ = 'Flyer - datos.json';
const FLYER_MAX_BYTES_ = 4 * 1024 * 1024;
const FLYER_TEXTOS_ = 'Flyer - textos.json';
const FLYER_IMAGEN_ = 'Flyer - imagen.jpg';
const FLYER_FORMATOS_ = { st: 'Historia 9:16', li: 'Post 4:5', ls: 'LinkedIn apaisado' };
const flyerImagenFmt_ = fmt => 'Flyer - imagen ' + fmt + '.jpg';

function flyerHero_(perfilId) {
  try {
    const rows = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Perfiles_Hero').getDataRange().getValues();
    const fila = rows.slice(1).find(r => r[0] === perfilId);
    if (!fila) return {};
    const o = {}; rows[0].forEach((h, i) => o[h] = fila[i]);
    return o;
  } catch (e) { return {}; }
}

function flyerArchivo_(cfg) {
  const it = carpetaPerfil_(cfg).getFilesByName(FLYER_ARCHIVO_);
  return it.hasNext() ? it.next() : null;
}

function flyerDatos_(perfilId) {
  const cfg = getConfig(perfilId);
  const hero = flyerHero_(perfilId);
  const puesto = (typeof cfg.nombre === 'string' && cfg.nombre) ? cfg.nombre : String(hero.posicion || perfilId);
  const empresa = empresaKey_(cfg);
  let guardado = '', fecha = '';
  try {
    const f = flyerArchivo_(cfg);
    if (f) { guardado = f.getBlob().getDataAsString(); fecha = Utilities.formatDate(f.getLastUpdated(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm'); }
  } catch (e) { }
  return {
    perfil: perfilId,
    puesto: puesto,
    empresa: busquedaCiega_(cfg) ? empresaParaCandidatos_(cfg) : empresa,
    ciega: busquedaCiega_(cfg),
    empresaNeutral: empresaParaCandidatos_(cfg),
    ubicacion: String(hero.ubicacion || ''),
    instancia: String(hero.instancia || ''),
    linkPostulacion: URL_WEBAPP + '?tipo=postulacion&perfil=' + encodeURIComponent(perfilId),
    mail: String(cfg.email_postulacion || 'empleos@blumiaconsulting.com'),
    linkViejo: (typeof cfg.link_flyer === 'string') ? cfg.link_flyer : '',
    guardado: guardado,
    fecha: fecha
  };
}

function flyerHTML_(perfilId) {
  const datos = flyerDatos_(perfilId);
  const js = o => JSON.stringify(o).replace(/</g, '\\u003c');
  let html = HtmlService.createTemplateFromFile('FlyerPagina').getRawContent();
  const lib = HtmlService.createTemplateFromFile('FlyerQRLib').getRawContent();
  const fuente = HtmlService.createTemplateFromFile('FuenteMarca').getRawContent();
  html = html.replace('%%FUENTE%%', () => fuente).replace('%%QRLIB%%', () => lib).replace('%%ASSETS%%', () => js(FLYER_ASSETS_)).replace('%%DATOS%%', () => js(datos));
  return HtmlService.createHtmlOutput(html).setTitle('Flyer · ' + datos.puesto + ' | Blumia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Guarda el diseño del flyer de la búsqueda (un solo archivo que se pisa; la foto ya viene comprimida desde la pantalla)
function flyerSocial_(perfilId) {
  const cache = CacheService.getScriptCache(), ck = 'fsoc_' + perfilId;
  try { const hit = cache.get(ck); if (hit) return JSON.parse(hit); } catch (e) { }
  const cfg = getConfig(perfilId), carpeta = carpetaPerfil_(cfg);
  const link = URL_WEBAPP + '?tipo=postulacion&perfil=' + encodeURIComponent(perfilId);
  const rep = t => String(t || '').replace(/\[LINK\]/g, link);
  let s = {};
  const ft = carpeta.getFilesByName(FLYER_TEXTOS_);
  if (ft.hasNext()) { try { s = JSON.parse(ft.next().getBlob().getDataAsString()) || {}; } catch (e) { } }
  else { // primera vez: se saca del archivo grande del diseño (trae la foto) y se deja una copia chica para las próximas
    const fd = carpeta.getFilesByName(FLYER_ARCHIVO_);
    if (fd.hasNext()) { try { s = (JSON.parse(fd.next().getBlob().getDataAsString()).social) || {}; } catch (e) { } try { carpeta.createFile(FLYER_TEXTOS_, JSON.stringify(s), MimeType.PLAIN_TEXT); } catch (e) { } }
  }
  let imagen = '', imagenId = '', fecha = '';
  try { const it = carpeta.getFilesByName(FLYER_IMAGEN_); if (it.hasNext()) { const fi = it.next(); imagen = fi.getUrl(); imagenId = fi.getId(); fecha = Utilities.formatDate(fi.getLastUpdated(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm'); } } catch (e) { }
  const imagenes = [];
  Object.keys(FLYER_FORMATOS_).forEach(fm => {
    try { const it = carpeta.getFilesByName(flyerImagenFmt_(fm)); if (it.hasNext()) { const fi = it.next(); imagenes.push({ fmt: fm, label: FLYER_FORMATOS_[fm], url: fi.getUrl(), id: fi.getId(), ts: fi.getLastUpdated().getTime(), fecha: Utilities.formatDate(fi.getLastUpdated(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm') }); } } catch (e) { }
  });
  imagenes.sort((a, b) => b.ts - a.ts); // la más reciente primero
  const out = { linkedin: rep(s.linkedin), instagram: rep(s.instagram), whatsapp: rep(s.whatsapp), link: link, fecha: fecha, imagen: imagen, imagenId: imagenId, imagenes: imagenes };
  try { cache.put(ck, JSON.stringify(out), 600); } catch (e) { }
  return out;
}

function flyerCacheBorrar_(perfilId) { try { CacheService.getScriptCache().remove('fsoc_' + perfilId); } catch (e) { } }

function flyerTextosEscribir_(carpeta, social) {
  const txt = JSON.stringify(social || {});
  const it = carpeta.getFilesByName(FLYER_TEXTOS_);
  if (it.hasNext()) it.next().setContent(txt); else carpeta.createFile(FLYER_TEXTOS_, txt, MimeType.PLAIN_TEXT);
}

function difusionGuardar_(perfilId, t) {
  const cfg = getConfig(perfilId);
  const lim = (v, n) => String(v == null ? '' : v).replace(/[<>]/g, '').slice(0, n);
  const f = flyerArchivo_(cfg);
  let g = {};
  if (f) { try { g = JSON.parse(f.getBlob().getDataAsString()) || {}; } catch (e) { g = {}; } }
  g.social = { linkedin: lim(t && t.linkedin, 3000), instagram: lim(t && t.instagram, 2200), whatsapp: lim(t && t.whatsapp, 1000) };
  const txt = JSON.stringify(g);
  if (txt.length > FLYER_MAX_BYTES_) throw new Error('No se pudo guardar: el archivo del flyer es demasiado grande.');
  if (f) f.setContent(txt); else carpetaPerfil_(cfg).createFile(FLYER_ARCHIVO_, txt, MimeType.PLAIN_TEXT);
  try { flyerTextosEscribir_(carpetaPerfil_(cfg), g.social); } catch (e) { }
  flyerCacheBorrar_(perfilId);
  try { registrarHistorial_(cfg.hoja_nombre || perfilId, '(búsqueda)', 'Difusión', 'Textos de difusión guardados', QUIEN_); } catch (e) { }
  return { fecha: Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm') };
}

// Redacta solo los tres textos de difusión (no toca el diseño del flyer)
function difusionTextosIA_(perfilId) {
  const d = flyerDatos_(perfilId);
  let g = {};
  try { g = JSON.parse(d.guardado || '{}') || {}; } catch (e) { }
  const mostrar = (g.mostrarEmpresa !== undefined) ? !!g.mostrarEmpresa : d.empresa === 'MACS';
  const r = flyerTextosIA_(perfilId, { mostrarEmpresa: mostrar });
  const rep = t => String(t || '').replace(/\[LINK\]/g, d.linkPostulacion);
  return { linkedin: rep(r.linkedin), instagram: rep(r.instagram), whatsapp: rep(r.whatsapp) };
}

// Revisión de diseño con IA: Claude mira la imagen del flyer, da sugerencias y, en cada una, una acción concreta que el editor puede aplicar solo.
// Solo acepta acciones de una lista cerrada y con valores acotados (la IA no toca nada fuera de eso).
const FLYER_COLORES_ = ['#0F3F27', '#1A1A1A', '#FFFFFF', '#FBC67B', '#FEE795', '#F9BDBB', '#B4E1E0'];
const FLYER_BLOQUES_ = ['logo', 'cli', 'bucle', 'aro', 'badge', 'tit', 'baj', 'secD', 'secR', 'qr', 'fTxt', 'arrow', 'fMail'];
const FLYER_BLOQUES_TEXTO_ = ['badge', 'tit', 'baj', 'secD', 'secR', 'qr', 'fTxt', 'arrow', 'fMail'];

function flyerAccionOk_(a) {
  if (!a || typeof a !== 'object') return null;
  const t = String(a.tipo || '');
  const num = (v, lo, hi) => { const n = Number(v); return isFinite(n) ? Math.max(lo, Math.min(hi, Math.round(n))) : null; };
  const col = v => { const c = String(v || '').toUpperCase(); return FLYER_COLORES_.indexOf(c) >= 0 ? c : null; };
  const blq = (v, lista) => (lista || FLYER_BLOQUES_).indexOf(String(v)) >= 0 ? String(v) : null;
  let r = null;
  if (t === 'colorTitulo') { const c = col(a.color); if (c) r = { tipo: t, color: c }; }
  else if (t === 'colorBloque') { const c = col(a.color), b = blq(a.bloque, FLYER_BLOQUES_TEXTO_); if (c && b) r = { tipo: t, bloque: b, color: c }; }
  else if (t === 'difuminado') { const n = num(a.valor, 0, 100); if (n !== null) r = { tipo: t, valor: n }; }
  else if (t === 'tamanoTexto') { const n = num(a.valor, 70, 130); if (n !== null) r = { tipo: t, valor: n }; }
  else if (t === 'tamanoBloque') { const n = num(a.valor, 40, 200), b = blq(a.bloque); if (n !== null && b) r = { tipo: t, bloque: b, valor: n }; }
  else if (t === 'mover') { const dx = num(a.dx, -600, 600), dy = num(a.dy, -600, 600), b = blq(a.bloque); if (dx !== null && dy !== null && b) r = { tipo: t, bloque: b, dx: dx, dy: dy }; }
  else if (t === 'dejarItems') { const n = num(a.n, 2, 4); if (n !== null) r = { tipo: t, n: n }; }
  else if (t === 'logoOscuro') r = { tipo: t, valor: !!a.valor };
  else if (t === 'altoPie') { const n = num(a.valor, 60, 150); if (n !== null) r = { tipo: t, valor: n }; }
  else if (t === 'serif') r = { tipo: t, valor: !!a.valor };
  else if (t === 'ocultar') { const b = blq(a.bloque, ['bucle', 'aro']); if (b) r = { tipo: t, bloque: b }; }
  else if (t === 'negrita') { const b = blq(a.bloque, FLYER_BLOQUES_TEXTO_); if (b) r = { tipo: t, bloque: b, valor: !!a.valor }; }
  else if (t === 'zonaSegura') r = { tipo: t };
  else if (t === 'acomodar') r = { tipo: t };
  else if (t === 'proponerTexto') { const b = blq(a.bloque, ['baj', 'fTxt', 'badge']); const x = String(a.texto || '').replace(/[<>]/g, '').trim().slice(0, 300); if (b && x) r = { tipo: t, bloque: b, texto: x }; }
  return r;
}

function flyerRevisarIA_(perfilId, base64, fmt, estado, ronda) {
  const cfg = getConfig(perfilId);
  const b64 = String(base64 || '').replace(/^data:image\/jpeg;base64,/, '');
  if (b64.length < 100 || b64.length > 6 * 1024 * 1024) throw new Error('La imagen para revisar no es válida o es demasiado grande.');
  const puesto = (typeof cfg.nombre === 'string' && cfg.nombre) ? cfg.nombre : perfilId;
  const sistema = 'Sos director/a de arte de Blumia Consulting (consultora de RR. HH. de Argentina). Revisás flyers de búsquedas laborales que se publican en historias de Instagram, posteos de LinkedIn y mensajes de WhatsApp. ' +
    'Marca: verde #0F3F27, naranja #FBC67B, rosa #F9BDBB, amarillo #FEE795, celeste #B4E1E0; tipografía Pangram; estilo cálido, limpio y profesional. ' +
    'Mirá la imagen y respondé SOLO con un JSON válido, sin texto antes ni después, con esta forma exacta: ' +
    '{"puntaje":entero de 1 a 10,"resumen":"una frase (máx. 140 caracteres)","sugerencias":[{"prioridad":"alta|media|baja","que":"qué mejorar (máx. 110 caracteres)","como":"cómo arreglarlo, en palabras simples (máx. 150 caracteres)","accion":null o un objeto de la lista de abajo}]}. ' +
    'Dá de 3 a 5 sugerencias, de la más a la menos importante. Evaluá en primer lugar la DISTRIBUCIÓN: todo centrado o alineado al mismo margen, márgenes parejos a izquierda y derecha, nada pegado a los bordes, y espacios parejos entre bloques sin huecos grandes ni elementos amontonados; si algo de eso está mal, la primera sugerencia es la acción acomodar. Después evaluá: legibilidad (contraste y tamaño de letra pensando en un celular), jerarquía, foto, equilibrio, zonas que tapa Instagram (250 px arriba y abajo en historia), coherencia con la marca y claridad de cómo postularse. ' +
    'ACCIONES (el editor las ejecuta solas; usá solo estas, con estos valores; si ninguna sirve, accion = null): ' +
    '{"tipo":"colorTitulo","color":"#FFFFFF"} (color de la paleta: #0F3F27 #1A1A1A #FFFFFF #FBC67B #FEE795 #F9BDBB #B4E1E0); ' +
    '{"tipo":"colorBloque","bloque":"baj","color":"#1A1A1A"}; {"tipo":"difuminado","valor":0-100} (más alto = la foto se funde más con el fondo claro); {"tipo":"tamanoTexto","valor":70-130}; ' +
    '{"tipo":"tamanoBloque","bloque":"tit","valor":40-200}; {"tipo":"mover","bloque":"fMail","dx":-600..600,"dy":-600..600} (px en un lienzo de 1080 de ancho; dy negativo sube); {"tipo":"dejarItems","n":2-4}; ' +
    '{"tipo":"logoOscuro","valor":true|false}; {"tipo":"altoPie","valor":60-150}; {"tipo":"serif","valor":true|false}; {"tipo":"ocultar","bloque":"bucle|aro"}; {"tipo":"negrita","bloque":"baj","valor":true|false}; ' +
    '{"tipo":"zonaSegura"} (saca de las zonas de Instagram lo que esté tapado); {"tipo":"acomodar"} (centra, alinea al mismo margen, separa de los bordes y parejiza los espacios entre bloques; el editor lo hace con medidas exactas); {"tipo":"proponerTexto","bloque":"baj|fTxt|badge","texto":"texto nuevo, más corto"} (solo si el texto es muy largo o confuso). ' +
    'Bloques válidos: logo, cli, bucle, aro, badge, tit, baj, secD, secR, qr, fTxt, arrow, fMail. ' +
    (Number(ronda) >= 2 ? 'Esta es la SEGUNDA revisión, después de aplicar cambios: marcá solo lo que siga mal; si el flyer ya está bien (puntaje 8 o más) devolvé pocas o ninguna sugerencia. ' : '') +
    'No inventes problemas: si algo está bien, decilo en el resumen. Español rioplatense, sin tecnicismos, con vos.';
  let est = '';
  try { est = JSON.stringify(estado || {}).slice(0, 1800); } catch (e) { }
  const usuario = [
    { type: 'image', source: { type: 'base64', media_type: 'image/jpeg', data: b64 } },
    { type: 'text', text: 'Formato: ' + (FLYER_FORMATOS_[fmt] || 'sin dato') + '. Puesto: ' + puesto + '.' + (busquedaCiega_(cfg) ? ' Es una búsqueda ciega: el nombre de la empresa cliente no debe aparecer.' : '') + (est ? ' Estado actual del editor (para que tus acciones tengan sentido): ' + est : '') }
  ];
  const r = extraerJSON_(procesarConClaude(sistema, usuario, 'redactar'));
  const t = (v, n) => String(v == null ? '' : v).replace(/[<>]/g, '').trim().slice(0, n);
  const pr = { alta: 1, media: 1, baja: 1 };
  return {
    puntaje: Math.max(1, Math.min(10, Math.round(Number(r.puntaje) || 0))),
    resumen: t(r.resumen, 200),
    sugerencias: (Array.isArray(r.sugerencias) ? r.sugerencias : []).slice(0, 6).map(x => ({ prioridad: pr[String(x && x.prioridad).toLowerCase()] ? String(x.prioridad).toLowerCase() : 'media', que: t(x && x.que, 160), como: t(x && x.como, 220), accion: flyerAccionOk_(x && x.accion) })).filter(x => x.que)
  };
}

function flyerGuardar_(perfilId, json) {
  const cfg = getConfig(perfilId);
  const txt = String(json || '');
  if (txt.length < 2 || txt.length > FLYER_MAX_BYTES_) throw new Error('El flyer es demasiado grande para guardarlo (probá con una foto más liviana).');
  try { JSON.parse(txt); } catch (e) { throw new Error('El diseño del flyer no es válido.'); }
  const f = flyerArchivo_(cfg);
  let out = txt;
  if (f) { // los textos de difusión se editan en ⚙ Configurar → Difusión: el editor del flyer no los pisa
    try { const nuevo = JSON.parse(txt), previo = JSON.parse(f.getBlob().getDataAsString() || '{}'); if (previo && previo.social) { nuevo.social = previo.social; out = JSON.stringify(nuevo); } } catch (e) { }
  }
  if (f) f.setContent(out); else carpetaPerfil_(cfg).createFile(FLYER_ARCHIVO_, out, MimeType.PLAIN_TEXT);
  flyerCacheBorrar_(perfilId);
  try { registrarHistorial_(cfg.hoja_nombre || perfilId, '(búsqueda)', 'Flyer', 'Diseño del flyer guardado', QUIEN_); } catch (e) { }
  return { fecha: Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm') };
}

// Guarda la imagen terminada del flyer (JPG) en la carpeta de la búsqueda con link de lectura, para enviarla por WhatsApp sin abrir el editor
function flyerImagen_(perfilId, base64, fmt) {
  const cfg = getConfig(perfilId);
  const b64 = String(base64 || '').replace(/^data:image\/jpeg;base64,/, '');
  if (b64.length < 100 || b64.length > 12 * 1024 * 1024) throw new Error('La imagen del flyer no es válida o es demasiado grande.');
  fmt = FLYER_FORMATOS_[fmt] ? fmt : '';
  const carpeta = carpetaPerfil_(cfg);
  const nombres = [FLYER_IMAGEN_].concat(fmt ? [flyerImagenFmt_(fmt)] : []);
  let url = '';
  nombres.forEach(nom => {
    const it = carpeta.getFilesByName(nom);
    while (it.hasNext()) it.next().setTrashed(true);
    const f = carpeta.createFile(Utilities.newBlob(Utilities.base64Decode(b64), 'image/jpeg', nom));
    f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    if (!url) url = f.getUrl();
  });
  flyerCacheBorrar_(perfilId);
  return { imagen: url };
}

// Valores cargados en el perfil del puesto (relevamiento) como texto «campo: valor», para que la IA tenga de dónde sacar desafíos y requisitos
function flyerPerfilTexto_(cfg) {
  try {
    const u = ultimoPerfilDatos_(cfg);
    if (!u) return '';
    const d = JSON.parse(u.texto), K = PERFIL_CAMPOS_;
    const L = [];
    K.campos.forEach((c, idx) => {
      if (NO_RELLENAR_PERFIL_.test(c.l + ' ' + c.s)) return;
      let v = d.v ? d.v[idx] : '';
      if (c.t === 'select') { const k = Number(v); v = (k > 0 && c.o && c.o[k]) ? c.o[k] : ''; }
      v = String(v == null ? '' : v).trim();
      if (!v) return;
      L.push((c.s ? c.s + ' › ' : '') + (c.l || c.h || '') + ': ' + v);
    });
    return L.join('\n').slice(0, 14000);
  } catch (e) { return ''; }
}

// Opciones: { mostrarEmpresa: true|false }. Devuelve { badge, bajada, desafios[], requisitos[], linkedin, instagram, whatsapp }
function flyerTextosIA_(perfilId, opciones) {
  const cfg = getConfig(perfilId);
  const d = flyerDatos_(perfilId);
  const mostrar = !!(opciones && opciones.mostrarEmpresa) && !d.ciega;
  const empresa = mostrar ? d.empresa : d.empresaNeutral;
  const perfil = flyerPerfilTexto_(cfg);
  const extra = (typeof cfg.prompt_extra === 'string') ? cfg.prompt_extra.slice(0, 5000) : '';
  if (!perfil && !extra) throw new Error('Esta búsqueda todavía no tiene datos para redactar el flyer: cargá el perfil del puesto en ⚙ Configurar (o escribí los textos a mano en la pantalla).');
  const sistema = 'Sos redactor/a de avisos de empleo de Blumia Consulting (Argentina). Con los datos de una búsqueda armás el texto de un flyer y tres posteos. ' +
    'Respondé SOLO con un JSON válido, sin texto antes ni después ni markdown, con esta forma exacta: ' +
    '{"badge":"ubicación · modalidad (máx. 40 caracteres, ej. «Pilar · Presencial»)","bajada":"2 o 3 frases (máx. 280 caracteres) que presentan a la empresa y qué busca","desafios":["4 o 5 ítems, cada uno máx. 95 caracteres, empiezan con verbo en infinitivo"],"requisitos":["4 ítems, cada uno máx. 95 caracteres"],' +
    '"linkedin":"posteo para LinkedIn (máx. 900 caracteres, con saltos de línea, 2 o 3 emojis como mucho, termina invitando a postularse en [LINK])","instagram":"posteo para Instagram (máx. 600 caracteres, con 4 o 5 hashtags, invita a postularse en el link de la bio o por [LINK])","whatsapp":"mensaje corto para difundir por WhatsApp (máx. 450 caracteres, con [LINK])"}. ' +
    'Reglas: español rioplatense profesional (vos); usá lenguaje inclusivo con barra cuando el puesto lo requiera (Coordinador/a); ' +
    'no inventes datos que no estén en el material; no pongas sueldo, banda salarial, puntajes ni criterios internos de evaluación; ' +
    'el nombre de la empresa es exactamente «' + empresa + '»' + (mostrar ? '' : ' (es confidencial: no nombres ni insinúes qué empresa es)') + '; ' +
    'escribí [LINK] tal cual donde va el link de postulación.';
  const usuario = 'PUESTO: ' + d.puesto + '\nEMPRESA: ' + empresa + '\nUBICACIÓN: ' + (d.ubicacion || '(sin dato)') + 
    (perfil ? '\n\nPERFIL DEL PUESTO (relevamiento):\n' + perfil : '') + (extra ? '\n\nOTROS DATOS DE LA BÚSQUEDA (uso interno, no copies criterios de evaluación):\n' + extra : '');
  const r = extraerJSON_(procesarConClaude(sistema, usuario, 'redactar'));
  const lista = (a, n, max) => (Array.isArray(a) ? a : []).slice(0, n).map(x => limpiarTxt_(x).slice(0, max)).filter(Boolean);
  const txt = (v, max) => String(v == null ? '' : v).replace(/[<>]/g, '').trim().slice(0, max);
  return {
    badge: txt(r.badge, 60), bajada: txt(r.bajada, 400),
    desafios: lista(r.desafios, 6, 140), requisitos: lista(r.requisitos, 6, 140),
    linkedin: txt(r.linkedin, 1500), instagram: txt(r.instagram, 1000), whatsapp: txt(r.whatsapp, 800)
  };
}
