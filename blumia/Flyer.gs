/* ============================================================
   FLYER Y TEXTOS PARA REDES de una búsqueda (menú 📚 Búsqueda → 📢 Flyer y textos).
   - flyerHTML_: sirve Flyer.html con los datos de la búsqueda y el último diseño guardado.
   - flyerGuardar_: guarda el diseño (textos, colores, foto de fondo ya comprimida) como «Flyer - datos.json» en la carpeta de la búsqueda.
   - flyerTextosIA_: Claude redacta el flyer (bajada, desafíos, requisitos) y los posteos de LinkedIn / Instagram / WhatsApp.
   La foto cambia casi en cada búsqueda: se sube desde la pantalla y queda guardada con el diseño.
   ============================================================ */

const FLYER_ARCHIVO_ = 'Flyer - datos.json';
const FLYER_MAX_BYTES_ = 4 * 1024 * 1024;

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
    empresa: empresa,
    empresaNeutral: empresaParaCandidatos_(cfg),
    ubicacion: String(hero.ubicacion || ''),
    instancia: String(hero.instancia || ''),
    linkPostulacion: URL_WEBAPP + '?tipo=postulacion&perfil=' + encodeURIComponent(perfilId),
    mail: 'empleos@blumiaconsulting.com',
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
  const d = flyerDatos_(perfilId);
  let s = {};
  try { s = (JSON.parse(d.guardado || '{}').social) || {}; } catch (e) { }
  const rep = t => String(t || '').replace(/\[LINK\]/g, d.linkPostulacion);
  return { linkedin: rep(s.linkedin), instagram: rep(s.instagram), whatsapp: rep(s.whatsapp), link: d.linkPostulacion, fecha: d.fecha };
}

function flyerGuardar_(perfilId, json) {
  const cfg = getConfig(perfilId);
  const txt = String(json || '');
  if (txt.length < 2 || txt.length > FLYER_MAX_BYTES_) throw new Error('El flyer es demasiado grande para guardarlo (probá con una foto más liviana).');
  try { JSON.parse(txt); } catch (e) { throw new Error('El diseño del flyer no es válido.'); }
  const f = flyerArchivo_(cfg);
  if (f) f.setContent(txt); else carpetaPerfil_(cfg).createFile(FLYER_ARCHIVO_, txt, MimeType.PLAIN_TEXT);
  try { registrarHistorial_(cfg.hoja_nombre || perfilId, '(búsqueda)', 'Flyer', 'Diseño del flyer guardado', QUIEN_); } catch (e) { }
  return { fecha: Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm') };
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
  const mostrar = !!(opciones && opciones.mostrarEmpresa);
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
