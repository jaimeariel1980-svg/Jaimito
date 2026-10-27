/* ============================================================
   PERFIL DEL PUESTO con IA: lee el Word (o texto) del puesto y completa los campos del relevamiento.
   - Búsqueda nueva: crearBusquedaIA_ llama a armarPerfilDesdeTexto_ y deja la primera versión guardada en Drive.
   - En el formulario: botón «Completar con IA desde un Word» (PerfilIA.html) → completarPerfilIA (acceso temporal del perfil).
   Solo completa lo que el documento dice; lo demás queda vacío para cargarlo a mano.
   ============================================================ */

const NO_RELLENAR_PERFIL_ = /firma|observaciones internas|notas internas/i;

function normOpc_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, ' ').trim(); }

function opcionesSelect_(c) { return (c.o || []).filter((o, k) => !(k === 0 && /^selecci/i.test(o))); }

function promptPerfilIA_(campos, pills) {
  const lc = campos.filter(c => !NO_RELLENAR_PERFIL_.test(c.l + ' ' + c.s)).map(c =>
    '#' + c.i + ' [' + c.t + '] ' + (c.s ? c.s + ' › ' : '') + (c.l || c.h || 'sin título') +
    (c.t === 'select' ? ' · opciones: ' + opcionesSelect_(c).join(' | ') : '') +
    (c.t === 'date' ? ' · formato AAAA-MM-DD' : '') + (c.t === 'number' ? ' · solo número' : '')).join('\n');
  const lp = pills.map(g => '@' + g.id + ' ' + g.s + ' · opciones: ' + g.o.join(' | ')).join('\n');
  return 'Sos consultor/a senior de selección de personal en Blumia Consulting (Argentina). Te paso el texto de un relevamiento o descripción de un puesto y la lista de campos de un formulario de «perfil del puesto». Completá SOLO los campos que el documento permite responder, con sus palabras ordenadas y claras (español rioplatense). No inventes datos, cifras, nombres ni requisitos: si el documento no lo dice, no incluyas ese campo. En los campos con opciones usá exactamente una de las opciones. Los campos de texto largo pueden llevar varias líneas.\n\nCAMPOS (número, tipo, sección › título):\n' + lc +
    '\n\nBOTONES DE OPCIÓN (id, sección, opciones):\n' + lp +
    '\n\nRespondé SOLO con un JSON válido, sin texto antes ni después ni markdown: {"c":{"<número>":"valor"},"p":{"<id sin @>":"opción"},"avisos":["datos importantes que el documento no decía o supuestos a confirmar"]}';
}

// devuelve {c:{indice:valor|opcion}, p:{id:opcion}, avisos:[]} ya limpio y validado contra los campos
function perfilIAValores_(campos, pills, texto) {
  texto = String(texto || '').replace(/\u0000/g, '').trim();
  if (texto.length < 200) throw new Error('El texto es muy corto para armar el perfil del puesto.');
  const r = extraerJSON_(procesarConClaude(promptPerfilIA_(campos, pills), texto.slice(0, 60000), 'analisis'));
  const out = { c: {}, p: {}, avisos: [] };
  const porI = {}; campos.forEach(c => porI[c.i] = c);
  Object.keys(r.c || {}).forEach(k => {
    const c = porI[k]; if (!c || NO_RELLENAR_PERFIL_.test(c.l + ' ' + c.s)) return;
    let v = limpiarTxt_(r.c[k]).slice(0, 4000);
    if (!v) return;
    if (c.t === 'select') { const o = opcionesSelect_(c).find(x => normOpc_(x) === normOpc_(v)); if (!o) return; v = o; }
    else if (c.t === 'number') { v = v.replace(/[^0-9.,]/g, ''); v = /^\d{1,3}(\.\d{3})+$/.test(v) ? v.replace(/\./g, '') : v.replace(',', '.'); if (!v || isNaN(Number(v))) return; }
    else if (c.t === 'date') { if (!/^\d{4}-\d{2}-\d{2}$/.test(v)) return; }
    else if (c.t === 'url') { if (!/^https?:\/\//i.test(v)) return; }
    else if (c.t === 'textarea') v = String(r.c[k]).replace(/[<>]/g, '').trim().slice(0, 4000);
    out.c[k] = v;
  });
  const gp = {}; pills.forEach(g => gp[g.id] = g);
  Object.keys(r.p || {}).forEach(id => {
    const g = gp[id]; if (!g) return;
    const o = g.o.find(x => normOpc_(x) === normOpc_(r.p[id])); if (o) out.p[id] = o;
  });
  out.avisos = (Array.isArray(r.avisos) ? r.avisos : []).slice(0, 8).map(a => limpiarTxt_(a).slice(0, 250)).filter(Boolean);
  return out;
}

// Desde el formulario abierto con el acceso temporal del perfil (mismo control que guardarPerfilDatos)
function completarPerfilIA(perfilId, tok, camposJson, pillsJson, texto) {
  if (!tok || CacheService.getScriptCache().get('pp_' + tok) !== String(perfilId)) throw new Error('Acceso vencido. Volvé a abrir el perfil.');
  let campos, pills;
  try { campos = JSON.parse(camposJson); pills = JSON.parse(pillsJson); } catch (e) { throw new Error('Datos del formulario inválidos.'); }
  if (!Array.isArray(campos) || campos.length > 200 || !Array.isArray(pills) || pills.length > 40) throw new Error('El formulario no tiene el formato esperado.');
  const cut = (v, n) => String(v == null ? '' : v).slice(0, n);
  campos = campos.map(c => ({ i: Number(c.i), t: cut(c.t, 12), s: cut(c.s, 120), l: cut(c.l, 200), h: cut(c.h, 200), o: Array.isArray(c.o) ? c.o.slice(0, 40).map(x => cut(x, 120)) : undefined }));
  pills = pills.map(g => ({ id: cut(g.id, 40), s: cut(g.s, 120), o: (g.o || []).slice(0, 12).map(x => cut(x, 120)) }));
  return JSON.stringify(perfilIAValores_(campos, pills, texto));
}

// Arma y guarda la primera versión del perfil del puesto de una búsqueda nueva (formato de guardarPerfilDatos: {v, p, n, pasos})
function armarPerfilDesdeTexto_(cfg, texto) {
  const K = PERFIL_CAMPOS_;
  const r = perfilIAValores_(K.campos, K.pills, texto);
  const v = K.campos.map(c => {
    const x = r.c[c.i];
    if (c.t === 'select') { const k = x ? c.o.findIndex(o => normOpc_(o) === normOpc_(x)) : 0; return k < 0 ? 0 : k; }
    return x || '';
  });
  const p = {};
  K.pills.forEach(g => { const k = r.p[g.id] ? g.o.findIndex(o => normOpc_(o) === normOpc_(r.p[g.id])) : g.on; p[g.id] = k; });
  const datos = JSON.stringify({ v: v, p: p, n: K.n, pasos: K.pasos });
  const sello = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH-mm-ss');
  carpetaPerfil_(cfg).createFile('Perfil del puesto - datos ' + sello + '.json', datos, MimeType.PLAIN_TEXT);
  return { campos: Object.keys(r.c).length, avisos: r.avisos };
}

// Para una búsqueda que todavía no tiene perfil propio: deja vinculado el formulario en blanco (reemplaza el link externo)
function activarPerfilPuesto_(pin, perfilId) {
  pinValido_(pin);
  const cfg = getConfig(perfilId);
  if (cfg.perfil_html) return { ok: true, yaTenia: true };
  setConfigCampo_(perfilId, 'perfil_html', 'RelevamientoPuesto');
  setConfigCampo_(perfilId, 'perfil_protegido', 'SI');
  return { ok: true };
}
