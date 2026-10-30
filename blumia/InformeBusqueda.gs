/* Informe general de la búsqueda para el cliente, en 3 pasos (pedido de Ariel 29/10):
   1) la IA analiza los números y propone hallazgos; 2) hace unas pocas preguntas con respuesta sugerida sobre lo que los datos no dicen;
   3) redacta el informe con lo que Ariel dejó y respondió. Los números salen siempre de la planilla (nunca los inventa la IA).
   Mismo formato y contenido que el Informe de proceso (embudo, realidad salarial, industrias, descartes, fuentes) más rango de edad,
   género y zona agregados. No viajan nombres de candidatos ni datos individuales ni situación familiar. Página: ?tipo=informebusqueda&perfil=<id> */

// El informe siempre arranca con este título (pedido de Ariel); se puede editar en el informe armado
const IB_TITULO_ = 'El proceso está en marcha, el mercado respondió.';

function informeBusquedaHTML_(perfilId, alcance) {
  const cfg = getConfig(perfilId);
  const emp = alcance === 'empresa';
  const datos = { perfil: perfilId, alcance: emp ? 'empresa' : '', puesto: emp ? 'Todas las búsquedas' : String(cfg.nombre || perfilId), empresa: empresaInforme_(cfg), volver: URL_WEBAPP + '?tipo=kanban&perfil=' + encodeURIComponent(perfilId) };
  const js = o => JSON.stringify(o).replace(/</g, '\\u003c');
  let html = HtmlService.createTemplateFromFile('InformeBusquedaPagina').getRawContent();
  const fuente = HtmlService.createTemplateFromFile('FuenteMarca').getRawContent();
  html = html.replace('%%FUENTE%%', () => fuente).replace('%%DATOS%%', () => js(datos));
  return HtmlService.createHtmlOutput(html).setTitle('Informe de la búsqueda · ' + datos.puesto + ' | Blumia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Números de la búsqueda, con las mismas reglas que el Resumen y el Informe de proceso (metEstado_).
// Edad, género y zona se leen del análisis del CV (columna Requisitos) y se muestran solo agregados.
function ibEstUna_(perfilId, conResumenes) {
  const cfg = getConfig(perfilId);
  const filas = JSON.parse(doGetKanban_(perfilId));
  const banda = Number(String(cfg.banda_salarial || '').replace(/\D/g, '')) || 0;
  const e = { total: 0, scr: 0, comp: 0, emp: 0, res: 0, desc: 0, dsin: 0, dpost: 0, dcomp: 0, activos: 0,
    notasCV: { A: 0, 'B+': 0, B: 0, C: 0, sin: 0 }, motivos: {}, fuentes: {}, formacion: {}, pret: { debajo: 0, linea: 0, encima: 0, total: 0 }, banda: banda,
    edad: { 'Menos de 30': 0, '30 a 39': 0, '40 a 49': 0, '50 o más': 0, 'Sin dato': 0 }, edadScr: { 'Menos de 30': 0, '30 a 39': 0, '40 a 49': 0, '50 o más': 0, 'Sin dato': 0 },
    genero: { Mujeres: 0, Varones: 0, 'Sin dato': 0 }, generoScr: { Mujeres: 0, Varones: 0, 'Sin dato': 0 }, zona: { 'En zona': 0, 'Fuera de zona': 0, 'Sin dato': 0 } };
  const res = [];
  let primero = 0;
  const campo = (txt, k) => { const m = String(txt || '').match(new RegExp(k + ':\\s*([^|]+)', 'i')); return m ? m[1].trim() : ''; };
  filas.forEach(c => {
    const m = metEstado_(c.estado);
    if (!m) return;
    e.total++;
    if (m.scr) e.scr++;
    if (m.comp) e.comp++;
    if (m.emp) e.emp++;
    if (m.res) e.res++;
    if (m.t === 'desc') {
      e.desc++;
      if (m.d === 'cv') e.dsin++; else if (m.d === 'post') e.dpost++; else e.dcomp++;
      const mo = String(c.motivo || '').trim();
      if (mo) e.motivos[mo] = (e.motivos[mo] || 0) + 1;
    }
    if (m.t === 'act') e.activos++;
    const n = String(c.nota || '').trim();
    if (e.notasCV[n] !== undefined) e.notasCV[n]++; else e.notasCV.sin++;
    const f = String(c.fuente || '').trim();
    if (f) e.fuentes[f] = (e.fuentes[f] || 0) + 1;
    const fo = String(c.formacion || '').trim().slice(0, 60);
    if (fo) e.formacion[fo] = (e.formacion[fo] || 0) + 1;
    if (banda && m.scr && c.pretNum > 0) {
      e.pret.total++;
      if (c.pretNum < banda * 0.95) e.pret.debajo++; else if (c.pretNum > banda * 1.05) e.pret.encima++; else e.pret.linea++;
    }
    const ed = parseInt((campo(c.requisitos, 'Edad').match(/\d{2}/) || [''])[0], 10);
    const kE = (ed >= 18 && ed <= 75) ? (ed < 30 ? 'Menos de 30' : ed < 40 ? '30 a 39' : ed < 50 ? '40 a 49' : '50 o más') : 'Sin dato';
    const g = campo(c.requisitos, 'G[eé]nero').toLowerCase();
    const kG = /^(f|muj)/.test(g) ? 'Mujeres' : /^(m|var|hom)/.test(g) ? 'Varones' : 'Sin dato';
    const z = campo(c.requisitos, 'Zona').toLowerCase();
    const kZ = /fuera|no\b/.test(z) ? 'Fuera de zona' : /en zona|^s[ií]|dentro|caba|gba|zona/.test(z) && !/n\/d/.test(z) ? 'En zona' : 'Sin dato';
    e.edad[kE]++; e.genero[kG]++; e.zona[kZ]++;
    if (m.scr) { e.edadScr[kE]++; e.generoScr[kG]++; }
    if (conResumenes && m.scr && res.length < 60 && c.resumen) res.push(String(c.resumen).replace(/\s+/g, ' ').slice(0, 170));
    if (c.recMs && (!primero || c.recMs < primero)) primero = c.recMs;
  });
  e._primero = primero; e.nombre = String(cfg.nombre || perfilId);
  if (conResumenes) e._res = res;
  return e;
}

// Búsquedas que entran en el informe: solo esta o, con alcance «empresa», todas las de la misma empresa (consolidado)
function ibIds_(perfilId, alcance) {
  if (alcance !== 'empresa') return [perfilId];
  const emp = empresaKey_(getConfig(perfilId));
  const ids = [];
  listarPerfiles().forEach(p => { if (p.id === 'operaciones') return; try { if (empresaKey_(getConfig(p.id)) === emp) ids.push(p.id); } catch (e) { } });
  return ids.length ? ids : [perfilId];
}

function ibSumar_(a, b) {
  Object.keys(b).forEach(k => {
    const v = b[k];
    if (typeof v === 'number') a[k] = (a[k] || 0) + v;
    else if (v && typeof v === 'object' && !Array.isArray(v)) { a[k] = a[k] || {}; ibSumar_(a[k], v); }
  });
  return a;
}

function ibEstadisticas_(perfilId, conResumenes, alcance) {
  const ids = ibIds_(perfilId, alcance);
  const raws = ids.map(id => ibEstUna_(id, conResumenes));
  const e = raws.length === 1 ? raws[0] : raws.reduce((acc, r) => ibSumar_(acc, r), {});
  if (raws.length > 1) {
    e.banda = 0;
    e._res = [].concat.apply([], raws.map(r => r._res || [])).slice(0, 60);
    e._primero = Math.min.apply(null, raws.map(r => r._primero).filter(Boolean).concat([Infinity]));
    if (!isFinite(e._primero)) e._primero = 0;
    e.porBusqueda = raws.map(r => ({ nombre: r.nombre, total: r.total, scr: r.scr, comp: r.comp, emp: r.emp, res: r.res, desc: r.desc }));
    delete e.nombre;
  }
  e.dias = e._primero ? Math.max(1, Math.round((Date.now() - e._primero) / 864e5)) : 0;
  delete e._primero;
  const top = (o, n) => Object.keys(o || {}).sort((a, b) => o[b] - o[a]).slice(0, n).map(k => ({ nombre: k, cantidad: o[k] }));
  e.motivosTop = top(e.motivos, 5); e.fuentesTop = top(e.fuentes, 5); e.formacionTop = top(e.formacion, 4);
  delete e.motivos; delete e.fuentes; delete e.formacion;
  if (!conResumenes) delete e._res;
  return e;
}

function ibAnalizar_(perfilId, alcance) {
  const cfg = getConfig(perfilId);
  const est = ibEstadisticas_(perfilId, true, alcance);
  const resumenes = est._res || []; delete est._res;
  const puesto = alcance === 'empresa' ? 'Todas las búsquedas de ' + empresaInforme_(cfg) : String(cfg.nombre || perfilId);
  const sistema = 'Sos consultor/a senior de selección de personal en Blumia Consulting (Argentina). Vas a preparar el informe general de ' + (alcance === 'empresa' ? 'TODAS las búsquedas de un cliente (consolidado: "porBusqueda" trae los números de cada una; incluí hallazgos que comparen las búsquedas entre sí)' : 'una búsqueda') + ' para el CLIENTE. ' +
    'Recibís los números reales de la búsqueda (incluye distribución agregada por rango de edad, género y zona; "edadScr" y "generoScr" son de quienes pasaron screening) y resúmenes anónimos de los candidatos que pasaron screening. Respondé SOLO con un JSON válido, sin texto antes ni después, con esta forma exacta: ' +
    '{"hallazgos":[{"titulo":"máx. 70 caracteres","detalle":"1 o 2 frases con el dato concreto","tipo":"positivo|atencion|dato"}],"industrias":[{"nombre":"sector","cantidad":entero}],"preguntas":[{"id":"q1","texto":"pregunta corta","sugeridas":["respuesta posible 1","respuesta posible 2","respuesta posible 3"]}]}. ' +
    'De 5 a 7 hallazgos, ordenados por importancia para el cliente, apoyados SOLO en los números recibidos (no inventes datos, causas ni porcentajes que no se puedan calcular con ellos). Incluí al menos uno sobre el perfil de las candidaturas (edad, género, formación o zona) en términos descriptivos de la muestra y uno sobre la expectativa salarial si hay datos. ' +
    'Nunca sugieras ni insinúes filtrar, preferir o descartar por edad o género: es solo una descripción de quiénes se postularon. ' +
    'Industrias: de 3 a 8 sectores de origen, deducidos de los resúmenes, con la cantidad aproximada de candidatos de cada uno; si no hay resúmenes, lista vacía. ' +
    'Si la muestra es chica (menos de 15 CVs o menos de 5 en screening) decilo en un hallazgo y no saques conclusiones fuertes. ' +
    'De 4 a 6 preguntas para quien lidera la búsqueda, SOLO sobre lo que los números no dicen y que cambia cómo se escribe el informe: qué quiere destacar el cliente, si la banda salarial es negociable, qué se recomienda hacer con los candidatos en reserva, qué próximos pasos proponer, qué tono usar, si hay algo del mercado que explique lo observado. Cada pregunta con 2 o 3 respuestas sugeridas cortas (máx. 90 caracteres). ' +
    'Nada de nombres de personas. Español rioplatense profesional.';
  const usuario = 'Puesto: ' + puesto + '. Empresa: ' + empresaInforme_(cfg) + '. Números de la búsqueda: ' + JSON.stringify(est) + (resumenes.length ? '. Resúmenes anónimos (uno por candidato): ' + JSON.stringify(resumenes) : '');
  let r;
  try { r = extraerJSON_(procesarConClaude(sistema, usuario, 'analisis')); } catch (err) { r = { hallazgos: [], preguntas: [], industrias: [] }; }
  const t = (v, n) => String(v == null ? '' : v).replace(/[<>]/g, '').trim().slice(0, n);
  const tipos = { positivo: 1, atencion: 1, dato: 1 };
  const hallazgos = (Array.isArray(r.hallazgos) ? r.hallazgos : []).slice(0, 7).map((h, i) => ({ id: 'h' + (i + 1), titulo: t(h && h.titulo, 90), detalle: t(h && h.detalle, 320), tipo: tipos[h && h.tipo] ? h.tipo : 'dato' })).filter(h => h.titulo);
  const industrias = (Array.isArray(r.industrias) ? r.industrias : []).slice(0, 8).map(x => t(x && x.nombre, 50)).filter(Boolean);
  let preguntas = (Array.isArray(r.preguntas) ? r.preguntas : []).slice(0, 6).map((p, i) => ({ id: 'q' + (i + 1), texto: t(p && p.texto, 160), sugeridas: (Array.isArray(p && p.sugeridas) ? p.sugeridas : []).slice(0, 3).map(s => t(s, 110)).filter(Boolean) })).filter(p => p.texto);
  if (!preguntas.length) preguntas = [
    { id: 'q1', texto: '¿Qué querés destacar del proceso ante el cliente?', sugeridas: ['La velocidad y el volumen de candidatos', 'La calidad de los finalistas', 'El ajuste a la banda salarial'] },
    { id: 'q2', texto: '¿Qué recomendás hacer con los candidatos en reserva?', sugeridas: ['Mantenerlos para una próxima búsqueda', 'Retomarlos si los finalistas no cierran'] },
    { id: 'q3', texto: '¿Qué tono preferís para el informe?', sugeridas: ['Profesional y cercano', 'Formal'] }];
  return { est: est, hallazgos: hallazgos, industrias: industrias, preguntas: preguntas };
}

// Escribe el informe con los hallazgos que se dejaron y las respuestas, lo guarda editable en Drive y devuelve el link
function ibRedactar_(perfilId, hallazgos, respuestas, extra) {
  const cfg = getConfig(perfilId);
  const alcance = extra && extra.alcance === 'empresa' ? 'empresa' : '';
  const est = ibEstadisticas_(perfilId, false, alcance);
  const lim = (v, n) => String(v == null ? '' : v).replace(/[<>]/g, '').trim().slice(0, n);
  const hs = (Array.isArray(hallazgos) ? hallazgos : []).slice(0, 8).map(h => ({ titulo: lim(h && h.titulo, 90), detalle: lim(h && h.detalle, 400) })).filter(h => h.titulo);
  const rs = (Array.isArray(respuestas) ? respuestas : []).slice(0, 8).map(r => ({ pregunta: lim(r && r.pregunta, 160), respuesta: lim(r && r.respuesta, 500) })).filter(r => r.respuesta);
  const ex = extra || {};
  const industrias = (Array.isArray(ex.industrias) ? ex.industrias : []).slice(0, 10).map(x => lim(x, 50)).filter(Boolean);
  const inc = Object.assign({ salarial: true, industrias: true, edad: true, genero: true, zona: true, fuentes: true }, ex.incluir || {});
  if (!hs.length) throw new Error('Dejá al menos un hallazgo para armar el informe.');
  const empresa = empresaInforme_(cfg), puesto = alcance ? 'Todas las búsquedas' : String(cfg.nombre || perfilId);
  const sistema = 'Sos consultor/a senior de selección de personal en Blumia Consulting (Argentina). Redactás el informe general de ' + (alcance ? 'todas las búsquedas (consolidado)' : 'una búsqueda') + ' dirigido al CLIENTE (' + empresa + '). ' +
    'Usá SOLO los números y los hallazgos recibidos y las respuestas de quien lidera la búsqueda; no inventes datos, porcentajes ni causas. Respetá el tono pedido en las respuestas (si no hay, profesional y cercano). ' +
    'No nombres candidatos, no menciones situación familiar ni criterios internos de evaluación; edad y género solo como descripción agregada de quiénes se postularon, nunca como criterio ni recomendación. Hablá de la búsqueda en tercera persona ("la búsqueda", "se recibieron"). Español rioplatense profesional. ' +
    'Respondé SOLO con un JSON válido con esta forma exacta: {"resumen":"párrafo de 3 a 5 frases","hallazgos":[{"titulo":"","texto":"2 o 3 frases"}],"recomendaciones":["acción concreta"],"proximos_pasos":["paso concreto"],"cierre":"1 o 2 frases"}. ' +
    'Hallazgos: reescribí los recibidos, en el mismo orden. 2 a 4 recomendaciones y 2 a 4 próximos pasos, coherentes con las respuestas.';
  const usuario = 'Puesto: ' + puesto + '. Números: ' + JSON.stringify(est) + '. Hallazgos a incluir: ' + JSON.stringify(hs) + '. Respuestas de quien lidera la búsqueda: ' + JSON.stringify(rs);
  const r = extraerJSON_(procesarConClaude(sistema, usuario, 'informe'));
  const inf = {
    titulo: IB_TITULO_,
    resumen: lim(r.resumen, 900),
    hallazgos: (Array.isArray(r.hallazgos) ? r.hallazgos : []).slice(0, 8).map(h => ({ titulo: lim(h && h.titulo, 100), texto: lim(h && h.texto, 600) })).filter(h => h.titulo),
    recomendaciones: (Array.isArray(r.recomendaciones) ? r.recomendaciones : []).slice(0, 5).map(x => lim(x, 300)).filter(Boolean),
    proximos: (Array.isArray(r.proximos_pasos) ? r.proximos_pasos : []).slice(0, 5).map(x => lim(x, 300)).filter(Boolean),
    cierre: lim(r.cierre, 400), industrias: industrias, incluir: inc
  };
  const html = ibArmarHTML_(cfg, puesto, empresa, est, inf);
  const it = DriveApp.getFoldersByName('Informes Clientes');
  const carpeta = it.hasNext() ? it.next() : DriveApp.createFolder('Informes Clientes');
  const nombre = alcance ? 'Informe_Empresa_' + String(empresaKey_(cfg)).replace(/[^\w-]+/g, '_') + '.html' : 'Informe_Busqueda_' + perfilId + '.html';
  const exs = carpeta.getFilesByName(nombre);
  const arch = exs.hasNext() ? exs.next() : carpeta.createFile(nombre, '<html></html>', MimeType.HTML);
  arch.setContent(conFuentePdf_(html).replace(/__FILEID__/g, arch.getId()));
  try { registrarHistorial_(cfg.hoja_nombre || perfilId, '(búsqueda)', 'Informe', (alcance ? 'Informe consolidado de la empresa' : 'Informe general para el cliente') + ' armado', QUIEN_); } catch (e) { }
  return { url: URL_WEBAPP + '?tipo=verinforme&id=' + arch.getId(), informe: inf };
}

// Mismo formato que el Informe de proceso (cssInformeProceso_) más las secciones de perfil de las candidaturas y recomendaciones
function ibArmarHTML_(cfg, puesto, empresa, e, inf) {
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const ce = (s, tag, cls) => '<' + tag + (cls ? ' class="' + cls + '"' : '') + ' contenteditable="true">' + esc(s) + '</' + tag + '>';
  const pct = (a, b) => b ? Math.round(a / b * 100) + '%' : '—';
  const inc = inf.incluir || {};
  const LOGO = 'https://lh3.googleusercontent.com/d/1457Wfl9zeY1V3sAGtY1IT2BmF2lqpoy8';
  let n = 0;
  const sec = (tit, em, cuerpo) => { n++; return '<div class="section"><div class="sec-num">0' + n + ' · ' + tit + '</div>' + cuerpo + '</div>'; };
  const stage = (num, lb, p, det, cls, dark) => '<div class="fnl-stage"><div class="fnl-shape ' + cls + '"><div class="fnl-num' + (dark ? ' dark' : '') + '" contenteditable="true">' + num + '</div><div class="fnl-lb' + (dark ? ' dark' : '') + '">' + lb + '</div></div><div class="fnl-side"><div class="fnl-pct" contenteditable="true">' + p + '</div><div class="fnl-det">' + det + '</div></div></div>';
  const barras = (obj, total) => '<div class="bars">' + Object.keys(obj).filter(k => k !== 'Sin dato' || obj[k]).map(k => '<div class="bar"><span class="bl">' + esc(k) + '</span><span class="bt"><i style="width:' + (total ? Math.round(obj[k] / total * 100) : 0) + '%"></i></span><span class="bv" contenteditable="true">' + pct(obj[k], total) + '</span></div>').join('') + '</div>';
  const conDato = o => Object.keys(o).reduce((a, k) => a + (k === 'Sin dato' ? 0 : o[k]), 0);
  let cuerpo = '';
  cuerpo += sec('Embudo', '', '<h3 class="sec-title">Del volumen a la <em>precisión</em></h3><div class="funnel">' +
    stage(e.total, 'CVS RECIBIDOS', '100%', 'base del proceso', 'f1', false) + stage(e.scr, 'SCREENINGS', pct(e.scr, e.total), 'entrevistas telefónicas', 'f2', false) +
    stage(e.comp, 'COMPETENCIAS', pct(e.comp, e.total), 'evaluación profunda', 'f3', false) + stage(e.emp, 'AVANZAN A EMPRESA', pct(e.emp, e.total), 'finalistas presentados', 'f4', true) +
    '</div>' + (e.res ? '<div class="fnl-extra" contenteditable="true">+ ' + e.res + ' candidatos en reserva activa</div>' : '') + (e.dias ? '<div class="fnl-extra">Tiempo desde el primer CV recibido: ' + e.dias + ' días</div>' : ''));
  if (e.porBusqueda && e.porBusqueda.length > 1) cuerpo += sec('Por búsqueda', '', '<h3 class="sec-title">Cada <em>búsqueda</em></h3><div class="rk-tw"><table class="pb"><thead><tr><th>Búsqueda</th><th>CVs</th><th>Screening</th><th>Competencias</th><th>Empresa</th><th>Reserva</th><th>Descartados</th></tr></thead><tbody>' +
    e.porBusqueda.map(b => '<tr><td contenteditable="true">' + esc(b.nombre) + '</td><td>' + b.total + '</td><td>' + b.scr + '</td><td>' + b.comp + '</td><td>' + b.emp + '</td><td>' + b.res + '</td><td>' + b.desc + '</td></tr>').join('') + '</tbody></table></div>');
  if (inc.salarial !== false && (e.banda || e.porBusqueda) && e.pret.total >= 5) cuerpo += sec('Realidad salarial', '', '<h3 class="sec-title">Lo que <em>pretende</em> el mercado</h3><div class="salary-grid">' +
    '<div class="sal-card dark"><div class="sal-pct dark" contenteditable="true">' + pct(e.pret.debajo, e.pret.total) + '</div><div class="sal-l">Por debajo</div><div class="sal-d">de la banda de referencia</div></div>' +
    '<div class="sal-card green"><div class="sal-pct green" contenteditable="true">' + pct(e.pret.linea, e.pret.total) + '</div><div class="sal-l">En línea</div><div class="sal-d">con la banda</div></div>' +
    '<div class="sal-card amber"><div class="sal-pct amber" contenteditable="true">' + pct(e.pret.encima, e.pret.total) + '</div><div class="sal-l">Por encima</div><div class="sal-d">de la banda</div></div></div><div class="fnl-extra">Sobre ' + e.pret.total + ' candidatos que informaron pretensión · ' + (e.banda ? 'banda de referencia $' + Number(e.banda).toLocaleString('es-AR') : 'respecto de la banda de cada búsqueda') + '</div>');
  if (inc.industrias !== false && inf.industrias && inf.industrias.length) cuerpo += sec('Industrias de origen', '', '<h3 class="sec-title">De dónde <em>vienen</em></h3><div class="ind-grid">' + inf.industrias.map(x => '<div class="ind-chip" contenteditable="true">' + esc(x) + '</div>').join('') + '</div>');
  const tieneEdad = inc.edad !== false && conDato(e.edadScr) >= 5, tieneGen = inc.genero !== false && conDato(e.generoScr) >= 5, tieneZona = inc.zona !== false && (e.zona['En zona'] + e.zona['Fuera de zona']) >= 5;
  if (tieneEdad || tieneGen || tieneZona) {
    let c = '<h3 class="sec-title">Quiénes <em>se postularon</em></h3><div class="perfil-grid">';
    if (tieneEdad) c += '<div class="pf"><div class="pf-t">Rango de edad</div>' + barras(e.edad, e.total) + '</div>';
    if (tieneGen) c += '<div class="pf"><div class="pf-t">Distribución por género</div>' + barras(e.genero, e.total) + '</div>';
    if (tieneZona) c += '<div class="pf"><div class="pf-t">Ubicación</div>' + barras(e.zona, e.total) + '</div>';
    cuerpo += sec('Perfil de las candidaturas', '', c + '</div><div class="fnl-extra">Sobre el total de CVs recibidos. Es una descripción de quiénes se postularon, no un criterio de selección.</div>');
  }
  cuerpo += sec('Descartes', '', '<h3 class="sec-title">Filtro por <em>etapa</em></h3><div class="stats-grid three">' +
    '<div class="stat-card acc3"><div class="stat-num gray" contenteditable="true">' + e.dsin + '</div><div class="stat-lbl">En análisis de CV</div></div>' +
    '<div class="stat-card acc3"><div class="stat-num gray" contenteditable="true">' + e.dpost + '</div><div class="stat-lbl">Post screening</div></div>' +
    '<div class="stat-card acc3"><div class="stat-num gray" contenteditable="true">' + e.dcomp + '</div><div class="stat-lbl">En competencias</div></div></div>' +
    (e.motivosTop.length ? '<div class="ind-grid" style="margin-top:14px">' + e.motivosTop.map(m => '<div class="ind-chip">' + esc(m.nombre) + ' · <b>' + m.cantidad + '</b></div>').join('') + '</div>' : ''));
  if (inc.fuentes !== false && e.fuentesTop.length) cuerpo += sec('Fuentes', '', '<h3 class="sec-title">Canales de <em>atracción</em></h3><div class="fuentes-grid">' + e.fuentesTop.map(f => '<div class="fuente-item">✓ ' + esc(f.nombre) + ' · ' + f.cantidad + '</div>').join('') + '</div>');
  cuerpo += sec('Hallazgos', '', '<h3 class="sec-title">Lo que <em>vimos</em></h3>' + inf.hallazgos.map(h => '<div class="insight"><span class="insight-tag" contenteditable="true">' + esc(h.titulo) + '</span><div class="insight-text" contenteditable="true">' + esc(h.texto) + '</div></div>').join(''));
  const lista = a => '<ul class="lista">' + a.map(x => '<li contenteditable="true">' + esc(x) + '</li>').join('') + '</ul>';
  if (inf.recomendaciones.length) cuerpo += sec('Recomendaciones', '', '<h3 class="sec-title">Qué <em>sugerimos</em></h3>' + lista(inf.recomendaciones));
  if (inf.proximos.length) cuerpo += sec('Próximos pasos', '', '<h3 class="sec-title">Cómo <em>seguimos</em></h3>' + lista(inf.proximos));
  const hstats = [['CVs Recibidos', e.total], ['Screenings', e.scr], ['Competencias', e.comp], ['Avanzan a empresa', e.emp]].map(x => '<div class="hs-item"><div class="hs-num" contenteditable="true">' + x[1] + '</div><div class="hs-lbl">' + x[0] + '</div></div>').join('');
  const cierre = '<section class="resumen"><div class="res-eye">Cierre</div><h2 class="res-tit" contenteditable="true">' + esc(inf.cierre || 'La búsqueda avanza y el pipeline está activo') + '</h2><div class="res-grid">' +
    [[e.total, 'CVs recibidos', ''], [e.emp, 'Avanzan a empresa', ' amber'], [e.res, 'En reserva', ''], [e.desc, 'Descartados', '']].map(x => '<div class="res-card"><div class="res-num' + x[2] + '" contenteditable="true">' + x[0] + '</div><div class="res-lbl">' + x[1] + '</div></div>').join('') + '</div></section>';
  const extra = 'body{font-size:16px}.hero{padding-top:44px}.h-title{font-size:48px}.hs-lbl{font-size:12px}.bsq-sub{font-size:15px}.sec-title{font-size:26px}.sec-num{font-size:11px}.insight-text{font-size:16px;line-height:1.7}.insight-tag{font-size:14px!important}.lista li{font-size:16px;line-height:1.6}.fuente-item,.ind-chip{font-size:15px}.fnl-det,.sal-d,.stat-lbl,.fnl-extra{font-size:13px}.res-lbl{font-size:12px}.res-lect{font-size:16px}' +
    '#barra-guardar{display:flex;gap:8px;align-items:center;flex-wrap:wrap;justify-content:flex-end}#edPdf{background:#0F3F27!important;color:#fff!important}#barra-guardar button{padding:10px 16px}#edBtn.on{background:#0F3F27;color:#fff}.ed-x,.ed-add{display:none}body.ed .ed-h{position:relative;outline:1px dashed rgba(15,63,39,.35);outline-offset:3px}body.ed .ed-x{display:flex;position:absolute;top:-10px;right:-10px;width:26px;height:26px;border-radius:50%;background:#b3261e;color:#fff;border:2px solid #fff;font:800 13px/1 inherit;font-family:inherit;align-items:center;justify-content:center;cursor:pointer;z-index:20;padding:0;box-shadow:0 1px 4px rgba(0,0,0,.3)}body.ed .ed-add{display:inline-flex;margin:8px 8px 0 0;background:#fff;border:1.5px dashed #0F3F27;color:#0F3F27;border-radius:8px;padding:7px 14px;font:700 13px inherit;font-family:inherit;cursor:pointer}.ed-p{font-size:16px;line-height:1.7;max-width:820px;margin-top:12px}' +
    '@media print{.ed-x,.ed-add,#barra-guardar,#pdfPanel{display:none!important}.ed-h{outline:none!important}}' +
    '.perfil-grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(240px,1fr));gap:14px}.pf{background:#fff;border:1px solid var(--line);border-radius:4px;padding:18px}.pf-t{font-size:11px;font-weight:800;letter-spacing:.08em;text-transform:uppercase;margin-bottom:12px}.bar{display:grid;grid-template-columns:96px 1fr 44px;gap:8px;align-items:center;margin-bottom:8px;font-size:12px;font-weight:600}.bar .bt{background:var(--bg2);height:10px;border-radius:5px;overflow:hidden}.bar .bt i{display:block;height:100%;background:var(--green)}.bar .bv{text-align:right;color:var(--mid)}.busqueda>.insight{margin-bottom:34px}.rk-tw{overflow-x:auto}table.pb{width:100%;border-collapse:collapse;font-size:15px}table.pb th{text-align:left;font-size:11px;letter-spacing:.1em;text-transform:uppercase;color:var(--mid);padding:8px;border-bottom:2px solid var(--dark)}table.pb td{padding:10px 8px;border-bottom:1px solid var(--line)}.lista{padding-left:20px;max-width:820px}.lista li{margin-bottom:8px}.insight-tag{text-transform:none!important;letter-spacing:0!important;font-size:13px!important}@media(max-width:700px){.h-stats,.res-grid{grid-template-columns:1fr 1fr}.salary-grid,.stats-grid.three{grid-template-columns:1fr}.hero,.busqueda,.resumen{padding-left:20px;padding-right:20px}.h-title{font-size:30px}}';
  const script = '<script>(' + ibEditorJS_.toString() + ')();</scr' + 'ipt>';
  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Informe de la búsqueda · ' + esc(puesto) + '</title><style>' + cssInformeProceso_() + extra + '</style></head><body>' +
    '<div id="barra-guardar" contenteditable="false"><button type="button" id="edBtn" title="Quitar o agregar textos y puntos">✏️ Editar</button><button type="button" id="edUn" style="display:none">↶ Deshacer</button><button type="button" id="edPdf">📑 PDF para el cliente</button><button type="button" id="edSave">GUARDAR CAMBIOS</button></div>' +
    '<section class="hero">' + heroLogoCli_(cfg, false) + '<img class="hero-logo" src="' + LOGO + '"><div class="eyebrow">Atracción y Elección · Informe de búsqueda</div><div class="cliente-nombre" contenteditable="true">' + esc(empresa) + '</div>' +
    '<h1 class="h-title" contenteditable="true">El proceso está <em>en marcha</em>, el mercado <em>respondió</em>.</h1><div class="h-stats">' + hstats + '</div></section>' +
    '<section class="busqueda"><div class="bsq-head"><div class="bsq-eye">' + (e.porBusqueda ? 'Consolidado · ' + e.porBusqueda.length + ' búsquedas' : 'Búsqueda') + '</div><h2 class="bsq-title" contenteditable="true">' + esc(puesto) + '</h2>' + (e.banda ? '<div class="bsq-sub" contenteditable="true">Banda salarial de referencia: $' + Number(e.banda).toLocaleString('es-AR') + '</div>' : '') + '</div>' +
    '<div class="insight"><span class="insight-tag">Resumen</span><div class="insight-text" contenteditable="true">' + esc(inf.resumen) + '</div></div>' + cuerpo + '</section>' + cierre +
    '<footer>Blumia Consulting · Documento preparado para ' + esc(empresa) + '</footer>' + script + '</body></html>';
}

// Estilo del informe de proceso (lo usan generarHTMLProceso y el informe de la búsqueda)
function cssInformeProceso_() {
  return ':root{--bg:#F7F5F1;--bg2:#EEECEA;--white:#FFF;--dark:#111;--text:#1A1A1A;--mid:#555;--light:#999;--line:#DDDBD7;--green:#1A6645;--green2:#238055;--glight:#E6F2EC;--amber:#F5C842;--font:\'PP Pangram Sans\',\'Plus Jakarta Sans\',sans-serif;}*{box-sizing:border-box;margin:0;padding:0;-webkit-print-color-adjust:exact;print-color-adjust:exact;}body{font-family:var(--font);color:var(--text);font-size:14px;line-height:1.5;background:var(--bg);}[contenteditable]:focus{outline:2px solid var(--green);outline-offset:2px;background:rgba(245,200,66,.12);}#barra-guardar{position:fixed;top:14px;right:14px;z-index:9999;}#barra-guardar button{background:var(--amber);color:var(--dark);border:none;padding:11px 20px;border-radius:3px;font-weight:800;font-size:11px;letter-spacing:.14em;text-transform:uppercase;cursor:pointer;box-shadow:0 3px 10px rgba(0,0,0,.25);font-family:var(--font);}.hero{padding:52px 40px 40px;background:linear-gradient(180deg,#FCFAF6 0%,var(--bg) 100%);border-bottom:1px solid var(--line);}.hero-logo{height:56px;margin-bottom:20px;display:block;}.eyebrow{display:block;font-size:11px;font-weight:800;letter-spacing:.2em;text-transform:uppercase;color:var(--green);background:var(--glight);padding:6px 14px;border-radius:2px;margin-bottom:14px;width:fit-content;}.cliente-nombre{font-size:15px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:var(--mid);margin-bottom:16px;}.h-title{font-size:46px;line-height:1.06;letter-spacing:-.03em;font-weight:800;margin-bottom:10px;max-width:900px;}.h-title em{color:var(--green);font-style:normal;}.h-stats{display:grid;grid-template-columns:repeat(4,1fr);gap:24px;margin-top:34px;padding-top:30px;border-top:1px solid var(--line);}.hs-num{font-size:50px;font-weight:800;color:var(--green);letter-spacing:-.04em;line-height:1;}.hs-lbl{font-size:11px;font-weight:700;margin-top:8px;text-transform:uppercase;letter-spacing:.08em;}.busqueda{padding:44px 40px;border-bottom:1px solid var(--line);background:#fff;}.bsq-head{margin-bottom:30px;}.bsq-eye{font-size:11px;font-weight:800;letter-spacing:.2em;text-transform:uppercase;color:var(--green);margin-bottom:8px;}.bsq-title{font-size:32px;font-weight:800;letter-spacing:-.025em;line-height:1.1;}.bsq-sub{font-size:14px;color:var(--mid);margin-top:8px;}.section{margin-bottom:40px;}.sec-num{font-size:10px;font-weight:800;letter-spacing:.2em;text-transform:uppercase;color:var(--mid);margin-bottom:8px;}.sec-title{font-size:22px;font-weight:800;letter-spacing:-.02em;margin-bottom:20px;}.sec-title em{color:var(--green);font-style:normal;}.funnel{max-width:640px;margin:0 auto;}.fnl-stage{display:grid;grid-template-columns:1fr 170px;gap:20px;align-items:center;margin-bottom:6px;break-inside:avoid;}.fnl-shape{padding:12px 20px;text-align:center;color:#fff;margin:0 auto;width:100%;}.fnl-shape.f1{background:linear-gradient(180deg,#1A6645,#238055);clip-path:polygon(0 0,100% 0,93% 100%,7% 100%);}.fnl-shape.f2{background:linear-gradient(180deg,#238055,#3FA670);clip-path:polygon(7% 0,93% 0,86% 100%,14% 100%);width:86%;}.fnl-shape.f3{background:#3FA670;clip-path:polygon(14% 0,86% 0,79% 100%,21% 100%);width:70%;}.fnl-shape.f4{background:linear-gradient(180deg,#F5C842,#E0A82E);clip-path:polygon(21% 0,79% 0,72% 100%,28% 100%);width:54%;}.fnl-num{font-size:26px;font-weight:800;letter-spacing:-1px;line-height:1;}.fnl-num.dark{color:#1A1A1A;}.fnl-lb{font-size:9px;font-weight:700;letter-spacing:1.6px;margin-top:3px;opacity:.9;}.fnl-lb.dark{color:#1A1A1A;opacity:1;}.fnl-side{padding-left:12px;border-left:2px solid var(--line);}.fnl-pct{font-size:15px;font-weight:800;color:var(--mid);}.fnl-det{font-size:11px;color:var(--light);margin-top:2px;}.fnl-extra{font-size:12px;color:var(--mid);margin-top:12px;text-align:center;font-weight:600;}.salary-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;}.sal-card{background:#fff;border:1px solid var(--line);padding:24px;border-radius:4px;}.sal-card.green{border-top:3px solid var(--green);}.sal-card.amber{border-top:3px solid var(--amber);}.sal-card.dark{border-top:3px solid var(--dark);}.sal-pct{font-size:40px;font-weight:800;letter-spacing:-.03em;line-height:1;}.sal-pct.green{color:var(--green);}.sal-pct.amber{color:#B8902F;}.sal-pct.dark{color:var(--dark);}.sal-l{font-size:13px;font-weight:700;margin-top:8px;}.sal-d{font-size:12px;color:var(--mid);margin-top:4px;}.ind-grid{display:flex;flex-wrap:wrap;gap:8px;}.ind-chip{background:#FCFAF6;border:1px solid var(--line);border-left:3px solid var(--amber);padding:8px 14px;font-size:12px;border-radius:2px;font-weight:600;}.stats-grid{display:grid;gap:14px;}.stats-grid.three{grid-template-columns:repeat(3,1fr);}.stat-card{background:#fff;border:1px solid var(--line);padding:20px;border-radius:4px;}.stat-card.acc3{border-top:3px solid var(--mid);}.stat-num{font-size:34px;font-weight:800;letter-spacing:-.03em;line-height:1;}.stat-num.gray{color:var(--mid);}.stat-lbl{font-size:11px;font-weight:700;margin-top:6px;text-transform:uppercase;letter-spacing:.06em;}.insight{padding:20px 24px;background:linear-gradient(180deg,#FCFAF6 0%,#F2EFEA 100%);border-left:4px solid var(--green);border-radius:2px;margin-top:16px;break-inside:avoid;}.insight-tag{font-size:10px;font-weight:800;letter-spacing:.2em;text-transform:uppercase;color:var(--green);margin-bottom:8px;display:block;}.insight-text{font-size:13px;line-height:1.6;}.fuentes-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;}.fuente-item{font-size:13px;color:#333;font-weight:600;}.resumen{padding:52px 40px;background:var(--green);color:#fff;}.res-eye{font-size:11px;font-weight:800;letter-spacing:.2em;text-transform:uppercase;color:var(--amber);margin-bottom:14px;}.res-tit{font-size:34px;font-weight:800;letter-spacing:-.03em;line-height:1.1;margin-bottom:24px;}.res-tit em{color:var(--amber);font-style:normal;}.res-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:18px;}.res-card{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);border-radius:4px;padding:22px;}.res-num{font-size:44px;font-weight:800;letter-spacing:-.03em;line-height:1;}.res-num.amber{color:var(--amber);}.res-lbl{font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;margin-top:8px;}.res-lect{font-size:14px;line-height:1.7;color:rgba(255,255,255,.92);margin-top:24px;}footer{padding:28px;text-align:center;font-size:11px;color:var(--mid);background:#fff;border-top:1px solid var(--line);}@media print{@page{size:A4;margin:12mm 10mm;}#barra-guardar{display:none!important;}.busqueda,.resumen{page-break-before:always;}.hero{page-break-after:avoid;}.section,.insight,.fnl-stage{break-inside:avoid;}}';
}

/* PDF del informe para el cliente (pedido de Ariel 29/10: «fundamental»): convierte el informe guardado (con las ediciones) con PDFShift,
   como el «PDF para el cliente», y lo deja en la carpeta «Informes Clientes» con link para ver. Opcionalmente lo manda por mail con el PDF adjunto. */
function ibArchivoInforme_(id, extension) {
  const f = DriveApp.getFileById(String(id));
  let ok = false;
  const ps = f.getParents();
  while (ps.hasNext()) if (ps.next().getName() === 'Informes Clientes') ok = true;
  if (!ok || !/^(Informe|Ranking)_/.test(f.getName()) || (extension && f.getName().slice(-extension.length) !== extension)) throw new Error('Ese archivo no es un informe para el cliente.');
  return f;
}

function informePDF_(fileId, quien) {
  QUIEN_ = quien || '';
  const f = ibArchivoInforme_(fileId, '.html');
  const html = f.getBlob().getDataAsString().replace(/<script[\s\S]*?<\/script>/gi, '');
  const resp = UrlFetchApp.fetch('https://api.pdfshift.io/v3/convert/pdf', {
    method: 'post', headers: { 'Authorization': 'Basic ' + Utilities.base64Encode('api:' + pdfshiftKey_()) }, contentType: 'application/json',
    payload: JSON.stringify({ source: html, format: 'A4', use_print: true }), muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) throw new Error('No se pudo armar el PDF (' + resp.getResponseCode() + '): ' + resp.getContentText().substring(0, 200));
  const carpeta = f.getParents().next();
  const nombre = f.getName().replace(/\.html$/, '.pdf');
  const viejos = carpeta.getFilesByName(nombre);
  while (viejos.hasNext()) viejos.next().setTrashed(true);
  const pdf = carpeta.createFile(resp.getBlob().setName(nombre));
  pdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  try { registrarHistorial_('(informes)', f.getName(), 'Informe PDF', 'PDF generado', quien); } catch (e) { }
  return { url: pdf.getUrl(), id: pdf.getId(), nombre: nombre };
}

function informeEnviar_(pdfId, para, asunto, mensaje, quien) {
  QUIEN_ = quien || '';
  para = String(para || '').trim();
  if (!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(para)) throw new Error('Escribí un correo válido (uno solo).');
  const f = ibArchivoInforme_(pdfId, '.pdf');
  const lim = (v, n) => String(v == null ? '' : v).replace(/[\r\n]+/g, ' ').trim().slice(0, n);
  const cuerpo = String(mensaje || '').slice(0, 3000) || 'Te comparto el informe de la búsqueda. Cualquier duda, quedo a disposición.';
  const env = enviarMailBlumia_(para, lim(asunto, 150) || 'Informe de la búsqueda', cuerpo + '\n\nBlumia Consulting', { attachments: [f.getBlob()] });
  try { registrarHistorial_('(informes)', f.getName(), 'Informe enviado', 'Enviado a ' + para + ' desde ' + env.desde, quien); } catch (e) { }
  return { ok: true, desde: env.desde, alias: env.alias };
}

// Modo edición del informe armado (corre dentro del HTML del informe, no en Apps Script): ✕ para sacar cualquier bloque, punto o dato,
// ＋ para agregar puntos o textos, ↶ para deshacer; los controles no salen al guardar ni al imprimir.
function ibEditorJS_() {
  var on = false, pila = [];
  var SEL = '.section,.insight,.ind-chip,.stat-card,.sal-card,.fnl-stage,.pf,.fuente-item,.res-card,.hs-item,.fnl-extra,li,.bar,.ed-p';
  function ctl(t, c) { var b = document.createElement('button'); b.type = 'button'; b.className = c; b.textContent = t; b.setAttribute('contenteditable', 'false'); return b; }
  function estado() { document.getElementById('edUn').style.display = pila.length ? 'inline-flex' : 'none'; }
  function marcar() {
    [].forEach.call(document.querySelectorAll(SEL), function (e) {
      if (e.getAttribute('data-ed')) return;
      e.setAttribute('data-ed', '1'); e.classList.add('ed-h');
      var x = ctl('✕', 'ed-x'); x.title = 'Eliminar';
      x.onclick = function (ev) { ev.stopPropagation(); ev.preventDefault(); var p = e.parentNode, sg = e.nextSibling; pila.push({ e: e, p: p, s: sg }); p.removeChild(e); estado(); };
      e.appendChild(x);
    });
    [].forEach.call(document.querySelectorAll('ul.lista'), function (ul) {
      if (ul.nextSibling && ul.nextSibling.className === 'ed-add') return;
      var a = ctl('＋ Agregar punto', 'ed-add');
      a.onclick = function () { var li = document.createElement('li'); li.setAttribute('contenteditable', 'true'); li.textContent = 'Nuevo punto'; ul.appendChild(li); marcar(); li.focus(); };
      ul.parentNode.insertBefore(a, ul.nextSibling);
    });
    [].forEach.call(document.querySelectorAll('.section'), function (sec) {
      if (sec.querySelector(':scope>.ed-add.tx')) return;
      var a = ctl('＋ Agregar texto', 'ed-add tx');
      a.onclick = function () { var p = document.createElement('div'); p.className = 'ed-p'; p.setAttribute('contenteditable', 'true'); p.textContent = 'Escribí acá…'; sec.insertBefore(p, a); marcar(); p.focus(); };
      sec.appendChild(a);
    });
  }
  function limpiar(raiz) {
    [].forEach.call(raiz.querySelectorAll('.ed-x,.ed-add'), function (n) { n.parentNode.removeChild(n); });
    [].forEach.call(raiz.querySelectorAll('[data-ed]'), function (n) { n.removeAttribute('data-ed'); n.classList.remove('ed-h'); });
  }
  function panel(h) { var p = document.getElementById('pdfPanel'); if (!p) { p = document.createElement('div'); p.id = 'pdfPanel'; p.setAttribute('contenteditable', 'false'); p.style.cssText = 'position:fixed;top:64px;right:14px;z-index:9998;background:#fff;border:2px solid #0F3F27;border-radius:12px;padding:16px;width:340px;max-width:calc(100vw - 28px);box-shadow:0 8px 30px rgba(0,0,0,.3);font:14px/1.45 inherit;font-family:inherit'; document.body.appendChild(p); } p.innerHTML = h; return p; }
  function esc(t) { return String(t == null ? '' : t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/"/g, '&quot;'); }
  function encender() { on = true; document.body.classList.add('ed'); document.getElementById('edBtn').classList.add('on'); document.getElementById('edBtn').textContent = '✔ Listo'; marcar(); }
  function apagar() { on = false; document.body.classList.remove('ed'); document.getElementById('edBtn').classList.remove('on'); document.getElementById('edBtn').textContent = '✏️ Editar'; limpiar(document); }
  document.getElementById('edBtn').onclick = function () { if (on) apagar(); else encender(); };
  document.getElementById('edUn').onclick = function () { var u = pila.pop(); if (u) u.p.insertBefore(u.e, u.s && u.s.parentNode === u.p ? u.s : null); estado(); };
  function guardar(cb, falla) {
    var estaba = on; if (on) apagar();
    var clon = document.documentElement.cloneNode(true);
    var pp = clon.querySelector('#pdfPanel'); if (pp) pp.parentNode.removeChild(pp);
    var b = clon.querySelector('#edSave'); if (b) { b.textContent = 'GUARDAR CAMBIOS'; b.disabled = false; }
    var eb = clon.querySelector('#edBtn'); if (eb) { eb.classList.remove('on'); eb.textContent = '✏️ Editar'; }
    var eu = clon.querySelector('#edUn'); if (eu) eu.style.display = 'none';
    var ep = clon.querySelector('#edPdf'); if (ep) { ep.textContent = '📑 PDF para el cliente'; ep.disabled = false; }
    google.script.run.withSuccessHandler(function (r) { if (estaba) encender(); if (r && r.ok) cb(); else falla('No se pudo guardar el informe.'); })
      .withFailureHandler(function (e) { if (estaba) encender(); falla((e && e.message) || 'No se pudo guardar el informe.'); }).guardarInformeEditado('__FILEID__', '<!DOCTYPE html>' + clon.outerHTML);
  }
  document.getElementById('edPdf').onclick = function () {
    var bt = this, t = bt.textContent; bt.disabled = true; bt.textContent = 'Armando el PDF…';
    var fin = function (m) { bt.disabled = false; bt.textContent = t; panel('<b>No se pudo armar el PDF</b><p style="margin:8px 0">' + esc(m) + '</p><button type="button" id="pdfX" style="height:36px;padding:0 14px;border-radius:8px;border:1.5px solid #0F3F27;background:#fff;font-weight:700;cursor:pointer">Cerrar</button>'); document.getElementById('pdfX').onclick = function () { document.getElementById('pdfPanel').remove(); }; };
    guardar(function () {
      google.script.run.withSuccessHandler(function (r) {
        bt.disabled = false; bt.textContent = t;
        var asunto = (document.title || 'Informe') + '';
        var p = panel('<b style="color:#0F3F27">✔ PDF listo</b><div style="margin:10px 0;display:flex;gap:8px;flex-wrap:wrap"><a href="' + esc(r.url) + '" target="_blank" style="display:inline-flex;align-items:center;height:36px;padding:0 14px;border-radius:8px;background:#0F3F27;color:#fff;font-weight:700;text-decoration:none">Abrir PDF</a><button type="button" id="pdfCp" style="height:36px;padding:0 14px;border-radius:8px;border:1.5px solid #0F3F27;background:#fff;font-weight:700;cursor:pointer">Copiar link</button></div>' +
          '<div style="font-weight:700;margin-top:6px">Enviarlo por mail al cliente</div><input id="pdfMail" type="email" placeholder="correo del cliente" style="width:100%;height:38px;margin:6px 0;border:1.5px solid #D5E0CE;border-radius:8px;padding:0 10px;font:inherit"><textarea id="pdfMsg" rows="3" style="width:100%;border:1.5px solid #D5E0CE;border-radius:8px;padding:8px 10px;font:inherit">Te comparto el informe de la búsqueda. Cualquier duda, quedo a disposición.</textarea><div id="pdfSt" style="font-size:12px;color:#666;margin:6px 0"></div><div style="display:flex;gap:8px"><button type="button" id="pdfEnv" style="height:38px;padding:0 14px;border-radius:8px;border:none;background:#FEE795;font-weight:800;cursor:pointer">Enviar por mail</button><button type="button" id="pdfX" style="height:38px;padding:0 14px;border-radius:8px;border:1.5px solid #0F3F27;background:#fff;font-weight:700;cursor:pointer">Cerrar</button></div>');
        document.getElementById('pdfX').onclick = function () { p.remove(); };
        document.getElementById('pdfCp').onclick = function () { try { navigator.clipboard.writeText(r.url); this.textContent = 'Copiado ✔'; } catch (e) { window.prompt('Copiá el link:', r.url); } };
        var env = document.getElementById('pdfEnv'), confirmado = false;
        env.onclick = function () {
          var para = document.getElementById('pdfMail').value.trim(), st = document.getElementById('pdfSt');
          if (!/^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(para)) { st.style.color = '#b3261e'; st.textContent = 'Escribí un correo válido.'; return; }
          if (!confirmado) { confirmado = true; env.textContent = 'Confirmar envío a ' + para; st.style.color = '#666'; st.textContent = 'Se manda desde la cuenta de Blumia con el PDF adjunto.'; return; }
          env.disabled = true; env.textContent = 'Enviando…';
          google.script.run.withSuccessHandler(function (rr) { env.textContent = '✔ Enviado'; st.style.color = rr && rr.alias === false ? '#9A7500' : '#0F3F27'; st.textContent = 'Enviado a ' + para + (rr && rr.desde ? ' desde ' + rr.desde + '.' : '.') + (rr && rr.alias === false ? ' (La dirección de la empresa todavía no está dada de alta como remitente en Gmail: sale de esta cuenta, pero las respuestas van a la de Blumia.)' : ''); })
            .withFailureHandler(function (e) { env.disabled = false; confirmado = false; env.textContent = 'Enviar por mail'; st.style.color = '#b3261e'; st.textContent = (e && e.message) || 'No se pudo enviar.'; })
            .informeEnviar(r.id, para, asunto, document.getElementById('pdfMsg').value, '');
        };
      }).withFailureHandler(function (e) { fin((e && e.message) || e); }).informePDF('__FILEID__', '');
    }, fin);
  };
  document.getElementById('edSave').onclick = function () {
    var btn = this; btn.textContent = 'Guardando...'; btn.disabled = true;
    var fin = function (t) { btn.textContent = t; btn.disabled = false; setTimeout(function () { btn.textContent = 'GUARDAR CAMBIOS'; }, 2500); };
    guardar(function () { fin('GUARDADO'); }, function () { fin('ERROR'); });
  };
}
