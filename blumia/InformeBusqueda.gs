/* Informe general de la búsqueda para el cliente, en 3 pasos (pedido de Ariel 29/10):
   1) la IA analiza los números y propone hallazgos; 2) hace unas pocas preguntas con respuesta sugerida sobre lo que los datos no dicen;
   3) redacta el informe con lo que Ariel dejó y respondió. Los números salen siempre de la planilla (nunca los inventa la IA).
   No viajan nombres de candidatos ni edad ni situación familiar. Página: ?tipo=informebusqueda&perfil=<id> */

function informeBusquedaHTML_(perfilId) {
  const cfg = getConfig(perfilId);
  const datos = { perfil: perfilId, puesto: String(cfg.nombre || perfilId), empresa: empresaInforme_(cfg), volver: URL_WEBAPP + '?tipo=kanban&perfil=' + encodeURIComponent(perfilId) };
  const js = o => JSON.stringify(o).replace(/</g, '\\u003c');
  let html = HtmlService.createTemplateFromFile('InformeBusquedaPagina').getRawContent();
  const fuente = HtmlService.createTemplateFromFile('FuenteMarca').getRawContent();
  html = html.replace('%%FUENTE%%', () => fuente).replace('%%DATOS%%', () => js(datos));
  return HtmlService.createHtmlOutput(html).setTitle('Informe de la búsqueda · ' + datos.puesto + ' | Blumia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// Números de la búsqueda, calculados con las mismas reglas que el Resumen y el Informe de proceso (metEstado_)
function ibEstadisticas_(perfilId) {
  const cfg = getConfig(perfilId);
  const filas = JSON.parse(doGetKanban_(perfilId));
  const banda = Number(String(cfg.banda_salarial || '').replace(/\D/g, '')) || 0;
  const e = { total: 0, scr: 0, comp: 0, emp: 0, res: 0, desc: 0, dsin: 0, dpost: 0, dcomp: 0, activos: 0,
    notasCV: { A: 0, 'B+': 0, B: 0, C: 0, sin: 0 }, motivos: {}, fuentes: {}, formacion: {}, pret: { debajo: 0, linea: 0, encima: 0, total: 0 }, banda: banda };
  let primero = 0;
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
    if (c.recMs && (!primero || c.recMs < primero)) primero = c.recMs;
  });
  e.dias = primero ? Math.max(1, Math.round((Date.now() - primero) / 864e5)) : 0;
  const top = (o, n) => Object.keys(o).sort((a, b) => o[b] - o[a]).slice(0, n).map(k => ({ nombre: k, cantidad: o[k] }));
  e.motivosTop = top(e.motivos, 5); e.fuentesTop = top(e.fuentes, 5); e.formacionTop = top(e.formacion, 4);
  delete e.motivos; delete e.fuentes; delete e.formacion;
  return e;
}

function ibAnalizar_(perfilId) {
  const cfg = getConfig(perfilId);
  const est = ibEstadisticas_(perfilId);
  const puesto = String(cfg.nombre || perfilId);
  const sistema = 'Sos consultor/a senior de selección de personal en Blumia Consulting (Argentina). Vas a preparar el informe general de una búsqueda para el CLIENTE. ' +
    'Recibís los números reales de la búsqueda. Respondé SOLO con un JSON válido, sin texto antes ni después, con esta forma exacta: ' +
    '{"hallazgos":[{"titulo":"máx. 70 caracteres","detalle":"1 o 2 frases con el dato concreto","tipo":"positivo|atencion|dato"}],"preguntas":[{"id":"q1","texto":"pregunta corta","sugeridas":["respuesta posible 1","respuesta posible 2","respuesta posible 3"]}]}. ' +
    'De 4 a 6 hallazgos, ordenados por importancia para el cliente, apoyados SOLO en los números recibidos (no inventes datos, causas ni porcentajes que no se puedan calcular con ellos). ' +
    'Si la muestra es chica (menos de 15 CVs o menos de 5 en screening) decilo en un hallazgo y no saques conclusiones fuertes. ' +
    'De 4 a 6 preguntas para quien lidera la búsqueda, SOLO sobre lo que los números no dicen y que cambia cómo se escribe el informe: qué quiere destacar el cliente, si la banda salarial es negociable, qué se recomienda hacer con los candidatos en reserva, qué próximos pasos proponer, qué tono usar, si hay algo del mercado que explique lo observado. Cada pregunta con 2 o 3 respuestas sugeridas cortas (máx. 90 caracteres). ' +
    'Nada de nombres de personas. Español rioplatense profesional.';
  const usuario = 'Puesto: ' + puesto + '. Empresa: ' + empresaInforme_(cfg) + '. Números de la búsqueda: ' + JSON.stringify(est);
  let r;
  try { r = extraerJSON_(procesarConClaude(sistema, usuario, 'analisis')); } catch (err) { r = { hallazgos: [], preguntas: [] }; }
  const t = (v, n) => String(v == null ? '' : v).replace(/[<>]/g, '').trim().slice(0, n);
  const tipos = { positivo: 1, atencion: 1, dato: 1 };
  let hallazgos = (Array.isArray(r.hallazgos) ? r.hallazgos : []).slice(0, 6).map((h, i) => ({ id: 'h' + (i + 1), titulo: t(h && h.titulo, 90), detalle: t(h && h.detalle, 320), tipo: tipos[h && h.tipo] ? h.tipo : 'dato' })).filter(h => h.titulo);
  let preguntas = (Array.isArray(r.preguntas) ? r.preguntas : []).slice(0, 6).map((p, i) => ({ id: 'q' + (i + 1), texto: t(p && p.texto, 160), sugeridas: (Array.isArray(p && p.sugeridas) ? p.sugeridas : []).slice(0, 3).map(s => t(s, 110)).filter(Boolean) })).filter(p => p.texto);
  if (!preguntas.length) preguntas = [
    { id: 'q1', texto: '¿Qué querés destacar del proceso ante el cliente?', sugeridas: ['La velocidad y el volumen de candidatos', 'La calidad de los finalistas', 'El ajuste a la banda salarial'] },
    { id: 'q2', texto: '¿Qué recomendás hacer con los candidatos en reserva?', sugeridas: ['Mantenerlos para una próxima búsqueda', 'Retomarlos si los finalistas no cierran'] },
    { id: 'q3', texto: '¿Qué tono preferís para el informe?', sugeridas: ['Profesional y cercano', 'Formal'] }];
  return { est: est, hallazgos: hallazgos, preguntas: preguntas };
}

// Escribe el informe con los hallazgos que se dejaron y las respuestas, lo guarda editable en Drive y devuelve el link
function ibRedactar_(perfilId, hallazgos, respuestas) {
  const cfg = getConfig(perfilId);
  const est = ibEstadisticas_(perfilId);
  const lim = (v, n) => String(v == null ? '' : v).replace(/[<>]/g, '').trim().slice(0, n);
  const hs = (Array.isArray(hallazgos) ? hallazgos : []).slice(0, 8).map(h => ({ titulo: lim(h && h.titulo, 90), detalle: lim(h && h.detalle, 400) })).filter(h => h.titulo);
  const rs = (Array.isArray(respuestas) ? respuestas : []).slice(0, 8).map(r => ({ pregunta: lim(r && r.pregunta, 160), respuesta: lim(r && r.respuesta, 500) })).filter(r => r.respuesta);
  if (!hs.length) throw new Error('Dejá al menos un hallazgo para armar el informe.');
  const empresa = empresaInforme_(cfg), puesto = String(cfg.nombre || perfilId);
  const sistema = 'Sos consultor/a senior de selección de personal en Blumia Consulting (Argentina). Redactás el informe general de una búsqueda dirigido al CLIENTE (' + empresa + '). ' +
    'Usá SOLO los números y los hallazgos recibidos y las respuestas de quien lidera la búsqueda; no inventes datos, porcentajes ni causas. Respetá el tono pedido en las respuestas (si no hay, profesional y cercano). ' +
    'No nombres candidatos, no menciones edad, género ni situación familiar, ni criterios internos de evaluación. Hablá de la búsqueda en tercera persona ("la búsqueda", "se recibieron"). Español rioplatense profesional. ' +
    'Respondé SOLO con un JSON válido con esta forma exacta: {"titulo":"máx. 90 caracteres, con gancho y sin exagerar","resumen":"párrafo de 3 a 5 frases","hallazgos":[{"titulo":"","texto":"2 o 3 frases"}],"recomendaciones":["acción concreta"],"proximos_pasos":["paso concreto"],"cierre":"1 o 2 frases"}. ' +
    'Hallazgos: reescribí los recibidos, en el mismo orden. 2 a 4 recomendaciones y 2 a 4 próximos pasos, coherentes con las respuestas.';
  const usuario = 'Puesto: ' + puesto + '. Números: ' + JSON.stringify(est) + '. Hallazgos a incluir: ' + JSON.stringify(hs) + '. Respuestas de quien lidera la búsqueda: ' + JSON.stringify(rs);
  const r = extraerJSON_(procesarConClaude(sistema, usuario, 'informe'));
  const t = (v, n) => lim(v, n);
  const inf = {
    titulo: t(r.titulo, 120) || ('Informe de la búsqueda · ' + puesto),
    resumen: t(r.resumen, 900),
    hallazgos: (Array.isArray(r.hallazgos) ? r.hallazgos : []).slice(0, 8).map(h => ({ titulo: t(h && h.titulo, 100), texto: t(h && h.texto, 600) })).filter(h => h.titulo),
    recomendaciones: (Array.isArray(r.recomendaciones) ? r.recomendaciones : []).slice(0, 5).map(x => t(x, 300)).filter(Boolean),
    proximos: (Array.isArray(r.proximos_pasos) ? r.proximos_pasos : []).slice(0, 5).map(x => t(x, 300)).filter(Boolean),
    cierre: t(r.cierre, 400)
  };
  const html = ibArmarHTML_(cfg, puesto, empresa, est, inf);
  const it = DriveApp.getFoldersByName('Informes Clientes');
  const carpeta = it.hasNext() ? it.next() : DriveApp.createFolder('Informes Clientes');
  const nombre = 'Informe_Busqueda_' + perfilId + '.html';
  const ex = carpeta.getFilesByName(nombre);
  const arch = ex.hasNext() ? ex.next() : carpeta.createFile(nombre, '<html></html>', MimeType.HTML);
  arch.setContent(conFuentePdf_(html).replace('__FILEID__', arch.getId()));
  try { registrarHistorial_(cfg.hoja_nombre || perfilId, '(búsqueda)', 'Informe', 'Informe general para el cliente armado', QUIEN_); } catch (e) { }
  return { url: URL_WEBAPP + '?tipo=verinforme&id=' + arch.getId(), informe: inf };
}

function ibArmarHTML_(cfg, puesto, empresa, e, inf) {
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const ce = s => '<span contenteditable="true">' + esc(s) + '</span>';
  const pct = (a, b) => b ? Math.round(a / b * 100) + '%' : '—';
  const fecha = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
  const lista = a => a.length ? '<ul>' + a.map(x => '<li contenteditable="true">' + esc(x) + '</li>').join('') + '</ul>' : '';
  const emb = [['CVs recibidos', e.total, '100%'], ['Pasaron screening', e.scr, pct(e.scr, e.total)], ['Evaluación de competencias', e.comp, pct(e.comp, e.total)], ['Avanzan a la empresa', e.emp, pct(e.emp, e.total)]]
    .map(x => '<div class="k"><div class="n">' + x[1] + '</div><div class="l">' + x[0] + '</div><div class="p">' + x[2] + ' del total</div></div>').join('');
  const desc = '<div class="g3"><div class="k2"><div class="n2">' + e.dsin + '</div><div class="l">Análisis de CV</div></div><div class="k2"><div class="n2">' + e.dpost + '</div><div class="l">Post screening</div></div><div class="k2"><div class="n2">' + e.dcomp + '</div><div class="l">Competencias y empresa</div></div></div>';
  const mot = e.motivosTop.length ? '<ul class="mot">' + e.motivosTop.map(m => '<li>' + esc(m.nombre) + ' <b>' + m.cantidad + '</b></li>').join('') + '</ul>' : '';
  const sal = (e.banda && e.pret.total >= 5) ? '<div class="sec"><div class="eye">Expectativa salarial</div><h3>Lo que pretenden los candidatos</h3><div class="g3"><div class="k2"><div class="n2">' + pct(e.pret.debajo, e.pret.total) + '</div><div class="l">Por debajo de la banda</div></div><div class="k2"><div class="n2">' + pct(e.pret.linea, e.pret.total) + '</div><div class="l">En línea con la banda</div></div><div class="k2"><div class="n2">' + pct(e.pret.encima, e.pret.total) + '</div><div class="l">Por encima de la banda</div></div></div><p class="nota">Sobre ' + e.pret.total + ' candidatos que informaron pretensión.</p></div>' : '';
  const hall = inf.hallazgos.map((h, i) => '<div class="h"><div class="hn">' + (i + 1) + '</div><div><h4 contenteditable="true">' + esc(h.titulo) + '</h4><p contenteditable="true">' + esc(h.texto) + '</p></div></div>').join('');
  const css = ':root{--g:#0F3F27;--a:#FEE795;--bg:#EEF7EE;--l:#D5E0CE;--t:#1A1A1A;--m:#555;--f:\'PP Pangram Sans\',\'Poppins\',sans-serif}*{box-sizing:border-box;margin:0;padding:0;-webkit-print-color-adjust:exact;print-color-adjust:exact}body{font-family:var(--f);color:var(--t);background:#fff;font-size:14px;line-height:1.55}[contenteditable]:focus{outline:2px solid var(--g);outline-offset:2px;background:rgba(254,231,149,.25)}' +
    '#bg{position:fixed;top:14px;right:14px;z-index:9}#bg button{background:var(--a);border:none;padding:11px 20px;border-radius:8px;font:800 11px var(--f);letter-spacing:.12em;text-transform:uppercase;cursor:pointer;box-shadow:0 3px 10px rgba(0,0,0,.25)}' +
    '.hero{background:var(--g);color:#fff;padding:54px 44px 44px}.hero .eye{color:var(--a);font-size:11px;font-weight:800;letter-spacing:.2em;text-transform:uppercase;margin-bottom:12px}.hero h1{font-size:38px;line-height:1.1;letter-spacing:-.02em;max-width:880px}.hero .meta{margin-top:16px;color:rgba(255,255,255,.8);font-size:13px}' +
    '.wrap{padding:0 44px}.sec{padding:34px 0;border-bottom:1px solid var(--l)}.sec .eye{font-size:10px;font-weight:800;letter-spacing:.2em;text-transform:uppercase;color:var(--m);margin-bottom:6px}.sec h3{font-size:22px;color:var(--g);margin-bottom:16px;letter-spacing:-.02em}.sec p{max-width:820px}' +
    '.g4{display:grid;grid-template-columns:repeat(4,1fr);gap:14px}.g3{display:grid;grid-template-columns:repeat(3,1fr);gap:14px}.k{background:var(--bg);border-radius:12px;padding:20px;border-top:4px solid var(--g)}.k .n{font-size:42px;font-weight:800;color:var(--g);line-height:1}.k .l{font-weight:700;margin-top:8px;font-size:13px}.k .p{font-size:12px;color:var(--m)}' +
    '.k2{border:1px solid var(--l);border-radius:12px;padding:18px}.n2{font-size:32px;font-weight:800;color:var(--g);line-height:1}.l{font-size:12px;font-weight:700;margin-top:6px;color:var(--m)}.nota{font-size:12px;color:var(--m);margin-top:10px}.mot{margin-top:14px;columns:2;list-style:none}.mot li{padding:4px 0;font-size:13px;break-inside:avoid}.mot b{color:var(--g)}' +
    '.h{display:flex;gap:16px;margin-bottom:18px;break-inside:avoid}.hn{flex:none;width:38px;height:38px;border-radius:50%;background:var(--a);color:var(--g);font-weight:800;display:flex;align-items:center;justify-content:center}.h h4{font-size:16px;color:var(--g);margin-bottom:4px}' +
    '.sec ul{padding-left:20px;max-width:820px}.sec li{margin-bottom:8px}.cierre{background:var(--bg);padding:30px 44px;font-size:15px}footer{padding:22px 44px;font-size:11px;color:var(--m);text-align:center}' +
    '@media(max-width:700px){.g4,.g3{grid-template-columns:1fr 1fr}.hero h1{font-size:28px}.mot{columns:1}.hero,.wrap,.cierre{padding-left:20px;padding-right:20px}}@media print{@page{size:A4;margin:12mm 10mm}#bg{display:none!important}.sec,.h,.k{break-inside:avoid}}';
  const script = '<script>function guardarCambios(){var b=document.querySelector("#bg button");b.textContent="Guardando...";b.disabled=true;var c=document.documentElement.cloneNode(true);var x=c.querySelector("#bg button");if(x){x.textContent="GUARDAR CAMBIOS";x.disabled=false}google.script.run.withSuccessHandler(function(r){b.textContent=(r&&r.ok)?"GUARDADO":"ERROR";b.disabled=false;setTimeout(function(){b.textContent="GUARDAR CAMBIOS"},2500)}).withFailureHandler(function(){b.textContent="ERROR";b.disabled=false}).guardarInformeEditado("__FILEID__","<!DOCTYPE html>"+c.outerHTML)}</scr' + 'ipt>';
  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Informe · ' + esc(puesto) + '</title><style>' + css + '</style></head><body>' +
    '<div id="bg"><button onclick="guardarCambios()">GUARDAR CAMBIOS</button></div>' +
    '<section class="hero"><div class="eye">Blumia Consulting · Informe de búsqueda</div><h1 contenteditable="true">' + esc(inf.titulo) + '</h1><div class="meta">' + esc(empresa) + ' · ' + esc(puesto) + ' · ' + fecha + '</div></section>' +
    '<div class="wrap"><div class="sec"><div class="eye">Resumen</div><h3>Qué pasó en esta búsqueda</h3><p contenteditable="true">' + esc(inf.resumen) + '</p></div>' +
    '<div class="sec"><div class="eye">El proceso en números</div><h3>Del volumen a la precisión</h3><div class="g4">' + emb + '</div>' + (e.res ? '<p class="nota">Además, ' + e.res + (e.res === 1 ? ' candidato queda' : ' candidatos quedan') + ' en reserva.</p>' : '') + (e.dias ? '<p class="nota">Tiempo desde el primer CV recibido: ' + e.dias + ' días.</p>' : '') + '</div>' +
    '<div class="sec"><div class="eye">Descartes</div><h3>Filtro por etapa</h3>' + desc + mot + '</div>' + sal +
    '<div class="sec"><div class="eye">Hallazgos</div><h3>Lo que vimos</h3>' + hall + '</div>' +
    (inf.recomendaciones.length ? '<div class="sec"><div class="eye">Recomendaciones</div><h3>Qué sugerimos</h3>' + lista(inf.recomendaciones) + '</div>' : '') +
    (inf.proximos.length ? '<div class="sec"><div class="eye">Próximos pasos</div><h3>Cómo seguimos</h3>' + lista(inf.proximos) + '</div>' : '') + '</div>' +
    (inf.cierre ? '<div class="cierre" contenteditable="true">' + esc(inf.cierre) + '</div>' : '') +
    '<footer>Blumia Consulting · Documento preparado para ' + esc(empresa) + '</footer>' + script + '</body></html>';
}
