/* Ranking de candidatos para presentar al cliente (pedido de Ariel 29/10): se ordena a mano en la web (arrastrando, con flechas en el celular),
   el orden y un comentario por persona se guardan por búsqueda en Drive («Ranking - orden.json» en la carpeta de la búsqueda)
   y de ahí sale el informe para el cliente. El sistema solo sugiere el orden por las notas; el orden final es criterio de quien lidera la búsqueda.
   Página: ?tipo=ranking&perfil=<id> (menú 🧰 Candidatos → Ranking). El informe no lleva edad, género ni situación familiar. */

const RANKING_ARCHIVO_ = 'Ranking - orden.json';
const RANKING_VAL_ = { 'A': 4, 'B+': 3, 'B': 2, 'C': 0 };

function rankingHTML_(perfilId) {
  const cfg = getConfig(perfilId);
  const datos = { perfil: perfilId, puesto: String(cfg.nombre || perfilId), empresa: empresaInforme_(cfg), volver: URL_WEBAPP + '?tipo=kanban&perfil=' + encodeURIComponent(perfilId) };
  const js = o => JSON.stringify(o).replace(/</g, '\\u003c');
  let html = HtmlService.createTemplateFromFile('RankingPagina').getRawContent();
  const fuente = HtmlService.createTemplateFromFile('FuenteMarca').getRawContent();
  html = html.replace('%%FUENTE%%', () => fuente).replace('%%DATOS%%', () => js(datos)).replace('%%SORTABLE%%', () => HtmlService.createHtmlOutputFromFile('SortableLib').getContent());
  return HtmlService.createHtmlOutput(html).setTitle('Ranking · ' + datos.puesto + ' | Blumia')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function rankingArchivo_(cfg) {
  const it = carpetaPerfil_(cfg).getFilesByName(RANKING_ARCHIVO_);
  return it.hasNext() ? it.next() : null;
}

function rankingLeer_(cfg) {
  const f = rankingArchivo_(cfg);
  if (!f) return null;
  try { return JSON.parse(f.getBlob().getDataAsString()) || null; } catch (e) { return null; }
}

// personas que siguen en carrera (no descartadas ni eliminadas) y ya tienen alguna nota de screening o competencias
function rankingPool_(perfilId) {
  const filas = JSON.parse(doGetKanban_(perfilId));
  const zona = txt => { const m = String(txt || '').match(/Zona:\s*([^|(]+)/i); return m ? m[1].trim() : ''; };
  const out = [];
  filas.forEach(c => {
    const m = metEstado_(c.estado);
    if (!m || m.t === 'desc') return;
    if (!(m.lv >= 1 || c.notaScr || c.notaComp)) return;
    const emp = c.notaPresEmp || c.notaScrEmp || '';
    out.push({ fila: c.fila, nombre: String(c.nombre || '').trim(), estado: c.estado, etapa: ['CV', 'Screening', 'Competencias', 'Empresa', 'Empresa', 'Contratado'][m.lv] || '', reserva: m.t === 'res',
      notas: { cv: String(c.nota || ''), scr: String(c.notaScr || ''), comp: String(c.notaComp || ''), emp: String(emp) },
      pret: Number(c.pretNum) || 0, formacion: String(c.formacion || '').slice(0, 80), zona: zona(c.requisitos) });
  });
  return out;
}

function rankingScore_(c) {
  const v = k => RANKING_VAL_[c.notas[k]] || 0;
  return v('emp') * 1000 + v('comp') * 100 + v('scr') * 10 + v('cv') - (c.reserva ? 5000 : 0);
}

function rankingDatos_(perfilId) {
  const cfg = getConfig(perfilId);
  const pool = rankingPool_(perfilId);
  const guard = rankingLeer_(cfg);
  const porNombre = {}, porFila = {};
  pool.forEach(c => { porNombre[c.nombre.toLowerCase()] = c; porFila[c.fila] = c; });
  let orden = [], sugerido = false;
  if (guard && Array.isArray(guard.orden)) {
    guard.orden.forEach(o => {
      const c = porNombre[String(o.nombre || '').toLowerCase()] || porFila[o.fila];
      if (c && !orden.some(x => x.fila === c.fila)) orden.push({ fila: c.fila, com: String(o.com || '') });
    });
  } else {
    sugerido = true;
    orden = pool.filter(c => !c.reserva && (c.notas.emp || ['A', 'B+'].indexOf(c.notas.comp) >= 0)).sort((a, b) => rankingScore_(b) - rankingScore_(a)).map(c => ({ fila: c.fila, com: '' }));
  }
  return { candidatos: pool, orden: orden, sugerido: sugerido, mostrarPret: guard && guard.mostrarPret === false ? false : true, fecha: guard && guard.fecha || '' };
}

function rankingSugerir_(perfilId) {
  return rankingPool_(perfilId).filter(c => !c.reserva && (c.notas.emp || ['A', 'B+', 'B'].indexOf(c.notas.comp) >= 0 || ['A', 'B+'].indexOf(c.notas.scr) >= 0))
    .sort((a, b) => rankingScore_(b) - rankingScore_(a)).map(c => ({ fila: c.fila, com: '' }));
}

function rankingGuardar_(perfilId, orden, mostrarPret) {
  const cfg = getConfig(perfilId);
  const pool = rankingPool_(perfilId), porFila = {};
  pool.forEach(c => { porFila[c.fila] = c; });
  const lim = (v, n) => String(v == null ? '' : v).replace(/[<>]/g, '').trim().slice(0, n);
  const lista = [];
  (Array.isArray(orden) ? orden : []).slice(0, 60).forEach(o => {
    const c = porFila[Number(o && o.fila)];
    if (c && !lista.some(x => x.fila === c.fila)) lista.push({ fila: c.fila, nombre: c.nombre, com: lim(o.com, 500) });
  });
  const txt = JSON.stringify({ orden: lista, mostrarPret: mostrarPret !== false, fecha: Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy HH:mm') });
  const f = rankingArchivo_(cfg);
  if (f) f.setContent(txt); else carpetaPerfil_(cfg).createFile(RANKING_ARCHIVO_, txt, MimeType.PLAIN_TEXT);
  try { registrarHistorial_(cfg.hoja_nombre || perfilId, '(búsqueda)', 'Ranking', 'Ranking guardado (' + lista.length + ' personas)', QUIEN_); } catch (e) { }
  return { ok: true, n: lista.length };
}

// Informe para el cliente a partir del orden guardado (mismo formato que el informe de proceso)
function rankingInforme_(perfilId) {
  const cfg = getConfig(perfilId);
  const guard = rankingLeer_(cfg);
  if (!guard || !Array.isArray(guard.orden) || !guard.orden.length) throw new Error('Armá y guardá el ranking antes de generar el informe.');
  const pool = rankingPool_(perfilId), porNombre = {}, porFila = {};
  pool.forEach(c => { porNombre[c.nombre.toLowerCase()] = c; porFila[c.fila] = c; });
  const items = [];
  guard.orden.forEach(o => { const c = porNombre[String(o.nombre || '').toLowerCase()] || porFila[o.fila]; if (c) items.push({ c: c, com: String(o.com || '') }); });
  if (!items.length) throw new Error('Las personas del ranking ya no están en carrera. Volvé a armarlo.');
  const empresa = empresaInforme_(cfg), puesto = String(cfg.nombre || perfilId);
  const html = rankingArmarHTML_(puesto, empresa, items, guard.mostrarPret !== false);
  const it = DriveApp.getFoldersByName('Informes Clientes');
  const carpeta = it.hasNext() ? it.next() : DriveApp.createFolder('Informes Clientes');
  const nombre = 'Ranking_' + perfilId + '.html';
  const ex = carpeta.getFilesByName(nombre);
  const arch = ex.hasNext() ? ex.next() : carpeta.createFile(nombre, '<html></html>', MimeType.HTML);
  arch.setContent(conFuentePdf_(html).replace('__FILEID__', arch.getId()));
  try { registrarHistorial_(cfg.hoja_nombre || perfilId, '(búsqueda)', 'Ranking', 'Informe de ranking generado (' + items.length + ' personas)', QUIEN_); } catch (e) { }
  return { url: URL_WEBAPP + '?tipo=verinforme&id=' + arch.getId(), n: items.length };
}

function rankingArmarHTML_(puesto, empresa, items, mostrarPret) {
  const esc = s => String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  const LOGO = 'https://lh3.googleusercontent.com/d/1457Wfl9zeY1V3sAGtY1IT2BmF2lqpoy8';
  const money = n => n ? '$' + Number(n).toLocaleString('es-AR') : '—';
  const chip = (lb, v) => v ? '<span class="rk-n rk-' + String(v).replace('+', 'p') + '"><i>' + lb + '</i>' + esc(v) + '</span>' : '';
  const fecha = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
  const tabla = '<div class="rk-tw"><table class="rk-t"><thead><tr><th>#</th><th>Candidato</th><th>Screening</th><th>Competencias</th><th>Formación</th>' + (mostrarPret ? '<th>Pretensión</th>' : '') + '</tr></thead><tbody>' +
    items.map((x, i) => '<tr><td><b>' + (i + 1) + '</b></td><td contenteditable="true">' + esc(x.c.nombre) + '</td><td>' + esc(x.c.notas.scr || '—') + '</td><td>' + esc(x.c.notas.comp || '—') + '</td><td>' + esc(x.c.formacion || '—') + '</td>' + (mostrarPret ? '<td>' + money(x.c.pret) + '</td>' : '') + '</tr>').join('') + '</tbody></table></div>';
  const cards = items.map((x, i) => '<div class="rk-card"><div class="rk-pos">' + (i + 1) + '</div><div class="rk-body"><div class="rk-nom" contenteditable="true">' + esc(x.c.nombre) + '</div>' +
    '<div class="rk-meta">' + [x.c.formacion, x.c.zona && /fuera/i.test(x.c.zona) ? 'Fuera de zona' : '', mostrarPret && x.c.pret ? 'Pretensión ' + money(x.c.pret) : ''].filter(Boolean).map(esc).join(' · ') + '</div>' +
    '<div class="rk-ch">' + chip('Screening', x.c.notas.scr) + chip('Competencias', x.c.notas.comp) + chip('Empresa', x.c.notas.emp) + '</div>' +
    '<div class="rk-com" contenteditable="true">' + (x.com ? esc(x.com) : 'Escribí acá la recomendación sobre esta persona.') + '</div></div></div>').join('');
  const extra = '.rk-tw{overflow-x:auto}.rk-t{width:100%;border-collapse:collapse;font-size:13px}.rk-t th{text-align:left;font-size:10px;letter-spacing:.12em;text-transform:uppercase;color:var(--mid);padding:8px;border-bottom:2px solid var(--dark)}.rk-t td{padding:10px 8px;border-bottom:1px solid var(--line)}' +
    '.rk-card{display:flex;gap:18px;border:1px solid var(--line);border-radius:4px;padding:20px;margin-bottom:14px;break-inside:avoid;background:#fff}.rk-pos{flex:none;width:52px;height:52px;border-radius:50%;background:var(--green);color:#fff;font-size:24px;font-weight:800;display:flex;align-items:center;justify-content:center}.rk-card:first-child .rk-pos{background:var(--amber);color:var(--dark)}' +
    '.rk-body{flex:1;min-width:0}.rk-nom{font-size:20px;font-weight:800;letter-spacing:-.02em}.rk-meta{font-size:12px;color:var(--mid);margin:2px 0 8px}.rk-ch{display:flex;gap:8px;flex-wrap:wrap;margin-bottom:10px}.rk-n{font-size:12px;font-weight:800;border-radius:4px;padding:4px 10px;background:var(--bg2)}.rk-n i{font-style:normal;font-weight:600;color:var(--mid);margin-right:6px}.rk-A{background:#1A6645;color:#fff}.rk-A i{color:#cfe}.rk-Bp{background:#B7E0C6;color:#14502F}.rk-B{background:#F7DC6F;color:#6B5800}.rk-C{background:#E06666;color:#fff}' +
    '.rk-com{font-size:13px;line-height:1.6;color:#333}@media(max-width:700px){.rk-card{flex-direction:column;gap:10px}.rk-t{font-size:11px}.hero,.busqueda{padding-left:20px;padding-right:20px}.h-title{font-size:30px}.h-stats{grid-template-columns:1fr 1fr}}';
  const script = '<script>function guardarCambios(){var btn=document.querySelector("#barra-guardar button");btn.textContent="Guardando...";btn.disabled=true;var clon=document.documentElement.cloneNode(true);var b=clon.querySelector("#barra-guardar button");if(b){b.textContent="GUARDAR CAMBIOS";b.disabled=false;}google.script.run.withSuccessHandler(function(r){btn.textContent=(r&&r.ok)?"GUARDADO":"ERROR";btn.disabled=false;setTimeout(function(){btn.textContent="GUARDAR CAMBIOS";},2500);}).withFailureHandler(function(){btn.textContent="ERROR";btn.disabled=false;}).guardarInformeEditado("__FILEID__","<!DOCTYPE html>"+clon.outerHTML);}</scr' + 'ipt>';
  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Ranking · ' + esc(puesto) + '</title><style>' + cssInformeProceso_() + extra + '</style></head><body>' +
    '<div id="barra-guardar" contenteditable="false"><button onclick="guardarCambios()">GUARDAR CAMBIOS</button></div>' +
    '<section class="hero"><img class="hero-logo" src="' + LOGO + '"><div class="eyebrow">Atracción y Elección · Candidatos recomendados</div><div class="cliente-nombre" contenteditable="true">' + esc(empresa) + '</div>' +
    '<h1 class="h-title" contenteditable="true">Los candidatos que <em>recomendamos</em> para ' + esc(puesto) + '</h1><div class="h-stats"><div class="hs-item"><div class="hs-num">' + items.length + '</div><div class="hs-lbl">Candidatos presentados</div></div><div class="hs-item"><div class="hs-num" style="font-size:26px;line-height:1.5">' + fecha + '</div><div class="hs-lbl">Fecha</div></div></div></section>' +
    '<section class="busqueda"><div class="section"><div class="sec-num">01 · Ranking</div><h3 class="sec-title">Vista <em>comparativa</em></h3>' + tabla + '</div>' +
    '<div class="section"><div class="sec-num">02 · Detalle</div><h3 class="sec-title">Cada <em>candidato</em></h3>' + cards + '</div></section>' +
    '<footer>Blumia Consulting · Documento confidencial preparado para ' + esc(empresa) + '</footer>' + script + '</body></html>';
}
