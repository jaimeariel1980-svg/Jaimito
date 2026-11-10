/* ============================================================
   ENTREVISTA DE EXPLORACIÓN (10/11, Ariel: la cliente no quiere competencias sin validar antes ciertas condiciones).
   Etapa opt-in por búsqueda (Config etapa_exploracion = SI) entre Screening y Competencias. No lleva puntaje ni pesos: es una evaluación
   general de la entrevista (A / B+ / B / C). Cada candidato tiene su propia ficha: perfil predominante, síntesis, foco, preguntas (5 a medida y
   una común por búsqueda) y «qué buscaría validar»; después de la entrevista se completa el registro (evidencias, fortalezas, alertas, conclusión).
   Pantalla: Exploracion.html (?tipo=exploracion&perfil=<id>&nombre=<candidato>). Las respuestas se guardan como las de screening y competencias
   (pestaña Evaluaciones, etapa «exploracion», un .json en la carpeta «Exploración» de la búsqueda) y el PDF queda en esa carpeta.
   ============================================================ */

function colExplPdf_(sh, crear) {
  const h = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  let c = h.indexOf('Exploración PDF') + 1;
  if (!c && crear) {
    c = Math.max(sh.getLastColumn(), 46) + 1;
    if (sh.getMaxColumns() < c) sh.insertColumnsAfter(sh.getMaxColumns(), c - sh.getMaxColumns());
    sh.getRange(1, c).setValue('Exploración PDF').setFontWeight('bold');
  }
  return c;
}

// carpeta «Exploración» dentro de la carpeta de la búsqueda (se crea la primera vez y queda anotada en Config)
function carpetaExpl_(cfg, perfilId) {
  if (cfg.folder_pdf_exploracion) { try { return DriveApp.getFolderById(cfg.folder_pdf_exploracion); } catch (e) { } }
  const madre = carpetaPerfil_(cfg);
  const it = madre.getFoldersByName('Exploración');
  const f = it.hasNext() ? it.next() : madre.createFolder('Exploración');
  try { setConfigCampo_(perfilId, 'folder_pdf_exploracion', f.getId()); } catch (e) { }
  return f;
}

function explBuscarFila_(sh, nombre) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const col = sh.getRange(2, 2, n - 1, 1).getValues(), k = evalNorm_(nombre);
  for (let i = 0; i < col.length; i++) if (evalNorm_(col[i][0]) === k) return i + 2;
  return 0;
}

function cvIdDeFila_(sh, fila) {
  const cel = sh.getRange(fila, 8);
  let url = '';
  const f = String(cel.getFormula() || '').match(/HYPERLINK\("([^"]+)"/i);
  if (f) url = f[1];
  if (!url) { const rt = cel.getRichTextValue(); if (rt && rt.getLinkUrl()) url = rt.getLinkUrl(); }
  if (!url) url = String(cel.getValue() || '');
  const m = url.match(/[?&]id=([-\w]{25,})/) || url.match(/[-\w]{25,}/);
  return m ? (m[1] || m[0]) : '';
}

// lo que necesita la pantalla además de getDatosPerfil: link del CV, pregunta común y qué validar (de ⚙ Configurar → Datos)
function explDatos_(perfilId, nombre) {
  const cfg = getConfig(perfilObligatorio_(perfilId));
  const out = { activa: explActiva_(cfg), comun: String(cfg.expl_pregunta_comun || '').trim(), validar: String(cfg.expl_validar || '').trim(), cv: '', formacion: '', zona: '' };
  try {
    const sh = getSheet(cfg), f = explBuscarFila_(sh, nombre);
    if (f) {
      const id = cvIdDeFila_(sh, f);
      if (id) out.cv = 'https://drive.google.com/file/d/' + id + '/view';
      out.formacion = String(sh.getRange(f, 23).getValue() || '');
      out.zona = String(sh.getRange(f, 24).getValue() || '');
    }
  } catch (e) { Logger.log('explDatos_: ' + e); }
  return out;
}

const EXPL_SISTEMA_ = 'Sos consultor/a senior de selección en Blumia Consulting. Preparás la FICHA DE LA ENTREVISTA DE EXPLORACIÓN de un candidato: una entrevista personalizada, ' +
  'posterior al screening y previa a la evaluación de competencias, para validar lo que la empresa cliente necesita confirmar antes de seguir. ' +
  'Respondé SOLO con un JSON válido, sin texto antes ni después, con esta forma exacta: ' +
  '{"perfil_predominante":"una línea: el perfil que muestra el CV (ej. «operaciones, supply chain, finanzas y gestión de proyectos»)",' +
  '"sintesis":"2 o 3 oraciones: qué muestra su trayectoria (empresas y roles reales del CV) y qué NO evidencia todavía para este puesto",' +
  '"foco":"1 a 3 oraciones: qué hay que determinar en esta entrevista con esta persona en particular",' +
  '"preguntas":[{"titulo":"2 a 4 palabras (ej. Orientación comercial)","pregunta":"pregunta abierta en voseo"}],' +
  '"validar":"1 a 3 oraciones: qué buscarías validar al terminar la entrevista"}. ' +
  'REGLAS: ' +
  '1) Exactamente 5 preguntas, todas a medida de esta persona: tomá sus empresas, roles y sectores reales del CV. No preguntes por lo que el CV ya demuestra; profundizá en lo que no evidencia o en lo que la empresa necesita confirmar. ' +
  '2) Cada pregunta es abierta, en voseo rioplatense, y pide un caso o situación concreta con su resultado. Sin preguntas de sí o no. ' +
  '3) Usá solo lo que dice el CV. No inventes empleadores, cargos, cifras ni logros. Si el CV dice poco, decilo en la síntesis y preguntá por eso. ' +
  '4) Priorizá lo que la empresa cliente quiere validar (se te indica) y la ficha del puesto. ' +
  '5) NUNCA uses ni menciones edad, género, estado civil, situación familiar, nacionalidad, salud ni religión: no son criterio. ' +
  '6) Tono profesional, directo y respetuoso. Nada de adjetivos vacíos.';

// Arma con IA el borrador de la ficha a partir del CV, el perfil del puesto y lo que se quiere validar. No guarda nada: la persona revisa y edita en pantalla.
function explPreparar_(perfilId, nombre) {
  const cfg = getConfig(perfilObligatorio_(perfilId)), sh = getSheet(cfg);
  const fila = explBuscarFila_(sh, nombre);
  if (!fila) throw new Error('No encontré a «' + nombre + '» en la búsqueda.');
  let cv = '';
  try { const id = cvIdDeFila_(sh, fila); if (id) cv = String(textoDeArchivo_(DriveApp.getFileById(id), 'Temp_Exploracion') || ''); } catch (e) { Logger.log('explPreparar_ CV: ' + e); }
  // el análisis del CV trae edad y género: no se le pasan a la IA
  const req = String(sh.getRange(fila, 6).getValue() || '').split(' || DECLARADO:')[0].split(/\s*\|\s*/)
    .filter(x => x && !/^\s*(edad|g[eé]nero|sexo|estado civil|situaci[oó]n familiar|nacionalidad)\b/i.test(x)).join(' | ');
  const resumen = String(sh.getRange(fila, 7).getValue() || '');
  if (!cv.trim() && !resumen.trim()) throw new Error('Este candidato no tiene un CV legible ni un resumen: completá la ficha a mano.');
  const validar = String(cfg.expl_validar || '').trim();
  const usuario = 'PUESTO: ' + String(cfg.nombre_perfil || cfg.nombre || '') + (cfg.empresa && !busquedaCiega_(cfg) ? ' · ' + cfg.empresa : '') + '\n\n' +
    'LO QUE LA EMPRESA QUIERE VALIDAR ANTES DE COMPETENCIAS' + (validar ? ':\n' + validar : ': (no se cargó; deducilo de la ficha del puesto)') + '\n\n' +
    'FICHA DEL PUESTO:\n' + (perfilFicha_(cfg, 'cv') || String(cfg.prompt_extra || '').slice(0, 6000) || '(sin ficha cargada)') + '\n\n' +
    'CANDIDATO: ' + nombre + '\n' +
    'ANÁLISIS DEL CV (resumen de la IA): ' + resumen + '\n' + (req ? 'DATOS DEL ANÁLISIS: ' + req + '\n' : '') +
    '\nTEXTO DEL CV:\n' + (cv.trim() ? cv.slice(0, 9000) : '(no se pudo leer el archivo del CV; usá el resumen)');
  const r = extraerJSON_(procesarConClaude(EXPL_SISTEMA_, usuario, 'analisis')) || {};
  const lim = (s, n) => String(s || '').replace(/[<>]/g, '').replace(/\s+/g, ' ').trim().slice(0, n);
  const preguntas = (Array.isArray(r.preguntas) ? r.preguntas : []).map(p => ({ titulo: lim(p && p.titulo, 60), pregunta: lim(p && p.pregunta, 600) })).filter(p => p.pregunta).slice(0, 5);
  if (preguntas.length < 3) throw new Error('La IA no pudo armar las preguntas. Probá de nuevo en un rato o completá la ficha a mano.');
  return { perfil_predominante: lim(r.perfil_predominante, 200), sintesis: lim(r.sintesis, 900), foco: lim(r.foco, 900), preguntas: preguntas, validar: lim(r.validar, 700) };
}

// guarda la ficha y lo escrito hasta ahora sin dar nota ni mover de etapa (para preparar la entrevista antes de hacerla)
function explBorrador_(perfilId, nombre, estadoJSON, quien) {
  const cfg = getConfig(perfilObligatorio_(perfilId)), sh = getSheet(cfg);
  const fila = explBuscarFila_(sh, nombre);
  if (!fila) throw new Error('No encontré a «' + nombre + '» en la búsqueda.');
  const c = colExpl_(sh, false);
  const nota = c ? String(sh.getRange(fila, c).getValue() || '').trim() : '';
  guardarEstadoEval_(cfg, perfilObligatorio_(perfilId), nombre, 'exploracion', estadoJSON, nota, quien);
  if (!evalFila_(evalHoja_(), perfilObligatorio_(perfilId), nombre, 'exploracion')) throw new Error('No se pudo guardar la ficha. Probá de nuevo.');
  registrarHistorial_(sh.getName(), sh.getRange(fila, 2).getValue(), 'Exploración', 'Ficha guardada (borrador)', quien || 'sistema');
  return { ok: true };
}

function procesarExploracion_(data) {
  QUIEN_ = String(data.entrevistador || '');
  const perfilId = perfilObligatorio_(data.perfil), cfg = getConfig(perfilId), sheet = getSheet(cfg);
  const fila = explBuscarFila_(sheet, data.nombre);
  if (!fila) throw new Error("No encontré a '" + data.nombre + "' en la pestaña '" + cfg.hoja_nombre + "'. Verificá perfil/pestaña.");
  const v = String(data.resultado || '').trim().toUpperCase();
  if (['A', 'B+', 'B', 'C'].indexOf(v) < 0) throw new Error('Falta elegir el resultado: A, B+, B o C.');
  const cNota = colExpl_(sheet, true);
  const notaAntes = String(sheet.getRange(fila, cNota).getValue() || '').trim();
  const celda = sheet.getRange(fila, cNota);
  celda.setDataValidation(null).setValue(v);
  const colores = { A: '#E6F2EC', 'B+': '#E6F2EC', B: '#FEF3D8', C: '#FDEAE4' };
  celda.setBackground(colores[v] || '#fff').setFontWeight('bold').setHorizontalAlignment('center');
  registrarHistorial_(sheet.getName(), sheet.getRange(fila, 2).getValue(), data.editar ? 'Exploración editada' : 'Exploración',
    (data.editar ? (notaAntes || '—') + ' → ' + v : 'Formulario guardado · resultado ' + v) + (data.entrevistador ? ' · entrevistó ' + data.entrevistador : ''));
  if (!data.editar || aplicarEdicion_(sheet, fila, 'exploracion', notaAntes, v)) aplicarEstadoNota(sheet, fila, 'exploracion', v);
  motivoDesdeForm_(sheet, fila, data.motivo);
  try { guardarProximo_(sheet, fila, data.proximo, v === 'A' || v === 'B+', data.entrevistador || 'sistema'); } catch (e) { Logger.log('próximo entrevistador: ' + e); }
  guardarEstadoEval_(cfg, perfilId, data.nombre, 'exploracion', data.estado, v, data.entrevistador);
  try {
    const folder = carpetaExpl_(cfg, perfilId);
    const blob = Utilities.newBlob(logoParaPdf_(data.htmlFull), 'text/html', 'Exploracion_' + String(data.nombre).replace(/\s+/g, '_') + '.html');
    const pdf = folder.createFile(blob.getAs('application/pdf'));
    const c = colExplPdf_(sheet, true), cel = sheet.getRange(fila, c);
    cel.clearFormat();
    cel.setFormula('=HYPERLINK("' + pdf.getUrl() + '"; "Ver detalle")');
    cel.setFontLine('underline').setFontColor('#000');
    return 'OK';
  } catch (e) {
    throw new Error('Guardé el resultado, pero el PDF falló: ' + e.message);
  }
}
