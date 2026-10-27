const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
const SHEET_MASTER_ID = '151V8cc-w00tP9C6mSxVaKNkRxgKXgUDI4L0KWXeNJGI';
const CORREO_DESTINO = 'empleos@blumiaconsulting.com';
const URL_DEPLOY = 'https://script.google.com/macros/s/AKfycbytekPKN31bZLxczDVknMr433K9KflzmrR-dCWsC1hLlx2MO7tbw6SrMjFXa1PnucQo/exec';
function URL_BASE() { return ScriptApp.getService().getUrl(); }

// Mientras se arma una página de solo lectura (kanban, inicio) la pestaña Config se lee UNA vez: cada lectura trae todos los requisitos para la IA (miles de caracteres por búsqueda).
let CFG_MEMO_ = null;
function configValores_() {
  if (CFG_MEMO_ && CFG_MEMO_.v) return CFG_MEMO_.v;
  const config = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config');
  if (!config) throw new Error('Falta pestaña Config. Corré setupInicial()');
  const v = config.getDataRange().getValues();
  if (CFG_MEMO_) CFG_MEMO_.v = v;
  return v;
}

function getConfig(perfilId) {
  const rows = configValores_();
  const headers = rows[0];
  const fila = rows.find(r => r[0] === perfilId);
  if (!fila) throw new Error('Perfil no encontrado: ' + perfilId);
  const cfg = {};
  headers.forEach((h, i) => cfg[h] = fila[i]);
  return cfg;
}

function listarPerfiles() {
  const config = configValores_();
  return config.slice(1).filter(r => r[0]).map(r => ({ id: r[0], nombre: r[1] }));
}

function getSheet(cfg) {
  return SpreadsheetApp.openById(cfg.sheet_id).getSheetByName(cfg.hoja_nombre);
}

function getDatosPerfil_(perfilId) {
  const cfg = getConfig(perfilId);
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const leer = (nombre) => {
    const sh = ss.getSheetByName(nombre);
    if (!sh) return [];
    const rows = sh.getDataRange().getValues();
    if (rows.length < 2) return [];
    const headers = rows[0];
    return rows.slice(1).filter(r => r[0] === perfilId).map(r => {
      const o = {};
      // google.script.run no devuelve fechas: si Sheets tomó "8.5" como 8 de mayo, se recupera el número
      headers.forEach((h, i) => {
        const v = r[i];
        o[h] = v instanceof Date
          ? (/^criterio_/.test(h) ? parseFloat(v.getDate() + '.' + (v.getMonth() + 1)) : Utilities.formatDate(v, 'GMT-3', 'dd/MM/yyyy'))
          : v;
      });
      return o;
    });
  };
  const heroArr = leer('Perfiles_Hero');
  const hero = heroArr[0] || {};
  return {
    perfil_id: perfilId,
    hero: hero,
    eval_items: leer('Perfiles_Screening_Eval').map(r => ({ id: r.id_eval, label: r.etiqueta, peso: Number(r.peso) || 1 })),
    val_items: leer('Perfiles_Screening_Val').map(r => ({
      id: r.id_val,
      label: r.etiqueta,
      opts: String(r.opciones || '').split('|').map(s => s.trim()).filter(Boolean)
    })),
    competencias: leer('Perfiles_Competencias').map(r => ({
      id: r.id_comp,
      nombre: r.nombre_comp,
      descripcion: r.descripcion,
      peso: Number(r.peso) || 1,
      preguntas: String(r.preguntas || '').split('|').map(s => s.trim()).filter(Boolean)
    })),
    datosInternos: String(cfg.datos_internos || '').toUpperCase() !== 'NO',
    entrevistadoras: String(hero.entrevistadoras || entrevistadoresDefault_().join(',')).split(',').map(s => s.trim()).filter(Boolean)
  };
}

// Un link sin «perfil» no cae en ninguna búsqueda por defecto: se avisa que está incompleto
function perfilObligatorio_(p) {
  const v = String(p || '').trim();
  if (!v) throw new Error('El link está incompleto: falta indicar la búsqueda.');
  return v;
}

function linkIncompleto_() {
  return HtmlService.createHtmlOutput('<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Link incompleto | Blumia</title><link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;600&display=swap" rel="stylesheet"></head><body style="margin:0;background:#F3F8F1;font-family:Poppins,sans-serif;color:#243D2C;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:24px;box-sizing:border-box"><div style="max-width:460px;text-align:center;background:#fff;border-radius:16px;padding:36px 28px;box-shadow:0 4px 18px rgba(36,61,44,.12)"><h2 style="margin:0 0 10px">Este link está incompleto</h2><p style="margin:0;line-height:1.5;font-size:14px">Falta indicar de qué búsqueda se trata. Pedí el link de nuevo a quien te lo envió o escribí a <b>empleos@blumiaconsulting.com</b>.</p></div></body></html>')
    .setTitle('Link incompleto | Blumia').addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doGetPagina_(e) {
  const tipo = e?.parameter?.tipo;
  const perfil = e?.parameter?.perfil || '';
  let template, titulo;
  const tipoQuePideBusqueda = ['competencias', 'postulacion', 'kanban', 'flyer', 'perfilpuesto', 'pdfcliente', 'informe', 'procesohtml'].indexOf(tipo) >= 0 || (!tipo && !!e?.parameter?.nombre);
  if (!perfil && tipoQuePideBusqueda) return linkIncompleto_();

  if (tipo === 'competencias') {
    template = HtmlService.createTemplateFromFile('Index');
    template.candidato = e.parameter.nombre || '';
    template.entrevistador = e.parameter.entrevistador || '';
    template.perfil = perfil;
    template.editar = e.parameter.editar === '1';
    titulo = 'Evaluación de Competencias | Blumia';
  } else if (tipo === 'postulacion') {
    let cerrada = false;
    try { cerrada = estadoBusqueda_(getConfig(perfil)) === 'CERRADA'; } catch (err) { }
    if (cerrada) {
      const h = getHeroPostulacion(perfil);
      return HtmlService.createHtmlOutput('<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Búsqueda cerrada | Blumia</title><link href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;600&display=swap" rel="stylesheet"></head><body style="margin:0;background:#F3F8F1;font-family:Poppins,sans-serif;color:#243D2C;display:flex;align-items:center;justify-content:center;min-height:100vh;padding:24px;box-sizing:border-box"><div style="max-width:460px;text-align:center;background:#fff;border-radius:16px;padding:36px 28px;box-shadow:0 4px 18px rgba(36,61,44,.12)"><h2 style="margin:0 0 10px">Esta búsqueda ya está cerrada</h2><p style="margin:0 0 6px;line-height:1.5">' + String(h.pos).replace(/[<>&]/g, '') + '</p><p style="margin:14px 0 0;line-height:1.5;font-size:14px">Gracias por tu interés. Podés seguir las próximas búsquedas en <b>blumiaconsulting.com</b>.</p></div></body></html>')
        .setTitle('Búsqueda cerrada | Blumia').addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    }
    template = HtmlService.createTemplateFromFile('Postulacion');
    const hero = getHeroPostulacion(perfil);
    template.perfil = perfil;
    template.pos = hero.pos;
    template.loc = hero.loc;
    template.logo = hero.logo;
    template.empresa = hero.empresa;
    template.base = JSON.stringify(preguntasBase_(getConfig(perfil)));
    titulo = 'Postulación · ' + hero.pos + ' | Blumia';
  } else if (tipo === 'pdfcliente') {
    const urlPdf = generarPDFClientes(perfil);
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:sans-serif;padding:40px;text-align:center">' +
      '<h2 style="color:#1A6645">PDF generado ✓</h2>' +
      '<p>El PDF con los informes está listo en la carpeta "Informes Clientes" de Drive.</p>' +
      '<a href="' + urlPdf + '" target="_blank" style="display:inline-block;background:#1A6645;color:#fff;padding:12px 24px;border-radius:4px;text-decoration:none;font-weight:bold">Abrir PDF</a>' +
      '</body></html>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (tipo === 'verinforme') {
    const contenido = DriveApp.getFileById(e.parameter.id).getBlob().getDataAsString();
    return HtmlService.createHtmlOutput(contenido)
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } else if (tipo === 'informe') {
    const url = generarInformeCandidato_(perfil, Number(e.parameter.fila));
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:sans-serif;padding:40px;text-align:center">' +
      '<h2 style="color:#1A6645">Informe generado ✓</h2>' +
      '<p>El informe de este candidato está listo.</p>' +
      '<a href="' + url + '" target="_blank" style="display:inline-block;background:#1A6645;color:#fff;padding:12px 24px;border-radius:4px;text-decoration:none;font-weight:bold">Abrir informe</a>' +
      '<p style="color:#888;margin-top:20px">También quedó el link "Ver informe" en la planilla.</p>' +
      '</body></html>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (tipo === 'perfilpuesto') {
    return perfilPuestoHTML(perfil, e.parameter.t, e.parameter.m);
  } else if (tipo === 'inicio') {
    return inicioHTML_();
  } else if (tipo === 'kanban') {
    return kanbanHTML(perfil);
  } else if (tipo === 'banco') {
    return bancoHTML(e.parameter.desde || '');
  } else if (tipo === 'flyer') {
    return flyerHTML_(perfil);
  } else if (tipo === 'procesohtml') {
    let empProceso = empresaDefault_();
    try { if (e.parameter.perfil) empProceso = empresaKey_(getConfig(e.parameter.perfil)); } catch (err) { }
    generarInformeProceso(empProceso);
    completarProcesoIA(empProceso);
    const url = generarHTMLProceso(empProceso);
    return HtmlService.createHtmlOutput('<!DOCTYPE html><html><body style="font-family:sans-serif;padding:40px;text-align:center"><h2 style="color:#1A6645">Informe de proceso listo</h2><a href="' + url + '" target="_blank" style="display:inline-block;background:#1A6645;color:#fff;padding:14px 28px;border-radius:6px;text-decoration:none;font-weight:800">Abrir informe</a></body></html>').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (!tipo && !e?.parameter?.perfil && !e?.parameter?.nombre) {
    return inicioHTML_(); // entrar al link base sin parámetros: pantalla de Inicio
  } else {
    template = HtmlService.createTemplateFromFile('Formulario');
    template.candidato = e.parameter.nombre || '';
    template.entrevistador = e.parameter.entrevistador || '';
    template.perfil = perfil;
    template.editar = e.parameter.editar === '1';
    titulo = 'Screening Blumia';
  }

  return template.evaluate()
    .setTitle(titulo)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function procesarTodo(perfilId) {
  const cfg = getConfig(perfilId);
  procesarMails(cfg);
  procesarDriveManual(cfg);
}

function procesarTodosLosPerfiles() {
  procesarTodo_v2();
}

function procesarMails(cfg) {
  const sheet = getSheet(cfg);
  const folderRecibidos = DriveApp.getFolderById(cfg.folder_recibidos);
 
  // label para no reprocesar
  const LABEL = 'Procesado-Blumia';
  let label = GmailApp.getUserLabelByName(LABEL);
  if (!label) label = GmailApp.createLabel(LABEL);
 
  // excluye los ya procesados con -label
  const threads = GmailApp.search(`to:${CORREO_DESTINO} newer_than:3d -label:${LABEL}`, 0, 50);
 
  threads.forEach(thread => {
    thread.getMessages().forEach(msg => {
      msg.getAttachments().forEach(att => {
        const mime = att.getContentType();
        if (mime.includes('pdf') || mime.includes('officedocument') || mime.includes('word')) {
          const file = folderRecibidos.createFile(att);
          const datos = procesarConChatGPT(file, cfg);
          if (datos && !esDuplicado(datos, sheet)) insertarFilaLibre(sheet, datos, file.getUrl(), msg.getDate(), 'Mail');
        }
      });
      const body = msg.getBody();
      const links = [...new Set([
        ...(body.match(/https:\/\/drive\.google\.com\/file\/d\/([a-zA-Z0-9_-]+)/g) || []),
        ...(body.match(/https:\/\/drive\.google\.com\/open\?id=([a-zA-Z0-9_-]+)/g) || [])
      ])];
      links.forEach(link => {
        try {
          const fileId = link.match(/[a-zA-Z0-9_-]{20,}/)[0];
          const orig = DriveApp.getFileById(fileId);
          const mime = orig.getMimeType();
          if (mime.includes('pdf') || mime.includes('officedocument') || mime.includes('word')) {
            const file = folderRecibidos.createFile(orig.getBlob()).setName(orig.getName());
            const datos = procesarConChatGPT(file, cfg);
            if (datos && !esDuplicado(datos, sheet)) insertarFilaLibre(sheet, datos, file.getUrl(), msg.getDate(), 'Mail');
          }
        } catch (e) { console.log('Error link ' + link); }
      });
    });
    thread.addLabel(label); // marca el thread como procesado
  });
}

/* ============================================================
   FIX: procesarDriveManual ahora escribe en la PRIMERA FILA LIBRE
   (antes usaba appendRow y caía después de las fórmulas, fila 251+).
   Reemplazá tu procesarDriveManual por esta.
   ============================================================ */

function procesarDriveManual(cfg) {
  const sheet = getSheet(cfg);
  const madre = DriveApp.getFolderById(cfg.folder_madre);
  const recibidos = DriveApp.getFolderById(cfg.folder_recibidos);
  const files = madre.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const datos = procesarConChatGPT(file, cfg);
    if (datos && !esDuplicado(datos, sheet)) {
      insertarFilaLibre(sheet, datos, file.getUrl(), file.getDateCreated(), 'CV directo');
    }
    file.moveTo(recibidos);
  }
}

function insertarFilaLibre(sheet, datos, url, fecha, fuente) {
  const colB = sheet.getRange(1, 2, Math.max(sheet.getLastRow(), 1), 1).getValues();
  let fila = -1;
  for (let i = 1; i < colB.length; i++) { if (!String(colB[i][0]).trim()) { fila = i + 1; break; } }
  if (fila === -1) fila = sheet.getLastRow() + 1;

  sheet.getRange(fila, 1, 1, 10).setValues([[
    fecha, datos.nombre, datos.telefono, datos.dni, datos.nivel_ajuste,
    datos.requisitos, datos.resumen, url, 'Nuevo', fuente || 'CV directo'
  ]]);
  if (datos.justificacion) sheet.getRange(fila, 5).setNote(datos.justificacion); // comentario nota
  sheet.getRange(fila, 22).setValue(datos.email || 'Sin dato');       // Correo
  ponerFormulaWhatsApp(sheet, fila);                                  // Contactar (21)
  registrarHistorial_(sheet.getName(), datos.nombre, 'Ingreso', 'Llegó por ' + (fuente || 'CV directo'), 'sistema');
  sheet.getRange(fila, 23).setValue(datos.formacion || 'Sin dato');   // Formación
  sheet.getRange(fila, 24).setValue(datos.zona || 'Sin dato');        // Zona
  sheet.getRange(fila, 27).setValue('Sin dato');                      // Posgrado
  sheet.getRange(fila, 28).setValue('Sin dato');                      // Excel
  sheet.getRange(fila, 29).setValue('Sin dato');                      // Power BI
  sheet.getRange(fila, 30).setValue('Sin dato');                      // IA
}

function esDuplicado(datos, sheet) {
  const data = sheet.getDataRange().getValues();
  const dniN = String(datos.dni || '').replace(/\D/g, '');
  const nombreN = String(datos.nombre || '').toLowerCase().trim();
  Logger.log('Hoja leída: ' + sheet.getName());
  Logger.log('Buscado=[' + nombreN + ' / ' + dniN + ']');
  if (data.length < 2) return false;
  for (let i = 1; i < data.length; i++) {
    const dniE = String(data[i][3] || '').replace(/\D/g, '');
    const nombreE = String(data[i][1] || '').toLowerCase().trim();
    if ((dniN && dniE === dniN) || (nombreN && nombreE === nombreN)) return true;
  }
  return false;
}


function procesarConChatGPT(file, cfg, textoYaExtraido) {
  try {
    let textoCV;
    if (textoYaExtraido && String(textoYaExtraido).trim().length >= 100) {
      // reutilizar texto extraído por el router → sin segunda conversión a Doc
      textoCV = String(textoYaExtraido).substring(0, 10000);
    } else {
      if (file.getMimeType() !== 'application/pdf') return null;
      const blob = file.getBlob();
      const tempFile = Drive.Files.create({ name: 'Temp_CV', mimeType: MimeType.GOOGLE_DOCS }, blob);
      textoCV = DocumentApp.openById(tempFile.id).getBody().getText().substring(0, 10000);
      Drive.Files.remove(tempFile.id);
    }
    if (textoCV.trim().length < 100) throw new Error('PDF vacío o corrupto');

    const promptSistema = `Sos un analista senior de headhunting. Analizá este CV para el siguiente perfil:

${cfg.prompt_extra || ''}

Devolvé SOLO JSON válido con este formato:
{
  "nombre": "Nombre Real", "email": "correo@ejemplo.com", "telefono": "Numero", "dni": "Solo numeros",
  "formacion": "Máximo nivel + título (ej: Universitario completo - Lic. en RRHH)",
  "localidad": "Ciudad/localidad textual del CV",
  "nivel_ajuste": "A/B+/B/C",
  "justificacion": "Una sola oración corta explicando el porqué de la nota",
  "requisitos": { "zona": "", "liderazgo": "", "formacion": "", "genero": "", "edad": "", "exp_rubro": "" },
  "resumen_experiencia": "3-4 oraciones narrativas"
}
REGLAS:
- email: extraé el correo del CV. Si no hay, dejá "".
- formacion: el máximo nivel educativo y título que figure en el CV. Si no hay, dejá "".
- En requisitos.zona poné la evaluación ("Zona OK" / "Fuera").
- localidad: la ciudad textual del CV (ej "Pilar", "CABA"). Si no figura, dejá "".
- nivel_ajuste (NOTA): A cumple todo sin ruido; B+ muy bueno con una cuestión menor/mediana; B cumple lo mínimo con dudas o borderline (reserva); C no cumple excluyente o bandera roja (descartado).
- justificacion: una sola oración directa explicando la nota, ej "B+: cumple experiencia pero sin convenio".`;

    const content = analizarJSONconIA_(cfg, promptSistema, textoCV);
    const req = content.requisitos || {};
    const sd = v => (v && String(v).trim()) ? String(v).trim() : 'Sin dato';
    return {
      nombre: String(content.nombre || 'No detectado'),
      email: sd(content.email),
      telefono: String(content.telefono || ''),
      dni: String(content.dni || ''),
      formacion: sd(content.formacion),
      zona: sd(content.localidad),
      nivel_ajuste: String(content.nivel_ajuste || 'C'),
      justificacion: String(content.justificacion || ''),
      requisitos: formatearRequisitos(req, content.localidad),
      resumen: String(content.resumen_experiencia || '')
    };
  } catch (e) {
    return { nombre: 'Error', email: 'Sin dato', formacion: 'Sin dato', zona: 'Sin dato', requisitos: 'Error', resumen: e.toString() };
  }
}

function procesarScreening_(data) {
  QUIEN_ = String(data.entrevistador || '');
  const cfg = getConfig(perfilObligatorio_(data.perfil));
  const sheet = getSheet(cfg);
  const rows = sheet.getDataRange().getValues();
  const norm = s => s.toString().toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ');
  const buscado = norm(data.nombre);

  let fila = -1;
  for (let i = 1; i < rows.length; i++) {
    if (norm(rows[i][1] || '') === buscado) { fila = i + 1; break; }
  }

  if (fila === -1) {
    const primeros = rows.slice(1, 6).map(r => r[1]).filter(Boolean).join(', ');
    throw new Error("No encontré a '" + data.nombre + "' en la pestaña '" + cfg.hoja_nombre + "'. Primeros: [" + (primeros || 'ninguno') + "]. Verificá perfil/pestaña.");
  }

  const notaAntesScr = String(sheet.getRange(fila, 12).getValue() || '').trim();
  sheet.getRange(fila, 26).setValue('LLEGUE resultado=' + data.resultado);
  const celdaL = sheet.getRange(fila, 12);
  celdaL.setDataValidation(null).setValue(data.resultado);
  const colores = { A: '#E6F2EC', B: '#FEF3D8', C: '#FDEAE4' };
  celdaL.setBackground(colores[data.resultado] || '#fff').setFontWeight('bold').setHorizontalAlignment('center');

  if (data.entrevistador) sheet.getRange(fila, 14).setValue(data.entrevistador);
  sheet.getRange(fila, 26).setValue('DEBUG resultado=' + data.resultado);
  registrarHistorial_(sheet.getName(), sheet.getRange(fila, 2).getValue(), data.editar ? 'Screening editado' : 'Screening', (data.editar ? (notaAntesScr || '—') + ' → ' + data.resultado : 'Formulario guardado · resultado ' + data.resultado) + (data.entrevistador ? ' · entrevistó ' + data.entrevistador : ''));
  if (!data.editar || aplicarEdicion_(sheet, fila, 'screening', notaAntesScr, data.resultado)) aplicarEstadoNota(sheet, fila, 'screening', data.resultado);
  guardarEstadoEval_(cfg, perfilObligatorio_(data.perfil), data.nombre, 'screening', data.estado, data.resultado, data.entrevistador);
  try {
    const folder = DriveApp.getFolderById(cfg.folder_pdf_screening);
    const blob = Utilities.newBlob(logoParaPdf_(data.htmlFull), 'text/html', 'Screening_' + data.nombre.replace(/\s+/g, '_') + '.html');
    const pdf = folder.createFile(blob.getAs('application/pdf'));
    const celdaM = sheet.getRange(fila, 13);
    celdaM.clearFormat();
    celdaM.setFormula('=HYPERLINK("' + pdf.getUrl() + '"; "Ver detalle")');
    celdaM.setFontLine('underline').setFontColor('#000');
    return 'OK';
  } catch (e) {
    throw new Error('Guardé el resultado, pero el PDF falló: ' + e.message);
  }
}
function llenarColumnaO(perfilId) {
  const cfg = getConfig(perfilId);
  const sheet = getSheet(cfg);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && !data[i][14]) {
      const url = URL_BASE() + '?tipo=competencias&perfil=' + perfilId + '&nombre=' + encodeURIComponent(data[i][1]) + '&entrevistador=' + encodeURIComponent(data[i][17] || '');
      sheet.getRange(i + 1, 15).setFormula('=HYPERLINK("' + url + '";"Evaluar competencias")');
    }
  }
}

/** El conversor HTML→PDF de Apps Script no siempre dibuja imágenes embebidas: el logo del encabezado se pide por URL pública (Cloudflare, azcubot/netlify). */
function logoParaPdf_(html) {
  return String(html || '').replace(/(<img[^>]*class="hero-logo"[^>]*?src=")data:image\/[a-z]+;base64,[^"]+(")/, '$1https://azcubot.pages.dev/blumia-logo-blanco.png$2');
}

function guardarEnPlanilla_(nombreCandidato, htmlContent, ajustePonderacion, perfilId, quien, estadoJSON, editar) {
  QUIEN_ = quien || '';
  try {
    console.log("Perfil recibido: " + perfilId);
    const cfg = getConfig(perfilObligatorio_(perfilId));
    const sheet = getSheet(cfg);
    const data = sheet.getDataRange().getValues();
    
    // Log para ver qué carajo está entrando
    console.log("Input recibido: '" + nombreCandidato + "' (largo: " + nombreCandidato.length + ")");
    
    // Normalización: quitamos acentos y normalizamos a NFD
    const norm = (s) => s.toString().toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
    
    const buscado = norm(nombreCandidato);
    let fila = -1;
    
    for (let i = 1; i < data.length; i++) {
      const nombreEnSheet = norm(data[i][1] || "");
      if (nombreEnSheet === buscado) {
        fila = i + 1;
        break;
      }
    }
    
    if (fila === -1) {
      console.error("ERROR: No hubo match.");
      console.error("El formulario envió: '" + nombreCandidato + "'");
      // Acá vemos qué hay en el Excel para comparar
      console.error("Ejemplo de lo que hay en Excel (primeras 5 filas): " + data.slice(1, 6).map(r => r[1]).join(", "));
      return 'Error: Candidato no encontrado. Revisá los logs de ejecución.';
    }

    // ... (el resto de tu código de guardado sigue igual) ...
    const folder = DriveApp.getFolderById(cfg.folder_pdf_competencias);
    console.log("htmlContent tipo: " + typeof htmlContent);
    console.log("htmlContent valor: " + htmlContent);
    const blob = Utilities.newBlob(logoParaPdf_(htmlContent), MimeType.HTML, 'temp.html');
    const pdf = folder.createFile(blob.getAs(MimeType.PDF).setName('Competencias_' + nombreCandidato + '.pdf'));
    
    const notaAntesComp = String(sheet.getRange(fila, 16).getValue() || '').trim();
    if (ajustePonderacion) sheet.getRange(fila, 16).setValue(ajustePonderacion);
    registrarHistorial_(sheet.getName(), sheet.getRange(fila, 2).getValue(), editar ? 'Competencias editado' : 'Competencias', editar ? (notaAntesComp || '—') + ' → ' + (ajustePonderacion || '—') : 'Formulario guardado' + (ajustePonderacion ? ' · resultado ' + ajustePonderacion : ''));
    if (ajustePonderacion && (!editar || aplicarEdicion_(sheet, fila, 'competencias', notaAntesComp, ajustePonderacion))) aplicarEstadoNota(sheet, fila, 'competencias', ajustePonderacion);
    guardarEstadoEval_(cfg, perfilObligatorio_(perfilId), nombreCandidato, 'competencias', estadoJSON, ajustePonderacion, quien);
    sheet.getRange(fila, 17).setFormula('=HYPERLINK("' + pdf.getUrl() + '"; "Ver PDF")');
    
    return 'OK';
  } catch (e) {
    return 'Error: ' + e.toString();
  }
}

function onEditDescartado(e) {
  if (!e || !e.range) return;
  const sh = e.range.getSheet();
  const row = e.range.getRow();
  const col = e.range.getColumn();
  if (row === 1) return;
  if (String(sh.getRange(row, 9).getValue() || '').trim() === 'Contratado') return;

  const NOTA_COLS = { 5: 'cv', 12: 'screening', 16: 'competencias', 20: 'empresa' };
  if (NOTA_COLS[col]) {
    let v = String(e.range.getValue() || '').trim().toUpperCase();
    const m = { 'ALTO': 'A', 'MEDIO': 'B', 'BAJO': 'C' };
    if (m[v]) v = m[v];
    if (['A','B+','B','C'].indexOf(v) === -1) return;
    aplicarEstadoNota(sh, row, NOTA_COLS[col], v);
    return;
  }

  // Si se edita el ESTADO (col 9) a mano y se pone "Avanza a empresa", poner botón informe
  if (col === 9) {
    const estado = String(e.range.getValue() || '').trim();
    try { if (String(e.oldValue || '') !== estado) registrarHistorial_(sh.getName(), sh.getRange(row, 2).getValue(), 'Estado', (e.oldValue || '—') + ' → ' + estado + ' (editado en la planilla)', (e.user && e.user.getEmail && e.user.getEmail()) || ''); } catch (err) { }
    if (estado === 'Avanza a empresa') {
      try {
        const yaTiene = sh.getRange(row, 32).getRichTextValue();
        if (!yaTiene || !yaTiene.getLinkUrl()) {
          const nombreHoja = sh.getName();
          let perfilId = '';
          listarPerfiles().forEach(p => {
            try { if (getConfig(p.id).hoja_nombre === nombreHoja) perfilId = p.id; } catch(e){}
          });
          if (perfilId) {
            const url = URL_WEBAPP + '?tipo=informe&perfil=' + encodeURIComponent(perfilId) + '&fila=' + row;
            const rich = SpreadsheetApp.newRichTextValue().setText('📄 Generar informe').setLinkUrl(url).build();
            sh.getRange(row, 32).setRichTextValue(rich);
          }
        }
      } catch(err) {}
    }
  }
}

function aplicarEstadoNota(sheet, fila, etapa, nota) {
  let v = String(nota || '').trim().toUpperCase();
  const m = { 'ALTO': 'A', 'MEDIO': 'B', 'BAJO': 'C' };
  if (m[v]) v = m[v];
  if (['A','B+','B','C'].indexOf(v) === -1) return;
  if (String(sheet.getRange(fila, 9).getValue() || '').trim() === 'Contratado') return;

  let estado;
  if (v === 'C') {
    estado = { cv:'Descartado sin screening', screening:'Descartado post screening', competencias:'Descartado en competencias', empresa:'Descartado en empresa' }[etapa];
  } else if (v === 'B' && etapa !== 'cv' && etapa !== 'screening') {
    estado = (etapa === 'competencias') ? 'Reserva competencias' : 'En reserva';
  } else {
    // A, B+ siempre avanzan; B en CV también avanza
    estado = { cv:'Avanza a screening', screening:'Avanza a competencias', competencias:'Avanza a screening empresa', empresa:'En cliente' }[etapa];
  }

  let bg = '#D9D9D9', fc = '#333333';
  if (v === 'A') { bg = '#1A6645'; fc = '#FFFFFF'; }
  else if (v === 'B+') { bg = '#B7E0C6'; fc = '#14502F'; }
  else if (v === 'B') { bg = '#F7DC6F'; fc = '#6B5800'; }
  else if (v === 'C') { bg = '#E06666'; fc = '#FFFFFF'; }

  const celda = sheet.getRange(fila, 9);
  const regla = celda.getDataValidation();
  const estadoAntes = String(celda.getValue() || '').trim();
  celda.setDataValidation(null).setValue(estado);
  registrarHistorial_(sheet.getName(), sheet.getRange(fila, 2).getValue(), 'Nota ' + etapa, v + ' · ' + (estadoAntes || '—') + ' → ' + estado);
  if (regla) celda.setDataValidation(regla);
  celda.setBackground(bg).setFontColor(fc).setFontWeight('bold').setHorizontalAlignment('center');

  // ENGANCHE: si avanza a empresa, poner botón "Generar informe" en col 32
  if (estado === 'Avanza a empresa') {
    try {
      const yaTiene = sheet.getRange(fila, 32).getRichTextValue();
      if (!yaTiene || !yaTiene.getLinkUrl()) {
        const nombreHoja = sheet.getName();
        let perfilId = '';
        listarPerfiles().forEach(p => {
          try { if (getConfig(p.id).hoja_nombre === nombreHoja) perfilId = p.id; } catch(e){}
        });
        if (perfilId) {
          const url = URL_WEBAPP + '?tipo=informe&perfil=' + encodeURIComponent(perfilId) + '&fila=' + fila;
          const rich = SpreadsheetApp.newRichTextValue().setText('📄 Generar informe').setLinkUrl(url).build();
          sheet.getRange(fila, 32).setRichTextValue(rich);
        }
      }
    } catch(e) {}
  }
}

function crearTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  ScriptApp.newTrigger('syncCalendarConfirmados').timeBased().everyMinutes(15).create();
  ScriptApp.newTrigger('procesarTodosLosPerfiles').timeBased().everyMinutes(30).create();
  ScriptApp.newTrigger('onEditDescartado').forSpreadsheet(ss).onEdit().create();
  console.log('Triggers OK');
}
function getPreguntasPostulacion(perfilId) {
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Postulacion_Preguntas');
  if (!sh) return [];
  const rows = sh.getDataRange().getValues();
  if (rows.length < 2) return [];
  // headers: perfil_id, id_preg, etiqueta, tipo, opciones, requerido
  return rows.slice(1)
    .filter(r => r[0] === perfilId && r[1])
    .map(r => ({
      id: String(r[1]),
      etiqueta: String(r[2] || ''),
      tipo: String(r[3] || 'texto').toLowerCase(),
      opciones: String(r[4] || ''),
      requerido: String(r[5] || '').toUpperCase()
    }));
}
function recibirPostulacion(payload) {
  try {
    const cfg = getConfig(perfilObligatorio_(payload.perfil));
    if (estadoBusqueda_(cfg) === 'CERRADA') throw new Error('BUSQUEDA_CERRADA');
    const sheet = getSheet(cfg);

    const d = payload.datos || {};
    if (esDuplicado(d, sheet)) return 'DUPLICADO';

    if (!payload.cvBase64 || String(payload.cvBase64).length > 14.5 * 1024 * 1024) throw new Error('El CV falta o pesa más de 10 MB.'); // control del servidor (el formulario ya lo valida)
    if (!/\.(pdf|doc|docx)$/i.test(String(payload.cvNombre || ''))) throw new Error('El CV tiene que ser PDF o Word.');
    const mime = payload.cvTipo || MimeType.PDF;
    const blob = Utilities.newBlob(Utilities.base64Decode(payload.cvBase64), mime, payload.cvNombre || 'CV');
    const file = DriveApp.getFolderById(cfg.folder_recibidos).createFile(blob);
    const url = file.getUrl();

    const formacion = [d.nivelFormacion, d.estadoFormacion].filter(Boolean).join(' ') +
                      (d.tituloFormacion ? ' — ' + d.tituloFormacion : '');
    const posgrado = d.posgradoEstado ? d.posgradoEstado + (d.posgradoCual ? ' — ' + d.posgradoCual : '') : 'Sin posgrado';
    const ia = (d.ia || '') + (d.iaTools ? ' — ' + d.iaTools : '');

    const colB = sheet.getRange(1, 2, Math.max(sheet.getLastRow(), 1), 1).getValues();
    let fila = -1;
    for (let i = 1; i < colB.length; i++) { if (!String(colB[i][0]).trim()) { fila = i + 1; break; } }
    if (fila === -1) fila = sheet.getLastRow() + 1;

    sheet.getRange(fila, 1, 1, 10).setValues([[
      new Date(), d.nombre || '', d.telefono || '', d.dni || '',
      '', 'Procesando…', '', url, 'Pendiente', 'Postulación'
    ]]);
    sheet.getRange(fila, 22, 1, 9).setValues([[
      d.email || '', formacion, '', 0,
      JSON.stringify({ perfil: perfilObligatorio_(payload.perfil), datos: d, declaradas: payload.declaradas || {}, cvId: file.getId() }),
      posgrado, d.excel || '', d.powerbi || '', ia
    ]]);
    ponerFormulaWhatsApp(sheet, fila);
    registrarHistorial_(sheet.getName(), d.nombre || '', 'Ingreso', 'Llegó por Postulación (formulario)', 'sistema');

    return 'OK';

  } catch (err) {
    if (String(err && err.message) === 'BUSQUEDA_CERRADA') throw new Error('Esta búsqueda ya está cerrada.');
    registrarErrorPostulacion(payload, err);
    throw new Error('No se pudo procesar la postulación'); // dispara el mensaje amable en el form
  }
}

function registrarErrorPostulacion(payload, err) {
  const d = (payload && payload.datos) || {};
  const detalle = String(err);
  const perfil = (payload && payload.perfil) || 'desconocido';

  // 1) hoja "Errores"
  try {
    const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
    let sh = ss.getSheetByName('Errores');
    if (!sh) {
      sh = ss.insertSheet('Errores');
      sh.appendRow(['Fecha','Perfil','Nombre','Email','Teléfono','Error']);
      sh.getRange(1,1,1,6).setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
    }
    sh.appendRow([new Date(), perfil, d.nombre || '', d.email || '', d.telefono || '', detalle]);
  } catch (e) {}

  // 2) mail de aviso
  try {
    MailApp.sendEmail({
      to: CORREO_DESTINO,
      subject: '⚠️ Falló una postulación — ' + perfil,
      body: 'No se pudo procesar una postulación.\n\n' +
            'Perfil: ' + perfil + '\n' +
            'Nombre: ' + (d.nombre || '(sin dato)') + '\n' +
            'Email: ' + (d.email || '(sin dato)') + '\n' +
            'Teléfono: ' + (d.telefono || '(sin dato)') + '\n\n' +
            'Error: ' + detalle + '\n\n' +
            'El candidato fue invitado a escribir a empleos@blumiaconsulting.com.'
    });
  } catch (e) {}
}
function getHeroPostulacion(perfilId) {
  try {
    const rows = SpreadsheetApp.openById(SHEET_MASTER_ID)
      .getSheetByName('Perfiles_Hero').getDataRange().getValues();
    const headers = rows[0];
    const fila = rows.slice(1).find(r => r[0] === perfilId);
    if (!fila) return { pos: 'Búsqueda', loc: '', logo: '', empresa: '' };
    const o = {};
    headers.forEach((h, i) => o[h] = fila[i]);
    const loc = [o.empresa, o.ubicacion].filter(Boolean).join(' · ');
    let logo = '', empresa = String(o.empresa || '');
    try {
      const cfg = getConfig(perfilId);
      logo = logoCliente_(cfg);
      if (!empresa) empresa = empresaKey_(cfg);
    } catch (e) { }
    return { pos: String(o.posicion || 'Búsqueda'), loc: loc, logo: logo, empresa: empresa };
  } catch (err) {
    return { pos: 'Búsqueda', loc: '', logo: '', empresa: '' };
  }
}

// Logo claro del cliente para el encabezado verde de la postulación: Config.logo_cliente (link https a un PNG blanco con fondo transparente).
// Una empresa de EMPRESAS_CONOCIDAS_ (MACS) que todavía no lo tiene cargado usa el logo publicado en la web de Azcu.
function logoCliente_(cfg) {
  const l = (cfg && typeof cfg.logo_cliente === 'string') ? cfg.logo_cliente.trim() : '';
  if (/^https:\/\//i.test(l)) return l;
  return logoConocido_(empresaKey_(cfg));
}

// Preguntas fijas de la postulación que se muestran en cada búsqueda (Config.preguntas_base, separadas con coma; «ninguna» = ninguna).
// Una búsqueda sin el dato (las anteriores) sigue mostrando todas; las nuevas arrancan sin ninguna.
const PREG_BASE_ = ['excel', 'powerbi', 'ingles', 'ia', 'fulltime', 'presencial'];
function preguntasBase_(cfg) {
  const v = (cfg && typeof cfg.preguntas_base === 'string') ? cfg.preguntas_base.trim().toLowerCase() : '';
  if (!v) return PREG_BASE_.slice();
  return v.split(',').map(x => x.trim()).filter(x => PREG_BASE_.indexOf(x) >= 0);
}
const URL_WEBAPP = 'https://script.google.com/macros/s/AKfycbytekPKN31bZLxczDVknMr433K9KflzmrR-dCWsC1hLlx2MO7tbw6SrMjFXa1PnucQo/exec';

function generarTodosLosLinks() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let sh = ss.getSheetByName('Links');
  if (!sh) sh = ss.insertSheet('Links');
  sh.clear();

  const headers = ['Perfil', 'Postulación (público)', 'Screening', 'Competencias'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);

  const perfiles = listarPerfiles();
  let fila = 2;
  perfiles.forEach(p => {
    const base = URL_WEBAPP + '?';
    const post = base + 'tipo=postulacion&perfil=' + p.id;
    const scr  = base + 'perfil=' + p.id;
    const comp = base + 'tipo=competencias&perfil=' + p.id;
    sh.getRange(fila, 1).setValue(p.nombre);
    sh.getRange(fila, 2).setFormula('=HYPERLINK("' + post + '";"Abrir postulación")');
    sh.getRange(fila, 3).setFormula('=HYPERLINK("' + scr  + '";"Abrir screening")');
    sh.getRange(fila, 4).setFormula('=HYPERLINK("' + comp + '";"Abrir competencias")');
    fila++;
  });

  sh.autoResizeColumns(1, headers.length);
  Logger.log('Tablero "Links" generado con ' + perfiles.length + ' perfiles.');
}
function procesarPendientes() {
  const MAX_POR_CORRIDA = 15;
  let procesados = 0;

  const perfiles = listarPerfiles();
  for (const p of perfiles) {
    if (procesados >= MAX_POR_CORRIDA) break;
    let cfg;
    try { cfg = getConfig(p.id); } catch (e) { continue; }
    const sheet = getSheet(cfg);
    const last = sheet.getLastRow();
    if (last < 2) continue;

    const estados = sheet.getRange(2, 9, last - 1, 1).getValues(); // col I
    for (let r = 0; r < estados.length; r++) {
      if (procesados >= MAX_POR_CORRIDA) break;
      if (String(estados[r][0]).trim() !== 'Pendiente') continue;

      const fila = r + 2;
      const metaRaw = sheet.getRange(fila, 26).getValue();
      let meta;
      try { meta = JSON.parse(metaRaw); } catch (e) {
        sheet.getRange(fila, 9).setValue('Revisar');
        continue;
      }

      let intentos = Number(sheet.getRange(fila, 25).getValue()) || 0;
      intentos++;
      sheet.getRange(fila, 25).setValue(intentos);

      try {
        const d = meta.datos || {};
        const formacion = [d.nivelFormacion, d.estadoFormacion].filter(Boolean).join(' ') +
                          (d.tituloFormacion ? ' — ' + d.tituloFormacion : '');
        const lineas = [
          'Nombre: ' + (d.nombre || ''), 'Email: ' + (d.email || ''),
          'Teléfono: ' + (d.telefono || ''), 'DNI: ' + (d.dni || ''),
          'Formación: ' + formacion,
          'Posgrado: ' + (d.posgradoEstado ? d.posgradoEstado + (d.posgradoCual ? ' (' + d.posgradoCual + ')' : '') : 'Sin posgrado')
        ];
        const declaradasBase = [
          d.excel ? 'Excel: ' + d.excel : '',
          d.powerbi ? 'Power BI: ' + d.powerbi : '',
          d.ingles ? 'Inglés: ' + d.ingles : '',
          d.ia ? 'Herramientas de IA: ' + d.ia + (d.iaTools ? ' — usa: ' + d.iaTools : '') : '',
          'Remuneración pretendida: ' + (d.remuneracion || ''),
          'Localidad: ' + (d.localidad || ''),
          'Movilidad propia: ' + (d.movilidad || ''),
          d.fulltime ? 'Disponibilidad full-time: ' + d.fulltime : '',
          d.presencial ? 'Trabajo 100% presencial: ' + d.presencial : ''
        ].filter(Boolean);
        declaradasBase.forEach(l => lineas.push(l));
        const decl = meta.declaradas || {};
        getPreguntasPostulacion(meta.perfil).forEach(q => {
          if (decl[q.id] != null && decl[q.id] !== '') lineas.push(q.etiqueta + ': ' + decl[q.id]);
        });

        const textoCV = textoDeCV(meta.cvId);
        const gpt = analizarPostulacionGPT(textoCV, lineas.join('\n'), cfg);
        if (!gpt.ok) throw new Error(gpt.error || 'GPT falló');

        const celdaNota = sheet.getRange(fila, 5);
        celdaNota.setValue(gpt.nivel_ajuste);                                                 // E nota
        if (gpt.justificacion) celdaNota.setNote(gpt.justificacion);                          // comentario de celda
        sheet.getRange(fila, 6).setValue(gpt.requisitos + ' || DECLARADO: ' + lineas.slice(6).join(' · ')); // F
        sheet.getRange(fila, 7).setValue(gpt.resumen);                                        // G
        sheet.getRange(fila, 24).setValue(gpt.zona);                                          // Zona
        sheet.getRange(fila, 26).clearContent();                                              // limpia meta
        try { sheet.getRange(fila, 31).setFormula('=HYPERLINK("' + generarPDFPostulacion(d, decl, cfg, meta.perfil) + '";"Ver postulación")'); } catch(e) {}
        sheet.getRange(fila, 9).setValue('Nuevo');                                            // Estado queda en Nuevo
        procesados++;
        Utilities.sleep(1500);
      } catch (err) {
        if (intentos >= 3) {
          sheet.getRange(fila, 9).setValue('Revisar');
          sheet.getRange(fila, 6).setValue('Falló el análisis tras 3 intentos: ' + err);
        }
        Utilities.sleep(1500);
      }
    }
  }
  Logger.log('Pendientes procesados en esta corrida: ' + procesados);
}

function textoDeCV(fileId) {
  const blob = DriveApp.getFileById(fileId).getBlob();
  const temp = Drive.Files.create({ name: 'Temp_CV', mimeType: MimeType.GOOGLE_DOCS }, blob);
  const t = DocumentApp.openById(temp.id).getBody().getText().substring(0, 10000);
  Drive.Files.remove(temp.id);
  if (t.trim().length < 50) throw new Error('CV vacío o ilegible');
  return t;
}
function analizarPostulacionGPT(textoCV, bloqueDeclarado, cfg) {
  try {
    const promptSistema = `Sos un analista senior de headhunting. Analizá esta postulación para el siguiente perfil:

${cfg.prompt_extra || ''}

Tenés DOS fuentes: (A) los datos DECLARADOS por el candidato en el formulario, y (B) el texto del CV.
Cruzá ambas. Si el candidato declara algo que el CV no respalda o contradice, reflejalo en la nota y en requisitos.

=== DATOS DECLARADOS (verdad declarada por el candidato, no inventar) ===
${bloqueDeclarado}
=== FIN DECLARADOS ===

Devolvé SOLO JSON válido con este formato:
{
  "nombre": "Nombre Real", "telefono": "Numero", "dni": "Solo numeros", "localidad": "Ciudad/localidad textual declarada o del CV",
  "nivel_ajuste": "A/B+/B/C",
  "justificacion": "Una sola oración corta explicando el porqué de la nota",
  "requisitos": { "zona": "", "liderazgo": "", "formacion": "", "genero": "", "edad": "", "exp_rubro": "" },
  "resumen_experiencia": "3-4 oraciones narrativas, atributos/actitud, sin repetir el checklist"
}
REGLAS DE LA NOTA (nivel_ajuste):
- A: cumple todo, perfil potente sin ruido.
- B+: muy bueno, cumple lo excluyente, con una cuestión menor o mediana que hace ruido.
- B: cumple lo mínimo pero con dudas relevantes, o borderline con factores compensatorios (reserva).
- C: no cumple un requisito excluyente, o cae en bandera roja clara (descartado).
- justificacion: una sola oración directa, ej "B+: cumple experiencia y Excel, pero sin experiencia en convenio".
- zona: "Zona OK" si la localidad declarada entra en las zonas del perfil, "Fuera" si no.`;

    const content = analizarJSONconIA_(cfg, promptSistema, 'CV:\n' + textoCV);
    const req = content.requisitos || {};
    return {
      ok: true,
      nombre: String(content.nombre || ''),
      telefono: String(content.telefono || ''),
      dni: String(content.dni || ''),
      nivel_ajuste: String(content.nivel_ajuste || 'C'),
      justificacion: String(content.justificacion || ''),
      zona: String(content.localidad || req.zona || 'N/D'),
      requisitos: formatearRequisitos(req, content.localidad),
      resumen: String(content.resumen_experiencia || '')
    };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}
function setupPostulacionAsync() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const estados = ['Pendiente', 'Nuevo', 'Contactado', 'Entrevista agendada', 'Entrevistado',
                   'Avanza', 'En evaluación cliente', 'Oferta', 'Contratado', 'Descartado', 'Revisar'];
  const reglaEstado = SpreadsheetApp.newDataValidation().requireValueInList(estados, true).build();

  listarPerfiles().forEach(p => {
    let cfg; try { cfg = getConfig(p.id); } catch (e) { return; }
    const sh = ss.getSheetByName(cfg.hoja_nombre);
    if (!sh) { Logger.log('No existe hoja: ' + cfg.hoja_nombre); return; }

    // headers 23-26
    const h = sh.getRange(1, 23, 1, 4).getValues()[0];
    const setH = (col, txt) => sh.getRange(1, col).setValue(txt)
      .setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
    if (h[0] !== 'Formación declarada') setH(23, 'Formación declarada');
    if (h[1] !== 'Zona') setH(24, 'Zona');
    if (h[2] !== 'Intentos') setH(25, 'Intentos');
    if (h[3] !== '_meta') setH(26, '_meta');

    // validación de Estado (col 9) ampliada, filas 2-250
    sh.getRange(2, 9, 249, 1).setDataValidation(reglaEstado);

    // extender fórmulas de links a 250 filas (col 11 screening, col 15 competencias)
    for (let i = 2; i <= 250; i++) {
      sh.getRange(i, 11).setFormula(
        '=IF(B' + i + '<>"";HYPERLINK("' + URL_DEPLOY + '?perfil=' + p.id +
        '&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(N' + i + ');"Iniciar Screening");"")');
      sh.getRange(i, 15).setFormula(
        '=IF(B' + i + '<>"";HYPERLINK("' + URL_DEPLOY + '?tipo=competencias&perfil=' + p.id +
        '&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(R' + i + ');"Evaluar competencias");"")');
    }
    Logger.log('Configurada: ' + cfg.hoja_nombre);
  });

  // trigger cada 5 min (sin duplicar)
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'procesarPendientes') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('procesarPendientes').timeBased().everyMinutes(5).create();

  Logger.log('Setup async completo: estados, headers 23-26, fórmulas a 250, trigger 5 min.');
}
function ponerFormulaWhatsApp(sheet, fila) {
  // limpia el teléfono (solo dígitos), saca 0 y 15 iniciales, antepone 549
  const f =
    '=IF(C' + fila + '<>"";HYPERLINK("https://wa.me/549"&' +
    'REGEXREPLACE(REGEXREPLACE(SUBSTITUTE(SUBSTITUTE(C' + fila + ';"+54";"");"+";"");"[^0-9]";"");"^0";"");' +
    '"Abrir WhatsApp");"")';
  sheet.getRange(fila, 21).setFormula(f);
}
function formatearRequisitos(req, localidad) {
  if (!req || typeof req !== 'object') return 'Error';
  const labels = { zona: 'Zona', liderazgo: 'Liderazgo', formacion: 'Formación', genero: 'Género', edad: 'Edad', exp_rubro: 'Exp. rubro' };
  return Object.keys(labels).map(k => {
    let val = req[k] || 'N/D';
    if (k === 'zona' && localidad && String(localidad).trim()) val = val + ' (' + String(localidad).trim() + ')';
    return `${labels[k]}: ${val}`;
  }).join(' | ');
}
/* ============================================================
   ROUTER DE MAIL MULTI-PERFIL (opción A)
   Reemplaza el barrido por-perfil de procesarMails(cfg).
   NO llamar más a procesarMails(cfg) en un loop de perfiles:
   correr SOLO procesarMailsRouter().
   ============================================================ */

// --- perfiles activos leídos de Config (excluye operaciones / vacíos) ---
function perfilesActivos_() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const cfg = ss.getSheetByName('Config').getDataRange().getValues();
  const cEst = cfg[0].indexOf('estado_busqueda');
  const out = [];
  for (let i = 1; i < cfg.length; i++) {
    const id = cfg[i][0];
    if (!id || id === 'operaciones') continue;
    if (cEst >= 0 && String(cfg[i][cEst] || '').toUpperCase().trim() === 'CERRADA') continue; // búsqueda cerrada: no recibe CVs por mail ni por carpeta
    out.push({ id: id, nombre: cfg[i][1], prompt: cfg[i][8] || '' });
  }
  return out;
}

function rutearPorAsunto_(asunto) {
  if (!asunto) return null;
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const sh = ss.getSheetByName('Mail_Routing');
  if (!sh) return null;
  const a = asunto.toLowerCase();
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const pid = data[i][0];
    const claves = String(data[i][1] || '').toLowerCase().split('|').map(s => s.trim()).filter(Boolean);
    if (claves.some(k => a.indexOf(k) !== -1) && perfilesActivos_().some(p => p.id === pid)) return pid;
  }
  // sin fila en Mail_Routing: si el asunto nombra el puesto (o su id), alcanza
  const norm = t => String(t || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, ' ').trim();
  const an = ' ' + norm(asunto) + ' ';
  const hit = perfilesActivos_().filter(p => {
    const n = norm(p.nombre), i = norm(p.id);
    return (n.length > 3 && an.indexOf(' ' + n + ' ') !== -1) || (i.length > 3 && an.indexOf(' ' + i + ' ') !== -1);
  });
  return hit.length === 1 ? hit[0].id : null;
}

// --- etiqueta de Gmail por búsqueda: "Blumia/<puesto>" (o "Blumia/Sin clasificar") ---
function etiquetarHilo_(thread, perfilId) {
  try {
    const nombre = perfilId ? ((perfilesActivos_().filter(p => p.id === perfilId)[0] || {}).nombre || perfilId) : 'Sin clasificar';
    const nom = 'Blumia/' + String(nombre).replace(/\//g, '-');
    thread.addLabel(GmailApp.getUserLabelByName(nom) || GmailApp.createLabel(nom));
  } catch (e) { }
}

// --- extrae texto del CV (PDF / Google Doc / Word) ---
function extraerTextoCV_(file) {
  try {
    const mime = file.getMimeType();
    if (mime === 'application/vnd.google-apps.document') {
      return DocumentApp.openById(file.getId()).getBody().getText();
    }
    const blob = file.getBlob();
    const tempFile = Drive.Files.create({ name: 'Temp_CV_router', mimeType: MimeType.GOOGLE_DOCS }, blob);
    const texto = DocumentApp.openById(tempFile.id).getBody().getText();
    Drive.Files.remove(tempFile.id);
    return texto;
  } catch (e) {
    Logger.log('extraerTextoCV_ falló: ' + e.message);
    return '';
  }
}

// --- clasificador GPT independiente (llamada liviana) ---
function clasificarPerfilGPT(textoCV) {
  const perfiles = perfilesActivos_();
  if (!perfiles.length || !textoCV) return null;

  const lista = perfiles.map(p =>
    `- ${p.id}: ${p.nombre}. ${p.prompt.split('\n')[0]}`
  ).join('\n');

  const promptSistema =
`Sos un clasificador de CVs para búsquedas laborales. Te paso un CV y una lista de perfiles disponibles.
Decidí a qué perfil corresponde mejor el CV según experiencia y formación.
PERFILES:
${lista}

Respondé SOLO JSON: {"perfil_id":"<id exacto de la lista o null>","confianza":"alta|media|baja"}
Si el CV no encaja claramente con ninguno, devolvé perfil_id null y confianza baja.`;

  try {
    const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: { Authorization: 'Bearer ' + OPENAI_API_KEY },
      contentType: 'application/json',
      payload: JSON.stringify({
        model: 'gpt-4o-mini',
        response_format: { type: 'json_object' },
        messages: [
          { role: 'system', content: promptSistema },
          { role: 'user', content: textoCV.slice(0, 2000) }
        ]
      }),
      muteHttpExceptions: true
    });
    const parsed = JSON.parse(response.getContentText());
    if (parsed.error) throw new Error('OpenAI: ' + parsed.error.message);
    const content = JSON.parse(parsed.choices[0].message.content);
    const ids = perfiles.map(p => p.id);
    if (content.perfil_id && ids.indexOf(content.perfil_id) !== -1 && content.confianza !== 'baja') {
      return content.perfil_id;
    }
    return null;
  } catch (e) {
    Logger.log('clasificarPerfilGPT falló: ' + e.message);
    return null;
  }
}

// --- hoja de descartes de ruteo ---
function hojaSinClasificar_() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let sh = ss.getSheetByName('Sin clasificar');
  if (!sh) {
    sh = ss.insertSheet('Sin clasificar');
    sh.getRange(1, 1, 1, 5).setValues([['Fecha', 'Asunto', 'Archivo', 'Link CV', 'Motivo']]);
  }
  return sh;
}

// --- ROUTER: barre la casilla UNA vez y decide el perfil por CV ---
function procesarMailsRouter() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const LABEL = 'Procesado-Blumia';
  let label = GmailApp.getUserLabelByName(LABEL);
  if (!label) label = GmailApp.createLabel(LABEL);

  const threads = GmailApp.search(`to:${CORREO_DESTINO} newer_than:3d -label:${LABEL}`, 0, 50);
  let ok = 0, sin = 0;

  threads.forEach(thread => {
    const msgs = thread.getMessages();
    msgs.forEach(msg => {
      const asunto = msg.getSubject();
      msg.getAttachments().forEach(att => {
        const mime = att.getContentType();
        if (!(mime.includes('pdf') || mime.includes('officedocument') || mime.includes('word'))) return;

        // 1) intentar por asunto (gratis)
        let perfil = rutearPorAsunto_(asunto);

        // guardamos el archivo en Drive una sola vez (en carpeta del perfil si ya lo sabemos, si no temporal en raíz)
        let cfgTmp = perfil ? getConfig(perfil) : null;
        const folderInicial = cfgTmp
          ? DriveApp.getFolderById(cfgTmp.folder_recibidos)
          : DriveApp.getRootFolder();
        const file = folderInicial.createFile(att);

        // extraemos el texto UNA sola vez (se reutiliza para clasificar y analizar)
        const textoCV = extraerTextoCV_(file);

        // 2) si el asunto no resolvió, clasificar con GPT (usa solo los primeros 2000 chars)
        if (!perfil) {
          perfil = clasificarPerfilGPT(textoCV);
        }

        // 3) sin clasificar → hoja aparte
        etiquetarHilo_(thread, perfil);
        if (!perfil) {
          hojaSinClasificar_().appendRow([msg.getDate(), asunto, att.getName(), file.getUrl(), 'GPT no pudo clasificar / asunto sin pista']);
          sin++;
          return;
        }

        // 4) procesar al perfil resuelto con TUS funciones (texto reutilizado → sin 2ª conversión)
        const cfg = getConfig(perfil);
        const sheet = getSheet(cfg);
        const datos = procesarConChatGPT(file, cfg, textoCV);
        if (datos && !esDuplicado(datos, sheet)) {
          insertarFilaLibre(sheet, datos, file.getUrl(), msg.getDate(), 'Mail');
          ok++;
        }
      });
    });
    thread.addLabel(label);
  });

  Logger.log(`✅ Router: ${ok} ruteados y procesados · ${sin} a "Sin clasificar".`);
}
function procesarTodo_v2() {
  asegurarCopiaDiaria_();
  // 1) MAIL — barrido único con router (decide perfil solo)
  try {
    procesarMailsRouter();
  } catch (e) {
    Logger.log('⚠️ Router de mail falló: ' + e.message);
  }

  // 2) DRIVE — barrido por perfil (el perfil ya está dado por la carpeta)
  perfilesActivos_().forEach(p => {
    try {
      const cfg = getConfig(p.id);
      procesarDriveManual(cfg);
      Logger.log('✅ Drive procesado: ' + p.id);
    } catch (e) {
      Logger.log('⚠️ Drive falló en ' + p.id + ': ' + e.message);
    }
  });

  Logger.log('LISTO — procesarTodo_v2 (mail + drive) completado.');
}
function generarPDFPostulacion(d, decl, cfg, perfil) {
  const preguntas = getPreguntasPostulacion(perfil);
  let dyn = '';
  preguntas.forEach(q => {
    if (decl[q.id] != null && decl[q.id] !== '') {
      dyn += `<tr><td class="k">${q.etiqueta}</td><td class="v">${decl[q.id]}</td></tr>`;
    }
  });
  const formacion = [d.nivelFormacion, d.estadoFormacion].filter(Boolean).join(' ') +
                    (d.tituloFormacion ? ' — ' + d.tituloFormacion : '');
  const posgrado = d.posgradoEstado ? d.posgradoEstado + (d.posgradoCual ? ' (' + d.posgradoCual + ')' : '') : 'Sin posgrado';
  const ia = (d.ia || '') + (d.iaTools ? ' — usa: ' + d.iaTools : '');
  const conoc = [['Excel', d.excel], ['Power BI', d.powerbi], ['Inglés', d.ingles], ['Herramientas de IA', ia]]
    .filter(x => x[1]).map(x => '<tr><td class="k">' + x[0] + '</td><td class="v">' + x[1] + '</td></tr>').join('');

  const html = `<!DOCTYPE html><html><head><meta charset="utf-8"><style>
    body{font-family:'Plus Jakarta Sans',Arial,sans-serif;margin:0;color:#111;background:#F7F5F1}
    .hero{background:#1A6645;padding:30px 34px;color:#fff}
    .hero .eyebrow{color:rgba(255,255,255,.7);font-size:11px;font-weight:700;letter-spacing:1.4px;text-transform:uppercase;margin-bottom:8px}
    .hero .pos{font-size:24px;font-weight:800}
    .hero .loc{color:rgba(255,255,255,.65);font-size:13px;margin-top:6px}
    .body{padding:28px 34px}
    .sec{color:#1A6645;font-weight:700;font-size:12px;letter-spacing:1px;text-transform:uppercase;margin:22px 0 8px;border-bottom:2px solid #F4D03F;padding-bottom:4px}
    table{width:100%;border-collapse:collapse}
    td{padding:7px 0;font-size:13px;vertical-align:top;border-bottom:1px solid #eee}
    td.k{color:#666;width:42%;font-weight:600}
    td.v{color:#111}
  </style></head><body>
    <div class="hero">
      <div class="eyebrow">Ficha de postulación</div>
      <div class="pos">${d.nombre || ''}</div>
      <div class="loc">${cfg.posicion || perfil}</div>
    </div>
    <div class="body">
      <div class="sec">Datos personales</div>
      <table>
        <tr><td class="k">Email</td><td class="v">${d.email || '-'}</td></tr>
        <tr><td class="k">Teléfono</td><td class="v">${d.telefono || '-'}</td></tr>
        <tr><td class="k">DNI</td><td class="v">${d.dni || '-'}</td></tr>
        <tr><td class="k">Localidad</td><td class="v">${d.localidad || '-'}</td></tr>
        <tr><td class="k">Movilidad propia</td><td class="v">${d.movilidad || '-'}</td></tr>
        ${d.fulltime ? '<tr><td class="k">Disponibilidad full-time</td><td class="v">' + d.fulltime + '</td></tr>' : ''}
        ${d.presencial ? '<tr><td class="k">Trabajo 100% presencial</td><td class="v">' + d.presencial + '</td></tr>' : ''}
        <tr><td class="k">Remuneración pretendida</td><td class="v">${d.remuneracion || '-'}</td></tr>
      </table>
      <div class="sec">Formación</div>
      <table>
        <tr><td class="k">Formación</td><td class="v">${formacion || '-'}</td></tr>
        <tr><td class="k">Posgrado</td><td class="v">${posgrado}</td></tr>
      </table>
      ${conoc ? '<div class="sec">Conocimientos</div><table>' + conoc + '</table>' : ''}
      ${dyn ? '<div class="sec">Sobre esta búsqueda</div><table>' + dyn + '</table>' : ''}
    </div>
  </body></html>`;

  const folderId = cfg.folder_pdf_postulacion || cfg.folder_recibidos;
  const blob = Utilities.newBlob(html, 'text/html', 'Postulacion_' + (d.nombre || 'candidato').replace(/\s+/g, '_') + '.html');
  const pdf = DriveApp.getFolderById(folderId).createFile(blob.getAs('application/pdf'));
  return pdf.getUrl();
}
/* EXPORTA candidatos contactados a una PESTAÑA "Contactos" del Sheet */

function getLinkScreening_(perfilId, nombreCandidato) {
  try {
    const cfg = getConfig(perfilId);
    const sh = getSheet(cfg);
    const datos = sh.getDataRange().getValues();
    const norm = s => String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim();
    const buscado = norm(nombreCandidato);
    for (let i = 1; i < datos.length; i++) {
      if (norm(datos[i][1]) === buscado) {
        const celda = sh.getRange(i+1, 13);  // col 13 = Detalle (PDF screening)
        let link = '';
        // 1) probar fórmula HYPERLINK
        const formula = celda.getFormula();
        if (formula) {
          const m = formula.match(/HYPERLINK\("([^"]+)"/i);
          if (m) link = m[1];
        }
        // 2) probar RichText
        if (!link) {
          const rt = celda.getRichTextValue();
          if (rt && rt.getLinkUrl()) link = rt.getLinkUrl();
        }
        // 3) probar valor plano
        if (!link) {
          const v = String(celda.getValue() || '');
          if (v.indexOf('http') === 0) link = v;
        }
        const nota = String(datos[i][11] || '');  // col 12 = nota screening
        return { link: link, nota: nota };
      }
    }
    return { link: '', nota: '' };
  } catch (e) {
    return { link: '', nota: '' };
  }
}

function setupColumnasInforme() {
  const perfiles = perfilesActivos_();
  perfiles.forEach(p => {
    const perfilId = (typeof p === 'string') ? p : p.id;
    const cfg = getConfig(perfilId);
    const sh = getSheet(cfg);
    sh.getRange(1, 32).setValue('Generar Informe');
    sh.getRange(1, 33).setValue('Informe');
    sh.getRange(1, 32, 1, 2).setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
  });
  SpreadsheetApp.getUi().alert('Listo: se agregaron las columnas del informe.');
}
function setupColumnasInforme() {
  const perfiles = perfilesActivos_();
  perfiles.forEach(p => {
    const perfilId = (typeof p === 'string') ? p : p.id;
    const cfg = getConfig(perfilId);
    const sh = getSheet(cfg);
    sh.getRange(1, 32).setValue('Generar Informe');
    sh.getRange(1, 33).setValue('Informe');
    sh.getRange(1, 32, 1, 2).setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
    Logger.log('Escrito en: ' + sh.getName() + ' → col32=' + sh.getRange(1,32).getValue());
  });
}
function ponerBotonesInforme() {
  const ESTADOS_OK = ['Avanza a empresa','Entrevista empresa agendada','En cliente','Contratado'];
  const perfiles = perfilesActivos_();
  let total = 0;
  perfiles.forEach(p => {
    const perfilId = (typeof p === 'string') ? p : p.id;
    const cfg = getConfig(perfilId);
    const sh = getSheet(cfg);
    const datos = sh.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      const estado = String(datos[i][8] || '');
      if (ESTADOS_OK.indexOf(estado) !== -1) {
        const url = URL_WEBAPP + '?tipo=informe&perfil=' + encodeURIComponent(perfilId) + '&fila=' + (i+1);
        const rich = SpreadsheetApp.newRichTextValue()
          .setText('📄 Generar informe')
          .setLinkUrl(url)
          .build();
        sh.getRange(i+1, 32).setRichTextValue(rich);
        total++;
      }
    }
  });
  Logger.log('Botones puestos: ' + total);
}
const ANTHROPIC_API_KEY = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');

/* Modelos de Claude por tarea (precio/calidad). Se pueden cambiar sin publicar con las propiedades del script
   MODELO_ANALISIS / MODELO_INFORME / MODELO_REDACTAR / MODELO_CV (ej.: claude-opus-5-5). */
const IA_TAREAS_ = {
  analisis: { modelo: 'claude-sonnet-5-5', effort: 'medium', maxTokens: 12000 },
  informe:  { modelo: 'claude-opus-5-5',   effort: 'medium', maxTokens: 12000 },
  redactar: { modelo: 'claude-sonnet-5-5', effort: 'low',    maxTokens: 4000 },
  cv:       { modelo: 'claude-sonnet-5-5', effort: 'low',    maxTokens: 4000 }
};

function procesarConClaude(promptSistema, contenidoUsuario, tarea) {
  tarea = IA_TAREAS_[tarea] ? tarea : 'informe';
  const base = IA_TAREAS_[tarea];
  const prop = PropertiesService.getScriptProperties().getProperty('MODELO_' + tarea.toUpperCase());
  const cfg = { modelo: prop || base.modelo, effort: base.effort, maxTokens: base.maxTokens };
  try {
    return llamarClaude_(promptSistema, contenidoUsuario, cfg);
  } catch (e) {
    // Si el modelo más grande tarda demasiado (Apps Script corta las esperas largas), reintenta con uno más rápido.
    if (/timeout|timed out|deadline/i.test(String(e && e.message)) && cfg.modelo !== 'claude-sonnet-5-5') {
      Logger.log('Claude tardó demasiado con ' + cfg.modelo + ': reintento con claude-sonnet-5-5');
      return llamarClaude_(promptSistema, contenidoUsuario, { modelo: 'claude-sonnet-5-5', effort: 'low', maxTokens: cfg.maxTokens });
    }
    throw e;
  }
}

function llamarClaude_(promptSistema, contenidoUsuario, cfg) {
  const response = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    headers: { 'x-api-key': ANTHROPIC_API_KEY, 'anthropic-version': '2023-06-01' },
    contentType: 'application/json',
    payload: JSON.stringify({
      model: cfg.modelo,
      max_tokens: cfg.maxTokens,
      output_config: { effort: cfg.effort },
      system: promptSistema,
      messages: [{ role: 'user', content: contenidoUsuario }]
    }),
    muteHttpExceptions: true
  });
  let parsed;
  try { parsed = JSON.parse(response.getContentText()); }
  catch (e) { throw new Error('Anthropic respondió algo ilegible (código ' + response.getResponseCode() + ')'); }
  if (parsed.error) throw new Error('Anthropic: ' + parsed.error.message);
  if (parsed.stop_reason === 'refusal') throw new Error('Anthropic rechazó procesar este contenido.');
  if (parsed.stop_reason === 'max_tokens') throw new Error('La respuesta de la IA se cortó por largo. Probá de nuevo.');
  const texto = (parsed.content || []).filter(b => b.type === 'text').map(b => b.text).join('');
  if (!texto) throw new Error('Anthropic no devolvió texto.');
  return texto;
}
function extraerFileId(sheet, fila, col) {
  let url = '';
  const formula = sheet.getRange(fila, col).getFormula();
  if (formula) {
    const m = formula.match(/HYPERLINK\("([^"]+)"/i);
    if (m) url = m[1];
  }
  if (!url) {
    const rt = sheet.getRange(fila, col).getRichTextValue();
    if (rt && rt.getLinkUrl()) url = rt.getLinkUrl();
  }
  if (!url) url = String(sheet.getRange(fila, col).getValue() || '');

  // si el link es de la web app (?...&id=XXX), sacar ese id primero
  const mId = url.match(/[?&]id=([-\w]{25,})/);
  if (mId) return mId[1];

  // si no, buscar un id de Drive normal en la URL
  const m2 = url.match(/[-\w]{25,}/);
  return m2 ? m2[0] : '';
}
function generarInformeCandidato_(perfilId, fila) {
  const cfg = getConfig(perfilId);
  const sh = getSheet(cfg);
  const nombre = String(sh.getRange(fila, 2).getValue() || '');
  const posicion = (typeof cfg.nombre === 'string') ? cfg.nombre : perfilId;
  const LOGO_URL = 'https://lh3.googleusercontent.com/d/1WdoIWop2mZ_KnSa9raQM1Xb5TFP2JvGO';
  const LOGO_VERDE = 'https://lh3.googleusercontent.com/d/1457Wfl9zeY1V3sAGtY1IT2BmF2lqpoy8';

  let cvUrl = '';
  try {
    const idCV0 = extraerFileId(sh, fila, 8);
    if (idCV0) {
      try { DriveApp.getFileById(idCV0).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch(e){}
      cvUrl = 'https://drive.google.com/file/d/' + idCV0 + '/view';
    }
  } catch(e){}

  let textoCV = '', textoScr = '', textoComp = '';
  try { const idCV = extraerFileId(sh, fila, 8);  if (idCV) textoCV = textoDeCV(idCV); } catch(e){}
  try { const idS  = extraerFileId(sh, fila, 13); if (idS)  textoScr = textoDeCV(idS); } catch(e){}
  try { const idC  = extraerFileId(sh, fila, 17); if (idC)  textoComp = textoDeCV(idC); } catch(e){}

  const promptSistema = 'Sos consultor senior de Blumia Consulting. Analizá las fuentes del candidato "' + nombre + '" (posición: ' + posicion + ', cliente ' + empresaInforme_(cfg) + ') y devolvé SOLO un JSON válido, sin texto adicional ni markdown, con esta estructura exacta:\n' +
'{\n' +
'  "edad": "", "localidad": "", "formacion": "", "ultimo_empleo": "",\n' +
'  "pretension": "", "familia": "", "telefono": "", "email": "",\n' +
'  "movilidad": "", "disponibilidad": "",\n' +
'  "herramientas": "Excel, Power BI, IA u otras que surjan del screening; vacío si no hay",\n' +
'  "perfil": "Un párrafo de trayectoria y perfil (4-5 oraciones).",\n' +
'  "motivo_egreso": "Una oración.",\n' +
'  "motivacion": "Una oración.",\n' +
'  "recorrido": [ {"periodo":"2014-2025","empresa":"Empresa · Puesto","detalle":"1 oración breve"} ],\n' +
'  "competencias": [ {"nombre":"Nombre de la competencia evaluada","nivel":4,"texto":"1-2 oraciones MÁXIMO describiendo el nivel demostrado"} ],\n' +
'  "conclusion": "Dos párrafos cortos separados por un salto de línea: el primero con fortalezas, el segundo con aspectos a observar. Conductual y actitudinal, sin elogios excesivos."\n' +
'}\n\n' +
'REGLAS:\n' +
'- Las competencias son las que REALMENTE se evaluaron en la fuente de competencias. Usá el nombre y nivel tal cual aparecen.\n' +
'- "nivel" es 1 a 4: 1=No demuestra, 2=Señales de desarrollo, 3=Demuestra efectividad, 4=Demuestra solidez.\n' +
'- Evidencias CORTAS (1-2 oraciones).\n' +
'- El recorrido laboral va ordenado del empleo MÁS RECIENTE al más antiguo (el último empleo primero).\n' +
'- NUNCA menciones al entrevistador ni digas "no se registró evidencia". Si no hay un ejemplo concreto, describí el nivel demostrado en general, sin forzar un ejemplo y sin señalar la ausencia como aspecto a mejorar.\n' +
'- IMPORTANTE: buscá con atención en TODAS las fuentes (CV, screening y sobre todo COMPETENCIAS) los datos personales: edad, composición familiar, teléfono, email, localidad, movilidad, disponibilidad, pretensión. Estos datos suelen figurar en el formulario de competencias o screening. Extraelos aunque estén en medio del texto.\n' +
'- No inventes hechos ni datos que no estén en las fuentes; si un dato realmente no aparece en ninguna fuente, dejá el campo vacío "".\n' +
'- La conclusión NUNCA dice "apto para presentar al cliente" ni recomienda avanzar; es analítica.';

  const contenido = '=== CV ===\n' + (textoCV || '(no disponible)') + '\n\n=== SCREENING ===\n' + (textoScr || '(no disponible)') + '\n\n=== COMPETENCIAS ===\n' + (textoComp || '(no disponible)');

  let raw = procesarConClaude(promptSistema, contenido);
  raw = raw.replace(/```json/gi,'').replace(/```/g,'').trim();
  let d;
  try { d = JSON.parse(raw); } catch(e){ throw new Error('Claude no devolvió JSON válido: ' + raw.substring(0,200)); }

  const NIVEL_TXT = {1:'No demuestra',2:'Señales de desarrollo',3:'Demuestra efectividad',4:'Demuestra solidez'};
  const internos = String(cfg.datos_internos || '').toUpperCase() !== 'NO'; // Edad y Composición familiar: uso interno, se pueden ocultar por búsqueda (⚙ Configurar → Datos)
  const dato = (l,v) => '<div class="dl">'+l+'</div><div class="dv">'+(v && String(v).trim() ? v : 'Sin dato')+'</div>';
  const datosHTML = [
    internos ? dato('Edad', d.edad) : '', dato('Localidad', d.localidad), dato('Formación', d.formacion),
    dato('Último empleo', d.ultimo_empleo), dato('Pretensión', d.pretension), internos ? dato('Composición familiar', d.familia) : '',
    dato('Teléfono', d.telefono), dato('Email', d.email), dato('Movilidad', d.movilidad),
    dato('Disponibilidad', d.disponibilidad), dato('Herramientas', d.herramientas)
  ].filter(Boolean).join('');
  const recHTML = (d.recorrido||[]).map(r =>
    '<div class="rec-item"><div class="rec-meta"><span class="rec-per">'+(r.periodo||'')+'</span><span class="rec-emp">'+(r.empresa||'')+'</span></div><p class="rec-det">'+(r.detalle||'')+'</p></div>'
  ).join('');
  const compHTML = (d.competencias||[]).map(c =>
    '<div class="cmp"><div class="cmp-head"><span class="cmp-name">'+(c.nombre||'')+'</span><span class="cmp-nivel">'+(c.nivel||'')+'/4 · '+(NIVEL_TXT[c.nivel]||'')+'</span></div><p>'+(c.texto||'')+'</p></div>'
  ).join('');
  const cvBtn = cvUrl ? '<a href="'+cvUrl+'" target="_blank" contenteditable="false" style="display:inline-block;background:#fff;color:#1A6645;padding:9px 20px;border-radius:4px;text-decoration:none;font-weight:800;font-size:9.5pt">📄 Ver Currículum →</a>' : '';
  const conclHTML = (d.conclusion||'').split('\n').filter(Boolean).map(p=>'<p>'+p+'</p>').join('');

  const htmlFull = '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Informe · '+nombre+'</title><style>'+
'*{box-sizing:border-box;margin:0;padding:0;-webkit-print-color-adjust:exact;print-color-adjust:exact;}'+
'body{font-family:\'PP Pangram Sans\',\'Plus Jakarta Sans\',sans-serif;color:#1A1A1A;background:#F7F5F1;font-size:10pt;line-height:1.45;}'+
'.page{background:#fff;max-width:210mm;margin:16px auto;box-shadow:0 1px 4px rgba(0,0,0,.06);}'+
'.body-cell{padding:0 11mm 11mm;}'+
'.hero{background:#1A6645;color:#fff;padding:20px 24px;border-bottom:4px solid #FBC67B;margin:0 -11mm 16px;}'+
'.hero img{height:38px;margin-bottom:12px;display:block;}'+
'.hero .tag{font-size:8.5pt;font-weight:800;letter-spacing:.22em;text-transform:uppercase;color:#A8E6C7;margin-bottom:5px;}'+
'.hero h1{font-size:23pt;font-weight:800;letter-spacing:-.02em;}'+
'.herorow{display:flex;justify-content:space-between;align-items:flex-end;gap:16px;flex-wrap:wrap;}'+
'.hero .sub{font-size:13pt;font-weight:700;margin-top:6px;background:rgba(255,255,255,.15);display:inline-block;padding:3px 12px;border-radius:4px;}'+
'.ficha-data{background:#FCFAF6;border-radius:4px;padding:12px 18px;margin:16px 0;display:grid;grid-template-columns:repeat(2,1fr);gap:5px 24px;}'+
'.dl{font-size:7.5pt;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:#555;}'+
'.dv{font-size:9.5pt;margin-bottom:4px;}'+
'h3{font-size:10pt;font-weight:800;text-transform:uppercase;letter-spacing:.1em;color:#1A1A1A;border-bottom:1px solid #DDD;padding-bottom:4px;margin:14px 0 8px;}'+
'.perfil-text{text-align:justify;margin-bottom:10px;}'+
'.extra{display:grid;grid-template-columns:1fr 1fr;gap:12px;background:#FCFAF6;border-left:3px solid #1A6645;padding:10px 14px;border-radius:3px;margin-bottom:6px;}'+
'.extra .l{font-size:8pt;font-weight:800;text-transform:uppercase;letter-spacing:.1em;color:#1A6645;}'+
'.extra .v{font-size:9.5pt;}'+
'.rec-item{border-left:2px solid #1A6645;padding:2px 0 2px 12px;margin-bottom:6px;}'+
'.rec-meta{display:flex;gap:10px;align-items:baseline;flex-wrap:wrap;}'+
'.rec-per{font-size:7.5pt;font-weight:800;color:#1A6645;background:#E6F2EC;padding:1px 6px;border-radius:2px;}'+
'.rec-emp{font-size:9.5pt;font-weight:700;}'+
'.rec-det{font-size:8.5pt;}'+
'.cmp-grid{column-count:2;column-gap:14px;}'+
'.cmp{background:#FCFAF6;border-left:2.5px solid #1A6645;padding:8px 12px;border-radius:0 2px 2px 0;margin-bottom:8px;break-inside:avoid;display:inline-block;width:100%;}'+
'.cmp-head{display:flex;justify-content:space-between;align-items:baseline;gap:6px;margin-bottom:3px;}'+
'.cmp-name{font-size:9pt;font-weight:800;color:#1A6645;line-height:1.2;}'+
'.cmp-nivel{font-size:7.5pt;font-weight:800;color:#1A6645;background:#E6F2EC;padding:1px 7px;border-radius:10px;white-space:nowrap;}'+
'.cmp p{font-size:8.5pt;line-height:1.4;}'+
'.conclusion{background:#1A6645;color:#fff;padding:14px 18px;border-radius:4px;margin-top:12px;break-inside:avoid;}'+
'.conclusion .eye{font-size:8.5pt;font-weight:800;letter-spacing:.22em;text-transform:uppercase;color:#A8E6C7;margin-bottom:6px;}'+
'.conclusion p{font-size:9.5pt;line-height:1.55;text-align:justify;margin-bottom:6px;}'+
'.foot{margin-top:20px;padding-top:14px;border-top:1px solid #DDD;text-align:center;break-inside:avoid;}'+
'.foot img{height:26px;}'+
'#barra-guardar{position:fixed;top:12px;right:12px;z-index:9999;}'+
'#barra-guardar button{background:#FBC67B;color:#1A1A1A;border:none;padding:10px 18px;border-radius:6px;font-weight:800;font-size:11pt;cursor:pointer;box-shadow:0 2px 8px rgba(0,0,0,.2);}'+
'@media print{body{background:#fff;}.page{box-shadow:none;margin:0;max-width:100%;}@page{margin:10mm 0;}#barra-guardar{display:none!important;}}'+
'</style></head><body contenteditable="true">'+
'<div id="barra-guardar" contenteditable="false"><button onclick="guardarCambios()">💾 Guardar cambios</button></div>'+
'<div class="page"><div class="body-cell">'+
'<div class="hero"><img src="'+LOGO_URL+'" alt="Blumia"><div class="tag">Informe de Candidato · '+empresaInforme_(cfg)+'</div>'+
'<div class="herorow"><div><h1>'+nombre+'</h1><div class="sub">'+posicion+'</div></div>'+cvBtn+'</div></div>'+
'<div class="ficha-data">'+datosHTML+'</div>'+
'<h3>Trayectoria y perfil</h3><p class="perfil-text">'+(d.perfil||'')+'</p>'+
'<div class="extra"><div><div class="l">Motivo de egreso / situación</div><div class="v">'+(d.motivo_egreso||'')+'</div></div>'+
'<div><div class="l">Motivación para el cambio</div><div class="v">'+(d.motivacion||'')+'</div></div></div>'+
'<h3>Recorrido laboral</h3>'+recHTML+
'<h3>Evaluación por competencias</h3><div class="cmp-grid">'+compHTML+'</div>'+
'<div class="conclusion"><div class="eye">Conclusión del perfil</div>'+conclHTML+'</div>'+
'<div class="foot"><img src="'+LOGO_VERDE+'" alt="Blumia"></div>'+
'</div></div>'+
'<script>'+
'function guardarCambios(){var btn=document.querySelector("#barra-guardar button");btn.textContent="Guardando...";btn.disabled=true;'+
'var clon=document.documentElement.cloneNode(true);'+
'var b=clon.querySelector("#barra-guardar button");if(b){b.textContent="💾 Guardar cambios";b.disabled=false;}'+
'var htmlLimpio="<!DOCTYPE html>"+clon.outerHTML;'+
'google.script.run.withSuccessHandler(function(r){btn.textContent=(r&&r.ok)?"✓ Guardado":"Error";btn.disabled=false;setTimeout(function(){btn.textContent="💾 Guardar cambios";},2500);})'+
'.withFailureHandler(function(e){btn.textContent="Error";btn.disabled=false;}).guardarInformeEditado("__FILEID__",htmlLimpio);}'+
'document.addEventListener("click",function(e){var a=e.target.closest("a");if(a&&a.href){e.preventDefault();window.open(a.href,"_blank");}});'+
'</script>'+
'</body></html>';

  const carpeta = DriveApp.getRootFolder();
  const archivo = carpeta.createFile('Informe_' + nombre.replace(/\s+/g,'_') + '.html', htmlFull, MimeType.HTML);
  const fileId = archivo.getId();
  archivo.setContent(htmlFull.replace('__FILEID__', fileId));
  const url = URL_WEBAPP + '?tipo=verinforme&id=' + fileId;
  const rich = SpreadsheetApp.newRichTextValue().setText('📄 Ver informe').setLinkUrl(url).build();
  sh.getRange(fila, 33).setRichTextValue(rich);
  return url;
}
function generarPDFClientes(perfilId) {
  const PDFSHIFT_KEY = pdfshiftKey_();
  const cfg = getConfig(perfilId);
  const sh = getSheet(cfg);
  const posicion = (typeof cfg.nombre === 'string') ? cfg.nombre : perfilId;
  const datos = sh.getDataRange().getValues();

  let items = [];
  let estilo = '';

  for (let i = 1; i < datos.length; i++) {
    const fileId = extraerFileId(sh, i+1, 33);
    if (!fileId) continue;
    let html = '';
    try { html = DriveApp.getFileById(fileId).getBlob().getDataAsString(); } catch(e){ continue; }

    if (!estilo) {
      const ms = html.match(/<style>([\s\S]*?)<\/style>/i);
      if (ms) estilo = ms[1];
    }
    const mb = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
    if (!mb) continue;

    const mn = html.match(/<h1>([\s\S]*?)<\/h1>/i);
    const nombre = mn ? mn[1].replace(/<[^>]+>/g,'').trim() : 'Candidato';
    const loc = (html.match(/Localidad<\/div><div class="dv">([\s\S]*?)<\/div>/i)||[])[1] || '';
    const emp = (html.match(/Último empleo<\/div><div class="dv">([\s\S]*?)<\/div>/i)||[])[1] || '';
    const rem = (html.match(/Pretensión<\/div><div class="dv">([\s\S]*?)<\/div>/i)||[])[1] || '';

    const orden = Number(sh.getRange(i+1, 35).getValue()) || 999;
    const valoracion = String(sh.getRange(i+1, 36).getValue() || '');

    items.push({ nombre, loc, emp, rem, orden, valoracion, body: mb[1] });
  }

  if (items.length === 0) throw new Error('No hay informes generados en este puesto (col 33 vacía).');

  // ordenar por "Orden cliente" (col 35); los sin orden van al final
  items.sort((a,b) => a.orden - b.orden);

  const filas = items.map(r =>
    '<tr><td class="cn">'+r.nombre+'</td><td>'+(r.loc||'—')+'</td><td>'+(r.emp||'—')+'</td><td>'+(r.rem||'—')+'</td><td class="cs">'+(r.valoracion||'—')+'</td></tr>'
  ).join('');

  const LOGO_URL = 'https://lh3.googleusercontent.com/d/1WdoIWop2mZ_KnSa9raQM1Xb5TFP2JvGO';
  const portada =
    '<div class="cuadro-page"><div class="cuadro-hero"><img src="'+LOGO_URL+'"><div class="ctag">Informe de Candidatos · '+empresaInforme_(cfg)+'</div>'+
    '<div class="ctitle">'+posicion+'</div></div>'+
    '<h3 class="cuadro-h3">Cuadro comparativo</h3>'+
    '<table class="cuadro"><thead><tr><th>Candidato</th><th>Localidad</th><th>Último empleo</th><th>Remuneración</th><th>Valoración</th></tr></thead>'+
    '<tbody>'+filas+'</tbody></table></div>';

  const estiloCuadro =
    '.cuadro-page{padding:0;page-break-after:always;}'+
    '.cuadro-hero{background:#1A6645;color:#fff;padding:24px 28px;border-bottom:4px solid #FBC67B;}'+
    '.cuadro-hero img{height:40px;margin-bottom:14px;display:block;}'+
    '.ctag{font-size:9pt;font-weight:800;letter-spacing:.22em;text-transform:uppercase;color:#A8E6C7;margin-bottom:6px;}'+
    '.ctitle{font-size:22pt;font-weight:800;}'+
    '.cuadro-h3{font-size:11pt;font-weight:800;text-transform:uppercase;letter-spacing:.1em;color:#1A1A1A;padding:0 11mm;margin:20px 0 10px;}'+
    '.cuadro{width:calc(100% - 22mm);margin:0 11mm;border-collapse:collapse;font-size:10pt;}'+
    '.cuadro th{background:#1A6645;color:#fff;text-align:left;padding:8px 12px;font-size:8.5pt;text-transform:uppercase;letter-spacing:.05em;}'+
    '.cuadro td{padding:8px 12px;border-bottom:1px solid #E6E4E0;}'+
    '.cuadro tr:nth-child(even) td{background:#FCFAF6;}'+
    '.cuadro .cn{font-weight:700;}'+
    '.cuadro .cs{font-weight:800;color:#1A6645;}';

  const sep = '<div style="page-break-before:always"></div>';
  const cuerpos = items.map(it => it.body).join(sep);
  const htmlParaPDF = '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><style>' + estilo + estiloCuadro + ' .page{page-break-after:always;}</style></head><body>' + portada + cuerpos + '</body></html>';

  const resp = UrlFetchApp.fetch('https://api.pdfshift.io/v3/convert/pdf', {
    method: 'post',
    headers: { 'Authorization': 'Basic ' + Utilities.base64Encode('api:' + PDFSHIFT_KEY) },
    contentType: 'application/json',
    payload: JSON.stringify({ source: htmlParaPDF, format: 'A4', use_print: true }),
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) {
    throw new Error('PDFShift error ' + resp.getResponseCode() + ': ' + resp.getContentText().substring(0,300));
  }

  let carpeta;
  const it = DriveApp.getFoldersByName('Informes Clientes');
  carpeta = it.hasNext() ? it.next() : DriveApp.createFolder('Informes Clientes');
  const pdfBlob = resp.getBlob().setName('Informes_' + posicion.replace(/\s+/g,'_') + '.pdf');
  const pdfFile = carpeta.createFile(pdfBlob);
  pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  Logger.log('PDF generado (' + items.length + ' informes): ' + pdfFile.getUrl());
  return pdfFile.getUrl();
}
function ponerBotonPDFCliente() {
  const perfiles = perfilesActivos_();
  perfiles.forEach(p => {
    const perfilId = (typeof p === 'string') ? p : p.id;
    const cfg = getConfig(perfilId);
    const sh = getSheet(cfg);
    if (sh.getMaxColumns() < 34) {
      sh.insertColumnsAfter(sh.getMaxColumns(), 34 - sh.getMaxColumns());
    }
    const url = URL_WEBAPP + '?tipo=pdfcliente&perfil=' + encodeURIComponent(perfilId);
    const rich = SpreadsheetApp.newRichTextValue()
      .setText('📑 Generar PDF cliente')
      .setLinkUrl(url)
      .build();
    sh.getRange(1, 34).setRichTextValue(rich);
    sh.getRange(1, 34).setBackground('#FBC67B').setFontColor('#1A1A1A').setFontWeight('bold');
  });
  Logger.log('Botón PDF cliente puesto en col 34.');
}
function guardarInformeEditado_(fileId, htmlContenido) {
  try {
    DriveApp.getFileById(fileId).setContent(htmlContenido);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}
function setupColumnasCliente() {
  const perfiles = perfilesActivos_();
  perfiles.forEach(p => {
    const perfilId = (typeof p === 'string') ? p : p.id;
    const cfg = getConfig(perfilId);
    const sh = getSheet(cfg);
    if (sh.getMaxColumns() < 36) sh.insertColumnsAfter(sh.getMaxColumns(), 36 - sh.getMaxColumns());
    sh.getRange(1, 35).setValue('Orden cliente');
    sh.getRange(1, 36).setValue('Valoración cliente');
    sh.getRange(1, 35, 1, 2).setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
    // desplegable en col 36
    const regla = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Se ajusta', 'Se ajusta con observaciones'], true)
      .setAllowInvalid(false)
      .build();
    sh.getRange(2, 36, sh.getMaxRows() - 1, 1).setDataValidation(regla);
  });
  Logger.log('Columnas 35 (Orden) y 36 (Valoración) creadas.');
}
function generarHTMLProceso(empresa) {
  empresa = empresa || empresaDefault_();
  const nombreEmp = empresaInforme_({ empresa: empresa });
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const sh = ss.getSheetByName(hojaProceso_(empresa));
  if (!sh) throw new Error('Corré primero generarInformeProceso');
  const filas = sh.getDataRange().getValues();
  const LOGO = 'https://lh3.googleusercontent.com/d/1457Wfl9zeY1V3sAGtY1IT2BmF2lqpoy8';

  let bs=[], cons=null;
  for(let i=1;i<filas.length;i++){
    const r=filas[i]; if(!r[0])continue;
    const o={nombre:r[0],cv:+r[1]||0,scr:+r[2]||0,comp:+r[3]||0,emp:+r[4]||0,res:+r[5]||0,dsin:+r[6]||0,dpost:+r[7]||0,dcomp:+r[8]||0,banda:r[9],pDeb:r[10],pLin:r[11],pEnc:r[12],ind:r[13],lectura:r[14],fuentes:r[15]};
    if(String(r[0]).indexOf('CONSOLIDADO')===0)cons=o; else bs.push(o);
  }
  const pct=(a,b)=>b?Math.round(a/b*100)+'%':'0%';
  const ce=(t)=>'<span contenteditable="true">'+t+'</span>';
  const fu=f=>String(f||'').split('\n').filter(Boolean).map(x=>'<div class="fuente-item">'+x+'</div>').join('');

  let n=0;
  const secciones=bs.map(b=>{
    n++;
    const stage=(num,lb,p,det,cls,dark)=>'<div class="fnl-stage"><div class="fnl-shape '+cls+'"><div class="fnl-num'+(dark?' dark':'')+'" contenteditable="true">'+num+'</div><div class="fnl-lb'+(dark?' dark':'')+'">'+lb+'</div></div><div class="fnl-side"><div class="fnl-pct" contenteditable="true">'+p+'</div><div class="fnl-det">'+det+'</div></div></div>';
    return '<section class="busqueda"><div class="bsq-head"><div class="bsq-eye">Búsqueda 0'+n+'</div>'+
      '<h2 class="bsq-title" contenteditable="true">'+b.nombre+'</h2>'+
      '<div class="bsq-sub" contenteditable="true">Banda salarial de referencia: $'+(b.banda||'—')+'</div></div>'+
      '<div class="section"><div class="sec-num">01 · Embudo</div><h3 class="sec-title">Del volumen a la <em>precisión</em></h3><div class="funnel">'+
        stage(b.cv,'CVS RECIBIDOS','100%','base del proceso','f1',false)+
        stage(b.scr,'SCREENINGS',pct(b.scr,b.cv),'entrevistas telefónicas','f2',false)+
        stage(b.comp,'COMPETENCIAS',pct(b.comp,b.cv),'evaluación profunda','f3',false)+
        stage(b.emp,'AVANZAN A EMPRESA',pct(b.emp,b.cv),'finalistas presentados','f4',true)+
      '</div>'+(b.res?'<div class="fnl-extra" contenteditable="true">+ '+b.res+' candidatos en reserva activa</div>':'')+'</div>'+
      ((b.pDeb||b.pLin||b.pEnc)?'<div class="section"><div class="sec-num">02 · Realidad salarial</div><h3 class="sec-title">Lo que <em>pretende</em> el mercado</h3><div class="salary-grid">'+
        '<div class="sal-card dark"><div class="sal-pct dark" contenteditable="true">'+(b.pDeb||'—')+'</div><div class="sal-l">Por debajo</div><div class="sal-d">de la banda de referencia</div></div>'+
        '<div class="sal-card green"><div class="sal-pct green" contenteditable="true">'+(b.pLin||'—')+'</div><div class="sal-l">En línea</div><div class="sal-d">con la banda</div></div>'+
        '<div class="sal-card amber"><div class="sal-pct amber" contenteditable="true">'+(b.pEnc||'—')+'</div><div class="sal-l">Por encima</div><div class="sal-d">de la banda</div></div>'+
      '</div></div>':'')+
      (b.ind?'<div class="section"><div class="sec-num">03 · Industrias de origen</div><h3 class="sec-title">De dónde <em>vienen</em></h3><div class="ind-grid">'+String(b.ind).split(',').map(x=>'<div class="ind-chip" contenteditable="true">'+x.trim()+'</div>').join('')+'</div></div>':'')+
      '<div class="section"><div class="sec-num">04 · Descartes</div><h3 class="sec-title">Filtro por <em>etapa</em></h3><div class="stats-grid three">'+
        '<div class="stat-card acc3"><div class="stat-num gray" contenteditable="true">'+b.dsin+'</div><div class="stat-lbl">En análisis de CV</div></div>'+
        '<div class="stat-card acc3"><div class="stat-num gray" contenteditable="true">'+b.dpost+'</div><div class="stat-lbl">Post screening</div></div>'+
        '<div class="stat-card acc3"><div class="stat-num gray" contenteditable="true">'+b.dcomp+'</div><div class="stat-lbl">En competencias</div></div>'+
      '</div></div>'+
      (b.lectura?'<div class="insight"><span class="insight-tag">Lectura del proceso</span><div class="insight-text" contenteditable="true">'+String(b.lectura).replace(/\n/g,'<br>')+'</div></div>':'')+
      (b.fuentes?'<div class="section"><div class="sec-num">05 · Fuentes</div><h3 class="sec-title">Canales de <em>atracción</em></h3><div class="fuentes-grid">'+fu(b.fuentes)+'</div></div>':'')+
    '</section>';
  }).join('');

  const hstats=cons?['CVs Recibidos|'+cons.cv,'Screenings|'+cons.scr,'Competencias|'+cons.comp,'Avanzan a empresa|'+cons.emp].map(s=>{const p=s.split('|');return '<div class="hs-item"><div class="hs-num" contenteditable="true">'+p[1]+'</div><div class="hs-lbl">'+p[0]+'</div></div>';}).join(''):'';

  const cierre=cons?'<section class="resumen"><div class="res-eye">Consolidado · '+bs.length+' búsquedas</div><h2 class="res-tit" contenteditable="true">El proceso <em>avanza</em>, el pipeline está activo</h2><div class="res-grid">'+
    '<div class="res-card"><div class="res-num" contenteditable="true">'+cons.cv+'</div><div class="res-lbl">CVs recibidos</div></div>'+
    '<div class="res-card"><div class="res-num amber" contenteditable="true">'+cons.emp+'</div><div class="res-lbl">Avanzan a empresa</div></div>'+
    '<div class="res-card"><div class="res-num" contenteditable="true">'+cons.res+'</div><div class="res-lbl">En reserva</div></div>'+
    '<div class="res-card"><div class="res-num" contenteditable="true">'+(cons.dsin+cons.dpost+cons.dcomp)+'</div><div class="res-lbl">Descartados</div></div>'+
  '</div><div class="res-lect" contenteditable="true">'+String(cons.lectura||'Escribí acá la lectura general del proceso.').replace(/\n/g,'<br>')+'</div></section>':'';

  const html='<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Informe de Proceso · '+nombreEmp+'</title><style>:root{--bg:#F7F5F1;--bg2:#EEECEA;--white:#FFF;--dark:#111;--text:#1A1A1A;--mid:#555;--light:#999;--line:#DDDBD7;--green:#1A6645;--green2:#238055;--glight:#E6F2EC;--amber:#F5C842;--font:\'PP Pangram Sans\',\'Plus Jakarta Sans\',sans-serif;}*{box-sizing:border-box;margin:0;padding:0;-webkit-print-color-adjust:exact;print-color-adjust:exact;}body{font-family:var(--font);color:var(--text);font-size:14px;line-height:1.5;background:var(--bg);}[contenteditable]:focus{outline:2px solid var(--green);outline-offset:2px;background:rgba(245,200,66,.12);}#barra-guardar{position:fixed;top:14px;right:14px;z-index:9999;}#barra-guardar button{background:var(--amber);color:var(--dark);border:none;padding:11px 20px;border-radius:3px;font-weight:800;font-size:11px;letter-spacing:.14em;text-transform:uppercase;cursor:pointer;box-shadow:0 3px 10px rgba(0,0,0,.25);font-family:var(--font);}.hero{padding:52px 40px 40px;background:linear-gradient(180deg,#FCFAF6 0%,var(--bg) 100%);border-bottom:1px solid var(--line);}.hero-logo{height:56px;margin-bottom:20px;display:block;}.eyebrow{display:block;font-size:11px;font-weight:800;letter-spacing:.2em;text-transform:uppercase;color:var(--green);background:var(--glight);padding:6px 14px;border-radius:2px;margin-bottom:14px;width:fit-content;}.cliente-nombre{font-size:15px;font-weight:800;letter-spacing:.1em;text-transform:uppercase;color:var(--mid);margin-bottom:16px;}.h-title{font-size:46px;line-height:1.06;letter-spacing:-.03em;font-weight:800;margin-bottom:10px;max-width:900px;}.h-title em{color:var(--green);font-style:normal;}.h-stats{display:grid;grid-template-columns:repeat(4,1fr);gap:24px;margin-top:34px;padding-top:30px;border-top:1px solid var(--line);}.hs-num{font-size:50px;font-weight:800;color:var(--green);letter-spacing:-.04em;line-height:1;}.hs-lbl{font-size:11px;font-weight:700;margin-top:8px;text-transform:uppercase;letter-spacing:.08em;}.busqueda{padding:44px 40px;border-bottom:1px solid var(--line);background:#fff;}.bsq-head{margin-bottom:30px;}.bsq-eye{font-size:11px;font-weight:800;letter-spacing:.2em;text-transform:uppercase;color:var(--green);margin-bottom:8px;}.bsq-title{font-size:32px;font-weight:800;letter-spacing:-.025em;line-height:1.1;}.bsq-sub{font-size:14px;color:var(--mid);margin-top:8px;}.section{margin-bottom:40px;}.sec-num{font-size:10px;font-weight:800;letter-spacing:.2em;text-transform:uppercase;color:var(--mid);margin-bottom:8px;}.sec-title{font-size:22px;font-weight:800;letter-spacing:-.02em;margin-bottom:20px;}.sec-title em{color:var(--green);font-style:normal;}.funnel{max-width:640px;margin:0 auto;}.fnl-stage{display:grid;grid-template-columns:1fr 170px;gap:20px;align-items:center;margin-bottom:6px;break-inside:avoid;}.fnl-shape{padding:12px 20px;text-align:center;color:#fff;margin:0 auto;width:100%;}.fnl-shape.f1{background:linear-gradient(180deg,#1A6645,#238055);clip-path:polygon(0 0,100% 0,93% 100%,7% 100%);}.fnl-shape.f2{background:linear-gradient(180deg,#238055,#3FA670);clip-path:polygon(7% 0,93% 0,86% 100%,14% 100%);width:86%;}.fnl-shape.f3{background:#3FA670;clip-path:polygon(14% 0,86% 0,79% 100%,21% 100%);width:70%;}.fnl-shape.f4{background:linear-gradient(180deg,#F5C842,#E0A82E);clip-path:polygon(21% 0,79% 0,72% 100%,28% 100%);width:54%;}.fnl-num{font-size:26px;font-weight:800;letter-spacing:-1px;line-height:1;}.fnl-num.dark{color:#1A1A1A;}.fnl-lb{font-size:9px;font-weight:700;letter-spacing:1.6px;margin-top:3px;opacity:.9;}.fnl-lb.dark{color:#1A1A1A;opacity:1;}.fnl-side{padding-left:12px;border-left:2px solid var(--line);}.fnl-pct{font-size:15px;font-weight:800;color:var(--mid);}.fnl-det{font-size:11px;color:var(--light);margin-top:2px;}.fnl-extra{font-size:12px;color:var(--mid);margin-top:12px;text-align:center;font-weight:600;}.salary-grid{display:grid;grid-template-columns:repeat(3,1fr);gap:14px;}.sal-card{background:#fff;border:1px solid var(--line);padding:24px;border-radius:4px;}.sal-card.green{border-top:3px solid var(--green);}.sal-card.amber{border-top:3px solid var(--amber);}.sal-card.dark{border-top:3px solid var(--dark);}.sal-pct{font-size:40px;font-weight:800;letter-spacing:-.03em;line-height:1;}.sal-pct.green{color:var(--green);}.sal-pct.amber{color:#B8902F;}.sal-pct.dark{color:var(--dark);}.sal-l{font-size:13px;font-weight:700;margin-top:8px;}.sal-d{font-size:12px;color:var(--mid);margin-top:4px;}.ind-grid{display:flex;flex-wrap:wrap;gap:8px;}.ind-chip{background:#FCFAF6;border:1px solid var(--line);border-left:3px solid var(--amber);padding:8px 14px;font-size:12px;border-radius:2px;font-weight:600;}.stats-grid{display:grid;gap:14px;}.stats-grid.three{grid-template-columns:repeat(3,1fr);}.stat-card{background:#fff;border:1px solid var(--line);padding:20px;border-radius:4px;}.stat-card.acc3{border-top:3px solid var(--mid);}.stat-num{font-size:34px;font-weight:800;letter-spacing:-.03em;line-height:1;}.stat-num.gray{color:var(--mid);}.stat-lbl{font-size:11px;font-weight:700;margin-top:6px;text-transform:uppercase;letter-spacing:.06em;}.insight{padding:20px 24px;background:linear-gradient(180deg,#FCFAF6 0%,#F2EFEA 100%);border-left:4px solid var(--green);border-radius:2px;margin-top:16px;break-inside:avoid;}.insight-tag{font-size:10px;font-weight:800;letter-spacing:.2em;text-transform:uppercase;color:var(--green);margin-bottom:8px;display:block;}.insight-text{font-size:13px;line-height:1.6;}.fuentes-grid{display:grid;grid-template-columns:1fr 1fr;gap:8px;}.fuente-item{font-size:13px;color:#333;font-weight:600;}.resumen{padding:52px 40px;background:var(--green);color:#fff;}.res-eye{font-size:11px;font-weight:800;letter-spacing:.2em;text-transform:uppercase;color:var(--amber);margin-bottom:14px;}.res-tit{font-size:34px;font-weight:800;letter-spacing:-.03em;line-height:1.1;margin-bottom:24px;}.res-tit em{color:var(--amber);font-style:normal;}.res-grid{display:grid;grid-template-columns:repeat(4,1fr);gap:18px;}.res-card{background:rgba(255,255,255,.08);border:1px solid rgba(255,255,255,.15);border-radius:4px;padding:22px;}.res-num{font-size:44px;font-weight:800;letter-spacing:-.03em;line-height:1;}.res-num.amber{color:var(--amber);}.res-lbl{font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;margin-top:8px;}.res-lect{font-size:14px;line-height:1.7;color:rgba(255,255,255,.92);margin-top:24px;}footer{padding:28px;text-align:center;font-size:11px;color:var(--mid);background:#fff;border-top:1px solid var(--line);}@media print{@page{size:A4;margin:12mm 10mm;}#barra-guardar{display:none!important;}.busqueda,.resumen{page-break-before:always;}.hero{page-break-after:avoid;}.section,.insight,.fnl-stage{break-inside:avoid;}}</style></head><body>'+
    '<div id="barra-guardar" contenteditable="false"><button onclick="guardarCambios()">GUARDAR CAMBIOS</button></div>'+
    '<section class="hero"><img class="hero-logo" src="'+LOGO+'"><div class="eyebrow">Atracción y Elección · Informe de Proceso</div>'+'<div class="cliente-nombre" contenteditable="true">'+nombreEmp+'</div>'+
    '<h1 class="h-title" contenteditable="true">El proceso está <em>en marcha</em>, el mercado <em>respondió</em>.</h1>'+
    '<div class="h-stats">'+hstats+'</div></section>'+
    secciones+cierre+
    '<footer>Blumia Consulting · Documento generado para '+nombreEmp+'</footer>'+
    '<script>function guardarCambios(){var btn=document.querySelector("#barra-guardar button");btn.textContent="Guardando...";btn.disabled=true;var clon=document.documentElement.cloneNode(true);var b=clon.querySelector("#barra-guardar button");if(b){b.textContent="GUARDAR CAMBIOS";b.disabled=false;}var htmlLimpio="<!DOCTYPE html>"+clon.outerHTML;google.script.run.withSuccessHandler(function(r){btn.textContent=(r&&r.ok)?"GUARDADO":"ERROR";btn.disabled=false;setTimeout(function(){btn.textContent="GUARDAR CAMBIOS";},2500);}).withFailureHandler(function(e){btn.textContent="ERROR";btn.disabled=false;}).guardarInformeEditado("__FILEID__",htmlLimpio);}</scr'+'ipt></body></html>';

  const it=DriveApp.getFoldersByName('Informes Clientes');
  const carpeta=it.hasNext()?it.next():DriveApp.createFolder('Informes Clientes');
  let arch; const ex=carpeta.getFilesByName('Informe_Proceso_'+empresa+'.html');
  if(ex.hasNext()){arch=ex.next();arch.setContent(html);}else{arch=carpeta.createFile('Informe_Proceso_'+empresa+'.html',html,MimeType.HTML);}
  const fileId=arch.getId();
  arch.setContent(html.replace('__FILEID__',fileId));
  const url=URL_WEBAPP+'?tipo=verinforme&id='+fileId;
  Logger.log('HTML editable proceso: '+url);
  return url;
}
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Blumia')
    .addItem('1 · Actualizar números', 'generarInformeProceso')
    .addItem('2 · Completar IA', 'completarProcesoIA')
    .addItem('3 · Generar HTML editable', 'menuHTMLProceso')
    .addItem('4 · Generar PDF final', 'menuPDFProceso')
    .addSeparator()
    .addItem('▶ Proceso completo (1→2→3)', 'menuProcesoCompleto')
    .addToUi();
}

function menuHTMLProceso() {
  const url = generarHTMLProceso();
  SpreadsheetApp.getUi().alert('HTML generado:\n\n' + url);
}

function menuPDFProceso() {
  const url = generarPDFProceso();
  SpreadsheetApp.getUi().alert('PDF generado:\n\n' + url);
}

function menuProcesoCompleto() {
  generarInformeProceso();
  completarProcesoIA();
  const url = generarHTMLProceso();
  SpreadsheetApp.getUi().alert('Proceso completo.\n\nHTML editable:\n' + url + '\n\nEditá y después usá "4 · Generar PDF final".');
}
function crearMenu() {
  SpreadsheetApp.getUi().createMenu('Blumia')
    .addItem('1 · Actualizar números', 'generarInformeProceso')
    .addItem('2 · Completar IA', 'completarProcesoIA')
    .addItem('3 · HTML editable', 'menuHTMLProceso')
    .addItem('4 · PDF final', 'menuPDFProceso')
    .addSeparator()
    .addItem('▶ Completo', 'menuProcesoCompleto')
    .addToUi();
}
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Blumia')
    .addItem('1 · Actualizar números', 'generarInformeProceso')
    .addItem('2 · Completar IA', 'completarProcesoIA')
    .addItem('3 · HTML editable', 'menuHTMLProceso')
    .addItem('4 · PDF final', 'menuPDFProceso')
    .addItem('▶ Proceso completo', 'menuProcesoCompleto')
    .addSeparator()
    .addSubMenu(ui.createMenu('Candidatos')
      .addItem('Generar informe (fila activa)', 'menuInformeFilaActiva')
      .addItem('PDF cliente (perfil activo)', 'menuPDFClienteActivo'))
    .addToUi();
}
function perfilDeHojaActiva_() {
  const nombre = SpreadsheetApp.getActiveSheet().getName();
  let id = '';
  listarPerfiles().forEach(p => { try { if (getConfig(p.id).hoja_nombre === nombre) id = p.id; } catch(e){} });
  return id;
}

function menuInformeFilaActiva() {
  const ui = SpreadsheetApp.getUi();
  const perfilId = perfilDeHojaActiva_();
  if (!perfilId) { ui.alert('Esta hoja no es un perfil.'); return; }
  const fila = SpreadsheetApp.getActiveRange().getRow();
  if (fila < 2) { ui.alert('Seleccioná la fila de un candidato.'); return; }
  const url = generarInformeCandidato_(perfilId, fila);
  ui.alert('Informe generado:\n\n' + url);
}

function menuPDFClienteActivo() {
  const ui = SpreadsheetApp.getUi();
  const perfilId = perfilDeHojaActiva_();
  if (!perfilId) { ui.alert('Esta hoja no es un perfil.'); return; }
  const url = generarPDFClientes(perfilId);
  ui.alert('PDF cliente generado:\n\n' + url);
}
function generarInformeProceso(empresa) {
  empresa = empresa || empresaDefault_();
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let sh = ss.getSheetByName(hojaProceso_(empresa));
  const existe = !!sh;
  if (!sh) sh = ss.insertSheet(hojaProceso_(empresa));

  const DESC_SIN='Descartado sin screening', DESC_POST='Descartado post screening', DESC_COMP='Descartado en competencias', EN_RESERVA='En reserva', AVANZA_EMP='Avanza a empresa';
  const EN_COMPETENCIAS=['Avanza a empresa','En reserva','Descartado en competencias','Avanza a competencias','Competencias agendada'];
  const HIZO_SCREENING_EXTRA=['Descartado post screening','Screening agendado','Avanza a screening'];

  // A..I números | J banda(manual) | K %debajo L %linea M %encima (IA) | N industrias (IA) | O lectura(manual) | P fuentes(manual)
  const headers=['Búsqueda','CVs recibidos','Screenings','Competencias','Avanzan a empresa','En reserva','Desc. sin screening','Desc. post screening','Desc. en competencias',
    'Banda salarial ($)','% debajo','% en línea','% encima','Industrias (comp.+empresa)','Lectura del proceso','Fuentes de reclutamiento'];

  // conservar lo que se carga a mano: J(banda,10) O(lectura,15) P(fuentes,16)
  const manualPrev={};
  if(existe){
    const prev=sh.getDataRange().getValues();
    for(let i=1;i<prev.length;i++){ if(prev[i][0]) manualPrev[prev[i][0]]={banda:prev[i][9]||'',lectura:prev[i][14]||'',fuentes:prev[i][15]||''}; }
  }
  sh.clear();

  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);

  const EXCLUIR=['operaciones'];
  const perfiles=listarPerfiles().filter(p=>EXCLUIR.indexOf(p.id)===-1);
  let fila=2;
  let tot={cv:0,scr:0,comp:0,emp:0,res:0,dsin:0,dpost:0,dcomp:0};
  const FUENTES_DEF='✓ Redes Blumia\n✓ CompuTrabajo\n✓ ZonaJobs\n✓ Portal Universidad Austral';

  perfiles.forEach(p=>{
    let cfg; try{cfg=getConfig(p.id);}catch(e){return;}
    if (empresaKey_(cfg) !== empresa) return;
    const shP=getSheet(cfg);
    const datos=shP.getDataRange().getValues();
    const c=contarBusqueda_(datos.slice(1));
    const m=manualPrev[p.nombre]||{};
    sh.getRange(fila,1,1,16).setValues([[p.nombre,c.cv,c.scr,c.comp,c.emp,c.res,c.dsin,c.dpost,c.dcomp,
      m.banda||String(cfg.banda_salarial||'').replace(/\D/g,''),'','','','',m.lectura||'',m.fuentes||FUENTES_DEF]]);
    fila++;
    Object.keys(tot).forEach(k=>tot[k]+=c[k]);
  });

  const mc=manualPrev['CONSOLIDADO']||{};
  sh.getRange(fila,1,1,16).setValues([['CONSOLIDADO',tot.cv,tot.scr,tot.comp,tot.emp,tot.res,tot.dsin,tot.dpost,tot.dcomp,
    '','','','','',mc.lectura||'',mc.fuentes||FUENTES_DEF]]).setFontWeight('bold').setBackground('#E6F2EC');

  sh.setColumnWidth(1,220);
  sh.setColumnWidth(14,300); sh.setColumnWidth(15,340); sh.setColumnWidth(16,220);
  sh.getRange(2,1,fila-1,16).setWrap(true).setVerticalAlignment('top');
  Logger.log('Informe_Proceso generado.');
}
function completarProcesoIA(empresa) {
  empresa = empresa || empresaDefault_();
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const sh = ss.getSheetByName(hojaProceso_(empresa));
  if (!sh) { Logger.log('Corré primero generarInformeProceso'); return; }
  const filas = sh.getDataRange().getValues();

  const HIZO_SCREENING = ['Avanza a empresa','En reserva','Descartado en competencias','Avanza a competencias','Competencias agendada','Descartado post screening','Screening agendado','Avanza a screening'];
  const AVANZO = ['Avanza a empresa','En reserva','Descartado en competencias','Avanza a competencias','Competencias agendada'];
  const EXCLUIR = ['operaciones'];
  const perfiles = listarPerfiles().filter(p => EXCLUIR.indexOf(p.id) === -1);

  perfiles.forEach(p => {
    // ubicar la fila de este perfil en Informe_Proceso
    let filaInf = -1;
    for (let i=1;i<filas.length;i++){ if(filas[i][0]===p.nombre){ filaInf=i+1; break; } }
    if (filaInf === -1) return;

    const banda = Number(String(sh.getRange(filaInf,10).getValue()).replace(/\D/g,'')) || 0;

    let cfg; try{cfg=getConfig(p.id);}catch(e){return;}
    const shP = getSheet(cfg);
    const datos = shP.getDataRange().getValues();

    // pretensiones de los que hicieron screening
    let debajo=0, linea=0, encima=0, totalSal=0;
    let resumenesAvanzan = [];
    for (let i=1;i<datos.length;i++){
      if(!datos[i][1])continue;
      const estado=String(datos[i][8]||'').trim();
      const req=String(datos[i][5]||'');
      const mSal=req.match(/Remuneraci[oó]n pretendida:\s*\$?\s*([\d.\s]+)/i);
      const mt = metEstado_(estado);
      if (banda && mt && mt.scr && mSal){
        const val=Number(mSal[1].replace(/[.\s]/g,''));
        if(val>0){
          totalSal++;
          if(val < banda*0.95) debajo++;
          else if(val > banda*1.05) encima++;
          else linea++;
        }
      }
      if (mt && mt.comp && datos[i][6]){
        resumenesAvanzan.push(String(datos[i][6]).substring(0,300));
      }
    }

    if (totalSal>0){
      sh.getRange(filaInf,11,1,3).setNumberFormat('@');
      sh.getRange(filaInf,11).setValue(Math.round(debajo/totalSal*100)+'%');
      sh.getRange(filaInf,12).setValue(Math.round(linea/totalSal*100)+'%');
      sh.getRange(filaInf,13).setValue(Math.round(encima/totalSal*100)+'%');
    }

    // industrias con GPT-4o-mini
    if (resumenesAvanzan.length>0){
      try {
        const prompt='Extraé las INDUSTRIAS/SECTORES de origen de estos candidatos. Devolvé SOLO una lista separada por comas, sin repetir, máximo 8 industrias, sin explicaciones.\n\n'+resumenesAvanzan.join('\n---\n');
        const resp=UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions',{
          method:'post',
          headers:{Authorization:'Bearer '+OPENAI_API_KEY},
          contentType:'application/json',
          payload:JSON.stringify({model:'gpt-4o-mini',messages:[{role:'user',content:prompt}]}),
          muteHttpExceptions:true
        });
        const parsed=JSON.parse(resp.getContentText());
        if(parsed.choices){
          sh.getRange(filaInf,14).setValue(parsed.choices[0].message.content.trim());
        }
      } catch(e){ Logger.log('IA industrias falló '+p.id+': '+e); }
    }
    Utilities.sleep(800);
  });
  Logger.log('Proceso IA completado.');
}
// medición opcional de tiempos (la usa ⏱ Velocidad): cada paso del armado del tablero con su duración
let TIMING_ = null, TIMING_T_ = 0, KANBAN_KB_ = 0;
function lapT_(n) { if (!TIMING_) return; const x = Date.now(); TIMING_.push({ n: n, ms: x - TIMING_T_ }); TIMING_T_ = x; }

function kanbanHTML(perfilId) {
  CFG_MEMO_ = {};
  try { return kanbanHTML_(perfilId); } finally { CFG_MEMO_ = null; }
}

function kanbanHTML_(perfilId) {
  const cfg = getConfig(perfilId);
  const emp = empresaKey_(cfg);
  const titulo = ((typeof cfg.nombre === 'string' && cfg.nombre) ? cfg.nombre : (typeof cfg.nombre_perfil === 'string' && cfg.nombre_perfil) ? cfg.nombre_perfil : perfilId) + ' · ' + emp;
  lapT_('Leer la configuración de la búsqueda');
  const data = doGetKanban_(perfilId);
  KANBAN_KB_ = Math.round(data.length / 1024);
  const empresas = JSON.stringify(empresasYPerfiles());
  lapT_('Lista de búsquedas para el selector');
  const empresaActual = empresaKey_(cfg);
  const flyer = URL_WEBAPP + '?tipo=flyer&perfil=' + encodeURIComponent(perfilId); // editor de flyer propio (el link viejo de Config se ve adentro)
  const perfilProt = String(cfg.perfil_protegido || '').toUpperCase() === 'SI';
  const perfilPdf = (!perfilProt && typeof cfg.link_perfil === 'string') ? cfg.link_perfil : '';
  let html = KANBAN_TPL(perfilId, titulo, data, empresas, empresaActual, flyer, perfilPdf, perfilProt);
  html = html.replace('%%SORTABLE%%', () => HtmlService.createHtmlOutputFromFile('SortableLib').getContent()); // arrastrar tarjetas no depende de una CDN externa
  lapT_('Armar la plantilla del tablero');
  let bandaRef = 0;
  try { bandaRef = Number(datosInformeBusqueda(perfilId).banda) || 0; } catch (e) { }
  lapT_('Banda salarial (datos del informe)');
  const enlaces = [['Formulario de postulación (para los candidatos)', URL_WEBAPP + '?tipo=postulacion&perfil=' + encodeURIComponent(perfilId)], ['Tablero de esta búsqueda', URL_WEBAPP + '?tipo=kanban&perfil=' + encodeURIComponent(perfilId)], ['Flyer y textos para redes', flyer], ['Perfil del puesto (documento)', perfilPdf], ['Agenda · screening', cfg.link_screening], ['Agenda · competencias', cfg.link_competencias], ['Agenda · entrevista con la empresa (virtual)', cfg.link_empresa_virtual], ['Agenda · entrevista con la empresa (presencial)', cfg.link_empresa_presencial]].filter(x => typeof x[1] === 'string' && /^https?:\/\//i.test(x[1]));
  try { const extra = '<script>var metEstado = ' + metEstado_.toString() + '; var BANDA_REF = ' + bandaRef + '; var URL_APP = ' + JSON.stringify(URL_BASE()) + '; var ENLACES = ' + JSON.stringify(enlaces).replace(/</g, '\\u003c') + ';</script>' + ['ConfigPanel', 'ResumenPanel', 'TableroCards'].map(f => HtmlService.createHtmlOutputFromFile(f).getContent()).join(''); html = html.replace(/<\/html>\s*$/, () => extra + '</html>'); } catch (e) { }
  lapT_('Sumar los paneles (configurar, resumen, tablero)');
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// links de una columna entera: primero el texto enriquecido (trae la URL ya armada) y la fórmula HYPERLINK solo para las celdas con texto que no tienen link
// (una lectura por columna en vez de una por celda; las fórmulas se piden solo si hace falta)
function linksColumna_(sh, col, n, valores) {
  const out = new Array(n).fill('');
  try {
    const rts = sh.getRange(2, col, n, 1).getRichTextValues();
    let faltan = false;
    rts.forEach((r, i) => {
      const rt = r[0];
      if (rt && rt.getLinkUrl()) out[i] = rt.getLinkUrl();
      else if (String(valores[i] == null ? '' : valores[i]).trim() !== '') faltan = true;
    });
    if (faltan) {
      sh.getRange(2, col, n, 1).getFormulas().forEach((f, i) => {
        if (out[i]) return;
        const m = String(f[0] || '').match(/HYPERLINK\("([^"]+)"/i);
        if (m) out[i] = m[1];
      });
    }
  } catch (e) { }
  return out;
}

// id de archivo de Drive de la columna del CV: primero la fórmula, después el link del texto y por último el valor (igual que extraerFileId)
function idsCVColumna_(sh, col, n, valores) {
  const fs = sh.getRange(2, col, n, 1).getFormulas();
  const url = fs.map(f => { const m = String(f[0] || '').match(/HYPERLINK\("([^"]+)"/i); return m ? m[1] : ''; });
  if (url.some((u, i) => !u && String(valores[i] == null ? '' : valores[i]).trim() !== '')) {
    const rts = sh.getRange(2, col, n, 1).getRichTextValues();
    rts.forEach((r, i) => { if (!url[i]) { const rt = r[0]; if (rt && rt.getLinkUrl()) url[i] = rt.getLinkUrl(); } });
  }
  return url.map((u, i) => {
    u = u || String(valores[i] || '');
    const mId = u.match(/[?&]id=([-\w]{25,})/);
    if (mId) return mId[1];
    const m2 = u.match(/[-\w]{25,}/);
    return m2 ? m2[0] : '';
  });
}

function doGetKanban_(perfilId) {
  const sh = getSheet(getConfig(perfilId));
  const d = sh.getDataRange().getValues();
  lapT_('Leer los candidatos (' + (d.length - 1) + ' filas)');
  const n = d.length - 1;
  const L = {};
  if (n > 0) [13, 17, 31, 33].forEach(c => { L[c] = linksColumna_(sh, c, n, d.slice(1).map(r => r[c - 1])); });
  const cvs = n > 0 ? idsCVColumna_(sh, 8, n, d.slice(1).map(r => r[7])) : [];
  lapT_('Leer los links de cada candidato');
  const link = (r, col) => (L[col] || [])[r - 2] || '';
  const out = [];
  const cMot = d[0].indexOf('Motivo de descarte');
  const cObs = d[0].indexOf('Observaciones');
  const movs = ultimosMovimientos_(sh.getName());
  lapT_('Leer el historial de movimientos');
  const prev = previosDeFilas_(perfilId, d.slice(1));
  lapT_('Buscar si ya se postularon a otra búsqueda (con cache)');
  for (let i = 1; i < d.length; i++) {
    if (!d[i][1]) continue;
    const m = String(d[i][5] || '').match(/Remuneraci[oó]n pretendida:\s*\$?\s*([\d.\s]+)/i);
    out.push({
      fila: i + 1, nombre: d[i][1], estado: String(d[i][8] || '').trim(), motivo: cMot >= 0 ? String(d[i][cMot] || '') : '', obs: cObs >= 0 ? String(d[i][cObs] || '') : '', movMs: movs[String(d[i][1]).trim()] || 0, recMs: d[i][0] ? new Date(d[i][0]).getTime() : 0, previos: prev[i - 1] || [],
      recibido: d[i][0] ? Utilities.formatDate(new Date(d[i][0]), 'GMT-3', 'dd/MM/yyyy') : '',
      nota: d[i][4] || '', notaScr: d[i][11] || '', notaComp: d[i][15] || '',
      pretension: m ? m[1].trim() : '', zona: d[i][23] || '',
      tel: d[i][2] || '', dni: String(d[i][3] || ''), mail: d[i][21] || '',
      resumen: d[i][6] || '', cv: cvs[i - 1] || '',
      informe: link(i + 1, 33),
      scrPDF: link(i + 1, 13), compPDF: link(i + 1, 17), postulacion: link(i + 1, 31), cita: d[i][36] || '', historial: d[i][37] || '', mailDescarte: d[i][39] || '', fuente: String(d[i][9] || ''), contactado: d[i][40] || '', notaScrEmp: d[i][43] || '', notaPresEmp: d[i][44] || '', notaScrEmp: d[i][43] || '', notaPresEmp: d[i][44] || '', formacion: d[i][22] || '', requisitos: d[i][5] || '', pretNum: (function(){var m=String(d[i][5]||'').match(/Remuneraci[oó]n pretendida:\s*\$?\s*([\d.\s]+)/i);return m?Number(m[1].replace(/[.\s]/g,'')):0;})()
    });
  }
  return JSON.stringify(out);
}

function moverEstado_(perfilId, fila, estado, quien) {
  QUIEN_ = quien || '';
  const sh = getSheet(getConfig(perfilId));
  const c = sh.getRange(fila, 9);
  const antes = String(c.getValue() || '').trim();
  const r = c.getDataValidation();
  c.setDataValidation(null).setValue(estado);
  if (r) c.setDataValidation(r);
  limpiarMotivoSiNoDescartado_(sh, fila, estado);
  if (antes !== estado) registrarHistorial_(sh.getName(), sh.getRange(fila, 2).getValue(), 'Estado', (antes || '—') + ' → ' + estado + ' (arrastre en el tablero)');
  return true;
}

function KANBAN_TPL(perfilId, titulo, dataJSON, empresasJSON, empresaActual, flyerUrl, perfilUrl, perfilProt) {
  const LOGO=LOGO_BLUMIA_;
  return `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Kanban · ${titulo}</title>
%%SORTABLE%%
<style>
:root{--green:#243D2C;--glight:#E3EEDC;--amber:#F9E7A0;--bg:#F3F8F1;--line:#D5E0CE;--mid:#555;--red:#E06666;--font:'Poppins','Plus Jakarta Sans',sans-serif;--cream:#F9E7A0;--mint:#DCE8D3;--teal:#5BB5A5}
*{box-sizing:border-box;margin:0;padding:0}body{font-family:var(--font);background:var(--bg);color:#1A1A1A;font-size:14px}
header{background:var(--green);color:#fff;padding:16px 20px;display:flex;align-items:center;gap:14px;position:sticky;top:0;z-index:10;flex-wrap:wrap}
header img{height:34px}header h1{font-size:18px;font-weight:800}
.nav{display:flex;gap:8px;margin-left:auto;flex-wrap:wrap;align-items:center}
.nav button,.nav .dd>button{padding:8px 14px;border-radius:8px;font-size:13px;font-weight:700;color:#fff;background:rgba(255,255,255,.16);border:1px solid rgba(255,255,255,.22);cursor:pointer;font-family:var(--font);white-space:nowrap}
.nav button:hover,.nav .dd>button:hover{background:rgba(255,255,255,.28)}
.nav button.avt{order:99;margin-left:6px;width:38px;height:38px;border-radius:50%;padding:0;background:#FBE794;color:#0D3F29;border:none;font-weight:800;font-size:13px;display:inline-flex;align-items:center;justify-content:center}.nav button.avt:hover{background:#f6dc72}
.nav select{padding:8px 10px;border-radius:8px;border:none;font-family:var(--font);font-weight:700;font-size:13px;cursor:pointer;max-width:260px}
.nav .dd{position:relative}
.nav .ddm{display:none;position:absolute;right:0;top:calc(100% + 6px);background:#fff;color:#1A1A1A;border-radius:10px;box-shadow:0 8px 24px rgba(0,0,0,.25);min-width:230px;padding:6px;z-index:60}
.nav .ddm.show{display:block}
.nav .ddm a{display:block;padding:10px 12px;border-radius:7px;font-size:13px;font-weight:600;cursor:pointer;color:#1A1A1A}
.nav .ddm a:hover{background:var(--glight)}
.filtros{padding:12px 16px 0;display:flex;gap:8px;flex-wrap:wrap;align-items:center}
.fwrap{position:relative;display:inline-block}
.fwrap svg{position:absolute;left:9px;top:50%;transform:translateY(-50%);opacity:.4;pointer-events:none}
.filtros input,.filtros select{padding:10px;border:1px solid var(--line);border-radius:8px;font-family:var(--font);font-size:13px;background:#fff}
.fwrap input{padding-left:30px}
.filtros button{padding:10px 14px;border:1px solid var(--line);border-radius:8px;background:#fff;cursor:pointer;font-family:var(--font);font-size:13px}
.board{display:flex;gap:12px;padding:16px;overflow-x:auto;align-items:flex-start}
.col{background:#fff;border:1px solid var(--line);border-radius:8px;min-width:235px;max-width:235px;flex-shrink:0}
.col-h{padding:12px 14px;border-bottom:2px solid var(--green);display:flex;justify-content:space-between;align-items:center;gap:6px}
.col-h span{font-size:11px;font-weight:800;text-transform:uppercase;letter-spacing:.08em;color:var(--green)}
.col-h b{font-size:20px;font-weight:800;color:var(--green)}
.col.res .col-h{border-bottom-color:var(--amber)}
.col.res .col-h span,.col.res .col-h b{color:#8a6d00}
.col.desc .col-h{border-bottom-color:var(--red)}
.col.desc .col-h span,.col.desc .col-h b{color:var(--red)}
.col-body{padding:10px;display:flex;flex-direction:column;gap:8px;min-height:60px;max-height:70vh;overflow-y:auto}
.acc{padding:8px;max-height:70vh;overflow-y:auto}
.acc details{border-bottom:1px solid var(--line)}
.acc details:last-child{border-bottom:none}
.acc summary{cursor:pointer;list-style:none;padding:9px 4px;font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.04em;display:flex;justify-content:space-between;align-items:center;gap:6px}
.acc summary::-webkit-details-marker{display:none}
.acc summary:hover{background:#FAFAF8}
.acc summary .cnt{background:var(--line);color:#333;border-radius:9px;padding:1px 7px;font-size:10px;font-weight:800;flex-shrink:0}
.col.desc .acc summary{color:var(--red)}
.col.res .acc summary{color:#8a6d00}
.dbody{display:flex;flex-direction:column;gap:8px;min-height:38px;padding:2px 0 8px}
.card{background:#FCFAF6;border:1px solid var(--line);border-left:3px solid var(--green);border-radius:6px;padding:8px 10px;cursor:grab;position:relative}.card .cx{position:absolute;top:4px;right:4px;border:0;background:transparent;color:#B0A99F;font-size:13px;line-height:1;cursor:pointer;padding:2px 5px;border-radius:3px}.card .cx:hover{background:#FDEAE4;color:#B91C1C}
.card:hover{box-shadow:0 2px 8px rgba(0,0,0,.1)}
.card .n{font-weight:800;font-size:13px;margin-bottom:3px}
.card .meta{font-size:11px;color:var(--mid);line-height:1.5}
.cita-tag{font-size:10px;font-weight:700;color:var(--green);background:var(--glight);padding:3px 8px;border-radius:5px;margin-top:6px;display:inline-block}
.tag{display:inline-block;font-size:9px;font-weight:800;text-transform:uppercase;letter-spacing:.05em;padding:2px 7px;border-radius:10px;margin-top:6px}
.tag.ag{background:#DBEAFE;color:#1E40AF}.tag.co{background:#FEF3C7;color:#92400E}
.nota{display:inline-block;width:20px;height:20px;line-height:20px;text-align:center;border-radius:4px;font-weight:800;font-size:11px;color:#fff;float:right}
.nA{background:#1A6645}.nB{background:#F5C842;color:#1A1A1A}.nC{background:#E06666}
.saved{position:fixed;bottom:20px;right:20px;background:var(--green);color:#fff;padding:10px 18px;border-radius:6px;font-weight:800;font-size:12px;opacity:0;transition:opacity .3s;z-index:100}
.saved.show{opacity:1}
.modal-bg{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:200;padding:20px;overflow-y:auto}
.modal-bg.show{display:flex;align-items:flex-start;justify-content:center}
.modal{background:#fff;border-radius:10px;max-width:560px;width:100%;margin:auto}
.modal-h{background:var(--green);color:#fff;padding:20px 24px;border-radius:10px 10px 0 0;display:flex;justify-content:space-between;align-items:center}
.modal-h h2{font-size:22px;font-weight:800}
.mx{background:none;border:none;color:#fff;font-size:20px;cursor:pointer}
.modal-b{padding:24px}
.frow{display:flex;justify-content:space-between;padding:9px 0;border-bottom:1px solid var(--line);font-size:13px;gap:12px}
.fl{font-weight:800;color:var(--mid);text-transform:uppercase;font-size:11px;letter-spacing:.05em;flex-shrink:0}
.fv{text-align:right}
.fres{margin:16px 0;font-size:13px;line-height:1.6;color:#333}
.notabox{margin:16px 0;padding:14px;background:var(--glight);border-radius:8px}
.notatit{font-size:11px;font-weight:800;text-transform:uppercase;color:var(--green);margin-bottom:10px;letter-spacing:.05em}
.notarow{display:flex;align-items:center;gap:6px;margin-bottom:8px}
.notarow span{font-size:12px;font-weight:700;width:110px;color:var(--mid)}
.nb{width:36px;height:30px;border:1px solid var(--line);background:#fff;border-radius:5px;font-weight:800;font-size:12px;cursor:pointer}
.nb:hover{background:var(--green);color:#fff}
.mbtns{display:flex;flex-wrap:wrap;gap:8px;margin-top:8px}
.mbtn{display:inline-block;background:var(--green);color:#fff;padding:10px 20px;border-radius:6px;text-decoration:none;font-weight:800;font-size:13px;border:none;cursor:pointer;font-family:var(--font)}
.mbtn.alt{background:#fff;color:var(--green);border:2px solid var(--green)}
.mbtn.gen{background:var(--amber);color:#1A1A1A}
.mbtn.wa{background:#25D366;color:#fff}
</style></head><body>
<header><img src="${LOGO}"><h1>${titulo}</h1><div class="nav" id="nav"></div></header>
<div class="filtros">
<span class="fwrap"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#1A1A1A" stroke-width="2.5"><circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/></svg><input id="fNombre" type="text" placeholder="Nombre" oninput="filtrar()" style="width:150px"></span>
<span class="fwrap"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#1A1A1A" stroke-width="2.5"><circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/></svg><input id="fZona" type="text" placeholder="Zona / ubicación" oninput="filtrar()" style="width:160px"></span>
<span class="fwrap"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#1A1A1A" stroke-width="2.5"><circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/></svg><input id="fForm" type="text" placeholder="Formación / profesión" oninput="filtrar()" style="width:180px"></span>
<span class="fwrap"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#1A1A1A" stroke-width="2.5"><circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/></svg><input id="fExp" type="text" placeholder="Experiencia / palabra clave" oninput="filtrar()" style="width:200px"></span>
<select id="fEstado" onchange="filtrar()"><option value="">Estado: todos</option></select>
<input id="fPret" type="number" placeholder="Pretensión hasta $" oninput="filtrar()" style="width:160px">
<button onclick="limpiarFiltros()">Limpiar</button>
<span id="contadorF" style="font-size:12px;color:#555;font-weight:700"></span>
</div>
<div class="board" id="board"></div>
<div class="saved" id="saved">✓ Guardado</div>
<div class="modal-bg" id="modal" onclick="if(event.target===this)cerrar()"><div class="modal"><div class="modal-h"><h2 id="mtit"></h2><button class="mx" onclick="cerrar()">✕</button></div><div class="modal-b" id="mbody"></div></div></div>
<script>
const DATA_ALL=${dataJSON};
const DATA=DATA_ALL.filter(function(c){return !/^Eliminado\\|/.test(c.estado);});
const ELIM=DATA_ALL.filter(function(c){return /^Eliminado\\|/.test(c.estado);});
const EMPRESAS=${empresasJSON};
const EMPRESA_ACTUAL='${empresaActual}';
const PERFIL='${perfilId}';
function pedirPerfilPin(){google.script.run.withSuccessHandler(function(u){abrirPerfilLink(u);}).withFailureHandler(function(e){alert((e&&e.message)||e);}).linkPerfilConPin(PERFIL,'sesion','ver');}
function abrirPerfilLink(u){var d=document.createElement('div');d.style.cssText='position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.45);z-index:99999;display:flex;align-items:center;justify-content:center';d.innerHTML='<div style="background:#fff;border-radius:10px;padding:26px 30px;text-align:center;font-family:sans-serif;max-width:340px"><p style="margin:0 0 16px;font-size:15px">Listo ✓</p><a id="pplk" target="_blank" rel="noopener" style="display:inline-block;background:#243D2C;color:#fff;padding:11px 22px;border-radius:99px;text-decoration:none;font-weight:700">Abrir perfil del puesto</a><div style="margin-top:14px"><a href="#" id="ppcl" style="color:#777;font-size:13px">Cerrar</a></div></div>';document.body.appendChild(d);var l=document.getElementById('pplk');l.href=u;l.onclick=function(){setTimeout(function(){d.remove();},400);};document.getElementById('ppcl').onclick=function(ev){ev.preventDefault();d.remove();};}
let navHTML='<button id="avBtn" title="Avisos pendientes" onclick="verAvisos()" style="font-size:17px;padding:6px 12px">🔔<span id="avCnt" style="display:none;margin-left:6px;background:#fff;color:#B91C1C;border-radius:12px;font-size:13px;font-weight:800;padding:1px 9px"></span></button><select onchange="window.top.location.href=\\'${URL_WEBAPP}?tipo=kanban&perfil=\\'+this.value">';
Object.keys(EMPRESAS).forEach(function(emp){
  navHTML+='<optgroup label="'+emp+'">';
  EMPRESAS[emp].forEach(function(p){navHTML+='<option value="'+p[0]+'"'+(p[0]===PERFIL?' selected':'')+'>'+p[1]+'</option>';});
  navHTML+='</optgroup>';
});
navHTML+='</select>';
navHTML+='<button title="Volver al inicio (todas las búsquedas)" onclick="window.top.location.href=\\'${URL_WEBAPP}?tipo=inicio\\'">🏠 Inicio</button>';
navHTML+='<button id="quienBtn" class="avt" title="Tocá para salir" onclick="blMenuUsuario()">👤</button>';
navHTML+='<div class="dd"><button onclick="ddToggle(this)">📄 Informes ▾</button><div class="ddm"><a onclick="window.open(\\'${URL_WEBAPP}?tipo=procesohtml&perfil=${perfilId}\\',\\'_blank\\')">📊 Informe de proceso (para la empresa)</a><a onclick="blConfirm(\\'¿Generar PDF cliente de esta búsqueda?\\',function(){window.open(\\'${URL_WEBAPP}?tipo=pdfcliente&perfil=${perfilId}\\',\\'_blank\\')})">📄 PDF para el cliente</a></div></div>';
navHTML+='<div class="dd"><button onclick="ddToggle(this)">📚 Búsqueda ▾</button><div class="ddm">'+(${perfilProt?'true':'false'}?'<a onclick="pedirPerfilPin()">🔒 Perfil del puesto</a>':'<a onclick="window.open(\\'${perfilUrl}\\',\\'_blank\\')">📋 Perfil del puesto</a>')+'<a onclick="window.open(\\'${flyerUrl}\\',\\'_blank\\')">📢 Flyer y textos</a><a onclick="verEnlaces()">🔗 Enlaces</a><a onclick="verHistorial()">📜 Historial de la búsqueda</a></div></div>';
navHTML+='<div class="dd"><button onclick="ddToggle(this)">🧰 Candidatos ▾</button><div class="ddm"><a onclick="descartesMasivo()">✉ Mails de descarte</a><a title="Baja un archivo para importar a los contactos de Google" onclick="bajarContactos(this)">📇 Contactos al celular</a><a onclick="window.top.location.href=\\'${URL_WEBAPP}?tipo=banco&desde=${perfilId}\\'">★ Banco de talento</a><a onclick="abrirPapelera()">🗑 Papelera ('+ELIM.length+')</a></div></div>';
document.addEventListener('click',function(e){if(!e.target.closest('.dd'))document.querySelectorAll('.ddm.show').forEach(function(m){m.classList.remove('show');});});
function ddToggle(b){var m=b.nextElementSibling,ab=m.classList.contains('show');document.querySelectorAll('.ddm.show').forEach(function(x){x.classList.remove('show');});if(!ab)m.classList.add('show');}
document.getElementById('nav').innerHTML=navHTML;
avInit();
try{var _q=localStorage.getItem('blumia_quien'),_b=document.getElementById('quienBtn');if(_q&&_b){var _p=_q.trim().split(/\\s+/).filter(Boolean);_b.textContent=(_p[0].charAt(0)+(_p.length>1?_p[_p.length-1].charAt(0):'')).toUpperCase();_b.title=_q+' · tocá para salir';}}catch(e){}
const COLS=[['Análisis de CV',['Nuevo','Contactado','Pendiente','']],['Screening',['Avanza a screening','Screening agendado']],['Competencias',['Avanza a competencias','Competencias agendada']],['Empresa',['Avanza a screening empresa','Screening empresa agendado','Avanza a presencial empresa','Avanza a empresa','Entrevista empresa agendada','En cliente']],['Contratado',['Contratado']]];
const BASE={'Análisis de CV':'Nuevo','Screening':'Avanza a screening','Competencias':'Avanza a competencias','Empresa':'Avanza a presencial empresa','Contratado':'Contratado'};
const DESC_TIPOS=['Descartado sin screening','Descartado post screening','Descartado en competencias','Descartado en screening empresa','Descartado presencial empresa','Descartado Fuera de zona','Descartado NO le interesa'];
const RES_TIPOS=['En reserva','Reserva competencias','Reserva screening empresa','Reserva presencial empresa'];
const nota=n=>n?'<span class="nota n'+String(n).replace('+','')+'">'+n+'</span>':'';
const tag=e=>/agendad/i.test(e)?'<span class="tag ag">Agendado</span>':'';
const card=c=>'<div class="card" data-fila="'+c.fila+'" onclick="ficha('+c.fila+')"><button class="cx" title="Eliminar" onclick="event.stopPropagation();eliminar('+c.fila+')">✕</button>'+nota(c.notaPresEmp||c.notaScrEmp||c.notaComp||c.notaScr||c.nota)+'<div class="n">'+c.nombre+'</div><div class="meta">'+(c.zona||'—')+(c.pretension?' · $'+c.pretension:'')+'</div>'+(c.formacion&&c.formacion!=='Sin dato'?'<div class="meta" title="'+String(c.formacion).replace(/"/g,'')+'">🎓 '+String(c.formacion).slice(0,70)+'</div>':'')+(c.cita?'<div class="cita-tag">📅 '+c.cita+'</div>':'')+tagsTarjeta(c)+(c.contactado?'<span class="tag" style="background:#D1FAE5;color:#065F46">✓ Contact. '+c.contactado+' <button title="Borrar la marca de contacto" style="border:none;background:none;color:#065F46;cursor:pointer;font-weight:800;padding:0 0 0 4px" onclick="event.stopPropagation();borrarMarca('+c.fila+')">✕</button></span>':'')+tag(c.estado)+'</div>';
const board=document.getElementById('board');
function acordeon(titulo,clase,tipos,extras,corto){
  var lista=tipos.slice();
  DATA.forEach(function(c){if(extras(c.estado)&&lista.indexOf(c.estado)===-1)lista.push(c.estado);});
  var total=DATA.filter(function(c){return extras(c.estado);}).length;
  var h='<div class="col '+clase+'"><div class="col-h"><span>'+titulo+'</span><b>'+total+'</b></div><div class="acc">';
  lista.forEach(function(tipo){
    var items=DATA.filter(function(c){return c.estado===tipo;});
    var etq=corto?tipo.replace(corto,'').trim():tipo;
    h+='<details><summary><span>▸ '+(etq||tipo)+'</span><span class="cnt">'+items.length+'</span></summary><div class="dbody" data-estado="'+tipo+'">'+items.map(card).join('')+'</div></details>';
  });
  return h+'</div></div>';
}
COLS.forEach(function(x){
  var t=x[0],est=x[1];
  var items=DATA.filter(function(c){return est.indexOf(c.estado)>-1;});
  board.innerHTML+='<div class="col"><div class="col-h"><span>'+t+'</span><b>'+items.length+'</b></div><div class="col-body" data-estado="'+BASE[t]+'">'+items.map(card).join('')+'</div></div>';
  if(t==='Análisis de CV')board.innerHTML+=acordeon('Descartados','desc',DESC_TIPOS,function(e){return /^Descartad/.test(e);},'Descartado');
});
board.innerHTML+=acordeon('Reservas','res',RES_TIPOS,function(e){return /reserva/i.test(e);},'');
function toast(){var s=document.getElementById('saved');s.classList.add('show');setTimeout(function(){s.classList.remove('show');},1500);}
document.querySelectorAll('.col-body,.dbody').forEach(function(col){new Sortable(col,{group:'kb',animation:150,forceFallback:true,fallbackOnBody:true,scroll:true,bubbleScroll:true,scrollSensitivity:80,scrollSpeed:15,onEnd:function(evt){var fila=+evt.item.dataset.fila;var estado=evt.to.dataset.estado;var cc=DATA_ALL.find(function(x){return x.fila===fila;});if(cc&&estado)cc.estado=estado;if(estado&&evt.from!==evt.to){google.script.run.withSuccessHandler(function(){toast();if(/^Descartad/i.test(estado))pedirMotivo(fila,true);}).withFailureHandler(function(e){alert('Error: '+e.message);window.top.location.reload();}).moverEstado(PERFIL,fila,estado,quienSoy());}}});});
function nrm(s){return String(s||'').toLowerCase().normalize('NFD').replace(/[\\u0300-\\u036f]/g,'');}
function raiz(p){p=nrm(p);if(p.length>5)return p.slice(0,-2);if(p.length>3)return p.slice(0,-1);return p;}
function coincide(texto,consulta){if(!consulta.trim())return true;var t=nrm(texto);var ps=consulta.trim().split(/\\s+/);for(var i=0;i<ps.length;i++){if(t.indexOf(raiz(ps[i]))===-1)return false;}return true;}
function filtrar(){
  var fNom=document.getElementById('fNombre').value;
  var fZon=document.getElementById('fZona').value;
  var fFor=document.getElementById('fForm').value;
  var fExp=document.getElementById('fExp').value;
  var fEst=document.getElementById('fEstado').value;
  var fPre=Number(document.getElementById('fPret').value)||0;
  var hay=fNom||fZon||fFor||fExp||fEst||fPre;
  var visibles=0;
  document.querySelectorAll('.card').forEach(function(el){
    var c=DATA.find(function(x){return x.fila===+el.dataset.fila;});
    var ok=true;
    if(c){
      if(ok&&fNom&&!coincide(c.nombre,fNom))ok=false;
      if(ok&&fZon&&!coincide(c.zona+' '+c.requisitos,fZon))ok=false;
      if(ok&&fFor&&!coincide(c.formacion+' '+c.requisitos+' '+c.resumen,fFor))ok=false;
      if(ok&&fExp&&!coincide(c.resumen+' '+c.requisitos+' '+c.formacion,fExp))ok=false;
      if(ok&&fEst&&c.estado!==fEst)ok=false;
      if(ok&&fPre&&(!c.pretNum||c.pretNum>fPre))ok=false;
    }
    el.style.display=ok?'':'none';
    if(ok)visibles++;
  });
  document.querySelectorAll('.acc details').forEach(function(d){
    var vis=d.querySelectorAll('.card:not([style*="none"])').length;
    d.querySelector('.cnt').textContent=vis;
    if(hay)d.open=vis>0;
  });
  document.getElementById('contadorF').textContent=hay?visibles+' resultado'+(visibles===1?'':'s'):'';
}
function limpiarFiltros(){['fNombre','fZona','fForm','fExp','fPret'].forEach(function(id){document.getElementById(id).value='';});document.getElementById('fEstado').value='';document.querySelectorAll('.acc details').forEach(function(d){d.open=false;});filtrar();}
(function(){var est=[];DATA.forEach(function(c){if(c.estado&&est.indexOf(c.estado)===-1)est.push(c.estado);});est.sort();var s=document.getElementById('fEstado');est.forEach(function(e){var o=document.createElement('option');o.value=e;o.text=e;s.appendChild(o);});})();
const MOTIVOS=${JSON.stringify(MOTIVOS_DESCARTE_)};
function tagsTarjeta(c){
  var h='',T={contactar:'Falta contactar',norespondio:'Sin respuesta',pronto:'Entrevista pronto',sincargar:'Falta cargar formulario',sinfecha:'No está en el calendario',reserva:'En reserva hace tiempo',faltascreening:'Falta screening',faltacompetencias:'Falta competencias'};
  if(/^Descartad/i.test(c.estado)){h+=c.motivo?'<span class="tag" style="background:#FDE2E2;color:#9B1C1C" title="Motivo de descarte">✖ '+String(c.motivo).replace(/;\\s*/g,' · ')+'</span>':'<span class="tag" style="background:#FFF3CD;color:#7a5500;cursor:pointer" onclick="event.stopPropagation();pedirMotivo('+c.fila+')">Sin motivo · cargar</span>';}
  if(c.mailDescarte)h+='<span class="tag">✉ Descarte enviado</span>';
  try{avPendientes([c]).filter(function(x){return x.k.indexOf('sinmov')!==0;}).forEach(function(x){var kk=x.k.split('|')[0],rj=x.urg||kk.indexOf('falta')===0&&kk!=='falta',dn=kk==='norespondio'?(String(x.txt).match(/hace (\\d+) d/)||[])[1]:'';h+='<span class="tag" style="background:'+(rj?'#FDE2E2':'#FFF3CD')+';color:'+(rj?'#9B1C1C':'#7a5500')+'">'+x.ic+' '+(dn?'Sin respuesta hace '+dn+' d':(T[kk]||x.txt))+'</span>';});}catch(e){}
  return h;
}
function pedirMotivo(f,recargar){
  var c=DATA_ALL.find(function(x){return x.fila===f;})||{};var ya=String(c.motivo||'');
  var h='<div style="font-size:13px;color:#555;margin-bottom:10px">Elegí el o los motivos del descarte de <b>'+c.nombre+'</b>.</div>';
  MOTIVOS.forEach(function(m,i){h+='<label style="display:flex;gap:8px;align-items:center;margin:6px 0;font-size:14px"><input type="checkbox" class="mot" value="'+m+'"'+(ya.indexOf(m)>-1?' checked':'')+'> '+m+'</label>';});
  var otro=(ya.match(/Otro: ([^;]+)/)||[])[1]||'';
  h+='<label style="display:block;margin-top:8px;font-size:13px">Otro (máximo tres palabras)<input id="motOtro" style="width:100%;padding:7px;border:1px solid #ccc;border-radius:6px;margin-top:3px" value="'+otro+'"></label>';
  h+='<div style="display:flex;gap:8px;justify-content:flex-end;margin-top:14px"><button class="mbtn alt" onclick="cerrar();'+(recargar?'window.top.location.reload()':'')+'">Después</button><button class="mbtn" style="background:#1A6645;color:#fff" onclick="guardarMotivo('+f+','+(recargar?1:0)+',this)">Guardar motivo</button></div>';
  document.getElementById('mtit').textContent='Motivo de descarte';document.getElementById('mbody').innerHTML=h;document.getElementById('modal').classList.add('show');
}
function guardarMotivo(f,rec,btn){
  var sel=[].slice.call(document.querySelectorAll('.mot:checked')).map(function(x){return x.value;});var o=document.getElementById('motOtro').value;
  btn.disabled=true;btn.textContent='Guardando...';
  google.script.run.withSuccessHandler(function(){toast();cerrar();window.top.location.reload();}).withFailureHandler(function(e){btn.disabled=false;btn.textContent='Guardar motivo';alert(e.message);}).guardarMotivoDescarte(PERFIL,f,sel,o,quienSoy());
}
function ficha(f){var c=DATA.find(function(x){return x.fila===f;});var row=function(l,v){return v?'<div class="frow"><span class="fl">'+l+'</span><span class="fv">'+v+'</span></div>':'';};var btn=function(u,t){return u?'<a href="'+u+'" target="_blank" class="mbtn">'+t+'</a>':'';};var waBtn='<button class="mbtn wa" onclick="menuWA('+c.fila+')">💬 WhatsApp ▾</button>';var mailBtn=(c.mail&&c.mail!=='Sin dato'&&c.mail.indexOf('@')>-1)?'<a href="https://mail.google.com/mail/?view=cm&to='+c.mail+'" target="_blank" class="mbtn" style="background:#EA4335;color:#fff">✉ Mail</a>':'';var cvBtn=c.cv?'<a href="https://drive.google.com/file/d/'+c.cv+'/view" target="_blank" class="mbtn alt">Ver CV</a>':'';var genBtn=c.informe?'':'<button class="mbtn gen" onclick="var b=this;blConfirm(\\'Generar informe con IA. Puede tener costo. ¿Continuar?\\',function(){generar('+c.fila+',b)})">Generar informe</button>';var delBtn='<button class="mbtn" style="background:#B91C1C;color:#fff" onclick="eliminar('+c.fila+')">🗑 Eliminar</button>';var bancoBtn='<button class="mbtn" style="background:#7C3AED" onclick="menuBanco('+c.fila+')">★ Banco de talento</button>';var notasEmp=/screening empresa|presencial empresa|avanza a empresa|entrevista empresa|reserva screening empresa|reserva presencial empresa/i.test(c.estado)?['screening-empresa','presencial-empresa'].map(function(ep){return '<div class="notarow"><span>'+(ep==='screening-empresa'?'Emp. screening':'Emp. presencial')+'</span>'+['A','B','C'].map(function(n){return '<button class="nb" onclick="setNotaEmp('+c.fila+',\\''+ep+'\\',\\''+n+'\\',this)">'+n+'</button>';}).join('')+'</div>';}).join(''):'';document.getElementById('mbody').innerHTML=row('Recibido',c.recibido)+row('Estado',c.estado)+(/^Descartad/i.test(c.estado)?'<div class="frow"><span class="fl">Motivo de descarte</span><span class="fv">'+(c.motivo?String(c.motivo).replace(/;\\s*/g,' · '):'sin cargar')+' <a href="#" onclick="pedirMotivo('+c.fila+');return false" style="font-size:12px">'+(c.motivo?'editar':'cargar')+'</a></span></div>':'')+row('Cita agendada',c.cita)+(c.historial?'<div class="frow"><span class="fl">Historial entrevistas</span><span class="fv" style="white-space:pre-line">'+c.historial+'</span></div>':'')+(c.contactado?'<div class="frow"><span class="fl">Último contacto</span><span class="fv">'+c.contactado+' <a href="#" onclick="borrarMarca('+c.fila+');return false" style="color:#aaa;font-size:11px;text-decoration:none;margin-left:6px">✕ borrar</a></span></div>':'')+'<div class="frow"><span class="fl">Registro de contactos</span><span class="fv" id="regCont" style="white-space:pre-line">cargando…</span></div><div class="frow"><span class="fl">Historial</span><span class="fv" id="histCand">cargando…</span></div>'+row('Teléfono',c.tel)+row('Email',c.mail)+row('Zona',c.zona)+row('Formación',c.formacion)+row('Pretensión',c.pretension?'$'+c.pretension:'')+row('Nota screening',c.notaScr)+row('Nota competencias',c.notaComp)+row('Nota screening empresa',c.notaScrEmp)+row('Nota presencial empresa',c.notaPresEmp)+'<div class="fres">'+(c.resumen||'Sin resumen')+'</div><div class="notabox"><div class="notatit">Poner nota manual</div>'+['cv','screening','competencias'].map(function(et){return '<div class="notarow"><span>'+({cv:'CV',screening:'Screening',competencias:'Competencias'}[et])+'</span>'+['A','B+','B','C'].map(function(n){return '<button class="nb" onclick="setNota('+c.fila+',\\''+et+'\\',\\''+n+'\\',this)">'+n+'</button>';}).join('')+'</div>';}).join('')+notasEmp+'</div><div class="mbtns">'+waBtn+mailBtn+cvBtn+btn(c.scrPDF,'Ver screening')+btn(c.compPDF,'Ver competencias')+btn(c.postulacion,'Ver postulación')+btn(c.scrLink,'Screening')+btn(c.compLink,'Competencias')+btn(c.informe,'Ver informe')+genBtn+bancoBtn+delBtn+'</div>';document.getElementById('mtit').textContent=c.nombre;document.getElementById('modal').classList.add('show');cargarRegCont(f);cargarHist(f);}
function histLin(x,conNombre){return '<div style="font-size:12px;line-height:1.4;padding:3px 0;border-bottom:1px solid #f0eee9"><b>'+x.fecha+'</b> · '+(conNombre?'<b>'+avEsc(x.candidato)+'</b> · ':'')+avEsc(x.tipo)+': '+avEsc(x.detalle)+' <span style="color:#1A6645;font-weight:700">· '+avEsc(x.quien)+'</span></div>';}
function cargarHist(f){google.script.run.withSuccessHandler(function(l){var e=document.getElementById('histCand');if(!e)return;e.innerHTML=l.length?'<details><summary style="cursor:pointer;font-weight:700">'+l.length+' movimiento'+(l.length===1?'':'s')+'</summary><div style="max-height:260px;overflow:auto;margin-top:6px">'+l.map(function(x){return histLin(x,false);}).join('')+'</div></details>':'Sin movimientos registrados todavía';}).withFailureHandler(function(){var e=document.getElementById('histCand');if(e)e.textContent='—';}).getHistorialCandidato(PERFIL,f);}
function verHistorial(){document.getElementById('mtit').textContent='📜 Historial de la búsqueda';document.getElementById('mbody').innerHTML='Cargando…';document.getElementById('modal').classList.add('show');google.script.run.withSuccessHandler(function(l){document.getElementById('mbody').innerHTML='<div style="font-size:12px;color:#666;margin-bottom:8px">Últimos '+l.length+' movimientos: quién tocó qué y cuándo.</div>'+(l.length?l.map(function(x){return histLin(x,true);}).join(''):'Todavía no hay movimientos registrados.');}).withFailureHandler(function(e){document.getElementById('mbody').textContent='No se pudo cargar: '+e.message;}).getHistorialBusqueda(PERFIL);}
function cargarRegCont(f){google.script.run.withSuccessHandler(function(l){var e=document.getElementById('regCont');if(!e)return;e.innerHTML=(l.length?l.map(function(x){return avEsc(x.fecha+' · '+x.tipo+(x.quien?' · '+x.quien:''))+' <a href="#" title="Borrar este contacto" onclick="borrarRegCont('+f+',\\''+x.fecha+'\\');return false" style="color:#aaa;font-size:11px;text-decoration:none">✕</a>';}).join('<br>'):'Sin contactos registrados')+' <a href="#" onclick="regContacto('+f+');return false" style="font-size:12px">+ registrar</a>';}).withFailureHandler(function(){var e=document.getElementById('regCont');if(e)e.textContent='—';}).getContactosRegistro(PERFIL,f);}
function borrarRegCont(f,fe){blConfirm('¿Borrar este contacto del registro?',function(){google.script.run.withSuccessHandler(function(){cargarRegCont(f);}).withFailureHandler(function(e){alert('Error: '+e.message);}).borrarContactoRegistro(PERFIL,f,fe);},{si:'Borrar',peligro:true});}
function regContacto(f){blPrompt('¿Qué contacto hiciste? (ej: Llamé, Dejé mensaje, Respondió por mail)','',function(t){var qs=quienSoy();google.script.run.withSuccessHandler(function(){cargarRegCont(f);}).withFailureHandler(function(e){alert('Error: '+e.message);}).registrarContactoManual(PERFIL,f,t,qs);},{titulo:'Registrar contacto'});}
function bajarContactos(b){b.disabled=true;google.script.run.withSuccessHandler(function(r){b.disabled=false;var a=document.createElement('a');a.href=URL.createObjectURL(new Blob([r.csv],{type:'text/csv'}));a.download='contactos-blumia.csv';a.click();alert(r.n+' contactos de esta búsqueda (los que ya contactaste o avanzaron). Importalo en contacts.google.com → Importar.');}).withFailureHandler(function(e){b.disabled=false;alert('Error: '+e.message);}).exportarContactosCSV(PERFIL);}
function menuWA(f){
  var c=DATA.find(function(x){return x.fila===f;});
  var ops=[['contacto-screening','📞 Contacto screening'],['contacto-competencias','💼 Contacto competencias'],['contacto-empresa-virtual','🏢 Empresa (virtual)'],['contacto-empresa-presencial','🏢 Empresa (presencial)'],['descarte-competencias','❌ Descarte competencias'],['descarte-empresa','❌ Descarte empresa']];
  var h='<div style="font-size:13px;margin-bottom:12px;color:#555">Elegí qué mensaje enviar a <b>'+c.nombre+'</b>:</div>';
  ops.forEach(function(o){h+='<button class="mbtn wa" style="display:block;width:100%;margin-bottom:8px;text-align:left" onclick="enviarWA('+f+',\\''+o[0]+'\\',this)">'+o[1]+'</button>';});
  h+='<button class="mbtn alt" style="width:100%;margin-top:6px" onclick="ficha('+f+')">← Volver</button>';
  document.getElementById('mtit').textContent='WhatsApp · '+c.nombre;
  document.getElementById('mbody').innerHTML=h;
}
function quienSoy(){var q='';try{q=localStorage.getItem('blumia_quien')||'';}catch(e){}return q;}
function enviarWA(f,tipo,btn){
  btn.textContent='Preparando...';
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.ok){window.open(r.url,'_blank');btn.textContent='✓ Abierto';}
    else{alert('Error: '+(r?r.error:'desconocido'));}
  }).withFailureHandler(function(e){alert('Error: '+e.message);}).armarComunicacion(PERFIL,f,tipo,quienSoy());
}
function menuBanco(f){
  var c=DATA.find(function(x){return x.fila===f;});
  var CATS=['RRHH / Gestión de Personas','Datos / BI','Compras / Abastecimiento','Calidad / SGI','Administración / Finanzas','Operaciones / Producción','Comercial / Ventas','Otros'];
  var SEN=['Jr','Ssr','Sr','Coordinador','Supervisor','Jefatura / Gerencia'];
  var h='<div style="font-size:13px;margin-bottom:14px;color:#555">Enviar a <b>'+c.nombre+'</b> al banco de talento.</div>';
  h+='<div class="notatit">Categoría</div><select id="bCat" style="width:100%;padding:10px;border:1px solid #DDD;border-radius:8px;font-family:inherit;font-size:13px;margin-bottom:12px">'+CATS.map(function(x){return '<option>'+x+'</option>';}).join('')+'</select>';
  h+='<div class="notatit">Seniority</div><select id="bSen" style="width:100%;padding:10px;border:1px solid #DDD;border-radius:8px;font-family:inherit;font-size:13px;margin-bottom:12px">'+SEN.map(function(x){return '<option>'+x+'</option>';}).join('')+'</select>';
  h+='<div class="notatit">Nota (opcional)</div><textarea id="bNota" placeholder="Por qué lo guardás, para qué tipo de búsqueda sirve..." style="width:100%;height:70px;padding:8px;border:1px solid #DDD;border-radius:8px;font-family:inherit;font-size:12px;margin-bottom:12px"></textarea>';
  h+='<div class="mbtns"><button class="mbtn" style="background:#7C3AED" onclick="guardarBanco('+f+',this)">★ Guardar en banco</button><button class="mbtn alt" onclick="ficha('+f+')">← Volver</button></div>';
  document.getElementById('mtit').textContent='Banco de talento';
  document.getElementById('mbody').innerHTML=h;
}
function guardarBanco(f,btn){
  var cat=document.getElementById('bCat').value;
  var sen=document.getElementById('bSen').value;
  var nota=document.getElementById('bNota').value;
  btn.textContent='Guardando...';btn.disabled=true;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.ok){btn.textContent='✓ Guardado en banco';btn.style.background='#999';}
    else{alert(r?r.error:'Error');btn.textContent='★ Guardar en banco';btn.disabled=false;}
  }).withFailureHandler(function(e){alert('Error: '+e.message);btn.textContent='★ Guardar en banco';btn.disabled=false;}).enviarABanco(PERFIL,f,cat,sen,nota,quienSoy());
}
function setNota(f,et,n,btn){btn.textContent='...';google.script.run.withSuccessHandler(function(r){toast();cerrar();if(r&&r.estado&&/^Descartad/i.test(r.estado)){pedirMotivo(f,true);}else setTimeout(function(){window.top.location.reload();},800);}).withFailureHandler(function(e){alert('Error: '+e.message);}).notaManual(PERFIL,f,et,n,quienSoy());}
function setNotaEmp(f,et,n,btn){btn.textContent='...';google.script.run.withSuccessHandler(function(r){toast();cerrar();if(r&&r.estado&&/^Descartad/i.test(r.estado)){pedirMotivo(f,true);}else setTimeout(function(){window.top.location.reload();},800);}).withFailureHandler(function(e){alert('Error: '+e.message);}).notaEmpresa(PERFIL,f,et,n,quienSoy());}
function generar(f,btn){btn.textContent='Generando...';btn.disabled=true;google.script.run.withSuccessHandler(function(url){window.open(url,'_blank');btn.textContent='✓ Listo';}).withFailureHandler(function(e){alert('Error: '+e.message);btn.textContent='Generar informe';btn.disabled=false;}).generarInformeCandidato(PERFIL,f);}
function descartesMasivo(){google.script.run.withSuccessHandler(function(p){if(p.sinScr===0&&p.postScr===0){alert('No hay descartes pendientes de mail en esta búsqueda.');return;}var h='<div style="font-size:13px;margin-bottom:14px">Se van a enviar <b>'+(p.sinScr+p.postScr)+'</b> mails en '+p.puesto+'.<br>Sin screening: <b>'+p.sinScr+'</b> · Post screening: <b>'+p.postScr+'</b><br><span style="color:#888">[NOMBRE] se reemplaza por el nombre de cada persona.</span></div>';h+='<div class="notatit">Texto · Descartado sin screening ('+p.sinScr+')</div><textarea id="txtSin" style="width:100%;height:150px;font-family:inherit;font-size:12px;padding:8px;border:1px solid #DDD;border-radius:6px;margin-bottom:12px">'+p.textoSin+'</textarea>';h+='<div class="notatit">Texto · Descartado post screening ('+p.postScr+')</div><textarea id="txtPost" style="width:100%;height:150px;font-family:inherit;font-size:12px;padding:8px;border:1px solid #DDD;border-radius:6px;margin-bottom:12px">'+p.textoPost+'</textarea>';h+='<div class="mbtns"><button class="mbtn" style="background:#E06666" onclick="enviarDesc(this)">✉ Enviar '+(p.sinScr+p.postScr)+' mails</button></div>';document.getElementById('mtit').textContent='Mails de descarte';document.getElementById('mbody').innerHTML=h;document.getElementById('modal').classList.add('show');}).withFailureHandler(function(e){alert('Error: '+e.message);}).previaDescartes(PERFIL);}
function enviarDesc(btn){blConfirm('¿Confirmás el envío? No se puede deshacer.',function(){var s=document.getElementById('txtSin').value;var p=document.getElementById('txtPost').value;btn.textContent='Enviando...';btn.disabled=true;google.script.run.withSuccessHandler(function(r){blAlert('Enviados: '+r.enviados+(r.sinMail?' · Sin mail válido: '+r.sinMail:''),function(){window.top.location.reload();});}).withFailureHandler(function(e){alert('Error: '+e.message);btn.disabled=false;btn.textContent='✉ Enviar';}).enviarDescartesMasivo(PERFIL,s,p,quienSoy());},{si:'Enviar',peligro:true});}
function borrarMarca(f){blConfirm('¿Borrar la marca de contacto?',function(){google.script.run.withSuccessHandler(function(){toast();cerrar();setTimeout(function(){window.top.location.reload();},800);}).withFailureHandler(function(e){alert('Error: '+e.message);}).borrarMarcaContacto(PERFIL,f,quienSoy());},{si:'Borrar',peligro:true});}
/* ===== Avisos 🔔: pendientes y novedades (igual que PAMI) ===== */
var AV_PEND,AV_NOV;
function avLS(k,v){try{if(v===undefined)return JSON.parse(localStorage.getItem('blumia_av_'+PERFIL+'_'+k)||'null');localStorage.setItem('blumia_av_'+PERFIL+'_'+k,JSON.stringify(v));}catch(e){return null;}}
function avEsc(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
function avFecha(s){var m=String(s||'').match(/(\\d{2})\\/(\\d{2})\\/(\\d{4})(?:[^\\d]+(\\d{2}):(\\d{2}))?/);if(!m)return null;return new Date(+m[3],+m[2]-1,+m[1],+(m[4]||0),+(m[5]||0)).getTime();}
function avActivo(c){return !/^(Descartad|Eliminado|Contratado|En cliente)/i.test(c.estado)&&!/reserva/i.test(c.estado);}
function avEtapa(e){
  if(/^(Avanza a screening empresa|Screening empresa)/i.test(e))return {k:'empresa',n:'el screening con la empresa',h:'Empresa'};
  if(/^(Avanza a presencial empresa|Avanza a empresa|Entrevista empresa)/i.test(e))return {k:'empresa',n:'la entrevista con la empresa',h:'Empresa'};
  if(/^(Avanza a competencias|Competencias)/i.test(e))return {k:'competencias',n:'competencias',h:'Competencias'};
  if(/^(Avanza a screening|Screening)/i.test(e))return {k:'screening',n:'screening',h:'Screening'};
  return null;
}
function avEtapasFaltan(c){var m=typeof metEstado==='function'?metEstado(c.estado):null,f=[];if(!m||m.t!=='act')return f;if(m.lv>=2&&!String(c.notaScr||'').trim())f.push('screening');if(m.lv>=3&&!String(c.notaComp||'').trim())f.push('competencias');return f;}
function avPendientes(L){
  var out=[],now=Date.now();
  L.forEach(function(c){
    if(!avActivo(c))return;
    var p=function(k,ic,txt,urg){out.push({k:k+'|'+c.fila,ic:ic,txt:txt,f:c.fila,n:c.nombre,urg:!!urg});};
    avEtapasFaltan(c).forEach(function(e){p('falta'+e,'⚠','Falta la etapa de '+e+' (no tiene nota cargada)');});
    var ep=avEtapa(c.estado);if(!ep)return;
    if(!/agendad/i.test(c.estado)){
      var ct=String(c.contactado||'').split(' · ')[0].trim().toLowerCase();
      if(ct!==ep.k){p('contactar','➜','Falta contactar para '+ep.n);return;}
      var f=avFecha(c.contactado);
      if(!c.cita&&f){var d=Math.floor((now-f)/864e5);if(d>=2){p('norespondio','⏳','Contactado para '+ep.n+' hace '+d+' días y todavía no agendó');return;}}
      var mv=c.movMs,dm=mv?Math.floor((now-mv)/864e5):0;
      if(dm>=7)p('sinmov','🕒',dm+' días sin movimiento en «'+c.estado+'»');
      return;
    }
    if(c.cita){
      var t=avFecha(c.cita);
      if(t&&t>=now-15*6e4&&t-now<=30*6e4)p('pronto','⏰','Entrevista a las '+String(c.cita).slice(-5)+(t<now?' (ya empezó)':' (en '+Math.max(1,Math.round((t-now)/6e4))+' min)'),true);
      return;
    }
    var lines=String(c.historial||'').split('\\n').filter(function(l){return l.indexOf(ep.h)===0;});
    var f2=lines.length?avFecha(lines[0]):null;
    if(f2){if(f2<now-60*6e4)p('sincargar','📝','La entrevista de '+ep.n+' del '+String(lines[0]).split(' — ')[1]+' ya pasó y falta cargar el formulario',true);}
    else p('sinfecha','📅','Figura agendado para '+ep.n+' pero no está en el calendario');
  });
  L.forEach(function(c){if(!/reserva/i.test(c.estado)||/^Descartad|^Eliminado/i.test(c.estado))return;var mv=c.movMs,dm=mv?Math.floor((now-mv)/864e5):0;if(dm>=14)out.push({k:'reserva|'+c.fila,ic:'📌',txt:'En reserva hace '+dm+' días: ¿se retoma o se descarta?',f:c.fila,n:c.nombre,urg:false});});
  out.sort(function(a,b){return (b.urg?1:0)-(a.urg?1:0);});
  return out;
}
function avNovedades(L){
  var prev=avLS('snap'),snap={},nov=[],d=new Date(),dd=function(n){return ('0'+n).slice(-2);},cuando=dd(d.getDate())+'/'+dd(d.getMonth()+1)+' '+dd(d.getHours())+':'+dd(d.getMinutes());
  L.forEach(function(c){
    var s={e:c.estado,c:c.cita||'',n:(c.notaScr||'')+'|'+(c.notaComp||'')};snap[c.fila]=s;if(!prev)return;
    var a=prev[c.fila],p=function(ic,txt){nov.push({ic:ic,txt:txt,f:c.fila,n:c.nombre,h:cuando});};
    if(!a){if(!/^Eliminado/.test(s.e))p('🆕','Nuevo CV recibido');return;}
    if(s.e!==a.e&&!/^Eliminado/.test(s.e)){
      if(s.e==='Contratado')p('🎉','Contratado/a');
      else if(/^Descartad/i.test(s.e))p('❌','Descartado/a: '+s.e.replace(/^Descartado\\s*/i,''));
      else p('🔄','Pasó a «'+s.e+'»'+(a.e?' (antes: '+a.e+')':''));
    }
    if(s.c&&s.c!==a.c)p('📅','Agendó para el '+s.c);
    if(a.n!==undefined&&s.n!==a.n&&s.n!=='|')p('📝','Nota cargada · screening '+(c.notaScr||'—')+' · competencias '+(c.notaComp||'—'));
  });
  avLS('snap',snap);return nov;
}
function avPintar(){var b=document.getElementById('avCnt');if(!b)return;var n=AV_PEND.length+(avLS('sinver')||0),u=AV_PEND.filter(function(x){return x.urg;}).length;b.textContent=n||'';b.style.display=n?'':'none';var bt=document.getElementById('avBtn');if(bt)bt.style.background=u?'#B91C1C':'';}
function avAbrir(f){if(DATA.some(function(x){return x.fila===f;}))ficha(f);else cerrar();}
function verAvisos(){
  avLS('sinver',0);avPintar();
  var item=function(x,nov){return '<div style="cursor:pointer;display:block;text-align:left;padding:10px 6px;border-bottom:1px solid #eee" onclick="avAbrir('+x.f+')"><div style="font-weight:700;font-size:14px">'+x.ic+' '+avEsc(x.n)+'</div><div style="font-size:12px;color:#555;margin-top:2px">'+avEsc(x.txt)+(nov?' · '+x.h:'')+'</div></div>';};
  document.getElementById('mtit').textContent='🔔 Avisos';
  document.getElementById('mbody').innerHTML='<div style="font-weight:800;color:#B91C1C;margin:2px 0 6px">Pendientes ('+AV_PEND.length+')</div>'+(AV_PEND.length?AV_PEND.map(function(x){return item(x,false);}).join(''):'<div style="color:#777;font-size:14px">No hay pendientes. 👌</div>')+'<div style="font-weight:800;color:#1A6645;margin:16px 0 6px">Novedades</div>'+(AV_NOV.length?AV_NOV.slice(0,40).map(function(x){return item(x,true);}).join(''):'<div style="color:#777;font-size:14px">Sin novedades desde la última visita.</div>');
  document.getElementById('modal').classList.add('show');
}
function avPopups(L){
  var cont=document.getElementById('avPop');
  if(!cont){cont=document.createElement('div');cont.id='avPop';cont.style.cssText='position:fixed;bottom:14px;left:14px;z-index:150;display:flex;flex-direction:column-reverse;gap:8px;max-width:min(340px,calc(100vw - 110px))';document.body.appendChild(cont);}
  var mk=function(x){var d=document.createElement('div');d.style.cssText='background:#fff;border-left:4px solid '+(x.urg?'#B91C1C':'#1A6645')+';box-shadow:0 4px 14px rgba(0,0,0,.2);border-radius:6px;padding:10px 12px;font-size:13px;cursor:pointer';d.innerHTML='<button type="button" aria-label="Cerrar aviso" style="float:right;border:none;background:none;font-size:16px;line-height:1;cursor:pointer;color:#777;margin:-4px -6px 0 8px;padding:4px" onclick="event.stopPropagation();this.parentNode.remove()">✕</button><b>'+x.ic+' '+avEsc(x.n)+'</b><br>'+avEsc(x.txt);d.onclick=function(){d.remove();if(x.panel)verAvisos();else if(x.f)avAbrir(x.f);};cont.appendChild(d);setTimeout(function(){if(d.parentNode)d.remove();},x.urg?15000:8000);};
  L.slice(0,5).forEach(mk);
  if(L.length>5)mk({ic:'🔔',n:'Y '+(L.length-5)+' más',txt:'Tocá para ver la lista',panel:1});
}
function avRevisar(L){
  AV_PEND=avPendientes(L);
  var nuevas=avNovedades(L),hist=avLS('nov')||[];
  if(nuevas.length){hist=nuevas.concat(hist).slice(0,80);avLS('nov',hist);avLS('sinver',(avLS('sinver')||0)+nuevas.length);}
  AV_NOV=hist;
  var avisados=avLS('urg')||{},hoy=new Date().toISOString().slice(0,10),pop=nuevas.slice(),vistos=avLS('pseen'),act={};
  AV_PEND.forEach(function(x){act[x.k]=1;if(x.urg&&avisados[x.k]!==hoy){avisados[x.k]=hoy;pop.push(x);}else if(vistos&&!vistos[x.k])pop.push(x);});
  avLS('urg',avisados);avLS('pseen',act);
  try{if(!sessionStorage.getItem('blumia_av_ses')&&AV_PEND.length){sessionStorage.setItem('blumia_av_ses','1');var nu=AV_PEND.filter(function(x){return x.urg;}).length;pop.unshift({ic:'🔔',n:'Tenés '+AV_PEND.length+' pendiente'+(AV_PEND.length===1?'':'s'),txt:(nu?nu+' urgente'+(nu===1?'':'s')+' · ':'')+'tocá para ver la lista',panel:1,urg:nu>0});}}catch(e){}
  avPintar();if(pop.length)avPopups(pop);
}
function avRefrescar(){
  google.script.run.withSuccessHandler(function(r){
    var nuevo;try{nuevo=JSON.parse(r);}catch(e){return;}
    var key=function(L){return L.map(function(c){return c.fila+'|'+c.estado;}).sort().join(',');};
    var cambio=key(nuevo)!==key(DATA_ALL);
    avRevisar(nuevo);
    if(cambio&&!document.getElementById('avBanner')){var b=document.createElement('div');b.id='avBanner';b.style.cssText='position:fixed;bottom:16px;left:50%;transform:translateX(-50%);background:#1A6645;color:#fff;padding:10px 18px;border-radius:20px;font-size:13px;font-weight:700;cursor:pointer;z-index:9999;box-shadow:0 4px 14px rgba(0,0,0,.25)';b.textContent='🔄 Hay cambios en el tablero · tocá para actualizar';b.onclick=function(){window.top.location.reload();};document.body.appendChild(b);}
  }).withFailureHandler(function(){}).doGetKanban(PERFIL);
}
function avInit(){avRevisar(DATA_ALL);setInterval(avRefrescar,180000);}
function eliminar(f){var c=DATA.find(function(x){return x.fila===f;});if(!c)return;blConfirm('¿Eliminar a '+c.nombre+' del tablero?\\nNo se pierde nada: queda en la 🗑 Papelera y se puede restaurar.',function(){google.script.run.withSuccessHandler(function(){toast();cerrar();setTimeout(function(){window.top.location.reload();},700);}).withFailureHandler(function(e){alert('Error: '+e.message);}).eliminarCandidato(PERFIL,f,quienSoy());},{si:'Eliminar',peligro:true});}
function abrirPapelera(){document.getElementById('mtit').textContent='Papelera';var h=ELIM.length?'':'<div style="color:#777;font-size:14px">La papelera está vacía.</div>';ELIM.forEach(function(c){var p=String(c.estado).split('|');h+='<div class="frow" style="align-items:center"><span class="fv"><b>'+c.nombre+'</b><br><small>'+(p[1]||'sin estado')+(p[2]?' · eliminado el '+p[2]:'')+'</small></span><button class="mbtn" onclick="restaurar('+c.fila+')">↩ Restaurar</button></div>';});document.getElementById('mbody').innerHTML=h;document.getElementById('modal').classList.add('show');}
function restaurar(f){google.script.run.withSuccessHandler(function(){toast();cerrar();setTimeout(function(){window.top.location.reload();},700);}).withFailureHandler(function(e){alert('Error: '+e.message);}).restaurarCandidato(PERFIL,f,quienSoy());}
function cerrar(){document.getElementById('modal').classList.remove('show');}
</script></body></html>`;
}
function notaManual_(perfilId, fila, etapa, nota, quien) {
  QUIEN_ = quien || '';
  const sh = getSheet(getConfig(perfilId));
  const col = {cv:5, screening:12, competencias:16}[etapa];
  sh.getRange(fila, col).setValue(nota);
  aplicarEstadoNota(sh, fila, etapa, nota);
  const est = String(sh.getRange(fila, 9).getValue() || '').trim();
  limpiarMotivoSiNoDescartado_(sh, fila, est);
  return { ok: true, estado: est };
}
function syncCitasCalendar() {
  const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
  const desde = new Date(); desde.setDate(desde.getDate() - 180);
  const hasta = new Date(); hasta.setDate(hasta.getDate() + 90);
  const eventos = cal.getEvents(desde, hasta);
  const norm = s => String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'');
  const perfiles = listarPerfiles().map(p => p.id).filter(id => id !== 'operaciones');

  const tipoDe = titulo => {
    const t = titulo.toLowerCase();
    if (t.includes('empresa')) return 'Empresa';
    if (t.includes('competencia')) return 'Competencias';
    if (t.includes('screening')) return 'Screening';
    return 'Entrevista';
  };
  const etapaDe = est => {
    const e = String(est).toLowerCase();
    if (e.includes('empresa') && e.includes('agendad')) return 'empresa';
    if (e.includes('competencia') && e.includes('agendad')) return 'competencias';
    if (e.includes('screening') && e.includes('agendad')) return 'screening';
    return '';
  };

  perfiles.forEach(id => {
    try {
      const sh = getSheet(getConfig(id));
      if (sh.getMaxColumns() < 38) sh.insertColumnsAfter(sh.getMaxColumns(), 38 - sh.getMaxColumns());
      sh.getRange(1, 37, sh.getMaxRows(), 2).clearDataValidations();
      sh.getRange(1,37).setValue('Cita');
      sh.getRange(1,38).setValue('Historial');
      const d = sh.getDataRange().getValues();
      for (let i = 1; i < d.length; i++) {
        if (!d[i][1]) continue;
        const nombreN = norm(d[i][1]);
        const mailN = String(d[i][21]||'').toLowerCase().trim();
        const suyos = eventos.filter(e => {
          if (mailN && mailN.includes('@')) {
            const invitados = e.getGuestList().map(g => g.getEmail().toLowerCase());
            if (invitados.includes(mailN)) return true;
          }
          const m = e.getTitle().match(/\(([^)]+)\)/);
          if(!m) return false;
          const evN = norm(m[1]);
          return evN === nombreN || (evN.length >= 10 && (nombreN.includes(evN) || evN.includes(nombreN)));
        });

        // historial completo (todas las entrevistas, más reciente arriba)
        if (suyos.length) {
          const hist = suyos.slice().sort((a,b) => b.getStartTime() - a.getStartTime())
            .map(e => tipoDe(e.getTitle()) + ' — ' + Utilities.formatDate(e.getStartTime(), 'GMT-3', 'dd/MM/yyyy HH:mm'))
            .join('\n');
          sh.getRange(i+1, 38).setValue(hist);
        } else {
          sh.getRange(i+1, 38).setValue('');
        }

        // cita pendiente (solo si estado agendado)
        const etapa = etapaDe(d[i][8]);
        if (!etapa) { sh.getRange(i+1,37).setValue(''); continue; }
        const ahora = new Date();
        const cands = suyos.filter(e => e.getTitle().toLowerCase().includes(etapa) && e.getStartTime() >= ahora);
        if (cands.length) {
          cands.sort((a,b) => a.getStartTime() - b.getStartTime());
          sh.getRange(i+1, 37).setValue(Utilities.formatDate(cands[0].getStartTime(), 'GMT-3', "dd/MM/yyyy 'a las' HH:mm"));
        } else {
          sh.getRange(i+1,37).setValue('');
        }
      }
      Logger.log(id + ' ok');
    } catch(e) { Logger.log('ERROR ' + id + ': ' + e.message); }
  });
  return true;
}

function crearTriggerCitas() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'syncCitasCalendar') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncCitasCalendar').timeBased().everyHours(1).create();
}
function notaEmpresa_(perfilId, fila, etapa, nota, quien) {
  QUIEN_ = quien || '';
  const sh = getSheet(getConfig(perfilId));
  const mapa = {
    'screening-empresa': { A:'Avanza a presencial empresa', B:'Reserva screening empresa', C:'Descartado en screening empresa' },
    'presencial-empresa': { A:'Contratado', B:'Reserva presencial empresa', C:'Descartado presencial empresa' }
  };
  const estado = mapa[etapa] && mapa[etapa][nota];
  if (!estado) return false;
  const celda = sh.getRange(fila, 9);
  const r = celda.getDataValidation();
  const estadoAntes = String(celda.getValue() || '').trim();
  celda.setDataValidation(null).setValue(estado);
  if (r) celda.setDataValidation(r);
  registrarHistorial_(sh.getName(), sh.getRange(fila, 2).getValue(), 'Nota ' + etapa, nota + ' · ' + (estadoAntes || '—') + ' → ' + estado);
  const colNota = (etapa === 'screening-empresa') ? 44 : 45;
  if (sh.getMaxColumns() < colNota) sh.insertColumnsAfter(sh.getMaxColumns(), colNota - sh.getMaxColumns());
  sh.getRange(fila, colNota).clearDataValidations().setValue(nota);
  limpiarMotivoSiNoDescartado_(sh, fila, estado);
  return { ok: true, estado: estado };
}
function enviarMailDescarte(perfilId, fila) {
  const cfg = getConfig(perfilId);
  const sh = getSheet(cfg);
  const nombre = String(sh.getRange(fila, 2).getValue() || '').split(' ')[0];
  const mail = String(sh.getRange(fila, 22).getValue() || '').trim();
  const estado = String(sh.getRange(fila, 9).getValue() || '').trim();
  const puesto = (typeof cfg.nombre === 'string') ? cfg.nombre : perfilId;
  const empresa = empresaParaCandidatos_(cfg);

  if (!mail || mail === 'Sin dato' || mail.indexOf('@') === -1) return { ok:false, error:'Sin mail válido' };

  const esSinScreening = /sin screening/i.test(estado);
  const asunto = 'Blumia Consulting · Proceso de selección';

  let cuerpo;
  if (esSinScreening) {
    cuerpo = 'Hola ' + nombre + ', ¿cómo estás?\n\n' +
      'Antes que nada, gracias por postularte a la búsqueda de ' + puesto + ' para ' + empresa + '.\n\n' +
      'En esta oportunidad avanzamos con otros perfiles que se ajustaban más a lo que el cliente estaba buscando, así que no vamos a seguir adelante con tu candidatura en este proceso puntual.\n\n' +
      'Queremos que sepas que tu CV queda en nuestra base, y que ante futuras búsquedas que encajen con tu perfil, vamos a tenerte en cuenta.\n\n' +
      'Te deseamos mucho éxito en lo que venga. ¡Gracias por tu interés en trabajar con nosotros!\n\n' +
      'Saludos,\nEquipo de Blumia Consulting';
  } else {
    cuerpo = 'Hola ' + nombre + ', ¿cómo estás?\n\n' +
      'Ante todo, gracias por el tiempo que compartimos durante la entrevista y por tu buena predisposición a lo largo del proceso.\n\n' +
      'Después de evaluar a todos los candidatos, en esta ocasión decidimos avanzar con otros perfiles que se ajustaban más a los requerimientos específicos de la búsqueda de ' + puesto + '. Así que no vamos a continuar con tu candidatura en este proceso.\n\n' +
      'Fue un gusto conocerte. Tu perfil queda en nuestra base y te vamos a tener presente para futuras búsquedas que puedan encajar mejor con tu experiencia.\n\n' +
      'Te deseamos mucho éxito en tus próximos pasos profesionales.\n\n' +
      'Saludos,\nEquipo de Blumia Consulting';
  }

  try {
    GmailApp.sendEmail(mail, asunto, cuerpo);
    sh.getRange(fila, 40).setValue('Enviado ' + Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy'));
    return { ok:true };
  } catch(e) {
    return { ok:false, error:String(e) };
  }
}
function enviarDescartesMasivo_(perfilId, textoSin, textoPost, quien) {
  QUIEN_ = quien || '';
  const cfg = getConfig(perfilId);
  const sh = getSheet(cfg);
  const d = sh.getDataRange().getValues();
  const puesto = (typeof cfg.nombre === 'string') ? cfg.nombre : perfilId;
  const empresa = empresaParaCandidatos_(cfg);
  const asunto = 'Blumia Consulting · Proceso de selección';
  let enviados = 0, sinMail = 0;

  for (let i = 1; i < d.length; i++) {
    if (!d[i][1]) continue;
    const estado = String(d[i][8] || '').trim();
    if (!/descartado sin screening|descartado post screening/i.test(estado)) continue;
    if (String(d[i][39] || '').trim()) continue; // ya enviado

    const mail = String(d[i][21] || '').trim();
    if (!mail || mail === 'Sin dato' || mail.indexOf('@') === -1) { sinMail++; continue; }

    const nombre = String(d[i][1]).split(' ')[0];
    const esSin = /sin screening/i.test(estado);
    let cuerpo;
    if (textoSin && textoPost) {
      cuerpo = (esSin ? textoSin : textoPost).replace(/\[NOMBRE\]/g, nombre);
    } else if (esSin) {
      cuerpo = 'Hola ' + nombre + ', ¿cómo estás?\n\n' +
        'Antes que nada, gracias por postularte a la búsqueda de ' + puesto + ' para ' + empresa + '.\n\n' +
        'En esta oportunidad avanzamos con otros perfiles que se ajustaban más a lo que el cliente estaba buscando, así que no vamos a seguir adelante con tu candidatura en este proceso puntual.\n\n' +
        'Queremos que sepas que tu CV queda en nuestra base, y que ante futuras búsquedas que encajen con tu perfil, vamos a tenerte en cuenta.\n\n' +
        'Te deseamos mucho éxito en lo que venga. ¡Gracias por tu interés en trabajar con nosotros!\n\n' +
        'Saludos,\nEquipo de Blumia Consulting';
    } else {
      cuerpo = 'Hola ' + nombre + ', ¿cómo estás?\n\n' +
        'Ante todo, gracias por el tiempo que compartimos durante la entrevista y por tu buena predisposición a lo largo del proceso.\n\n' +
        'Después de evaluar a todos los candidatos, en esta ocasión decidimos avanzar con otros perfiles que se ajustaban más a los requerimientos específicos de la búsqueda de ' + puesto + '. Así que no vamos a continuar con tu candidatura en este proceso.\n\n' +
        'Fue un gusto conocerte. Tu perfil queda en nuestra base y te vamos a tener presente para futuras búsquedas que puedan encajar mejor con tu experiencia.\n\n' +
        'Te deseamos mucho éxito en tus próximos pasos profesionales.\n\n' +
        'Saludos,\nEquipo de Blumia Consulting';
    }

    try {
      GmailApp.sendEmail(mail, asunto, cuerpo, {from:'empleos@blumiaconsulting.com', name:'Blumia Consulting'});
      sh.getRange(i + 1, 40).setValue('Enviado ' + Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy'));
      registrarHistorial_(sh.getName(), d[i][1], 'Mail de descarte', 'Enviado (' + (esSin ? 'sin screening' : 'post screening') + ')');
      enviados++;
    } catch(e) {}
  }
  return { ok:true, enviados: enviados, sinMail: sinMail };
}

function previaDescartes_(perfilId) {
  const cfg = getConfig(perfilId);
  const sh = getSheet(cfg);
  const d = sh.getDataRange().getValues();
  const puesto = (typeof cfg.nombre === 'string') ? cfg.nombre : perfilId;
  let sinScr = 0, postScr = 0;
  for (let i = 1; i < d.length; i++) {
    if (!d[i][1]) continue;
    const estado = String(d[i][8] || '').trim();
    if (String(d[i][39] || '').trim()) continue;
    const mail = String(d[i][21] || '').trim();
    if (!mail || mail === 'Sin dato' || mail.indexOf('@') === -1) continue;
    if (/descartado sin screening/i.test(estado)) sinScr++;
    else if (/descartado post screening/i.test(estado)) postScr++;
  }
  return {
    sinScr: sinScr, postScr: postScr, puesto: puesto,
    textoSin: 'Hola [NOMBRE], ¿cómo estás?\n\nAntes que nada, gracias por postularte a la búsqueda de ' + puesto + ' para ' + empresaParaCandidatos_(cfg) + '.\n\nEn esta oportunidad avanzamos con otros perfiles que se ajustaban más a lo que el cliente estaba buscando, así que no vamos a seguir adelante con tu candidatura en este proceso puntual.\n\nQueremos que sepas que tu CV queda en nuestra base, y que ante futuras búsquedas que encajen con tu perfil, vamos a tenerte en cuenta.\n\nTe deseamos mucho éxito en lo que venga. ¡Gracias por tu interés en trabajar con nosotros!\n\nSaludos,\nEquipo de Blumia Consulting',
    textoPost: 'Hola [NOMBRE], ¿cómo estás?\n\nAnte todo, gracias por el tiempo que compartimos durante la entrevista y por tu buena predisposición a lo largo del proceso.\n\nDespués de evaluar a todos los candidatos, en esta ocasión decidimos avanzar con otros perfiles que se ajustaban más a los requerimientos específicos de la búsqueda de ' + puesto + '. Así que no vamos a continuar con tu candidatura en este proceso.\n\nFue un gusto conocerte. Tu perfil queda en nuestra base y te vamos a tener presente para futuras búsquedas que puedan encajar mejor con tu experiencia.\n\nTe deseamos mucho éxito en tus próximos pasos profesionales.\n\nSaludos,\nEquipo de Blumia Consulting'
  };
}
function empresasYPerfiles(){
  const rows = configValores_();
  const emp = {};
  const cEst = rows[0].indexOf('estado_busqueda');
  for(let i=1;i<rows.length;i++){
    if(!rows[i][0]) continue;
    const e = rows[i][9] || 'Sin empresa';
    const cerr = cEst >= 0 && String(rows[i][cEst]||'').toUpperCase().trim() === 'CERRADA';
    (emp[e]=emp[e]||[]).push([rows[i][0], rows[i][1] + (cerr ? ' · cerrada' : '')]);
  }
  return emp;
}
function crearPestanaPlantillas(){
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let sh = ss.getSheetByName('Plantillas');
  if(sh){ Logger.log('Ya existe la pestaña Plantillas'); return; }
  sh = ss.insertSheet('Plantillas');
  sh.getRange(1,1,1,4).setValues([['clave','canal','link_calendar','texto']]);
  sh.getRange(1,1,1,4).setFontWeight('bold');
  const filas = [
    ['contacto-screening','WhatsApp','',''],
    ['contacto-competencias','WhatsApp','',''],
    ['contacto-empresa','WhatsApp','',''],
    ['descarte-sin','Mail','',''],
    ['descarte-post','Mail','','']
  ];
  sh.getRange(2,1,filas.length,4).setValues(filas);
  sh.setColumnWidth(4, 600);
  Logger.log('Pestaña Plantillas creada con 5 filas');
}
function armarComunicacion_(perfilId, fila, tipo, quien){
  const cfg = getConfig(perfilId);
  const sh = getSheet(cfg);
  const nombre = String(sh.getRange(fila, 2).getValue() || '').split(' ')[0];
  const tel = String(sh.getRange(fila, 3).getValue() || '').replace(/\D/g,'');
  const entrevistador = String(quien || '').trim() || String(sh.getRange(fila, 14).getValue() || 'Equipo Blumia');
  const puesto = (typeof cfg.nombre_perfil === 'string') ? cfg.nombre_perfil : perfilId;
  const empresa = empresaParaCandidatos_(cfg);

  const linkPorTipo = {
    'contacto-screening': cfg.link_screening,
    'contacto-competencias': cfg.link_competencias,
    'contacto-empresa-virtual': cfg.link_empresa_virtual,
    'contacto-empresa-presencial': cfg.link_empresa_virtual
  };
  const link = linkPorTipo[tipo] || '';

  let texto = plantillaTexto_(tipo, perfilId).texto;
  if(!texto) return {ok:false, error:'No encontré la plantilla ' + tipo};

  texto = firmaConPlaceholder_(texto).replace(/\[NOMBRE\]/g, nombre)
               .replace(/\[PUESTO\]/g, puesto)
               .replace(/la empresa \[EMPRESA\]/g, EMPRESAS_CONOCIDAS_[empresa] ? 'la empresa ' + empresa : empresa)
               .replace(/\[EMPRESA\]/g, empresa)
               .replace(/\[ENTREVISTADOR\]/g, entrevistador)
               .replace(/\[LINK\]/g, link);

  let telWa = tel;
  if(telWa.indexOf('54') !== 0) telWa = '549' + telWa.replace(/^0+/,'');
  var etiqueta = {'contacto-screening':'screening','contacto-competencias':'competencias','contacto-empresa-virtual':'empresa','contacto-empresa-presencial':'empresa','descarte-competencias':'descarte comp.','descarte-empresa':'descarte emp.'}[tipo] || '';
  if(etiqueta){
    sh.getRange(fila, 41).setValue(etiqueta + ' · ' + Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy'));
    registrarContacto_(perfilId, fila, sh.getRange(fila, 2).getValue(), tel, 'WhatsApp: ' + etiqueta, entrevistador);
    registrarHistorial_(sh.getName(), sh.getRange(fila, 2).getValue(), 'Contacto', 'WhatsApp: ' + etiqueta, entrevistador);
  }
  const url = 'https://wa.me/' + telWa + '?text=' + encodeURIComponent(texto);
  return {ok:true, url:url, texto:texto};
}
function crearPestanaBanco(){
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let sh = ss.getSheetByName('Banco');
  if(sh){ Logger.log('Ya existe la pestaña Banco'); return; }
  sh = ss.insertSheet('Banco');
  const H = ['fecha_alta','nombre','categoria','seniority','zona','formacion','pretension','tel','mail','cv','resumen','origen_perfil','origen_empresa','origen_fila','nota_banco'];
  sh.getRange(1,1,1,H.length).setValues([H]);
  sh.getRange(1,1,1,H.length).setFontWeight('bold').setBackground('#1A6645').setFontColor('#fff');
  sh.setFrozenRows(1);
  sh.setColumnWidth(11, 400);
  sh.setColumnWidth(15, 300);
  Logger.log('Pestaña Banco creada con ' + H.length + ' columnas');
}
function enviarABanco_(perfilId, fila, categoria, seniority, notaBanco, quien){
  QUIEN_ = quien || '';
  const cfg = getConfig(perfilId);
  const sh = getSheet(cfg);
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const banco = ss.getSheetByName('Banco');
  if(!banco) return {ok:false, error:'Falta la pestaña Banco'};

  const d = sh.getRange(fila, 1, 1, 41).getValues()[0];
  const nombre = String(d[1] || '');
  if(!nombre) return {ok:false, error:'Sin nombre'};

  const yaEsta = banco.getDataRange().getValues();
  for(let i=1;i<yaEsta.length;i++){
    if(String(yaEsta[i][1]).toLowerCase().trim() === nombre.toLowerCase().trim()){
      return {ok:false, error:'Ya está en el banco (agregado el ' + yaEsta[i][0] + ')'};
    }
  }

  function linkDe(col){
    try{
      const rt = sh.getRange(fila, col).getRichTextValue();
      if(rt && rt.getLinkUrl()) return rt.getLinkUrl();
      const f = sh.getRange(fila, col).getFormula();
      const m = f.match(/HYPERLINK\("([^"]+)"/);
      if(m) return m[1];
    }catch(e){}
    return '';
  }

  let cvId = '';
  try { cvId = extraerFileId(sh, fila, 8) || ''; } catch(e){}
  const pret = (String(d[5]||'').match(/Remuneraci[oó]n pretendida:\s*\$?\s*([\d.\s]+)/i)||[])[1] || '';

  banco.appendRow([
    Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy'),
    nombre,
    categoria,
    seniority,
    d[23] || '',
    d[22] || '',
    pret.replace(/[.\s]/g,''),
    d[2] || '',
    d[21] || '',
    '',
    d[6] || '',
    (typeof cfg.nombre_perfil === 'string') ? cfg.nombre_perfil : perfilId,
    (typeof cfg.empresa === 'string') ? cfg.empresa : '',
    fila,
    notaBanco || '',
    '', '', '', '',
    d[4] || '',
    d[11] || '',
    d[15] || '',
    d[8] || ''
  ]);

  const r = banco.getLastRow();
  const H = banco.getDataRange().getValues()[0];
  const cTel = H.indexOf('tel');
  if(cTel > -1){ banco.getRange(r, cTel+1).setNumberFormat('@').setValue(String(d[2]||'').replace(/^=/,'').replace(/^\+/,'+')); }
  const put = function(nombreCol, url, texto){
    const c = H.indexOf(nombreCol);
    if(c > -1 && url) banco.getRange(r, c+1).setFormula('=HYPERLINK("' + url + '"; "' + texto + '")');
  };
  if(cvId) put('cv', 'https://drive.google.com/file/d/' + cvId + '/view', 'Ver CV');
  put('screening_pdf', linkDe(13), 'Ver screening');
  put('competencias_pdf', linkDe(17), 'Ver competencias');
  put('informe', linkDe(33), 'Ver informe');
  put('postulacion', linkDe(31), 'Ver postulación');
  registrarHistorial_(sh.getName(), nombre, 'Banco de talento', 'Pasó al banco' + (categoria ? ' · ' + categoria : '') + (seniority ? ' · ' + seniority : ''));

  return {ok:true};
}
function doGetBanco(){
  const banco = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Banco');
  if(!banco) return '[]';
  const d = banco.getDataRange().getValues();
  const H = d[0];
  const idx = {};
  H.forEach(function(h,i){ if(h) idx[h] = i; });
  function link(fila, nombreCol){
    const c = idx[nombreCol];
    if(c === undefined) return '';
    try{
      const rt = banco.getRange(fila+1, c+1).getRichTextValue();
      if(rt && rt.getLinkUrl()) return rt.getLinkUrl();
      const f = banco.getRange(fila+1, c+1).getFormula();
      const m = f.match(/HYPERLINK\("([^"]+)"/);
      if(m) return m[1];
    }catch(e){}
    return '';
  }
  const out = [];
  for(let i=1;i<d.length;i++){
    if(!d[i][idx['nombre']]) continue;
    out.push({
      fila: i+1,
      alta: d[i][idx['fecha_alta']] || '',
      nombre: d[i][idx['nombre']] || '',
      categoria: d[i][idx['categoria']] || '',
      seniority: d[i][idx['seniority']] || '',
      zona: d[i][idx['zona']] || '',
      formacion: d[i][idx['formacion']] || '',
      pretension: d[i][idx['pretension']] || '',
      pretNum: Number(String(d[i][idx['pretension']]||'').replace(/\D/g,'')) || 0,
      tel: d[i][idx['tel']] || '',
      mail: d[i][idx['mail']] || '',
      resumen: d[i][idx['resumen']] || '',
      origen: (d[i][idx['origen_perfil']]||'') + (d[i][idx['origen_empresa']] ? ' · ' + d[i][idx['origen_empresa']] : ''),
      notaBanco: d[i][idx['nota_banco']] || '',
      notaCv: d[i][idx['nota_cv']] || '',
      notaScr: d[i][idx['nota_scr']] || '',
      notaComp: d[i][idx['nota_comp']] || '',
      estadoOrigen: d[i][idx['estado_origen']] || '',
      cv: link(i,'cv'),
      scrPDF: link(i,'screening_pdf'),
      compPDF: link(i,'competencias_pdf'),
      informe: link(i,'informe'),
      postulacion: link(i,'postulacion')
    });
  }
  return JSON.stringify(out);
}
function BANCO_TPL(dataJSON, empresasJSON, desde){
  const LOGO='data:image/png;base64,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';
  return `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Banco de Talento · Blumia</title>
<link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Poppins:wght@400;500;600;700;800&display=swap">
<style>
:root{--green:#243D2C;--glight:#E3EEDC;--amber:#F9E7A0;--violet:#243D2C;--bg:#F3F8F1;--line:#D5E0CE;--mid:#555;--font:'Poppins','Plus Jakarta Sans',sans-serif}
*{box-sizing:border-box;margin:0;padding:0}body{font-family:var(--font);background:var(--bg);color:#1A1A1A;font-size:14px}
header{background:var(--violet);color:#fff;padding:16px 20px;display:flex;align-items:center;gap:14px;position:sticky;top:0;z-index:10;flex-wrap:wrap}
header img{height:34px}header h1{font-size:18px;font-weight:800}
.nav{display:flex;gap:8px;margin-left:auto;flex-wrap:wrap}
.nav button{padding:8px 16px;border-radius:8px;font-size:12px;font-weight:700;color:#243D2C;background:#F9E7A0;border:none;cursor:pointer;font-family:var(--font)}
.filtros{padding:14px 16px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;background:#fff;border-bottom:1px solid var(--line)}
.fwrap{position:relative;display:inline-block}
.fwrap svg{position:absolute;left:9px;top:50%;transform:translateY(-50%);opacity:.4;pointer-events:none}
.filtros input,.filtros select{padding:10px;border:1px solid var(--line);border-radius:8px;font-family:var(--font);font-size:13px;background:#fff}
.fwrap input{padding-left:30px}
.filtros button{padding:10px 14px;border:1px solid var(--line);border-radius:8px;background:#fff;cursor:pointer;font-family:var(--font);font-size:13px}
.wrap{padding:16px;display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:12px}
.p{background:#fff;border:1px solid var(--line);border-left:5px solid #5BB5A5;border-radius:16px;padding:14px;cursor:pointer}
.p:hover{box-shadow:0 2px 10px rgba(0,0,0,.08)}
.p .n{font-weight:800;font-size:15px;margin-bottom:4px}
.p .cat{display:inline-block;font-size:10px;font-weight:800;text-transform:uppercase;letter-spacing:.04em;background:#FEF3D8;color:#7a5500;padding:3px 8px;border-radius:10px;margin-right:5px;margin-bottom:6px}
.p .sen{display:inline-block;font-size:10px;font-weight:800;background:var(--glight);color:var(--green);padding:3px 8px;border-radius:10px;margin-bottom:6px}
.p .meta{font-size:12px;color:var(--mid);line-height:1.6}
.p .org{font-size:11px;color:#888;margin-top:6px;border-top:1px solid var(--line);padding-top:6px}
.nota{display:inline-block;width:20px;height:20px;line-height:20px;text-align:center;border-radius:4px;font-weight:800;font-size:11px;color:#fff;float:right}
.nA{background:#1A6645}.nB{background:#F5C842;color:#1A1A1A}.nC{background:#E06666}
.vacio{padding:40px;text-align:center;color:#888;grid-column:1/-1}
.modal-bg{display:none;position:fixed;inset:0;background:rgba(0,0,0,.5);z-index:200;padding:20px;overflow-y:auto}
.modal-bg.show{display:flex;align-items:flex-start;justify-content:center}
.modal{background:#fff;border-radius:10px;max-width:560px;width:100%;margin:auto}
.modal-h{background:var(--violet);color:#1A1A1A;padding:20px 24px;border-radius:10px 10px 0 0;display:flex;justify-content:space-between;align-items:center}
.modal-h h2{font-size:22px;font-weight:800}
.mx{background:none;border:none;color:#1A1A1A;font-size:20px;cursor:pointer}
.modal-b{padding:24px}
.frow{display:flex;justify-content:space-between;padding:9px 0;border-bottom:1px solid var(--line);font-size:13px;gap:12px}
.fl{font-weight:800;color:var(--mid);text-transform:uppercase;font-size:11px;letter-spacing:.05em;flex-shrink:0}
.fv{text-align:right}
.fres{margin:16px 0;font-size:13px;line-height:1.6;color:#333}
.mbtns{display:flex;flex-wrap:wrap;gap:8px;margin-top:14px}
.mbtn{display:inline-block;background:var(--green);color:#fff;padding:10px 18px;border-radius:6px;text-decoration:none;font-weight:800;font-size:13px;border:none;cursor:pointer;font-family:var(--font)}
.mbtn.alt{background:#fff;color:var(--green);border:2px solid var(--green)}
.mbtn.wa{background:#25D366;color:#fff}

/* celular */
@media(max-width:640px){header{position:static;padding:10px 12px}header h1{font-size:15px;flex:1 1 100%}.filtros{display:grid;grid-template-columns:1fr 1fr;gap:6px;padding:8px 10px}.filtros .fwrap{width:100%}.filtros input,.filtros select{width:100%!important;box-sizing:border-box}.wrap{padding:10px;grid-template-columns:1fr}}
</style></head><body>
<header><img src="${LOGO}"><h1>★ Banco de Talento</h1><div class="nav"><button onclick="window.top.location.href='${URL_WEBAPP}?tipo=kanban&perfil=${desde}'">← Volver al proceso</button></div></header>
<div class="filtros">
<span class="fwrap"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#1A1A1A" stroke-width="2.5"><circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/></svg><input id="fNombre" type="text" placeholder="Nombre" oninput="filtrar()" style="width:150px"></span>
<select id="fCat" onchange="filtrar()"><option value="">Categoría: todas</option></select>
<select id="fSen" onchange="filtrar()"><option value="">Seniority: todos</option></select>
<span class="fwrap"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#1A1A1A" stroke-width="2.5"><circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/></svg><input id="fZona" type="text" placeholder="Zona" oninput="filtrar()" style="width:150px"></span>
<span class="fwrap"><svg width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="#1A1A1A" stroke-width="2.5"><circle cx="11" cy="11" r="7"/><path d="M20 20l-4-4"/></svg><input id="fExp" type="text" placeholder="Formación / experiencia" oninput="filtrar()" style="width:200px"></span>
<input id="fPret" type="number" placeholder="Pretensión hasta $" oninput="filtrar()" style="width:160px">
<button onclick="limpiar()">Limpiar</button>
<span id="cont" style="font-size:12px;color:#555;font-weight:700"></span>
</div>
<div class="wrap" id="wrap"></div>
<div class="modal-bg" id="modal" onclick="if(event.target===this)cerrar()"><div class="modal"><div class="modal-h"><h2 id="mtit"></h2><button class="mx" onclick="cerrar()">✕</button></div><div class="modal-b" id="mbody"></div></div></div>
<script>
const DATA=${dataJSON};
const EMPRESAS=${empresasJSON};
const nota=n=>n?'<span class="nota n'+String(n).replace('+','')+'">'+n+'</span>':'';
function pintar(){
  var w=document.getElementById('wrap');
  if(!DATA.length){w.innerHTML='<div class="vacio">Todavía no hay nadie en el banco de talento.<br>Enviá candidatos desde la ficha del Kanban con el botón ★ Banco de talento.</div>';return;}
  w.innerHTML=DATA.map(function(c){
    return '<div class="p" data-fila="'+c.fila+'" onclick="ficha('+c.fila+')">'+nota(c.notaComp||c.notaScr||c.notaCv)+
    '<div class="n">'+c.nombre+'</div>'+
    '<span class="cat">'+c.categoria+'</span><span class="sen">'+c.seniority+'</span>'+
    '<div class="meta">'+(c.zona||'—')+(c.pretension?' · $'+c.pretension:'')+(c.formacion?'<br>'+c.formacion:'')+'</div>'+
    '<div class="org">De: '+c.origen+' · '+c.alta+'</div></div>';
  }).join('');
}
pintar();
(function(){
  var cats=['RRHH / Gestión de Personas','Datos / BI','Compras / Abastecimiento','Calidad / SGI','Administración / Finanzas','Operaciones / Producción','Comercial / Ventas','Otros'];
  var sens=['Jr','Ssr','Sr','Coordinador','Supervisor','Jefatura / Gerencia'];
  DATA.forEach(function(c){
    if(c.categoria&&cats.indexOf(c.categoria)===-1)cats.push(c.categoria);
    if(c.seniority&&sens.indexOf(c.seniority)===-1)sens.push(c.seniority);
  });
  var sc=document.getElementById('fCat'),ss=document.getElementById('fSen');
  cats.forEach(function(x){var n=DATA.filter(function(c){return c.categoria===x;}).length;var o=document.createElement('option');o.value=x;o.text=x+' ('+n+')';sc.appendChild(o);});
  sens.forEach(function(x){var n=DATA.filter(function(c){return c.seniority===x;}).length;var o=document.createElement('option');o.value=x;o.text=x+' ('+n+')';ss.appendChild(o);});
})();
function nrm(s){return String(s||'').toLowerCase().normalize('NFD').replace(/[\\u0300-\\u036f]/g,'');}
function raiz(p){p=nrm(p);if(p.length>5)return p.slice(0,-2);if(p.length>3)return p.slice(0,-1);return p;}
function coincide(t,q){if(!q.trim())return true;var x=nrm(t);var ps=q.trim().split(/\\s+/);for(var i=0;i<ps.length;i++){if(x.indexOf(raiz(ps[i]))===-1)return false;}return true;}
function filtrar(){
  var fN=document.getElementById('fNombre').value;
  var fC=document.getElementById('fCat').value;
  var fS=document.getElementById('fSen').value;
  var fZ=document.getElementById('fZona').value;
  var fE=document.getElementById('fExp').value;
  var fP=Number(document.getElementById('fPret').value)||0;
  var hay=fN||fC||fS||fZ||fE||fP;
  var vis=0;
  document.querySelectorAll('.p').forEach(function(el){
    var c=DATA.find(function(x){return x.fila===+el.dataset.fila;});
    var ok=true;
    if(c){
      if(ok&&fN&&!coincide(c.nombre,fN))ok=false;
      if(ok&&fC&&c.categoria!==fC)ok=false;
      if(ok&&fS&&c.seniority!==fS)ok=false;
      if(ok&&fZ&&!coincide(c.zona,fZ))ok=false;
      if(ok&&fE&&!coincide(c.formacion+' '+c.resumen+' '+c.notaBanco,fE))ok=false;
      if(ok&&fP&&(!c.pretNum||c.pretNum>fP))ok=false;
    }
    el.style.display=ok?'':'none';
    if(ok)vis++;
  });
  document.getElementById('cont').textContent=hay?vis+' resultado'+(vis===1?'':'s'):DATA.length+' en el banco';
}
function limpiar(){['fNombre','fZona','fExp','fPret'].forEach(function(id){document.getElementById(id).value='';});document.getElementById('fCat').value='';document.getElementById('fSen').value='';filtrar();}
document.getElementById('cont').textContent=DATA.length+' en el banco';
function ficha(f){
  var c=DATA.find(function(x){return x.fila===f;});
  var row=function(l,v){return v?'<div class="frow"><span class="fl">'+l+'</span><span class="fv">'+v+'</span></div>':'';};
  var btn=function(u,t){return u?'<a href="'+u+'" target="_blank" class="mbtn">'+t+'</a>':'';};
  var tel=String(c.tel||'').replace(/\\D/g,'');
  var wa=tel?'<a href="https://wa.me/'+(tel.indexOf('54')===0?tel:'549'+tel.replace(/^0+/,''))+'" target="_blank" class="mbtn wa">💬 WhatsApp</a>':'';
  var mail=(c.mail&&c.mail.indexOf('@')>-1)?'<a href="https://mail.google.com/mail/?view=cm&to='+c.mail+'" target="_blank" class="mbtn" style="background:#EA4335">✉ Mail</a>':'';
  document.getElementById('mbody').innerHTML=
    row('Categoría',c.categoria)+row('Seniority',c.seniority)+row('En el banco desde',c.alta)+
    row('Zona',c.zona)+row('Formación',c.formacion)+row('Pretensión',c.pretension?'$'+c.pretension:'')+
    row('Teléfono',c.tel)+row('Email',c.mail)+
    row('Viene de',c.origen)+row('Estado en esa búsqueda',c.estadoOrigen)+
    row('Nota CV',c.notaCv)+row('Nota screening',c.notaScr)+row('Nota competencias',c.notaComp)+
    (c.notaBanco?'<div class="fres" style="background:#EDE9FE;padding:12px;border-radius:8px"><b>Nota del banco:</b><br>'+c.notaBanco+'</div>':'')+
    '<div class="fres">'+(c.resumen||'Sin resumen')+'</div>'+
    '<div class="mbtns">'+wa+mail+btn(c.cv,'Ver CV')+btn(c.scrPDF,'Ver screening')+btn(c.compPDF,'Ver competencias')+btn(c.informe,'Ver informe')+btn(c.postulacion,'Ver postulación')+'<button class="mbtn" style="background:#1A6645" onclick="menuBusqueda('+c.fila+')">→ Enviar a búsqueda</button>'+'<button class="mbtn" style="background:#E06666" onclick="borrarBanco('+c.fila+',this)">🗑 Quitar del banco</button></div>';
  document.getElementById('mtit').textContent=c.nombre;
  document.getElementById('modal').classList.add('show');
}
function volver(){if(history.length>1){history.back();}else{window.top.location.href='${URL_WEBAPP}?tipo=inicio';}}
function borrarBanco(f,btn){
  var c=DATA.find(function(x){return x.fila===f;});
  blConfirm('¿Quitar a '+c.nombre+' del banco de talento?\\nSigue estando en su búsqueda de origen.',function(){
  btn.textContent='Quitando...';btn.disabled=true;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.ok){btn.textContent='✓ Quitado';btn.style.background='#999';setTimeout(function(){window.top.location.href='${URL_WEBAPP}?tipo=banco';},600);}
    else{alert(r?r.error:'Error');btn.disabled=false;btn.textContent='🗑 Quitar del banco';}
  }).withFailureHandler(function(e){alert('Error: '+e.message);btn.disabled=false;btn.textContent='🗑 Quitar del banco';}).borrarDelBanco(f,c.nombre);
  },{si:'Quitar',peligro:true});
}
function menuBusqueda(f){
  try{
  var c=DATA.find(function(x){return x.fila===f;});
  var h='<div style="font-size:13px;margin-bottom:14px;color:#555">Enviar a <b>'+c.nombre+'</b> a otra búsqueda.<br><span style="color:#888">Entra como <b>Avanza a screening</b>, con fuente Banco de talento.</span></div>';
  h+='<div class="fl" style="margin-bottom:6px">Búsqueda de destino</div><select id="bDest" style="width:100%;padding:10px;border:1px solid #DDD;border-radius:8px;font-family:inherit;font-size:13px;margin-bottom:14px">';
  Object.keys(EMPRESAS).forEach(function(emp){
    h+='<optgroup label="'+emp+'">';
    EMPRESAS[emp].forEach(function(p){h+='<option value="'+p[0]+'">'+p[1]+'</option>';});
    h+='</optgroup>';
  });
  h+='</select>';
  h+='<div class="mbtns"><button class="mbtn" onclick="mandarABusqueda('+f+',this)">→ Enviar</button><button class="mbtn alt" onclick="ficha('+f+')">← Volver</button></div>';
  document.getElementById('mtit').textContent='Enviar a búsqueda';
  document.getElementById('mbody').innerHTML=h;
  }catch(e){alert('ERROR: '+e.message);}
}
function mandarABusqueda(f,btn){
  var dest=document.getElementById('bDest').value;
  btn.textContent='Enviando...';btn.disabled=true;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.ok){btn.textContent='✓ Enviado a '+r.puesto;btn.style.background='#999';}
    else{alert(r?r.error:'Error');btn.textContent='→ Enviar';btn.disabled=false;}
  }).withFailureHandler(function(e){alert('Error: '+e.message);btn.textContent='→ Enviar';btn.disabled=false;}).enviarABusqueda(f,dest);
}
function volver(){if(history.length>1){history.back();}else{window.top.location.href='${URL_WEBAPP}?tipo=inicio';}}
function cerrar(){document.getElementById('modal').classList.remove('show');}
</script></body></html>`;
}

function bancoHTML(desde){
  return HtmlService.createHtmlOutput(BANCO_TPL(doGetBanco(), JSON.stringify(empresasYPerfiles()), desde || ((listarPerfiles()[0] || {}).id || '')))
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
function borrarDelBanco_(fila, nombre){
  const b = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Banco');
  const v = String(b.getRange(fila, 2).getValue() || '').trim();
  if(v.toLowerCase() !== String(nombre||'').toLowerCase().trim()){
    return {ok:false, error:'La fila no coincide con el nombre. Actualizá la página y probá de nuevo.'};
  }
  b.deleteRow(fila);
  return {ok:true};
}
function enviarABusqueda_(filaBanco, perfilDestino){
  const banco = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Banco');
  const d = banco.getDataRange().getValues();
  const H = d[0];
  const idx = {}; H.forEach(function(h,i){ if(h) idx[h] = i; });
  const b = d[filaBanco-1];
  if(!b || !b[idx['nombre']]) return {ok:false, error:'No encontré a la persona en el banco'};
  const nombre = String(b[idx['nombre']]).trim();

  const cfg = getConfig(perfilDestino);
  const sh = getSheet(cfg);
  const dest = sh.getDataRange().getValues();
  for(let i=1;i<dest.length;i++){
    if(String(dest[i][1]).toLowerCase().trim() === nombre.toLowerCase().trim()){
      return {ok:false, error:'Ya está en esa búsqueda (fila ' + (i+1) + ')'};
    }
  }

  let cvUrl = '';
  try{
    const c = idx['cv'];
    const rt = banco.getRange(filaBanco, c+1).getRichTextValue();
    if(rt && rt.getLinkUrl()) cvUrl = rt.getLinkUrl();
    if(!cvUrl){
      const m = banco.getRange(filaBanco, c+1).getFormula().match(/HYPERLINK\("([^"]+)"/);
      if(m) cvUrl = m[1];
    }
  }catch(e){}

  const filaRef = sh.getLastRow();
  const nueva = filaRef + 1;

  sh.getRange(nueva, 1).setValue(Utilities.formatDate(new Date(),'GMT-3','dd/MM/yyyy'));
  sh.getRange(nueva, 2).setValue(nombre);
  sh.getRange(nueva, 3).setValue(b[idx['tel']] || '');
  sh.getRange(nueva, 5).setValue(b[idx['nota_cv']] || '');
  sh.getRange(nueva, 6).setValue('Zona: ' + (b[idx['zona']]||'') + (b[idx['pretension']] ? '\nRemuneración pretendida: ' + b[idx['pretension']] : ''));
  sh.getRange(nueva, 7).setValue(b[idx['resumen']] || '');
  if(cvUrl) sh.getRange(nueva, 8).setFormula('=HYPERLINK("' + cvUrl + '"; "Ver CV")');
  sh.getRange(nueva, 9).setDataValidation(null).setValue('Avanza a screening');
  sh.getRange(nueva, 10).setValue('Banco de talento');
  sh.getRange(nueva, 22).setValue(b[idx['mail']] || '');
  sh.getRange(nueva, 23).setValue(b[idx['formacion']] || '');
  sh.getRange(nueva, 24).setValue(b[idx['zona']] || '');

  [11, 21, 31].forEach(function(col){
    try{
      const f = sh.getRange(filaRef, col).getFormula();
      if(f) sh.getRange(filaRef, col).copyTo(sh.getRange(nueva, col));
    }catch(e){}
  });

  return {ok:true, puesto: (typeof cfg.nombre_perfil === 'string') ? cfg.nombre_perfil : perfilDestino, fila: nueva};
}
function borrarMarcaContacto_(perfilId, fila, quien){
  QUIEN_ = quien || '';
  const sh = getSheet(getConfig(perfilId));
  registrarHistorial_(sh.getName(), sh.getRange(fila, 2).getValue(), 'Contacto', 'Marca de contacto borrada (' + sh.getRange(fila, 41).getValue() + ')');
  sh.getRange(fila, 41).clearContent();
  return {ok:true};
}