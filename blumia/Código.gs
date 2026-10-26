const OPENAI_API_KEY = PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
const SHEET_MASTER_ID = '151V8cc-w00tP9C6mSxVaKNkRxgKXgUDI4L0KWXeNJGI';
const CORREO_DESTINO = 'empleos@blumiaconsulting.com';
const URL_DEPLOY = 'https://script.google.com/macros/s/AKfycbytekPKN31bZLxczDVknMr433K9KflzmrR-dCWsC1hLlx2MO7tbw6SrMjFXa1PnucQo/exec';
function URL_BASE() { return ScriptApp.getService().getUrl(); }

function getConfig(perfilId) {
  const config = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config');
  if (!config) throw new Error('Falta pestaña Config. Corré setupInicial()');
  const rows = config.getDataRange().getValues();
  const headers = rows[0];
  const fila = rows.find(r => r[0] === perfilId);
  if (!fila) throw new Error('Perfil no encontrado: ' + perfilId);
  const cfg = {};
  headers.forEach((h, i) => cfg[h] = fila[i]);
  return cfg;
}

function listarPerfiles() {
  const config = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getDataRange().getValues();
  return config.slice(1).filter(r => r[0]).map(r => ({ id: r[0], nombre: r[1] }));
}

function getSheet(cfg) {
  return SpreadsheetApp.openById(cfg.sheet_id).getSheetByName(cfg.hoja_nombre);
}

function getDatosPerfil(perfilId) {
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
      headers.forEach((h, i) => o[h] = r[i]);
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
    entrevistadoras: String(hero.entrevistadoras || 'Ariel').split(',').map(s => s.trim()).filter(Boolean)
  };
}

function doGet(e) {
  const tipo = e?.parameter?.tipo;
  const perfil = e?.parameter?.perfil || 'operaciones';
  let template, titulo;

  if (tipo === 'competencias') {
    template = HtmlService.createTemplateFromFile('Index');
    template.candidato = e.parameter.nombre || '';
    template.entrevistador = e.parameter.entrevistador || '';
    template.perfil = perfil;
    titulo = 'Evaluación de Competencias | Blumia';
  } else if (tipo === 'postulacion') {
    template = HtmlService.createTemplateFromFile('Postulacion');
    const hero = getHeroPostulacion(perfil);
    template.perfil = perfil;
    template.pos = hero.pos;
    template.loc = hero.loc;
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
    const url = generarInformeCandidato(perfil, Number(e.parameter.fila));
    return HtmlService.createHtmlOutput(
      '<html><body style="font-family:sans-serif;padding:40px;text-align:center">' +
      '<h2 style="color:#1A6645">Informe generado ✓</h2>' +
      '<p>El informe de este candidato está listo.</p>' +
      '<a href="' + url + '" target="_blank" style="display:inline-block;background:#1A6645;color:#fff;padding:12px 24px;border-radius:4px;text-decoration:none;font-weight:bold">Abrir informe</a>' +
      '<p style="color:#888;margin-top:20px">También quedó el link "Ver informe" en la planilla.</p>' +
      '</body></html>'
    ).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else if (tipo === 'perfilpuesto') {
    return perfilPuestoHTML(perfil, e.parameter.t);
  } else if (tipo === 'kanban') {
    return kanbanHTML(perfil);
  } else if (tipo === 'banco') {
    return bancoHTML(e.parameter.desde || '');
  } else if (tipo === 'procesohtml') {
    let empProceso = 'MACS';
    try { if (e.parameter.perfil) empProceso = empresaKey_(getConfig(e.parameter.perfil)); } catch (err) { }
    generarInformeProceso(empProceso);
    completarProcesoIA(empProceso);
    const url = generarHTMLProceso(empProceso);
    return HtmlService.createHtmlOutput('<!DOCTYPE html><html><body style="font-family:sans-serif;padding:40px;text-align:center"><h2 style="color:#1A6645">Informe de proceso listo</h2><a href="' + url + '" target="_blank" style="display:inline-block;background:#1A6645;color:#fff;padding:14px 28px;border-radius:6px;text-decoration:none;font-weight:800">Abrir informe</a></body></html>').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } else {
    template = HtmlService.createTemplateFromFile('Formulario');
    template.candidato = e.parameter.nombre || '';
    template.entrevistador = e.parameter.entrevistador || '';
    template.perfil = perfil;
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
  listarPerfiles().forEach(p => {
    try { procesarTodo(p.id); } catch (e) { console.log('Error perfil ' + p.id + ': ' + e); }
  });
}

function configurarEstiloPlanilla(perfilId) {
  const cfg = getConfig(perfilId);
  const sheet = getSheet(cfg);
  const aplicar = (col, lista) => {
    const rango = sheet.getRange(col + '2:' + col + '5000');
    rango.setDataValidation(null);
    rango.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(lista, true).build());
  };
  aplicar('I', ['Nuevo', 'Contactado', 'Entrevista agendada', 'Entrevistado', 'Avanza', 'En evaluación cliente', 'Oferta', 'Contratado', 'Descartado']);
  aplicar('L', ['A', 'B', 'C']);
  aplicar('N', ['A', 'B', 'C']);
  aplicar('O', ['A', 'B', 'C']);
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
  sheet.getRange(fila, 23).setValue(datos.formacion || 'Sin dato');   // Formación
  sheet.getRange(fila, 24).setValue(datos.zona || 'Sin dato');        // Zona
  sheet.getRange(fila, 27).setValue('Sin dato');                      // Posgrado
  sheet.getRange(fila, 28).setValue('Sin dato');                      // Excel
  sheet.getRange(fila, 29).setValue('Sin dato');                      // Power BI
  sheet.getRange(fila, 30).setValue('Sin dato');                      // IA
}

function insertarFila(sheet, datos, url, fecha) {
  sheet.appendRow([fecha, datos.nombre, datos.telefono, datos.dni, datos.nivel_ajuste, datos.requisitos, datos.resumen, url, 'Nuevo']);
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

function procesarScreening(data) {
  const cfg = getConfig(data.perfil || 'operaciones');
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

  sheet.getRange(fila, 26).setValue('LLEGUE resultado=' + data.resultado);
  const celdaL = sheet.getRange(fila, 12);
  celdaL.setDataValidation(null).setValue(data.resultado);
  const colores = { A: '#E6F2EC', B: '#FEF3D8', C: '#FDEAE4' };
  celdaL.setBackground(colores[data.resultado] || '#fff').setFontWeight('bold').setHorizontalAlignment('center');

  if (data.entrevistador) sheet.getRange(fila, 14).setValue(data.entrevistador);
  sheet.getRange(fila, 26).setValue('DEBUG resultado=' + data.resultado);
  aplicarEstadoNota(sheet, fila, 'screening', data.resultado);
  try {
    const folder = DriveApp.getFolderById(cfg.folder_pdf_screening);
    const blob = Utilities.newBlob(data.htmlFull, 'text/html', 'Screening_' + data.nombre.replace(/\s+/g, '_') + '.html');
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

function guardarEnPlanilla(nombreCandidato, htmlContent, ajustePonderacion, perfilId) {
  try {
    console.log("Perfil recibido: " + perfilId);
    const cfg = getConfig(perfilId || 'operaciones');
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
    const blob = Utilities.newBlob(htmlContent, MimeType.HTML, 'temp.html');
    const pdf = folder.createFile(blob.getAs(MimeType.PDF).setName('Competencias_' + nombreCandidato + '.pdf'));
    
    if (ajustePonderacion) sheet.getRange(fila, 16).setValue(ajustePonderacion);
    if (ajustePonderacion) aplicarEstadoNota(sheet, fila, 'competencias', ajustePonderacion);
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
  celda.setDataValidation(null).setValue(estado);
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

function setupInicial() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);

  const crear = (nombre, headers, datos) => {
    let sh = ss.getSheetByName(nombre);
    if (sh) {
      console.log(nombre + ' ya existe, salteo');
      return sh;
    }
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
    if (datos.length) sh.getRange(2, 1, datos.length, headers.length).setValues(datos);
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, headers.length);
    return sh;
  };

  crear('Config',
    ['perfil_id', 'nombre_perfil', 'sheet_id', 'hoja_nombre', 'folder_madre', 'folder_recibidos', 'folder_pdf_screening', 'folder_pdf_competencias', 'prompt_extra'],
    [['operaciones', 'Supervisor de Operaciones', SHEET_MASTER_ID, 'Hoja 1',
      '1AaGPCE-UPY4Lr1wC8zD7Ramk-YaRG_n7',
      '1euERruY_wGie6mpdfS4LcUSgo01Iodcy',
      '15Y4LVnbG-H3I-fxQsvnZQYGNfvqJNy75',
      '10Uv1EPmT3PD0v2kOx6TeyT6yKxpapTRW',
      'PERFIL: Supervisor de Operaciones - Servicios medioambientales/limpieza industrial. ZONA: Ramallo, Villa Constitución, San Nicolás, Arroyo Seco (otras = Fuera). EXPERIENCIA MÍNIMA: 3+ años liderando equipos en planta industrial. DESEABLE: rubro medioambiental, limpieza industrial, plantas siderúrgicas (Acindar, Ternium, Sidersa). FORMACIÓN: Técnico Electromecánico/Químico/Seguridad e Higiene (deseable, no excluyente). GÉNERO/EDAD: indistinto. BANDERAS ROJAS: solo experiencia administrativa, sin liderazgo, sin experiencia industrial.'
    ]]);

  crear('Perfiles_Hero',
    ['perfil_id', 'empresa', 'posicion', 'ubicacion', 'instancia', 'entrevistadoras', 'criterio_avanza', 'criterio_reserva'],
    [['operaciones', 'MACS', 'Supervisor de Operaciones', 'San Nicolás · Planta Ternium', 'Screening · 1ª entrevista', 'Ariel, Analía, Flavia, Marcos', '8.5', '6.5']]);

  crear('Perfiles_Screening_Eval',
    ['perfil_id', 'id_eval', 'etiqueta', 'peso'],
    [
      ['operaciones', 'e1', 'Experiencia en rubro medioambiental/limpieza industrial', 3],
      ['operaciones', 'e2', 'Liderazgo de equipos en planta (3+ años)', 3],
      ['operaciones', 'e3', 'Experiencia en plantas industriales pesadas', 2],
      ['operaciones', 'e4', 'Formación técnica (Electromecánica/Química/Seg. e Hig.)', 1],
      ['operaciones', 'e5', 'Manejo de seguridad e higiene', 2],
      ['operaciones', 'e6', 'Gestión de turnos y guardias', 1],
      ['operaciones', 'e7', 'Resolución de problemas operativos', 1.5]
    ]);

  crear('Perfiles_Screening_Val',
    ['perfil_id', 'id_val', 'etiqueta', 'opciones'],
    [
      ['operaciones', 'v1', 'Remuneración pretendida compatible', 'Sí|No|A revisar'],
      ['operaciones', 'v2', 'Disponibilidad para turnos rotativos', 'Sí|No'],
      ['operaciones', 'v3', 'Movilidad propia / vive en zona', 'Sí|No'],
      ['operaciones', 'v4', 'Preaviso razonable', 'Sí|No'],
      ['operaciones', 'v5', 'Interés real en la posición', 'Alto|Medio|Bajo'],
      ['operaciones', 'v6', 'Disponibilidad para guardias / urgencias', 'Sí|No']
    ]);

  crear('Perfiles_Competencias',
    ['perfil_id', 'id_comp', 'nombre_comp', 'descripcion', 'peso', 'preguntas'],
    [
      ['operaciones', 'c1', 'Liderazgo: Coordinación de Equipos de Trabajo',
        'Coordinar y guiar al equipo de operadores en planta, asegurando el cumplimiento de las metas operativas y promoviendo un buen clima de trabajo. (25% del tiempo según rol oficial)',
        3,
        'Contame de una situación donde tuviste que coordinar a tu equipo en una tarea operativa exigente dentro de una planta cliente. ¿Cómo organizaste el trabajo y qué resultado obtuviste?|Dame un ejemplo de cuando heredaste un equipo con problemas de desempeño o falta de compromiso. ¿Qué cambios implementaste y cómo respondieron?|Describí una vez que tuviste que tomar una decisión difícil con alguien de tu equipo (sancionar, reasignar, dar feedback duro). ¿Cómo lo manejaste?'],
      ['operaciones', 'c2', 'Planificación y Organización',
        'Estructurar y priorizar tareas y recursos, asegurando el cumplimiento de objetivos operativos dentro de los plazos establecidos. (Cubre el 50% del rol: operaciones diarias + informes + mantenimiento)',
        2.5,
        'Contame cómo organizás un día o turno típico: qué priorizás, cómo asignás recursos y cómo controlás el avance de las tareas.|Dame un ejemplo de cuando tuviste que reorganizar la planificación por un cambio inesperado del cliente, una falla de equipo o una baja de personal. ¿Cómo lo manejaste?|Describí cómo planificás el mantenimiento preventivo o la compra de insumos para que no te frene la operación. ¿Qué herramientas o sistemas usás?'],
      ['operaciones', 'c3', 'Resolución de Problemas y Toma de Decisiones',
        'Analizar situaciones complejas, identificar soluciones efectivas y tomar decisiones informadas bajo presión, incluyendo negociación con los 4 actores (operadores, base, cliente, sindicato).',
        2.5,
        'Contame de un problema operativo crítico que tuviste que resolver rápido con recursos limitados o información incompleta. ¿Qué decisión tomaste y por qué?|Dame un ejemplo de una negociación con un delegado sindical o un grupo de operarios un tema sensible. ¿Cómo lo encaraste y qué resultado obtuviste?|Describí una situación con intereses cruzados entre el cliente y tu empresa, o entre operación y seguridad. ¿Cómo decidiste y cómo justificaste tu decisión?'],
      ['operaciones', 'c4', 'Foco en el Cliente',
        'Representar a MACS frente al cliente Ternium en planta, construyendo confianza y dando respuesta efectiva a sus necesidades operativas.',
        2,
        'Contame una situación donde un referente del cliente (jefe de planta, supervisor) estaba disconforme con el servicio. ¿Cómo lo manejaste?|Dame un ejemplo de un pedido del cliente que era difícil de cumplir con los recursos disponibles. ¿Cómo lo equilibraste?|Describí cómo fue tu mejor relación de trabajo con un cliente industrial. ¿Qué hiciste para construirla y sostenerla?'],
      ['operaciones', 'c5', 'Comunicación Efectiva',
        'Transmitir información clara y precisa a diferentes niveles (operadores, cliente, base, contratistas), incluyendo informes operativos diarios y mensuales.',
        2,
        'Contame de una vez que tuviste que transmitir información técnica o un cambio importante a personas con perfiles muy distintos (operarios, cliente, jefe). ¿Cómo lo hiciste?|Dame un ejemplo de un malentendido o falla de comunicación que generó un problema operativo. ¿Cómo lo resolviste y qué aprendiste?|Describí cómo elaborás un informe operativo diario o mensual para gerencia. ¿Qué priorizás y cómo lo presentás?'],
      ['operaciones', 'c6', 'Conciencia de Seguridad y Medio Ambiente',
        'Garantizar el cumplimiento de procedimientos de seguridad, uso de EPP, integridad de equipos y cuidado del medio ambiente. (Eje central de la misión del puesto)',
        2,
        'Contame de una situación donde detectaste un riesgo grave de seguridad o un incumplimiento ambiental en planta. ¿Qué hiciste y cómo lo escalaste?|Dame un ejemplo de cómo lograste que tu equipo internalizara el uso correcto de EPP o procedimientos cuando había resistencia o costumbre arraigada.|Describí un incidente, cuasi-accidente o evento ambiental que tuviste que gestionar. ¿Cómo actuaste en el momento y qué medidas implementaste después?']
    ]);

  console.log('Setup completo. Revisá las pestañas del Sheet.');
}
function fixConfig() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const config = ss.getSheetByName('Config');
  
  // Actualizar fila 2 con los IDs correctos de testing
  config.getRange(2, 1, 1, 9).setValues([[
    'operaciones',
    'Supervisor de Operaciones',
    SHEET_MASTER_ID,
    'Control Busqueda',
    '1AaGPCE-UPY4Lr1wC8zD7Ramk-YaRG_n7',
    '1zxVv8zBHC5rKWdWC6DLZFq0ngLE3Xp9O',
    '1ePdZqFr4QZE1wwQZ4uHpj7SpbDItU68K',
    '13JBx0Xd_56LwtU0iC4FuM0Zuw9Ayv2I0',
    'PERFIL: Supervisor de Operaciones - Servicios medioambientales/limpieza industrial. ZONA: Ramallo, Villa Constitución, San Nicolás, Arroyo Seco (otras = Fuera). EXPERIENCIA MÍNIMA: 3+ años liderando equipos en planta industrial. DESEABLE: rubro medioambiental, limpieza industrial, plantas siderúrgicas (Acindar, Ternium, Sidersa). FORMACIÓN: Técnico Electromecánico/Químico/Seguridad e Higiene (deseable, no excluyente). GÉNERO/EDAD: indistinto. BANDERAS ROJAS: solo experiencia administrativa, sin liderazgo, sin experiencia industrial.'
  ]]);
  
  console.log('Config actualizado con carpetas de testing');
  console.log('folder_recibidos: 1zxVv8zBHC5rKWdWC6DLZFq0ngLE3Xp9O');
  console.log('folder_pdf_screening: 1ePdZqFr4QZE1wwQZ4uHpj7SpbDItU68K');
  console.log('folder_pdf_competencias: 13JBx0Xd_56LwtU0iC4FuM0Zuw9Ayv2I0');
  console.log('hoja_nombre: Control Busqueda');
}
function testDatosPerfil() {
  const datos = getDatosPerfil('operaciones');
  console.log('Hero: ' + JSON.stringify(datos.hero));
  console.log('Eval items: ' + datos.eval_items.length);
  console.log('Val items: ' + datos.val_items.length);
  console.log('Competencias: ' + datos.competencias.length);
  console.log('Entrevistadoras: ' + datos.entrevistadoras);
}
function crearTodoPerfiles() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const carpetaMadre = DriveApp.getFolderById('18bhn7cD2PkcK_SldSfKHrjiieHvO3tYH');

  // ==========================================
  // PASO 1: RENOMBRAR OPERACIONES
  // ==========================================
  const hoja = ss.getSheetByName('Control Busqueda');
  if (hoja) {
    hoja.setName('Supervisor de Operaciones');
    console.log('Pestaña renombrada a: Supervisor de Operaciones');
  }

  const config = ss.getSheetByName('Config');
  const datosConfig = config.getDataRange().getValues();
  for (let i = 1; i < datosConfig.length; i++) {
    if (datosConfig[i][0] === 'operaciones') {
      config.getRange(i + 1, 4).setValue('Supervisor de Operaciones');
      console.log('Config actualizado: hoja_nombre = Supervisor de Operaciones');
      break;
    }
  }

  try {
    var carpetasHijas = carpetaMadre.getFolders();
    while (carpetasHijas.hasNext()) {
      var carpeta = carpetasHijas.next();
      var nombre = carpeta.getName();
      if (nombre.toLowerCase().includes('operaci') || nombre.toLowerCase().includes('macs') || nombre === 'ATRACCION') {
        console.log('Carpeta encontrada: ' + nombre + ' - revisá si es la de Operaciones y renombrala a mano');
      }
    }
  } catch(e) {
    console.log('Revisá la carpeta de Operaciones en Drive manualmente');
  }

  // ==========================================
  // PASO 2: CREAR 3 PERFILES NUEVOS
  // ==========================================
  const perfiles = [
    {
      id: 'coord-rrhh',
      nombre: 'Coordinador de Recursos Humanos',
      empresa: '',
      ubicacion: '',
      prompt: 'PERFIL: Coordinador de Recursos Humanos. Completar criterios del perfil.'
    },
    {
      id: 'analista-datos',
      nombre: 'Analista de Datos',
      empresa: '',
      ubicacion: '',
      prompt: 'PERFIL: Analista de Datos. Completar criterios del perfil.'
    },
    {
      id: 'analista-compras',
      nombre: 'Analista de Compras y Contrataciones',
      empresa: '',
      ubicacion: '',
      prompt: 'PERFIL: Analista de Compras y Contrataciones. Completar criterios del perfil.'
    }
  ];

  const hero = ss.getSheetByName('Perfiles_Hero');
  const evalSheet = ss.getSheetByName('Perfiles_Screening_Eval');
  const valSheet = ss.getSheetByName('Perfiles_Screening_Val');
  const compSheet = ss.getSheetByName('Perfiles_Competencias');

  perfiles.forEach(p => {
    console.log('--- Creando perfil: ' + p.nombre + ' ---');

    var carpetaPerfil = carpetaMadre.createFolder(p.nombre);
    var folderMadre = carpetaPerfil.createFolder('CVs nuevos');
    var folderRecibidos = carpetaPerfil.createFolder('CVs procesados');
    var folderScreening = carpetaPerfil.createFolder('Screening PDFs');
    var folderCompetencias = carpetaPerfil.createFolder('Competencias PDFs');
    console.log('Carpetas Drive creadas: ' + carpetaPerfil.getUrl());

    var nombrePestana = p.nombre;
    var hojaNueva = ss.getSheetByName(nombrePestana);
    if (!hojaNueva) {
      hojaNueva = ss.insertSheet(nombrePestana);
      var headers = ['Fecha', 'Nombre', 'Teléfono', 'DNI', 'Nivel de Ajuste', 'Requisitos', 'Resumen Experiencia', 'Link CV', 'Estado', 'Fuente', 'Screening', 'Impresión', 'Detalle', 'Entrevistó', 'Entrevista competencias', 'Impresión', 'Detalle', 'Entrevistó', 'Entrevista Compañia', 'Impresión', 'Contactar', 'Correo'];
      hojaNueva.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
      hojaNueva.setFrozenRows(1);
      hojaNueva.autoResizeColumns(1, headers.length);
      console.log('Pestaña creada: ' + nombrePestana);
    }

    config.appendRow([p.id, p.nombre, SHEET_MASTER_ID, nombrePestana, folderMadre.getId(), folderRecibidos.getId(), folderScreening.getId(), folderCompetencias.getId(), p.prompt]);

    hero.appendRow([p.id, p.empresa || '(completar)', p.nombre, p.ubicacion || '(completar)', 'Screening · 1ª entrevista', 'Ariel, Analía, Flavia, Marcos', '8.5', '6.5']);

    var evalBase = [
      ['Experiencia en posiciones similares', 3],
      ['Formación académica requerida', 2],
      ['Habilidades técnicas del puesto', 2],
      ['Liderazgo / gestión de equipos', 1],
      ['Conocimiento del sector/rubro', 1]
    ];
    evalBase.forEach(function(ev, idx) {
      evalSheet.appendRow([p.id, 'e' + (idx + 1), ev[0], ev[1]]);
    });

    var valBase = [
      ['Remuneración pretendida compatible', 'Sí|No|A revisar'],
      ['Disponibilidad horaria', 'Sí|No'],
      ['Ubicación / movilidad', 'Sí|No'],
      ['Preaviso razonable', 'Sí|No'],
      ['Interés real en la posición', 'Alto|Medio|Bajo']
    ];
    valBase.forEach(function(v, idx) {
      valSheet.appendRow([p.id, 'v' + (idx + 1), v[0], v[1]]);
    });

    var compBase = [
      ['Competencia 1 (completar)', 'Descripción de la competencia', 2, 'Pregunta STAR 1|Pregunta STAR 2|Pregunta STAR 3'],
      ['Competencia 2 (completar)', 'Descripción de la competencia', 2, 'Pregunta STAR 1|Pregunta STAR 2|Pregunta STAR 3'],
      ['Competencia 3 (completar)', 'Descripción de la competencia', 2, 'Pregunta STAR 1|Pregunta STAR 2|Pregunta STAR 3'],
      ['Competencia 4 (completar)', 'Descripción de la competencia', 1, 'Pregunta STAR 1|Pregunta STAR 2|Pregunta STAR 3'],
      ['Competencia 5 (completar)', 'Descripción de la competencia', 1, 'Pregunta STAR 1|Pregunta STAR 2|Pregunta STAR 3']
    ];
    compBase.forEach(function(c, idx) {
      compSheet.appendRow([p.id, 'c' + (idx + 1), c[0], c[1], c[2], c[3]]);
    });

    console.log('✅ ' + p.nombre + ' COMPLETO');
  });

  console.log('');
  console.log('========================================');
  console.log('TODO LISTO');
  console.log('========================================');
  console.log('1. Pestaña Operaciones renombrada');
  console.log('2. 3 perfiles nuevos creados');
  console.log('3. Carpetas Drive creadas');
  console.log('');
  console.log('FALTA COMPLETAR EN EL SHEET:');
  console.log('- Config: prompt_extra de cada perfil');
  console.log('- Perfiles_Hero: empresa y ubicación');
  console.log('- Perfiles_Screening_Eval: criterios reales');
  console.log('- Perfiles_Competencias: competencias y preguntas STAR reales');
  console.log('- Renombrar carpeta de Operaciones en Drive a mano');
}
function verificarCarpetas() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const config = ss.getSheetByName('Config').getDataRange().getValues();
  
  for (let i = 1; i < config.length; i++) {
    const perfil = config[i][0];
    if (!perfil) continue;
    console.log('=== ' + perfil + ' ===');
    
    const campos = ['folder_madre', 'folder_recibidos', 'folder_pdf_screening', 'folder_pdf_competencias'];
    const indices = [4, 5, 6, 7];
    
    campos.forEach((campo, idx) => {
      const id = config[i][indices[idx]];
      try {
        const folder = DriveApp.getFolderById(id);
        console.log('✅ ' + campo + ': ' + folder.getName() + ' (' + id + ')');
      } catch(e) {
        console.log('❌ ' + campo + ': NO EXISTE o SIN ACCESO (' + id + ')');
      }
    });
    console.log('');
  }
}
function actualizarFolderMadreOperaciones() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const config = ss.getSheetByName('Config');
  const datos = config.getDataRange().getValues();
  
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][0] === 'operaciones') {
      config.getRange(i + 1, 5).setValue('1w5P1j3RrwX9DgjPCuapBuNrvF4jWd-hx');
      console.log('✅ folder_madre de operaciones actualizado');
      break;
    }
  }
}
function configurarHojasNuevas() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const URL_WEBAPP = 'https://script.google.com/macros/s/AKfycbx8s6xzvm2KPLuW9MPwUVG7OYzgFY7OAlWg1dkfCnXA-91J5jBNbM8XpXI79wHpWjc8/exec';
  
  const perfilesNuevos = [
    { id: 'coord-rrhh',       hoja: 'Coordinador de Recursos Humanos' },
    { id: 'analista-datos',   hoja: 'Analista de Datos' },
    { id: 'analista-compras', hoja: 'Analista de Compras y Contrataciones' }
  ];

  const modelo = ss.getSheetByName('Supervisor de Operaciones');
  if (!modelo) { console.log('❌ No encontré la hoja modelo'); return; }

  // Leer validaciones del modelo
  const validaciones = {};
  const colsConValidacion = [9, 12, 14, 16, 17]; // I, L, N, P, R
  colsConValidacion.forEach(col => {
    const regla = modelo.getRange(2, col).getDataValidation();
    if (regla) validaciones[col] = regla;
  });

  // Leer ancho de columnas del modelo
  const lastCol = modelo.getLastColumn();

  perfilesNuevos.forEach(p => {
    const hoja = ss.getSheetByName(p.hoja);
    if (!hoja) { console.log('❌ No encontré: ' + p.hoja); return; }

    console.log('Configurando: ' + p.hoja);

    // 1. Copiar anchos de columna
    for (let c = 1; c <= lastCol; c++) {
      hoja.setColumnWidth(c, modelo.getColumnWidth(c));
    }

    // 2. Aplicar validaciones (desplegables)
    Object.keys(validaciones).forEach(col => {
      const rango = hoja.getRange(2, Number(col), 100, 1);
      rango.setDataValidation(validaciones[col]);
    });

    // 3. Entrevistadores en col R (índice 17)
    const entrevistadores = ['Ariel', 'Marcos', 'Ever', 'Analía', 'Flavia'];
    const reglaEntrevistadores = SpreadsheetApp.newDataValidation()
      .requireValueInList(entrevistadores, true)
      .build();
    hoja.getRange(2, 18, 100, 1).setDataValidation(reglaEntrevistadores);
    hoja.getRange(2, 14, 100, 1).setDataValidation(reglaEntrevistadores); // col N también

    // 4. Impresión A/B/C en cols L, P
    const reglaABC = SpreadsheetApp.newDataValidation()
      .requireValueInList(['A', 'B', 'C'], true)
      .build();
    hoja.getRange(2, 12, 100, 1).setDataValidation(reglaABC); // L
    hoja.getRange(2, 16, 100, 1).setDataValidation(reglaABC); // P

    // 5. Estado en col I
    const reglaEstado = SpreadsheetApp.newDataValidation()
      .requireValueInList(['Nuevo', 'Contactado', 'Entrevista agendada', 'Entrevistado', 'Avanza', 'En evaluación cliente', 'Oferta', 'Contratado', 'Descartado'], true)
      .build();
    hoja.getRange(2, 9, 100, 1).setDataValidation(reglaEstado);

    // 6. Links de screening en col K (índice 10)
    for (let i = 2; i <= 101; i++) {
      const nombreCelda = 'B' + i;
      const entrevistadorCelda = 'R' + i;
      const urlScreening = URL_WEBAPP + '?perfil=' + p.id + '&nombre=\'&' + nombreCelda + '&\'&entrevistador=\'&' + entrevistadorCelda;
      hoja.getRange(i, 11).setFormula(
        '=IF(B' + i + '<>"",HYPERLINK("' + URL_WEBAPP + '?perfil=' + p.id + '&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(R' + i + '),"Iniciar Screening"),"")'
      );
    }

    // 7. Links de competencias en col O (índice 14... col O = 15)
    for (let i = 2; i <= 101; i++) {
      hoja.getRange(i, 15).setFormula(
        '=IF(B' + i + '<>"",HYPERLINK("' + URL_WEBAPP + '?tipo=competencias&perfil=' + p.id + '&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(R' + i + '),"Evaluar competencias"),"")'
      );
    }

    console.log('✅ ' + p.hoja + ' configurada');
  });

  // También actualizar Supervisor de Operaciones col K y O
  const hojaOps = ss.getSheetByName('Supervisor de Operaciones');
  if (hojaOps) {
    console.log('Actualizando links en Supervisor de Operaciones...');
    for (let i = 2; i <= 101; i++) {
      hojaOps.getRange(i, 11).setFormula(
        '=IF(B' + i + '<>"",HYPERLINK("' + URL_WEBAPP + '?perfil=operaciones&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(R' + i + '),"Iniciar Screening"),"")'
      );
      hojaOps.getRange(i, 15).setFormula(
        '=IF(B' + i + '<>"",HYPERLINK("' + URL_WEBAPP + '?tipo=competencias&perfil=operaciones&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(R' + i + '),"Evaluar competencias"),"")'
      );
    }
    console.log('✅ Supervisor de Operaciones actualizado');
  }

  console.log('');
  console.log('========================================');
  console.log('TODO LISTO');
  console.log('Los links se generan automáticamente cuando hay nombre en col B');
  console.log('========================================');
}
function crearPerfilSGI() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const carpetaMadre = DriveApp.getFolderById('18bhn7cD2PkcK_SldSfKHrjiieHvO3tYH');
  const URL_WEBAPP = 'https://script.google.com/macros/s/AKfycbx8s6xzvm2KPLuW9MPwUVG7OYzgFY7OAlWg1dkfCnXA-91J5jBNbM8XpXI79wHpWjc8/exec';

  const p = {
    id: 'analista-sgi',
    nombre: 'Analista de SGI',
    prompt: 'PERFIL: Analista de SGI. Completar criterios del perfil.'
  };

  // 1. Carpetas Drive
  const carpetaPerfil = carpetaMadre.createFolder(p.nombre);
  const folderMadre = carpetaPerfil.createFolder('CVs nuevos');
  const folderRecibidos = carpetaPerfil.createFolder('CVs procesados');
  const folderScreening = carpetaPerfil.createFolder('Screening PDFs');
  const folderCompetencias = carpetaPerfil.createFolder('Competencias PDFs');
  console.log('Carpetas Drive creadas: ' + carpetaPerfil.getUrl());

  // 2. Pestaña candidatos
  let hoja = ss.getSheetByName(p.nombre);
  if (!hoja) {
    hoja = ss.insertSheet(p.nombre);
    const headers = ['Fecha','Nombre','Teléfono','DNI','Nivel de Ajuste','Requisitos','Resumen Experiencia','Link CV','Estado','Fuente','Screening','Impresión','Detalle','Entrevistó','Entrevista competencias','Impresión','Detalle','Entrevistó','Entrevista Compañia','Impresión','Contactar','Correo'];
    hoja.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
    hoja.setFrozenRows(1);
    hoja.autoResizeColumns(1, headers.length);
  }

  // 3. Config
  ss.getSheetByName('Config').appendRow([
    p.id, p.nombre, SHEET_MASTER_ID, p.nombre,
    folderMadre.getId(), folderRecibidos.getId(),
    folderScreening.getId(), folderCompetencias.getId(),
    p.prompt
  ]);

  // 4. Hero
  ss.getSheetByName('Perfiles_Hero').appendRow([
    p.id, '(completar)', p.nombre, '(completar)',
    'Screening · 1ª entrevista', 'Ariel, Analía, Flavia, Marcos', '8.5', '6.5'
  ]);

  // 5. Screening Eval
  [['Experiencia en posiciones similares',3],['Formación académica requerida',2],['Habilidades técnicas del puesto',2],['Liderazgo / gestión de equipos',1],['Conocimiento del sector/rubro',1]]
  .forEach((ev,i) => ss.getSheetByName('Perfiles_Screening_Eval').appendRow([p.id,'e'+(i+1),ev[0],ev[1]]));

  // 6. Screening Val
  [['Remuneración pretendida compatible','Sí|No|A revisar'],['Disponibilidad horaria','Sí|No'],['Ubicación / movilidad','Sí|No'],['Preaviso razonable','Sí|No'],['Interés real en la posición','Alto|Medio|Bajo']]
  .forEach((v,i) => ss.getSheetByName('Perfiles_Screening_Val').appendRow([p.id,'v'+(i+1),v[0],v[1]]));

  // 7. Competencias placeholder
  [1,2,3,4,5].forEach(i => ss.getSheetByName('Perfiles_Competencias').appendRow([
    p.id,'c'+i,'Competencia '+i+' (completar)','Descripción de la competencia', i<=3?2:1,'Pregunta STAR 1|Pregunta STAR 2|Pregunta STAR 3'
  ]));

  // 8. Formato + links (igual que los otros)
  const modelo = ss.getSheetByName('Supervisor de Operaciones');
  const reglaEstado = SpreadsheetApp.newDataValidation().requireValueInList(['Nuevo','Contactado','Entrevista agendada','Entrevistado','Avanza','En evaluación cliente','Oferta','Contratado','Descartado'],true).build();
  const reglaABC = SpreadsheetApp.newDataValidation().requireValueInList(['A','B','C'],true).build();
  const reglaEntrev = SpreadsheetApp.newDataValidation().requireValueInList(['Ariel','Marcos','Ever','Analía','Flavia'],true).build();

  for (let c = 1; c <= modelo.getLastColumn(); c++) hoja.setColumnWidth(c, modelo.getColumnWidth(c));
  hoja.getRange(2,9,100,1).setDataValidation(reglaEstado);
  hoja.getRange(2,12,100,1).setDataValidation(reglaABC);
  hoja.getRange(2,16,100,1).setDataValidation(reglaABC);
  hoja.getRange(2,14,100,1).setDataValidation(reglaEntrev);
  hoja.getRange(2,18,100,1).setDataValidation(reglaEntrev);

  for (let i = 2; i <= 101; i++) {
    hoja.getRange(i,11).setFormula('=IF(B'+i+'<>"",HYPERLINK("'+URL_WEBAPP+'?perfil='+p.id+'&nombre="&ENCODEURL(B'+i+')&"&entrevistador="&ENCODEURL(R'+i+'),"Iniciar Screening"),"")');
    hoja.getRange(i,15).setFormula('=IF(B'+i+'<>"",HYPERLINK("'+URL_WEBAPP+'?tipo=competencias&perfil='+p.id+'&nombre="&ENCODEURL(B'+i+')&"&entrevistador="&ENCODEURL(R'+i+'),"Evaluar competencias"),"")');
  }

  console.log('✅ Analista de SGI COMPLETO');
  console.log('Links: ' + URL_WEBAPP + '?perfil=analista-sgi');
}
function corregirFormulasLinks() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  // URL actualizada a la que me pasaste
  const URL = 'https://script.google.com/macros/s/AKfycbxu9CtgETmsBIuYEh7JT_9USx8jncfWpDzJPW7UlGc/dev';
  
  const perfiles = [
    { hoja: 'Supervisor de Operaciones', id: 'operaciones' },
    { hoja: 'Coordinador de Recursos Humanos', id: 'coord-rrhh' },
    { hoja: 'Analista de Datos', id: 'analista-datos' },
    { hoja: 'Analista de Compras y Contrataciones', id: 'analista-compras' },
    { hoja: 'Analista de SGI', id: 'analista-sgi' }
  ];

  perfiles.forEach(p => {
    const hoja = ss.getSheetByName(p.hoja);
    if (!hoja) { return; }
    
    for (let i = 2; i <= 101; i++) {
      // Screening
      hoja.getRange(i, 11).setValue(
        '=IF(B'+i+'<>""; HYPERLINK("'+URL+'?tipo=screening&perfil='+p.id+'&nombre=" & ENCODEURL(B'+i+') & "&entrevistador=" & ENCODEURL(R'+i+'); "Iniciar Screening"); "")'
      );
      // Competencias
      hoja.getRange(i, 15).setValue(
        '=IF(B'+i+'<>""; HYPERLINK("'+URL+'?tipo=competencias&perfil='+p.id+'&nombre=" & ENCODEURL(B'+i+') & "&entrevistador=" & ENCODEURL(R'+i+'); "Evaluar competencias"); "")'
      );
    }
  });
  
  console.log('URL actualizada. Probá ahora.');
}
function testProcesarScreening() {
  try {
    const resultado = procesarScreening({
      perfil: 'operaciones',
      nombre: 'María Elisabet Digiovani',  // Nombre exacto de B2
      resultado: 'TEST_A',
      comentarios: 'Test',
      htmlFull: '<html></html>'
    });
    Logger.log('✓ ÉXITO: ' + resultado);
  } catch(e) {
    Logger.log('✗ ERROR: ' + e.message);
  }
}
function debugSheets() {
  const ss = SpreadsheetApp.openById('151V8cc-w00tP9C6mSxVaKNkRxgKXgUDI4L0KWXeNJGI');
  ss.getSheets().forEach(s => Logger.log('"' + s.getName() + '"'));
}
function debugNombre() {
  const ss = SpreadsheetApp.openById('151V8cc-w00tP9C6mSxVaKNkRxgKXgUDI4L0KWXeNJGI');
  const hoja = ss.getSheetByName('Analista de Compras y Contrataciones');
  const nombre = hoja.getRange('B2').getValue();
  Logger.log('Valor: [' + nombre + ']');
  Logger.log('Largo: ' + nombre.length);
  Logger.log('Chars: ' + [...nombre].map(c => c.codePointAt(0)).join(', '));
}
function testPerfilRoto() {
  const datos = getDatosPerfil('coord-rrhh');
  Logger.log('Hero: ' + JSON.stringify(datos.hero));
  Logger.log('Eval: ' + datos.eval_items.length);
  Logger.log('Val: ' + datos.val_items.length);
  Logger.log('Comp: ' + datos.competencias.length);
  Logger.log('Entrevistadoras: ' + datos.entrevistadoras);
}
function arreglarTodosLosLinks() {
  const URL = 'https://script.google.com/macros/s/AKfycbytekPKN31bZLxczDVknMr433K9KflzmrR-dCWsC1hLlx2MO7tbw6SrMjFXa1PnucQo/exec';
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const perfiles = [
    { hoja: 'Supervisor de Operaciones', id: 'operaciones' },
    { hoja: 'Coordinador de Recursos Humanos', id: 'coord-rrhh' },
    { hoja: 'Analista de Datos', id: 'analista-datos' },
    { hoja: 'Analista de Compras y Contrataciones', id: 'analista-compras' },
    { hoja: 'Analista de SGI', id: 'analista-sgi' }
  ];
  perfiles.forEach(p => {
    const hoja = ss.getSheetByName(p.hoja);
    if (!hoja) { console.log('No existe: ' + p.hoja); return; }
    for (let i = 2; i <= 101; i++) {
      hoja.getRange(i, 11).setFormula(
        '=IF(B' + i + '<>"";HYPERLINK("' + URL + '?perfil=' + p.id + '&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(N' + i + ');"Iniciar Screening");"")'
      );
      hoja.getRange(i, 15).setFormula(
        '=IF(B' + i + '<>"";HYPERLINK("' + URL + '?tipo=competencias&perfil=' + p.id + '&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(R' + i + ');"Evaluar competencias");"")'
      );
    }
    console.log('OK: ' + p.hoja);
  });
  console.log('LISTO.');
}
function cargarPostulacionPreguntas() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let sh = ss.getSheetByName('Postulacion_Preguntas');
  if (!sh) sh = ss.insertSheet('Postulacion_Preguntas');
  sh.clear();

  const headers = ['perfil_id', 'id_preg', 'etiqueta', 'tipo', 'opciones', 'requerido'];
  sh.getRange(1, 1, 1, 6).setValues([headers])
    .setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);

  const filas = [
    ['operaciones', 'p5', '¿Trabajaste en extracción de barros, tratamiento de residuos o limpieza industrial?', 'select', 'Sí|No', 'SI'],
    ['operaciones', 'p6', 'Nivel en ese tipo de operación', 'select', 'Básico|Intermedio|Avanzado', 'SI'],
    ['coord-rrhh', 'p5', '¿Con qué sistema de liquidación de sueldos trabajaste? (Tango, Bejerman, SAP, otro)', 'texto', '', 'SI'],
    ['coord-rrhh', 'p6', 'Nivel en ese sistema', 'select', 'Básico|Intermedio|Avanzado', 'SI'],
    ['analista-datos', 'p5', '¿Qué herramientas manejás? (Excel, SQL, Power BI, Python…)', 'texto', '', 'SI'],
    ['analista-datos', 'p6', 'Nivel general en esas herramientas', 'select', 'Básico|Intermedio|Avanzado', 'SI'],
    ['analista-compras', 'p5', '¿Qué ERP o sistema de compras usaste? (SAP, Oracle, otro)', 'texto', '', 'SI'],
    ['analista-compras', 'p6', 'Nivel en ese ERP', 'select', 'Básico|Intermedio|Avanzado', 'SI'],
    ['analista-sgi', 'p5', '¿Qué normas ISO manejás? (9001, 14001, 45001)', 'texto', '', 'SI'],
    ['analista-sgi', 'p6', 'Nivel de dominio de esas normas', 'select', 'Básico|Intermedio|Avanzado', 'SI']
  ];

  sh.getRange(2, 1, filas.length, 6).setValues(filas);
  sh.autoResizeColumns(1, 6);
  Logger.log('Cargadas ' + filas.length + ' preguntas dinámicas');
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
    const cfg = getConfig(payload.perfil || 'operaciones');
    const sheet = getSheet(cfg);

    const d = payload.datos || {};
    if (esDuplicado(d, sheet)) return 'DUPLICADO';

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
      JSON.stringify({ perfil: payload.perfil || 'operaciones', datos: d, declaradas: payload.declaradas || {}, cvId: file.getId() }),
      posgrado, d.excel || '', d.powerbi || '', ia
    ]]);
    ponerFormulaWhatsApp(sheet, fila);

    return 'OK';

  } catch (err) {
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
    if (!fila) return { pos: 'Búsqueda', loc: '' };
    const o = {};
    headers.forEach((h, i) => o[h] = fila[i]);
    const loc = [o.empresa, o.ubicacion].filter(Boolean).join(' · ');
    return { pos: String(o.posicion || 'Búsqueda'), loc: loc };
  } catch (err) {
    return { pos: 'Búsqueda', loc: '' };
  }
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
          'Posgrado: ' + (d.posgradoEstado ? d.posgradoEstado + (d.posgradoCual ? ' (' + d.posgradoCual + ')' : '') : 'Sin posgrado'),
          'Excel: ' + (d.excel || ''),
          'Power BI: ' + (d.powerbi || ''),
          'Inglés: ' + (d.ingles || ''),
          'Herramientas de IA: ' + (d.ia || '') + (d.iaTools ? ' — usa: ' + d.iaTools : ''),
          'Remuneración pretendida: ' + (d.remuneracion || ''),
          'Localidad: ' + (d.localidad || ''),
          'Movilidad propia: ' + (d.movilidad || ''),
          'Disponibilidad full-time: ' + (d.fulltime || ''),
          'Trabajo 100% presencial: ' + (d.presencial || '')
        ];
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
function verPerfiles() {
  listarPerfiles().forEach(p => Logger.log('[' + p.id + '] → ' + p.nombre));
}
function chequearHoja() {
  const cfg = getConfig('coord-rrhh');
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(cfg.hoja_nombre);
  Logger.log('Hoja: [' + cfg.hoja_nombre + ']');
  Logger.log('Existe: ' + (sh ? 'SÍ' : 'NO'));
  if (!sh) return;
  Logger.log('Headers 23-26: ' + JSON.stringify(sh.getRange(1, 23, 1, 4).getValues()[0]));
  Logger.log('Formula K2: ' + sh.getRange(2, 11).getFormula());
  Logger.log('Formula K150: ' + sh.getRange(150, 11).getFormula());
  Logger.log('Ultima fila: ' + sh.getLastRow());
}
function configurarCoordRRHH() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const cfg = getConfig('coord-rrhh');
  const sh = ss.getSheetByName(cfg.hoja_nombre);

  try {
    // headers 23-26
    sh.getRange(1, 23).setValue('Formación declarada').setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
    sh.getRange(1, 24).setValue('Zona').setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
    sh.getRange(1, 25).setValue('Intentos').setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
    sh.getRange(1, 26).setValue('_meta').setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
    Logger.log('Headers OK');

    // estados ampliados 2-250
    const estados = ['Pendiente','Nuevo','Contactado','Entrevista agendada','Entrevistado','Avanza','En evaluación cliente','Oferta','Contratado','Descartado','Revisar'];
    sh.getRange(2, 9, 249, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(estados, true).build());
    Logger.log('Estados OK');

    // fórmulas 102-250
    for (let i = 102; i <= 250; i++) {
      sh.getRange(i, 11).setFormula('=IF(B'+i+'<>"";HYPERLINK("'+URL_DEPLOY+'?perfil=coord-rrhh&nombre="&ENCODEURL(B'+i+')&"&entrevistador="&ENCODEURL(N'+i+');"Iniciar Screening");"")');
      sh.getRange(i, 15).setFormula('=IF(B'+i+'<>"";HYPERLINK("'+URL_DEPLOY+'?tipo=competencias&perfil=coord-rrhh&nombre="&ENCODEURL(B'+i+')&"&entrevistador="&ENCODEURL(R'+i+');"Evaluar competencias");"")');
    }
    Logger.log('Formulas OK');
    Logger.log('TODO LISTO en ' + cfg.hoja_nombre);
  } catch (e) {
    Logger.log('ERROR: ' + e);
  }
}
function chequearTodas() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  listarPerfiles().forEach(p => {
    let cfg;
    try { cfg = getConfig(p.id); } catch (e) { Logger.log('❌ [' + p.id + '] sin Config'); return; }
    const sh = ss.getSheetByName(cfg.hoja_nombre);
    if (!sh) { Logger.log('❌ [' + p.id + '] NO existe la hoja "' + cfg.hoja_nombre + '"'); return; }

    const headers = sh.getRange(1, 23, 1, 4).getValues()[0];
    const headersOK = headers[0] === 'Formación declarada' && headers[1] === 'Zona' && headers[2] === 'Intentos' && headers[3] === '_meta';
    const f250 = sh.getRange(250, 11).getFormula();
    const formulasOK = f250 && f250.indexOf('HYPERLINK') > -1;
    const val = sh.getRange(2, 9).getDataValidation();
    const tieneVal = val ? val.getCriteriaValues()[0].indexOf('Pendiente') > -1 : false;

    const estado = (headersOK && formulasOK && tieneVal) ? '✅ COMPLETA' : '⚠️ FALTA';
    Logger.log(estado + ' [' + p.id + '] → headers:' + (headersOK?'ok':'NO') + ' | formulas250:' + (formulasOK?'ok':'NO') + ' | estados:' + (tieneVal?'ok':'NO'));
  });
  Logger.log('--- fin chequeo ---');
}
function configurarOperaciones() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const cfg = getConfig('operaciones');
  const sh = ss.getSheetByName(cfg.hoja_nombre);
  try {
    const estados = ['Pendiente','Nuevo','Contactado','Entrevista agendada','Entrevistado','Avanza','En evaluación cliente','Oferta','Contratado','Descartado','Revisar'];
    sh.getRange(2, 9, 249, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(estados, true).build());
    Logger.log('Estados OK');

    for (let i = 2; i <= 250; i++) {
      sh.getRange(i, 11).setFormula('=IF(B'+i+'<>"";HYPERLINK("'+URL_DEPLOY+'?perfil=operaciones&nombre="&ENCODEURL(B'+i+')&"&entrevistador="&ENCODEURL(N'+i+');"Iniciar Screening");"")');
      sh.getRange(i, 15).setFormula('=IF(B'+i+'<>"";HYPERLINK("'+URL_DEPLOY+'?tipo=competencias&perfil=operaciones&nombre="&ENCODEURL(B'+i+')&"&entrevistador="&ENCODEURL(R'+i+');"Evaluar competencias");"")');
    }
    Logger.log('Formulas OK');
    Logger.log('TODO LISTO en ' + cfg.hoja_nombre);
  } catch (e) {
    Logger.log('ERROR: ' + e);
  }
}
function chequearTodas() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  listarPerfiles().forEach(p => {
    let cfg;
    try { cfg = getConfig(p.id); } catch (e) { Logger.log('❌ [' + p.id + '] sin Config'); return; }
    const sh = ss.getSheetByName(cfg.hoja_nombre);
    if (!sh) { Logger.log('❌ [' + p.id + '] NO existe la hoja'); return; }
    const headers = sh.getRange(1, 23, 1, 4).getValues()[0];
    const headersOK = headers[0] === 'Formación declarada' && headers[1] === 'Zona' && headers[2] === 'Intentos' && headers[3] === '_meta';
    const f250 = sh.getRange(250, 11).getFormula();
    const formulasOK = f250 && f250.indexOf('HYPERLINK') > -1;
    const val = sh.getRange(2, 9).getDataValidation();
    const tieneVal = val ? val.getCriteriaValues()[0].indexOf('Pendiente') > -1 : false;
    const estado = (headersOK && formulasOK && tieneVal) ? '✅ COMPLETA' : '⚠️ FALTA';
    Logger.log(estado + ' [' + p.id + '] → headers:' + (headersOK?'ok':'NO') + ' | formulas250:' + (formulasOK?'ok':'NO') + ' | estados:' + (tieneVal?'ok':'NO'));
  });
}
function chequearFila2() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  ['coord-rrhh','analista-datos','analista-compras','analista-sgi'].forEach(id => {
    const cfg = getConfig(id);
    const sh = ss.getSheetByName(cfg.hoja_nombre);
    const k = sh.getRange(2, 11).getFormula();
    const o = sh.getRange(2, 15).getFormula();
    const ok = (k.indexOf('HYPERLINK')>-1 && o.indexOf('HYPERLINK')>-1) ? '✅' : '❌ FALTA';
    Logger.log(ok + ' [' + id + ']');
  });
}
function completarLinksCoordRRHH() {
  const cfg = getConfig('coord-rrhh');
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(cfg.hoja_nombre);
  for (let i = 2; i <= 250; i++) {
    sh.getRange(i, 11).setFormula('=IF(B'+i+'<>"";HYPERLINK("'+URL_DEPLOY+'?perfil=coord-rrhh&nombre="&ENCODEURL(B'+i+')&"&entrevistador="&ENCODEURL(N'+i+');"Iniciar Screening");"")');
    sh.getRange(i, 15).setFormula('=IF(B'+i+'<>"";HYPERLINK("'+URL_DEPLOY+'?tipo=competencias&perfil=coord-rrhh&nombre="&ENCODEURL(B'+i+')&"&entrevistador="&ENCODEURL(R'+i+');"Evaluar competencias");"")');
  }
  Logger.log('Links completados 2-250 en ' + cfg.hoja_nombre);
}
function crearColumnasConocimientos() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const cols = { 27: 'Posgrado', 28: 'Excel', 29: 'Power BI', 30: 'IA' };
  listarPerfiles().forEach(p => {
    let cfg; try { cfg = getConfig(p.id); } catch (e) { return; }
    const sh = ss.getSheetByName(cfg.hoja_nombre);
    if (!sh) { Logger.log('No existe: ' + cfg.hoja_nombre); return; }
    Object.keys(cols).forEach(c => {
      const col = Number(c);
      if (sh.getRange(1, col).getValue() !== cols[col]) {
        sh.getRange(1, col).setValue(cols[col])
          .setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
      }
    });
    Logger.log('Columnas OK: ' + cfg.hoja_nombre);
  });
  Logger.log('Listo: 27 Posgrado · 28 Excel · 29 Power BI · 30 IA');
}
function verCarpetasMadre() {
  const config = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getDataRange().getValues();
  for (let i = 1; i < config.length; i++) {
    if (!config[i][0]) continue;
    const id = config[i][4]; // folder_madre
    try {
      const f = DriveApp.getFolderById(id);
      Logger.log('✅ [' + config[i][0] + '] → ' + f.getName());
    } catch(e) {
      Logger.log('❌ [' + config[i][0] + '] carpeta inválida o vacía (' + id + ')');
    }
  }
}
/* ============================================================
   CARGA PERFIL: Coordinador de Gestión de Personas (coord-rrhh)
   Cliente MACS S.A. — Pilar, Bs. As.
   Correr UNA vez: cargarPerfil_CoordRRHH()
   Completa: Config.prompt_extra, Perfiles_Hero, Postulacion_Preguntas
   ============================================================ */

function cargarPerfil_CoordRRHH() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const PERFIL = 'coord-rrhh';

  // ---------- 1) prompt_extra en Config ----------
  const promptExtra =
`PERFIL: Coordinador de Gestión de Personas (Generalista de RRHH). Cliente: MACS S.A., planta en Pilar (Bs. As.).
MISIÓN: Diseñar e implementar los procesos de gestión de personas (atracción, selección, desarrollo, desempeño, compensaciones y salida), gestión de relaciones laborales y medicina ocupacional, asegurando cumplimiento normativo y bienestar. Apoya a los líderes en la gestión de sus equipos.

EXPERIENCIA (EXCLUYENTE): mínimo 3 años en áreas de RRHH / Gestión de Personas. Preferente en empresas industriales o de servicios con PERSONAL DE CONVENIO.
FORMACIÓN: Graduado/a en RRHH, Relaciones del Trabajo o Psicología (deseable, no excluyente). Deseable actualización en legislación laboral y administración de personal.
CONOCIMIENTOS CLAVE:
- Relaciones laborales y normativa laboral argentina (fuerte; clave para el rol).
- Administración de personal, nómina/legajos electrónicos, beneficios.
- Excel AVANZADO.
- Sistema de gestión de personas (ERP de RRHH) nivel intermedio.
- Inglés intermedio (no excluyente).
ZONA: planta en Pilar (Bs. As.). Valorar cercanía / disponibilidad para presencial. Modalidad presencial.

BANDERAS ROJAS: alta rotación sin justificación; conflictos frecuentes con superiores; SIN experiencia en contextos de convenio; desconocimiento de normativa laboral; perfiles provenientes EXCLUSIVAMENTE de áreas administrativas sin contacto con personal operativo.

CONTINUIDAD LABORAL (criterio importante para Blumia): analizá el historial. VALORÁ permanencia y crecimiento dentro de cada empresa. PENALIZÁ rotación alta con saltos cortos y frecuentes (job hopping) SIN justificación razonable (cierres, fin de obra/proyecto, mudanzas). La estabilidad y la progresión interna suman fuerte.
DESEABLE QUE SUMA (no excluyente): experiencia en PLANTAS INDUSTRIALES o entornos productivos (ayuda a entender personal de convenio y dinámica operativa).

EVALUACIÓN DE BORDERLINE — "incluir y señalar": NO descartes mecánicamente a quien no cumple el 100%. Si a un candidato le falta algún requisito pero presenta FACTORES COMPENSATORIOS (trayectoria sólida, continuidad laboral ejemplar, formación destacada, experiencia transferible, logros relevantes), clasificalo como MEDIO y en el resumen marcá explícitamente "REVISAR:" explicando QUÉ le falta y QUÉ lo compensa, para que el selector decida. Ante la duda, INCLUIR y señalar, no filtrar. Reservá BAJO para quienes claramente no tienen experiencia real de RRHH, no llegan a la base mínima, o caen en banderas rojas claras.

NIVEL_AJUSTE:
- Alto: cumple experiencia excluyente (3+ años RRHH) + relaciones laborales/convenio + Excel avanzado, con buena continuidad laboral.
- Medio: cumple lo básico pero le falta algo (convenio, normativa, o un requisito deseable) O es un borderline con factores compensatorios → marcar "REVISAR:" en el resumen.
- Bajo: no llega a 3 años, sin experiencia real de RRHH, rotación injustificada marcada, o banderas rojas.`;

  const config = ss.getSheetByName('Config');
  const cfgData = config.getDataRange().getValues();
  let filaCfg = -1;
  for (let i = 1; i < cfgData.length; i++) if (cfgData[i][0] === PERFIL) { filaCfg = i + 1; break; }
  if (filaCfg === -1) { Logger.log('❌ No existe coord-rrhh en Config'); return; }
  config.getRange(filaCfg, 9).setValue(promptExtra); // col I = prompt_extra
  Logger.log('✅ Config.prompt_extra cargado');

  // ---------- 2) Perfiles_Hero ----------
  const hero = ss.getSheetByName('Perfiles_Hero');
  const heroData = hero.getDataRange().getValues();
  let filaHero = -1;
  for (let i = 1; i < heroData.length; i++) if (heroData[i][0] === PERFIL) { filaHero = i + 1; break; }
  // headers: perfil_id, empresa, posicion, ubicacion, instancia, entrevistadoras, criterio_avanza, criterio_reserva
  const filaHeroVals = [PERFIL, 'MACS S.A.', 'Coordinador de Gestión de Personas', 'Pilar, Bs. As.',
                        'Screening · 1ª entrevista', 'Ariel, Analía, Flavia, Marcos', '8.5', '6.5'];
  if (filaHero === -1) hero.appendRow(filaHeroVals);
  else hero.getRange(filaHero, 1, 1, filaHeroVals.length).setValues([filaHeroVals]);
  Logger.log('✅ Perfiles_Hero cargado');

  // ---------- 3) Postulacion_Preguntas (dinámicas de esta búsqueda) ----------
  const pp = ss.getSheetByName('Postulacion_Preguntas');
  // borrar las viejas de coord-rrhh
  const ppData = pp.getDataRange().getValues();
  for (let i = ppData.length - 1; i >= 1; i--) if (ppData[i][0] === PERFIL) pp.deleteRow(i + 1);
  // headers: perfil_id, id_preg, etiqueta, tipo, opciones, requerido
  const preguntas = [
    [PERFIL, 'p1', '¿Tenés experiencia gestionando personal de convenio / relaciones laborales?', 'select', 'Sí, directa|Sí, parcial|No', 'SI'],
    [PERFIL, 'p2', '¿Con qué sistema de gestión de personas (ERP de RRHH) trabajaste? (Ej: SAP SuccessFactors, Visma, Rhpro, Meta4…)', 'texto', '', 'SI'],
    [PERFIL, 'p3', 'Nivel en administración de nómina y legajos', 'select', 'Básico|Intermedio|Avanzado', 'SI']
  ];
  pp.getRange(pp.getLastRow() + 1, 1, preguntas.length, 6).setValues(preguntas);
  Logger.log('✅ Postulacion_Preguntas: ' + preguntas.length + ' preguntas cargadas');

  Logger.log('LISTO — perfil coord-rrhh configurado.');
}
function procesarAhoraCoordRRHH() {
  procesarTodo('coord-rrhh');
  Logger.log('Listo, revisá la hoja de coord-rrhh');
}
function verCarpetaCoord() {
  const cfg = getConfig('coord-rrhh');
  const f = DriveApp.getFolderById(cfg.folder_madre);
  Logger.log('Nombre: ' + f.getName());
  Logger.log('URL: ' + f.getUrl());
  const files = f.getFiles();
  let n = 0;
  while (files.hasNext()) { Logger.log('Dentro: ' + files.next().getName()); n++; }
  if (n === 0) Logger.log('(vacía)');
}
function chequearFunciones() {
  Logger.log('procesarDriveManual: ' + (typeof procesarDriveManual === 'function'));
  Logger.log('procesarMails: ' + (typeof procesarMails === 'function'));
  Logger.log('procesarTodo: ' + (typeof procesarTodo === 'function'));
  Logger.log('procesarConChatGPT: ' + (typeof procesarConChatGPT === 'function'));
}
function buscarCarlos() {
  const config = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getDataRange().getValues();
  for (let i = 1; i < config.length; i++) {
    if (!config[i][0]) continue;
    const perfil = config[i][0];
    [['folder_madre', config[i][4]], ['folder_recibidos', config[i][5]]].forEach(([nombre, id]) => {
      try {
        const f = DriveApp.getFolderById(id);
        const files = f.getFiles();
        while (files.hasNext()) {
          const file = files.next();
          if (file.getName().toLowerCase().includes('carlos') || file.getName().toLowerCase().includes('nuñez') || file.getName().toLowerCase().includes('nunez')) {
            Logger.log('🔴 ENCONTRADO en [' + perfil + '] ' + nombre + ': ' + file.getName() + ' → ' + file.getUrl());
          }
        }
      } catch(e) {}
    });
  }
  Logger.log('--- fin búsqueda ---');
}
function buscarMailCarlos() {
  const threads = GmailApp.search('carlos OR nuñez OR nunez newer_than:7d', 0, 20);
  Logger.log('Threads encontrados: ' + threads.length);
  threads.forEach(t => {
    Logger.log('—');
    Logger.log('Asunto: ' + t.getFirstMessageSubject());
    Logger.log('Fecha: ' + t.getLastMessageDate());
    t.getMessages().forEach(m => {
      const adj = m.getAttachments();
      adj.forEach(a => Logger.log('  Adjunto: ' + a.getName()));
    });
  });
}
function verHeaders() {
  const cfg = getConfig('coord-rrhh');
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(cfg.hoja_nombre);
  const h = sh.getRange(1, 1, 1, 31).getValues()[0];
  h.forEach((nombre, i) => Logger.log('Col ' + (i+1) + ': ' + nombre));
}
function ponerFormulaWhatsApp(sheet, fila) {
  // limpia el teléfono (solo dígitos), saca 0 y 15 iniciales, antepone 549
  const f =
    '=IF(C' + fila + '<>"";HYPERLINK("https://wa.me/549"&' +
    'REGEXREPLACE(REGEXREPLACE(SUBSTITUTE(SUBSTITUTE(C' + fila + ';"+54";"");"+";"");"[^0-9]";"");"^0";"");' +
    '"Abrir WhatsApp");"")';
  sheet.getRange(fila, 21).setFormula(f);
}
function ponerWhatsAppTodos() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  listarPerfiles().forEach(p => {
    let cfg; try { cfg = getConfig(p.id); } catch (e) { return; }
    const sh = ss.getSheetByName(cfg.hoja_nombre);
    if (!sh) return;
    for (let i = 2; i <= 250; i++) ponerFormulaWhatsApp(sh, i);
    Logger.log('WhatsApp OK: ' + cfg.hoja_nombre);
  });
  Logger.log('Listo.');
}
function verLinks() {
  const cfg = getConfig('coord-rrhh');
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(cfg.hoja_nombre);
  Logger.log('K2 (screening): ' + sh.getRange(2,11).getFormula());
  Logger.log('O2 (competencias): ' + sh.getRange(2,15).getFormula());
  Logger.log('U2 (whatsapp): ' + sh.getRange(2,21).getFormula());
}
function restaurarLinksTodos() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  listarPerfiles().forEach(p => {
    let cfg; try { cfg = getConfig(p.id); } catch(e){ return; }
    const sh = ss.getSheetByName(cfg.hoja_nombre);
    if (!sh) return;
    for (let i = 2; i <= 250; i++) {
      sh.getRange(i, 11).setFormula('=IF(B'+i+'<>"";HYPERLINK("'+URL_DEPLOY+'?perfil='+p.id+'&nombre="&ENCODEURL(B'+i+')&"&entrevistador="&ENCODEURL(N'+i+');"Iniciar Screening");"")');
      sh.getRange(i, 15).setFormula('=IF(B'+i+'<>"";HYPERLINK("'+URL_DEPLOY+'?tipo=competencias&perfil='+p.id+'&nombre="&ENCODEURL(B'+i+')&"&entrevistador="&ENCODEURL(R'+i+');"Evaluar competencias");"")');
    }
    Logger.log('Links OK: ' + cfg.hoja_nombre);
  });
  Logger.log('Listo, links restaurados.');
}
function armarFormulas() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  listarPerfiles().forEach(p => {
    let cfg; try { cfg = getConfig(p.id); } catch (e) { return; }
    const sh = ss.getSheetByName(cfg.hoja_nombre);
    if (!sh) return;
    for (let i = 2; i <= 250; i++) {
      sh.getRange(i, 11).setFormula('=IF(B' + i + '<>"";HYPERLINK("' + URL_DEPLOY + '?perfil=' + p.id + '&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(N' + i + ');"Iniciar Screening");"")');
      sh.getRange(i, 15).setFormula('=IF(B' + i + '<>"";HYPERLINK("' + URL_DEPLOY + '?tipo=competencias&perfil=' + p.id + '&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(R' + i + ');"Evaluar competencias");"")');
      sh.getRange(i, 21).setFormula('=IF(C' + i + '<>"";HYPERLINK("https://wa.me/549"&REGEXREPLACE(REGEXREPLACE(SUBSTITUTE(SUBSTITUTE(C' + i + ';"+54";"");"+";"");"[^0-9]";"");"^0";"");"Abrir WhatsApp");"")');
    }
    Logger.log('Fórmulas OK: ' + cfg.hoja_nombre);
  });
  Logger.log('Listo — screening + competencias + whatsapp en las 5 hojas.');
}
function setupEstados() {
  const ESTADOS = ['Pendiente','Nuevo','Contactado','Avanza a screening','Screening agendado','Avanza a competencias','Competencias agendada','Avanza a empresa','Entrevista empresa agendada','En cliente','Contratado','En reserva','Descartado sin screening','Descartado post screening','Descartado en competencias','Descartado en empresa'];
  const regla = SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).build();
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  listarPerfiles().forEach(p => {
    let cfg; try { cfg = getConfig(p.id); } catch (e) { return; }
    const sh = ss.getSheetByName(cfg.hoja_nombre);
    if (!sh) return;
    sh.getRange(2, 9, 249, 1).setDataValidation(regla);
  });
  Logger.log('Listo');
}
function probarEstado() {
  const cfg = getConfig('coord-rrhh');
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(cfg.hoja_nombre);
  aplicarEstadoNota(sh, 2, 'cv', 'A');
  Logger.log('Estado fila 2: ' + sh.getRange(2,9).getValue());
}
function ponerDesplegables() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const NOTAS = ['A','B+','B','C'];
  const ESTADOS = ['Pendiente','Nuevo','Contactado','Avanza a screening','Screening agendado','Avanza a competencias','Competencias agendada','Avanza a empresa','Entrevista empresa agendada','En cliente','Contratado','En reserva','Descartado sin screening','Descartado post screening','Descartado en competencias','Descartado en empresa'];
  const rNota = SpreadsheetApp.newDataValidation().requireValueInList(NOTAS, true).build();
  const rEstado = SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).build();

  listarPerfiles().forEach(p => {
    let cfg; try { cfg = getConfig(p.id); } catch(e){ return; }
    const sh = ss.getSheetByName(cfg.hoja_nombre);
    if (!sh) return;
    sh.getRange(2, 9, 249, 1).setDataValidation(rEstado);   // Estado
    sh.getRange(2, 5, 249, 1).setDataValidation(rNota);     // Nivel de Ajuste (CV)
    sh.getRange(2, 12, 249, 1).setDataValidation(rNota);    // Screening
    sh.getRange(2, 16, 249, 1).setDataValidation(rNota);    // Competencias
    sh.getRange(2, 20, 249, 1).setDataValidation(rNota);    // Empresa
  });
  Logger.log('Listo');
}
function ponerDesplegablesRRHH() {
  const cfg = getConfig('coord-rrhh');
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(cfg.hoja_nombre);
  Logger.log('Hoja: ' + cfg.hoja_nombre);
  Logger.log('Existe: ' + (sh ? 'SI' : 'NO'));
  if (!sh) return;
  const NOTAS = ['A','B+','B','C'];
  const ESTADOS = ['Pendiente','Nuevo','Contactado','Avanza a screening','Screening agendado','Avanza a competencias','Competencias agendada','Avanza a empresa','Entrevista empresa agendada','En cliente','Contratado','En reserva','Descartado sin screening','Descartado post screening','Descartado en competencias','Descartado en empresa'];
  const rNota = SpreadsheetApp.newDataValidation().requireValueInList(NOTAS, true).build();
  const rEstado = SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).build();
  sh.getRange(2, 9, 249, 1).setDataValidation(rEstado);
  sh.getRange(2, 5, 249, 1).setDataValidation(rNota);
  sh.getRange(2, 12, 249, 1).setDataValidation(rNota);
  sh.getRange(2, 16, 249, 1).setDataValidation(rNota);
  sh.getRange(2, 20, 249, 1).setDataValidation(rNota);
  Logger.log('Desplegables puestos en ' + cfg.hoja_nombre);
}
function migrarNotasViejas() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const map = { 'ALTO':'A', 'MEDIO':'B', 'BAJO':'C' };
  listarPerfiles().forEach(p => {
    let cfg; try { cfg = getConfig(p.id); } catch(e){ return; }
    const sh = ss.getSheetByName(cfg.hoja_nombre);
    if (!sh) return;
    const last = sh.getLastRow();
    for (let f = 2; f <= last; f++) {
      const v = String(sh.getRange(f, 5).getValue() || '').trim().toUpperCase();
      if (map[v]) {
        sh.getRange(f, 5).setValue(map[v]);
      }
    }
  });
  Logger.log('Notas migradas');
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
function cargarPerfil_AnalistaDatos() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const PERFIL = 'analista-datos';

  const promptExtra =
`PERFIL: Analista de Datos. Cliente: MACS S.A., planta en Pilar (Bs. As.). Reporta a Dirección de Ingeniería y Proyectos.
MISIÓN: Centralizar, procesar y transformar el flujo de información operativa en datos para la toma de decisiones. Diseñar e implementar tableros de control (KPIs) que visibilicen en tiempo real la salud económica y operativa de los proyectos, optimizar asignación de recursos y dar soporte analítico a la dirección.

EXPERIENCIA (EXCLUYENTE): mínimo 3 años en análisis de datos. Preferente en construcción, ingeniería, minería, oil & gas, medio ambiente o industria con proyectos de campo.
FORMACIÓN (EXCLUYENTE): Graduado/a en Ciencias Económicas (Lic. en Administración, Economía, Contador) o afín. Deseable posgrado/diplomatura en Business Intelligence / análisis de datos.
CONOCIMIENTOS CLAVE:
- Excel AVANZADO (excluyente).
- Business Intelligence (BI), SQL, DAX nivel intermedio (clave para el rol).
- Diseño y automatización de dashboards / KPIs dinámicos.
- Análisis económico-financiero de proyectos (flujo de caja, costos directos/indirectos, márgenes, control de desvíos vs presupuesto y plazos).
- IA generativa aplicada (ChatGPT, Claude, Copilot) para optimización de fórmulas, depuración SQL/DAX y aceleración del análisis (excluyente).
- Inglés intermedio (no excluyente).
ZONA: planta en Pilar (Bs. As.). Modalidad presencial. Valorar cercanía / disponibilidad.

CONTINUIDAD LABORAL (criterio importante para Blumia): analizá el historial. VALORÁ permanencia y crecimiento dentro de cada empresa. PENALIZÁ rotación alta con saltos cortos y frecuentes (job hopping) SIN justificación razonable (cierres, fin de obra/proyecto, mudanzas). La estabilidad y la progresión interna suman fuerte.

EVALUACIÓN DE BORDERLINE — "incluir y señalar": NO descartes mecánicamente a quien no cumple el 100%. Si a un candidato le falta algún requisito pero presenta FACTORES COMPENSATORIOS (trayectoria sólida, continuidad laboral ejemplar, formación destacada, experiencia transferible, logros relevantes), clasificalo como MEDIO y en el resumen marcá explícitamente "REVISAR:" explicando QUÉ le falta y QUÉ lo compensa, para que el selector decida. Ante la duda, INCLUIR y señalar, no filtrar. Reservá BAJO para quienes claramente no tienen experiencia real de análisis de datos, no llegan a la base mínima, o caen en banderas rojas claras.

NIVEL_AJUSTE:
- Alto: cumple experiencia excluyente (3+ años) + Cs. Económicas + Excel avanzado + BI/SQL + IA generativa, con buena continuidad laboral.
- Medio: cumple lo básico pero le falta algo (BI/SQL flojo, sin IA, o un deseable) O es un borderline con factores compensatorios → marcar "REVISAR:" en el resumen.
- Bajo: no llega a 3 años, sin experiencia real de análisis de datos, formación fuera de Cs. Económicas/afín sin compensación, o banderas rojas.`;

  const config = ss.getSheetByName('Config');
  const cfgData = config.getDataRange().getValues();
  let filaCfg = -1;
  for (let i = 1; i < cfgData.length; i++) if (cfgData[i][0] === PERFIL) { filaCfg = i + 1; break; }
  if (filaCfg === -1) { Logger.log('❌ No existe analista-datos en Config'); return; }
  config.getRange(filaCfg, 9).setValue(promptExtra);
  Logger.log('✅ Config.prompt_extra cargado');

  const hero = ss.getSheetByName('Perfiles_Hero');
  const heroData = hero.getDataRange().getValues();
  let filaHero = -1;
  for (let i = 1; i < heroData.length; i++) if (heroData[i][0] === PERFIL) { filaHero = i + 1; break; }
  const filaHeroVals = [PERFIL, 'MACS S.A.', 'Analista de Datos', 'Pilar, Bs. As.',
                        'Screening · 1ª entrevista', 'Ariel, Analía, Flavia, Marcos', '8.5', '6.5'];
  if (filaHero === -1) hero.appendRow(filaHeroVals);
  else hero.getRange(filaHero, 1, 1, filaHeroVals.length).setValues([filaHeroVals]);
  Logger.log('✅ Perfiles_Hero cargado');

  const pp = ss.getSheetByName('Postulacion_Preguntas');
  const ppData = pp.getDataRange().getValues();
  for (let i = ppData.length - 1; i >= 1; i--) if (ppData[i][0] === PERFIL) pp.deleteRow(i + 1);
  const preguntas = [
    [PERFIL, 'p1', '¿Con qué herramientas de Business Intelligence trabajaste? (Ej: Power BI, Tableau, Looker, Qlik…)', 'texto', '', 'SI'],
    [PERFIL, 'p2', 'Nivel de SQL', 'select', 'Básico|Intermedio|Avanzado', 'SI'],
    [PERFIL, 'p3', '¿Tenés experiencia en análisis económico-financiero de proyectos (costos, márgenes, flujo de caja)?', 'select', 'Sí, directa|Sí, parcial|No', 'SI']
  ];
  pp.getRange(pp.getLastRow() + 1, 1, preguntas.length, 6).setValues(preguntas);
  Logger.log('✅ Postulacion_Preguntas: ' + preguntas.length + ' preguntas cargadas');

  Logger.log('LISTO — perfil analista-datos configurado.');
}

function cargarPerfil_AnalistaCompras() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const PERFIL = 'analista-compras';

  const promptExtra =
`PERFIL: Analista de Compras y Contrataciones. Cliente: MACS S.A., planta en Pilar (Bs. As.). Reporta a Directora de Administración y Gestión.
MISIÓN: Gestionar de manera eficiente y ágil el proceso de compras y contrataciones, asegurando el abastecimiento oportuno de materiales, equipos y servicios especializados/tercerizados requeridos para proyectos de ingeniería, infraestructura y servicios ambientales.

EXPERIENCIA (EXCLUYENTE): mínimo 5 años en compras y contrataciones. Preferente en construcción, ingeniería, minería, oil & gas, medio ambiente o industria con proyectos de campo.
FORMACIÓN (EXCLUYENTE): Graduado/a en Ingeniería Industrial, Administración de Empresas o afín. Deseable posgrado/diplomatura/especialización en Compras o Supply Chain.
CONOCIMIENTOS CLAVE:
- Gestión del proceso de compras de punta a punta (solicitud, cotización, adjudicación, OC, seguimiento, recepción, pago).
- Negociación, cuadros comparativos, licitaciones y concursos de precios, administración de contratos a largo plazo.
- Administración de cartera de proveedores y cumplimiento normativo/técnico.
- Compras en el exterior / importación: despachantes, pagos, comex.
- Gestión de stock de materiales e insumos.
- Excel AVANZADO (excluyente).
- ERP o sistema de gestión de compras (SAP, Oracle, Odoo o similar) — EXCLUYENTE.
- Manejo de IA (no excluyente). Inglés (no excluyente).
ZONA: planta en Pilar (Bs. As.). Modalidad presencial 8-17. Valorar cercanía / disponibilidad.

OBSERVACIÓN DEL CLIENTE: importancia de antecedentes comprobables.

CONTINUIDAD LABORAL (criterio importante para Blumia): analizá el historial. VALORÁ permanencia y crecimiento dentro de cada empresa. PENALIZÁ rotación alta con saltos cortos y frecuentes (job hopping) SIN justificación razonable (cierres, fin de obra/proyecto, mudanzas). La estabilidad y la progresión interna suman fuerte.

EVALUACIÓN DE BORDERLINE — "incluir y señalar": NO descartes mecánicamente a quien no cumple el 100%. Si a un candidato le falta algún requisito pero presenta FACTORES COMPENSATORIOS (trayectoria sólida, continuidad laboral ejemplar, formación destacada, experiencia transferible, logros relevantes), clasificalo como MEDIO y en el resumen marcá explícitamente "REVISAR:" explicando QUÉ le falta y QUÉ lo compensa, para que el selector decida. Ante la duda, INCLUIR y señalar, no filtrar. Reservá BAJO para quienes claramente no tienen experiencia real de compras, no llegan a la base mínima, o caen en banderas rojas claras.

NIVEL_AJUSTE:
- Alto: cumple experiencia excluyente (5+ años compras) + formación + Excel avanzado + ERP de compras, con buena continuidad laboral. Suma importación/comex.
- Medio: cumple lo básico pero le falta algo (sin importación, ERP flojo, o un deseable) O es un borderline con factores compensatorios → marcar "REVISAR:" en el resumen.
- Bajo: no llega a 5 años, sin experiencia real de compras, sin ERP, o banderas rojas.`;

  const config = ss.getSheetByName('Config');
  const cfgData = config.getDataRange().getValues();
  let filaCfg = -1;
  for (let i = 1; i < cfgData.length; i++) if (cfgData[i][0] === PERFIL) { filaCfg = i + 1; break; }
  if (filaCfg === -1) { Logger.log('❌ No existe analista-compras en Config'); return; }
  config.getRange(filaCfg, 9).setValue(promptExtra);
  Logger.log('✅ Config.prompt_extra cargado');

  const hero = ss.getSheetByName('Perfiles_Hero');
  const heroData = hero.getDataRange().getValues();
  let filaHero = -1;
  for (let i = 1; i < heroData.length; i++) if (heroData[i][0] === PERFIL) { filaHero = i + 1; break; }
  const filaHeroVals = [PERFIL, 'MACS S.A.', 'Analista de Compras y Contrataciones', 'Pilar, Bs. As.',
                        'Screening · 1ª entrevista', 'Ariel, Analía, Flavia, Marcos', '8.5', '6.5'];
  if (filaHero === -1) hero.appendRow(filaHeroVals);
  else hero.getRange(filaHero, 1, 1, filaHeroVals.length).setValues([filaHeroVals]);
  Logger.log('✅ Perfiles_Hero cargado');

  const pp = ss.getSheetByName('Postulacion_Preguntas');
  const ppData = pp.getDataRange().getValues();
  for (let i = ppData.length - 1; i >= 1; i--) if (ppData[i][0] === PERFIL) pp.deleteRow(i + 1);
  const preguntas = [
    [PERFIL, 'p1', '¿Con qué ERP o sistema de gestión de compras trabajaste? (Ej: SAP, Oracle, Odoo…)', 'texto', '', 'SI'],
    [PERFIL, 'p2', '¿Tenés experiencia en compras en el exterior / importación (comex, despachantes)?', 'select', 'Sí, directa|Sí, parcial|No', 'SI'],
    [PERFIL, 'p3', '¿Gestionaste el ciclo de compras de punta a punta (cotización, adjudicación, OC, recepción, pago)?', 'select', 'Sí, completo|Sí, parcial|No', 'SI']
  ];
  pp.getRange(pp.getLastRow() + 1, 1, preguntas.length, 6).setValues(preguntas);
  Logger.log('✅ Postulacion_Preguntas: ' + preguntas.length + ' preguntas cargadas');

  Logger.log('LISTO — perfil analista-compras configurado.');
}

function cargarPerfil_AnalistaSGI() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const PERFIL = 'analista-sgi';

  const promptExtra =
`PERFIL: Analista de Sistema de Gestión (Integrado). Cliente: MACS S.A., planta en Pilar (Bs. As.). Reporta a Directora de Administración y Gestión.
MISIÓN: Asegurar la correcta implementación, mantenimiento y mejora del Sistema de Gestión Integrado (calidad, seguridad y medio ambiente), garantizando cumplimiento legal, normativo y contractual. Sostener estándares con foco en el cuidado de las personas y la integridad de la empresa.

EXPERIENCIA (EXCLUYENTE): mínimo 3 años en gestión e implementación de sistemas de calidad, seguridad y medio ambiente. Preferente en proyectos técnicos e industriales (construcción, ingeniería, industrial, químico, minería).
FORMACIÓN: Graduado/a en Ingeniería (Industrial, en Seguridad, Ambiental o afín) — no excluyente.
CONOCIMIENTOS CLAVE:
- Normas ISO 9001, ISO 45001, ISO 14001 — EXCLUYENTE (cursos/diplomatura/certificaciones).
- Auditorías internas y externas; gestión de hallazgos y no conformidades.
- Administración y control documental del sistema de gestión (procedimientos, instructivos, registros).
- Cumplimiento de requisitos legales y normativos (calidad, seguridad, MA).
- Auditor interno / capacitador en sistemas de gestión (deseable).
- Excel AVANZADO. ERP (SAP, Oracle, Odoo o similar) intermedio. Software de gestión documental.
- Conocimientos de IA.
ZONA: planta en Pilar (Bs. As.). Modalidad presencial. Valorar cercanía / disponibilidad.

CONTINUIDAD LABORAL (criterio importante para Blumia): analizá el historial. VALORÁ permanencia y crecimiento dentro de cada empresa. PENALIZÁ rotación alta con saltos cortos y frecuentes (job hopping) SIN justificación razonable (cierres, fin de obra/proyecto, mudanzas). La estabilidad y la progresión interna suman fuerte.

EVALUACIÓN DE BORDERLINE — "incluir y señalar": NO descartes mecánicamente a quien no cumple el 100%. Si a un candidato le falta algún requisito pero presenta FACTORES COMPENSATORIOS (trayectoria sólida, continuidad laboral ejemplar, formación destacada, experiencia transferible, logros relevantes), clasificalo como MEDIO y en el resumen marcá explícitamente "REVISAR:" explicando QUÉ le falta y QUÉ lo compensa, para que el selector decida. Ante la duda, INCLUIR y señalar, no filtrar. Reservá BAJO para quienes claramente no tienen experiencia real en sistemas de gestión, no llegan a la base mínima, o caen en banderas rojas claras.

NIVEL_AJUSTE:
- Alto: cumple experiencia excluyente (3+ años SGI) + ISO 9001/45001/14001 + auditorías + Excel avanzado, con buena continuidad laboral. Suma Ingeniería y experiencia industrial.
- Medio: cumple lo básico pero le falta algo (una norma, auditorías flojas, o un deseable) O es un borderline con factores compensatorios → marcar "REVISAR:" en el resumen.
- Bajo: no llega a 3 años, sin ISO, sin experiencia real en sistemas de gestión, o banderas rojas.`;

  const config = ss.getSheetByName('Config');
  const cfgData = config.getDataRange().getValues();
  let filaCfg = -1;
  for (let i = 1; i < cfgData.length; i++) if (cfgData[i][0] === PERFIL) { filaCfg = i + 1; break; }
  if (filaCfg === -1) { Logger.log('❌ No existe analista-sgi en Config'); return; }
  config.getRange(filaCfg, 9).setValue(promptExtra);
  Logger.log('✅ Config.prompt_extra cargado');

  const hero = ss.getSheetByName('Perfiles_Hero');
  const heroData = hero.getDataRange().getValues();
  let filaHero = -1;
  for (let i = 1; i < heroData.length; i++) if (heroData[i][0] === PERFIL) { filaHero = i + 1; break; }
  const filaHeroVals = [PERFIL, 'MACS S.A.', 'Analista de Sistema de Gestión', 'Pilar, Bs. As.',
                        'Screening · 1ª entrevista', 'Ariel, Analía, Flavia, Marcos', '8.5', '6.5'];
  if (filaHero === -1) hero.appendRow(filaHeroVals);
  else hero.getRange(filaHero, 1, 1, filaHeroVals.length).setValues([filaHeroVals]);
  Logger.log('✅ Perfiles_Hero cargado');

  const pp = ss.getSheetByName('Postulacion_Preguntas');
  const ppData = pp.getDataRange().getValues();
  for (let i = ppData.length - 1; i >= 1; i--) if (ppData[i][0] === PERFIL) pp.deleteRow(i + 1);
  const preguntas = [
    [PERFIL, 'p1', '¿Con qué normas ISO trabajaste? (Ej: 9001, 45001, 14001)', 'texto', '', 'SI'],
    [PERFIL, 'p2', '¿Tenés experiencia liderando o participando en auditorías internas/externas?', 'select', 'Sí, liderando|Sí, participando|No', 'SI'],
    [PERFIL, 'p3', '¿Sos auditor interno certificado?', 'select', 'Sí|No', 'SI']
  ];
  pp.getRange(pp.getLastRow() + 1, 1, preguntas.length, 6).setValues(preguntas);
  Logger.log('✅ Postulacion_Preguntas: ' + preguntas.length + ' preguntas cargadas');

  Logger.log('LISTO — perfil analista-sgi configurado.');
}
function cargarCompetencias_MACS() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const sh = ss.getSheetByName('Perfiles_Competencias');
  const data = sh.getDataRange().getValues();

  const DATOS = {
    'coord-rrhh': [
      ['Impacto e influencia', 'Capacidad de generar adhesión y movilizar a líderes y colaboradores hacia decisiones de gestión de personas, sin autoridad formal directa. Evaluar evidencia de haber instalado políticas, cambiado prácticas o destrabado conflictos influyendo sobre otras áreas.',
        'Contame de una política o cambio de RRHH que impulsaste y que al principio tuvo resistencia de líderes o colaboradores. ¿Cómo lograste que lo adoptaran?|Dame un ejemplo donde tuviste que convencer a un gerente de una decisión de personas que él no compartía. ¿Qué argumentos usaste y qué resultado obtuviste?|Describí una situación donde tu intervención cambió el rumbo de una decisión importante sobre el equipo. ¿Cómo influiste?'],
      ['Comunicación efectiva', 'Transmitir mensajes claros y adaptados a distintos públicos (dirección, líderes, personal de convenio, delegados), incluyendo temas sensibles. Evaluar evidencia de manejo de comunicaciones difíciles y escucha activa.',
        'Contame de una vez que tuviste que comunicar una noticia difícil al personal (reestructuración, cambio de condiciones, desvinculación). ¿Cómo lo encaraste?|Dame un ejemplo donde un mensaje mal interpretado generó un conflicto interno. ¿Cómo lo resolviste?|Describí cómo adaptás tu comunicación cuando hablás con un gerente versus con personal operativo de convenio.'],
      ['Negociación', 'Alcanzar acuerdos equilibrados en escenarios de intereses contrapuestos (relaciones laborales, convenio, delegados, compensaciones). Evaluar evidencia de manejo de conflictos laborales y cierre de acuerdos sostenibles.',
        'Contame de una negociación con un delegado o representante gremial. ¿Cómo la preparaste y a qué acuerdo llegaste?|Dame un ejemplo de un conflicto laboral individual o colectivo que tuviste que destrabar. ¿Cómo lo manejaste?|Describí una situación donde tuviste que equilibrar lo que pedía el colaborador con lo que la empresa podía dar. ¿Cómo cerraste?'],
      ['Foco en el cliente', 'Entender y dar respuesta a las necesidades de los líderes y colaboradores como clientes internos, sosteniendo cercanía y calidad de servicio. Evaluar evidencia de orientación al servicio interno.',
        'Contame de un líder al que asististe en la gestión de su equipo. ¿Qué necesitaba y cómo lo acompañaste?|Dame un ejemplo de una demanda interna difícil de satisfacer con los recursos disponibles. ¿Cómo la resolviste?|Describí cómo construís una relación de confianza con los líderes a los que das servicio desde RRHH.'],
      ['Trabajo en equipo', 'Colaborar de forma efectiva con otras áreas y dentro del equipo de gestión, aportando al objetivo común por sobre lo individual. Evaluar evidencia de articulación transversal.',
        'Contame de un proyecto de RRHH que requirió coordinar con otras áreas (operaciones, administración, seguridad). ¿Cómo articulaste?|Dame un ejemplo donde tuviste que dejar de lado tu postura por el bien del equipo o del proyecto. ¿Qué pasó?|Describí una situación de tensión dentro de tu propio equipo. ¿Qué hiciste para recomponer la colaboración?']
    ],
    'analista-datos': [
      ['Pensamiento Analítico y Rigurosidad', 'Descomponer problemas complejos en datos confiables y conclusiones sólidas, con atención al detalle y control de calidad de la información. Evaluar evidencia de análisis profundo y validación de datos.',
        'Contame de un análisis complejo que hiciste donde los datos venían sucios o incompletos. ¿Cómo aseguraste que la conclusión fuera confiable?|Dame un ejemplo de un desvío o patrón que detectaste en los datos que otros no habían visto. ¿Qué impacto tuvo?|Describí tu proceso para validar la calidad de un dataset antes de presentar un resultado a la dirección.'],
      ['Proactividad y Autonomía', 'Avanzar sin supervisión constante, anticiparse a necesidades de información y proponer análisis no solicitados que agregan valor. Evaluar evidencia de iniciativa propia.',
        'Contame de un tablero o análisis que desarrollaste por iniciativa propia, sin que te lo pidieran. ¿Qué te llevó a hacerlo y cómo se usó?|Dame un ejemplo de cuando anticipaste un problema mirando los datos antes de que escalara. ¿Qué hiciste?|Describí una situación donde resolviste sola/o un bloqueo técnico (SQL, fórmula, fuente de datos) sin ayuda.'],
      ['Comunicación Efectiva y Capacidad de Síntesis', 'Traducir análisis técnicos en mensajes claros y accionables para audiencias no técnicas (dirección, gerencia). Evaluar evidencia de storytelling con datos y síntesis.',
        'Contame de una vez que tuviste que presentar un análisis complejo a la dirección. ¿Cómo lo simplificaste para que tomaran una decisión?|Dame un ejemplo de un dashboard que rediseñaste para que fuera más claro o útil. ¿Qué cambiaste?|Describí cómo decidís qué dejar afuera de un reporte para no saturar a quien lo lee.'],
      ['Orientación a Resultados y Mejora Continua', 'Enfocar el análisis en decisiones e impacto concreto, e identificar oportunidades de optimización de procesos. Evaluar evidencia de resultados medibles y mejoras implementadas.',
        'Contame de un análisis tuyo que derivó en una decisión concreta con impacto económico u operativo. ¿Cuál fue el resultado?|Dame un ejemplo de un proceso de reporte que automatizaste o mejoraste. ¿Cuánto tiempo o error ahorraste?|Describí una ineficiencia que detectaste en los datos operativos y la recomendación que propusiste.'],
      ['Organización y Planificación', 'Gestionar múltiples requerimientos de información y proyectos de datos en paralelo, cumpliendo plazos. Evaluar evidencia de priorización y método.',
        'Contame de un período con varios pedidos de análisis urgentes a la vez. ¿Cómo priorizaste y cumpliste?|Dame un ejemplo de un proyecto de datos largo que planificaste de punta a punta. ¿Cómo lo organizaste?|Describí cómo llevás el control de tus tareas y fuentes de datos para no perder trazabilidad.']
    ],
    'analista-compras': [
      ['Orientación al cliente', 'Responder a las necesidades de las áreas internas (operaciones, mantenimiento, proyectos) asegurando abastecimiento oportuno. Evaluar evidencia de servicio al cliente interno y cumplimiento de plazos.',
        'Contame de un pedido urgente de obra o mantenimiento que tuviste que abastecer contra reloj. ¿Cómo lo resolviste?|Dame un ejemplo de un área interna disconforme con una compra o un plazo. ¿Cómo recompusiste la relación?|Describí cómo equilibrás la presión de las áreas por tener el material ya contra el proceso de compra correcto.'],
      ['Pensamiento analítico y atención al detalle', 'Analizar cotizaciones, cuadros comparativos y condiciones para decidir la mejor compra, sin errores en cantidades, precios o especificaciones. Evaluar evidencia de rigor en el análisis de compra.',
        'Contame de una decisión de compra compleja donde el más barato no era la mejor opción. ¿Cómo armaste el cuadro comparativo y qué decidiste?|Dame un ejemplo de un error de especificación o cantidad que detectaste antes de emitir la OC. ¿Qué pasó?|Describí cómo analizás una cotización para asegurarte de que comparás peras con peras.'],
      ['Negociación', 'Obtener mejores condiciones de precio, plazo y calidad con proveedores, sosteniendo relaciones de largo plazo. Evaluar evidencia de ahorros y acuerdos logrados.',
        'Contame de una negociación con un proveedor donde lograste una mejora concreta de precio o condiciones. ¿Cómo la encaraste?|Dame un ejemplo de un proveedor crítico o monopólico con el que tuviste que negociar. ¿Qué estrategia usaste?|Describí una situación donde tuviste que sostener una relación de largo plazo a pesar de un conflicto puntual.'],
      ['Organización y planificación', 'Gestionar el ciclo de compras de punta a punta y múltiples órdenes en paralelo cumpliendo plazos. Evaluar evidencia de método y seguimiento.',
        'Contame cómo organizás el seguimiento de muchas órdenes de compra abiertas a la vez. ¿Qué sistema o método usás?|Dame un ejemplo de una compra de importación que planificaste (despacho, pagos, plazos). ¿Cómo la coordinaste?|Describí una vez que un quiebre de stock te obligó a reorganizar prioridades de compra. ¿Cómo lo manejaste?'],
      ['Trabajo en equipo', 'Articular con operaciones, mantenimiento, administración y finanzas para cerrar el circuito de compra-recepción-pago. Evaluar evidencia de colaboración transversal.',
        'Contame de una compra que requirió coordinar con varias áreas (operaciones, recepción en obra, finanzas). ¿Cómo articulaste?|Dame un ejemplo de un desacuerdo con otra área sobre una compra o una factura. ¿Cómo lo resolviste?|Describí cómo trabajás con recepción y administración para validar que lo comprado llegó y se pagó bien.']
    ],
    'analista-sgi': [
      ['Organización y control', 'Mantener el sistema documental y los registros del SGI ordenados, vigentes y auditables (procedimientos, instructivos, no conformidades). Evaluar evidencia de control documental riguroso.',
        'Contame de cómo organizaste o saneaste la documentación de un sistema de gestión. ¿En qué estado estaba y cómo lo dejaste?|Dame un ejemplo de un descontrol documental (versiones vencidas, registros faltantes) que tuviste que corregir. ¿Cómo lo hiciste?|Describí tu método para asegurar que los procedimientos estén siempre actualizados y disponibles.'],
      ['Análisis y resolución de problemas', 'Identificar causas raíz de no conformidades y desvíos, y proponer acciones correctivas efectivas. Evaluar evidencia de análisis de causa y mejora.',
        'Contame de una no conformidad importante que tuviste que gestionar. ¿Cómo analizaste la causa raíz y qué acción correctiva implementaste?|Dame un ejemplo de un hallazgo de auditoría recurrente que lograste eliminar de fondo. ¿Cómo?|Describí una situación donde un problema de calidad o seguridad requería una solución rápida pero también sostenible. ¿Cómo decidiste?'],
      ['Comunicación efectiva', 'Transmitir requisitos normativos y resultados de auditoría a áreas operativas y dirección, logrando comprensión y acción. Evaluar evidencia de comunicación clara de lo normativo.',
        'Contame de una vez que tuviste que explicar un requisito normativo complejo a un área operativa que lo veía como burocracia. ¿Cómo lograste que lo adoptaran?|Dame un ejemplo de cómo presentaste resultados de auditoría a la dirección. ¿Qué priorizaste?|Describí una situación donde una mala comunicación de un procedimiento generó un incumplimiento. ¿Cómo la corregiste?'],
      ['Impacto e influencia', 'Lograr que las áreas operativas cumplan los lineamientos del SGI sin autoridad jerárquica directa. Evaluar evidencia de instalación de cultura de gestión.',
        'Contame de cómo lograste que un área resistente adoptara una práctica del sistema de gestión. ¿Qué hiciste?|Dame un ejemplo donde tu intervención mejoró el cumplimiento normativo de un sector. ¿Cómo influiste?|Describí una situación donde tuviste que sostener un estándar de calidad o seguridad frente a la presión de la operación por avanzar.'],
      ['Trabajo en equipo', 'Articular con áreas operativas y de soporte para implementar el SGI de forma colaborativa. Evaluar evidencia de articulación transversal.',
        'Contame de una auditoría o implementación que requirió coordinar con varias áreas. ¿Cómo articulaste?|Dame un ejemplo donde dependías de otras áreas para cerrar una acción correctiva. ¿Cómo aseguraste que cumplieran?|Describí una situación de tensión con un área operativa por temas del sistema de gestión. ¿Cómo la resolviste?']
    ]
  };

  let escritas = 0;
  for (let i = 1; i < data.length; i++) {
    const perfil = data[i][0];
    const idComp = data[i][1];
    if (!DATOS[perfil]) continue;
    const idx = parseInt(String(idComp).replace('c', ''), 10) - 1;
    const fila = DATOS[perfil][idx];
    if (!fila) continue;
    sh.getRange(i + 1, 3).setValue(fila[0]); // nombre_comp
    sh.getRange(i + 1, 4).setValue(fila[1]); // descripcion
    sh.getRange(i + 1, 6).setValue(fila[2]); // preguntas
    escritas++;
  }
  Logger.log('✅ Competencias actualizadas: ' + escritas + ' filas (coord-rrhh + 3 analistas)');
}
function cargarChecksScreening_MACS() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const sv = ss.getSheetByName('Perfiles_Screening_Val');

  const CHECKS = {
    'coord-rrhh': [
      ['v6', 'Gestionó personal de convenio', 'Sí|No'],
      ['v7', 'Experiencia gremial del lado de la empresa', 'Sí|No'],
      ['v8', 'Salidas laborales justificadas (sin job hopping)', 'Sí|No|A revisar']
    ],
    'analista-datos': [
      ['v6', 'Usa IA generativa para tareas reales (SQL/fórmulas/DAX)', 'Sí|No'],
      ['v7', 'Armó dashboards de punta a punta', 'Sí|No'],
      ['v8', 'Experiencia en análisis económico-financiero de proyectos', 'Sí|No']
    ],
    'analista-compras': [
      ['v6', '5+ años en compras con ERP', 'Sí|No'],
      ['v7', 'Emitía órdenes de compra', 'Sí|No'],
      ['v8', 'Experiencia en importación / comex', 'Sí|No'],
      ['v9', 'Referencias verificables', 'Sí|No|A revisar']
    ],
    'analista-sgi': [
      ['v6', 'ISO 9001/45001/14001 con experiencia real', 'Sí|No'],
      ['v7', 'Implementó SGI (no solo teoría)', 'Sí|No'],
      ['v8', 'Auditorías internas/externas', 'Sí|No']
    ]
  };

  // checks comunes a los 4 perfiles (excluyentes transversales)
  const COMUNES = [
    ['Excel avanzado', 'Sí|No'],
    ['Manejo de IA', 'Sí|No']
  ];
  Object.keys(CHECKS).forEach(pid => {
    let n = CHECKS[pid].length;
    COMUNES.forEach(c => { n++; CHECKS[pid].push(['v' + (n), c[0], c[1]]); });
  });

  const data = sv.getDataRange().getValues();
  const existe = {};
  for (let i = 1; i < data.length; i++) existe[data[i][0] + '|' + data[i][1]] = true;

  const nuevas = [];
  Object.keys(CHECKS).forEach(pid => {
    CHECKS[pid].forEach(c => {
      if (!existe[pid + '|' + c[0]]) nuevas.push([pid, c[0], c[1], c[2]]);
    });
  });

  if (nuevas.length === 0) { Logger.log('Sin cambios — los checks ya estaban cargados.'); return; }
  sv.getRange(sv.getLastRow() + 1, 1, nuevas.length, 4).setValues(nuevas);
  Logger.log('✅ ' + nuevas.length + ' checks agregados a Screening_Val.');
}
function actualizarSGI_MACS() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const PERFIL = 'analista-sgi';

  const promptExtra =
`PERFIL: Analista de Sistema de Gestión de Calidad, Seguridad, Medio Ambiente y Salud (SGI). Cliente: MACS S.A., planta en Pilar (Bs. As.). Reporta a Directora de Administración y Gestión.
MISIÓN: Asegurar la correcta implementación, mantenimiento y mejora del Sistema de Gestión Integrado (ISO 9001, ISO 45001, ISO 14001), garantizando cumplimiento legal, normativo y contractual. Sostener estándares de calidad, seguridad y medio ambiente con foco en el cuidado de las personas.

EXPERIENCIA (EXCLUYENTE): mínimo 5 años en gestión e implementación de sistemas de calidad, seguridad y medio ambiente. Preferente en proyectos técnicos e industriales (construcción, ingeniería, industrial, químico, minería).
FORMACIÓN: Graduado/a en Ingeniería o Licenciatura (Industrial, en Seguridad, Ambiental o afín) — no excluyente. Cursos/diplomatura/certificaciones en normas ISO (9001/45001/14001) — EXCLUYENTE.
CONOCIMIENTOS CLAVE:
- Normas ISO 9001, ISO 45001, ISO 14001 con experiencia real (no solo teórica).
- Auditorías internas y externas (planifica, coordina y participa); gestión de hallazgos y no conformidades.
- Administración y control documental del SGI (procedimientos, instructivos, registros).
- Cumplimiento de requisitos legales y normativos (calidad, seguridad, MA). Reportes e indicadores / tablero de control.
- Auditor interno ISO 9000 (capacitación deseable).
- Excel AVANZADO. ERP (SAP, Oracle, Odoo o similar) intermedio. Software de gestión documental. Conocimientos de IA.
ZONA: planta en Pilar (Bs. As.). Modalidad presencial. Valorar cercanía / disponibilidad.

BANDERAS ROJAS (DESCARTANTES): el cliente busca un perfil de ANALISTA, NO de auditor puro ni de supervisor de seguridad. Descartar a candidatos cuyo perfil sea exclusivamente auditoría o exclusivamente seguridad e higiene operativa. Descartar perfiles provenientes EXCLUSIVAMENTE de áreas administrativas sin contacto con personal operativo.

CONTINUIDAD LABORAL (criterio importante para Blumia): analizá el historial. VALORÁ permanencia y crecimiento dentro de cada empresa. PENALIZÁ rotación alta con saltos cortos y frecuentes (job hopping) SIN justificación razonable (cierres, fin de obra/proyecto, mudanzas). La estabilidad y la progresión interna suman fuerte.

EVALUACIÓN DE BORDERLINE — "incluir y señalar": NO descartes mecánicamente a quien no cumple el 100%. Si a un candidato le falta algún requisito pero presenta FACTORES COMPENSATORIOS (trayectoria sólida, continuidad laboral ejemplar, formación destacada, experiencia transferible, logros relevantes), clasificalo como MEDIO y en el resumen marcá explícitamente "REVISAR:" explicando QUÉ le falta y QUÉ lo compensa, para que el selector decida. Ante la duda, INCLUIR y señalar, no filtrar. Reservá BAJO para quienes claramente no tienen experiencia real en sistemas de gestión, no llegan a la base mínima, son perfil de auditor/supervisor de seguridad, o caen en banderas rojas claras.

NIVEL_AJUSTE:
- Alto: cumple experiencia excluyente (5+ años SGI) + ISO 9001/45001/14001 con experiencia real + control documental + Excel avanzado, con buena continuidad laboral. Suma Ingeniería y experiencia industrial.
- Medio: cumple lo básico pero le falta algo (una norma, menos años, o un deseable) O es un borderline con factores compensatorios → marcar "REVISAR:" en el resumen.
- Bajo: no llega a 5 años, sin ISO, perfil de auditor puro o supervisor de seguridad, sin experiencia real en SGI, o banderas rojas.`;

  // Config.prompt_extra (col 9)
  const config = ss.getSheetByName('Config');
  const cfg = config.getDataRange().getValues();
  for (let i = 1; i < cfg.length; i++) if (cfg[i][0] === PERFIL) { config.getRange(i + 1, 9).setValue(promptExtra); break; }
  Logger.log('✅ Config.prompt_extra SGI actualizado (5 años + red flag analista)');

  // Perfiles_Hero.posicion (col 3) — lo que ve el candidato
  const hero = ss.getSheetByName('Perfiles_Hero');
  const h = hero.getDataRange().getValues();
  for (let i = 1; i < h.length; i++) if (h[i][0] === PERFIL) { hero.getRange(i + 1, 3).setValue('Analista de Sistema de Gestión (Calidad, Seguridad, Medio Ambiente y Salud)'); break; }
  Logger.log('✅ Perfiles_Hero.posicion SGI actualizado');

  // Screening_Eval e1 → 5 años (col 3)
  const se = ss.getSheetByName('Perfiles_Screening_Eval');
  const s = se.getDataRange().getValues();
  for (let i = 1; i < s.length; i++) if (s[i][0] === PERFIL && s[i][1] === 'e1') { se.getRange(i + 1, 3).setValue('Experiencia 5+ años en sistemas de gestión (QSMA)'); break; }
  Logger.log('✅ Screening_Eval e1 SGI → 5 años');

  // Competencia c4 nombre → "Impacto e Influencia: gestión de cambio" (col 3)
  const pc = ss.getSheetByName('Perfiles_Competencias');
  const p = pc.getDataRange().getValues();
  for (let i = 1; i < p.length; i++) if (p[i][0] === PERFIL && p[i][1] === 'c4') { pc.getRange(i + 1, 3).setValue('Impacto e Influencia: gestión de cambio'); break; }
  Logger.log('✅ Competencia c4 SGI renombrada');

  // Check nuevo en Screening_Val: perfil analista (no auditor / no supervisor seg.) — idempotente
  const sv = ss.getSheetByName('Perfiles_Screening_Val');
  const v = sv.getDataRange().getValues();
  let existe = false;
  for (let i = 1; i < v.length; i++) if (v[i][0] === PERFIL && v[i][1] === 'v9') { existe = true; break; }
  if (!existe) {
    sv.getRange(sv.getLastRow() + 1, 1, 1, 4).setValues([[PERFIL, 'v9', 'Perfil de analista (no auditor puro / no supervisor de seguridad)', 'Sí|No|A revisar']]);
    Logger.log('✅ Check v9 SGI agregado');
  } else Logger.log('Check v9 SGI ya existía');

  Logger.log('LISTO — SGI actualizado.');
}
function agregarExcelIA_MACS() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const sv = ss.getSheetByName('Perfiles_Screening_Val');
  const PERFILES = ['coord-rrhh', 'analista-datos', 'analista-compras', 'analista-sgi'];
  const COMUNES = [
    ['Excel avanzado', 'Sí|No'],
    ['Manejo de IA', 'Sí|No']
  ];

  const data = sv.getDataRange().getValues();

  // próximo vN libre y etiquetas existentes por perfil
  const maxV = {}, etiquetas = {};
  PERFILES.forEach(p => { maxV[p] = 0; etiquetas[p] = {}; });
  for (let i = 1; i < data.length; i++) {
    const pid = data[i][0];
    if (maxV[pid] === undefined) continue;
    const m = String(data[i][1]).match(/^v(\d+)$/);
    if (m) maxV[pid] = Math.max(maxV[pid], parseInt(m[1]));
    etiquetas[pid][data[i][2]] = true;
  }

  const nuevas = [];
  PERFILES.forEach(p => {
    COMUNES.forEach(c => {
      if (!etiquetas[p][c[0]]) {
        maxV[p]++;
        nuevas.push([p, 'v' + maxV[p], c[0], c[1]]);
      }
    });
  });

  if (nuevas.length === 0) { Logger.log('Sin cambios — Excel/IA ya estaban.'); return; }
  sv.getRange(sv.getLastRow() + 1, 1, nuevas.length, 4).setValues(nuevas);
  Logger.log('✅ ' + nuevas.length + ' filas agregadas:');
  nuevas.forEach(n => Logger.log('  ' + n[0] + ' ' + n[1] + ' — ' + n[2]));
}
function unificarScreeningVal_MACS() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const sv = ss.getSheetByName('Perfiles_Screening_Val');
  const MACS = ['coord-rrhh', 'analista-datos', 'analista-compras', 'analista-sgi'];

  const ESC = 'Avanzado|Intermedio|Básico';

  // bloque limpio y ordenado por perfil (contexto → red flag/excluyente → comunes)
  const BLOQUE = {
    'coord-rrhh': [
      ['Remuneración pretendida compatible', 'Sí|No|A revisar'],
      ['Disponibilidad horaria', 'Sí|No'],
      ['Ubicación / movilidad', 'Sí|No'],
      ['Preaviso razonable', 'Sí|No'],
      ['Interés real en la posición', 'Alto|Medio|Bajo'],
      ['Gestionó personal de convenio', 'Sí|No'],
      ['Experiencia gremial del lado de la empresa', 'Sí|No'],
      ['Salidas laborales justificadas (sin job hopping)', 'Sí|No|A revisar'],
      ['Excel avanzado', ESC],
      ['Manejo de IA', ESC]
    ],
    'analista-datos': [
      ['Remuneración pretendida compatible', 'Sí|No|A revisar'],
      ['Disponibilidad horaria', 'Sí|No'],
      ['Ubicación / movilidad', 'Sí|No'],
      ['Preaviso razonable', 'Sí|No'],
      ['Interés real en la posición', 'Alto|Medio|Bajo'],
      ['Usa IA generativa para tareas reales (SQL/fórmulas/DAX)', 'Sí|No'],
      ['Armó dashboards de punta a punta', 'Sí|No'],
      ['Experiencia en análisis económico-financiero de proyectos', 'Sí|No'],
      ['Excel avanzado', ESC],
      ['Manejo de IA', ESC]
    ],
    'analista-compras': [
      ['Remuneración pretendida compatible', 'Sí|No|A revisar'],
      ['Disponibilidad horaria', 'Sí|No'],
      ['Ubicación / movilidad', 'Sí|No'],
      ['Preaviso razonable', 'Sí|No'],
      ['Interés real en la posición', 'Alto|Medio|Bajo'],
      ['5+ años en compras con ERP', 'Sí|No'],
      ['Emitía órdenes de compra', 'Sí|No'],
      ['Experiencia en importación / comex', 'Sí|No'],
      ['Referencias verificables', 'Sí|No|A revisar'],
      ['Excel avanzado', ESC],
      ['Manejo de IA', ESC]
    ],
    'analista-sgi': [
      ['Remuneración pretendida compatible', 'Sí|No|A revisar'],
      ['Disponibilidad horaria', 'Sí|No'],
      ['Ubicación / movilidad', 'Sí|No'],
      ['Preaviso razonable', 'Sí|No'],
      ['Interés real en la posición', 'Alto|Medio|Bajo'],
      ['ISO 9001/45001/14001 con experiencia real', 'Sí|No'],
      ['Implementó SGI (no solo teoría)', 'Sí|No'],
      ['Auditorías internas/externas', 'Sí|No'],
      ['Perfil de analista (no auditor puro / no supervisor de seguridad)', 'Sí|No|A revisar'],
      ['Excel avanzado', ESC],
      ['Manejo de IA', ESC]
    ]
  };

  const all = sv.getDataRange().getValues();
  // conservar header + filas de perfiles NO-MACS (ej: operaciones)
  const keep = all.slice(1).filter(r => r[0] !== '' && MACS.indexOf(r[0]) === -1);

  // armar filas MACS renumeradas v1..vN
  const macsRows = [];
  MACS.forEach(p => {
    BLOQUE[p].forEach((c, idx) => macsRows.push([p, 'v' + (idx + 1), c[0], c[1]]));
  });

  const final = keep.concat(macsRows);

  // limpiar de fila 2 en adelante y reescribir
  if (sv.getLastRow() > 1) sv.getRange(2, 1, sv.getLastRow() - 1, 4).clearContent();
  sv.getRange(2, 1, final.length, 4).setValues(final);

  Logger.log('✅ Screening_Val unificado: ' + macsRows.length + ' filas MACS + ' + keep.length + ' conservadas (operaciones u otros).');
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
  const out = [];
  for (let i = 1; i < cfg.length; i++) {
    const id = cfg[i][0];
    if (!id || id === 'operaciones') continue;
    out.push({ id: id, nombre: cfg[i][1], prompt: cfg[i][8] || '' });
  }
  return out;
}

// --- tabla Mail_Routing: atajo opcional por asunto (la editás vos) ---
function setupMailRouting() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let sh = ss.getSheetByName('Mail_Routing');
  if (!sh) sh = ss.insertSheet('Mail_Routing');
  sh.clear();
  const filas = [
    ['perfil_id', 'palabras_clave_asunto (separadas por |)'],
    ['coord-rrhh', 'rrhh|recursos humanos|gestión de personas|coordinador rrhh|personas'],
    ['analista-datos', 'datos|data|analista de datos|bi|power bi|sql'],
    ['analista-compras', 'compras|contrataciones|abastecimiento|comprador'],
    ['analista-sgi', 'sgi|sistema de gestión|iso|calidad|seguridad|medio ambiente|auditor']
  ];
  sh.getRange(1, 1, filas.length, 2).setValues(filas);
  Logger.log('✅ Mail_Routing creada. Editá las palabras clave a gusto.');
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
    if (claves.some(k => a.indexOf(k) !== -1)) return pid;
  }
  return null;
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
function patchAntiDescarte_MACS() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const config = ss.getSheetByName('Config');
  const data = config.getDataRange().getValues();
  const PERFILES = ['coord-rrhh', 'analista-datos', 'analista-compras', 'analista-sgi'];

  const BLOQUE =
`

REGLAS ANTI-DESCARTE (CRÍTICAS, sobreescriben cualquier criterio anterior):
1. La falta de experiencia en la industria, el rubro o el contexto específico (convenio, sector, tipo de empresa) NUNCA es por sí sola motivo de C. Si el candidato tiene experiencia REAL y sólida en la función del rol, esa falta baja la nota a B (en reserva) con nota "REVISAR: falta [X], pero compensa con [Y]", para que el selector decida. Reservá C SOLO para: no tener experiencia real en la función del puesto, no llegar a la base mínima excluyente de la función, o banderas rojas claras.
2. SOBRECALIFICACIÓN: si el candidato EXCEDE el perfil (es claramente más senior o de mayor alcance que el puesto), NO lo descartes ni lo subas a A. Asignale B (en reserva) con nota "REVISAR: perfil por encima del puesto, evaluar encaje real y expectativa salarial". Un perfil que excede es B, nunca B+ ni A.`;

  let n = 0;
  for (let i = 1; i < data.length; i++) {
    const id = data[i][0];
    if (PERFILES.indexOf(id) === -1) continue;
    let prompt = String(data[i][8] || '');
    if (prompt.indexOf('REGLAS ANTI-DESCARTE') !== -1) continue; // idempotente
    config.getRange(i + 1, 9).setValue(prompt + BLOQUE);
    n++;
    Logger.log('✅ Patch aplicado: ' + id);
  }
  Logger.log(n === 0 ? 'Sin cambios — ya estaba aplicado.' : 'LISTO — ' + n + ' perfiles parcheados.');
}
function probarError() {
  const payloadFalso = {
    perfil: 'coord-rrhh',
    datos: { nombre: 'PRUEBA Error', email: 'prueba@test.com', telefono: '3364000000' }
  };
  registrarErrorPostulacion(payloadFalso, new Error('Esto es una prueba de error, ignorar'));
  Logger.log('Listo: revisá tu mail y la hoja Errores');
}
function probarMail() {
  Logger.log('CORREO_DESTINO = ' + CORREO_DESTINO);
  try {
    MailApp.sendEmail(CORREO_DESTINO, 'Test Blumia', 'Esto es una prueba de envío.');
    Logger.log('Mail enviado OK');
  } catch(e) {
    Logger.log('ERROR al enviar: ' + e);
  }
  Logger.log('Cuota restante hoy: ' + MailApp.getRemainingDailyQuota());
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
        <tr><td class="k">Disponibilidad full-time</td><td class="v">${d.fulltime || '-'}</td></tr>
        <tr><td class="k">Trabajo 100% presencial</td><td class="v">${d.presencial || '-'}</td></tr>
        <tr><td class="k">Remuneración pretendida</td><td class="v">${d.remuneracion || '-'}</td></tr>
      </table>
      <div class="sec">Formación</div>
      <table>
        <tr><td class="k">Formación</td><td class="v">${formacion || '-'}</td></tr>
        <tr><td class="k">Posgrado</td><td class="v">${posgrado}</td></tr>
      </table>
      <div class="sec">Conocimientos</div>
      <table>
        <tr><td class="k">Excel</td><td class="v">${d.excel || '-'}</td></tr>
        <tr><td class="k">Power BI</td><td class="v">${d.powerbi || '-'}</td></tr>
        <tr><td class="k">Inglés</td><td class="v">${d.ingles || '-'}</td></tr>
        <tr><td class="k">Herramientas de IA</td><td class="v">${ia || '-'}</td></tr>
      </table>
      ${dyn ? '<div class="sec">Sobre esta búsqueda</div><table>' + dyn + '</table>' : ''}
    </div>
  </body></html>`;

  const folderId = cfg.folder_pdf_postulacion || cfg.folder_recibidos;
  const blob = Utilities.newBlob(html, 'text/html', 'Postulacion_' + (d.nombre || 'candidato').replace(/\s+/g, '_') + '.html');
  const pdf = DriveApp.getFolderById(folderId).createFile(blob.getAs('application/pdf'));
  return pdf.getUrl();
}
function headerPostulacionPDF() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  listarPerfiles().forEach(p => {
    let cfg; try { cfg = getConfig(p.id); } catch(e){ return; }
    const sh = ss.getSheetByName(cfg.hoja_nombre);
    if (sh) sh.getRange(1,31).setValue('Postulación').setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
  });
}
function debugCalendar() {
  const cal = CalendarApp.getDefaultCalendar();
  const ini = new Date();
  const fin = new Date(ini.getTime() + 15 * 24 * 60 * 60 * 1000);
  const eventos = cal.getEvents(ini, fin);
  const norm = s => String(s||'').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'');

  const cfg = getConfig('coord-rrhh');
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(cfg.hoja_nombre);
  const datos = sh.getDataRange().getValues();

  eventos.forEach(ev => {
    const titulo = norm(ev.getTitle());
    Logger.log('EVENTO: [' + titulo + ']');
    Logger.log('  tiene screening?: ' + (titulo.indexOf('screening') > -1));
    for (let i = 1; i < datos.length; i++) {
      const nombre = norm(datos[i][1]);
      if (!nombre) continue;
      if (titulo.indexOf(nombre) > -1) {
        Logger.log('  >>> MATCH con fila ' + (i+1) + ': [' + nombre + ']');
      }
    }
  });
}
function debugCalendar2() {
  const cal = CalendarApp.getDefaultCalendar();
  const ini = new Date();
  const fin = new Date(ini.getTime() + 15 * 24 * 60 * 60 * 1000);
  const eventos = cal.getEvents(ini, fin);
  const norm = s => String(s||'').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'');

  const cfg = getConfig('coord-rrhh');
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(cfg.hoja_nombre);
  const datos = sh.getDataRange().getValues();

  let salida = '';
  eventos.forEach(ev => {
    const titulo = norm(ev.getTitle());
    salida += 'EVENTO:[' + titulo + '] screening?:' + (titulo.indexOf('screening')>-1) + ' | ';
    for (let i = 1; i < datos.length; i++) {
      const nombre = norm(datos[i][1]);
      if (nombre && titulo.indexOf(nombre) > -1) salida += 'MATCH fila' + (i+1) + ':[' + nombre + '] ';
    }
  });
  // escribe el resultado en Config!Z1 para que lo leas
  SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getRange('Z1').setValue(salida || 'sin eventos');
}
function verCalendarios() {
  const cals = CalendarApp.getAllCalendars();
  let s = 'DEFAULT: [' + CalendarApp.getDefaultCalendar().getName() + '] || TODOS: ';
  cals.forEach(c => { s += '[' + c.getName() + '] '; });
  SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getRange('Z1').setValue(s);
}
function debugCalendar3() {
  const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
  const ini = new Date();
  const fin = new Date(ini.getTime() + 15 * 24 * 60 * 60 * 1000);
  const eventos = cal.getEvents(ini, fin);
  let s = 'Eventos: ' + eventos.length + ' | ';
  eventos.forEach(ev => { s += '[' + ev.getTitle() + '] '; });
  SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getRange('Z1').setValue(s);
}
function debugMatch() {
  const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
  const ini = new Date();
  const fin = new Date(ini.getTime() + 15 * 24 * 60 * 60 * 1000);
  const eventos = cal.getEvents(ini, fin);
  const norm = s => String(s||'').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'');

  const cfg = getConfig('analista-compras');
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(cfg.hoja_nombre);
  const datos = sh.getDataRange().getValues();

  let s = '';
  eventos.forEach(ev => {
    const t = norm(ev.getTitle());
    s += 'TITULO:[' + t + '] | ';
    for (let i = 1; i < datos.length; i++) {
      const n = norm(datos[i][1]);
      if (!n) continue;
      if (t.indexOf(n) > -1) s += 'MATCH:[' + n + '] ';
    }
  });
  s += ' || NOMBRES EN HOJA: ';
  for (let i = 1; i < Math.min(datos.length, 10); i++) {
    if (datos[i][1]) s += '[' + norm(datos[i][1]) + '] ';
  }
  SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getRange('Z1').setValue(s);
}
function debugSync() {
  const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
  const ini = new Date();
  const fin = new Date(ini.getTime() + 15 * 24 * 60 * 60 * 1000);
  const eventos = cal.getEvents(ini, fin);
  const norm = s => String(s||'').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'');

  let s = 'Eventos:' + eventos.length + ' | ';

  ['analista-compras','coord-rrhh','analista-datos','analista-sgi'].forEach(pid => {
    let cfg; try { cfg = getConfig(pid); } catch(e){ return; }
    const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(cfg.hoja_nombre);
    if (!sh) { s += pid + ':SIN HOJA | '; return; }
    const datos = sh.getDataRange().getValues();

    eventos.forEach(ev => {
      const t = norm(ev.getTitle());
      let est = null;
      if (t.indexOf('screening')>-1) est='Screening agendado';
      else if (t.indexOf('competencia')>-1) est='Competencias agendada';
      else if (t.indexOf('empresa')>-1) est='Entrevista empresa agendada';

      for (let i = 1; i < datos.length; i++) {
        const n = norm(datos[i][1]);
        if (n && t.indexOf(n) > -1) {
          s += 'MATCH ' + pid + ' fila' + (i+1) + ' [' + n + '] estado-a-poner:' + est + ' | ';
        }
      }
    });
  });
  SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getRange('Z1').setValue(s);
}
function debugNombres() {
  const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
  const ini = new Date();
  const fin = new Date(ini.getTime() + 15 * 24 * 60 * 60 * 1000);
  const eventos = cal.getEvents(ini, fin);

  let s = 'EVENTOS: ';
  eventos.forEach(ev => { s += '[' + ev.getTitle() + '] '; });

  s += ' || COMPRAS: ';
  const cfg = getConfig('analista-compras');
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(cfg.hoja_nombre);
  const datos = sh.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1]) s += '[' + datos[i][1] + '] ';
  }
  SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getRange('Z1').setValue(s);
}
function debugEduardo() {
  const norm = s => String(s||'').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
  const titulo = norm('MACS - Screening Analista de compras y contrataciones (Eduardo Bodean)');
  const cfg = getConfig('analista-compras');
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(cfg.hoja_nombre);
  const datos = sh.getDataRange().getValues();

  let s = 'titulo norm:[' + titulo + '] | ';
  for (let i = 1; i < datos.length; i++) {
    const n = norm(datos[i][1]);
    if (n === norm('Eduardo Bodean')) {
      s += 'fila' + (i+1) + ' nombre:[' + n + '] incluido?:' + (titulo.indexOf(n) > -1);
    }
  }
  SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getRange('Z1').setValue(s);
}
function debugSyncReal() {
  let s = '';
  try {
    const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
    const ini = new Date();
    const fin = new Date(ini.getTime() + 15 * 24 * 60 * 60 * 1000);
    const eventos = cal.getEvents(ini, fin);
    const norm = x => String(x||'').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'');
    const cfg = getConfig('analista-compras');
    const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(cfg.hoja_nombre);
    const datos = sh.getDataRange().getValues();

    eventos.forEach(ev => {
      const t = norm(ev.getTitle());
      if (t.indexOf('screening') === -1) return;
      for (let i = 1; i < datos.length; i++) {
        const n = norm(datos[i][1]);
        if (n && t.indexOf(n) > -1) {
          const celda = sh.getRange(i+1, 9);
          celda.setValue('Screening agendado').setBackground('#D1E9FF');
          s += 'ESCRITO fila' + (i+1) + ' [' + n + '] | ';
        }
      }
    });
    s = s || 'no entro a ningun match';
  } catch(e) {
    s = 'ERROR: ' + e;
  }
  SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getRange('Z1').setValue(s);
}
function debugTituloReal() {
  const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
  const ini = new Date();
  const fin = new Date(ini.getTime() + 15 * 24 * 60 * 60 * 1000);
  const eventos = cal.getEvents(ini, fin);
  const norm = x => String(x||'').toLowerCase().trim().normalize('NFD').replace(/[\u0300-\u036f]/g,'');

  let s = '';
  eventos.forEach((ev, idx) => {
    const t = ev.getTitle();
    const tn = norm(t);
    s += 'EV' + idx + ' crudo:[' + t + '] norm:[' + tn + '] screening?:' + (tn.indexOf('screening')>-1) + ' bodean?:' + (tn.indexOf('eduardo bodean')>-1) + ' || ';
  });
  SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getRange('Z1').setValue(s);
}
/* ============================================================
   WHATSAPP con mensaje + link de agenda de screening por perfil.
   Reemplaza la fórmula simple anterior de la col 21.
   Correr: ponerWhatsAppAgenda()  (aplica a las 4 búsquedas activas)
   ============================================================ */

/* WhatsApp con mensaje + link de agenda — versión EN INGLÉS (Sheet en inglés) */

function ponerWhatsAppAgenda() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);

  const MAP = {
    'analista-compras': { link: 'https://calendar.app.google/JxPyBZGrP1i4FhKY9', pos: 'Analista de Compras y Contrataciones' },
    'coord-rrhh':       { link: 'https://calendar.app.google/WrYYgPKMCVbAS56Q7', pos: 'Coordinador/a de Gestión de Personas' },
    'analista-sgi':     { link: 'https://calendar.app.google/7Y3g2QVh4Uuhzojh6', pos: 'Analista de SGI / Calidad' },
    'analista-datos':   { link: 'https://calendar.app.google/R3XSVkLUpRJVyFaX6', pos: 'Analista de Datos' }
  };

  Object.keys(MAP).forEach(pid => {
    let cfg; try { cfg = getConfig(pid); } catch(e){ return; }
    const sh = ss.getSheetByName(cfg.hoja_nombre);
    if (!sh) return;
    const link = MAP[pid].link;
    const pos = MAP[pid].pos;

    for (let i = 2; i <= 250; i++) {
      const f =
        '=IF(C' + i + '<>"";HYPERLINK("https://api.whatsapp.com/send?phone="&' +
        'SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(C' + i + ';"+";"");" ";"");"-";"");"(";"");")";"")&' +
        '"&text="&ENCODEURL("Hola "&IFERROR(LEFT(B' + i + ';FIND(" ";B' + i + ')-1);B' + i + ')&' +
        '", ¿cómo estás? Soy "&N' + i + '&" de Blumia Consulting."&CHAR(10)&CHAR(10)&' +
        '"Te contacto por la búsqueda de ' + pos + ' para la empresa MACS."&CHAR(10)&CHAR(10)&' +
        '"Me gustaría coordinar una charla breve con vos. Si te interesa, elegí el horario que mejor te quede en este link y ya nos queda agendado: ' + link + '"&CHAR(10)&CHAR(10)&' +
        '"Saludos!");"WHATSAPP ▶");"")';
      sh.getRange(i, 21).setFormula(f);
    }
    Logger.log('WhatsApp+agenda OK: ' + cfg.hoja_nombre);
  });
  Logger.log('Listo.');
}
/* EXPORTA candidatos contactados a una PESTAÑA "Contactos" del Sheet */

function exportarContactos() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);

  const ETIQUETA = {
    'coord-rrhh': 'RRHH',
    'analista-datos': 'Analista Datos',
    'analista-compras': 'Analista Compras',
    'analista-sgi': 'Analista SGI'
  };
  const ESTADOS_OK = ['Contactado','Screening agendado','Avanza a competencias',
    'Competencias agendada','Avanza a empresa','Entrevista empresa agendada','En cliente','Contratado'];

  const filas = [['Name','Given Name','Phone 1 - Value','E-mail 1 - Value','Group Membership']];

  Object.keys(ETIQUETA).forEach(pid => {
    let cfg; try { cfg = getConfig(pid); } catch(e){ return; }
    const sh = ss.getSheetByName(cfg.hoja_nombre);
    if (!sh) return;
    const datos = sh.getDataRange().getValues();
    const etq = ETIQUETA[pid];

    for (let i = 1; i < datos.length; i++) {
      const nombre = String(datos[i][1] || '').trim();
      const tel = String(datos[i][2] || '').trim();
      const estado = String(datos[i][8] || '').trim();
      const email = String(datos[i][21] || '').trim();
      if (!nombre || !tel) continue;
      if (ESTADOS_OK.indexOf(estado) === -1) continue;

      const nombreCompleto = etq + ' - ' + nombre;
      const telLimpio = tel.replace(/[^0-9+]/g, '');
      const emailOk = (email && email.indexOf('@') > -1) ? email : '';
      filas.push([nombreCompleto, nombreCompleto, telLimpio, emailOk, 'MACS ' + etq]);
    }
  });

  // crea o limpia la pestaña Contactos
  let sh = ss.getSheetByName('Contactos');
  if (!sh) sh = ss.insertSheet('Contactos');
  else sh.clearContents();

  sh.getRange(1, 1, filas.length, 5).setValues(filas);
  sh.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
  Logger.log('Contactos exportados: ' + (filas.length - 1) + ' → pestaña Contactos');
}
function debugEmiliano() {
  const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
  const ini = new Date();
  const fin = new Date(ini.getTime() + 15 * 86400000);
  const eventos = cal.getEvents(ini, fin);
  let s = 'Eventos en calendar: ' + eventos.length + ' | ';
  eventos.forEach(ev => { s += '[' + ev.getTitle() + '] '; });
  SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getRange('Z1').setValue(s);
}
function debugMatchEmiliano() {
  const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
  const ini = new Date();
  const fin = new Date(ini.getTime() + 15 * 86400000);
  const eventos = cal.getEvents(ini, fin);
  const norm = s => String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim();

  const cfg = getConfig('coord-rrhh');
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(cfg.hoja_nombre);
  const datos = sh.getDataRange().getValues();

  let s = '';
  eventos.forEach(ev => {
    const t = norm(ev.getTitle());
    s += 'TITULO:[' + t + '] screening?:' + (t.indexOf('screening')>-1) + ' | ';
    for (let i = 1; i < datos.length; i++) {
      const n = norm(datos[i][1]);
      if (n && n.indexOf('emiliano') > -1) {
        s += 'HOJA fila' + (i+1) + ':[' + n + '] incluido?:' + (t.indexOf(n)>-1) + ' | ';
      }
    }
  });
  SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getRange('Z1').setValue(s);
}
function buscarDuplicados() {
  const ss = ScriptApp.getProjectTriggers();
  Logger.log('Revisá manualmente con Ctrl+F en cada archivo (Codigo, Screening, Competencias, CalendarSync) estas funciones comunes:');
  Logger.log('procesarConChatGPT, insertarFilaLibre, esDuplicado, getConfig, getSheet, formatearRequisitos, aplicarEstadoNota');
}
function getNotaScreening(perfilId, nombreCandidato) {
  try {
    const cfg = getConfig(perfilId);
    const sh = getSheet(cfg);
    const datos = sh.getDataRange().getValues();
    const norm = s => String(s||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/\s+/g,' ').trim();
    const buscado = norm(nombreCandidato);
    for (let i = 1; i < datos.length; i++) {
      if (norm(datos[i][1]) === buscado) {
        return String(datos[i][11] || '');  // col 12 = Impresión screening
      }
    }
    return '';
  } catch (e) {
    return '';
  }
}
function getLinkScreening(perfilId, nombreCandidato) {
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

function reconciliarCompras() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const cfg = getConfig('analista-compras');
  const sheet = getSheet(cfg);

  const idDe = url => { const m = String(url || '').match(/[-\w]{25,}/); return m ? m[0] : ''; };
  const norm = s => String(s || '').toLowerCase().replace(/\.(pdf|docx?|doc)$/, '').replace(/\s+/g, ' ').trim();

  // lo que ya está en la hoja: por ID y por texto del link (col 8)
  const data = sheet.getDataRange().getValues();
  const ids = {}, textos = [];
  for (let i = 1; i < data.length; i++) {
    const cell = data[i][7];
    const id = idDe(cell);
    if (id) ids[id] = true;
    textos.push(norm(cell));
  }

  // recorrer carpeta de recibidos de compras
  const folder = DriveApp.getFolderById(cfg.folder_recibidos);
  const files = folder.getFiles();
  const faltan = [];
  let total = 0;
  while (files.hasNext()) {
    const f = files.next(); total++;
    const nombre = norm(f.getName());
    const presente = ids[f.getId()] || textos.some(t => t && (t === nombre || t.indexOf(nombre) !== -1));
    if (!presente) faltan.push([f.getName(), f.getUrl()]);
  }

  // salida
  let out = ss.getSheetByName('CVs_Faltantes_Compras');
  if (!out) out = ss.insertSheet('CVs_Faltantes_Compras');
  out.clear();
  out.getRange(1, 1, 1, 2).setValues([['Archivo en carpeta SIN fila en hoja', 'Link']]);
  if (faltan.length) out.getRange(2, 1, faltan.length, 2).setValues(faltan);

  Logger.log(`Compras → carpeta: ${total} archivos · en hoja: ${Object.keys(ids).length} IDs · faltan: ${faltan.length}. Ver pestaña "CVs_Faltantes_Compras".`);
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
  SpreadsheetApp.getUi().alert('Columnas 32 y 33 creadas.');
}
function verHojas() {
  const perfiles = perfilesActivos_();
  perfiles.forEach(p => {
    const perfilId = (typeof p === 'string') ? p : p.id;
    const cfg = getConfig(perfilId);
    const sh = getSheet(cfg);
    Logger.log(perfilId + ' → hoja: ' + sh.getName());
  });
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
function generarInformeCandidato(perfilId, fila) {
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
  const dato = (l,v) => '<div class="dl">'+l+'</div><div class="dv">'+(v && String(v).trim() ? v : 'Sin dato')+'</div>';
  const datosHTML = [
    dato('Edad', d.edad), dato('Localidad', d.localidad), dato('Formación', d.formacion),
    dato('Último empleo', d.ultimo_empleo), dato('Pretensión', d.pretension), dato('Composición familiar', d.familia),
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
function vistaPreviaInforme() {
  const nombre = 'Juan Pérez (EJEMPLO)';
  const posicion = 'Analista de Datos';
  const LOGO_URL = 'https://lh3.googleusercontent.com/d/1WdoIWop2mZ_KnSa9raQM1Xb5TFP2JvGO';
  const LOGO_VERDE = 'https://lh3.googleusercontent.com/d/1457Wfl9zeY1V3sAGtY1IT2BmF2lqpoy8';
  const cvUrl = 'https://drive.google.com/file/d/11fgltSUxtLEC5yaEjl3ZMW66IE0EXEkd/view';

  const d = {
    edad:'38 años', localidad:'Pilar, Bs. As.', formacion:'Lic. en Administración',
    ultimo_empleo:'Industrias ACME', pretension:'$2.800.000', familia:'Casado, dos hijos',
    telefono:'11-5555-5555', email:'juan@mail.com', movilidad:'Sí', disponibilidad:'15 días',
    herramientas:'Excel avanzado, Power BI, SQL, ChatGPT y Copilot',
    perfil:'Profesional con más de 8 años en análisis de datos en entornos industriales. Diseñó tableros de control y automatizó reportes de gestión. Combina sólida formación económica con manejo avanzado de herramientas de BI e IA generativa. Orientado a resultados y con buena lectura del negocio.',
    motivo_egreso:'Reestructuración del área.', motivacion:'Busca un proyecto con mayor desafío analítico.',
    recorrido:[
      {periodo:'2019-2025', empresa:'Industrias ACME · Analista de Datos Sr.', detalle:'Diseño de dashboards de KPIs y automatización de reportes para dirección.'},
      {periodo:'2016-2019', empresa:'Consultora XYZ · Analista', detalle:'Análisis financiero de proyectos y control de desvíos presupuestarios.'}
    ],
    competencias:[
      {nombre:'Pensamiento analítico', nivel:4, texto:'Demuestra solidez. Estructura problemas complejos y los traduce en métricas accionables.'},
      {nombre:'Manejo de herramientas BI', nivel:4, texto:'Demuestra solidez. Domina Power BI y SQL, con automatización de flujos de datos.'},
      {nombre:'Comunicación de resultados', nivel:3, texto:'Demuestra efectividad. Presenta hallazgos con claridad adaptándose a la audiencia.'},
      {nombre:'Autonomía', nivel:3, texto:'Demuestra efectividad. Gestiona sus entregables con buen criterio de priorización.'}
    ],
    conclusion:'Perfil técnico sólido, con fuerte dominio de herramientas analíticas y buena lectura del negocio. Se destaca por su autonomía y su capacidad para transformar datos en información útil para la toma de decisiones.\nComo aspecto a observar, conviene profundizar en su experiencia liderando proyectos transversales, ya que su trayectoria reciente fue más individual que de coordinación de equipos.'
  };

  const NIVEL_TXT = {1:'No demuestra',2:'Señales de desarrollo',3:'Demuestra efectividad',4:'Demuestra solidez'};
  const dato = (l,v) => v ? '<div class="dl">'+l+'</div><div class="dv">'+v+'</div>' : '';
  const datosHTML = [
    dato('Edad', d.edad), dato('Localidad', d.localidad), dato('Formación', d.formacion),
    dato('Último empleo', d.ultimo_empleo), dato('Pretensión', d.pretension), dato('Composición familiar', d.familia),
    dato('Teléfono', d.telefono), dato('Email', d.email), dato('Movilidad', d.movilidad),
    dato('Disponibilidad', d.disponibilidad), dato('Herramientas', d.herramientas)
  ].filter(Boolean).join('');
  const recHTML = d.recorrido.map(r =>
    '<div class="rec-item"><div class="rec-meta"><span class="rec-per">'+r.periodo+'</span><span class="rec-emp">'+r.empresa+'</span></div><p class="rec-det">'+r.detalle+'</p></div>'
  ).join('');
  const compHTML = d.competencias.map(c =>
    '<div class="cmp"><div class="cmp-head"><span class="cmp-name">'+c.nombre+'</span><span class="cmp-nivel">'+c.nivel+'/4 · '+NIVEL_TXT[c.nivel]+'</span></div><p>'+c.texto+'</p></div>'
  ).join('');
  const cvBtn = '<a href="'+cvUrl+'" target="_blank" style="display:inline-block;background:#fff;color:#1A6645;padding:9px 20px;border-radius:4px;text-decoration:none;font-weight:800;font-size:9.5pt">📄 Ver Currículum →</a>';
  const conclHTML = d.conclusion.split('\n').filter(Boolean).map(p=>'<p>'+p+'</p>').join('');

  const htmlFull = '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Vista previa</title><style>'+
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
'@media print{body{background:#fff;}.page{box-shadow:none;margin:0;max-width:100%;}@page{margin:10mm 0;}}'+
'</style></head><body contenteditable="true"><div class="page"><div class="body-cell">'+
'<div class="hero"><img src="'+LOGO_URL+'" alt="Blumia"><div class="tag">Informe de Candidato · MACS S.A.</div>'+
'<div class="herorow"><div><h1>'+nombre+'</h1><div class="sub">'+posicion+'</div></div>'+cvBtn+'</div></div>'+
'<div class="ficha-data">'+datosHTML+'</div>'+
'<h3>Trayectoria y perfil</h3><p class="perfil-text">'+d.perfil+'</p>'+
'<div class="extra"><div><div class="l">Motivo de egreso / situación</div><div class="v">'+d.motivo_egreso+'</div></div>'+
'<div><div class="l">Motivación para el cambio</div><div class="v">'+d.motivacion+'</div></div></div>'+
'<h3>Recorrido laboral</h3>'+recHTML+
'<h3>Evaluación por competencias</h3><div class="cmp-grid">'+compHTML+'</div>'+
'<div class="conclusion"><div class="eye">Conclusión del perfil</div>'+conclHTML+'</div>'+
'<div class="foot"><img src="'+LOGO_VERDE+'" alt="Blumia"></div>'+
'</div></div>'+
'<script>document.addEventListener("click",function(e){var a=e.target.closest("a");if(a&&a.href){e.preventDefault();window.open(a.href,"_blank");}});</script>'+
'</body></html>';

  const carpeta = DriveApp.getRootFolder();
  const archivo = carpeta.createFile('VISTA_PREVIA_Informe.html', htmlFull, MimeType.HTML);
  const fileId = archivo.getId();
  Logger.log('Vista previa: ' + URL_WEBAPP + '?tipo=verinforme&id=' + fileId);
}
function verCvUrl() {
  const cfg = getConfig('analista-compras');
  const sh = getSheet(cfg);
  const datos = sh.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][1]).indexOf('Bodean') !== -1) {
      const idCV = extraerFileId(sh, i+1, 8);
      Logger.log('Eduardo fila ' + (i+1));
      Logger.log('fileId CV: ' + idCV);
      Logger.log('valor col8: ' + sh.getRange(i+1, 8).getValue());
      const rt = sh.getRange(i+1, 8).getRichTextValue();
      Logger.log('link col8: ' + (rt ? rt.getLinkUrl() : 'no rich'));
      break;
    }
  }
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
function probarPDFvistaPrevia() {
  const PDFSHIFT_KEY = pdfshiftKey_();
  // genera el HTML de vista previa y agarra su contenido
  vistaPreviaInforme();
  // busca el archivo de vista previa recién creado
  const files = DriveApp.getFilesByName('VISTA_PREVIA_Informe.html');
  if (!files.hasNext()) { Logger.log('No encontré la vista previa'); return; }
  const html = files.next().getBlob().getDataAsString();

  const resp = UrlFetchApp.fetch('https://api.pdfshift.io/v3/convert/pdf', {
    method: 'post',
    headers: { 'Authorization': 'Basic ' + Utilities.base64Encode('api:' + PDFSHIFT_KEY) },
    contentType: 'application/json',
    payload: JSON.stringify({ source: html, format: 'A4', use_print: true }),
    muteHttpExceptions: true
  });
  if (resp.getResponseCode() !== 200) { Logger.log('Error PDFShift: ' + resp.getContentText().substring(0,300)); return; }

  const carpeta = DriveApp.getRootFolder();
  const pdf = carpeta.createFile(resp.getBlob().setName('PRUEBA_PDF.pdf'));
  pdf.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  Logger.log('PDF de prueba: ' + pdf.getUrl());
}
function probarPDFClientes() {
  const url = generarPDFClientes('analista-compras');
  Logger.log('Listo: ' + url);
}
function verScreeningFacundo() {
  const info = getLinkScreening('analista-datos', 'Facundo San Martino');
  Logger.log('link: ' + info.link);
  Logger.log('nota: ' + info.nota);
}
function verCarpetas() {
  const cfg = getConfig('analista-compras');
  Logger.log('folder_madre: ' + cfg.folder_madre);
  Logger.log('folder_recibidos: ' + cfg.folder_recibidos);
}
function diagnosticarCV() {
  const cfg = getConfig('analista-compras');
  const file = DriveApp.getFileById('1Ya-62HJZvpApqcnWwnnzeibJOrJk6p_d');
  Logger.log('Archivo: ' + file.getName());

  // 1. ¿Se puede leer el texto?
  try {
    const texto = textoDeCV('1Ya-62HJZvpApqcnWwnnzeibJOrJk6p_d');
    Logger.log('TEXTO OK (' + texto.length + ' chars). Primeros 300:\n' + texto.substring(0,300));
  } catch(e) {
    Logger.log('FALLO LECTURA: ' + e);
  }

  // 2. ¿Qué devuelve GPT?
  try {
    const datos = procesarConChatGPT(file, cfg);
    Logger.log('GPT devolvió: ' + JSON.stringify(datos));
  } catch(e) {
    Logger.log('FALLO GPT: ' + e);
  }
}
function vistaPreviaGuardar() {
  const nombre = 'Juan Pérez (EJEMPLO)';
  const posicion = 'Analista de Datos';
  const LOGO_URL = 'https://lh3.googleusercontent.com/d/1WdoIWop2mZ_KnSa9raQM1Xb5TFP2JvGO';
  const LOGO_VERDE = 'https://lh3.googleusercontent.com/d/1457Wfl9zeY1V3sAGtY1IT2BmF2lqpoy8';

  const htmlBase = '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><title>Prueba guardar</title><style>'+
'*{box-sizing:border-box;margin:0;padding:0;-webkit-print-color-adjust:exact;print-color-adjust:exact;}'+
'body{font-family:sans-serif;color:#1A1A1A;background:#F7F5F1;font-size:11pt;padding:20px;}'+
'.page{background:#fff;max-width:210mm;margin:16px auto;padding:20px;}'+
'.hero{background:#1A6645;color:#fff;padding:20px;border-bottom:4px solid #FBC67B;}'+
'.hero h1{font-size:23pt;}'+
'.campo{padding:12px;border-bottom:1px solid #ddd;}'+
'.dl{font-size:8pt;font-weight:800;text-transform:uppercase;color:#555;}'+
'#barra-guardar{position:fixed;top:12px;right:12px;z-index:9999;}'+
'#barra-guardar button{background:#FBC67B;color:#1A1A1A;border:none;padding:10px 18px;border-radius:6px;font-weight:800;font-size:11pt;cursor:pointer;}'+
'@media print{#barra-guardar{display:none!important;}}'+
'</style></head><body contenteditable="true">'+
'<div id="barra-guardar" contenteditable="false"><button onclick="guardarCambios()">💾 Guardar cambios</button></div>'+
'<div class="page"><div class="hero"><h1>'+nombre+'</h1><div>'+posicion+'</div></div>'+
'<div class="campo"><div class="dl">Localidad</div><div>Editá esto y guardá</div></div>'+
'<div class="campo"><div class="dl">Formación</div><div>Probá cambiar este texto</div></div>'+
'</div>'+
'<script>'+
'function guardarCambios(){var btn=document.querySelector("#barra-guardar button");btn.textContent="Guardando...";btn.disabled=true;'+
'var clon=document.documentElement.cloneNode(true);'+
'var b=clon.querySelector("#barra-guardar button");if(b){b.textContent="💾 Guardar cambios";b.disabled=false;}'+
'var htmlLimpio="<!DOCTYPE html>"+clon.outerHTML;'+
'google.script.run.withSuccessHandler(function(r){btn.textContent=(r&&r.ok)?"✓ Guardado":"Error: "+(r?r.error:"?");btn.disabled=false;setTimeout(function(){btn.textContent="💾 Guardar cambios";},2500);})'+
'.withFailureHandler(function(e){btn.textContent="FALLO: "+e.message;btn.disabled=false;}).guardarInformeEditado("__FILEID__",htmlLimpio);}'+
'</script>'+
'</body></html>';

  const carpeta = DriveApp.getRootFolder();
  const archivo = carpeta.createFile('PRUEBA_GUARDAR.html', htmlBase, MimeType.HTML);
  const fileId = archivo.getId();
  archivo.setContent(htmlBase.replace('__FILEID__', fileId));
  Logger.log('Prueba guardar: ' + URL_WEBAPP + '?tipo=verinforme&id=' + fileId);
}
function guardarInformeEditado(fileId, htmlContenido) {
  try {
    DriveApp.getFileById(fileId).setContent(htmlContenido);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: String(e) };
  }
}
function probarGuardarNota() {
  const cfg = getConfig('analista-compras');
  const sheet = getSheet(cfg);
  Logger.log('Hoja: ' + sheet.getName());
  // buscar primer candidato con nombre
  const datos = sheet.getDataRange().getValues();
  for (let i = 1; i < datos.length; i++) {
    if (datos[i][1]) {
      Logger.log('Probando fila ' + (i+1) + ' (' + datos[i][1] + ')');
      try {
        aplicarEstadoNota(sheet, i+1, 'competencias', 'A');
        Logger.log('aplicarEstadoNota OK. Estado ahora: ' + sheet.getRange(i+1, 9).getValue());
      } catch(e) {
        Logger.log('FALLO aplicarEstadoNota: ' + e);
      }
      break;
    }
  }
}
function probarBenitez() {
  const cfg = getConfig('analista-sgi');
  const sheet = getSheet(cfg);
  const data = sheet.getDataRange().getValues();
  const norm = s => s.toString().toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  const buscado = norm('Emmanuel Benitez');
  Logger.log('Buscado: "' + buscado + '"');
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] && norm(data[i][1]).indexOf('benitez') > -1) {
      Logger.log('Encontrado: "' + data[i][1] + '" → normalizado: "' + norm(data[i][1]) + '"');
      Logger.log('¿Matchea exacto con "emmanuel benitez"?: ' + (norm(data[i][1]) === buscado));
    }
  }
}
function buscarBenitezTodos() {
  const norm = s => s.toString().toLowerCase().trim().normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  listarPerfiles().forEach(p => {
    let cfg; try { cfg = getConfig(p.id); } catch(e){ return; }
    const sheet = getSheet(cfg);
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] && norm(data[i][1]).indexOf('benitez') > -1) {
        Logger.log(p.id + ' → fila ' + (i+1) + ': "' + data[i][1] + '" | nota(col16): "' + data[i][15] + '" | estado(col9): "' + data[i][8] + '"');
      }
    }
  });
}
function probarGuardarBenitez() {
  const cfg = getConfig('analista-sgi');
  const sheet = getSheet(cfg);
  const fila = 16;
  Logger.log('Antes → nota col16: "' + sheet.getRange(fila,16).getValue() + '" | estado: "' + sheet.getRange(fila,9).getValue() + '"');
  sheet.getRange(fila, 16).setValue('B');
  Logger.log('Después de escribir B → nota col16: "' + sheet.getRange(fila,16).getValue() + '"');
}
function verificarLinksCompetencias() {
  listarPerfiles().forEach(p => {
    let cfg; try { cfg = getConfig(p.id); } catch(e){ return; }
    const sh = getSheet(cfg);
    const datos = sh.getDataRange().getValues();
    for (let i = 1; i < datos.length; i++) {
      if (!datos[i][1]) continue;
      const rt = sh.getRange(i+1, 15).getRichTextValue(); // col 15 = Evaluar competencias
      let url = rt ? rt.getLinkUrl() : '';
      if (!url) { const f = sh.getRange(i+1,15).getFormula(); const m = f.match(/HYPERLINK\("([^"]+)"/i); if(m) url = m[1]; }
      if (url) {
        const mp = url.match(/perfil=([^&]+)/);
        const perfilEnLink = mp ? decodeURIComponent(mp[1]) : '?';
        if (perfilEnLink !== p.id) {
          Logger.log('⚠️ ' + p.id + ' fila ' + (i+1) + ' (' + datos[i][1] + ') → link dice: ' + perfilEnLink);
        }
      }
      break; // solo el primero de cada hoja para muestra
    }
  });
  Logger.log('--- fin ---');
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
  empresa = empresa || 'MACS';
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
  const url = generarInformeCandidato(perfilId, fila);
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
  empresa = empresa || 'MACS';
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
    let c={cv:0,scr:0,comp:0,emp:0,res:0,dsin:0,dpost:0,dcomp:0};
    for(let i=1;i<datos.length;i++){
      if(!datos[i][1])continue;
      if(/^Eliminado\|/.test(String(datos[i][8]||'')))continue;
      c.cv++;
      const e=String(datos[i][8]||'').trim();
      if(e===AVANZA_EMP)c.emp++;
      if(e===EN_RESERVA)c.res++;
      if(e===DESC_SIN)c.dsin++;
      if(e===DESC_POST)c.dpost++;
      else if(e.indexOf('Descartado')===0&&[DESC_SIN,DESC_COMP].indexOf(e)===-1)c.dpost++;
      if(e===DESC_COMP)c.dcomp++;
      if(EN_COMPETENCIAS.indexOf(e)!==-1)c.comp++;
      if(EN_COMPETENCIAS.indexOf(e)!==-1||HIZO_SCREENING_EXTRA.indexOf(e)!==-1)c.scr++;
    }
    const m=manualPrev[p.nombre]||{};
    sh.getRange(fila,1,1,16).setValues([[p.nombre,c.cv,c.scr,c.comp,c.emp,c.res,c.dsin,c.dpost,c.dcomp,
      m.banda||'','','','','',m.lectura||'',m.fuentes||FUENTES_DEF]]);
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
  empresa = empresa || 'MACS';
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
      if (banda && HIZO_SCREENING.indexOf(estado)!==-1 && mSal){
        const val=Number(mSal[1].replace(/[.\s]/g,''));
        if(val>0){
          totalSal++;
          if(val < banda*0.95) debajo++;
          else if(val > banda*1.05) encima++;
          else linea++;
        }
      }
      if (AVANZO.indexOf(estado)!==-1 && datos[i][6]){
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
function kanbanHTML(perfilId) {
  const cfg = getConfig(perfilId);
  const emp = (typeof cfg.empresa === 'string') ? cfg.empresa : 'MACS';
  const titulo = ((typeof cfg.nombre === 'string') ? cfg.nombre : perfilId) + ' · ' + emp;
  const data = doGetKanban(perfilId);
  const empresas = JSON.stringify(empresasYPerfiles());
  const empresaActual = (typeof cfg.empresa === 'string') ? cfg.empresa : 'MACS';
  const flyer = (typeof cfg.link_flyer === 'string') ? cfg.link_flyer : '';
  const perfilProt = String(cfg.perfil_protegido || '').toUpperCase() === 'SI';
  const perfilPdf = (!perfilProt && typeof cfg.link_perfil === 'string') ? cfg.link_perfil : '';
  return HtmlService.createHtmlOutput(KANBAN_TPL(perfilId, titulo, data, empresas, empresaActual, flyer, perfilPdf, perfilProt))
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function doGetKanban(perfilId) {
  const sh = getSheet(getConfig(perfilId));
  const d = sh.getDataRange().getValues();
const link = (r, col) => {
    const rt = sh.getRange(r, col).getRichTextValue();
    if (rt && rt.getLinkUrl()) return rt.getLinkUrl();
    const f = sh.getRange(r, col).getFormula();
    const m = f.match(/HYPERLINK\("([^"]+)"/i);
    return m ? m[1] : '';
  };
  const out = [];
  for (let i = 1; i < d.length; i++) {
    if (!d[i][1]) continue;
    const m = String(d[i][5] || '').match(/Remuneraci[oó]n pretendida:\s*\$?\s*([\d.\s]+)/i);
    out.push({
      fila: i + 1, nombre: d[i][1], estado: String(d[i][8] || '').trim(),
      recibido: d[i][0] ? Utilities.formatDate(new Date(d[i][0]), 'GMT-3', 'dd/MM/yyyy') : '',
      nota: d[i][4] || '', notaScr: d[i][11] || '', notaComp: d[i][15] || '',
      pretension: m ? m[1].trim() : '', zona: d[i][23] || '',
      tel: d[i][2] || '', mail: d[i][21] || '',
      resumen: d[i][6] || '', cv: extraerFileId(sh, i + 1, 8) || '',
      scrLink: link(i + 1, 11), compLink: link(i + 1, 15), informe: link(i + 1, 33),
      scrPDF: link(i + 1, 13), compPDF: link(i + 1, 17), postulacion: link(i + 1, 31), cita: d[i][36] || '', historial: d[i][37] || '', whatsapp: link(i + 1, 21), mailDescarte: d[i][39] || '', contactado: d[i][40] || '', notaScrEmp: d[i][43] || '', notaPresEmp: d[i][44] || '', notaScrEmp: d[i][43] || '', notaPresEmp: d[i][44] || '', formacion: d[i][22] || '', requisitos: d[i][5] || '', pretNum: (function(){var m=String(d[i][5]||'').match(/Remuneraci[oó]n pretendida:\s*\$?\s*([\d.\s]+)/i);return m?Number(m[1].replace(/[.\s]/g,'')):0;})()
    });
  }
  return JSON.stringify(out);
}

function moverEstado(perfilId, fila, estado) {
  const sh = getSheet(getConfig(perfilId));
  const c = sh.getRange(fila, 9);
  const r = c.getDataValidation();
  c.setDataValidation(null).setValue(estado);
  if (r) c.setDataValidation(r);
  return true;
}

function KANBAN_TPL(perfilId, titulo, dataJSON, empresasJSON, empresaActual, flyerUrl, perfilUrl, perfilProt) {
  const LOGO='https://lh3.googleusercontent.com/d/1WdoIWop2mZ_KnSa9raQM1Xb5TFP2JvGO';
  return `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Kanban · ${titulo}</title>
<script src="https://cdn.jsdelivr.net/npm/sortablejs@1.15.0/Sortable.min.js"><\/script>
<style>
:root{--green:#1A6645;--glight:#E6F2EC;--amber:#F5C842;--bg:#F7F5F1;--line:#DDDBD7;--mid:#555;--red:#E06666;--font:'PP Pangram Sans','Plus Jakarta Sans',sans-serif}
*{box-sizing:border-box;margin:0;padding:0}body{font-family:var(--font);background:var(--bg);color:#1A1A1A;font-size:14px}
header{background:var(--green);color:#fff;padding:16px 20px;display:flex;align-items:center;gap:14px;position:sticky;top:0;z-index:10;flex-wrap:wrap}
header img{height:34px}header h1{font-size:18px;font-weight:800}
.nav{display:flex;gap:8px;margin-left:auto;flex-wrap:wrap;align-items:center}
.nav button{padding:7px 14px;border-radius:6px;font-size:12px;font-weight:800;color:#fff;background:rgba(255,255,255,.15);border:none;cursor:pointer;font-family:var(--font)}
.nav .proc{background:var(--amber);color:#1A1A1A}
.nav select{padding:7px 10px;border-radius:6px;border:none;font-family:var(--font);font-weight:800;font-size:12px;cursor:pointer}
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
function pedirPerfilPin(){var p=prompt('PIN para ver el perfil del puesto (solo selectores):');if(!p)return;google.script.run.withSuccessHandler(function(u){abrirPerfilLink(u);}).withFailureHandler(function(e){alert((e&&e.message)||e);}).linkPerfilConPin(PERFIL,p);}
function abrirPerfilLink(u){var d=document.createElement('div');d.style.cssText='position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,.45);z-index:99999;display:flex;align-items:center;justify-content:center';d.innerHTML='<div style="background:#fff;border-radius:10px;padding:26px 30px;text-align:center;font-family:sans-serif;max-width:340px"><p style="margin:0 0 16px;font-size:15px">PIN correcto ✓</p><a id="pplk" target="_blank" rel="noopener" style="display:inline-block;background:#1A6645;color:#fff;padding:11px 22px;border-radius:6px;text-decoration:none;font-weight:700">Abrir perfil del puesto</a><div style="margin-top:14px"><a href="#" id="ppcl" style="color:#777;font-size:13px">Cerrar</a></div></div>';document.body.appendChild(d);var l=document.getElementById('pplk');l.href=u;l.onclick=function(){setTimeout(function(){d.remove();},400);};document.getElementById('ppcl').onclick=function(ev){ev.preventDefault();d.remove();};}
let navHTML='<select onchange="window.top.location.href=\\'${URL_WEBAPP}?tipo=kanban&perfil=\\'+this.value">';
Object.keys(EMPRESAS).forEach(function(emp){
  navHTML+='<optgroup label="'+emp+'">';
  EMPRESAS[emp].forEach(function(p){navHTML+='<option value="'+p[0]+'"'+(p[0]===PERFIL?' selected':'')+'>'+p[1]+'</option>';});
  navHTML+='</optgroup>';
});
navHTML+='</select>';
navHTML+='<button onclick="window.top.location.href=\\'${URL_WEBAPP}?tipo=kanban&perfil=${perfilId}\\'">↻ Actualizar</button>';
navHTML+='<button class="proc" onclick="window.open(\\'${URL_WEBAPP}?tipo=procesohtml&perfil=${perfilId}\\',\\'_blank\\')">Informe proceso</button>';
navHTML+='<button class="proc" onclick="if(confirm(\\'¿Generar PDF cliente de esta búsqueda?\\'))window.open(\\'${URL_WEBAPP}?tipo=pdfcliente&perfil=${perfilId}\\',\\'_blank\\')">PDF cliente</button>';
navHTML+='<button style="background:#E06666;color:#fff" onclick="descartesMasivo()">✉ Descartes</button>';
navHTML+='<button onclick="bajarContactos(this)">📇 Contactos</button>';
navHTML+='<button id="avBtn" onclick="verAvisos()" style="position:relative">🔔 Avisos<span id="avCnt" style="display:none;position:absolute;top:-6px;right:-6px;background:#B91C1C;color:#fff;border-radius:10px;font-size:11px;font-weight:800;padding:1px 6px"></span></button>';
navHTML+='<button onclick="abrirPapelera()">🗑 Papelera ('+ELIM.length+')</button>';
navHTML+='<button style="background:#F5C842;color:#1A1A1A" onclick="window.top.location.href=\\'${URL_WEBAPP}?tipo=banco&desde=${perfilId}\\'">★ Banco</button>';
navHTML+=${perfilProt?'true':'false'}?'<button onclick="pedirPerfilPin()">🔒 Perfil del puesto</button>':'<button onclick="window.open(\\'${perfilUrl}\\',\\'_blank\\')">📋 Perfil del puesto</button>';
navHTML+='<button onclick="window.open(\\'${flyerUrl}\\',\\'_blank\\')">📢 Flyer</button>';
navHTML+='<button onclick="prompt(\\'Link del formulario de postulación (copialo):\\',\\'${URL_WEBAPP}?tipo=postulacion&perfil=${perfilId}\\')">🔗 Link postulación</button>';
document.getElementById('nav').innerHTML=navHTML;
avInit();
const COLS=[['Análisis de CV',['Nuevo','Contactado','Pendiente','']],['Screening',['Avanza a screening','Screening agendado']],['Competencias',['Avanza a competencias','Competencias agendada']],['Screening empresa',['Avanza a screening empresa','Screening empresa agendado']],['Presencial empresa',['Avanza a presencial empresa','Avanza a empresa','Entrevista empresa agendada','En cliente']],['Contratado',['Contratado']]];
const BASE={'Análisis de CV':'Nuevo','Screening':'Avanza a screening','Competencias':'Avanza a competencias','Screening empresa':'Avanza a screening empresa','Presencial empresa':'Avanza a presencial empresa','Contratado':'Contratado'};
const DESC_TIPOS=['Descartado sin screening','Descartado post screening','Descartado en competencias','Descartado en screening empresa','Descartado presencial empresa','Descartado Fuera de zona','Descartado NO le interesa'];
const RES_TIPOS=['En reserva','Reserva competencias','Reserva screening empresa','Reserva presencial empresa'];
const nota=n=>n?'<span class="nota n'+String(n).replace('+','')+'">'+n+'</span>':'';
const tag=e=>/agendad/i.test(e)?'<span class="tag ag">Agendado</span>':'';
const card=c=>'<div class="card" data-fila="'+c.fila+'" onclick="ficha('+c.fila+')"><button class="cx" title="Eliminar" onclick="event.stopPropagation();eliminar('+c.fila+')">✕</button>'+nota(c.notaPresEmp||c.notaScrEmp||c.notaComp||c.notaScr||c.nota)+'<div class="n">'+c.nombre+'</div><div class="meta">'+(c.zona||'—')+(c.pretension?' · $'+c.pretension:'')+'</div>'+(c.cita?'<div class="cita-tag">📅 '+c.cita+'</div>':'')+(c.contactado?'<span class="tag" style="background:#D1FAE5;color:#065F46">✓ Contact. '+c.contactado+'</span>':'')+tag(c.estado)+'</div>';
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
document.querySelectorAll('.col-body,.dbody').forEach(function(col){new Sortable(col,{group:'kb',animation:150,forceFallback:true,fallbackOnBody:true,scroll:true,bubbleScroll:true,scrollSensitivity:80,scrollSpeed:15,onEnd:function(evt){var fila=+evt.item.dataset.fila;var estado=evt.to.dataset.estado;var cc=DATA_ALL.find(function(x){return x.fila===fila;});if(cc&&estado)cc.estado=estado;if(estado&&evt.from!==evt.to){google.script.run.withSuccessHandler(toast).withFailureHandler(function(e){alert('Error: '+e.message);window.top.location.reload();}).moverEstado(PERFIL,fila,estado);}}});});
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
function ficha(f){var c=DATA.find(function(x){return x.fila===f;});var row=function(l,v){return v?'<div class="frow"><span class="fl">'+l+'</span><span class="fv">'+v+'</span></div>':'';};var btn=function(u,t){return u?'<a href="'+u+'" target="_blank" class="mbtn">'+t+'</a>':'';};var waBtn='<button class="mbtn wa" onclick="menuWA('+c.fila+')">💬 WhatsApp ▾</button>';var mailBtn=(c.mail&&c.mail!=='Sin dato'&&c.mail.indexOf('@')>-1)?'<a href="https://mail.google.com/mail/?view=cm&to='+c.mail+'" target="_blank" class="mbtn" style="background:#EA4335;color:#fff">✉ Mail</a>':'';var cvBtn=c.cv?'<a href="https://drive.google.com/file/d/'+c.cv+'/view" target="_blank" class="mbtn alt">Ver CV</a>':'';var genBtn=c.informe?'':'<button class="mbtn gen" onclick="if(confirm(\\'Generar informe con IA. Puede tener costo. ¿Continuar?\\'))generar('+c.fila+',this)">Generar informe</button>';var delBtn='<button class="mbtn" style="background:#B91C1C;color:#fff" onclick="eliminar('+c.fila+')">🗑 Eliminar</button>';var bancoBtn='<button class="mbtn" style="background:#7C3AED" onclick="menuBanco('+c.fila+')">★ Banco de talento</button>';var notasEmp=/screening empresa|presencial empresa|avanza a empresa|entrevista empresa|reserva screening empresa|reserva presencial empresa/i.test(c.estado)?['screening-empresa','presencial-empresa'].map(function(ep){return '<div class="notarow"><span>'+(ep==='screening-empresa'?'Emp. screening':'Emp. presencial')+'</span>'+['A','B','C'].map(function(n){return '<button class="nb" onclick="setNotaEmp('+c.fila+',\\''+ep+'\\',\\''+n+'\\',this)">'+n+'</button>';}).join('')+'</div>';}).join(''):'';document.getElementById('mbody').innerHTML=row('Recibido',c.recibido)+row('Estado',c.estado)+row('Cita agendada',c.cita)+(c.historial?'<div class="frow"><span class="fl">Historial entrevistas</span><span class="fv" style="white-space:pre-line">'+c.historial+'</span></div>':'')+(c.contactado?'<div class="frow"><span class="fl">Último contacto</span><span class="fv">'+c.contactado+' <a href="#" onclick="borrarMarca('+c.fila+');return false" style="color:#aaa;font-size:11px;text-decoration:none;margin-left:6px">✕ borrar</a></span></div>':'')+'<div class="frow"><span class="fl">Registro de contactos</span><span class="fv" id="regCont" style="white-space:pre-line">cargando…</span></div>'+row('Teléfono',c.tel)+row('Email',c.mail)+row('Zona',c.zona)+row('Formación',c.formacion)+row('Pretensión',c.pretension?'$'+c.pretension:'')+row('Nota screening',c.notaScr)+row('Nota competencias',c.notaComp)+row('Nota screening empresa',c.notaScrEmp)+row('Nota presencial empresa',c.notaPresEmp)+'<div class="fres">'+(c.resumen||'Sin resumen')+'</div><div class="notabox"><div class="notatit">Poner nota manual</div>'+['cv','screening','competencias'].map(function(et){return '<div class="notarow"><span>'+({cv:'CV',screening:'Screening',competencias:'Competencias'}[et])+'</span>'+['A','B+','B','C'].map(function(n){return '<button class="nb" onclick="setNota('+c.fila+',\\''+et+'\\',\\''+n+'\\',this)">'+n+'</button>';}).join('')+'</div>';}).join('')+notasEmp+'</div><div class="mbtns">'+waBtn+mailBtn+cvBtn+btn(c.scrPDF,'Ver screening')+btn(c.compPDF,'Ver competencias')+btn(c.postulacion,'Ver postulación')+btn(c.scrLink,'Screening')+btn(c.compLink,'Competencias')+btn(c.informe,'Ver informe')+genBtn+bancoBtn+delBtn+'</div>';document.getElementById('mtit').textContent=c.nombre;document.getElementById('modal').classList.add('show');cargarRegCont(f);}
function cargarRegCont(f){google.script.run.withSuccessHandler(function(l){var e=document.getElementById('regCont');if(!e)return;e.innerHTML=(l.length?l.map(function(x){return avEsc(x.fecha+' · '+x.tipo);}).join('<br>'):'Sin contactos registrados')+' <a href="#" onclick="regContacto('+f+');return false" style="font-size:12px">+ registrar</a>';}).withFailureHandler(function(){var e=document.getElementById('regCont');if(e)e.textContent='—';}).getContactosRegistro(PERFIL,f);}
function regContacto(f){var t=prompt('¿Qué contacto hiciste? (ej: Llamé, Dejé mensaje, Respondió por mail)');if(!t||!t.trim())return;google.script.run.withSuccessHandler(function(){cargarRegCont(f);}).withFailureHandler(function(e){alert('Error: '+e.message);}).registrarContactoManual(PERFIL,f,t.trim());}
function bajarContactos(b){b.disabled=true;google.script.run.withSuccessHandler(function(r){b.disabled=false;var a=document.createElement('a');a.href=URL.createObjectURL(new Blob([r.csv],{type:'text/csv'}));a.download='contactos-blumia.csv';a.click();alert(r.n+' contactos. Importalo en contacts.google.com → Importar.');}).withFailureHandler(function(e){b.disabled=false;alert('Error: '+e.message);}).exportarContactosCSV();}
function menuWA(f){
  var c=DATA.find(function(x){return x.fila===f;});
  var ops=[['contacto-screening','📞 Contacto screening'],['contacto-competencias','💼 Contacto competencias'],['contacto-empresa-virtual','🏢 Empresa (virtual)'],['contacto-empresa-presencial','🏢 Empresa (presencial)'],['descarte-competencias','❌ Descarte competencias'],['descarte-empresa','❌ Descarte empresa']];
  var h='<div style="font-size:13px;margin-bottom:12px;color:#555">Elegí qué mensaje enviar a <b>'+c.nombre+'</b>:</div>';
  ops.forEach(function(o){h+='<button class="mbtn wa" style="display:block;width:100%;margin-bottom:8px;text-align:left" onclick="enviarWA('+f+',\\''+o[0]+'\\',this)">'+o[1]+'</button>';});
  h+='<button class="mbtn alt" style="width:100%;margin-top:6px" onclick="ficha('+f+')">← Volver</button>';
  document.getElementById('mtit').textContent='WhatsApp · '+c.nombre;
  document.getElementById('mbody').innerHTML=h;
}
function enviarWA(f,tipo,btn){
  btn.textContent='Preparando...';
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.ok){window.open(r.url,'_blank');btn.textContent='✓ Abierto';}
    else{alert('Error: '+(r?r.error:'desconocido'));}
  }).withFailureHandler(function(e){alert('Error: '+e.message);}).armarComunicacion(PERFIL,f,tipo);
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
  }).withFailureHandler(function(e){alert('Error: '+e.message);btn.textContent='★ Guardar en banco';btn.disabled=false;}).enviarABanco(PERFIL,f,cat,sen,nota);
}
function setNota(f,et,n,btn){btn.textContent='...';google.script.run.withSuccessHandler(function(){toast();cerrar();setTimeout(function(){window.top.location.reload();},800);}).withFailureHandler(function(e){alert('Error: '+e.message);}).notaManual(PERFIL,f,et,n);}
function setNotaEmp(f,et,n,btn){btn.textContent='...';google.script.run.withSuccessHandler(function(){toast();cerrar();setTimeout(function(){window.top.location.reload();},800);}).withFailureHandler(function(e){alert('Error: '+e.message);}).notaEmpresa(PERFIL,f,et,n);}
function generar(f,btn){btn.textContent='Generando...';btn.disabled=true;google.script.run.withSuccessHandler(function(url){window.open(url,'_blank');btn.textContent='✓ Listo';}).withFailureHandler(function(e){alert('Error: '+e.message);btn.textContent='Generar informe';btn.disabled=false;}).generarInformeCandidato(PERFIL,f);}
function descartesMasivo(){google.script.run.withSuccessHandler(function(p){if(p.sinScr===0&&p.postScr===0){alert('No hay descartes pendientes de mail en esta búsqueda.');return;}var h='<div style="font-size:13px;margin-bottom:14px">Se van a enviar <b>'+(p.sinScr+p.postScr)+'</b> mails en '+p.puesto+'.<br>Sin screening: <b>'+p.sinScr+'</b> · Post screening: <b>'+p.postScr+'</b><br><span style="color:#888">[NOMBRE] se reemplaza por el nombre de cada persona.</span></div>';h+='<div class="notatit">Texto · Descartado sin screening ('+p.sinScr+')</div><textarea id="txtSin" style="width:100%;height:150px;font-family:inherit;font-size:12px;padding:8px;border:1px solid #DDD;border-radius:6px;margin-bottom:12px">'+p.textoSin+'</textarea>';h+='<div class="notatit">Texto · Descartado post screening ('+p.postScr+')</div><textarea id="txtPost" style="width:100%;height:150px;font-family:inherit;font-size:12px;padding:8px;border:1px solid #DDD;border-radius:6px;margin-bottom:12px">'+p.textoPost+'</textarea>';h+='<div class="mbtns"><button class="mbtn" style="background:#E06666" onclick="enviarDesc(this)">✉ Enviar '+(p.sinScr+p.postScr)+' mails</button></div>';document.getElementById('mtit').textContent='Mails de descarte';document.getElementById('mbody').innerHTML=h;document.getElementById('modal').classList.add('show');}).withFailureHandler(function(e){alert('Error: '+e.message);}).previaDescartes(PERFIL);}
function enviarDesc(btn){if(!confirm('¿Confirmás el envío? No se puede deshacer.'))return;var s=document.getElementById('txtSin').value;var p=document.getElementById('txtPost').value;btn.textContent='Enviando...';btn.disabled=true;google.script.run.withSuccessHandler(function(r){alert('Enviados: '+r.enviados+(r.sinMail?' · Sin mail válido: '+r.sinMail:''));window.top.location.reload();}).withFailureHandler(function(e){alert('Error: '+e.message);btn.disabled=false;btn.textContent='✉ Enviar';}).enviarDescartesMasivo(PERFIL,s,p);}
function borrarMarca(f){if(!confirm('¿Borrar la marca de contacto?'))return;google.script.run.withSuccessHandler(function(){toast();cerrar();setTimeout(function(){window.top.location.reload();},800);}).withFailureHandler(function(e){alert('Error: '+e.message);}).borrarMarcaContacto(PERFIL,f);}
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
function avPendientes(L){
  var out=[],now=Date.now();
  L.forEach(function(c){
    if(!avActivo(c))return;
    var ep=avEtapa(c.estado);if(!ep)return;
    var p=function(k,ic,txt,urg){out.push({k:k+'|'+c.fila,ic:ic,txt:txt,f:c.fila,n:c.nombre,urg:!!urg});};
    if(!/agendad/i.test(c.estado)){
      var ct=String(c.contactado||'').split(' · ')[0].trim().toLowerCase();
      if(ct!==ep.k){p('contactar','➜','Falta contactar para '+ep.n);return;}
      var f=avFecha(c.contactado);
      if(!c.cita&&f){var d=Math.floor((now-f)/864e5);if(d>=2)p('norespondio','⏳','Contactado para '+ep.n+' hace '+d+' días y todavía no agendó');}
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
  out.sort(function(a,b){return (b.urg?1:0)-(a.urg?1:0);});
  return out;
}
function avNovedades(L){
  var prev=avLS('snap'),snap={},nov=[],d=new Date(),dd=function(n){return ('0'+n).slice(-2);},cuando=dd(d.getDate())+'/'+dd(d.getMonth()+1)+' '+dd(d.getHours())+':'+dd(d.getMinutes());
  L.forEach(function(c){
    var s={e:c.estado,c:c.cita||''};snap[c.fila]=s;if(!prev)return;
    var a=prev[c.fila],p=function(ic,txt){nov.push({ic:ic,txt:txt,f:c.fila,n:c.nombre,h:cuando});};
    if(!a){if(!/^Eliminado/.test(s.e))p('🆕','Nuevo CV recibido');return;}
    if(s.e!==a.e&&!/^Eliminado/.test(s.e)){
      if(s.e==='Contratado')p('🎉','Contratado/a');
      else if(/^Descartad/i.test(s.e))p('❌','Descartado/a: '+s.e.replace(/^Descartado\\s*/i,''));
      else p('🔄','Pasó a «'+s.e+'»'+(a.e?' (antes: '+a.e+')':''));
    }
    if(s.c&&s.c!==a.c)p('📅','Agendó para el '+s.c);
  });
  avLS('snap',snap);return nov;
}
function avPintar(){var b=document.getElementById('avCnt');if(!b)return;var n=AV_PEND.filter(function(x){return x.urg;}).length+(avLS('sinver')||0);b.textContent=n||'';b.style.display=n?'':'none';}
function avAbrir(f){if(DATA.some(function(x){return x.fila===f;}))ficha(f);else cerrar();}
function verAvisos(){
  avLS('sinver',0);avPintar();
  var item=function(x,nov){return '<div class="frow" style="cursor:pointer;align-items:flex-start" onclick="avAbrir('+x.f+')"><span class="fv"><b>'+x.ic+' '+avEsc(x.n)+'</b><br><small>'+avEsc(x.txt)+(nov?' · '+x.h:'')+'</small></span></div>';};
  document.getElementById('mtit').textContent='🔔 Avisos';
  document.getElementById('mbody').innerHTML='<div style="font-weight:800;color:#B91C1C;margin:2px 0 6px">Pendientes ('+AV_PEND.length+')</div>'+(AV_PEND.length?AV_PEND.map(function(x){return item(x,false);}).join(''):'<div style="color:#777;font-size:14px">No hay pendientes. 👌</div>')+'<div style="font-weight:800;color:#1A6645;margin:16px 0 6px">Novedades</div>'+(AV_NOV.length?AV_NOV.slice(0,40).map(function(x){return item(x,true);}).join(''):'<div style="color:#777;font-size:14px">Sin novedades desde la última visita.</div>');
  document.getElementById('modal').classList.add('show');
}
function avPopups(L){
  var cont=document.getElementById('avPop');
  if(!cont){cont=document.createElement('div');cont.id='avPop';cont.style.cssText='position:fixed;top:14px;right:14px;z-index:9999;display:flex;flex-direction:column;gap:8px;max-width:340px';document.body.appendChild(cont);}
  var mk=function(x){var d=document.createElement('div');d.style.cssText='background:#fff;border-left:4px solid '+(x.urg?'#B91C1C':'#1A6645')+';box-shadow:0 4px 14px rgba(0,0,0,.2);border-radius:6px;padding:10px 12px;font-size:13px;cursor:pointer';d.innerHTML='<b>'+x.ic+' '+avEsc(x.n)+'</b><br>'+avEsc(x.txt);d.onclick=function(){d.remove();if(x.panel)verAvisos();else if(x.f)avAbrir(x.f);};cont.appendChild(d);setTimeout(function(){if(d.parentNode)d.remove();},x.urg?15000:8000);};
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
function eliminar(f){var c=DATA.find(function(x){return x.fila===f;});if(!c)return;if(!confirm('¿Eliminar a '+c.nombre+' del tablero?\\nNo se pierde nada: queda en la 🗑 Papelera y se puede restaurar.'))return;google.script.run.withSuccessHandler(function(){toast();cerrar();setTimeout(function(){window.top.location.reload();},700);}).withFailureHandler(function(e){alert('Error: '+e.message);}).eliminarCandidato(PERFIL,f);}
function abrirPapelera(){document.getElementById('mtit').textContent='Papelera';var h=ELIM.length?'':'<div style="color:#777;font-size:14px">La papelera está vacía.</div>';ELIM.forEach(function(c){var p=String(c.estado).split('|');h+='<div class="frow" style="align-items:center"><span class="fv"><b>'+c.nombre+'</b><br><small>'+(p[1]||'sin estado')+(p[2]?' · eliminado el '+p[2]:'')+'</small></span><button class="mbtn" onclick="restaurar('+c.fila+')">↩ Restaurar</button></div>';});document.getElementById('mbody').innerHTML=h;document.getElementById('modal').classList.add('show');}
function restaurar(f){google.script.run.withSuccessHandler(function(){toast();cerrar();setTimeout(function(){window.top.location.reload();},700);}).withFailureHandler(function(e){alert('Error: '+e.message);}).restaurarCandidato(PERFIL,f);}
function cerrar(){document.getElementById('modal').classList.remove('show');}
</script></body></html>`;
}
function verLinkPDF() {
  const sh = getSheet(getConfig('analista-compras'));
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (d[i][16]) { // col 17 competencias tiene algo
      Logger.log('Fila ' + (i+1) + ' - ' + d[i][1]);
      Logger.log('Valor: ' + d[i][16]);
      Logger.log('Formula: ' + sh.getRange(i+1, 17).getFormula());
      const rt = sh.getRange(i+1, 17).getRichTextValue();
      Logger.log('RichText link: ' + (rt ? rt.getLinkUrl() : 'no hay'));
      break;
    }
  }
}
function notaManual(perfilId, fila, etapa, nota) {
  const sh = getSheet(getConfig(perfilId));
  const col = {cv:5, screening:12, competencias:16}[etapa];
  sh.getRange(fila, col).setValue(nota);
  aplicarEstadoNota(sh, fila, etapa, nota);
  return true;
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

function testCitas() {
  const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
  const desde = new Date(); desde.setDate(desde.getDate() - 30);
  const hasta = new Date(); hasta.setDate(hasta.getDate() + 60);
  const eventos = cal.getEvents(desde, hasta);
  Logger.log('Cantidad de eventos encontrados: ' + eventos.length);
  eventos.slice(0, 15).forEach(e => Logger.log('Evento: "' + e.getTitle() + '"'));
}
function testPerfiles() {
  const perfiles = listarPerfiles();
  perfiles.forEach(p => {
    try {
      const cfg = getConfig(p.id);
      const sh = getSheet(cfg);
      Logger.log('OK: ' + p.id + ' → hoja: ' + sh.getName());
    } catch(e) {
      Logger.log('FALLA en ' + p.id + ': ' + e.message);
    }
  });
}
function testUno() {
  const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
  const desde = new Date(); desde.setDate(desde.getDate() - 30);
  const hasta = new Date(); hasta.setDate(hasta.getDate() + 60);
  const eventos = cal.getEvents(desde, hasta);
  const norm = s => String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'');

  const sh = getSheet(getConfig('analista-datos'));
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (!d[i][1]) continue;
    if (norm(d[i][1]).includes('vergniaud')) {
      Logger.log('Encontré en Sheet fila ' + (i+1) + ': "' + d[i][1] + '" → normalizado: ' + norm(d[i][1]));
      const ev = eventos.find(e => { const m = e.getTitle().match(/\(([^)]+)\)/); return m && norm(m[1]) === norm(d[i][1]); });
      Logger.log('Match evento: ' + (ev ? ev.getTitle() : 'NO ENCONTRÓ'));
      const evParcial = eventos.find(e => e.getTitle().toLowerCase().includes('vergniaud'));
      Logger.log('Evento que contiene vergniaud: ' + (evParcial ? evParcial.getTitle() : 'ninguno'));
    }
  }
}
function crearTriggerCitas() {
  ScriptApp.getProjectTriggers().forEach(t => { if (t.getHandlerFunction() === 'syncCitasCalendar') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('syncCitasCalendar').timeBased().everyHours(1).create();
}
function testMariano(){
  const sh=getSheet(getConfig('analista-datos'));
  const d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++){
    if(String(d[i][1]).toLowerCase().includes('vergniaud')){
      Logger.log('Estado: "'+d[i][8]+'"');
      Logger.log('Cita col37: "'+d[i][36]+'"');
    }
  }
}
function testCeleste() {
  const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
  const desde = new Date(); desde.setDate(desde.getDate() - 180);
  const hasta = new Date(); hasta.setDate(hasta.getDate() + 90);
  const eventos = cal.getEvents(desde, hasta);
  const norm = s => String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'');

  const sh = getSheet(getConfig('analista-sgi'));
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][1]).toLowerCase().includes('cuellar')) {
      Logger.log('En Sheet: "' + d[i][1] + '" → norm: ' + norm(d[i][1]));
      Logger.log('Estado: "' + d[i][8] + '"');
      // buscar eventos que contengan cuellar
      const suyos = eventos.filter(e => e.getTitle().toLowerCase().includes('cuellar'));
      Logger.log('Eventos con "cuellar": ' + suyos.length);
      suyos.forEach(e => {
        const m = e.getTitle().match(/\(([^)]+)\)/);
        Logger.log('  Título: "' + e.getTitle() + '"');
        Logger.log('  Nombre en paréntesis: "' + (m?m[1]:'SIN PARÉNTESIS') + '" → norm: ' + (m?norm(m[1]):'—'));
        Logger.log('  ¿Matchea?: ' + (m && norm(m[1])===norm(d[i][1])));
      });
    }
  }
}
function testMailEvento() {
  const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
  const desde = new Date(); desde.setDate(desde.getDate() - 180);
  const hasta = new Date(); hasta.setDate(hasta.getDate() + 90);
  const eventos = cal.getEvents(desde, hasta);
  const ev = eventos.find(e => e.getTitle().toLowerCase().includes('cuellar'));
  if (ev) {
    Logger.log('Título: ' + ev.getTitle());
    Logger.log('Invitados: ' + ev.getGuestList().map(g => g.getEmail()).join(', '));
    Logger.log('Descripción: ' + ev.getDescription().substring(0, 300));
  }
}
function testMailSheet() {
  const sh = getSheet(getConfig('analista-sgi'));
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][1]).toLowerCase().includes('cuellar')) {
      Logger.log('col22 (mail): "' + d[i][21] + '"');
      Logger.log('col6 requisitos: "' + String(d[i][5]).substring(0,200) + '"');
    }
  }
}
function buscarMail() {
  const sh = getSheet(getConfig('analista-sgi'));
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (!d[i][1]) continue;
    // busca una fila donde alguna celda tenga un @
    const idx = d[i].findIndex(c => String(c).includes('@'));
    if (idx >= 0) {
      Logger.log(d[i][1] + ' → mail en columna ' + (idx+1) + ': "' + d[i][idx] + '"');
    }
  }
}
function testCarolina(){
  const sh=getSheet(getConfig('analista-sgi'));
  const d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++){
    if(String(d[i][1]).toLowerCase().includes('azocar')||String(d[i][1]).toLowerCase().includes('azócar')){
      Logger.log('Estado col9: "'+d[i][8]+'"');
    }
  }
}
function testCeleste2(){
  const cal=CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
  const desde=new Date();desde.setDate(desde.getDate()-180);
  const hasta=new Date();hasta.setDate(hasta.getDate()+90);
  const eventos=cal.getEvents(desde,hasta);
  const norm=s=>String(s).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').replace(/[^a-z0-9]/g,'');
  const sh=getSheet(getConfig('analista-sgi'));
  const d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++){
    if(String(d[i][1]).toLowerCase().includes('cuellar')){
      const nombreN=norm(d[i][1]);
      const mailN=String(d[i][21]||'').toLowerCase().trim();
      Logger.log('nombreN: '+nombreN+' | mail: '+mailN);
      const suyos=eventos.filter(e=>{
        if(mailN&&mailN.includes('@')){const inv=e.getGuestList().map(g=>g.getEmail().toLowerCase());if(inv.includes(mailN))return true;}
        const m=e.getTitle().match(/\(([^)]+)\)/);if(!m)return false;const evN=norm(m[1]);
        return evN===nombreN||(evN.length>=10&&(nombreN.includes(evN)||evN.includes(nombreN)));
      });
      Logger.log('Eventos match: '+suyos.length);
      const ahora=new Date();
      suyos.forEach(e=>Logger.log('  '+e.getTitle()+' | '+e.getStartTime()+' | futura: '+(e.getStartTime()>=ahora)));
    }
  }
}
function verTriggers(){
  ScriptApp.getProjectTriggers().forEach(t=>Logger.log(t.getHandlerFunction()+' | '+t.getEventType()));
}
function reponerDesplegableEstados() {
  const ESTADOS = [
    'Pendiente','Nuevo','Contactado',
    'Avanza a screening','Screening agendado',
    'Avanza a competencias','Competencias agendada',
    'Avanza a screening empresa','Screening empresa agendado',
    'Avanza a presencial empresa','Entrevista empresa agendada',
    'En cliente','Contratado','En reserva',
    'Descartado sin screening','Descartado post screening','Descartado en competencias',
    'Descartado en screening empresa','Descartado presencial empresa',
    'Descartado Fuera de zona','Descartado NO le interesa'
  ];
  const perfiles = ['coord-rrhh','analista-datos','analista-compras','analista-sgi'];
  const regla = SpreadsheetApp.newDataValidation().requireValueInList(ESTADOS, true).setAllowInvalid(true).build();

  perfiles.forEach(id => {
    const sh = getSheet(getConfig(id));
    const d = sh.getDataRange().getValues();
    // corregir estados: "Avanza a empresa"→"Avanza a presencial empresa" y typo "Desacartado"
    for (let i = 1; i < d.length; i++) {
      if (!d[i][1]) continue;
      let est = String(d[i][8]||'').trim();
      let nuevo = est;
      if (est === 'Avanza a empresa') nuevo = 'Avanza a presencial empresa';
      if (est === 'Desacartado NO le interesa') nuevo = 'Descartado NO le interesa';
      if (nuevo !== est) sh.getRange(i+1,9).setValue(nuevo);
    }
    // poner el desplegable en toda la columna 9 (desde fila 2)
    sh.getRange(2, 9, sh.getMaxRows()-1, 1).setDataValidation(regla);
    Logger.log(id + ' ok');
  });
  return true;
}

function notaEmpresa(perfilId, fila, etapa, nota) {
  const sh = getSheet(getConfig(perfilId));
  const mapa = {
    'screening-empresa': { A:'Avanza a presencial empresa', B:'Reserva screening empresa', C:'Descartado en screening empresa' },
    'presencial-empresa': { A:'Contratado', B:'Reserva presencial empresa', C:'Descartado presencial empresa' }
  };
  const estado = mapa[etapa] && mapa[etapa][nota];
  if (!estado) return false;
  const celda = sh.getRange(fila, 9);
  const r = celda.getDataValidation();
  celda.setDataValidation(null).setValue(estado);
  if (r) celda.setDataValidation(r);
  const colNota = (etapa === 'screening-empresa') ? 44 : 45;
  if (sh.getMaxColumns() < colNota) sh.insertColumnsAfter(sh.getMaxColumns(), colNota - sh.getMaxColumns());
  sh.getRange(fila, colNota).clearDataValidations().setValue(nota);
  return true;
}
function verDescartesEmpresa(){
  const perfiles=['coord-rrhh','analista-datos','analista-compras','analista-sgi'];
  perfiles.forEach(id=>{
    const sh=getSheet(getConfig(id));
    const rango=sh.getRange(2,9).getDataValidation();
    if(rango){
      const vals=rango.getCriteriaValues()[0];
      const emp=vals.filter(v=>/empresa/i.test(v));
      Logger.log(id+' → estados con "empresa": '+JSON.stringify(emp));
    }
  });
}
function contarNotaEmpresa(){
  Logger.log(typeof notaEmpresa);
}
function verAplicarEstado(){
  Logger.log(aplicarEstadoNota.toString().substring(0,600));
}
function verAplicarEstado2(){
  Logger.log(aplicarEstadoNota.toString().substring(600,1400));
}
function verEstadoReserva(){
  const perfiles=['coord-rrhh','analista-datos','analista-compras','analista-sgi'];
  perfiles.forEach(id=>{
    const sh=getSheet(getConfig(id));
    const d=sh.getDataRange().getValues();
    for(let i=1;i<d.length;i++){
      if(String(d[i][1]).toLowerCase().includes('capparelli')){
        Logger.log(id+' → "'+d[i][1]+'" → Estado: "'+d[i][8]+'"');
      }
    }
  });
}
function volverCapparelli(){
  const sh=getSheet(getConfig('coord-rrhh'));
  const d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++){
    if(String(d[i][1]).includes('Capparelli')){
      const celda=sh.getRange(i+1,9);
      const r=celda.getDataValidation();
      celda.setDataValidation(null).setValue('Avanza a presencial empresa');
      if(r)celda.setDataValidation(r);
      Logger.log('Listo: '+d[i][1]+' → Avanza a presencial empresa');
    }
  }
}
function borrarInformeBenvenuto(){
  const sh=getSheet(getConfig('analista-datos'));
  const d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++){
    if(String(d[i][1]).includes('Benvenuto')){
      sh.getRange(i+1,33).clearContent();
      Logger.log('Listo: borrado informe de '+d[i][1]);
    }
  }
}
function verTriggers(){
  ScriptApp.getProjectTriggers().forEach(t=>Logger.log(t.getHandlerFunction()+' | '+t.getEventType()));
}
function verWhatsApp(){
  const sh=getSheet(getConfig('analista-datos'));
  const d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++){
    if(!d[i][1])continue;
    const rt=sh.getRange(i+1,21).getRichTextValue();
    const link=rt?rt.getLinkUrl():'';
    const formula=sh.getRange(i+1,21).getFormula();
    Logger.log(d[i][1]+' → valor:"'+d[i][20]+'" link:"'+link+'" formula:"'+formula.substring(0,80)+'"');
    if(i>3)break;
  }
}
function verWhatsAppKanban(){
  const j=JSON.parse(doGetKanban('coord-rrhh'));
  const c=j.find(x=>x.whatsapp);
  Logger.log('WhatsApp link: '+(c?c.whatsapp:'ninguno'));
}
function verConversion(){
  const j=JSON.parse(doGetKanban('coord-rrhh'));
  const c=j.find(x=>x.whatsapp);
  const original=c.whatsapp;
  const convertido=original.replace('https://api.whatsapp.com/send?phone=','https://wa.me/').replace('&text=','?text=');
  Logger.log('CONVERTIDO: '+convertido);
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
function probarMailDescarte(){
  const mail='jaimeariel1980@gmail.com';
  const nombre='Ariel';
  const puesto='Analista de Datos';
  const empresa='MACS';
  const asunto='Blumia Consulting · Proceso de selección';

  const sinScr='Hola '+nombre+', ¿cómo estás?\n\n'+
    'Antes que nada, gracias por postularte a la búsqueda de '+puesto+' para '+empresa+'.\n\n'+
    'En esta oportunidad avanzamos con otros perfiles que se ajustaban más a lo que el cliente estaba buscando, así que no vamos a seguir adelante con tu candidatura en este proceso puntual.\n\n'+
    'Queremos que sepas que tu CV queda en nuestra base, y que ante futuras búsquedas que encajen con tu perfil, vamos a tenerte en cuenta.\n\n'+
    'Te deseamos mucho éxito en lo que venga. ¡Gracias por tu interés en trabajar con nosotros!\n\n'+
    'Saludos,\nEquipo de Blumia Consulting';

  const postScr='Hola '+nombre+', ¿cómo estás?\n\n'+
    'Ante todo, gracias por el tiempo que compartimos durante la entrevista y por tu buena predisposición a lo largo del proceso.\n\n'+
    'Después de evaluar a todos los candidatos, en esta ocasión decidimos avanzar con otros perfiles que se ajustaban más a los requerimientos específicos de la búsqueda de '+puesto+'. Así que no vamos a continuar con tu candidatura en este proceso.\n\n'+
    'Fue un gusto conocerte. Tu perfil queda en nuestra base y te vamos a tener presente para futuras búsquedas que puedan encajar mejor con tu experiencia.\n\n'+
    'Te deseamos mucho éxito en tus próximos pasos profesionales.\n\n'+
    'Saludos,\nEquipo de Blumia Consulting';

  GmailApp.sendEmail(mail, asunto+' (SIN screening)', sinScr);
  GmailApp.sendEmail(mail, asunto+' (POST screening)', postScr);
  Logger.log('Enviados a '+mail);
}
function enviarDescartesMasivo(perfilId, textoSin, textoPost) {
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
      enviados++;
    } catch(e) {}
  }
  return { ok:true, enviados: enviados, sinMail: sinMail };
}

function contarDescartesPendientes(perfilId) {
  const sh = getSheet(getConfig(perfilId));
  const d = sh.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < d.length; i++) {
    if (!d[i][1]) continue;
    const estado = String(d[i][8] || '').trim();
    if (!/descartado sin screening|descartado post screening/i.test(estado)) continue;
    if (String(d[i][39] || '').trim()) continue;
    const mail = String(d[i][21] || '').trim();
    if (!mail || mail === 'Sin dato' || mail.indexOf('@') === -1) continue;
    n++;
  }
  return n;
}
function probarMasivoSeguro(){
  const perfilId='coord-rrhh';
  const cfg=getConfig(perfilId);
  const sh=getSheet(cfg);
  const d=sh.getDataRange().getValues();
  const puesto=(typeof cfg.nombre==='string')?cfg.nombre:perfilId;
  let cuenta=0, lista=[];
  for(let i=1;i<d.length;i++){
    if(!d[i][1])continue;
    const estado=String(d[i][8]||'').trim();
    if(!/descartado sin screening|descartado post screening/i.test(estado))continue;
    if(String(d[i][39]||'').trim())continue;
    const mail=String(d[i][21]||'').trim();
    if(!mail||mail==='Sin dato'||mail.indexOf('@')===-1)continue;
    cuenta++;
    lista.push(d[i][1]+' ('+estado+') → '+mail);
  }
  GmailApp.sendEmail('jaimeariel1980@gmail.com','PRUEBA descarte masivo · '+puesto,
    'Se enviarían '+cuenta+' mails de descarte en '+puesto+':\n\n'+lista.join('\n'));
  Logger.log('Prueba enviada. Cantidad: '+cuenta);
}
function previaDescartes(perfilId) {
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
function verEntrevistador(){
  const sh=getSheet(getConfig('coord-rrhh'));
  const d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++){
    if(String(d[i][1]).toLowerCase().includes('APELLIDO')){
      Logger.log('Col14 (entrevistador): "'+d[i][13]+'"');
    }
  }
}
function verEntrevistador(){
  const sh=getSheet(getConfig('analista-compras'));
  const d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++){
    if(String(d[i][1]).toLowerCase().includes('APELLIDO')){
      Logger.log('Col14 (entrevistador): "'+d[i][13]+'"');
    }
  }
}
function verTodosEntrevistadores(){
  const perfiles=['coord-rrhh','analista-datos','analista-compras','analista-sgi'];
  perfiles.forEach(id=>{
    const sh=getSheet(getConfig(id));
    const d=sh.getDataRange().getValues();
    for(let i=1;i<d.length;i++){
      if(d[i][13]){
        Logger.log(id+' → '+d[i][1]+' → entrevistador: "'+d[i][13]+'"');
      }
    }
  });
}
function verLinkScreening(){
  const j=JSON.parse(doGetKanban('analista-sgi'));
  const c=j.find(x=>x.scrLink);
  Logger.log('Link screening: '+(c?c.scrLink:'ninguno'));
}
function verUltimo(){
  const sh=getSheet(getConfig('analista-sgi'));
  const d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++){
    if(String(d[i][1]).toLowerCase().includes('APELLIDO')){
      Logger.log('Col14: "'+d[i][13]+'"');
    }
  }
}
function verRomina(){
  const sh=getSheet(getConfig('coord-rrhh'));
  const d=sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++){
    if(String(d[i][1]).toLowerCase().includes('rosales')){
      Logger.log('Fila '+(i+1)+' → '+d[i][1]+' → col14: "'+d[i][13]+'" | estado: "'+d[i][8]+'"');
    }
  }
}
function verAlias(){
  Logger.log('Alias/cuentas para enviar: ' + GmailApp.getAliases().join(', '));
}
function probarRemitente(){
  GmailApp.sendEmail('jaimeariel1980@gmail.com', 'Prueba remitente Blumia', 'Este mail debería figurar como enviado desde empleos@blumiaconsulting.com', {from:'empleos@blumiaconsulting.com', name:'Blumia Consulting'});
  Logger.log('Enviado');
}
function verEmpresas(){
  const config = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config');
  const rows = config.getDataRange().getValues();
  const emp = {};
  for(let i=1;i<rows.length;i++){
    if(!rows[i][0]) continue;
    const e = rows[i][9] || 'Sin empresa';
    (emp[e]=emp[e]||[]).push(rows[i][0]+' ('+rows[i][1]+')');
  }
  Logger.log(JSON.stringify(emp, null, 2));
}
function empresasYPerfiles(){
  const config = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config');
  const rows = config.getDataRange().getValues();
  const emp = {};
  for(let i=1;i<rows.length;i++){
    if(!rows[i][0]) continue;
    const e = rows[i][9] || 'Sin empresa';
    (emp[e]=emp[e]||[]).push([rows[i][0], rows[i][1]]);
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
function cargarPlantillas(){
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let sh = ss.getSheetByName('Plantillas');
  if(!sh) sh = ss.insertSheet('Plantillas');
  sh.clear();
  sh.getRange(1,1,1,4).setValues([['clave','canal','link_calendar','texto']]);
  sh.getRange(1,1,1,4).setFontWeight('bold');

  const P = [];
  P.push(['contacto-screening','WhatsApp','',
'Hola [NOMBRE], ¿cómo estás? Soy [ENTREVISTADOR] de Blumia Consulting.\nTe contacto por la búsqueda de [PUESTO] para la empresa [EMPRESA].\nMe gustaría coordinar una charla breve con vos. Si te interesa, elegí el horario que mejor te quede en este link y ya nos queda agendado: [LINK]\nSaludos!']);

  P.push(['contacto-competencias','WhatsApp','',
'Hola [NOMBRE], ¿cómo estás? Soy Ariel de Blumia Consulting.\nTe contacto por la búsqueda de [PUESTO] para la empresa [EMPRESA].\nNos gustaría avanzar con tu perfil a la siguiente etapa del proceso. Se trata de una entrevista por Google Meet, con una duración aproximada de una hora, en la que profundizaremos sobre tu experiencia y recorrido profesional.\nSi te interesa continuar, te invito a elegir el día y horario que mejor te quede desde el siguiente enlace para dejar la entrevista agendada:\n[LINK]\n¡Nos vemos pronto!\nSaludos,\nAriel\nBlumia Consulting']);

  P.push(['contacto-empresa-virtual','WhatsApp','',
'Hola [NOMBRE], ¿cómo estás?\nQueríamos contarte que nos gustaría avanzar con tu perfil a la siguiente etapa del proceso para la posición de [PUESTO] en [EMPRESA].\nLa próxima instancia consiste en una entrevista virtual por Google Meet con referentes de la empresa, con una duración aproximada de 30 minutos. Será un primer encuentro para que puedas presentar tu experiencia, conocer al equipo y despejar las dudas que tengas sobre la empresa y la posición.\nSi te interesa continuar, te invitamos a elegir el día y horario que mejor te quede desde el siguiente enlace para dejar la entrevista agendada:\n[LINK]\nSaludos,\nAriel\nBlumia Consulting']);

  P.push(['contacto-empresa-presencial','WhatsApp','',
'Hola [NOMBRE], ¿cómo estás?\nQueríamos contarte que nos gustaría avanzar con tu perfil a la siguiente etapa del proceso para la posición de [PUESTO] en [EMPRESA].\nLa próxima instancia consiste en una entrevista presencial con referentes de la empresa, con una duración aproximada de una hora. Será una conversación en la que podrán profundizar sobre tu experiencia, conversar acerca de los desafíos de la posición y responder las consultas que puedan surgir de ambas partes.\nSi te interesa continuar, te invitamos a elegir el día y horario que mejor te quede desde el siguiente enlace para dejar la entrevista agendada:\n[LINK]\nSaludos,\nAriel\nBlumia Consulting']);

  P.push(['descarte-sin','Mail','',
'Hola [NOMBRE],\nQueremos agradecerte sinceramente por tu interés en la posición de [PUESTO] y por el tiempo dedicado a postularte.\nLuego de revisar cuidadosamente los curriculums recibidos, en esta oportunidad decidimos avanzar con candidatos cuyos perfiles se encuentran más alineados con los desafíos y requerimientos específicos de la posición.\nValoramos mucho tu trayectoria profesional y el interés en formar parte de este proceso. Conservaremos tu perfil para futuras oportunidades que puedan resultar acordes a tu experiencia.\nTe deseamos muchos éxitos en tus próximos desafíos profesionales.\nSaludos cordiales,\nEquipo de Selección de Blumia Consulting']);

  P.push(['descarte-post','Mail','',
'Hola [NOMBRE],\nMuchas gracias por el tiempo compartido durante el proceso de selección para la posición de [PUESTO].\nQueremos agradecer especialmente tu apertura y predisposición durante la conversación. Luego de analizar todas las informaciones recabadas hemos decidido avanzar con otros perfiles que presentan un ajuste más cercano a las necesidades actuales de la posición y del contexto organizacional.\nValoramos mucho tu experiencia y el intercambio mantenido, y esperamos poder considerar tu perfil para futuras oportunidades.\nTe deseamos muchos éxitos en tus próximos proyectos y desafíos profesionales.\nSaludos cordiales,\nEquipo de Selección de Blumia Consulting']);

  P.push(['descarte-competencias','Mail','',
'Hola [NOMBRE],\nAntes que nada, queremos agradecerte el tiempo que compartiste con nosotros durante las entrevistas para la posición de [PUESTO].\nValoramos mucho tu predisposición, la apertura con la que nos contaste tu experiencia y el interés demostrado a lo largo del proceso.\nLuego de evaluar integralmente los perfiles, en esta oportunidad hemos decidido continuar con candidatos que en principio presentan un ajuste más cercano a los desafíos específicos de la posición.\nAgradecemos sinceramente tu participación y esperamos poder considerar tu perfil en futuras oportunidades que se alineen mejor con tu experiencia.\nTe deseamos muchos éxitos en los próximos desafíos profesionales.\nSaludos cordiales,\nEquipo de Selección\nBlumia Consulting']);

  P.push(['descarte-empresa','Mail','',
'Hola [NOMBRE],\nAntes que nada, muchas gracias por el tiempo, el compromiso y la buena predisposición que demostraste durante todo el proceso de selección para la posición de [PUESTO].\nLuego de la entrevista con [EMPRESA], decidimos avanzar con otro perfil, principalmente porque [MOTIVO TÉCNICO].\nQueremos que sepas que nos resultó muy interesante conocerte y que tu perfil nos dejó una muy buena impresión. Por ese motivo, con tu autorización, mantendremos tus antecedentes en nuestra base para tenerte en cuenta en futuras búsquedas que se ajusten a tu experiencia y recorrido profesional.\nTe agradecemos nuevamente por haber confiado en Blumia Consulting y te deseamos muchos éxitos en los desafíos que vienen.\nSaludos,\nEquipo de Selección\nBlumia Consulting']);

  sh.getRange(2,1,P.length,4).setValues(P);
  sh.setColumnWidth(4, 600);
  sh.getRange(2,4,P.length,1).setWrap(true);
  Logger.log('Cargadas '+P.length+' plantillas');
}
function cargarLinks(){
  const config = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config');
  const rows = config.getDataRange().getValues();
  const H = rows[0];
  const cScr = H.indexOf('link_screening');
  const cComp = H.indexOf('link_competencias');
  const cEmp = H.indexOf('link_empresa_virtual');
  const links = {
    'coord-rrhh':['https://calendar.app.google/du9d2VuFZuoi6M6Q9','https://calendar.app.google/cuDHpykBPjam6cNN6','https://calendar.app.google/5YVWRVg61qB4uXTk9'],
    'analista-sgi':['https://calendar.app.google/QMmxZWwgwqqxbALPA','https://calendar.app.google/JW5Dwq71pcFcE4st8','https://calendar.app.google/QufPrSDQx6i7iDR8A'],
    'analista-compras':['https://calendar.app.google/T7RGSJAgLvazHSqG9','https://calendar.app.google/uV1zkjCFcKnt7cH27','https://calendar.app.google/gHpLoWyWKHWFs6qr5'],
    'analista-datos':['https://calendar.app.google/sCMqXshZNJqkoFmM6','https://calendar.app.google/YJdwHAXT7Ge41SxGA','https://calendar.app.google/BtLikgYjEUsbBZxu7']
  };
  for(let i=1;i<rows.length;i++){
    const id = rows[i][0];
    if(links[id]){
      config.getRange(i+1, cScr+1).setValue(links[id][0]);
      config.getRange(i+1, cComp+1).setValue(links[id][1]);
      config.getRange(i+1, cEmp+1).setValue(links[id][2]);
    }
  }
  Logger.log('Links cargados en las 4 búsquedas de MACS');
}
function crearColumnasLinks(){
  const config = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config');
  const rows = config.getDataRange().getValues();
  const headers = rows[0];
  const nuevas = ['link_screening','link_competencias','link_empresa_virtual','link_empresa_presencial'];
  let col = headers.length + 1;
  nuevas.forEach(function(n){
    if(headers.indexOf(n) === -1){
      config.getRange(1, col).setValue(n);
      col++;
    }
  });
  Logger.log('Columnas creadas. Encabezados: ' + config.getDataRange().getValues()[0].filter(String).join(' | '));
}
function verLinks(){
  ['coord-rrhh','analista-datos','analista-compras','analista-sgi'].forEach(id=>{
    const c = getConfig(id);
    Logger.log(id + ' → scr: ' + c.link_screening + ' | comp: ' + c.link_competencias + ' | emp: ' + c.link_empresa_virtual);
  });
}
function armarComunicacion(perfilId, fila, tipo){
  const cfg = getConfig(perfilId);
  const sh = getSheet(cfg);
  const nombre = String(sh.getRange(fila, 2).getValue() || '').split(' ')[0];
  const tel = String(sh.getRange(fila, 3).getValue() || '').replace(/\D/g,'');
  const entrevistador = String(sh.getRange(fila, 14).getValue() || 'Ariel');
  const puesto = (typeof cfg.nombre_perfil === 'string') ? cfg.nombre_perfil : perfilId;
  const empresa = empresaParaCandidatos_(cfg);

  const linkPorTipo = {
    'contacto-screening': cfg.link_screening,
    'contacto-competencias': cfg.link_competencias,
    'contacto-empresa-virtual': cfg.link_empresa_virtual,
    'contacto-empresa-presencial': cfg.link_empresa_virtual
  };
  const link = linkPorTipo[tipo] || '';

  const plantillas = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Plantillas').getDataRange().getValues();
  let texto = '';
  for(let i=1;i<plantillas.length;i++){
    if(plantillas[i][0] === tipo){ texto = plantillas[i][3]; break; }
  }
  if(!texto) return {ok:false, error:'No encontré la plantilla ' + tipo};

  texto = texto.replace(/\[NOMBRE\]/g, nombre)
               .replace(/\[PUESTO\]/g, puesto)
               .replace(/la empresa \[EMPRESA\]/g, empresa === 'MACS' ? 'la empresa MACS' : empresa)
               .replace(/\[EMPRESA\]/g, empresa)
               .replace(/\[ENTREVISTADOR\]/g, entrevistador)
               .replace(/\[LINK\]/g, link);

  let telWa = tel;
  if(telWa.indexOf('54') !== 0) telWa = '549' + telWa.replace(/^0+/,'');
  var etiqueta = {'contacto-screening':'screening','contacto-competencias':'competencias','contacto-empresa-virtual':'empresa','contacto-empresa-presencial':'empresa','descarte-competencias':'descarte comp.','descarte-empresa':'descarte emp.'}[tipo] || '';
  if(etiqueta){
    sh.getRange(fila, 41).setValue(etiqueta + ' · ' + Utilities.formatDate(new Date(), 'GMT-3', 'dd/MM/yyyy'));
    registrarContacto_(perfilId, fila, sh.getRange(fila, 2).getValue(), tel, 'WhatsApp: ' + etiqueta);
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
function enviarABanco(perfilId, fila, categoria, seniority, notaBanco){
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

  return {ok:true};
}
function arreglarLinksBanco(){
  const banco = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Banco');
  const d = banco.getDataRange().getValues();
  let n = 0;
  for(let i=1;i<d.length;i++){
    const v = String(d[i][9] || '').trim();
    if(v && v.indexOf('http') === -1 && v.length > 20){
      banco.getRange(i+1, 10).setFormula('=HYPERLINK("https://drive.google.com/file/d/' + v + '/view"; "Ver CV")');
      n++;
    }
  }
  Logger.log('Arreglados ' + n + ' links de CV');
}
function ampliarBanco(){
  const banco = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Banco');
  const H = banco.getDataRange().getValues()[0];
  const nuevas = ['screening_pdf','competencias_pdf','informe','postulacion','nota_cv','nota_scr','nota_comp','estado_origen'];
  let col = H.length + 1;
  nuevas.forEach(function(n){
    if(H.indexOf(n) === -1){
      banco.getRange(1, col).setValue(n).setFontWeight('bold').setBackground('#1A6645').setFontColor('#fff');
      col++;
    }
  });
  Logger.log('Columnas del Banco: ' + banco.getDataRange().getValues()[0].filter(String).join(' | '));
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
  const LOGO='https://lh3.googleusercontent.com/d/1WdoIWop2mZ_KnSa9raQM1Xb5TFP2JvGO';
  return `<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Banco de Talento · Blumia</title>
<style>
:root{--green:#1A6645;--glight:#E6F2EC;--amber:#F5C842;--violet:#F5C842;--bg:#F7F5F1;--line:#DDDBD7;--mid:#555;--font:'PP Pangram Sans','Plus Jakarta Sans',sans-serif}
*{box-sizing:border-box;margin:0;padding:0}body{font-family:var(--font);background:var(--bg);color:#1A1A1A;font-size:14px}
header{background:var(--violet);color:#fff;padding:16px 20px;display:flex;align-items:center;gap:14px;position:sticky;top:0;z-index:10;flex-wrap:wrap}
header img{height:34px}header h1{font-size:18px;font-weight:800}
.nav{display:flex;gap:8px;margin-left:auto;flex-wrap:wrap}
.nav button{padding:7px 14px;border-radius:6px;font-size:12px;font-weight:800;color:#1A1A1A;background:rgba(255,255,255,.55);border:none;cursor:pointer;font-family:var(--font)}
.filtros{padding:14px 16px;display:flex;gap:8px;flex-wrap:wrap;align-items:center;background:#fff;border-bottom:1px solid var(--line)}
.fwrap{position:relative;display:inline-block}
.fwrap svg{position:absolute;left:9px;top:50%;transform:translateY(-50%);opacity:.4;pointer-events:none}
.filtros input,.filtros select{padding:10px;border:1px solid var(--line);border-radius:8px;font-family:var(--font);font-size:13px;background:#fff}
.fwrap input{padding-left:30px}
.filtros button{padding:10px 14px;border:1px solid var(--line);border-radius:8px;background:#fff;cursor:pointer;font-family:var(--font);font-size:13px}
.wrap{padding:16px;display:grid;grid-template-columns:repeat(auto-fill,minmax(320px,1fr));gap:12px}
.p{background:#fff;border:1px solid var(--line);border-left:3px solid var(--violet);border-radius:8px;padding:14px;cursor:pointer}
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
function volver(){if(history.length>1){history.back();}else{window.top.location.href='${URL_WEBAPP}?tipo=kanban&perfil=coord-rrhh';}}
function borrarBanco(f,btn){
  var c=DATA.find(function(x){return x.fila===f;});
  if(!confirm('¿Quitar a '+c.nombre+' del banco de talento?\\nSigue estando en su búsqueda de origen.'))return;
  btn.textContent='Quitando...';btn.disabled=true;
  google.script.run.withSuccessHandler(function(r){
    if(r&&r.ok){btn.textContent='✓ Quitado';btn.style.background='#999';setTimeout(function(){window.top.location.href='${URL_WEBAPP}?tipo=banco';},600);}
    else{alert(r?r.error:'Error');btn.disabled=false;btn.textContent='🗑 Quitar del banco';}
  }).withFailureHandler(function(e){alert('Error: '+e.message);btn.disabled=false;btn.textContent='🗑 Quitar del banco';}).borrarDelBanco(f,c.nombre);
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
function volver(){if(history.length>1){history.back();}else{window.top.location.href='${URL_WEBAPP}?tipo=kanban&perfil=coord-rrhh';}}
function cerrar(){document.getElementById('modal').classList.remove('show');}
</script></body></html>`;
}

function bancoHTML(desde){
  return HtmlService.createHtmlOutput(BANCO_TPL(doGetBanco(), JSON.stringify(empresasYPerfiles()), desde || 'coord-rrhh'))
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport','width=device-width, initial-scale=1');
}
function verLinksBanco(){
  const b = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Banco');
  const d = b.getDataRange().getValues();
  const H = d[0];
  Logger.log('Columnas: ' + H.filter(String).join(' | '));
  for(let i=1;i<d.length && i<4;i++){
    Logger.log(d[i][1] + ' → scr:"' + b.getRange(i+1, H.indexOf('screening_pdf')+1).getFormula() + '" comp:"' + b.getRange(i+1, H.indexOf('competencias_pdf')+1).getFormula() + '" inf:"' + b.getRange(i+1, H.indexOf('informe')+1).getFormula() + '"');
  }
}
function borrarDelBanco(fila, nombre){
  const b = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Banco');
  const v = String(b.getRange(fila, 2).getValue() || '').trim();
  if(v.toLowerCase() !== String(nombre||'').toLowerCase().trim()){
    return {ok:false, error:'La fila no coincide con el nombre. Actualizá la página y probá de nuevo.'};
  }
  b.deleteRow(fila);
  return {ok:true};
}
function enviarABusqueda(filaBanco, perfilDestino){
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
function verGids(){
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  ss.getSheets().forEach(function(s){
    Logger.log(s.getName() + ' → gid: ' + s.getSheetId());
  });
}
function verFlyers(){
  const f = DriveApp.getFolderById('13sWdDTx-YJ5ocITn33MAfPfdVLxSaFzw');
  const it = f.getFiles();
  while(it.hasNext()){
    const a = it.next();
    Logger.log(a.getName() + ' → ' + a.getUrl());
  }
  const sub = f.getFolders();
  while(sub.hasNext()){
    const c = sub.next();
    Logger.log('[CARPETA] ' + c.getName() + ' → ' + c.getUrl());
  }
}
function cargarFlyers(){
  const config = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config');
  const rows = config.getDataRange().getValues();
  const H = rows[0];
  let c = H.indexOf('link_flyer');
  if(c === -1){ c = H.length; config.getRange(1, c+1).setValue('link_flyer'); }
  const F = {
    'analista-datos':'https://drive.google.com/file/d/1mEZ4KiuEgAIJFkRgFWVO5dVbuXP2VZPk/view',
    'analista-compras':'https://drive.google.com/file/d/1Izlncu72X5wYVivHnv1531NRc6rFFmzY/view',
    'analista-sgi':'https://drive.google.com/file/d/1o0cfBDTERNgA2EQ8cVFouma-jLSRgnhQ/view',
    'coord-rrhh':'https://drive.google.com/file/d/1f2UNmx7xLqx1i6EcqevNA4NJ-ZyTuUOx/view'
  };
  for(let i=1;i<rows.length;i++){
    if(F[rows[i][0]]) config.getRange(i+1, c+1).setValue(F[rows[i][0]]);
  }
  Logger.log('Flyers cargados en columna ' + (c+1));
}
function cargarPerfilesPuesto(){
  const config = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config');
  const rows = config.getDataRange().getValues();
  const H = rows[0];
  let c = H.indexOf('link_perfil');
  if(c === -1){ c = H.length; config.getRange(1, c+1).setValue('link_perfil'); }
  const P = {
    'analista-compras':'https://drive.google.com/file/d/1HA4DRQtSy0_4IOOvTXgeLr9bsEk6EPuz/view',
    'coord-rrhh':'https://drive.google.com/file/d/1706O7fEINVM-LKiVHAh_9uGGF9BuV7zJ/view',
    'analista-datos':'https://drive.google.com/file/d/17cTkRCpA0CHi2oue6jmhlSBSUsz-L3K8/view',
    'analista-sgi':'https://drive.google.com/file/d/1EdIMj2yeCaviqivihaOfQctIyDPyey3W/view'
  };
  for(let i=1;i<rows.length;i++){
    if(P[rows[i][0]]) config.getRange(i+1, c+1).setValue(P[rows[i][0]]);
  }
  Logger.log('Perfiles de puesto cargados en columna ' + (c+1));
}
function borrarMarcaContacto(perfilId, fila){
  const sh = getSheet(getConfig(perfilId));
  sh.getRange(fila, 41).clearContent();
  return {ok:true};
}
function verCitaVieja(){
  const sh = getSheet(getConfig('coord-rrhh'));
  const d = sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++){
    if(String(d[i][1]).toLowerCase().includes('bodean')){
      Logger.log('Estado (col9): "'+d[i][8]+'"');
      Logger.log('Cita (col37): "'+d[i][36]+'"');
      Logger.log('Historial (col38): "'+d[i][37]+'"');
    }
  }
}
function verCitaVieja(){
  ['coord-rrhh','analista-datos','analista-compras','analista-sgi','operaciones'].forEach(function(id){
    const sh = getSheet(getConfig(id));
    const d = sh.getDataRange().getValues();
    for(let i=1;i<d.length;i++){
      if(String(d[i][1]).toLowerCase().includes('bodean')){
        Logger.log(id+' → estado:"'+d[i][8]+'" cita:"'+d[i][36]+'" hist:"'+d[i][37]+'"');
      }
    }
  });
}
function verNotasEmpresa(){
  const sh = getSheet(getConfig('analista-compras'));
  const d = sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++){
    if(/empresa|contratado|reserva/i.test(String(d[i][8]))){
      Logger.log(d[i][1]+' → estado:"'+d[i][8]+'"');
    }
  }
}
function verPatricio(){
  const sh = getSheet(getConfig('analista-compras'));
  const d = sh.getDataRange().getValues();
  for(let i=1;i<d.length;i++){
    if(String(d[i][1]).toLowerCase().includes('patricio')){
      Logger.log('estado: "'+d[i][8]+'"');
      Logger.log('notaScr col12: "'+d[i][11]+'"');
      Logger.log('notaComp col16: "'+d[i][15]+'"');
      Logger.log('scrEmp col44: "'+d[i][43]+'"');
      Logger.log('presEmp col45: "'+d[i][44]+'"');
    }
  }
}
function chequearSistema(){
  const fns = ['notaEmpresa','enviarABanco','armarComunicacion','borrarMarcaContacto','doGetKanban','KANBAN_TPL'];
  fns.forEach(function(f){
    Logger.log(f + ': ' + (typeof this[f] === 'function' ? 'OK' : 'FALTA'));
  });
}
function verTelBanco(){
  const b = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Banco');
  const d = b.getDataRange().getValues();
  const H = d[0];
  const c = H.indexOf('tel');
  for(let i=1;i<d.length;i++){
    if(String(d[i][1]).toLowerCase().includes('maza')){
      Logger.log('valor: "'+d[i][c]+'" | formula: "'+b.getRange(i+1,c+1).getFormula()+'"');
    }
  }
}
function arreglarTelBanco(){
  const b = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Banco');
  const d = b.getDataRange().getValues();
  const c = d[0].indexOf('tel');
  let n = 0;
  for(let i=1;i<d.length;i++){
    const cel = b.getRange(i+1, c+1);
    const f = cel.getFormula();
    if(f && f.indexOf('+') === 0){
      cel.setNumberFormat('@').setValue(f);
      n++;
    }
  }
  Logger.log('Arreglados ' + n + ' teléfonos');
}
