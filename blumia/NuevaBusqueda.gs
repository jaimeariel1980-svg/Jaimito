/* ============================================================
   NUEVA BÚSQUEDA DESDE UN WORD (⚙ Configurar → ➕ Nueva búsqueda)
   1) El navegador lee el texto del Word (no se guarda el archivo).
   2) proponerBusquedaIA_: Claude arma la propuesta completa (solo propone, no crea nada).
   3) La persona la revisa y la corrige en pantalla.
   4) crearBusquedaIA_: crea carpetas, pestaña de candidatos, Config, Hero, screening,
      competencias y preguntas de postulación (mismo resultado que crearPerfilDirectorComercial).
   ============================================================ */

const CARPETA_MADRE_BUSQUEDAS_ = '18bhn7cD2PkcK_SldSfKHrjiieHvO3tYH';
const ESTADOS_BUSQUEDA_ = ['Pendiente', 'Nuevo', 'Contactado', 'Avanza a screening', 'Screening agendado', 'Avanza a competencias', 'Competencias agendada',
  'Avanza a screening empresa', 'Screening empresa agendado', 'Avanza a presencial empresa', 'Entrevista empresa agendada',
  'En cliente', 'Contratado', 'En reserva', 'Descartado sin screening', 'Descartado post screening', 'Descartado en competencias',
  'Descartado en screening empresa', 'Descartado presencial empresa', 'Descartado Fuera de zona', 'Descartado NO le interesa'];

function slugBusqueda_(s) {
  return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 40);
}

function promptNuevaBusqueda_(empresas) {
  return 'Sos consultor/a senior de selección de personal en Blumia Consulting (Argentina). Te paso el texto de un relevamiento o descripción de un puesto de un cliente. Armá la configuración COMPLETA de la búsqueda para el sistema de reclutamiento. Usá SOLO lo que dice el documento; lo que falte, dejalo breve y razonable pero marcá en "avisos" qué dato no estaba y tendrían que confirmar. No inventes cifras, nombres de empresas ni requisitos excluyentes que el documento no mencione.\n\n' +
    'Respondé SOLO con un JSON válido (sin texto antes ni después, sin markdown) con esta forma exacta:\n' +
    '{"nombre":"nombre del puesto, corto (ej: Director Comercial de Proyectos de Consultoría)",' +
    '"empresa":"empresa o sector cliente, corto (' + (empresas.length ? 'si encaja, usá una de estas: ' + empresas.join(' | ') : 'ej: Energía y Minería') + ')",' +
    '"ubicacion":"ciudad/zona y modalidad, corto",' +
    '"banda_salarial":"solo dígitos en pesos si el documento da un monto de referencia, si no cadena vacía",' +
    '"prompt_extra":"texto para que la IA puntúe CVs: PERFIL (misión, reporta a), EXPERIENCIA (EXCLUYENTE) , FORMACIÓN, CONOCIMIENTOS, IDIOMAS, ZONA, BANDERAS ROJAS, CONTINUIDAD LABORAL y EVALUACIÓN DE BORDERLINE (incluir y señalar si hay factores compensatorios), y al final NIVEL_AJUSTE con líneas \'- Alto: …\', \'- Medio: …\', \'- Bajo: …\'. En párrafos con esos títulos en mayúscula, en español rioplatense, entre 1500 y 4500 caracteres",' +
    '"screening_eval":[{"criterio":"criterio a evaluar en el screening","peso":1 a 3}] (4 a 6, los más importantes con peso 3),' +
    '"screening_val":[{"pregunta":"dato a validar con el candidato (remuneración, disponibilidad, modalidad, preaviso, interés, idiomas, etc.)","opciones":"Sí|No|A revisar"}] (4 a 7, opciones separadas con |),' +
    '"competencias":[{"nombre":"competencia","descripcion":"una oración","peso":1 a 3,"preguntas":["pregunta STAR 1","pregunta STAR 2","pregunta STAR 3"]}] (4 a 6, con 3 preguntas de entrevista por competencia que pidan ejemplos concretos),' +
    '"postulacion":[{"etiqueta":"pregunta del formulario de postulación","tipo":"texto o select","opciones":"opciones separadas con | si es select, si no \'\'","requerido":"SI o NO"}] (4 a 6, enfocadas en los requisitos excluyentes; con "select" cuando se pueda),' +
    '"avisos":["datos que faltaban o supuestos que hay que confirmar"]}';
}

function empresasConocidas_() {
  try {
    const rows = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getDataRange().getValues();
    const c = rows[0].indexOf('empresa');
    if (c < 0) return [];
    const s = {};
    rows.slice(1).forEach(r => { const e = String(r[c] || '').trim(); if (e) s[e] = 1; });
    return Object.keys(s).slice(0, 12);
  } catch (e) { return []; }
}

function idLibre_(base) {
  const ids = {};
  SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getDataRange().getValues().slice(1).forEach(r => { ids[String(r[0])] = 1; });
  let id = base || 'busqueda', n = 2;
  while (ids[id]) id = (base || 'busqueda') + '-' + n++;
  return id;
}

function listaEntera_(v, def) { return Math.max(1, Math.min(3, Math.round(Number(v) || def))); }

function normalizarPropuesta_(r) {
  const t = (v, n) => limpiarTxt_(v).slice(0, n);
  const p = {
    nombre: t(r.nombre, 120), empresa: t(r.empresa, 80), ubicacion: t(r.ubicacion, 120),
    banda_salarial: String(r.banda_salarial || '').replace(/\D/g, '').slice(0, 12),
    prompt_extra: limpiarTxt_(r.prompt_extra).slice(0, 12000),
    screening_eval: (Array.isArray(r.screening_eval) ? r.screening_eval : []).slice(0, 8).map(e => ({ criterio: t(e.criterio, 200), peso: listaEntera_(e.peso, 1) })).filter(e => e.criterio),
    screening_val: (Array.isArray(r.screening_val) ? r.screening_val : []).slice(0, 10).map(v => ({ pregunta: t(v.pregunta, 200), opciones: String(v.opciones || '').split('|').map(o => t(o, 40)).filter(Boolean).join('|') || 'Sí|No' })).filter(v => v.pregunta),
    competencias: (Array.isArray(r.competencias) ? r.competencias : []).slice(0, 8).map(c => ({
      nombre: t(c.nombre, 120), descripcion: t(c.descripcion, 400), peso: listaEntera_(c.peso, 1),
      preguntas: (Array.isArray(c.preguntas) ? c.preguntas : String(c.preguntas || '').split('|')).map(q => t(q, 300)).filter(Boolean).slice(0, 6)
    })).filter(c => c.nombre),
    postulacion: (Array.isArray(r.postulacion) ? r.postulacion : []).slice(0, 8).map(q => {
      const tipo = String(q.tipo).toLowerCase() === 'select' ? 'select' : 'texto';
      const ops = tipo === 'select' ? String(q.opciones || '').split('|').map(o => t(o, 60)).filter(Boolean).join('|') : '';
      return { etiqueta: t(q.etiqueta, 300), tipo: ops || tipo === 'texto' ? tipo : 'texto', opciones: ops, requerido: q.requerido === 'NO' ? 'NO' : 'SI' };
    }).filter(q => q.etiqueta),
    avisos: (Array.isArray(r.avisos) ? r.avisos : []).slice(0, 8).map(a => t(a, 250)).filter(Boolean)
  };
  return p;
}

function proponerBusquedaIA_(pin, texto, empresaSug) {
  pinValido_(pin);
  texto = String(texto || '').replace(/\u0000/g, '').trim();
  if (texto.length < 200) throw new Error('El texto es muy corto para armar una búsqueda (pegá o subí el relevamiento completo).');
  const empresas = empresasConocidas_();
  const r = extraerJSON_(procesarConClaude(promptNuevaBusqueda_(empresas), texto.slice(0, 60000), 'analisis'));
  const p = normalizarPropuesta_(r);
  const sug = limpiarTxt_(empresaSug).slice(0, 80);
  if (sug) p.empresa = sug;
  return p;
}

function validarPropuesta_(p) {
  if (!p || typeof p !== 'object') throw new Error('Falta la propuesta.');
  if (!p.nombre) throw new Error('Falta el nombre del puesto.');
  if (!p.empresa) throw new Error('Falta la empresa.');
  if (String(p.prompt_extra || '').length < 200) throw new Error('Los requisitos para la IA son muy cortos (mínimo 200 caracteres).');
  if (!/NIVEL_AJUSTE/.test(p.prompt_extra)) throw new Error('Los requisitos para la IA tienen que terminar con el bloque NIVEL_AJUSTE (Alto / Medio / Bajo).');
  if (p.screening_eval.length < 2) throw new Error('Se necesitan al menos 2 criterios de screening.');
  if (p.competencias.length < 2) throw new Error('Se necesitan al menos 2 competencias.');
  p.competencias.forEach(c => { if (!c.preguntas.length) throw new Error('La competencia «' + c.nombre + '» no tiene preguntas.'); });
  if (!p.postulacion.length) throw new Error('Se necesita al menos una pregunta de postulación.');
}

function crearBusquedaIA_(pin, prop, quien) {
  pinValido_(pin);
  quien = limpiarTxt_(quien) || 'sin nombre';
  const p = normalizarPropuesta_(prop || {});
  validarPropuesta_(p);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
    const id = idLibre_(slugBusqueda_(p.nombre));
    const hojaNombre = p.nombre.replace(/[\[\]\*\?\:\/\\]/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 60);
    if (ss.getSheetByName(hojaNombre)) throw new Error('Ya existe una pestaña llamada «' + hojaNombre + '». Cambiá el nombre del puesto.');
    const config = ss.getSheetByName('Config');
    const H = config.getDataRange().getValues()[0];

    const carpetaMadre = DriveApp.getFolderById(CARPETA_MADRE_BUSQUEDAS_);
    const itC = carpetaMadre.getFoldersByName(p.empresa);
    const carpetaEmpresa = itC.hasNext() ? itC.next() : carpetaMadre.createFolder(p.empresa);
    const carpetaPerfil = carpetaEmpresa.createFolder(p.nombre);
    const fNuevos = carpetaPerfil.createFolder('CVs nuevos');
    const fProc = carpetaPerfil.createFolder('CVs procesados');
    const fScr = carpetaPerfil.createFolder('Screening PDFs');
    const fComp = carpetaPerfil.createFolder('Competencias PDFs');

    const hoja = ss.insertSheet(hojaNombre);
    const headers = ['Fecha', 'Nombre', 'Teléfono', 'DNI', 'Nivel de Ajuste', 'Requisitos', 'Resumen Experiencia', 'Link CV', 'Estado', 'Fuente', 'Screening', 'Impresión', 'Detalle', 'Entrevistó', 'Entrevista competencias', 'Impresión', 'Detalle', 'Entrevistó', 'Entrevista Compañia', 'Impresión', 'Contactar', 'Correo'];
    hoja.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
    hoja.setFrozenRows(1);
    hoja.autoResizeColumns(1, headers.length);
    formatearHojaBusqueda_(ss, id, hoja);

    const valores = {
      perfil_id: id, nombre_perfil: p.nombre, nombre: p.nombre, sheet_id: SHEET_MASTER_ID, hoja_nombre: hojaNombre,
      folder_madre: fNuevos.getId(), folder_recibidos: fProc.getId(), folder_pdf_screening: fScr.getId(), folder_pdf_competencias: fComp.getId(),
      prompt_extra: p.prompt_extra, empresa: p.empresa
    };
    if (p.banda_salarial && Number(p.banda_salarial) >= 1000) valores.banda_salarial = p.banda_salarial;
    ['banda_salarial'].forEach(k => { if (valores[k] !== undefined && H.indexOf(k) === -1) { H.push(k); config.getRange(1, H.length).setValue(k); } });
    config.appendRow(H.map(h => valores[h] !== undefined ? valores[h] : ''));

    ss.getSheetByName('Perfiles_Hero').appendRow([id, p.empresa, p.nombre, p.ubicacion, 'Screening · 1ª entrevista', 'Ariel, Analía, Flavia, Marcos', '8.5', '6.5']);
    const ev = ss.getSheetByName('Perfiles_Screening_Eval');
    p.screening_eval.forEach((e, i) => ev.appendRow([id, 'e' + (i + 1), e.criterio, e.peso]));
    const va = ss.getSheetByName('Perfiles_Screening_Val');
    p.screening_val.forEach((v, i) => va.appendRow([id, 'v' + (i + 1), v.pregunta, v.opciones]));
    const co = ss.getSheetByName('Perfiles_Competencias');
    p.competencias.forEach((c, i) => co.appendRow([id, 'c' + (i + 1), c.nombre, c.descripcion, c.peso, c.preguntas.join('|')]));
    const pp = ss.getSheetByName('Postulacion_Preguntas');
    const filas = p.postulacion.map((q, i) => [id, 'p' + (i + 1), q.etiqueta, q.tipo, q.opciones, q.requerido]);
    pp.getRange(pp.getLastRow() + 1, 1, filas.length, 6).setValues(filas);

    logForm_(ss, quien, id, 'Búsqueda creada desde Word + IA', { nombre: p.nombre, empresa: p.empresa });
    try { registrarHistorial_(hojaNombre, '(búsqueda)', 'Búsqueda creada', 'Creada desde un Word con IA por ' + quien, quien); } catch (e) { }
    return { id: id, nombre: p.nombre, empresa: p.empresa, kanban: URL_DEPLOY + '?tipo=kanban&perfil=' + id, postulacion: URL_DEPLOY + '?tipo=postulacion&perfil=' + id, carpeta: carpetaPerfil.getUrl() };
  } finally { lock.releaseLock(); }
}

/* Desplegables (estado, notas, entrevistadores) y links «Iniciar Screening» / «Evaluar competencias» de la pestaña de una búsqueda nueva. */
function formatearHojaBusqueda_(ss, id, hoja) {
  const modelo = ss.getSheetByName('Supervisor de Operaciones');
  if (modelo) for (let c = 1; c <= Math.min(modelo.getLastColumn(), hoja.getMaxColumns()); c++) hoja.setColumnWidth(c, modelo.getColumnWidth(c));
  const lista = l => SpreadsheetApp.newDataValidation().requireValueInList(l, true).setAllowInvalid(true).build();
  const entrevistadores = ['Ariel', 'Marcos', 'Ever', 'Analía', 'Flavia'];
  hoja.getRange(2, 9, 250, 1).setDataValidation(lista(ESTADOS_BUSQUEDA_));
  hoja.getRange(2, 12, 250, 1).setDataValidation(lista(['A', 'B', 'C']));
  hoja.getRange(2, 16, 250, 1).setDataValidation(lista(['A', 'B', 'C']));
  hoja.getRange(2, 14, 250, 1).setDataValidation(lista(entrevistadores));
  hoja.getRange(2, 18, 250, 1).setDataValidation(lista(entrevistadores));
  const fScr = [], fComp = [];
  for (let i = 2; i <= 250; i++) {
    fScr.push(['=IF(B' + i + '<>"",HYPERLINK("' + URL_DEPLOY + '?perfil=' + id + '&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(N' + i + '),"Iniciar Screening"),"")']);
    fComp.push(['=IF(B' + i + '<>"",HYPERLINK("' + URL_DEPLOY + '?tipo=competencias&perfil=' + id + '&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(R' + i + '),"Evaluar competencias"),"")']);
  }
  hoja.getRange(2, 11, fScr.length, 1).setFormulas(fScr);
  hoja.getRange(2, 15, fComp.length, 1).setFormulas(fComp);
}
