/**
 * FORMULARIO DE PRECARGA — AZCUÉNAGA INMOBILIARIA
 * Backend (Código.gs)
 *
 * Responsabilidades:
 *  - Servir el formulario HTML (doGet)
 *  - Recibir envíos parciales o completos (doPost / precargarDatos)
 *  - Generar y resolver códigos de precarga
 *  - Crear/recuperar la carpeta de Drive de cada propiedad
 *  - Guardar archivos con nomenclatura estándar
 *  - Calcular el estado "presentado"/"pendiente" de cada documento
 *  - Enviar el mail de confirmación de precarga
 */

// ---------------------------------------------------------------------------
// CONFIGURACIÓN
// ---------------------------------------------------------------------------

const CONFIG = {
  SHEET_ID: '1O8VKF9ZUC073kGNYIc3SgKPcIrVpnvynpQQ69RAk9W8',
  SHEET_NAME: 'Ventas - Base de Datos Maestra',
  DRIVE_ROOT_FOLDER_ID: '1AGXKG5UN_GNlrosBojgNHhqDBaL9Z12m',
  REMITENTE_NOMBRE: 'Azcuénaga Inmobiliaria',
  MAX_DUEÑOS: 5,

  // Documentos requeridos para poder hacer el PRIMER envío de precarga
  // (vacío: toda la documentación es opcional, se puede enviar sin adjuntar nada)
  DOCUMENTOS_OBLIGATORIOS_PRECARGA: [],

  // Catálogo completo de documentos que puede tener una propiedad
  DOCUMENTOS: [
    { key: 'plano_edificacion', label: 'Plano de edificación', obligatorio: false },
    { key: 'plano_mensura',     label: 'Plano de mensura',     obligatorio: false },
    { key: 'escritura',         label: 'Escritura',            obligatorio: false },
    { key: 'dni_propietario',   label: 'DNI propietario',      obligatorio: false },
    { key: 'cuil_propietario',  label: 'CUIL Propietario',     obligatorio: false },
    { key: 'titulo_propiedad',  label: 'Título de propiedad',  obligatorio: false },
    { key: 'tgi_municipal',     label: 'TGI Municipal',        obligatorio: false },
    { key: 'api_provincial',    label: 'API Provincial',       obligatorio: false },
    { key: 'epe_luz',           label: 'EPE Luz',              obligatorio: false },
    { key: 'aguas_santafesinas',label: 'Aguas Santafesinas',   obligatorio: false },
    { key: 'litoral_gas',       label: 'Litoral Gas',          obligatorio: false },
    { key: 'autorizacion_venta',label: 'Autorización de venta',obligatorio: false }
  ],

  // Campos personales/de propiedad mínimos para el primer envío de precarga
  CAMPOS_OBLIGATORIOS_PRECARGA: [
    'dueño1_nombre', 'dueño1_dni', 'dueño1_fechaNac', 'dueño1_nacionalidad',
    'dueño1_domicilio', 'dueño1_celular', 'dueño1_email',
    'calle', 'numero', 'pasillo', 'entrecalle1', 'entrecalle2', 'ciudad'
  ]
};

// Documentos opcionales de DNI/CUIL para dueños adicionales (2 a MAX_DUEÑOS).
// Los del dueño 1 (dni_propietario / cuil_propietario) son los únicos obligatorios.
for (let n = 2; n <= CONFIG.MAX_DUEÑOS; n++) {
  CONFIG.DOCUMENTOS.push({ key: 'dni_dueño' + n,  label: 'DNI Dueño ' + n,  obligatorio: false });
  CONFIG.DOCUMENTOS.push({ key: 'cuil_dueño' + n, label: 'CUIL Dueño ' + n, obligatorio: false });
}

// Encabezados fijos del Sheet (además de las columnas dinámicas por dueño y por documento)
const COLUMNAS_META = [
  'Código de Precarga', 'Fecha de creación', 'Última actualización'
];

const COLUMNAS_PROPIEDAD = [
  'Calle', 'Número', 'Departamento', 'Piso', 'Pasillo', 'Entrecalle 1', 'Entrecalle 2', 'Ciudad',
  'Tipo de propiedad', 'Superficie Total (m2)', 'Superficie Cubierta (m2)',
  'Dormitorios', 'Ambientes (sin dormitorios)', 'Baños', 'Plantas',
  'Cochera', 'Patio delantero', 'Patio trasero', 'Gas Natural', 'Cloaca',
  'Antigüedad (años)', 'Orientación', 'Inscripción del dominio',
  'Carpeta Drive (URL)', 'Estado de la propiedad', '% Completitud',
  'Última alerta de documentación enviada'
];

// Campos de cada dueño/titular. La clave (nombre, dni, ...) debe coincidir
// con el sufijo usado en el formulario: dueño{N}_{clave}
const CAMPOS_DUEÑO = [
  { clave: 'nombre',       columna: 'Nombre y Apellido' },
  { clave: 'dni',          columna: 'DNI' },
  { clave: 'fechaNac',     columna: 'Fecha de Nacimiento' },
  { clave: 'nacionalidad', columna: 'Nacionalidad' },
  { clave: 'domicilio',    columna: 'Domicilio' },
  { clave: 'pasillo',      columna: 'Pasillo' },
  { clave: 'celular',      columna: 'Celular' },
  { clave: 'email',        columna: 'E-mail' }
];

function _columnasDueños() {
  const columnas = [];
  for (let n = 1; n <= CONFIG.MAX_DUEÑOS; n++) {
    CAMPOS_DUEÑO.forEach(c => columnas.push('Dueño ' + n + ' - ' + c.columna));
  }
  return columnas;
}

// ---------------------------------------------------------------------------
// ENTRADA WEB
// ---------------------------------------------------------------------------

/**
 * EJECUTAR UNA SOLA VEZ, A MANO, desde el editor de Apps Script (▶ Ejecutar),
 * eligiendo esta función en el desplegable. Reescribe la fila 1 de encabezados
 * con la estructura completa (incluidas las columnas de dueños que faltaban).
 *
 * IMPORTANTE: si ya tenés filas de datos cargadas con la estructura vieja,
 * borralas antes de correr esto (los datos quedarían desalineados, ya que
 * las columnas de dueños se insertan antes que las de la propiedad). Si es
 * solo la fila de prueba, borrala y volvé a enviar el formulario después.
 */
function regenerarEncabezados() {
  const sheet = _getSheet();
  const headers = _encabezadosCompletos();
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).clearContent();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}

/**
 * DIAGNÓSTICO — correr a mano desde el editor (▶ Ejecutar, eligiendo esta
 * función). No depende del despliegue del Web App: prueba la búsqueda
 * directamente contra la hoja y deja el detalle en el Registro de ejecución
 * (Ver → Registros de ejecución).
 *
 * Editá el valor de PROBAR_CODIGO por el código que estás buscando.
 */
function debugBuscarPrecarga() {
  const PROBAR_CODIGO = 'PRE-2026-0003'; // <-- cambiá esto por el código a probar

  const sheet = _getSheet();
  Logger.log('Hoja usada: "%s" (Sheet ID: %s)', sheet.getName(), CONFIG.SHEET_ID);

  const lastRow = sheet.getLastRow();
  Logger.log('Última fila con datos: %s', lastRow);

  const objetivo = _normalizarCodigo(PROBAR_CODIGO);
  Logger.log('Buscando: "%s" → normalizado: "%s"', PROBAR_CODIGO, objetivo);

  const valoresColumnaA = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(v => v[0]);
  Logger.log('Códigos existentes en columna A: %s', JSON.stringify(valoresColumnaA));
  Logger.log('Normalizados: %s', JSON.stringify(valoresColumnaA.map(_normalizarCodigo)));

  const fila = _buscarFilaPorCodigo(sheet, PROBAR_CODIGO);
  Logger.log('Resultado _buscarFilaPorCodigo: %s', fila);
}

function doGet(e) {
    if (e && e.parameter && e.parameter.azr) return azcuPoll(e);
    if (e && e.parameter && e.parameter.app === 'encuesta') {
    var te=HtmlService.createTemplateFromFile('Encuesta');
    te.codigo = e.parameter.prop || '';
    return te.evaluate().setTitle('Encuesta de Satisfacción — Azcuénaga')
      .addMetaTag('viewport','width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (e && e.parameter && e.parameter.app === 'documentos') {
    return HtmlService.createHtmlOutputFromFile('Documentos')
      .setTitle('Generador de documentos — Azcuénaga')
      .addMetaTag('viewport','width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (e && e.parameter && e.parameter.app === 'tablero') {
    return HtmlService.createHtmlOutputFromFile('Tablero')
      .setTitle('Tablero de propiedades — Azcuénaga')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  if (e && e.parameter && e.parameter.app === 'procedimiento') {
    return HtmlService.createHtmlOutputFromFile('ProcedimientoVentas')
      .setTitle('Procedimiento de Ventas — Azcuénaga')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }
  const codigoPrecarga = e && e.parameter && e.parameter.precarga ? e.parameter.precarga : '';
  const template = HtmlService.createTemplateFromFile('Precarga');
  template.codigoPrecarga = codigoPrecarga;
  template.datosPrevios = codigoPrecarga ? obtenerDatosPorCodigo(codigoPrecarga) : null;
  return template.evaluate()
    .setTitle('Formulario de Precarga — Azcuénaga Inmobiliaria')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Llamada desde el cliente (google.script.run) al enviar el formulario,
 * completo o parcial. Devuelve { ok, codigoPrecarga, mensaje }.
 */

function precargarDatos(datos, archivosBase64) {
  const sheet = _getSheet();
  const esNuevo = !datos.codigoPrecarga;
  const archivos = _archivosDesdeBase64(archivosBase64);

  if (esNuevo) {
    _validarCamposObligatoriosPrimerEnvio(datos, archivos);
    datos.codigoPrecarga = _generarCodigoPrecarga(sheet);
  }

  const carpeta = _getOrCreatePropertyFolder(datos.codigoPrecarga, datos.calle, datos.numero);
  const urlsGuardadas = _guardarArchivos(carpeta, datos.codigoPrecarga, datos.calle, datos.numero, archivos);
  const estadoDocs = _calcularEstadoDesdeCarpeta(carpeta);
  const verificacionIA = _verificarArchivosConIA(archivos, _nombresDueños(datos));

  _guardarFila(sheet, datos, carpeta.getUrl(), estadoDocs, verificacionIA);

  if (esNuevo) {
    _enviarMailConfirmacionPrecarga(datos);
  }

  return {
    ok: true,
    codigoPrecarga: datos.codigoPrecarga,
    archivosGuardados: urlsGuardadas
  };
}

/**
 * El cliente manda cada archivo como { base64, mimeType, filename }
 * (ver Precarga.html). Acá se reconstruye como Blob de Apps Script.
 */
function _archivosDesdeBase64(archivosBase64) {
  const archivos = {};
  if (!archivosBase64) return archivos;

  Object.keys(archivosBase64).forEach(key => {
    const info = archivosBase64[key];
    if (!info || !info.base64) return;
    const bytes = Utilities.base64Decode(info.base64);
    archivos[key] = Utilities.newBlob(bytes, info.mimeType, info.filename);
  });

  return archivos;
}

function obtenerDatosPorCodigo(codigoPrecarga) {
  const sheet = _getSheet();
  const fila = _buscarFilaPorCodigo(sheet, codigoPrecarga);
  if (!fila) return null;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const valores = sheet.getRange(fila, 1, 1, sheet.getLastColumn()).getValues()[0];
  const registro = {};
  headers.forEach((h, i) => {
    registro[h] = _valorSerializable(valores[i]);
  });
  return registro;
}

/**
 * google.script.run puede fallar en silencio si se le manda un objeto Date
 * (u otro tipo no plano) desde el servidor. Convertimos todo a texto/número
 * antes de devolverlo al cliente.
 */
function _valorSerializable(valor) {
  if (valor instanceof Date) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  }
  return valor;
}

/**
 * Une la búsqueda y, si no encuentra nada, el diagnóstico, en una sola
 * ejecución del servidor (evita depender de dos llamadas separadas).
 */
function obtenerDatosOFallar(codigoPrecarga) {
  const registro = obtenerDatosPorCodigo(codigoPrecarga);
  if (registro) return { encontrado: true, registro: registro };
  return { encontrado: false, diagnostico: diagnosticoBusquedaWeb(codigoPrecarga) };
}

/**
 * TEMPORAL — diagnóstico detallado, usado por obtenerDatosOFallar cuando no
 * se encuentra el código. Se puede borrar una vez resuelto el problema.
 */
function diagnosticoBusquedaWeb(codigoPrecarga) {
  const sheet = _getSheet();
  const lastRow = sheet.getLastRow();
  const objetivo = _normalizarCodigo(codigoPrecarga);
  const valoresColumnaA = lastRow < 2 ? [] : sheet.getRange(2, 1, lastRow - 1, 1).getValues().map(v => v[0]);

  const codigosDeCoincidenciaAproximada = valoresColumnaA.filter(v => _normalizarCodigo(v).replace(/-/g, '') === objetivo.replace(/-/g, ''));

  return {
    sheetId: CONFIG.SHEET_ID,
    sheetName: sheet.getName(),
    lastRow: lastRow,
    codigoRecibido: codigoPrecarga,
    objetivo: objetivo,
    codigosCharCodes: Array.from(codigoPrecarga).map(c => c.charCodeAt(0)),
    codigosExistentes: valoresColumnaA,
    normalizados: valoresColumnaA.map(_normalizarCodigo),
    primerCoincidenciaAproxCharCodes: codigosDeCoincidenciaAproximada[0]
      ? Array.from(codigosDeCoincidenciaAproximada[0]).map(c => c.charCodeAt(0))
      : null,
    filaEncontrada: _buscarFilaPorCodigo(sheet, codigoPrecarga)
  };
}

// ---------------------------------------------------------------------------
// CÓDIGO DE PRECARGA
// ---------------------------------------------------------------------------

function _generarCodigoPrecarga(sheet) {
  const anio = new Date().getFullYear();
  const props = PropertiesService.getScriptProperties();
  const clave = 'ULTIMO_CORRELATIVO_' + anio;
  const ultimo = Number(props.getProperty(clave) || '0');
  const siguiente = ultimo + 1;
  props.setProperty(clave, String(siguiente));
  return 'PRE-' + anio + '-' + String(siguiente).padStart(4, '0');
}

function _buscarFilaPorCodigo(sheet, codigoPrecarga) {
  const columnaCodigo = 1; // Columna A
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null; // solo hay encabezados, todavía no hay filas de datos

  const objetivo = _normalizarCodigo(codigoPrecarga);
  const valores = sheet.getRange(2, columnaCodigo, lastRow - 1, 1).getValues();
  for (let i = 0; i < valores.length; i++) {
    if (_normalizarCodigo(valores[i][0]) === objetivo) return i + 2; // +2: encabezado + índice base 0
  }
  return null;
}

function _normalizarCodigo(valor) {
  return String(valor || '')
    .normalize('NFKC')
    .replace(/[\u2010-\u2015\u2212\uFE58\uFE63\uFF0D]/g, '-') // guion largo/corto/matemático → "-"
    .replace(/\s+/g, '') // saca todos los espacios, no solo los de los extremos
    .toUpperCase();
}

// ---------------------------------------------------------------------------
// DRIVE: CARPETAS Y ARCHIVOS
// ---------------------------------------------------------------------------

function _getOrCreatePropertyFolder(codigoPrecarga, calle, numero) {
  const raiz = DriveApp.getFolderById(CONFIG.DRIVE_ROOT_FOLDER_ID);
  const nombreCarpeta = _sanear(codigoPrecarga + '_' + calle + '_' + numero);

  const existentes = raiz.getFoldersByName(nombreCarpeta);
  if (existentes.hasNext()) return existentes.next();

  return raiz.createFolder(nombreCarpeta);
}

/**
 * archivos: objeto { plano_edificacion: blob, escritura: blob, ... }
 * Nomenclatura de archivo: Calle_Numero.tipoDocumento.ext
 */
function _guardarArchivos(carpeta, codigoPrecarga, calle, numero, archivos) {
  const base = _sanear(calle + '_' + numero);
  const urls = {};

  CONFIG.DOCUMENTOS.forEach(doc => {
    const blob = archivos && archivos[doc.key];
    if (!blob) return;

    const ext = _extensionDesdeMimeType(blob.getContentType());
    const nombreArchivo = base + '.' + doc.key + '.' + ext;

    // Si ya existe un archivo con ese nombre, se reemplaza (se borra la versión anterior)
    const previos = carpeta.getFilesByName(nombreArchivo);
    while (previos.hasNext()) previos.next().setTrashed(true);

    const archivo = carpeta.createFile(blob).setName(nombreArchivo);
    urls[doc.key] = archivo.getUrl();
  });

  return urls;
}

function _extensionDesdeMimeType(mimeType) {
  const mapa = {
    'application/pdf': 'pdf',
    'image/jpeg': 'jpg',
    'image/png': 'png'
  };
  return mapa[mimeType] || 'dat';
}

function _sanear(texto) {
  return String(texto)
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '') // saca acentos
    .replace(/[^a-zA-Z0-9]+/g, '_')
    .replace(/^_+|_+$/g, '');
}

// ---------------------------------------------------------------------------
// ESTADO DE DOCUMENTACIÓN (presentado / pendiente)
// ---------------------------------------------------------------------------

/**
 * Recalcula, mirando directamente el contenido de la carpeta de Drive,
 * qué documentos están presentados. Se usa tanto al guardar como en la
 * alerta semanal, para que el estado nunca dependa de un valor "cacheado".
 *
 * IMPORTANTE: la carpeta se resuelve de forma DETERMINÍSTICA (código +
 * calle + número), nunca leyendo la columna "Carpeta Drive (URL)" del
 * Sheet, porque esa columna puede quedar desactualizada/desalineada si
 * cambia el esquema de encabezados.
 */
function calcularEstadoDocumentacion(codigoPrecarga) {
  const sheet = _getSheet();
  const fila = _buscarFilaPorCodigo(sheet, codigoPrecarga);
  if (!fila) return null;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const valores = sheet.getRange(fila, 1, 1, sheet.getLastColumn()).getValues()[0];
  const calle = valores[headers.indexOf('Calle')];
  const numero = valores[headers.indexOf('Número')];

  const carpeta = _getOrCreatePropertyFolder(codigoPrecarga, calle, numero);
  return _calcularEstadoDesdeCarpeta(carpeta);
}

/**
 * Misma lógica que calcularEstadoDocumentacion, pero recibe la carpeta ya
 * resuelta (evita una vuelta extra a Drive/Sheet cuando ya la tenemos a
 * mano, como en precargarDatos).
 */
function _calcularEstadoDesdeCarpeta(carpeta) {
  const estado = {};
  CONFIG.DOCUMENTOS.forEach(doc => {
    const it = carpeta.getFiles();
    let encontrado = false;
    while (it.hasNext()) {
      const nombre = it.next().getName();
      if (nombre.indexOf('.' + doc.key + '.') !== -1) { encontrado = true; break; }
    }
    estado[doc.key] = encontrado ? 'Presentado' : 'Pendiente';
  });
  return estado;
}

function _idDesdeUrl(url) {
  const match = String(url).match(/[-\w]{25,}/);
  return match ? match[0] : null;
}

// ---------------------------------------------------------------------------
// VERIFICACIÓN DE DOCUMENTACIÓN CON IA (OpenAI)
// ---------------------------------------------------------------------------
//
// La clave de API NUNCA se escribe acá. Se guarda como Propiedad del Script:
// editor de Apps Script → ícono de engranaje "Configuración del proyecto" →
// "Propiedades del script" → agregar propiedad "OPENAI_API_KEY" con el valor
// de la clave. Así queda privada al proyecto y no viaja por el código.
//
// Si la propiedad no está configurada, la verificación simplemente se salta
// (no rompe la precarga ni la subida de documentos) y queda anotado el motivo
// en la columna correspondiente.

function _openAiApiKey() {
  return PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY');
}

/**
 * A partir de los datos del formulario, arma { 1: 'Juan Pérez', 2: '...' }
 * con los nombres de los dueños ya cargados, para poder contrastar contra
 * lo que figura en el DNI/CUIL.
 */
function _nombresDueños(datos) {
  const nombres = {};
  for (let n = 1; n <= CONFIG.MAX_DUEÑOS; n++) {
    const valor = datos && datos['dueño' + n + '_nombre'];
    if (valor) nombres[n] = valor;
  }
  return nombres;
}

/**
 * Qué se espera que sea un documento (para el prompt a la IA) según su key,
 * y a nombre de quién debería estar (si corresponde verificar nombre).
 */
function _infoEsperadaDocumento(key, nombresDueños) {
  nombresDueños = nombresDueños || {};
  let m;

  if (key === 'dni_propietario') return { tipo: 'DNI (documento de identidad) argentino', nombre: nombresDueños[1] || null };
  if (key === 'cuil_propietario') return { tipo: 'constancia de CUIL/CUIT', nombre: nombresDueños[1] || null };
  if ((m = key.match(/^dni_dueño(\d)$/))) return { tipo: 'DNI (documento de identidad) argentino', nombre: nombresDueños[m[1]] || null };
  if ((m = key.match(/^cuil_dueño(\d)$/))) return { tipo: 'constancia de CUIL/CUIT', nombre: nombresDueños[m[1]] || null };
  if (key === 'plano_edificacion') return { tipo: 'plano de edificación de un inmueble', nombre: null };
  if (key === 'plano_mensura') return { tipo: 'plano de mensura de un inmueble', nombre: null };
  if (key === 'escritura') return { tipo: 'escritura de una propiedad', nombre: null };
  if (key === 'titulo_propiedad') return { tipo: 'título de propiedad', nombre: null };

  const doc = CONFIG.DOCUMENTOS.filter(d => d.key === key)[0];
  return { tipo: doc ? doc.label.toLowerCase() : key, nombre: null };
}

function _docPorLabel(label) {
  return CONFIG.DOCUMENTOS.filter(d => d.label === label)[0] || null;
}

/**
 * Recorre un objeto { key: blob } (como el que arma _archivosDesdeBase64) y
 * devuelve { key: 'estado de verificación' } llamando a la IA para cada uno.
 * Nunca tira error: si algo falla, deja anotado el motivo como estado.
 */
function _verificarArchivosConIA(archivos, nombresDueños) {
  const resultado = {};
  if (!archivos) return resultado;
  Object.keys(archivos).forEach(key => {
    const blob = archivos[key];
    if (!blob) return;
    const info = _infoEsperadaDocumento(key, nombresDueños);
    resultado[key] = _verificarDocumentoIA(blob, info.tipo, info.nombre).estado;
  });
  return resultado;
}

/**
 * Manda un archivo (imagen o PDF) a OpenAI para que diga si corresponde al
 * tipo de documento esperado, si es legible, y si el nombre que figura
 * coincide con el nombre esperado. Devuelve { estado, detalle }.
 *
 * estado es un texto corto pensado para mostrar directo en el Tablero:
 * "Verificado OK", "Ilegible / mal escaneo", "Documento incorrecto (...)",
 * "Nombre no coincide (...)", o "No verificado (...)" si hubo un problema
 * técnico (sin clave configurada, sin respuesta, etc.).
 */
/**
 * Helper compartido: manda un archivo (imagen o PDF, como data URI) a
 * OpenAI junto con instrucciones, y devuelve { ok:true, datos:{...} } con
 * el JSON ya parseado, o { ok:false, motivo:'...' } si algo falló en el
 * camino (sin clave, sin red, respuesta rara, etc.). No tira excepciones.
 */
function _consultarOpenAIConArchivo(dataUri, mimeType, filename, instrucciones) {
  const apiKey = _openAiApiKey();
  if (!apiKey) return { ok: false, motivo: 'falta configurar OPENAI_API_KEY' };

  const content = [{ type: 'input_text', text: instrucciones }];
  if (mimeType === 'application/pdf') {
    content.push({ type: 'input_file', filename: filename || 'documento.pdf', file_data: dataUri });
  } else {
    content.push({ type: 'input_image', image_url: dataUri });
  }

  const payload = {
    model: 'gpt-4o-mini',
    input: [{ role: 'user', content: content }]
  };

  const opciones = {
    method: 'post',
    contentType: 'application/json',
    headers: { Authorization: 'Bearer ' + apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let resp;
  try {
    resp = UrlFetchApp.fetch('https://api.openai.com/v1/responses', opciones);
  } catch (e) {
    return { ok: false, motivo: 'error de red con OpenAI (' + e.message + ')' };
  }

  const codigo = resp.getResponseCode();
  if (codigo < 200 || codigo >= 300) {
    return { ok: false, motivo: 'error OpenAI ' + codigo, textoCrudo: resp.getContentText().slice(0, 300) };
  }

  let json;
  try {
    json = JSON.parse(resp.getContentText());
  } catch (e) {
    return { ok: false, motivo: 'respuesta inválida de OpenAI' };
  }

  const texto = _extraerTextoRespuestaOpenAI(json);
  const datos = _parsearJsonIA(texto);
  if (!datos) {
    return { ok: false, motivo: 'no se pudo interpretar la respuesta de la IA', textoCrudo: texto };
  }

  return { ok: true, datos: datos };
}

function _tipoSoportadoPorIA(mimeType) {
  return ['application/pdf', 'image/jpeg', 'image/png'].indexOf(mimeType) !== -1;
}

function _verificarDocumentoIA(blob, tipoEsperado, nombreEsperado) {
  if (!blob) return { estado: 'No verificado (sin archivo)', detalle: null };

  const mimeType = blob.getContentType();
  if (!_tipoSoportadoPorIA(mimeType)) {
    return { estado: 'No verificado (formato no soportado por la IA)', detalle: mimeType };
  }

  const base64 = Utilities.base64Encode(blob.getBytes());
  const dataUri = 'data:' + mimeType + ';base64,' + base64;

  const instrucciones =
    'Sos un asistente que verifica documentación para una inmobiliaria. ' +
    'Te paso un archivo que debería ser: "' + tipoEsperado + '"' +
    (nombreEsperado ? (', a nombre de "' + nombreEsperado + '"') : '') + '. ' +
    'Mirá el archivo con atención y respondé ÚNICAMENTE con un JSON válido, sin texto adicional, ' +
    'sin bloques de código ni comillas triples, con exactamente esta forma: ' +
    '{"tipo_detectado": "<qué tipo de documento es en realidad, en pocas palabras>", ' +
    '"coincide_tipo": <true o false, si el documento corresponde al tipo esperado>, ' +
    '"legible": <true o false, si el contenido relevante se puede leer con claridad>, ' +
    '"nombre_detectado": "<nombre y apellido que figura en el documento, o cadena vacía si no aplica o no se ve>", ' +
    '"nombre_coincide": <true, false, o null si no corresponde verificar un nombre>, ' +
    '"observacion": "<una frase breve explicando el problema si hay alguno, o cadena vacía si está todo bien>"}';

  const resp = _consultarOpenAIConArchivo(dataUri, mimeType, blob.getName(), instrucciones);
  if (!resp.ok) {
    return { estado: 'No verificado (' + resp.motivo + ')', detalle: resp.textoCrudo || null };
  }

  const datosIA = resp.datos;
  let estado;
  if (datosIA.legible === false) {
    estado = 'Ilegible / mal escaneo';
  } else if (datosIA.coincide_tipo === false) {
    estado = 'Documento incorrecto' + (datosIA.tipo_detectado ? ' (parece: ' + datosIA.tipo_detectado + ')' : '');
  } else if (datosIA.nombre_coincide === false) {
    estado = 'Nombre no coincide' + (datosIA.nombre_detectado ? ' (dice: ' + datosIA.nombre_detectado + ')' : '');
  } else {
    estado = 'Verificado OK';
  }

  return { estado: estado, detalle: datosIA.observacion || null };
}

/**
 * Llamada desde el cliente (Precarga.html) apenas se selecciona el archivo
 * del DNI de un dueño, ANTES de enviar el formulario. Lee el documento con
 * IA y devuelve los datos para autocompletar: nombre completo, DNI, fecha
 * de nacimiento (AAAA-MM-DD, formato que espera <input type=date>) y
 * nacionalidad. { ok:false, motivo } si no se pudo leer.
 */
function extraerDatosDNI(archivoBase64) {
  if (!archivoBase64 || !archivoBase64.base64) return { ok: false, motivo: 'no se recibió ningún archivo' };

  const mimeType = archivoBase64.mimeType;
  if (!_tipoSoportadoPorIA(mimeType)) return { ok: false, motivo: 'formato no soportado' };

  const dataUri = 'data:' + mimeType + ';base64,' + archivoBase64.base64;

  const instrucciones =
    'Sos un asistente que lee documentos de identidad argentinos (DNI) para precargar un formulario. ' +
    'Mirá la imagen o archivo y respondé ÚNICAMENTE con un JSON válido, sin texto adicional, sin bloques ' +
    'de código ni comillas triples, con exactamente esta forma: ' +
    '{"es_dni": <true o false, si el archivo es un DNI argentino>, ' +
    '"legible": <true o false, si los datos se pueden leer con claridad>, ' +
    '"nombre": "<nombre/s de pila, tal como figuran>", ' +
    '"apellido": "<apellido/s, tal como figuran>", ' +
    '"dni": "<número de DNI, solo dígitos, sin puntos>", ' +
    '"fecha_nacimiento": "<fecha de nacimiento en formato AAAA-MM-DD, o cadena vacía si no se ve>", ' +
    '"nacionalidad": "<nacionalidad tal como figura, por ejemplo Argentina>", ' +
    '"observacion": "<motivo breve si no se pudo leer algo, o cadena vacía si está todo bien>"}';

  const resp = _consultarOpenAIConArchivo(dataUri, mimeType, archivoBase64.filename, instrucciones);
  if (!resp.ok) return { ok: false, motivo: resp.motivo };

  const d = resp.datos;
  if (d.legible === false || d.es_dni === false) {
    return { ok: false, motivo: d.observacion || (d.es_dni === false ? 'el archivo no parece un DNI' : 'no se pudo leer con claridad') };
  }

  return {
    ok: true,
    nombreCompleto: [d.nombre, d.apellido].filter(Boolean).join(' ').trim(),
    dni: d.dni || '',
    fechaNacimiento: d.fecha_nacimiento || '',
    nacionalidad: d.nacionalidad || ''
  };
}

/**
 * Llamada desde el cliente (Precarga.html) apenas se selecciona el archivo
 * de un plano o la escritura, ANTES de enviar el formulario. Lee el
 * documento con IA y devuelve todo lo que pueda encontrar sobre la
 * propiedad (superficie, ambientes, servicios, etc.) para autocompletar la
 * sección 2. Los campos que la IA no encuentre vienen en null y el cliente
 * los ignora; nunca pisa un campo que el usuario ya haya completado.
 */
function extraerDatosPropiedad(archivoBase64) {
  if (!archivoBase64 || !archivoBase64.base64) return { ok: false, motivo: 'no se recibió ningún archivo' };

  const mimeType = archivoBase64.mimeType;
  if (!_tipoSoportadoPorIA(mimeType)) return { ok: false, motivo: 'formato no soportado' };

  const dataUri = 'data:' + mimeType + ';base64,' + archivoBase64.base64;

  const instrucciones =
    'Sos un asistente que lee planos, escrituras y documentación de inmuebles para precargar un formulario ' +
    'de una inmobiliaria. Mirá el archivo con atención y extraé todos los datos que encuentres sobre la ' +
    'propiedad — pueden no estar todos, usá null en los que no figuren, nunca inventes un valor. Respondé ' +
    'ÚNICAMENTE con un JSON válido, sin texto adicional, sin bloques de código ni comillas triples, con ' +
    'exactamente esta forma: ' +
    '{"legible": <true o false>, ' +
    '"superficie_total_m2": <número o null>, ' +
    '"superficie_cubierta_m2": <número o null>, ' +
    '"dormitorios": <número o null>, ' +
    '"ambientes_sin_dormitorios": <número o null>, ' +
    '"baños": <número o null>, ' +
    '"plantas": <número o null>, ' +
    '"cochera": <true, false o null>, ' +
    '"patio_delantero": <true, false o null>, ' +
    '"patio_trasero": <true, false o null>, ' +
    '"gas_natural": <true, false o null>, ' +
    '"cloaca": <true, false o null>, ' +
    '"tipo_propiedad": "<uno de: Terreno, Casa, Departamento, PH, Local Comercial, Oficina Comercial, Garage - Cochera, Depósito — o null si no se puede determinar>", ' +
    '"orientacion": "<uno de: Norte, Sur, Este, Oeste, Noroeste, Noreste, Sudeste, Sudoeste — o null>", ' +
    '"antiguedad_años": <número o null>, ' +
    '"observacion": "<motivo breve si el archivo no se pudo leer bien, o cadena vacía si está todo bien>"}';

  const resp = _consultarOpenAIConArchivo(dataUri, mimeType, archivoBase64.filename, instrucciones);
  if (!resp.ok) return { ok: false, motivo: resp.motivo };

  const d = resp.datos;
  if (d.legible === false) {
    return { ok: false, motivo: d.observacion || 'no se pudo leer con claridad' };
  }

  return { ok: true, datos: d };
}

/**
 * Llamada desde el cliente al seleccionar el archivo de la escritura.
 * Extrae los titulares (nombre, DNI, domicilio, nacionalidad, fecha de
 * nacimiento) para precargar los bloques de dueño/a en la sección 1.
 * Incluye a TODAS las personas que figuren como titulares, en el orden en
 * que aparecen en el documento.
 */
function extraerDatosEscritura(archivoBase64) {
  if (!archivoBase64 || !archivoBase64.base64) return { ok: false, motivo: 'no se recibió ningún archivo' };

  const mimeType = archivoBase64.mimeType;
  if (!_tipoSoportadoPorIA(mimeType)) return { ok: false, motivo: 'formato no soportado' };

  const dataUri = 'data:' + mimeType + ';base64,' + archivoBase64.base64;

  const instrucciones =
    'Sos un asistente que lee escrituras de propiedades en Argentina, para precargar los datos de los ' +
    'titulares en un formulario. Mirá el documento con atención y respondé ÚNICAMENTE con un JSON válido, ' +
    'sin texto adicional, sin bloques de código ni comillas triples, con exactamente esta forma: ' +
    '{"es_escritura": <true o false>, ' +
    '"legible": <true o false>, ' +
    '"titulares": [ {"nombre": "<nombre/s de pila>", "apellido": "<apellido/s>", ' +
    '"dni": "<número de DNI, solo dígitos, sin puntos>", ' +
    '"domicilio": "<domicilio tal como figura: calle, número, piso, depto, ciudad>", ' +
    '"nacionalidad": "<nacionalidad tal como figura>", ' +
    '"fecha_nacimiento": "<fecha de nacimiento en formato AAAA-MM-DD si figura, o cadena vacía>"} ], ' +
    '"inscripcion_dominio": "<la inscripción de dominio del inmueble, tal como figura en la escritura ' +
    '(matrícula, folio real, tomo, folio, número, tal como esté redactado), o cadena vacía si no figura>", ' +
    '"observacion": "<motivo breve si no se pudo leer algo, o cadena vacía si está todo bien>"}. ' +
    'Incluí en "titulares" a TODAS las personas que figuren como propietarias/titulares del inmueble, en ' +
    'el orden en que aparecen. Si algún dato de una persona no figura, dejalo como cadena vacía, pero ' +
    'igual incluí a esa persona en la lista.';

  const resp = _consultarOpenAIConArchivo(dataUri, mimeType, archivoBase64.filename, instrucciones);
  if (!resp.ok) return { ok: false, motivo: resp.motivo };

  const d = resp.datos;
  if (d.legible === false || d.es_escritura === false) {
    return { ok: false, motivo: d.observacion || (d.es_escritura === false ? 'el archivo no parece una escritura' : 'no se pudo leer con claridad') };
  }

  const titulares = (d.titulares || []).map(t => ({
    nombreCompleto: [t.nombre, t.apellido].filter(Boolean).join(' ').trim(),
    dni: t.dni || '',
    domicilio: t.domicilio || '',
    nacionalidad: t.nacionalidad || '',
    fechaNacimiento: t.fecha_nacimiento || ''
  })).filter(t => t.nombreCompleto);

  if (!titulares.length) return { ok: false, motivo: 'no se detectaron titulares en el documento' };

  return { ok: true, titulares: titulares, inscripcionDominio: d.inscripcion_dominio || '' };
}

function _extraerTextoRespuestaOpenAI(json) {
  if (json && typeof json.output_text === 'string') return json.output_text;
  const out = (json && json.output) || [];
  for (let i = 0; i < out.length; i++) {
    const contenido = out[i] && out[i].content;
    if (!contenido) continue;
    for (let j = 0; j < contenido.length; j++) {
      const c = contenido[j];
      if (c && c.type === 'output_text' && c.text) return c.text;
    }
  }
  return '';
}

function _parsearJsonIA(texto) {
  const limpio = String(texto || '').trim()
    .replace(/^```json/i, '').replace(/^```/, '').replace(/```$/, '').trim();
  try {
    return JSON.parse(limpio);
  } catch (e) {
    return null;
  }
}

/**
 * DIAGNÓSTICO — correr a mano desde el editor para probar la conexión con
 * OpenAI sin depender de una precarga real. Requiere haber configurado
 * OPENAI_API_KEY en las Propiedades del script.
 */
function debugVerificarIA() {
  const carpeta = DriveApp.getFolderById(CONFIG.DRIVE_ROOT_FOLDER_ID);
  Logger.log('Configurá manualmente un blob de prueba si querés probar esta función.');
  Logger.log('Clave configurada: ' + (_openAiApiKey() ? 'sí' : 'NO — falta agregar OPENAI_API_KEY en Propiedades del script'));
}

// ---------------------------------------------------------------------------
// VALIDACIÓN
// ---------------------------------------------------------------------------

function _validarCamposObligatoriosPrimerEnvio(datos, archivos) {
  const faltantesCampos = CONFIG.CAMPOS_OBLIGATORIOS_PRECARGA.filter(c => !datos[c]);
  const faltantesDocs = CONFIG.DOCUMENTOS_OBLIGATORIOS_PRECARGA.filter(k => !(archivos && archivos[k]));

  if (faltantesCampos.length || faltantesDocs.length) {
    throw new Error(
      'Faltan datos obligatorios para el primer envío de precarga. ' +
      'Campos: ' + faltantesCampos.join(', ') + '. ' +
      'Documentos: ' + faltantesDocs.join(', ')
    );
  }
}

// ---------------------------------------------------------------------------
// PERSISTENCIA EN EL SHEET
// ---------------------------------------------------------------------------

function _getSheet() {
  const libro = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  let hoja = libro.getSheetByName(CONFIG.SHEET_NAME);
  if (!hoja) {
    hoja = libro.insertSheet(CONFIG.SHEET_NAME);
    hoja.appendRow(_encabezadosCompletos());
  }
  return hoja;
}

function _encabezadosCompletos() {
  // Meta + dueños + propiedad + una columna de estado por cada documento
  // + una columna de verificación por IA por cada documento (al final, para
  // no correr/desalinear columnas ya existentes en el Sheet real).
  const columnasDocs = CONFIG.DOCUMENTOS.map(d => 'Doc: ' + d.label);
  const columnasVerifIA = _columnasVerificacionIA();
  return COLUMNAS_META.concat(_columnasDueños()).concat(COLUMNAS_PROPIEDAD).concat(columnasDocs).concat(columnasVerifIA);
}

function _columnasVerificacionIA() {
  return CONFIG.DOCUMENTOS.map(d => 'Verificación IA: ' + d.label);
}

function _guardarFila(sheet, datos, urlCarpeta, estadoDocs, verificacionIA) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const filaExistente = _buscarFilaPorCodigo(sheet, datos.codigoPrecarga);
  estadoDocs = estadoDocs || calcularEstadoDocumentacion(datos.codigoPrecarga) || {};

  const mapaValores = {
    'Código de Precarga': datos.codigoPrecarga,
    'Fecha de creación': filaExistente ? undefined : new Date(),
    'Última actualización': new Date(),
    'Calle': datos.calle,
    'Número': datos.numero,
    'Departamento': datos.departamento,
    'Piso': datos.piso,
    'Pasillo': datos.pasillo,
    'Entrecalle 1': datos.entrecalle1,
    'Entrecalle 2': datos.entrecalle2,
    'Ciudad': datos.ciudad,
    'Tipo de propiedad': datos.tipoPropiedad,
    'Superficie Total (m2)': datos.superficieTotal,
    'Superficie Cubierta (m2)': datos.superficieCubierta,
    'Dormitorios': datos.dormitorios,
    'Ambientes (sin dormitorios)': datos.ambientes,
    'Baños': datos.baños,
    'Plantas': datos.plantas,
    'Cochera': datos.cochera,
    'Patio delantero': datos.patioDelantero,
    'Patio trasero': datos.patioTrasero,
    'Gas Natural': datos.gasNatural,
    'Cloaca': datos.cloaca,
    'Antigüedad (años)': datos.antiguedad,
    'Orientación': datos.orientacion,
    'Inscripción del dominio': datos.inscripcionDominio,
    'Carpeta Drive (URL)': urlCarpeta,
    'Estado de la propiedad': filaExistente ? undefined : 'Verificación para la carga'
    ,'Etapa': filaExistente ? undefined : 'Captación'
  };

  CONFIG.DOCUMENTOS.forEach(d => {
    mapaValores['Doc: ' + d.label] = estadoDocs[d.key] || 'Pendiente';
    if (verificacionIA && verificacionIA[d.key] !== undefined) {
      mapaValores['Verificación IA: ' + d.label] = verificacionIA[d.key];
    }
  });

  for (let n = 1; n <= CONFIG.MAX_DUEÑOS; n++) {
    CAMPOS_DUEÑO.forEach(c => {
      const valor = datos['dueño' + n + '_' + c.clave];
      if (valor !== undefined) {
        mapaValores['Dueño ' + n + ' - ' + c.columna] = valor;
      }
    });
  }

  const fila = filaExistente || sheet.getLastRow() + 1;
  headers.forEach((h, i) => {
    if (mapaValores[h] !== undefined) {
      sheet.getRange(fila, i + 1).setValue(mapaValores[h]);
    }
  });

  // % de completitud simple: proporción de documentos presentados
  const totalDocs = CONFIG.DOCUMENTOS.length;
  const presentados = Object.values(estadoDocs).filter(v => v === 'Presentado').length;
  const pctCol = headers.indexOf('% Completitud') + 1;
  if (pctCol) sheet.getRange(fila, pctCol).setValue(Math.round((presentados / totalDocs) * 100) + '%');
}

// ---------------------------------------------------------------------------
// NOTIFICACIONES
// ---------------------------------------------------------------------------

function _enviarMailConfirmacionPrecarga(datos) {
  const destinatario = datos['dueño1_email'];
  if (!destinatario) return;

  const nombre = datos['dueño1_nombre'] || '';
  const direccion = (datos.calle || '') + ' ' + (datos.numero || '');
  const asunto = 'Azcuénaga Inmobiliaria — Precarga registrada: ' + datos.codigoPrecarga;

  const cuerpoTexto =
    'Hola ' + nombre + ', registramos tu carga parcial para la propiedad en ' + direccion + '. ' +
    'Tu número de precarga es: ' + datos.codigoPrecarga + '. Podés retomar la carga en cualquier momento usando ese número; ' +
    'cuando tengas el resto de la documentación, volvé a ingresar al formulario y continuá desde donde quedaste.\n\n' +
    'Saludos, ' + CONFIG.REMITENTE_NOMBRE;

  const cuerpoHtml =
    '<div style="font-family:\'Segoe UI\',Helvetica,Arial,sans-serif;background:#faf7f1;padding:28px 16px;">' +
      '<div style="max-width:520px;margin:0 auto;background:#ffffff;border-radius:10px;overflow:hidden;border:1px solid #e7e1d6;">' +
        '<div style="background:#1b3866;color:#ffffff;padding:18px 24px;font-size:16px;font-weight:600;">' +
          'Azcuénaga Inmobiliaria' +
        '</div>' +
        '<div style="padding:26px 24px;">' +
          '<p style="text-align:justify;font-size:14px;line-height:1.6;color:#22293a;margin:0 0 18px;">' +
            'Hola ' + nombre + ', registramos tu carga parcial para la propiedad en <strong>' + direccion + '</strong>. ' +
            'Podés retomar la carga en cualquier momento usando tu número de precarga, y cuando tengas el resto de la documentación simplemente volvé a ingresar al formulario para continuar desde donde quedaste.' +
          '</p>' +
          '<div style="background:#fbf3ec;border-left:3px solid #b5673e;border-radius:0 8px 8px 0;padding:12px 18px;margin-bottom:20px;">' +
            '<span style="font-size:12.5px;color:#7a4a30;">Tu número de precarga</span><br>' +
            '<span style="font-size:19px;font-weight:700;color:#1b3866;">' + datos.codigoPrecarga + '</span>' +
          '</div>' +
          '<p style="font-size:13px;color:#6b7280;margin:0;">Saludos,<br><strong style="color:#1b3866;">' + CONFIG.REMITENTE_NOMBRE + '</strong></p>' +
        '</div>' +
      '</div>' +
    '</div>';

  MailApp.sendEmail({
    to: destinatario,
    subject: asunto,
    body: cuerpoTexto,
    htmlBody: cuerpoHtml,
    name: CONFIG.REMITENTE_NOMBRE
  });
}

/* ── Reporte al propietario por e-mail ── */

function enviarReportePropietario(codigo, logoDataUri, resumenCustom, ownerCustom) {
  var mae = hojaMaestra_();
  var fila = _filaPorCodigo_(mae, codigo);
  if (fila < 0) throw new Error('No encontré la propiedad ' + codigo);
  var head = mae.getRange(1, 1, 1, mae.getLastColumn()).getValues()[0];
  var row  = mae.getRange(fila, 1, 1, mae.getLastColumn()).getValues()[0];
  function c(n) { var i = head.indexOf(n); return i > -1 ? i : -1; }
  var calle = row[c('Calle')] || '', num = row[c('Número')] || '';
  var dir = (calle + ' ' + num).trim();
  var segMap = leerSeguimiento_();
  var seg = segMap[codigo] || [];
  var cons = seg.length;
  var vis = seg.filter(function(s) { return /visita/i.test(s.tipo) || /visitó|oferta/i.test(s.res); }).length;
  var ofertas = seg.filter(function(s) { return /propuesta|oferta/i.test(s.res); }).length;
  var hoy = new Date();
  var dias = null;
  seg.forEach(function(s) { if (s._d) { var dd = Math.floor((hoy - s._d) / 86400000); if (dias === null || dd < dias) dias = dd; } });

  var frase;
  if (vis >= 3) frase = 'Se mantiene un alto interés de compradores. Seguimos trabajando para concretar una propuesta.';
  else if (vis > 0) frase = 'Hubo visitas a la propiedad. Estamos trabajando para generar una propuesta concreta.';
  else if (cons > 0) frase = 'Se recibieron consultas pero todavía sin visitas. Sugerimos reforzar la comunicación con los interesados.';
  else frase = 'No se registraron consultas, sugerimos revisar precio.';

  var resumen = 'Durante el período, la propiedad ' + dir + ' registró ' + cons + ' consulta(s), ' + vis + ' visita(s) y ' + ofertas + ' propuesta(s). ' + frase;
  var resumenTxt = resumen;
  if (resumenCustom && String(resumenCustom).trim()) {
    resumen = String(resumenCustom).trim().replace(/<(?!\/?(b|i|u|br|span)\b)[^>]*>/gi, '').replace(/<span(?![^>]*style="font-size:\d+%")[^>]*>/gi, '<span>');
    resumenTxt = resumen.replace(/<br\s*\/?>/gi, '\n').replace(/<[^>]+>/g, '').replace(/&lt;/g,'<').replace(/&gt;/g,'>').replace(/&amp;/g,'&');
  }

  var destinatarios = [];
  var sinMail = [];
  for (var k = 1; k <= CONFIG.MAX_DUEÑOS; k++) {
    var ci = c('Dueño ' + k + ' - Nombre y Apellido');
    if (ci < 0) continue;
    var nombre = ('' + (row[ci] || '')).trim();
    if (!nombre) continue;
    var ei = c('Dueño ' + k + ' - E-mail');
    var email = ei > -1 ? ('' + (row[ei] || '')).trim() : '';
    if (email && email.indexOf('@') > -1) {
      destinatarios.push({ nombre: nombre, email: email });
    } else {
      sinMail.push(nombre);
    }
  }

  var inlineImages = {};
  if (logoDataUri) {
    var match = logoDataUri.match(/^data:image\/(png|jpe?g|gif);base64,(.+)$/i);
    if (match) {
      inlineImages.logo = Utilities.newBlob(Utilities.base64Decode(match[2]), 'image/' + match[1], 'logo.png');
    }
  }

  var fecha = Utilities.formatDate(hoy, 'GMT-3', 'dd/MM/yyyy');
  var enviados = [];
  destinatarios.forEach(function(dest) {
    var html = _buildReporteHtml_((ownerCustom && String(ownerCustom).trim()) || dest.nombre, dir, codigo, cons, vis, ofertas, dias, resumen, fecha);
    var textoPlano = 'Informe de gestión — ' + dir + '\n\n' +
      'Hola ' + dest.nombre + ', te compartimos el resumen de gestión de tu propiedad ' + dir + '.\n\n' +
      'Consultas: ' + cons + ' | Visitas: ' + vis + ' | Propuestas: ' + ofertas + ' | Días s/mov: ' + (dias === null ? '—' : dias) + '\n\n' +
      resumenTxt + '\n\nSaludos,\n' + CONFIG.REMITENTE_NOMBRE;
    MailApp.sendEmail({
      to: dest.email,
      subject: 'Informe de gestión — ' + dir,
      body: textoPlano,
      htmlBody: html,
      name: CONFIG.REMITENTE_NOMBRE,
      inlineImages: inlineImages
    });
    enviados.push(dest.nombre + ' (' + dest.email + ')');
  });

  return { ok: true, enviados: enviados, sinMail: sinMail };
}

function _buildReporteHtml_(nombre, dir, codigo, cons, vis, ofertas, dias, resumen, fecha) {
  var metricStyle = 'display:inline-block;text-align:center;background:#f0f4fa;border-radius:10px;padding:14px 18px;min-width:80px;margin:4px;';
  var numStyle = 'font-size:22px;font-weight:700;color:#1b3866;display:block;';
  var labelStyle = 'font-size:10px;color:#6b7280;text-transform:uppercase;letter-spacing:0.5px;display:block;margin-top:2px;';
  return '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body style="margin:0;padding:0;background:#f0f4fa;font-family:Arial,Helvetica,sans-serif;">' +
    '<div style="max-width:520px;margin:24px auto;background:#ffffff;border-radius:14px;overflow:hidden;box-shadow:0 2px 12px rgba(0,0,0,0.08);">' +
    '<div style="background:#1b3866;padding:18px 24px;color:#ffffff;">' +
    '<table cellpadding="0" cellspacing="0" border="0"><tr>' +
    '<td style="vertical-align:middle;padding-right:12px;"><img src="cid:logo" alt="Azcuénaga" style="height:36px;border-radius:50%;"></td>' +
    '<td style="vertical-align:middle;"><span style="font-size:11px;color:#b0c4e0;">Reporte de gestión</span><br><span style="font-size:15px;font-weight:700;">INFORME DE GESTIÓN</span></td>' +
    '</tr></table></div>' +
    '<div style="padding:24px 28px;">' +
    '<p style="font-size:14px;color:#333;margin:0 0 16px;">Estimado/a <strong>' + nombre + '</strong>, te compartimos el resumen de gestión de tu propiedad <strong>' + dir + '</strong>.</p>' +
    '<div style="text-align:center;margin:20px 0;">' +
    '<div style="' + metricStyle + '"><span style="' + numStyle + '">' + cons + '</span><span style="' + labelStyle + '">Consultas</span></div>' +
    '<div style="' + metricStyle + '"><span style="' + numStyle + '">' + vis + '</span><span style="' + labelStyle + '">Visitas</span></div>' +
    '<div style="' + metricStyle + '"><span style="' + numStyle + '">' + ofertas + '</span><span style="' + labelStyle + '">Propuestas</span></div>' +
    '<div style="' + metricStyle + '"><span style="' + numStyle + '">' + (dias === null ? '—' : dias) + '</span><span style="' + labelStyle + '">Días s/mov</span></div>' +
    '</div>' +
    '<p style="font-size:13.5px;color:#444;line-height:1.5;margin:16px 0 20px;">' + resumen + '</p>' +
    '<p style="font-size:12px;color:#999;margin:20px 0 0;border-top:1px solid #eee;padding-top:14px;">Saludos,<br><strong style="color:#1b3866;">' + CONFIG.REMITENTE_NOMBRE + '</strong><br>' +
    '<span style="font-size:11px;">Mendoza 4873 · Rosario · ' + fecha + '</span></p>' +
    '</div></div></body></html>';
}

/**
 * Pensado para un trigger semanal (time-driven trigger).
 * Recorre todas las propiedades no vendidas y avisa por mail la
 * documentación pendiente de cada una.
 */
function alertaSemanalDocumentacionPendiente() {
  const sheet = _getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; // todavía no hay propiedades cargadas

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const datos = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  const idxEstado = headers.indexOf('Estado de la propiedad');
  const idxCodigo = headers.indexOf('Código de Precarga');

  datos.forEach(fila => {
    if (fila[idxEstado] === 'Vendida') return;

    const codigoPrecarga = fila[idxCodigo];
    const estadoDocs = calcularEstadoDocumentacion(codigoPrecarga);
    const pendientes = CONFIG.DOCUMENTOS.filter(d => estadoDocs[d.key] === 'Pendiente').map(d => d.label);
    if (!pendientes.length) return;

    // TODO: reemplazar por el email del dueño 1 guardado en la fila, y/o WhatsApp
    // MailApp.sendEmail(emailDueño, 'Documentación pendiente — ' + codigoPrecarga, pendientes.join('\n'));
  });
}

/**
 * UTILITARIO MANUAL — correr una sola vez desde el editor de Apps Script
 * (Ejecutar > recalcularEstadosDocumentacion) para recalcular, fila por
 * fila, el estado de todos los "Doc: ..." usando la resolución
 * determinística de carpeta (en vez del valor viejo/cacheado leído de
 * "Carpeta Drive (URL)"). Salta las filas migradas (MIG-*) porque esas
 * propiedades no tienen carpeta creada por este mecanismo.
 */
function recalcularEstadosDocumentacion() {
  const sheet = _getSheet();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idxCodigo = headers.indexOf('Código de Precarga');
  const codigos = sheet.getRange(2, idxCodigo + 1, lastRow - 1, 1).getValues();

  let actualizados = 0;
  codigos.forEach((fila, i) => {
    const codigoPrecarga = fila[0];
    if (!codigoPrecarga || String(codigoPrecarga).indexOf('MIG-') === 0) return;

    const estadoDocs = calcularEstadoDocumentacion(codigoPrecarga);
    if (!estadoDocs) return;

    const numeroFila = i + 2;
    CONFIG.DOCUMENTOS.forEach(d => {
      const col = headers.indexOf('Doc: ' + d.label) + 1;
      if (col) sheet.getRange(numeroFila, col).setValue(estadoDocs[d.key] || 'Pendiente');
    });
    actualizados++;
  });

  Logger.log('Filas actualizadas: ' + actualizados);
}

// ---------------------------------------------------------------------------
// INFORME DE CARGA TOKKO (PDF con la información en el mismo orden del
// formulario de carga de Tokko, para el botón "Generar informe carga Tokko"
// del Tablero — solo aplica a propiedades en etapa "Captación")
// ---------------------------------------------------------------------------

function generarInformeTokko(codigo) {
  const mae = hojaMaestra_();
  const fila = _filaPorCodigo_(mae, codigo);
  if (fila < 0) throw new Error('No encontré la propiedad ' + codigo);

  const head = mae.getRange(1, 1, 1, mae.getLastColumn()).getValues()[0];
  const vals = mae.getRange(fila, 1, 1, mae.getLastColumn()).getValues()[0];
  const d = {};
  head.forEach((h, i) => { d[h] = vals[i]; });

  const FALTA = '— completar manualmente —';
  function v(campo) {
    const val = d[campo];
    if (val === undefined || val === null || val === '') return FALTA;
    if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), 'dd/MM/yyyy');
    return String(val);
  }
  function siNo(campo) {
    const val = d[campo];
    if (val === undefined || val === null || val === '') return FALTA;
    const s = String(val).trim().toLowerCase();
    if (['si','sí','true','1','x'].indexOf(s) !== -1) return 'Sí';
    if (['no','false','0'].indexOf(s) !== -1) return 'No';
    return String(val);
  }

  const dueños = [];
  for (let n = 1; n <= CONFIG.MAX_DUEÑOS; n++) {
    const nombre = d['Dueño ' + n + ' - Nombre y Apellido'];
    if (nombre) {
      dueños.push({
        nombre: nombre,
        dni: d['Dueño ' + n + ' - DNI'] || FALTA,
        telefono: d['Dueño ' + n + ' - Celular'] || FALTA,
        email: d['Dueño ' + n + ' - Email'] || FALTA
      });
    }
  }

  const doc = DocumentApp.create('Informe carga Tokko - ' + codigo);
  const body = doc.getBody();
  body.setMarginTop(36).setMarginBottom(36).setMarginLeft(40).setMarginRight(40);

  const tituloEstilo = {};
  tituloEstilo[DocumentApp.Attribute.FONT_SIZE] = 18;
  tituloEstilo[DocumentApp.Attribute.BOLD] = true;
  tituloEstilo[DocumentApp.Attribute.FOREGROUND_COLOR] = '#1b3866';

  const subEstilo = {};
  subEstilo[DocumentApp.Attribute.FONT_SIZE] = 13;
  subEstilo[DocumentApp.Attribute.BOLD] = true;
  subEstilo[DocumentApp.Attribute.FOREGROUND_COLOR] = '#1b3866';

  const subSubEstilo = {};
  subSubEstilo[DocumentApp.Attribute.FONT_SIZE] = 11;
  subSubEstilo[DocumentApp.Attribute.BOLD] = true;
  subSubEstilo[DocumentApp.Attribute.ITALIC] = true;
  subSubEstilo[DocumentApp.Attribute.FOREGROUND_COLOR] = '#b5673e';

  body.appendParagraph('Informe carga Tokko').setAttributes(tituloEstilo);
  body.appendParagraph('Propiedad: ' + codigo + '  ·  Generado: ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'))
    .setFontSize(10).setForegroundColor('#6b7280');
  body.appendHorizontalRule();

  function fila_(label, valor) {
    const p = body.appendParagraph('');
    p.appendText(label + ': ').setBold(true);
    p.appendText(valor);
  }

  // 1. Nueva Propiedad
  body.appendParagraph('1. Nueva Propiedad').setAttributes(subEstilo);
  fila_('Propietario (Privado)', dueños.length ? dueños.map(x => x.nombre).join(' / ') : FALTA);
  fila_('Tipo de propiedad', v('Tipo de propiedad'));
  fila_('Calle', v('Calle'));
  fila_('Número', v('Número'));
  fila_('Dirección para publicar', v('Calle') + ' ' + v('Número'));
  fila_('Título de publicación para portales', FALTA);
  fila_('País', 'Argentina');
  fila_('Región', 'Santa Fe');
  fila_('Localidad/Partido', v('Ciudad'));
  fila_('Sub-división', v('Pasillo'));

  // 2. Información general
  body.appendParagraph('2. Información general').setAttributes(subEstilo);

  body.appendParagraph('Características generales').setAttributes(subSubEstilo);
  fila_('Ambientes', v('Ambientes (sin dormitorios)'));
  fila_('Antigüedad', v('Antigüedad (años)'));
  fila_('Baños', v('Baños'));
  fila_('Condición', FALTA);
  fila_('Dormitorios', v('Dormitorios'));
  fila_('Ocupación', FALTA);
  fila_('Orientación', v('Orientación'));
  fila_('Plantas', v('Plantas'));
  fila_('Suites', FALTA);
  fila_('Suites con placares', FALTA);
  fila_('Toilettes', FALTA);
  fila_('Zonificación', FALTA);

  body.appendParagraph('Cocheras').setAttributes(subSubEstilo);
  fila_('Cocheras', siNo('Cochera'));
  fila_('Cocheras cubiertas', FALTA);
  fila_('Cocheras descubiertas', FALTA);

  body.appendParagraph('Salas comunes').setAttributes(subSubEstilo);
  fila_('Salas comunes', FALTA);
  fila_('Comedores', FALTA);
  fila_('Livings', FALTA);
  fila_('Salas de TV', FALTA);

  body.appendParagraph('Superficies y medidas').setAttributes(subSubEstilo);
  fila_('Terreno', v('Superficie Total (m2)'));
  fila_('Descubierta', FALTA);
  fila_('Superficie cubierta', v('Superficie Cubierta (m2)'));
  fila_('Superficie semicubierta', FALTA);
  fila_('Total construido', v('Superficie Cubierta (m2)'));
  fila_('Fondo', FALTA);
  fila_('Frente', FALTA);

  // 3. Servicios, ambientes y adicionales
  body.appendParagraph('3. Servicios, ambientes y adicionales').setAttributes(subEstilo);

  body.appendParagraph('Servicios').setAttributes(subSubEstilo);
  fila_('Agua Corriente', FALTA);
  fila_('Cable', FALTA);
  fila_('Cloaca', siNo('Cloaca'));
  fila_('Electricidad', FALTA);
  fila_('Gas Natural', siNo('Gas Natural'));
  fila_('Internet', FALTA);
  fila_('Pavimento', FALTA);
  fila_('Teléfono', FALTA);
  fila_('Wifi', FALTA);

  body.appendParagraph('Ambientes').setAttributes(subSubEstilo);
  fila_('Patio delantero', siNo('Patio delantero'));
  fila_('Patio trasero', siNo('Patio trasero'));
  fila_('Otros ambientes (Altillo, Balcón, Baulera, Biblioteca, Cocina, Comedor diario, Dependencia, Escritorio, Galería, etc.)', FALTA);

  body.appendParagraph('Adicionales').setAttributes(subSubEstilo);
  fila_('Aire Acondicionado / Calefacción / Amoblado / Alarma / Apto mascotas / Apto profesional, etc.', FALTA);

  // 4. Datos internos — NO publicar en Tokko
  body.appendParagraph('4. Datos internos — NO publicar en Tokko').setAttributes(subEstilo);
  fila_('Código de Precarga', codigo);
  dueños.forEach((x, i) => {
    fila_('Dueño ' + (i + 1), x.nombre + '  ·  DNI ' + x.dni + '  ·  Tel ' + x.telefono + '  ·  ' + x.email);
  });
  fila_('Entrecalle 1', v('Entrecalle 1'));
  fila_('Entrecalle 2', v('Entrecalle 2'));
  fila_('Estado de la propiedad', v('Estado de la propiedad'));

  doc.saveAndClose();
  const pdfBlob = DriveApp.getFileById(doc.getId()).getAs('application/pdf');
  const base64 = Utilities.base64Encode(pdfBlob.getBytes());
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  return { base64: base64, filename: 'Informe_carga_Tokko_' + codigo + '.pdf' };
}

function unificarMaestro(){
  var ID='1O8VKF9ZUC073kGNYIc3SgKPcIrVpnvynpQQ69RAk9W8';
  var ss=SpreadsheetApp.openById(ID);
  var hoja=null,sh=ss.getSheets();
  for(var i=0;i<sh.length;i++){
    var h=sh[i].getRange(1,1,1,sh[i].getLastColumn()).getValues()[0].join('|');
    if(h.indexOf('Código de Precarga')>-1){hoja=sh[i];break;}
  }
  if(!hoja)throw new Error('No encontré la pestaña con "Código de Precarga"');
  var nuevas=['Precio','Moneda','Tipo de operación','Etapa','Motivo de suspensión','Origen','Inmobiliaria colega','Contacto colega','Reparto de comisión','Vendedor asignado','Link publicación (fotos)','Link Tokko','Observaciones comerciales'];
  var last=hoja.getLastColumn();
  var headers=hoja.getRange(1,1,1,last).getValues()[0];
  nuevas.forEach(function(n){if(headers.indexOf(n)===-1){last++;hoja.getRange(1,last).setValue(n);}});
  if(!ss.getSheetByName('Seguimiento')){
    var seg=ss.insertSheet('Seguimiento');
    seg.getRange(1,1,1,10).setValues([['Código de Precarga','Fecha','Vía','Interesado','Teléfono','Resultado','Próximo paso','Fecha próximo contacto','Observaciones','Vendedor']]);
    seg.setFrozenRows(1);
  }
  if(!ss.getSheetByName('Agenda')){
    var ag=ss.insertSheet('Agenda');
    ag.getRange(1,1,1,8).setValues([['Código de Precarga','Fecha','Hora','Interesado','Teléfono','Vendedor','Nota','Estado']]);
    ag.setFrozenRows(1);
  }
  Logger.log('Listo: columnas y pestañas unificadas.');
}
function volcarVentas(){
  var ID_VENTAS='16-xykemcOS78QxNymU5Se7RGLwZqvxvXvZ1guhzUECQ';
  var ID_MAESTRO='1O8VKF9ZUC073kGNYIc3SgKPcIrVpnvynpQQ69RAk9W8';

  // 1. leer sheet viejo (busca la pestaña con Dirección/Estado/Tipo)
  var sv=SpreadsheetApp.openById(ID_VENTAS).getSheets(), base=null;
  for(var i=0;i<sv.length;i++){
    var h=sv[i].getRange(1,1,1,sv[i].getLastColumn()).getValues()[0].join('|');
    if(h.indexOf('Direcci')>-1 && h.indexOf('Estado')>-1 && h.indexOf('Tipo')>-1){base=sv[i];break;}
  }
  if(!base)throw new Error('No encontré la base de propiedades en el sheet viejo');
  var vd=base.getDataRange().getValues(), vh=vd[0];
  function col(nombre){for(var k=0;k<vh.length;k++){if((''+vh[k]).toLowerCase().indexOf(nombre.toLowerCase())>-1)return k;}return -1;}
  var cId=col('ID'),cDir=col('Direcc'),cTipo=col('Tipo'),cProp=col('Propietario'),cPrecio=col('Precio'),cEstado=col('Estado'),cObs=col('Observ'),cLink=col('Link');

  // 2. maestro: headers + códigos ya cargados
  var ms=SpreadsheetApp.openById(ID_MAESTRO).getSheets(), mae=null;
  for(var j=0;j<ms.length;j++){
    if(ms[j].getRange(1,1,1,ms[j].getLastColumn()).getValues()[0].join('|').indexOf('Código de Precarga')>-1){mae=ms[j];break;}
  }
  if(!mae)throw new Error('No encontré la pestaña maestra (Código de Precarga)');
  var mh=mae.getRange(1,1,1,mae.getLastColumn()).getValues()[0];
  var existentes={};
  if(mae.getLastRow()>1){mae.getRange(2,1,mae.getLastRow()-1,1).getValues().forEach(function(r){existentes[r[0]]=1;});}
  function idx(nombre){return mh.indexOf(nombre);}

  var em={'En venta':'Publicada','Reservada':'Reserva','Vendida':'Vendida','Suspendida':'Suspendida'};
  var nuevas=[], saltadas=0;

  for(var r=1;r<vd.length;r++){
    var id=(cId>-1?vd[r][cId]:'')||'';
    var dir=(cDir>-1?vd[r][cDir]:'')||'';
    if(!id && !dir)continue;
    var codigo='MIG-'+(id||('R'+r));
    if(existentes[codigo]){saltadas++;continue;}

    var d=(''+dir).trim(), calle=d, num='';
    var m=d.match(/^(.*?)[\s,\.]*(\d+)\s*$/);
    if(m){calle=m[1].trim();num=m[2];}
    var dl=d.toLowerCase(), ciudad='Rosario';
    if(dl.indexOf('funes')>-1)ciudad='Funes'; else if(dl.indexOf('ibarluce')>-1)ciudad='Ibarlucea';

    var precio=(''+(cPrecio>-1?vd[r][cPrecio]:'')), moneda='';
    if(/US\$|u\$s|usd/i.test(precio))moneda='USD'; else if(/\$/.test(precio))moneda='ARS';
    var precioNum=precio.replace(/[^\d]/g,'');
    var estado=(cEstado>-1?vd[r][cEstado]:'')||'';

    var fila=new Array(mh.length).fill('');
    function set(nombre,val){var p=idx(nombre);if(p>-1)fila[p]=val;}
    set('Código de Precarga',codigo);
    set('Fecha de creación',Utilities.formatDate(new Date(),'GMT-3','dd/MM/yyyy'));
    set('Dueño 1 - Nombre y Apellido',(cProp>-1?vd[r][cProp]:'')||'');
    set('Calle',calle); set('Número',num); set('Ciudad',ciudad);
    set('Tipo de propiedad',(cTipo>-1?vd[r][cTipo]:'')||'');
    set('Precio',precioNum); set('Moneda',moneda);
    set('Tipo de operación','Venta');
    set('Etapa',em[estado]||'Publicada');
    set('Estado de la propiedad','Migrado - revisar');
    set('Origen','Propia');
    set('Link publicación (fotos)',(cLink>-1?vd[r][cLink]:'')||'');
    set('Observaciones comerciales','Migrado de ventas '+id+'. Dirección original: "'+d+'". '+((cObs>-1?vd[r][cObs]:'')||''));
    nuevas.push(fila);
  }

  if(nuevas.length){mae.getRange(mae.getLastRow()+1,1,nuevas.length,mh.length).setValues(nuevas);}
  Logger.log('Volcadas: '+nuevas.length+' · Saltadas (ya estaban): '+saltadas);
  SpreadsheetApp.getUi().alert('Volcado listo.\nNuevas: '+nuevas.length+'\nSaltadas: '+saltadas);
}
var MAESTRO_ID='1O8VKF9ZUC073kGNYIc3SgKPcIrVpnvynpQQ69RAk9W8';

function hojaMaestra_(){
  var sh=SpreadsheetApp.openById(MAESTRO_ID).getSheets();
  for(var i=0;i<sh.length;i++){
    var head=sh[i].getRange(1,1,1,sh[i].getLastColumn()).getValues()[0].join('|');
    if(head.indexOf('Código de Precarga')>-1)return sh[i];
  }
  throw new Error('No encontré la pestaña con "Código de Precarga"');
}

function fmtPrecio_(v,mon){
  if(v===''||v==null)return '';
  var n=(''+v).replace(/[^\d]/g,'');
  if(!n)return '';
  var s=Number(n).toLocaleString('es-AR');
  return (mon==='USD'?'US$':(mon==='ARS'?'$':''))+s;
}

function _pdias_(fstr){
  if(!fstr)return null;
  var d;
  if(Object.prototype.toString.call(fstr)==='[object Date]'){ d=fstr; }
  else { var m=(''+fstr).trim().match(/(\d{1,2})\/(\d{1,2})\/(\d{2,4})/); if(!m)return null; var y=m[3].length===2?'20'+m[3]:m[3]; d=new Date(parseInt(y),parseInt(m[2])-1,parseInt(m[1])); }
  if(isNaN(d.getTime()))return null;
  return d;
}
function leerSeguimiento_(){
  var ss=SpreadsheetApp.openById(MAESTRO_ID);
  var seg=ss.getSheetByName('Seguimiento');
  var map={};
  if(!seg || seg.getLastRow()<2)return map;
  var d=seg.getRange(1,1,seg.getLastRow(),seg.getLastColumn()).getValues(), h=d[0];
  function c(n){for(var x=0;x<h.length;x++){if((''+h[x]).toLowerCase().indexOf(n.toLowerCase())>-1)return x;}return -1;}
  var cId=c('Código'), cF=c('Fecha'), cV=c('Vía'), cCli=c('Interesado'), cTel=c('Tel'), cRes=c('Resultado'), cProx=c('Próximo'), cObs=c('Observ'), cFProx=c('próximo contacto'), cVend=c('Vendedor'), cCargo=c('Cargó');
  for(var r=1;r<d.length;r++){
    var id=(''+d[r][cId]).trim(); if(!id)continue;
    if(!map[id])map[id]=[];
    var f=d[r][cF];
    map[id].push({ _row:(r+1), fechaProx:cFProx>-1?(''+d[r][cFProx]):'', fecha:(f&&f.getMonth!==undefined)?Utilities.formatDate(f,'GMT-3','dd/MM/yyyy'):(''+f), _d:_pdias_(f),
      tipo:(''+d[r][cV])||'', cli:(''+d[r][cCli])||'', tel:cTel>-1?(''+d[r][cTel]):'', res:(''+d[r][cRes])||'', prox:cProx>-1?(''+d[r][cProx]):'', obs:cObs>-1?(''+d[r][cObs]):'', vend:cVend>-1?(''+d[r][cVend]).trim():'', cargo:cCargo>-1?(''+d[r][cCargo]).trim():'' });
  }
  return map;
}
function getDatos(){
  var h=hojaMaestra_();
  if(h.getLastRow()<2)return [];
  var data=h.getRange(1,1,h.getLastRow(),h.getLastColumn()).getValues();
  var head=data[0];
  var _cm={};
  function c(n){ if(n in _cm)return _cm[n]; var i=head.indexOf(n); if(i<0){ var nl=(''+n).toLowerCase().trim();
    for(var x=0;x<head.length;x++){ if((''+head[x]).toLowerCase().trim()===nl){ i=x; break; } } } return (_cm[n]=i); }
  var docCols=[];
  head.forEach(function(x,i){ if((''+x).indexOf('Doc:')===0) docCols.push({nombre:(''+x).replace('Doc:','').trim(), i:i}); });
  var verifCols=[];
  head.forEach(function(x,i){ if((''+x).indexOf('Verificación IA:')===0) verifCols.push({nombre:(''+x).replace('Verificación IA:','').trim(), i:i}); });
  var segMap=leerSeguimiento_();
  var encMap=_mapaEncuestas_();
  var hoy=new Date();
  var out=[];
  for(var r=1;r<data.length;r++){
    var row=data[r];
    var id=row[c('Código de Precarga')]; if(!id)continue;
    var duenos=[];
    for(var k=1;k<=5;k++){ var nom=row[c('Dueño '+k+' - Nombre y Apellido')];
      if(nom) duenos.push({ nombre:nom, dni:row[c('Dueño '+k+' - DNI')]||'', tel:row[c('Dueño '+k+' - Celular')]||'', mail:row[c('Dueño '+k+' - E-mail')]||'', dom:row[c('Dueño '+k+' - Domicilio')]||'' }); }
        var nOwn=duenos.length;
    var papeles={}, docUrls={};
    docCols.forEach(function(d){
      var mo=d.nombre.match(/Due[ñn]o\s+(\d)/i);
      if(mo && parseInt(mo[1])>nOwn) return;
      var raw=(''+row[d.i]).trim();
      var val=raw.toLowerCase();
      var esError=/no coincide|error|inv[áa]lido|no encontrado|no legible|no corresponde|nombre no/.test(val);
      papeles[d.nombre]=(val && !esError && !/pendiente|falta|^no$|^—$|^-$/.test(val))?1:0;
      if(/^https?:/i.test(raw)) docUrls[d.nombre]=raw;
    });
    // Asegurar que todos los docs de CONFIG aparezcan aunque la columna no exista aún
    CONFIG.DOCUMENTOS.forEach(function(d){
      var mo=d.label.match(/Due[ñn]o\s+(\d)/i);
      if(mo && parseInt(mo[1])>nOwn) return;
      if(!(d.label in papeles)) papeles[d.label]=0;
    });
    var verifIA={};
    verifCols.forEach(function(d){
      var mo=d.nombre.match(/Due[ñn]o\s+(\d)/i);
      if(mo && parseInt(mo[1])>nOwn) return;
      var raw=(''+row[d.i]).trim();
      if(raw) verifIA[d.nombre]=raw;
    });
    var ci=c('% Aviso Tokko');
    var pctS=(''+((ci>-1&&row[ci]!=null)?row[ci]:'')).replace('%','').trim();
    var _pn=parseInt(pctS); var pct=(pctS===''||isNaN(_pn))?null:_pn;
    var calle=row[c('Calle')]||'', num=row[c('Número')]||'';
    var seg=segMap[id]||[];
    var cons=seg.length;
    var vis=seg.filter(function(s){return /visita/i.test(s.tipo)||/visitó|oferta/i.test(s.res);}).length;
    var dias=null;
    seg.forEach(function(s){ if(s._d){ var dd=Math.floor((hoy-s._d)/86400000); if(dias===null||dd<dias)dias=dd; } });
    var segOut=seg.map(function(s){return {_row:s._row,fecha:s.fecha,tipo:s.tipo,cli:s.cli,tel:s.tel,res:s.res,prox:s.prox,obs:s.obs,fechaProx:s.fechaProx,vend:s.vend||'',cargo:s.cargo||''};});
    out.push({
      id:id, dir:(calle+' '+num).trim(), ciudad:row[c('Ciudad')]||'', tipo:row[c('Tipo de propiedad')]||'',
      etapa:row[c('Etapa')]||'Publicada', estadoCarga:row[c('Estado de la propiedad')]||'',
      precio:fmtPrecio_(row[c('Precio')], row[c('Moneda')]), operacion:row[c('Tipo de operación')]||'',
      pct:pct, plan:row[c('Plan Tokko')]||'',
            papeles:papeles, docUrls:docUrls, encuesta:(encMap[id]||null),
      origen:row[c('Origen')]||'Propia', colega:row[c('Inmobiliaria colega')]||'', colegaTel:row[c('Contacto colega')]||'',
      comision:row[c('Reparto de comisión')]||'', vendedor:row[c('Vendedor asignado')]||'', tokko:row[c('Link Tokko')]||'',
      carpeta:row[c('Carpeta Drive (URL)')]||'', motivo:row[c('Motivo de suspensión')]||'',
      obs:row[c('Observaciones comerciales')]||'',
      duenos:duenos, prop:duenos.length?duenos[0].nombre:'',
      seg:segOut, cons:cons, vis:vis, dias:dias
    });
  }
  return out;
}

function testGetDatos(){
  var d=getDatos();
  Logger.log('Propiedades leídas: '+d.length);
  Logger.log('Primera: '+JSON.stringify(d[0],null,2));
  Logger.log('Docs de la primera: '+JSON.stringify(d[0]?d[0].papeles:{}));
}
function corregirMigracion(){
  var ID_VENTAS='16-xykemcOS78QxNymU5Se7RGLwZqvxvXvZ1guhzUECQ';
  var ID_MAESTRO='1O8VKF9ZUC073kGNYIc3SgKPcIrVpnvynpQQ69RAk9W8';

  var ms=SpreadsheetApp.openById(ID_MAESTRO).getSheets(), mae=null;
  for(var j=0;j<ms.length;j++){
    if(ms[j].getRange(1,1,1,ms[j].getLastColumn()).getValues()[0].join('|').indexOf('Código de Precarga')>-1){mae=ms[j];break;}
  }
  if(!mae)throw new Error('No encontré la pestaña maestra');
  if(mae.getLastRow()>1){
    var cods=mae.getRange(2,1,mae.getLastRow()-1,1).getValues();
    for(var i=cods.length-1;i>=0;i--){ if((''+cods[i][0]).indexOf('MIG-')===0) mae.deleteRow(i+2); }
  }

  var sv=SpreadsheetApp.openById(ID_VENTAS).getSheets(), base=null, hdr=0;
  for(var k=0;k<sv.length;k++){
    var nf=Math.min(6, sv[k].getLastRow());
    if(nf<1)continue;
    var top=sv[k].getRange(1,1,nf,sv[k].getLastColumn()).getValues();
    for(var rr=0;rr<top.length;rr++){
      var line=top[rr].join('|').toLowerCase();
      if(line.indexOf('propietario')>-1 && line.indexOf('estado')>-1 && line.indexOf('direcc')>-1){ base=sv[k]; hdr=rr; break; }
    }
    if(base)break;
  }
  if(!base)throw new Error('No encontré la BASE con Propietario');

  var all=base.getDataRange().getValues();
  var vh=all[hdr];
  function col(n){for(var x=0;x<vh.length;x++){if((''+vh[x]).toLowerCase().indexOf(n.toLowerCase())>-1)return x;}return -1;}
  var cId=col('ID'),cDir=col('Direcc'),cTipo=col('Tipo'),cProp=col('Propietario'),cPrecio=col('Precio'),cEstado=col('Estado'),cObs=col('Observ');

  var mh=mae.getRange(1,1,1,mae.getLastColumn()).getValues()[0];
  function idx(n){return mh.indexOf(n);}
  var em={'en venta':'Publicada','reservada':'Reserva','vendida':'Vendida','suspendida':'Suspendida','alquilada':'Publicada'};
  var nuevas=[];
  for(var r=hdr+1;r<all.length;r++){
    var row=all[r];
    var id=(cId>-1?row[cId]:'')||'';
    var dir=(cDir>-1?row[cDir]:'')||'';
    if(!id && !dir)continue;
    if((''+id).toLowerCase().indexOf('id')===0)continue;

    var d=(''+dir).trim(), calle=d, num='';
    var m=d.match(/^(.*?)[\s,\.]*(\d+)\s*$/);
    if(m){calle=m[1].trim();num=m[2];}
    var dl=d.toLowerCase(), ciudad='Rosario';
    if(dl.indexOf('funes')>-1)ciudad='Funes'; else if(dl.indexOf('ibarluce')>-1)ciudad='Ibarlucea';

    var precioRaw=(''+(cPrecio>-1?row[cPrecio]:''));
    var precioNum=precioRaw.replace(/\.\d{2}$/,'').replace(/[^\d]/g,'');
    var moneda= precioNum? 'USD' : '';

    var estado=((cEstado>-1?row[cEstado]:'')||'').toString().trim().toLowerCase();

    var fila=new Array(mh.length).fill('');
    function set(n,v){var p=idx(n);if(p>-1)fila[p]=v;}
    set('Código de Precarga','MIG-'+id);
    set('Fecha de creación',Utilities.formatDate(new Date(),'GMT-3','dd/MM/yyyy'));
    set('Dueño 1 - Nombre y Apellido',((cProp>-1?row[cProp]:'')||'').toString().trim());
    set('Calle',calle); set('Número',num); set('Ciudad',ciudad);
    set('Tipo de propiedad',((cTipo>-1?row[cTipo]:'')||'').toString().trim());
    set('Precio',precioNum); set('Moneda',moneda); set('Tipo de operación','Venta');
    set('Etapa',em[estado]||'Publicada');
    set('Estado de la propiedad','Migrado - revisar'); set('Origen','Propia');
    set('Observaciones comerciales','Migrado '+id+'. Dir original: "'+d+'". '+((cObs>-1?row[cObs]:'')||''));
    nuevas.push(fila);
  }
  if(nuevas.length)mae.getRange(mae.getLastRow()+1,1,nuevas.length,mh.length).setValues(nuevas);
  SpreadsheetApp.getUi().alert('Migración corregida.\nMigradas ahora: '+nuevas.length);
}
// 1) % de Tokko (corregido: decimales -> %)
function llenarTokko(){
  var ID_VENTAS='16-xykemcOS78QxNymU5Se7RGLwZqvxvXvZ1guhzUECQ';
  var ms=SpreadsheetApp.openById(MAESTRO_ID).getSheets(), mae=null;
  for(var j=0;j<ms.length;j++){ if(ms[j].getRange(1,1,1,ms[j].getLastColumn()).getValues()[0].join('|').indexOf('Código de Precarga')>-1){mae=ms[j];break;} }
  if(!mae)throw new Error('No encontré la pestaña maestra');
  var mh=mae.getRange(1,1,1,mae.getLastColumn()).getValues()[0];
  var cCod=mh.indexOf('Código de Precarga'), cPct=mh.indexOf('% Aviso Tokko'), cPlan=mh.indexOf('Plan Tokko');
  if(cPct===-1||cPlan===-1)throw new Error('Faltan columnas Tokko; corré agregarColumnasTokko primero');
  var sv=SpreadsheetApp.openById(ID_VENTAS).getSheets(), portal=null, phdr=0;
  for(var k=0;k<sv.length;k++){
    var nf=Math.min(4, sv[k].getLastRow()); if(nf<1)continue;
    var top=sv[k].getRange(1,1,nf,sv[k].getLastColumn()).getValues();
    for(var rr=0;rr<top.length;rr++){ var line=top[rr].join('|').toLowerCase();
      if(line.indexOf('porcentaje')>-1 && line.indexOf('plan')>-1){ portal=sv[k]; phdr=rr; break; } }
    if(portal)break;
  }
  if(!portal)throw new Error('No encontré la pestaña del portal');
  var pv=portal.getDataRange().getValues(), ph=pv[phdr];
  function pcol(n){for(var x=0;x<ph.length;x++){if((''+ph[x]).toLowerCase().indexOf(n.toLowerCase())>-1)return x;}return -1;}
  var pId=pcol('ID'), pPct=pcol('porcentaje'), pPlan=pcol('plan');
  var mapa={};
  for(var r=phdr+1;r<pv.length;r++){
    var id=(''+ (pId>-1?pv[r][pId]:'')).trim().toUpperCase(); if(!id || id.indexOf('P')!==0)continue;
    var pctRaw=(''+(pPct>-1?pv[r][pPct]:'')).replace('%','').trim();
    var _num=parseFloat(pctRaw); if(!isNaN(_num)&&_num<=1)_num=_num*100;
    var pctN=(pctRaw===''||isNaN(_num))?'':Math.round(_num);
    mapa[id]={pct:pctN, plan:((pPlan>-1?pv[r][pPlan]:'')||'').toString().trim()};
  }
  var n=mae.getLastRow()-1; if(n<1)return;
  var cods=mae.getRange(2,cCod+1,n,1).getValues();
  var ady=Math.abs(cPct-cPlan)===1, tocadas=0;
  if(ady){
    var startCol=Math.min(cPct,cPlan)+1, pctFirst=cPct<cPlan, out=[];
    for(var i=0;i<n;i++){ var cod=(''+cods[i][0]).toUpperCase().replace('MIG-','').trim(); var m=mapa[cod]; if(m)tocadas++;
      var a=m?m.pct:'', b=m?m.plan:''; out.push(pctFirst?[a,b]:[b,a]); }
    mae.getRange(2,startCol,n,2).setValues(out);
  }
  SpreadsheetApp.getUi().alert('% de Tokko cargado. Propiedades actualizadas: '+tocadas);
}

// 2) Volcar SEGUIMIENTO (consultas/visitas) del sheet viejo al maestro
function volcarSeguimiento(){
  var ID_VENTAS='16-xykemcOS78QxNymU5Se7RGLwZqvxvXvZ1guhzUECQ';
  var ss=SpreadsheetApp.openById(MAESTRO_ID);
  var seg=ss.getSheetByName('Seguimiento');
  if(!seg)throw new Error('No existe la pestaña "Seguimiento" en el maestro (corré unificarMaestro)');
  if(seg.getLastRow()>1) seg.getRange(2,1,seg.getLastRow()-1,seg.getLastColumn()).clearContent();
  var sv=SpreadsheetApp.openById(ID_VENTAS).getSheets(), src=null, hdr=0;
  for(var k=0;k<sv.length;k++){
    var nf=Math.min(5, sv[k].getLastRow()); if(nf<1)continue;
    var top=sv[k].getRange(1,1,nf,sv[k].getLastColumn()).getValues();
    for(var rr=0;rr<top.length;rr++){ var line=top[rr].join('|').toLowerCase();
      if((line.indexOf('interesado')>-1||line.indexOf('cliente')>-1) && line.indexOf('resultado')>-1){ src=sv[k]; hdr=rr; break; } }
    if(src)break;
  }
  if(!src)throw new Error('No encontré la pestaña de seguimiento en el sheet viejo');
  var sd=src.getDataRange().getValues(), sh=sd[hdr];
  function sc(n){for(var x=0;x<sh.length;x++){if((''+sh[x]).toLowerCase().indexOf(n.toLowerCase())>-1)return x;}return -1;}
  var cId=sc('ID'), cFecha=sc('Fecha'), cVia=sc('contacto'), cCli=(sc('Interesado')>-1?sc('Interesado'):sc('Cliente')),
      cTel=sc('Tel'), cRes=sc('Resultado'), cProx=sc('Próximo'), cFProx=sc('próximo contacto'), cObs=sc('Observ');
  var out=[];
  for(var r=hdr+1;r<sd.length;r++){
    var row=sd[r];
    var pid=(''+ (cId>-1?row[cId]:'')).trim().toUpperCase();
    var cli=(''+ (cCli>-1?row[cCli]:'')).trim();
    if(!pid && !cli)continue;
    if(pid.indexOf('P')!==0)continue;
    out.push(['MIG-'+pid,(cFecha>-1?row[cFecha]:''),(cVia>-1?row[cVia]:''),cli,(cTel>-1?row[cTel]:''),(cRes>-1?row[cRes]:''),(cProx>-1?row[cProx]:''),(cFProx>-1?row[cFProx]:''),(cObs>-1?row[cObs]:''),'']);
  }
  if(out.length) seg.getRange(2,1,out.length,10).setValues(out);
  Logger.log('Seguimiento volcado. Interacciones cargadas: '+out.length);
}

// 3) Limpiar la nota "Migrado ..." de Observaciones
function limpiarObservaciones(){
  var mae=hojaMaestra_();
  var mh=mae.getRange(1,1,1,mae.getLastColumn()).getValues()[0];
  var c=mh.indexOf('Observaciones comerciales'); if(c===-1)return;
  var n=mae.getLastRow()-1; if(n<1)return;
  var rng=mae.getRange(2,c+1,n,1), vals=rng.getValues(), cambiadas=0;
  for(var i=0;i<n;i++){
    var v=''+vals[i][0];
    var nuevo=v.replace(/^Migrado\s+\S+\.\s*Dir original:\s*"[^"]*"\.\s*/,'').trim();
    if(nuevo!==v){ vals[i][0]=nuevo; cambiadas++; }
  }
  rng.setValues(vals);
  SpreadsheetApp.getUi().alert('Observaciones limpiadas: '+cambiadas);
}
function misEnlaces(){
  var url = ScriptApp.getService().getUrl(); // la URL /exec publicada
  var dev = url.replace(/\/exec$/, '/dev');
  var msg =
    '\n=== ENLACES DEL PROYECTO ===\n\n' +
    'FORM (precarga):\n' + url + '\n\n' +
    'TABLERO (para trabajar/probar):\n' + dev + '?app=tablero\n\n' +
    'TABLERO (oficial, para el cliente):\n' + url + '?app=tablero\n\n' +
    '============================';
  Logger.log(msg);
}
function verLinea162(){
  var html = HtmlService.createHtmlOutputFromFile('Tablero').getContent();
  var lineas = html.split('\n');
  for (var i = 158; i <= 166 && i < lineas.length; i++){
    Logger.log('L' + (i+1) + ': ' + lineas[i].substring(0,300));
  }
}
function verServido(){
  var html = HtmlService.createHtmlOutputFromFile('Tablero').getContent();
  var lineas = html.split('\n');
  Logger.log('TOTAL LÍNEAS: ' + lineas.length);
  for (var i = 155; i <= 170 && i < lineas.length; i++){
    Logger.log('L' + (i+1) + ' [' + lineas[i].length + ' chars]: ' + lineas[i].substring(0,200));
  }
}
function partirScript(){
  var html = HtmlService.createHtmlOutputFromFile('Tablero').getContent();
  var ini = html.indexOf('<script>');
  var fin = html.lastIndexOf('</script>');
  var js = html.substring(ini+8, fin);
  Logger.log('largo del script: ' + js.length);
  Logger.log('mitad 1 (primeros 200): ' + js.substring(0,200));
  Logger.log('¿hay backslash-u raro?: ' + /\\u[0-9a-f]{4}/i.test(js));
  Logger.log('¿hay <script dentro?: ' + (js.indexOf('<script') > -1));
  Logger.log('¿hay </script dentro?: ' + (js.indexOf('</script') > -1));
}
var MAP_CAMPOS = {
  precio:'Precio', tipo:'Tipo de propiedad', obs:'Observaciones comerciales',
  tokko:'Link Tokko', prop:'Dueño 1 - Nombre y Apellido', tel:'Dueño 1 - Celular',
  mail:'Dueño 1 - E-mail', dni:'Dueño 1 - DNI', dom:'Dueño 1 - Domicilio',
  colega:'Inmobiliaria colega', colegaTel:'Contacto colega', comision:'Reparto de comisión',
  origen:'Origen', pct:'% Aviso Tokko', vendedor:'Vendedor asignado', motivo:'Motivo de suspensión'
};
function _filaPorCodigo_(mae, codigo){
  var col=mae.getRange(1,1,1,mae.getLastColumn()).getValues()[0].indexOf('Código de Precarga')+1;
  if(col<1)throw new Error('No hay columna Código de Precarga');
  var n=mae.getLastRow()-1; if(n<1)return -1;
  var cods=mae.getRange(2,col,n,1).getValues();
  for(var i=0;i<n;i++){ if((''+cods[i][0]).trim()===(''+codigo).trim()) return i+2; }
  return -1;
}
function guardarPropiedad(codigo, cambios){
  var mae=hojaMaestra_();
  var fila=_filaPorCodigo_(mae, codigo);
  if(fila<0)throw new Error('No encontré la propiedad '+codigo);
  var head=mae.getRange(1,1,1,mae.getLastColumn()).getValues()[0];
  var quien=String((cambios&&(cambios.quien||cambios.cargo))||'').trim().slice(0,40), difs=[], etapaDe='', etapaA='';
  var ETQ={precio:'Precio',tipo:'Tipo',obs:'Observaciones',tokko:'URL de ficha',origen:'Origen',colega:'Colega',colegaTel:'Contacto colega',comision:'Comisión',pct:'% aviso',vendedor:'Vendedor',motivo:'Motivo de suspensión',prop:'Propietario',tel:'Tel. propietario',mail:'Mail propietario',dni:'DNI propietario',dom:'Domicilio propietario'};
  function leer(col){ var i=head.indexOf(col); return i>-1 ? String(mae.getRange(fila,i+1).getValue()==null?'':mae.getRange(fila,i+1).getValue()).trim() : ''; }
  function set(colName, val){ var i=head.indexOf(colName); if(i>-1) mae.getRange(fila,i+1).setValue(val); }
  Object.keys(cambios||{}).forEach(function(campo){
    if(campo==='papeles'){
      var pap=cambios.papeles||{};
      Object.keys(pap).forEach(function(doc){
        var antes=leer('Doc: '+doc), antesOk=!!antes && !/^(pendiente|falta|no|—|-)$/i.test(antes);
        if(!!pap[doc]!==antesOk) difs.push('«'+doc+'» '+(pap[doc]?'marcado como cargado':'marcado como pendiente'));
        set('Doc: '+doc, pap[doc]?'Cargado':'Pendiente');
      });
    } else if(campo==='etapa'){
      var et=cambios.etapa==='Publicadas'?'Publicada':cambios.etapa;
      var ea=leer('Etapa')||'Publicada';
      if(ea!==et){ difs.push('Etapa: '+ea+' → '+et); etapaDe=ea; etapaA=et; }
      set('Etapa', et);
    } else if(MAP_CAMPOS[campo]){
      var col=MAP_CAMPOS[campo], av=leer(col), nv=(cambios[campo]==null?'':String(cambios[campo])).trim();
      if(av.replace(/\s+/g,' ')!==nv.replace(/\s+/g,' ') && ETQ[campo] && !(av==='' && (nv==='' || (campo==='origen' && nv==='Propia')))) difs.push(ETQ[campo]+': '+(av||'—')+' → '+(nv||'—'));
      set(col, cambios[campo]);
    }
  });
  set('Última actualización', new Date());
  var dir=_dirDe_(mae,fila,head);
  if(difs.length) _audit_(etapaA?'etapa':'campos', (etapaA?'🔶 ':'✏️ ')+difs.join('; ').slice(0,200), dir, quien);
  if(etapaA && !(typeof AZCU_NOAUDIT_!=='undefined' && AZCU_NOAUDIT_)){ try{ if(typeof azcuNotificarEtapa_==='function') azcuNotificarEtapa_(codigo, dir, etapaDe, etapaA, quien); }catch(e){} }
  try{ if(typeof azcuLimpiarCache_==='function') azcuLimpiarCache_(); }catch(e){}
  return {ok:true};
}
function _audit_(tipo, texto, dir, quien){ try{ if(typeof azcuAudit_==='function') azcuAudit_(tipo, texto, dir, quien); }catch(e){} }
function _dirDe_(mae, fila, head){ try{ return (String(mae.getRange(fila, head.indexOf('Calle')+1).getValue())+' '+String(mae.getRange(fila, head.indexOf('Número')+1).getValue())).trim(); }catch(e){ return ''; } }
function _invalidaNovedades_(){ try{ var c=CacheService.getScriptCache(); c.remove('azmov_90'); c.remove('azmov_30'); }catch(e){} }
function registrarInteraccion(codigo, datos){
  var ss=SpreadsheetApp.openById(MAESTRO_ID);
  var seg=ss.getSheetByName('Seguimiento');
  if(!seg)throw new Error('No existe la pestaña Seguimiento');
  var fecha=datos.fecha||Utilities.formatDate(new Date(),'GMT-3','dd/MM/yyyy');
  var fechaProx=datos.fechaProx||'';
  // si el resultado es positivo, crear recordatorio en Calendar a los 2 dias
  if(/Le interesó|Le intereso/i.test(datos.resultado||'')){
    try{
      var rec=new Date();
      rec.setDate(rec.getDate()+2);
      rec.setHours(10,0,0,0);
      var mae=hojaMaestra_();
      var fila=_filaPorCodigo_(mae,codigo);
      var head=mae.getRange(1,1,1,mae.getLastColumn()).getValues()[0];
      var calle=fila>0?mae.getRange(fila,head.indexOf('Calle')+1).getValue():'';
      var num=fila>0?mae.getRange(fila,head.indexOf('Número')+1).getValue():'';
      var dir=(calle+' '+num).trim()||codigo;
      var titulo='Seguimiento: '+(datos.interesado||'interesado')+' - '+dir;
      CalendarApp.getDefaultCalendar().createEvent(titulo,rec,new Date(rec.getTime()+30*60*1000),{
        description:'Contactar a '+(datos.interesado||'')+'\nTel: '+(datos.tel||'')+'\nPropiedad: '+codigo+' - '+dir+'\nLe interesó en la visita/consulta del '+fecha
      });
      fechaProx=Utilities.formatDate(rec,'GMT-3','dd/MM/yyyy');
    }catch(e){ Logger.log('Error Calendar recordatorio: '+e.message); }
  }
  seg.appendRow([
    codigo, fecha, datos.via||'', datos.interesado||'', datos.tel||'',
    datos.resultado||'', datos.prox||'', fechaProx, datos.obs||'', datos.vendedor||'', datos.cargo||''
  ]);
  try{ if(!String(seg.getRange(1,11).getValue()).trim()) seg.getRange(1,11).setValue('Cargó'); }catch(e){}
  _invalidaNovedades_();
  return {ok:true};
}
function editarInteraccion(filaSeg, datos){
  var ss=SpreadsheetApp.openById(MAESTRO_ID);
  var seg=ss.getSheetByName('Seguimiento');
  if(!seg || filaSeg<2)throw new Error('Fila inválida');
  var h=seg.getRange(1,1,1,seg.getLastColumn()).getValues()[0];
  function set(colName,val){var i=h.indexOf(colName);if(i>-1 && val!==undefined)seg.getRange(filaSeg,i+1).setValue(val);}
  set('Fecha',datos.fecha); set('Vía',datos.via); set('Interesado',datos.interesado);
  set('Teléfono',datos.tel); set('Resultado',datos.resultado); set('Próximo paso',datos.prox);
  set('Observaciones',datos.obs);
  try{ var rowE=seg.getRange(filaSeg,1,1,Math.max(seg.getLastColumn(),4)).getValues()[0]; _audit_('campos','✏️ Editó la interacción de '+(datos.interesado||rowE[3]||'—')+' ('+(datos.resultado||rowE[5]||'')+')', (typeof azcuMapaDirs_==='function'?(azcuMapaDirs_()[String(rowE[0]).trim()]||''):''), datos.cargo||datos.quien||''); }catch(e){}
  _invalidaNovedades_();
  return {ok:true};
}
function eliminarPropiedad(codigo, quien){
  var mae=hojaMaestra_();
  var fila=_filaPorCodigo_(mae, codigo);
  if(fila<0)throw new Error('No encontré la propiedad '+codigo);
  try{ var hE=mae.getRange(1,1,1,mae.getLastColumn()).getValues()[0]; var dE=_dirDe_(mae,fila,hE); _audit_('elimprop','🗑 Eliminó la propiedad '+(dE||codigo), dE, quien||''); }catch(e){}
  mae.deleteRow(fila);
  return {ok:true};
}
function testGuardar(){
  Logger.log(JSON.stringify(guardarPropiedad('MIG-P001', {obs:'prueba desde código'})));
}
function subirDocumento(codigo, docNombre, base64, filename, mime, quien){
  var mae=hojaMaestra_();
  var fila=_filaPorCodigo_(mae, codigo);
  if(fila<0)throw new Error('No encontré la propiedad '+codigo);
  var head=mae.getRange(1,1,1,mae.getLastColumn()).getValues()[0];
  var colCarpeta=head.indexOf('Carpeta Drive (URL)');
  var urlCarpeta=colCarpeta>-1?mae.getRange(fila,colCarpeta+1).getValue():'';
  var idc=(''+urlCarpeta).match(/[-\w]{25,}/);
  var carpeta;
  if(idc){ carpeta=DriveApp.getFolderById(idc[0]); }
  else {
    var calle=mae.getRange(fila, head.indexOf('Calle')+1).getValue();
    var num=mae.getRange(fila, head.indexOf('Número')+1).getValue();
    carpeta=_getOrCreatePropertyFolder(codigo, calle, num);
    if(colCarpeta>-1) mae.getRange(fila,colCarpeta+1).setValue(carpeta.getUrl());
  }
  var bytes=Utilities.base64Decode(base64);
  var blob=Utilities.newBlob(bytes, mime||'application/octet-stream', (docNombre+' - '+(filename||'')).trim());
  var archivo=carpeta.createFile(blob);
  var i=head.indexOf('Doc: '+docNombre);
  if(i>-1) mae.getRange(fila,i+1).setValue(archivo.getUrl());

  // Verificación con IA (misma lógica que del lado de Precarga)
  var doc = _docPorLabel(docNombre);
  if (doc) {
    var nombresDueños = {};
    for (var n = 1; n <= CONFIG.MAX_DUEÑOS; n++) {
      var colNom = head.indexOf('Dueño ' + n + ' - Nombre y Apellido');
      if (colNom > -1) {
        var vNom = mae.getRange(fila, colNom + 1).getValue();
        if (vNom) nombresDueños[n] = vNom;
      }
    }
    var infoIA = _infoEsperadaDocumento(doc.key, nombresDueños);
    var resultadoIA = _verificarDocumentoIA(blob, infoIA.tipo, infoIA.nombre);
    var colVerif = head.indexOf('Verificación IA: ' + docNombre);
    if (colVerif > -1) mae.getRange(fila, colVerif + 1).setValue(resultadoIA.estado);

    // Si lo que se subió es la Escritura, además de verificarla aprovechamos
    // para extraer la Inscripción del dominio (dato clave para el Generador
    // de documentos) y guardarla en su columna, sin pisarla si ya tenía algo
    // y la IA no pudo leer nada nuevo.
    if (doc.key === 'escritura') {
      var colInscripcion = head.indexOf('Inscripción del dominio');
      if (colInscripcion > -1) {
        var datosEscritura = extraerDatosEscritura({
          base64: base64,
          mimeType: mime || 'application/octet-stream',
          filename: filename || ''
        });
        if (datosEscritura.ok && datosEscritura.inscripcionDominio) {
          mae.getRange(fila, colInscripcion + 1).setValue(datosEscritura.inscripcionDominio);
        }
      }
    }
  }

  _audit_('doc','📄 Subió «'+docNombre+'»', _dirDe_(mae,fila,head), quien||'');
  return {ok:true, url:archivo.getUrl()};
}

function borrarDocumento(codigo, docNombre, quien){
  var mae=hojaMaestra_();
  var fila=_filaPorCodigo_(mae, codigo);
  if(fila<0)throw new Error('No encontré la propiedad '+codigo);
  var head=mae.getRange(1,1,1,mae.getLastColumn()).getValues()[0];
  var i=head.indexOf('Doc: '+docNombre);
  if(i<0)throw new Error('No existe ese documento');
  var val=(''+mae.getRange(fila,i+1).getValue()).trim();
  var m=val.match(/[-\w]{25,}/);
  if(m){ try{ DriveApp.getFileById(m[0]).setTrashed(true); }catch(e){} }
  mae.getRange(fila,i+1).setValue('Pendiente');
  _audit_('borrardoc','📄 Borró el documento «'+docNombre+'»', _dirDe_(mae,fila,head), quien||'');
  return {ok:true};
}
function testSubir(){
  try{
    var f=DriveApp.getFolderById('1wBjd8B8ZF-nbbTD67yq_os7Z3J386Wty');
    Logger.log('CARPETA OK: '+f.getName());
  }catch(e){
    Logger.log('FALLA CARPETA: '+e.message);
    return;
  }
  try{
    var r=subirDocumento('MIG-P001','Escritura', Utilities.base64Encode('prueba'), 'test.txt', 'text/plain');
    Logger.log('SUBIDA OK: '+JSON.stringify(r));
  }catch(e){
    Logger.log('FALLA SUBIDA: '+e.message);
  }
}
function testMotivo(){
  Logger.log(JSON.stringify(guardarPropiedad('MIG-P001', {motivo:'prueba motivo'})));
}
function testSubir2(){
  try{ Logger.log('RAIZ: '+DriveApp.getFolderById('1wBjd8B8ZF-nbbTD67yq_os7Z3J386Wty').getName()); }
  catch(e){ Logger.log('FALLA RAIZ: '+e.message); }
  try{ Logger.log(JSON.stringify(subirDocumento('MIG-P001','Escritura',Utilities.base64Encode('x'),'t.txt','text/plain'))); }
  catch(e){ Logger.log('FALLA SUBIDA en: '+e.message); }
}
function obtenerPropiedades(){
  var mae=hojaMaestra_();
  if(mae.getLastRow()<2)return [];
  var data=mae.getRange(1,1,mae.getLastRow(),mae.getLastColumn()).getValues();
  var h=data[0];
  function c(n){return h.indexOf(n);}
  var cCod=c('Código de Precarga'),cCalle=c('Calle'),cNum=c('Número'),
      cProp=c('Dueño 1 - Nombre y Apellido'),cDni=c('Dueño 1 - DNI'),cDom=c('Dueño 1 - Domicilio');
  var out=[];
  for(var i=1;i<data.length;i++){
    var r=data[i];var id=r[cCod];if(!id)continue;
    var dir=((cCalle>-1?r[cCalle]:'')+' '+(cNum>-1?r[cNum]:'')).trim();
    out.push({
      id:id,
      etiqueta:(dir||'(sin dirección)')+' — '+id,
      direccion:dir,
      propietario:cProp>-1?(''+r[cProp]):'',
      dni_propietario:cDni>-1?(''+r[cDni]):'',
      domicilio_propietario:cDom>-1?(''+r[cDom]):'',
      'matrícula':''
    });
  }
  return out;
}

function obtenerImagenes(){
  return { logo:'', sello:'' };
}

function guardarDocumento(nombre, base64){
  try{
    var carpeta=DriveApp.getFolderById('1AGXKG5UN_GNlrosBojgNHhqDBaL9Z12m');
    var it=carpeta.getFoldersByName('Documentos generados');
    var dest=it.hasNext()?it.next():carpeta.createFolder('Documentos generados');
    var bytes=Utilities.base64Decode(base64);
    var blob=Utilities.newBlob(bytes,'application/pdf',nombre);
    var f=dest.createFile(blob);
    return {ok:true, nombre:nombre, url:f.getUrl()};
  }catch(e){
    return {ok:false, error:e.message};
  }
}
function getAppUrl(){ return ScriptApp.getService().getUrl(); }
function testProps(){
  var p=obtenerPropiedades();
  Logger.log('cantidad: '+p.length);
  Logger.log('primera: '+JSON.stringify(p[0]));
}
function guardarEncuesta(datos){
  var ss=SpreadsheetApp.openById(MAESTRO_ID);
  var h=ss.getSheetByName('Encuestas');
  if(!h){
    h=ss.insertSheet('Encuestas');
    h.appendRow(['Código de Precarga','Fecha','Nombre','Operación','Calificación','Equipo','Tiempos respuesta','Tiempos operación','Recomendación','Comentarios']);
  }
  h.appendRow([
    datos.codigo||'', Utilities.formatDate(new Date(),'GMT-3','dd/MM/yyyy'),
    datos.nombre||'', datos.operacion||'', datos.calificacion||'', datos.equipo||'',
    datos.tiempos_respuesta||'', datos.tiempos_operacion||'', datos.recomendacion||'', datos.comentarios||''
  ]);
  return {ok:true};
}
function obtenerEncuesta(codigo){
  var ss=SpreadsheetApp.openById(MAESTRO_ID);
  var h=ss.getSheetByName('Encuestas');
  if(!h || h.getLastRow()<2) return null;
  var d=h.getRange(1,1,h.getLastRow(),h.getLastColumn()).getValues();
  var hd=d[0];
  for(var i=d.length-1;i>=1;i--){
    if((''+d[i][0]).trim()===(''+codigo).trim()){
      var o={};
      hd.forEach(function(k,j){
        var v=d[i][j];
        if(v && v.getMonth) v=Utilities.formatDate(v,'GMT-3','dd/MM/yyyy');
        o[''+k]=(v==null?'':(''+v));
      });
      return o;
    }
  }
  return null;
}
function getEncuestaUrl(){ return ScriptApp.getService().getUrl(); }
function testEnc(){
  var ss=SpreadsheetApp.openById(MAESTRO_ID);
  Logger.log('Pestañas: '+ss.getSheets().map(function(h){return h.getName();}).join(' | '));
  Logger.log('Resultado: '+JSON.stringify(obtenerEncuesta('PRE-2026-0004')));
}
function testTablero(){
  var h=HtmlService.createHtmlOutputFromFile('Tablero').getContent();
  Logger.log('tiene boton Encuesta: ' + (h.indexOf('>Encuesta</button>')>-1));
  Logger.log('largo: ' + h.length);
}
function testEnc2(){
  Logger.log('busco PRE-2026-0004: ' + JSON.stringify(obtenerEncuesta('PRE-2026-0004')));
}
function testFinal(){
  var r = obtenerEncuesta('PRE-2026-0004');
  Logger.log(r ? 'ENCONTRÓ: '+r.Nombre : 'NULL - no encontró');
}
function _mapaEncuestas_(){
  var ss=SpreadsheetApp.openById(MAESTRO_ID);
  var h=ss.getSheetByName('Encuestas');
  var m={};
  if(!h || h.getLastRow()<2) return m;
  var d=h.getRange(2,1,h.getLastRow()-1,h.getLastColumn()).getValues();
  for(var i=0;i<d.length;i++){
    var cod=(''+d[i][0]).trim();
    if(cod) m[cod]={cal:d[i][4], nombre:(''+d[i][2]), rec:(''+d[i][8])};
  }
  return m;
}
function agendarVisita(codigo, datos){
  var cal=CalendarApp.getDefaultCalendar();
  var ini=new Date(datos.cuando);
  var fin=new Date(ini.getTime()+60*60*1000);
  var titulo='Visita: '+(datos.dir||codigo)+(datos.interesado?(' - '+datos.interesado):'');
  var opts={
    description:'Propiedad: '+codigo+'\nInteresado: '+(datos.interesado||'')+'\nTel: '+(datos.tel||''),
    location:datos.dir||''
  };
  if(datos.mail) opts.guests=datos.mail;
  var ev=cal.createEvent(titulo, ini, fin, opts);
  ev.addPopupReminder(60);
  registrarInteraccion(codigo, {
    fecha: Utilities.formatDate(ini,'GMT-3','dd/MM/yyyy'),
    via: 'Visita',
    interesado: datos.interesado||'',
    tel: datos.tel||'',
    resultado: 'Visita agendada',
    prox: 'Visita '+Utilities.formatDate(ini,'GMT-3','dd/MM/yyyy HH:mm'),
    obs: datos.obs||'',
    vendedor: datos.vendedor||'',
    cargo: datos.quien||datos.cargo||''
  });
  return {ok:true, eventId:ev.getId()};
}

function eliminarInteraccion(filaSeg, quien){
  var ss=SpreadsheetApp.openById(MAESTRO_ID);
  var seg=ss.getSheetByName('Seguimiento');
  if(!seg || filaSeg<2)throw new Error('Fila inválida');
  try{ var rowD=seg.getRange(filaSeg,1,1,Math.max(seg.getLastColumn(),6)).getValues()[0]; _audit_('elimint','🗑 Eliminó la interacción de '+(rowD[3]||'—')+' ('+(rowD[2]||'')+', '+(rowD[5]||'')+')', (typeof azcuMapaDirs_==='function'?(azcuMapaDirs_()[String(rowD[0]).trim()]||''):''), quien||''); }catch(e){}
  seg.deleteRow(filaSeg);
  _invalidaNovedades_();
  return {ok:true};
}

function registrarPropuesta(codigo, datos){
  registrarInteraccion(codigo, {
    fecha: datos.fecha,
    via: 'Propuesta',
    interesado: datos.cliente,
    tel: datos.tel||'',
    resultado: 'Hizo propuesta',
    prox: 'Responder propuesta',
    obs: 'Propuesta: '+(datos.monto||'') + (datos.obs?(' — '+datos.obs):''),
    vendedor: datos.vendedor||'',
    cargo: datos.quien||datos.cargo||''
  });
  return {ok:true};
}
function testDocPermiso(){
  var doc = DocumentApp.create('test permiso');
  doc.getBody().setText('ok');
  Logger.log('permiso OK: ' + doc.getUrl());
  DriveApp.getFileById(doc.getId()).setTrashed(true);
}
function testUltimaInteraccion(){
  var ss=SpreadsheetApp.openById(MAESTRO_ID);
  var seg=ss.getSheetByName('Seguimiento');
  var ult=seg.getRange(seg.getLastRow(),1,1,seg.getLastColumn()).getValues()[0];
  Logger.log(JSON.stringify(ult));
}
function testRecordatorio(){
  try{
    var rec=new Date();
    rec.setDate(rec.getDate()+2);
    rec.setHours(10,0,0,0);
    var ev=CalendarApp.getDefaultCalendar().createEvent(
      'Test recordatorio',
      rec,
      new Date(rec.getTime()+30*60*1000)
    );
    Logger.log('OK: '+ev.getTitle()+' | '+rec);
  }catch(e){
    Logger.log('ERROR: '+e.message);
  }
}
function testTipos(){
  var d=getDatos();
  var tipos=d.map(function(p){return p.tipo;}).filter(function(t){return t;});
  Logger.log('Total props: '+d.length+' | Con tipo: '+tipos.length);
  Logger.log('Tipos únicos: '+JSON.stringify([...new Set(tipos)]));
}