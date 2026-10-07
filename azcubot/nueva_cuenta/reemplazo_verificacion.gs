function _verificarArchivosConIA(archivos, nombresDueños) {
  const resultado = {};
  if (!archivos) return resultado;
  const apiKey = _openAiApiKey();
  const claves = [], reqs = [];
  Object.keys(archivos).forEach(key => {
    const blob = archivos[key];
    if (!blob) return;
    const info = _infoEsperadaDocumento(key, nombresDueños);
    const mime = blob.getContentType();
    if (!apiKey) { resultado[key] = 'No verificado (falta configurar OPENAI_API_KEY)'; return; }
    if (!_tipoSoportadoPorIA(mime)) { resultado[key] = 'No verificado (formato no soportado por la IA)'; return; }
    const dataUri = 'data:' + mime + ';base64,' + Utilities.base64Encode(blob.getBytes());
    const content = [{ type: 'input_text', text: _instruccionesVerificacion_(info.tipo, info.nombre) }];
    content.push(mime === 'application/pdf'
      ? { type: 'input_file', filename: blob.getName() || 'documento.pdf', file_data: dataUri }
      : { type: 'input_image', image_url: dataUri });
    claves.push(key);
    reqs.push({
      url: 'https://api.openai.com/v1/responses',
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + apiKey },
      payload: JSON.stringify({ model: 'gpt-4o-mini', input: [{ role: 'user', content: content }] }),
      muteHttpExceptions: true
    });
  });
  if (!reqs.length) return resultado;
  let resps;
  try {
    resps = UrlFetchApp.fetchAll(reqs);
  } catch (e) {
    claves.forEach(k => { resultado[k] = 'No verificado (error de red con OpenAI)'; });
    return resultado;
  }
  resps.forEach((r, i) => {
    let estado;
    try {
      if (r.getResponseCode() < 200 || r.getResponseCode() >= 300) throw new Error('error OpenAI ' + r.getResponseCode());
      const datos = _parsearJsonIA(_extraerTextoRespuestaOpenAI(JSON.parse(r.getContentText())));
      if (!datos) throw new Error('respuesta no interpretable');
      estado = _estadoVerificacion_(datos);
    } catch (e) {
      estado = 'No verificado (' + e.message + ')';
    }
    resultado[claves[i]] = estado;
  });
  return resultado;
}

function _instruccionesVerificacion_(tipoEsperado, nombreEsperado) {
  return 'Sos un asistente que verifica documentación para una inmobiliaria. ' +
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
}

function _estadoVerificacion_(d) {
  if (d.legible === false) return 'Ilegible / mal escaneo';
  if (d.coincide_tipo === false) return 'Documento incorrecto' + (d.tipo_detectado ? ' (parece: ' + d.tipo_detectado + ')' : '');
  if (d.nombre_coincide === false) return 'Nombre no coincide' + (d.nombre_detectado ? ' (dice: ' + d.nombre_detectado + ')' : '');
  return 'Verificado OK';
}
