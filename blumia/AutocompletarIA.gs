/* Autocompletado con IA de las entrevistas (formulario de competencias, y dictado/redacción en los cuadros de notas).
   Todo exige el PIN de los selectores (pinValido_): gasta crédito y maneja audios de candidatos. El audio no se guarda. */

function iaCacheKey_(pref, perfilId, cand) {
  const raw = String(perfilId || '') + '|' + String(cand || '').toLowerCase().trim();
  const h = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, raw, Utilities.Charset.UTF_8)
    .map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
  return pref + h;
}

function transcribirAudio_(pin, perfilId, cand, b64, mime) {
  pinValido_(pin);
  if (!b64 || String(b64).length > 33000000) throw new Error('El audio está vacío o es demasiado largo');
  const tipo = String(mime || 'audio/webm').split(';')[0];
  const ext = /mp4|m4a/.test(tipo) ? 'm4a' : /ogg/.test(tipo) ? 'ogg' : /wav/.test(tipo) ? 'wav' : 'webm';
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), tipo, 'audio.' + ext);
  const pedir = modelo => UrlFetchApp.fetch('https://api.openai.com/v1/audio/transcriptions', {
    method: 'post', muteHttpExceptions: true, headers: { Authorization: 'Bearer ' + OPENAI_API_KEY },
    payload: { file: blob, model: modelo, language: 'es', prompt: 'Entrevista de selección de personal por competencias, en castellano rioplatense.' }
  });
  let parsed = JSON.parse(pedir('gpt-4o-mini-transcribe').getContentText());
  if (parsed.error) parsed = JSON.parse(pedir('whisper-1').getContentText());
  if (parsed.error) throw new Error('No se pudo transcribir: ' + parsed.error.message);
  const texto = String(parsed.text || '').trim();
  if (perfilId && cand) guardarTransParcial_(perfilId, cand, texto);
  return texto;
}

// Respaldo temporal (6 h) de lo transcripto, por si la pantalla se cierra antes de terminar. No se guarda el audio.
function guardarTransParcial_(perfilId, cand, t) {
  try {
    const c = CacheService.getScriptCache(), k = iaCacheKey_('tr_', perfilId, cand);
    let a = [];
    try { a = JSON.parse(c.get(k) || '[]'); } catch (e) { }
    a.push(String(t || ''));
    let j = JSON.stringify(a);
    while (j.length > 40000 && a.length > 1) { a.shift(); j = JSON.stringify(a); }
    c.put(k, j, 21600);
  } catch (e) { Logger.log('respaldo de transcripción: ' + e); }
}

function recuperarEntrevista_(pin, perfilId, cand) {
  pinValido_(pin);
  const c = CacheService.getScriptCache();
  let an = null, parcial = [];
  try { an = JSON.parse(c.get(iaCacheKey_('an_', perfilId, cand)) || 'null'); } catch (e) { }
  try { parcial = JSON.parse(c.get(iaCacheKey_('tr_', perfilId, cand)) || '[]'); } catch (e) { }
  return JSON.stringify({ analisis: an, parcial: parcial.join(' ') });
}

function promptAnalisisEntrevista_(hero, nombre, comps) {
  return 'Sos un analista senior de selección de personal de Blumia Consulting. Te paso la transcripción de una entrevista por competencias (metodología STAR) para el puesto «' + String(hero.posicion || '') + '»' + (hero.empresa ? ' de ' + hero.empresa : '') + '. La transcripción no indica quién habla: inferilo por el contexto (las preguntas son de quien entrevista, las respuestas de la persona entrevistada). Evaluá SOLO con lo que dijo la persona entrevistada.\n' +
    'PERSONA: ' + String(nombre || '') + '. Nombrala por su nombre de pila en lugar de «el candidato» o «la candidata» y usá el género gramatical que corresponde a su nombre; si es ambiguo, formas neutras.\n' +
    'COMPETENCIAS: ' + JSON.stringify(comps) + '\n' +
    'NIVELES: 0 = no hubo evidencia en la entrevista; 1 = No demuestra; 2 = Señales de desarrollo; 3 = Demuestra efectividad; 4 = Demuestra solidez.\n\n' +
    'CÓMO ESCRIBIR LA «evidencia» DE CADA COMPETENCIA (es lo más importante):\n' +
    '- Es el relato concreto de LO QUE LA PERSONA CONTÓ, no una definición. PROHIBIDO repetir, parafrasear o resumir la descripción de la competencia y PROHIBIDAS las fórmulas genéricas («demuestra capacidad de…», «tiene habilidad para…», «se destaca por…»).\n' +
    '- Tiene que nombrar el EJEMPLO ESPECÍFICO que dio: cuál fue la situación o el proyecto, en qué empresa, cuenta o momento, qué hizo ella/él puntualmente (acciones, decisiones, a quién involucró, cómo lo encaró), qué resultado obtuvo (con cifras, plazos, montos o nombres si los dijo) y, si lo mencionó, qué aprendió y si volvió a aplicarlo.\n' +
    '- Extensión: 3 a 6 oraciones cuando hubo ejemplos. Si dio más de un ejemplo para esa competencia, desarrollá cada uno por separado. Usá los datos tal como los dijo (empresas, clientes, proyectos, cifras, herramientas, cargos); no los inventes, no los redondees y no agregues detalles que no estén en la transcripción.\n' +
    '- Atribuí siempre al relato («contó que», «relató», «mencionó que», «según dijo»): lo que se sabe de su trabajo es lo que la persona CONTÓ, no algo que alguien haya visto. Lo observado en la propia entrevista (cómo se expresó, su actitud) se puede decir directamente.\n' +
    '- Si habló de la competencia solo en términos generales y SIN ejemplo concreto, dejalo asentado («Habló en términos generales de… sin dar un ejemplo concreto») y asigná como máximo nivel 2.\n' +
    '- Si no surgió, evidencia = «No surgió evidencia en la entrevista» y nivel 0.\n' +
    '- Si hubo contradicciones o dudas en lo que contó, anotalas como hecho.\n' +
    'EJEMPLO DE FORMATO (inventado, NO uses su contenido): MAL → «Demuestra capacidad de negociación con clientes.» BIEN → «Contó que en 2022, en su rol en Constructora Norte, renegoció con el cliente principal un contrato por USD 400 mil que estaba en riesgo: armó una propuesta con precios escalonados, la presentó al gerente de compras y logró renovarlo por dos años con una mejora del 8 % en el margen. Mencionó que desde entonces prepara siempre una alternativa de precios antes de cada reunión.»\n\n' +
    'REGLAS GENERALES: tono objetivo y neutral, sin adjetivos elogiosos. No interpretes ni deduzcas rasgos o motivaciones: informá hechos. Los datos de la persona (meta) completalos SOLO si los dijo explícitamente; si no, dejá "". «fortalezas» y «areas»: viñetas cortas que empiecen con «• » y se basen en hechos y ejemplos concretos de la entrevista (no en generalidades). «recomendacion»: 3 a 5 oraciones DESCRIPTIVAS con lo observado (cómo se desenvolvió, cuáles fueron sus ejemplos más sólidos y qué no surgió) y, solo si las hubo, alertas concretas dichas como hecho; coherente con los niveles que asignaste; sin decidir si contratar o descartar.\n' +
    'Devolvé SOLO un JSON válido, sin texto antes ni después, con este formato: {"comps":[{"id":"","nivel":0,"evidencia":""}],"fortalezas":"","areas":"","recomendacion":"","meta":{"edad":"","localidad":"","familia":"","formacion":"","empresaAct":"","puestoAct":"","antig":"","remAct":"","remPret":"","disp":"","motivo":"","movilidad":"","referencias":""}}';
}

function extraerJSON_(txt) {
  const t = String(txt || '');
  const i = t.indexOf('{'), j = t.lastIndexOf('}');
  if (i < 0 || j <= i) throw new Error('La IA no devolvió un resultado legible. Probá de nuevo.');
  return JSON.parse(t.slice(i, j + 1));
}

function analizarEntrevista_(pin, perfilId, cand, compsJson, heroJson, texto) {
  pinValido_(pin);
  texto = String(texto || '').trim();
  if (texto.length < 300) throw new Error('La transcripción es muy corta para analizarla');
  let comps = [], hero = {};
  try { comps = JSON.parse(compsJson); hero = JSON.parse(heroJson) || {}; } catch (e) { throw new Error('Datos del formulario inválidos'); }
  const cortar = (v, n) => String(v == null ? '' : v).trim().slice(0, n);
  comps = (Array.isArray(comps) ? comps : []).slice(0, 15).map(c => ({ id: cortar(c.id, 30), nombre: cortar(c.nombre, 120), descripcion: cortar(c.descripcion, 500) })).filter(c => c.id);
  if (!comps.length) throw new Error('El formulario no tiene competencias para evaluar');
  const cache = CacheService.getScriptCache(), kA = iaCacheKey_('an_', perfilId, cand);
  try { cache.put(kA, JSON.stringify({ texto: texto.slice(0, 40000), resultado: null }), 21600); } catch (e) { }

  const r = extraerJSON_(procesarConClaude(promptAnalisisEntrevista_(hero, cand, comps), texto.slice(0, 60000), 'analisis'));
  const out = { comps: [], fortalezas: cortar(r.fortalezas, 2000), areas: cortar(r.areas, 2000), recomendacion: cortar(r.recomendacion, 2000), meta: {} };
  (r.comps || []).forEach(c => {
    if (comps.some(x => x.id === String(c.id))) out.comps.push({ id: String(c.id), nivel: Math.max(0, Math.min(4, Math.round(Number(c.nivel) || 0))), evidencia: cortar(c.evidencia, 1800) });
  });
  ['edad', 'localidad', 'familia', 'formacion', 'empresaAct', 'puestoAct', 'antig', 'remAct', 'remPret', 'disp', 'motivo', 'movilidad', 'referencias'].forEach(f => { out.meta[f] = cortar((r.meta || {})[f], 200); });
  try { cache.put(kA, JSON.stringify({ texto: texto.slice(0, 40000), resultado: out }), 21600); } catch (e) { }
  return JSON.stringify(out);
}

// Reescribe notas sueltas como texto bien redactado (botón «Redactar con IA» de los cuadros de notas). No resume ni inventa.
function redactarIA_(pin, texto, indic) {
  pinValido_(pin);
  texto = String(texto || '').trim();
  if (texto.length < 8) throw new Error('Escribí o dictá algo primero');
  const sistema = 'Te paso las notas sueltas de quien entrevista en un proceso de selección de personal. Redactalas como un texto bien escrito, completo y fluido, en oraciones y párrafos (no en viñetas ni frases sueltas), en español rioplatense profesional; corregí ortografía y puntuación. NO resumas ni acortes: conservá todos los datos, ejemplos, nombres, cifras y matices y conectalos para que se lea como una descripción de lo que contó la persona. Atribuí lo que contó a su relato («contó que», «mencionó que»). No inventes nada ni agregues adjetivos elogiosos. Devolvé solo el texto, sin comillas ni explicaciones.';
  const usuario = texto.slice(0, 4000) + (String(indic || '').trim() ? '\n\nINDICACIONES (aplicalas sin inventar datos): ' + String(indic).trim().slice(0, 300) : '');
  return procesarConClaude(sistema, usuario, 'redactar').trim();
}

/* Lectura de un CV: devuelve el JSON que armó la IA. Usa Claude (Sonnet 5.5) en las búsquedas que no son de MACS y ChatGPT mini en las de MACS;
   se puede forzar por perfil con la columna opcional `cv_ia` de Config (claude / openai). */
function usarClaudeCV_(cfg) {
  const x = String((cfg && cfg.cv_ia) || '').toLowerCase().trim();
  if (x === 'claude') return true;
  if (x === 'openai' || x === 'chatgpt') return false;
  return empresaKey_(cfg) !== 'MACS';
}

function analizarJSONconIA_(cfg, promptSistema, contenidoUsuario) {
  if (usarClaudeCV_(cfg)) return extraerJSON_(procesarConClaude(promptSistema, contenidoUsuario, 'cv'));
  const response = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post',
    headers: { Authorization: 'Bearer ' + OPENAI_API_KEY },
    contentType: 'application/json',
    payload: JSON.stringify({
      model: 'gpt-4o-mini',
      response_format: { type: 'json_object' },
      messages: [{ role: 'system', content: promptSistema }, { role: 'user', content: contenidoUsuario }]
    }),
    muteHttpExceptions: true
  });
  const parsed = JSON.parse(response.getContentText());
  if (parsed.error) throw new Error('OpenAI: ' + parsed.error.message);
  return JSON.parse(parsed.choices[0].message.content);
}
