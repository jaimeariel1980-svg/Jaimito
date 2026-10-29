/* Prueba del circuito completo de Blumia (se corre desde el editor de Apps Script).
   - blumiaPruebaCompleta: crea una búsqueda «PRUEBA AUTOMÁTICA», pasa un candidato falso por todo el circuito (postulación,
     análisis del CV con IA, estados, agenda, screening y competencias con PDF, contactos, observaciones, descarte, papelera,
     flyer, informe y páginas) y al final borra todo lo que creó. El resultado queda en la pestaña `Pruebas_Blumia`.
   - Solo corre si la propiedad del script BLUMIA_PRUEBAS = SI (así nadie la dispara por error ni desde el link).
   - No manda mails a candidatos: el candidato falso usa el mail de quien corre la prueba (la invitación de la cita le llega a esa persona).
   - blumiaLimpiarPruebas: borra restos si una corrida anterior se cortó a la mitad. */

const PRUEBA_PREFIJO_ = 'prueba-automatica';
const PRUEBA_JPG_ = '/9j/4AAQSkZJRgABAQEASABIAAD/2wBDAP//////////////////////////////////////////////////////////////////////////////////////wgALCAABAAEBAREA/8QAFBABAAAAAAAAAAAAAAAAAAAAAP/aAAgBAQABPxA=';

function blumiaPruebaCompleta() {
  if (PropertiesService.getScriptProperties().getProperty('BLUMIA_PRUEBAS') !== 'SI') throw new Error('Para correr la prueba cargá la propiedad del script BLUMIA_PRUEBAS = SI (Configuración del proyecto → Propiedades del script).');
  const sesAntes = SES_;
  SES_ = { nombre: 'Prueba automática', rol: 'Administrador', fila: 0 };
  try { return pruebaCompleta_(); } finally { SES_ = sesAntes; }
}

function blumiaLimpiarPruebas() {
  if (PropertiesService.getScriptProperties().getProperty('BLUMIA_PRUEBAS') !== 'SI') throw new Error('Cargá BLUMIA_PRUEBAS = SI para limpiar.');
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const d = ss.getSheetByName('Config').getDataRange().getValues();
  const ids = d.slice(1).map(r => String(r[0] || '')).filter(id => id.indexOf(PRUEBA_PREFIJO_) === 0);
  const log = [];
  ids.forEach(id => limpiarPrueba_({ id: id }, log));
  Logger.log(log.join('\n') || 'No había restos de pruebas.');
  return log.length ? log : ['No había restos de pruebas.'];
}

function pruebaCompleta_(op) {
  const res = [], c = { quien: 'Prueba automática' };
  const email = (op && op.sinMail) ? '' : Session.getEffectiveUser().getEmail(); // la prueba semanal no invita ni manda mails
  const requiere = (cond, msg) => { if (!cond) throw new Error(msg); };
  const paso = (nombre, fn) => {
    const t = Date.now();
    try { const r = fn(); res.push([nombre, 'OK', String(r === undefined ? '' : r).slice(0, 300), Date.now() - t]); return true; }
    catch (e) { res.push([nombre, 'FALLÓ', String(e && e.message || e).slice(0, 300), Date.now() - t]); return false; }
  };
  const celda = (col) => String(getSheet(c.cfg).getRange(c.fila, col).getValue() || '').trim();
  const hoyMas = n => Utilities.formatDate(new Date(Date.now() + n * 86400000), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd');

  try {
    // 1. claves
    paso('Claves de IA (OpenAI y Anthropic)', () => { const r = probarClaves(); requiere(!/ERROR|FALTA/.test(r), r); return 'OK'; });

    // 2. crear la búsqueda
    const stamp = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyyMMdd-HHmm');
    const prop = {
      nombre: 'PRUEBA AUTOMÁTICA ' + stamp, empresa: 'Prueba Blumia', ubicacion: 'Rosario', banda_salarial: '1000000',
      prompt_extra: 'PERFIL: Director Comercial de una empresa de servicios. Reporta a la gerencia general. EXPERIENCIA (EXCLUYENTE): 8 años o más liderando equipos comerciales. FORMACIÓN: Licenciatura en Administración, Marketing o afín. ZONA: Rosario y alrededores. NIVEL_AJUSTE: Alto si cumple todo, Medio si le falta un requisito no excluyente, Bajo si no cumple la experiencia.',
      screening_eval: [{ criterio: 'Experiencia en el rubro', peso: 3 }, { criterio: 'Comunicación', peso: 2 }],
      screening_val: [{ pregunta: 'Disponibilidad inmediata', opciones: 'Sí|No|A revisar' }],
      competencias: [{ nombre: 'Liderazgo', descripcion: 'Conduce equipos', peso: 3, preguntas: ['Contá una situación', 'Qué hiciste', 'Qué resultado tuvo'] }, { nombre: 'Comunicación', descripcion: 'Se expresa con claridad', peso: 2, preguntas: ['Contá una situación', 'Qué hiciste', 'Qué resultado tuvo'] }],
      postulacion: [{ etiqueta: '¿Por qué te interesa la búsqueda?', tipo: 'texto', opciones: '', requerido: 'SI' }], links: {}
    };
    if (!paso('Crear la búsqueda de prueba', () => {
      const r = crearBusquedaIA_('', prop, c.quien, '');
      c.id = r.id; c.nombre = r.nombre;
      requiere(c.id.indexOf(PRUEBA_PREFIJO_) === 0, 'El id de la búsqueda no es de prueba: ' + c.id);
      return c.id;
    })) throw new Error('corte');

    paso('La búsqueda quedó en Config y con carpetas', () => {
      c.cfg = getConfig(c.id);
      ['folder_madre', 'folder_recibidos', 'folder_pdf_screening', 'folder_pdf_competencias', 'hoja_nombre'].forEach(k => requiere(c.cfg[k], 'Falta ' + k + ' en Config'));
      DriveApp.getFolderById(c.cfg.folder_pdf_screening);
      requiere(SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(c.cfg.hoja_nombre), 'No se creó la pestaña de candidatos');
      requiere(listarPerfiles().some(p => p.id === c.id), 'No aparece en la lista de búsquedas');
      requiere(getPreguntasPostulacion(c.id).length === 1, 'Las preguntas de postulación no se cargaron');
      return 'OK';
    });
    requiere(c.cfg, 'Sin Config no se puede seguir');

    // 3. páginas
    [['Página: tablero', { tipo: 'kanban', perfil: c.id }], ['Página: postulación', { tipo: 'postulacion', perfil: c.id }],
     ['Página: screening', { perfil: c.id, nombre: 'Prueba Candidato', entrevistador: 'Prueba' }], ['Página: competencias', { tipo: 'competencias', perfil: c.id, nombre: 'Prueba Candidato', entrevistador: 'Prueba' }],
     ['Página: flyer', { tipo: 'flyer', perfil: c.id }], ['Página: informe de la búsqueda', { tipo: 'informebusqueda', perfil: c.id }]].forEach(p => paso(p[0], () => {
      const h = doGetPagina_({ parameter: p[1] }).getContent();
      requiere(h.length > 3000, 'La página salió casi vacía (' + h.length + ' caracteres)');
      requiere(!/incompleto|Error al cargar/i.test(h.slice(0, 4000)), 'La página muestra un error');
      return Math.round(h.length / 1024) + ' KB';
    }));

    paso('Informe de la búsqueda: números y armado (sin IA)', () => {
      const e = ibEstadisticas_(c.id);
      requiere(typeof e.total === 'number' && typeof e.scr === 'number', 'No salieron los números');
      const h = ibArmarHTML_(getConfig(c.id), 'Prueba', 'Prueba', e, { titulo: 'T', resumen: 'R', hallazgos: [{ titulo: 'a', texto: 'b' }], recomendaciones: ['x'], proximos: ['y'], cierre: 'z' });
      requiere(h.length > 3000 && h.indexOf('__FILEID__') > 0, 'El informe salió incompleto');
      return e.total + ' candidatos';
    });

    paso('Vaciar búsqueda: pide el nombre y no borra con otro', () => {
      const i = vaciarBusquedaInfo_(c.id);
      let frenó = false;
      try { vaciarBusqueda_(c.id, 'nombre equivocado ' + Date.now(), 'Prueba'); } catch (e) { frenó = /No borré nada/.test(String(e && e.message)); }
      requiere(frenó, 'Vaciar la búsqueda no frenó con un nombre equivocado');
      return i.candidatos + ' candidatos (no se borró nada)';
    });

    paso('Agenda: quién entrevista según la etapa', () => {
      const e = entrevistadoresAgenda_();
      requiere(Array.isArray(e.lista) && e.lista.length > 0, 'No hay usuarios para elegir entrevistador');
      requiere(usuHace_({ entrevista: '' }, 'competencias') && !usuHace_({ entrevista: 'screening' }, 'competencias') && !usuHace_({ entrevista: 'ninguna' }, 'screening'), 'La regla de etapas por persona no anda');
      return e.lista.length + ' usuarios';
    });

    paso('Quién entrevistó (tarjeta): se lee de Evaluaciones', () => {
      const m = entrevistadoresEval_(c.id);
      requiere(m && typeof m === 'object', 'No devolvió un mapa');
      return Object.keys(m).length + ' evaluaciones con entrevistador';
    });

    paso('Remitente de los mails de Blumia', () => {
      const r = remitenteBlumia_();
      requiere(/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(r), 'El remitente no es una dirección válida: ' + r);
      let alias = false; try { alias = GmailApp.getAliases().some(a => String(a).toLowerCase() === r.toLowerCase()); } catch (e) { }
      return r + (alias ? ' (alias dado de alta en Gmail)' : ' (SIN dar de alta como remitente en Gmail: los mails salen de la cuenta del sistema)');
    });

    // 3b. búsqueda ciega
    paso('Búsqueda ciega: sin nombre ni logo de la empresa para el candidato', () => {
      try {
        setConfigCampo_(c.id, 'busqueda_ciega', 'SI');
        const cfgC = getConfig(c.id), h = getHeroPostulacion(c.id), fd = flyerDatos_(c.id);
        requiere(!h.logo && !h.empresa && String(h.loc).indexOf('Prueba Blumia') < 0, 'La postulación todavía muestra la empresa: «' + h.loc + '»');
        requiere(empresaParaCandidatos_(cfgC).indexOf('Prueba Blumia') < 0, 'Los mensajes todavía nombran a la empresa');
        requiere(fd.ciega && String(fd.empresa).indexOf('Prueba Blumia') < 0, 'El flyer todavía nombra a la empresa');
      } finally { setConfigCampo_(c.id, 'busqueda_ciega', 'NO'); }
      return 'OK';
    });

    // 3c. preguntas habituales propias
    paso('Preguntas habituales propias (lista general)', () => {
      const antes = leerHabituales_();
      try {
        const r = guardarHabituales_('', antes.concat([{ etiqueta: 'PRUEBA ¿Tenés registro de conducir?', tipo: 'select', opciones: 'Sí|No', requerido: 'SI' }]), c.quien);
        requiere(r.habituales.some(h => /PRUEBA ¿Tenés registro/.test(h.etiqueta)), 'No quedó guardada en la lista');
        requiere(getConfigBusqueda_('', c.id).habituales.length === antes.length + 1, 'Configurar no devuelve la lista');
      } finally { guardarHabituales_('', antes, c.quien); }
      requiere(leerHabituales_().length === antes.length, 'La lista no volvió a como estaba');
      return 'OK';
    });

    // 4. postulación
    const dni = '9' + String(Date.now()).slice(-7);
    const cand = { nombre: 'Prueba Candidato Uno', telefono: '3415550000', dni: dni, email: email, nivelFormacion: 'Universitario', estadoFormacion: 'Completo', tituloFormacion: 'Lic. en Administración', remuneracion: '1000000', localidad: 'Rosario', movilidad: 'Sí' };
    paso('Postulación con CV (recibirPostulacion)', () => {
      const html = '<html><body><h1>Prueba Candidato Uno</h1><p>Licenciado en Administración con 12 años de experiencia como Director Comercial en una empresa de servicios de Rosario. Lideró un equipo de 25 personas y llevó las ventas de 100 a 300 millones. Reporta a la gerencia general.</p></body></html>';
      const pdf = Utilities.newBlob(html, 'text/html', 'cv.html').getAs('application/pdf');
      c.cvB64 = Utilities.base64Encode(pdf.getBytes());
      const r = recibirPostulacion({ perfil: c.id, datos: cand, declaradas: {}, cvBase64: c.cvB64, cvNombre: 'cv-prueba.pdf', cvTipo: 'application/pdf' });
      requiere(r === 'OK', 'Respuesta inesperada: ' + r);
      const d = getSheet(c.cfg).getDataRange().getValues();
      for (let i = 1; i < d.length; i++) if (String(d[i][1]).trim() === cand.nombre) { c.fila = i + 1; break; }
      requiere(c.fila, 'El candidato no apareció en la pestaña');
      requiere(celda(9) === 'Pendiente', 'Estado inicial: «' + celda(9) + '»');
      requiere(/^https:\/\/drive\.google\.com/.test(celda(8)), 'No hay link al CV');
      return 'fila ' + c.fila;
    });
    paso('Postulación repetida se detecta como duplicado', () => {
      const r = recibirPostulacion({ perfil: c.id, datos: cand, declaradas: {}, cvBase64: c.cvB64, cvNombre: 'cv-prueba.pdf', cvTipo: 'application/pdf' });
      requiere(r === 'DUPLICADO', 'Devolvió «' + r + '»');
      return 'DUPLICADO';
    });
    requiere(c.fila, 'Sin candidato no se puede seguir');

    // 5. análisis del CV con IA
    paso('Lectura del CV y análisis con IA', () => {
      const meta = JSON.parse(getSheet(c.cfg).getRange(c.fila, 26).getValue());
      const texto = textoDeCV(meta.cvId);
      const g = analizarPostulacionGPT(texto, 'Nombre: ' + cand.nombre + '\nLocalidad: Rosario\nRemuneración pretendida: 1000000', c.cfg);
      requiere(g && g.ok, 'La IA falló: ' + (g && g.error));
      const nota = String(g.nivel_ajuste || '');
      const sh = getSheet(c.cfg);
      sh.getRange(c.fila, 5).setValue(nota || 'A'); sh.getRange(c.fila, 9).setValue('Nuevo');
      return 'nota ' + (nota || '(sin dato)') + ' · CV de ' + texto.length + ' caracteres';
    });

    // 6. tablero
    paso('El tablero trae al candidato (doGetKanban_)', () => {
      const lista = JSON.parse(doGetKanban_(c.id));
      const x = lista.filter(k => k.nombre === cand.nombre)[0];
      requiere(x, 'No aparece en los datos del tablero');
      requiere(x.cv && x.estado === 'Nuevo', 'Datos incompletos: estado «' + x.estado + '»');
      return 'OK';
    });
    paso('Cambios de estado (moverEstado_)', () => {
      moverEstado_(c.id, c.fila, 'Contactado', c.quien); requiere(celda(9) === 'Contactado', 'No pasó a Contactado');
      moverEstado_(c.id, c.fila, 'Avanza a screening', c.quien); requiere(celda(9) === 'Avanza a screening', 'No pasó a Avanza a screening');
      return 'OK';
    });
    paso('Contacto por WhatsApp y registro', () => {
      const r = armarComunicacion_(c.id, c.fila, 'contacto-screening', c.quien);
      requiere(r, 'No armó el mensaje');
      registrarContactoManual_(c.id, c.fila, 'Llamé', c.quien);
      requiere(getContactosRegistro_(c.id, c.fila).length >= 1, 'No quedó en el registro de contactos');
      return 'OK';
    });
    paso('Observaciones', () => { guardarObservaciones_(c.id, c.fila, 'Observación de prueba', c.quien); return 'OK'; });

    // 7. agenda
    paso('Agendar screening en el calendario', () => {
      const r = agendarCitaTablero_(c.id, c.fila, 'screening', hoyMas(3), '10:30', 30, 'Meet de prueba', true, c.quien);
      requiere(r && r.ok, 'No se pudo agendar');
      c.evento = r.evento;
      requiere(celda(9) === 'Screening agendado', 'Estado: «' + celda(9) + '»');
      requiere(celda(37), 'No quedó la cita en la ficha');
      return r.cita;
    });

    // 8. screening y competencias con PDF
    paso('Screening: nota, PDF y respuestas guardadas', () => {
      const html = '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body><h1>Screening de prueba</h1><p>Texto de prueba con acentos: áéíóú ñ.</p></body></html>';
      const r = procesarScreening_({ perfil: c.id, nombre: cand.nombre, resultado: 'A', entrevistador: 'Prueba', htmlFull: html, estado: JSON.stringify({ prueba: 1 }) });
      requiere(r === 'OK', 'Respuesta: ' + r);
      requiere(celda(12) === 'A', 'La nota de screening no quedó');
      requiere(/HYPERLINK/.test(String(getSheet(c.cfg).getRange(c.fila, 13).getFormula())), 'Falta el link al PDF');
      return 'estado «' + celda(9) + '»';
    });
    paso('Competencias: nota, PDF y respuestas guardadas', () => {
      const html = '<!DOCTYPE html><html><head><meta charset="utf-8"></head><body><h1>Competencias de prueba</h1></body></html>';
      const r = guardarEnPlanilla_(cand.nombre, html, 'A', c.id, c.quien, JSON.stringify({ prueba: 1 }), false);
      requiere(r === 'OK', 'Respuesta: ' + r);
      requiere(celda(16) === 'A', 'La nota de competencias no quedó');
      requiere(/HYPERLINK/.test(String(getSheet(c.cfg).getRange(c.fila, 17).getFormula())), 'Falta el link al PDF');
      return 'estado «' + celda(9) + '»';
    });
    paso('Se puede reabrir la evaluación guardada', () => { requiere(getEvaluacionPrevia_(c.id, cand.nombre, 'screening'), 'No hay respuestas guardadas'); return 'OK'; });
    paso('Informe del candidato', () => {
      const url = generarInformeCandidato_(c.id, c.fila);
      requiere(/verinforme&id=/.test(url), 'Link inesperado');
      c.informeId = (url.match(/id=([\w-]+)/) || [])[1];
      return 'OK';
    });

    // 9. flyer
    paso('Flyer: diseño, imagen y envío', () => {
      flyerGuardar_(c.id, JSON.stringify({ fmt: 'historia', txt: { badge: 'Prueba' } }));
      const im = flyerImagen_(c.id, PRUEBA_JPG_);
      const s = flyerSocial_(c.id);
      requiere(s.imagen && s.imagenId, 'No quedó la imagen del flyer');
      requiere(im.imagen, 'No devolvió el link de la imagen');
      requiere(flyerHTML_(c.id).getContent().length > 3000, 'La pantalla del flyer salió vacía');
      return 'OK';
    });

    // 10. descarte y papelera
    paso('Descarte con motivo y vista previa de mails (no envía)', () => {
      moverEstado_(c.id, c.fila, 'Descartado post screening', c.quien);
      guardarMotivoDescarte_(c.id, c.fila, [MOTIVOS_DESCARTE_[0]], '', c.quien);
      const p = previaDescartes_(c.id);
      requiere(p, 'La vista previa de descartes no respondió');
      return 'OK';
    });
    paso('Papelera: eliminar y restaurar', () => {
      eliminarCandidato_(c.id, c.fila, c.quien); requiere(/^Eliminado\|/.test(celda(9)), 'No pasó a la papelera');
      restaurarCandidato_(c.id, c.fila, c.quien); requiere(!/^Eliminado\|/.test(celda(9)), 'No se restauró');
      return 'OK';
    });
    paso('Historial del candidato', () => { requiere(getHistorialCandidato_(c.id, c.fila).length >= 5, 'El historial tiene muy pocos movimientos'); return 'OK'; });
  } catch (e) {
    if (String(e && e.message) !== 'corte') res.push(['Corte de la prueba', 'FALLÓ', String(e && e.message || e).slice(0, 300), 0]);
  }

  // limpieza (siempre)
  const log = [];
  if (c.id) {
    paso('Limpieza: se borra todo lo que creó la prueba', () => {
      limpiarPrueba_(c, log);
      return log.filter(l => /^No pude/.test(l)).length ? log.join(' | ') : 'OK (' + log.length + ' cosas borradas)';
    });
    if (log.some(l => /^No pude/.test(l))) res[res.length - 1][1] = 'FALLÓ';
  }
  return pruebaInforme_(res);
}

function pruebaInforme_(res) {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let sh = ss.getSheetByName('Pruebas_Blumia');
  if (!sh) { sh = ss.insertSheet('Pruebas_Blumia'); sh.appendRow(['Fecha', 'Paso', 'Resultado', 'Detalle', 'ms']); sh.getRange(1, 1, 1, 5).setFontWeight('bold'); }
  const fecha = new Date(), mal = res.filter(r => r[1] !== 'OK').length;
  const filas = res.map(r => [fecha, r[0], r[1], r[2], r[3]]);
  filas.push([fecha, 'RESUMEN', mal ? 'FALLÓ' : 'OK', res.length - mal + ' pasos bien · ' + mal + ' con problemas', res.reduce((s, r) => s + r[3], 0)]);
  sh.getRange(sh.getLastRow() + 1, 1, filas.length, 5).setValues(filas);
  sh.getRange(sh.getLastRow() - filas.length + 1, 3, filas.length, 1).setFontColor('#000');
  const out = res.map(r => (r[1] === 'OK' ? '✅ ' : '❌ ') + r[0] + (r[1] === 'OK' ? '' : ' → ' + r[2])).join('\n') + '\n\n' + (mal ? '❌ ' + mal + ' paso(s) con problemas' : '✅ Todo el circuito funciona');
  Logger.log(out);
  return out;
}

/* Borra todo lo de UNA búsqueda de prueba. Solo actúa si el id empieza con «prueba-automatica». */
function limpiarPrueba_(c, log, forzar) {
  const id = String(c.id || '');
  if (!forzar && id.indexOf(PRUEBA_PREFIJO_) !== 0) throw new Error('Por seguridad solo se borran búsquedas de prueba.');
  const hacer = (que, fn) => { try { const n = fn(); log.push(que + (n === undefined ? '' : ' (' + n + ')')); } catch (e) { log.push('No pude ' + que + ': ' + String(e && e.message || e).slice(0, 120)); } };
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let cfg = null;
  try { cfg = getConfig(id); } catch (e) { }
  const nombreBusq = cfg ? String(cfg.nombre || cfg.nombre_perfil || '') : '';
  if (!forzar && cfg && nombreBusq.indexOf('PRUEBA AUTOM') !== 0) throw new Error('Por seguridad solo se borran búsquedas de prueba.');

  hacer('borrar el evento del calendario', () => { if (c.evento) { const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com'); const ev = cal && cal.getEventById(c.evento); if (ev) ev.deleteEvent(); } });
  hacer('borrar el informe del candidato', () => { if (c.informeId) DriveApp.getFileById(c.informeId).setTrashed(true); });
  if (forzar) hacer('mandar a la papelera de Drive las 4 carpetas de la búsqueda (se recuperan 30 días)', () => {
    let n = 0;
    ['folder_madre', 'folder_recibidos', 'folder_pdf_screening', 'folder_pdf_competencias'].forEach(k => { if (cfg && cfg[k]) { const f = DriveApp.getFolderById(cfg[k]); let a = 0; const it = f.getFiles(); while (it.hasNext()) { it.next(); a++; } log.push('carpeta «' + f.getName() + '» con ' + a + ' archivo(s)'); f.setTrashed(true); n++; } });
    return n;
  });
  if (!forzar) hacer('borrar las carpetas de Drive', () => {
    if (!cfg || !cfg.folder_madre) return 0;
    const carpeta = DriveApp.getFolderById(cfg.folder_madre).getParents().next();
    const empresa = carpeta.getParents().hasNext() ? carpeta.getParents().next() : null;
    carpeta.setTrashed(true);
    if (empresa && empresa.getName() === 'Prueba Blumia') { const it = empresa.getFolders(); let hay = false; while (it.hasNext()) { const f = it.next(); if (f.getId() !== carpeta.getId() && !f.isTrashed()) hay = true; } if (!hay) empresa.setTrashed(true); }
    return 1;
  });
  hacer('borrar la pestaña de candidatos', () => { const s = cfg && ss.getSheetByName(cfg.hoja_nombre); if (s) ss.deleteSheet(s); });
  ['Config', 'Perfiles_Hero', 'Perfiles_Screening_Eval', 'Perfiles_Screening_Val', 'Perfiles_Competencias', 'Postulacion_Preguntas', 'Mail_Routing'].forEach(nom => hacer('borrar filas de ' + nom, () => {
    const s = ss.getSheetByName(nom); if (!s) return 0;
    const d = s.getDataRange().getValues(); let n = 0;
    for (let i = d.length - 1; i >= 1; i--) if (String(d[i][0]) === id) { s.deleteRow(i + 1); n++; }
    return n;
  }));
  [['Historial', 1], ['Evaluaciones', 1], ['Registro_Contactos', 1]].forEach(p => hacer('borrar filas de ' + p[0], () => {
    const s = ss.getSheetByName(p[0]); if (!s || s.getLastRow() < 2) return 0;
    const d = s.getRange(1, 1, s.getLastRow(), Math.min(s.getLastColumn(), 8)).getValues(); let n = 0;
    for (let i = d.length - 1; i >= 1 && n < 60; i--) if (d[i].some(v => { const t = String(v); return t === id || (nombreBusq && t === nombreBusq) || (cfg && cfg.hoja_nombre && t === cfg.hoja_nombre); })) { s.deleteRow(i + 1); n++; }
    return n;
  }));
  hacer('limpiar la memoria rápida', () => { const k = CacheService.getScriptCache(); k.remove('cfg_all'); });
}

/* Prueba semanal automática: los lunes de madrugada corre el mismo circuito y, si algo falla, avisa por mail a quien es dueño del script.
   El activador se crea solo desde procesarTodo_v2 (asegurarPruebaSemanal_). Para apagarla: propiedad del script BLUMIA_PRUEBA_SEMANAL = NO. */
function blumiaPruebaSemanal() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('BLUMIA_PRUEBA_SEMANAL') === 'NO') return;
  const cache = CacheService.getScriptCache();
  if (cache.get('prueba_semanal_corriendo')) return;
  cache.put('prueba_semanal_corriendo', '1', 600);
  const para = Session.getEffectiveUser().getEmail();
  const sesAntes = SES_;
  SES_ = { nombre: 'Prueba automática', rol: 'Administrador', fila: 0 };
  try {
    const out = pruebaCompleta_({ sinMail: true });
    if (/❌/.test(out)) MailApp.sendEmail({ to: para, subject: '⚠️ Blumia: la prueba semanal encontró problemas', body: out + '\n\nEl detalle completo está en la pestaña Pruebas_Blumia de la planilla.' });
  } catch (e) {
    try { MailApp.sendEmail({ to: para, subject: '⚠️ Blumia: la prueba semanal no pudo correr', body: String(e && e.message || e) + '\n\nSi quedaron restos, corré blumiaLimpiarPruebas desde el editor (con BLUMIA_PRUEBAS = SI).' }); } catch (e2) { }
  } finally { SES_ = sesAntes; cache.remove('prueba_semanal_corriendo'); }
}

function asegurarPruebaSemanal_() {
  try {
    const cache = CacheService.getScriptCache();
    if (cache.get('prueba_semanal_ok')) return;
    if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'blumiaPruebaSemanal')) ScriptApp.newTrigger('blumiaPruebaSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(4).create();
    cache.put('prueba_semanal_ok', '1', 21600);
  } catch (e) { Logger.log('asegurarPruebaSemanal_: ' + e); }
}

/* Borra una búsqueda vieja de prueba («Gerente de Logística», pedido de Ariel 29/10). Se corre desde el editor con BLUMIA_PRUEBAS = SI.
   Protecciones: tiene que haber exactamente UNA búsqueda cuyo nombre diga «logística» y su pestaña no puede tener candidatos.
   Manda a la papelera de Drive las 4 carpetas de la búsqueda (la carpeta madre de la búsqueda no se toca) y borra la pestaña y sus filas de configuración. */
function blumiaBorrarBusquedaLogistica() {
  if (PropertiesService.getScriptProperties().getProperty('BLUMIA_PRUEBAS') !== 'SI') throw new Error('Cargá BLUMIA_PRUEBAS = SI para borrar.');
  const norm = t => String(t || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const d = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config').getDataRange().getValues(), H = d[0];
  const col = k => H.indexOf(k);
  const hit = d.slice(1).filter(r => r[0] && ['nombre', 'nombre_perfil', 'hoja_nombre'].some(k => col(k) >= 0 && /gerente.*logistic|logistic/.test(norm(r[col(k)]))));
  if (hit.length !== 1) throw new Error('Esperaba encontrar UNA búsqueda de logística y encontré ' + hit.length + ': ' + hit.map(r => r[0]).join(', ') + '. No borré nada.');
  const id = String(hit[0][0]), cfg = getConfig(id), sh = getSheet(cfg);
  const n = sh ? sh.getDataRange().getValues().slice(1).filter(r => String(r[1] || '').trim()).length : 0;
  if (n > 0) throw new Error('La búsqueda «' + id + '» tiene ' + n + ' candidato(s) cargado(s). Por seguridad no borré nada.');
  const log = [];
  limpiarPrueba_({ id: id }, log, true);
  const out = 'Búsqueda «' + id + '» borrada:\n' + log.join('\n');
  Logger.log(out);
  return out;
}
