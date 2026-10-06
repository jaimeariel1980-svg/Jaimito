function admin_(k) {
  if (String(k) !== prop_('ADMIN_KEY')) throw new Error('No autorizado');
}

function fmt_(d, f) { return Utilities.formatDate(d, 'GMT-3', f); }

function filaPorId_(sh, id) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  const ids = sh.getRange(2, COL.id, n, 1).getValues();
  const i = ids.findIndex(r => String(r[0]) === String(id));
  return i === -1 ? 0 : i + 2;
}

function plantillasIniciales_() {
  return [
    ['contacto-screening', 'WhatsApp', 'Hola [NOMBRE], ¿cómo estás? Soy [ENTREVISTADOR] del equipo de Selección.\nTe contacto por la búsqueda de [PUESTO] en [EMPRESA].\nMe gustaría coordinar una charla breve con vos para conocerte. ¿Qué día y horario te quedan cómodos? [LINK]\nSaludos!'],
    ['contacto-competencias', 'WhatsApp', 'Hola [NOMBRE], ¿cómo estás? Soy [ENTREVISTADOR] del equipo de Selección.\nTe escribo por la búsqueda de [PUESTO] en [EMPRESA].\nNos gustaría avanzar con tu perfil a la siguiente etapa del proceso. Se trata de una entrevista de aproximadamente una hora, en la que vamos a profundizar sobre tu experiencia y tu recorrido profesional.\nSi te interesa continuar, contanos qué día y horario te quedan mejor para agendarla. [LINK]\n¡Nos vemos pronto!\nSaludos,\n[FIRMA]'],
    ['contacto-pami-virtual', 'WhatsApp', 'Hola [NOMBRE], ¿cómo estás?\nQueríamos contarte que nos gustaría avanzar con tu perfil a la siguiente etapa del proceso para la posición de [PUESTO] en [EMPRESA].\nLa próxima instancia consiste en una entrevista virtual con referentes de [EMPRESA], de aproximadamente 30 minutos. Será un primer encuentro para que puedas presentar tu experiencia, conocer al equipo y despejar las dudas que tengas sobre la posición.\nSi te interesa continuar, contanos qué día y horario te quedan mejor para agendarla. [LINK]\nSaludos,\n[FIRMA]'],
    ['contacto-pami-presencial', 'WhatsApp', 'Hola [NOMBRE], ¿cómo estás?\nQueríamos contarte que nos gustaría avanzar con tu perfil a la siguiente etapa del proceso para la posición de [PUESTO] en [EMPRESA].\nLa próxima instancia consiste en una entrevista presencial con referentes de [EMPRESA], de aproximadamente una hora. Será una conversación para profundizar sobre tu experiencia, conocer los desafíos de la posición y responder las consultas de ambas partes.\nSi te interesa continuar, contanos qué día y horario te quedan mejor para agendarla. [LINK]\nSaludos,\n[FIRMA]'],
    ['confirmacion-cita', 'WhatsApp', 'Hola [NOMBRE], te confirmamos tu entrevista de [TIPO] para la búsqueda de [PUESTO]:\n📅 [FECHA] a las [HORA]\n📍 [LUGAR]\nSi no podés asistir, avisanos con anticipación para reprogramarla. ¡Gracias!\n[FIRMA]'],
    ['descarte-sin', 'Mail', 'Hola [NOMBRE],\nQueremos agradecerte sinceramente por tu interés en la posición de [PUESTO] y por el tiempo dedicado a postularte.\nLuego de revisar cuidadosamente los currículums recibidos, en esta oportunidad decidimos avanzar con candidatos cuyos perfiles se encuentran más alineados con los requerimientos específicos de la posición.\nValoramos mucho tu trayectoria profesional. Conservaremos tu perfil para futuras oportunidades que puedan resultar acordes a tu experiencia.\nTe deseamos muchos éxitos en tus próximos desafíos profesionales.\nSaludos cordiales,\n[FIRMA]'],
    ['descarte-post', 'Mail', 'Hola [NOMBRE],\nMuchas gracias por el tiempo compartido durante el proceso de selección para la posición de [PUESTO].\nQueremos agradecer especialmente tu apertura y predisposición durante la conversación. Luego de analizar toda la información recabada, hemos decidido avanzar con otros perfiles que presentan un ajuste más cercano a las necesidades actuales de la posición.\nValoramos mucho tu experiencia y el intercambio mantenido, y esperamos poder considerar tu perfil para futuras oportunidades.\nTe deseamos muchos éxitos en tus próximos proyectos.\nSaludos cordiales,\n[FIRMA]'],
    ['descarte-competencias', 'WhatsApp', 'Hola [NOMBRE],\nAntes que nada, queremos agradecerte el tiempo que compartiste con nosotros durante las entrevistas para la posición de [PUESTO].\nValoramos mucho tu predisposición y el interés demostrado a lo largo del proceso.\nLuego de evaluar integralmente los perfiles, en esta oportunidad hemos decidido continuar con candidatos que presentan un ajuste más cercano a los desafíos específicos de la posición.\nEsperamos poder considerar tu perfil en futuras oportunidades. Te deseamos muchos éxitos.\nSaludos cordiales,\n[FIRMA]'],
    ['descarte-pami', 'WhatsApp', 'Hola [NOMBRE],\nAntes que nada, muchas gracias por el tiempo, el compromiso y la buena predisposición que demostraste durante todo el proceso de selección para la posición de [PUESTO].\nLuego de la entrevista con [EMPRESA], decidimos avanzar con otro perfil, principalmente porque [MOTIVO TÉCNICO].\nNos resultó muy interesante conocerte y tu perfil nos dejó una muy buena impresión. Por ese motivo, con tu autorización, mantendremos tus antecedentes para tenerte en cuenta en futuras búsquedas.\nTe agradecemos nuevamente y te deseamos muchos éxitos.\nSaludos,\n[FIRMA]']
  ];
}

function colorNota_(n) {
  return { A: ['#1E8E5A', '#FFFFFF'], 'B+': ['#8BD3A8', '#0F3A24'], B: ['#F2B705', '#4A3A00'], C: ['#D9534F', '#FFFFFF'] }[n] || null;
}

function estadoPorNota_(etapa, nota) {
  const C = { cv: 'Descartado sin screening', screening: 'Descartado post screening', competencias: 'Descartado en competencias', pami1: 'Descartado en PAMI virtual', pami2: 'Descartado en PAMI presencial' };
  const B = { cv: 'Avanza a screening', screening: 'Avanza a competencias', competencias: 'Reserva competencias', pami1: 'Reserva PAMI virtual', pami2: 'Reserva PAMI presencial' };
  const A = { cv: 'Avanza a screening', screening: 'Avanza a competencias', competencias: 'Avanza a PAMI virtual', pami1: 'Avanza a PAMI presencial', pami2: 'Contratado' };
  return (nota === 'C' ? C : nota === 'B' ? B : A)[etapa];
}

function etapaDeEstado_(estado) {
  const e = String(estado || '').trim();
  for (let i = 0; i < ESTADOS_COL.length; i++) if (ESTADOS_COL[i][1].indexOf(e) > -1) return i === 5 ? -1 : i;
  return -1;
}

function setEstado_(sh, fila, estado, nota) {
  const c = sh.getRange(fila, COL.estado);
  const regla = c.getDataValidation();
  c.setDataValidation(null).setValue(estado);
  if (regla) c.setDataValidation(regla);
  const col = nota ? colorNota_(nota) : null;
  if (col) c.setBackground(col[0]).setFontColor(col[1]).setFontWeight('bold').setHorizontalAlignment('center');
  sh.getRange(fila, COL.mov).setValue(new Date());
}

function registrarHistorial_(id, nombre, tipo, detalle, quien) {
  try { hoja_('Historial').appendRow([new Date(), id, nombre, tipo, detalle, quien || '']); } catch (e) { }
}

function aplicarNota_(sh, fila, etapa, nota, quien) {
  let v = String(nota || '').trim().toUpperCase();
  const m = { ALTO: 'A', MEDIO: 'B', BAJO: 'C' };
  if (m[v]) v = m[v];
  if (['A', 'B+', 'B', 'C'].indexOf(v) === -1 || !NOTA_COL[etapa]) return null;
  const r = sh.getRange(fila, 1, 1, HEAD_CAND.length).getValues()[0];
  const celda = sh.getRange(fila, NOTA_COL[etapa]);
  celda.clearDataValidations().setValue(v);
  const col = colorNota_(v);
  celda.setBackground(col[0]).setFontColor(col[1]).setFontWeight('bold').setHorizontalAlignment('center');
  const actual = String(r[COL.estado - 1]).trim();
  const etq = { cv: 'CV', screening: 'Screening', competencias: 'Competencias', pami1: 'PAMI virtual', pami2: 'PAMI presencial' }[etapa];
  if (actual === 'Contratado') {
    registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Nota ' + etq, v + ' (ya estaba Contratado, no cambia el estado)', quien);
    return actual;
  }
  const estado = estadoPorNota_(etapa, v);
  setEstado_(sh, fila, estado, v);
  registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Nota ' + etq, v + ' → ' + estado, quien);
  return estado;
}

function alEditar(e) {
  try {
    if (!e || !e.range) return;
    const sh = e.range.getSheet();
    if (sh.getName() !== 'Candidatos') return;
    const fila = e.range.getRow(), col = e.range.getColumn();
    if (fila < 2 || e.range.getNumRows() > 1) return;
    const etapa = Object.keys(NOTA_COL).find(k => NOTA_COL[k] === col);
    if (etapa) { aplicarNota_(sh, fila, etapa, e.range.getValue(), 'planilla'); return; }
    if (col === COL.estado) {
      sh.getRange(fila, COL.mov).setValue(new Date());
      registrarHistorial_(sh.getRange(fila, COL.id).getValue(), sh.getRange(fila, COL.nombre).getValue(), 'Estado', 'Cambiado a mano a "' + e.range.getValue() + '"', 'planilla');
    }
  } catch (err) { registrarError_('alEditar', err); }
}

function citaObj_(txt) {
  const m = String(txt || '').match(/^(\d{4}-\d{2}-\d{2} \d{2}:\d{2})\|(.*)$/);
  if (!m) return null;
  return { ts: m[1], tipo: m[2] };
}

function getTablero(k) {
  admin_(k);
  const aj = ajustes_();
  const d = hoja_('Candidatos').getDataRange().getValues();
  const cands = [];
  for (let i = 1; i < d.length; i++) {
    const r = d[i];
    if (!r[COL.nombre - 1]) continue;
    let resp = [];
    try { resp = JSON.parse(r[COL.resp - 1] || '[]'); } catch (e) { }
    const dt = v => v instanceof Date ? fmt_(v, 'dd/MM/yyyy') : String(v || '');
    cands.push({
      fila: i + 1, id: String(r[COL.id - 1]), fecha: dt(r[0]), mov: r[COL.mov - 1] instanceof Date ? fmt_(r[COL.mov - 1], "yyyy-MM-dd'T'HH:mm") : '',
      nombre: r[COL.nombre - 1], tel: String(r[COL.tel - 1] || ''), dni: String(r[COL.dni - 1] || ''), email: r[COL.email - 1], loc: r[COL.loc - 1],
      busq: r[COL.busq - 1], pret: String(r[COL.pret - 1] || ''), ajuste: r[COL.ajuste - 1], just: r[COL.just - 1], resumen: r[COL.resumen - 1], req: r[COL.req - 1],
      cv: r[COL.cv - 1], estado: String(r[COL.estado - 1] || 'Nuevo').trim(), entrev: r[COL.entrev - 1],
      notaScr: r[COL.notaScr - 1], notaComp: r[COL.notaComp - 1], notaP1: r[COL.notaP1 - 1], notaP2: r[COL.notaP2 - 1],
      recTipo: r[COL.recTipo - 1], recPor: r[COL.recPor - 1], coment: r[COL.coment - 1], resp: resp, proc: r[COL.proc - 1], fuente: r[COL.fuente - 1],
      cita: citaObj_(r[COL.cita - 1]), entrevistas: String(r[COL.entrevistas - 1] || ''), contactado: String(r[COL.contactado - 1] || ''),
      descarte: String(r[COL.descarte - 1] || ''), infScr: String(r[COL.infScr - 1] || ''), infComp: String(r[COL.infComp - 1] || '')
    });
  }
  return JSON.stringify({
    estadosCol: ESTADOS_COL, desc: ESTADOS_DESC, res: ESTADOS_RES, todos: todosEstados_(),
    entrevistadores: lista_(aj.entrevistadores, ','),
    busquedas: busquedas_().map(b => ({ id: b.id, titulo: b.titulo, recomendacion: b.recomendacion, link_perfil: b.link_perfil, ubicacion: b.ubicacion })),
    cfg: { cliente: aj.nombre_cliente || 'PAMI', diasAlerta: Number(aj.dias_alerta) || 7, linkAgenda: aj.link_agenda || '', duraciones: { screening: Number(aj.duracion_screening) || 30, competencias: Number(aj.duracion_competencias) || 60, pami1: Number(aj.duracion_pami1) || 30, pami2: Number(aj.duracion_pami2) || 60 } },
    ahora: fmt_(new Date(), "yyyy-MM-dd'T'HH:mm"), cands: cands
  });
}

function getHistorial(k, id) {
  admin_(k);
  const d = hoja_('Historial').getDataRange().getValues().slice(1).filter(r => String(r[1]) === String(id));
  return d.reverse().slice(0, 80).map(r => ({ fecha: r[0] instanceof Date ? fmt_(r[0], 'dd/MM/yyyy HH:mm') : String(r[0]), tipo: r[3], detalle: r[4], quien: r[5] }));
}

function moverEstado(k, fila, estado) {
  admin_(k);
  if (todosEstados_().indexOf(estado) === -1) throw new Error('Estado inválido');
  const sh = hoja_('Candidatos');
  fila = Number(fila);
  if (!(fila >= 2 && fila <= sh.getLastRow())) throw new Error('Fila inválida');
  const antes = String(sh.getRange(fila, COL.estado).getValue());
  if (antes === estado) return true;
  setEstado_(sh, fila, estado, null);
  registrarHistorial_(sh.getRange(fila, COL.id).getValue(), sh.getRange(fila, COL.nombre).getValue(), 'Estado', antes + ' → ' + estado, 'tablero');
  return true;
}

function ponerNota(k, fila, etapa, nota) {
  admin_(k);
  const sh = hoja_('Candidatos');
  fila = Number(fila);
  if (!(fila >= 2 && fila <= sh.getLastRow())) throw new Error('Fila inválida');
  if (!NOTA_COL[etapa]) throw new Error('Etapa inválida');
  return aplicarNota_(sh, fila, etapa, nota, 'tablero');
}

function guardarCandidato(k, fila, c) {
  admin_(k);
  const sh = hoja_('Candidatos');
  fila = Number(fila);
  if (!(fila >= 2 && fila <= sh.getLastRow())) throw new Error('Fila inválida');
  const corto = (s, n) => String(s == null ? '' : s).slice(0, n);
  const set = (col, v) => sh.getRange(fila, col).clearDataValidations().setValue(v);
  const id = sh.getRange(fila, COL.id).getValue(), nombre = sh.getRange(fila, COL.nombre).getValue();
  if (c.busq != null) {
    if (c.busq !== '' && !busquedas_().some(b => b.id === c.busq)) throw new Error('Búsqueda inválida');
    set(COL.busq, c.busq);
  }
  if (c.estado != null) moverEstado(k, fila, c.estado);
  if (c.entrev != null) set(COL.entrev, corto(c.entrev, 60));
  if (c.recTipo != null) set(COL.recTipo, corto(c.recTipo, 40));
  if (c.recPor != null) {
    set(COL.recPor, corto(c.recPor, 120));
    if (c.recTipo) registrarHistorial_(id, nombre, 'Recomendación', c.recTipo + (c.recPor ? ': ' + c.recPor : ''), 'tablero');
  }
  if (c.coment != null) set(COL.coment, corto(c.coment, 2000));
  return true;
}

function reanalizar(k, fila) {
  admin_(k);
  const sh = hoja_('Candidatos');
  sh.getRange(Number(fila), COL.intentos).setValue(0);
  analizarFila_(Number(fila));
  return sh.getRange(Number(fila), COL.proc).getValue();
}

function rellenar_(txt, v) {
  return String(txt).replace(/\[([A-ZÁÉÍÓÚÑ ]+)\]/g, (m, k) => (v[k] != null ? v[k] : m));
}

function textoPlantilla_(clave) {
  const f = hoja_('Plantillas').getDataRange().getValues().find(r => r[0] === clave);
  return f ? String(f[2]) : '';
}

function datosComunes_(r) {
  const aj = ajustes_();
  const b = busquedas_().find(x => x.id === r[COL.busq - 1]);
  return {
    NOMBRE: String(r[COL.nombre - 1]).trim().split(/\s+/)[0], PUESTO: b ? b.titulo : '', EMPRESA: aj.nombre_cliente || 'PAMI',
    ENTREVISTADOR: String(r[COL.entrev - 1] || lista_(aj.entrevistadores, ',')[0] || ''), LINK: aj.link_agenda || '', FIRMA: aj.firma || '', aj: aj
  };
}

function telWa_(t) {
  let d = String(t || '').replace(/\D/g, '');
  if (d.indexOf('54') === 0) d = d.slice(2);
  d = d.replace(/^0+/, '');
  if (d.charAt(0) === '9' && d.length >= 11) d = d.slice(1);
  return '549' + d;
}

function armarComunicacion(k, fila, tipo) {
  admin_(k);
  const sh = hoja_('Candidatos');
  fila = Number(fila);
  const r = sh.getRange(fila, 1, 1, HEAD_CAND.length).getValues()[0];
  const v = datosComunes_(r);
  let texto = textoPlantilla_(tipo);
  if (!texto) return { ok: false, error: 'No encontré la plantilla ' + tipo + ' en la pestaña Plantillas' };
  texto = rellenar_(texto, v);
  if (!String(r[COL.tel - 1]).replace(/\D/g, '')) return { ok: false, error: 'El candidato no tiene teléfono' };
  const etq = { 'contacto-screening': 'screening', 'contacto-competencias': 'competencias', 'contacto-pami-virtual': 'PAMI virtual', 'contacto-pami-presencial': 'PAMI presencial', 'descarte-competencias': 'descarte comp.', 'descarte-pami': 'descarte PAMI', 'confirmacion-cita': 'confirmación' }[tipo] || '';
  if (etq) {
    sh.getRange(fila, COL.contactado).setValue(etq + ' · ' + fmt_(new Date(), 'dd/MM/yyyy'));
    sh.getRange(fila, COL.mov).setValue(new Date());
    registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Contacto', 'WhatsApp: ' + etq, 'tablero');
  }
  return { ok: true, url: 'https://wa.me/' + telWa_(r[COL.tel - 1]) + '?text=' + encodeURIComponent(texto), texto: texto };
}

function borrarMarcaContacto(k, fila) {
  admin_(k);
  hoja_('Candidatos').getRange(Number(fila), COL.contactado).clearContent();
  return true;
}

function mailValido_(m) { return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(m || '').trim()); }

function mailDescarte_(sh, fila, textoSin, textoPost) {
  const r = sh.getRange(fila, 1, 1, HEAD_CAND.length).getValues()[0];
  const mail = String(r[COL.email - 1]).trim();
  if (!mailValido_(mail)) return { ok: false, error: 'Sin mail válido' };
  const v = datosComunes_(r);
  const sin = /sin screening/i.test(String(r[COL.estado - 1]));
  const base = sin ? (textoSin || textoPlantilla_('descarte-sin')) : (textoPost || textoPlantilla_('descarte-post'));
  if (!base) return { ok: false, error: 'No hay plantilla de descarte' };
  GmailApp.sendEmail(mail, v.aj.asunto_mail || 'Proceso de selección', rellenar_(base, v));
  sh.getRange(fila, COL.descarte).setValue('Enviado ' + fmt_(new Date(), 'dd/MM/yyyy'));
  registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Mail', 'Descarte enviado por mail', 'tablero');
  return { ok: true };
}

function enviarMailDescarte(k, fila) {
  admin_(k);
  const sh = hoja_('Candidatos');
  try { return mailDescarte_(sh, Number(fila)); } catch (e) { return { ok: false, error: String(e) }; }
}

function previaDescartes(k, busq) {
  admin_(k);
  const d = hoja_('Candidatos').getDataRange().getValues();
  let sin = 0, post = 0, sinMail = 0;
  for (let i = 1; i < d.length; i++) {
    if (!d[i][COL.nombre - 1] || (busq && d[i][COL.busq - 1] !== busq)) continue;
    const e = String(d[i][COL.estado - 1]).trim();
    if (!/^Descartado (sin screening|post screening)$/i.test(e) || String(d[i][COL.descarte - 1]).trim()) continue;
    if (!mailValido_(d[i][COL.email - 1])) { sinMail++; continue; }
    if (/sin screening/i.test(e)) sin++; else post++;
  }
  return { sin: sin, post: post, sinMail: sinMail, textoSin: textoPlantilla_('descarte-sin'), textoPost: textoPlantilla_('descarte-post') };
}

function enviarDescartesMasivo(k, busq, textoSin, textoPost) {
  admin_(k);
  const sh = hoja_('Candidatos');
  const d = sh.getDataRange().getValues();
  let enviados = 0, sinMail = 0, pendientes = 0;
  for (let i = 1; i < d.length; i++) {
    if (!d[i][COL.nombre - 1] || (busq && d[i][COL.busq - 1] !== busq)) continue;
    if (!/^Descartado (sin screening|post screening)$/i.test(String(d[i][COL.estado - 1]).trim()) || String(d[i][COL.descarte - 1]).trim()) continue;
    if (!mailValido_(d[i][COL.email - 1])) { sinMail++; continue; }
    if (enviados >= 40) { pendientes++; continue; }
    try { const res = mailDescarte_(sh, i + 1, textoSin, textoPost); if (res.ok) enviados++; } catch (e) { registrarError_('mail descarte fila ' + (i + 1), e); }
  }
  return { enviados: enviados, sinMail: sinMail, pendientes: pendientes };
}

function getAgenda(k) {
  admin_(k);
  const d = hoja_('Candidatos').getDataRange().getValues();
  const desde = fmt_(new Date(Date.now() - 3600000), 'yyyy-MM-dd HH:mm');
  const out = [];
  for (let i = 1; i < d.length; i++) {
    const c = citaObj_(d[i][COL.cita - 1]);
    if (!c || c.ts < desde || !d[i][COL.nombre - 1]) continue;
    out.push({ fila: i + 1, nombre: d[i][COL.nombre - 1], ts: c.ts, tipo: c.tipo, busq: d[i][COL.busq - 1], tel: String(d[i][COL.tel - 1] || '') });
  }
  out.sort((a, b) => a.ts < b.ts ? -1 : 1);
  return out;
}
