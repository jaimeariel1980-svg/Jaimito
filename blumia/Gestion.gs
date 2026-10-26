/* Gestión del candidato desde la ficha: observaciones (edición de datos, anular etapa y agendar entrevista). */
function colPorNombre_(sh, nombre, crear) {
  const h = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  let c = h.indexOf(nombre) + 1;
  if (!c && crear) {
    c = Math.max(sh.getLastColumn(), 46) + 1;
    if (sh.getMaxColumns() < c) sh.insertColumnsAfter(sh.getMaxColumns(), c - sh.getMaxColumns());
    sh.getRange(1, c).setValue(nombre).setFontWeight('bold');
  }
  return c;
}

function guardarObservaciones(perfilId, fila, texto, quien) {
  QUIEN_ = quien || '';
  const sh = getSheet(getConfig(perfilId));
  fila = filaCandidato_(sh, fila);
  const t = String(texto || '').replace(/[<>]/g, '').trim().slice(0, 4000);
  const c = colPorNombre_(sh, 'Observaciones', true);
  const antes = String(sh.getRange(fila, c).getValue() || '');
  if (antes === t) return { ok: true, obs: t };
  sh.getRange(fila, c).setNumberFormat('@').setValue(t);
  registrarHistorial_(sh.getName(), sh.getRange(fila, 2).getValue(), 'Observaciones', antes ? 'Se editaron las observaciones' : 'Se cargaron observaciones');
  return { ok: true, obs: t };
}

// copia de seguridad a demanda (el botón de la ficha / menú pide el PIN)
function copiaAhoraBlumia(pin) {
  pinValido_(pin);
  return { ok: true, fecha: hacerCopiaBlumia_() };
}

function miles_(n) { return String(Number(n)).replace(/\B(?=(\d{3})+(?!\d))/g, '.'); }

/* ---- Editar los datos del candidato (F1) ---- */
function guardarDatosCandidato(perfilId, fila, d, quien) {
  QUIEN_ = quien || '';
  const sh = getSheet(getConfig(perfilId));
  fila = filaCandidato_(sh, fila);
  const limpio = v => String(v == null ? '' : v).replace(/[<>]/g, '').trim();
  const nombre = limpio(d.nombre);
  if (!nombre) throw new Error('El nombre no puede quedar vacío.');
  const mail = limpio(d.mail);
  if (mail && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(mail)) throw new Error('El mail no parece válido.');
  const tel = limpio(d.tel);
  if (tel && tel.replace(/\D/g, '').length < 8) throw new Error('El teléfono tiene que tener al menos 8 números.');
  const campos = [[2, 'Nombre', nombre], [3, 'Teléfono', tel], [4, 'DNI', limpio(d.dni)], [22, 'Mail', mail], [23, 'Formación', limpio(d.formacion)], [24, 'Zona', limpio(d.zona)]];
  const cambios = [];
  campos.forEach(c => {
    const antes = String(sh.getRange(fila, c[0]).getValue() || '');
    if (antes !== c[2]) { sh.getRange(fila, c[0]).setNumberFormat('@').setValue(c[2]); cambios.push(c[1] + ': ' + (antes || '—') + ' → ' + (c[2] || '—')); }
  });
  const pret = String(limpio(d.pretension)).replace(/\D/g, '');
  if (pret) {
    const req = String(sh.getRange(fila, 6).getValue() || '');
    const re = /(Remuneraci[oó]n pretendida:\s*\$?\s*)([\d.\s]+)/i;
    const nuevo = re.test(req) ? req.replace(re, '$1' + miles_(pret)) : (req ? req + '\n' : '') + 'Remuneración pretendida: $' + miles_(pret);
    if (nuevo !== req) { sh.getRange(fila, 6).setValue(nuevo); cambios.push('Pretensión: $' + miles_(pret)); }
  }
  if (cambios.length) registrarHistorial_(sh.getName(), nombre, 'Datos editados', cambios.join(' · '));
  return { ok: true, cambios: cambios.length };
}

/* ---- Anular una etapa (F4): saca la nota, deja el PDF y el estado vuelve a «Avanza a …» ---- */
function anularEtapa(perfilId, fila, etapa, quien) {
  QUIEN_ = quien || '';
  const sh = getSheet(getConfig(perfilId));
  fila = filaCandidato_(sh, fila);
  const E = {
    'screening': { col: 12, nom: 'screening', vuelve: 'Avanza a screening', desde: ['Avanza a competencias', 'Descartado post screening', 'Competencias agendada', 'En reserva'] },
    'competencias': { col: 16, nom: 'competencias', vuelve: 'Avanza a competencias', desde: ['Avanza a screening empresa', 'Descartado en competencias', 'Reserva competencias', 'Screening empresa agendado'] },
    'screening-empresa': { col: 44, nom: 'screening con la empresa', vuelve: 'Avanza a screening empresa', desde: ['Avanza a presencial empresa', 'Reserva screening empresa', 'Descartado en screening empresa', 'Entrevista empresa agendada'] },
    'presencial-empresa': { col: 45, nom: 'entrevista final con la empresa', vuelve: 'Avanza a presencial empresa', desde: ['Contratado', 'Reserva presencial empresa', 'Descartado presencial empresa'] }
  }[etapa];
  if (!E) throw new Error('Etapa inválida.');
  const celda = sh.getRange(fila, E.col);
  const nota = String(celda.getValue() || '').trim();
  if (!nota) throw new Error('Esa etapa no tiene nota para anular.');
  const estado = String(sh.getRange(fila, 9).getValue() || '').trim();
  celda.clearContent().setBackground(null);
  let nuevoEstado = estado;
  if (E.desde.indexOf(estado) > -1) {
    const c9 = sh.getRange(fila, 9), r = c9.getDataValidation();
    c9.setDataValidation(null).setValue(E.vuelve);
    if (r) c9.setDataValidation(r);
    nuevoEstado = E.vuelve;
    limpiarMotivoSiNoDescartado_(sh, fila, nuevoEstado);
  }
  registrarHistorial_(sh.getName(), sh.getRange(fila, 2).getValue(), 'Etapa anulada', E.nom + ' (nota ' + nota + ')' + (nuevoEstado !== estado ? ' · «' + estado + '» → «' + nuevoEstado + '»' : '') + ' · el PDF sigue en Drive');
  return { ok: true, estado: nuevoEstado };
}

/* ---- Agendar una entrevista desde el tablero (C3): crea el evento en el Calendar de Blumia y deja la cita en la ficha ---- */
const AGENDA_TIPOS_ = {
  'screening': { titulo: 'Screening Blumia', estado: 'Screening agendado' },
  'competencias': { titulo: 'Competencias Blumia', estado: 'Competencias agendada' },
  'empresa-virtual': { titulo: 'Entrevista con la empresa (virtual)', estado: 'Screening empresa agendado' },
  'empresa-presencial': { titulo: 'Entrevista con la empresa (presencial)', estado: 'Entrevista empresa agendada' }
};

function agendarCitaTablero(perfilId, fila, tipo, fecha, hora, duracion, lugar, forzar, quien) {
  QUIEN_ = quien || '';
  const T = AGENDA_TIPOS_[tipo];
  if (!T) throw new Error('Tipo de entrevista inválido.');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(fecha)) || !/^\d{2}:\d{2}$/.test(String(hora))) throw new Error('Fecha u hora inválida.');
  const inicio = new Date(fecha + 'T' + hora + ':00-03:00');
  if (isNaN(inicio.getTime())) throw new Error('Fecha u hora inválida.');
  if (inicio.getTime() < Date.now() - 3600000) throw new Error('La fecha ya pasó.');
  const dur = Math.max(10, Math.min(240, Number(duracion) || 30));
  const fin = new Date(inicio.getTime() + dur * 60000);
  const sh = getSheet(getConfig(perfilId));
  fila = filaCandidato_(sh, fila);
  const nombre = String(sh.getRange(fila, 2).getValue() || '').trim();
  const mail = String(sh.getRange(fila, 22).getValue() || '').trim();
  const cal = CalendarApp.getCalendarById('blumiaconsulting@gmail.com');
  if (!cal) throw new Error('No tengo acceso al calendario de Blumia.');
  const choques = cal.getEvents(inicio, fin);
  if (choques.length && !forzar) return { ok: false, conflicto: true, eventos: choques.slice(0, 3).map(e => e.getTitle()) };
  const opts = { description: 'Candidato/a: ' + nombre + (lugar ? '\n' + lugar : '') + '\n\nAgendado desde el tablero de Blumia' };
  if (lugar) opts.location = String(lugar).slice(0, 300);
  if (/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(mail)) { opts.guests = mail; opts.sendInvites = true; }
  const ev = cal.createEvent(T.titulo + ' (' + nombre + ')', inicio, fin, opts);
  // estado y cita en la ficha (el sincronizador de Calendar los confirma después)
  const c9 = sh.getRange(fila, 9), r = c9.getDataValidation();
  c9.setDataValidation(null).setValue(T.estado);
  if (r) c9.setDataValidation(r);
  if (sh.getMaxColumns() < 38) sh.insertColumnsAfter(sh.getMaxColumns(), 38 - sh.getMaxColumns());
  const txt = Utilities.formatDate(inicio, 'GMT-3', "dd/MM/yyyy 'a las' HH:mm");
  sh.getRange(fila, 37).setValue(txt);
  const tipoHist = tipo.indexOf('empresa') === 0 ? 'Empresa' : (tipo === 'competencias' ? 'Competencias' : 'Screening');
  const hist = String(sh.getRange(fila, 38).getValue() || '');
  sh.getRange(fila, 38).setValue(tipoHist + ' — ' + Utilities.formatDate(inicio, 'GMT-3', 'dd/MM/yyyy HH:mm') + (hist ? '\n' + hist : ''));
  registrarHistorial_(sh.getName(), nombre, 'Agenda', T.estado + ': ' + txt + (opts.guests ? ' · invitación enviada a ' + mail : ' · sin mail, no se envió invitación') + ' (creada desde el tablero)');
  return { ok: true, cita: txt, invitado: !!opts.guests, evento: ev.getId() };
}
