/* ============================================================
   AVISOS A QUIEN ENTREVISTA (4/11, Ariel: «lo que no le llega es a la entrevistadora»).
   Cuando un candidato reserva solo en la página de citas del calendario de Blumia, Google invita al candidato pero no a quien lo va a entrevistar.
   Este archivo cierra ese hueco:
   1) agendaInvitar_: detecta la reserva (evento con el mail de un candidato y sin «Entrevista a cargo de»), invita a quien tiene asignado en
      «Próximo entrevistador» (le llega el mail de Google con día y hora) y anota la línea en la descripción. Si no hay nadie asignado, o la persona
      no tiene mail en Usuarios, le avisa a quien administra el script (una sola vez por evento).
   2) agendaRecordar_: unos 30 minutos antes de cada entrevista le manda un mail a quien entrevista.
   Se dispara con el calendario (cada vez que cambia) y cada 10 minutos de respaldo. No guarda nada en propiedades: lo que ya se hizo queda
   marcado en el propio evento (setTag).
   ============================================================ */

const AGENDA_CAL_ = 'blumiaconsulting@gmail.com';

function agendaVigilar() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return;
  try {
    try { agendaInvitar_(); } catch (e) { Logger.log('agendaInvitar_: ' + e); }
    try { agendaRecordar_(); } catch (e) { Logger.log('agendaRecordar_: ' + e); }
  } finally { lock.releaseLock(); }
}

// Crea los disparadores una sola vez: cambios en el calendario de Blumia + respaldo cada 10 minutos (se llama desde procesarTodo_v2)
function asegurarDisparadoresAgenda_() {
  try {
    const cache = CacheService.getScriptCache();
    if (cache.get('agenda_disp_ok')) return;
    const ts = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'agendaVigilar');
    if (!ts.some(t => t.getEventType() === ScriptApp.EventType.CLOCK)) ScriptApp.newTrigger('agendaVigilar').timeBased().everyMinutes(10).create();
    if (!ts.some(t => t.getEventType() === ScriptApp.EventType.CALENDAR)) {
      try { ScriptApp.newTrigger('agendaVigilar').forUserCalendar(AGENDA_CAL_).onEventUpdated().create(); } catch (e) { Logger.log('Disparador de calendario: ' + e); }
    }
    cache.put('agenda_disp_ok', '1', 21600);
  } catch (e) { Logger.log('asegurarDisparadoresAgenda_: ' + e); }
}

const agMailOk_ = m => /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(String(m || '').trim());
const agHora_ = d => Utilities.formatDate(d, 'GMT-3', "dd/MM/yyyy 'a las' HH:mm");
const agPrimerNombre_ = n => String(n || '').trim().split(' ')[0].toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
const agAdmin_ = () => { try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; } };

// quién entrevista según el evento: la línea «Entrevista a cargo de: X» de la descripción
function agConDe_(e) { const m = String(e.getDescription() || '').match(/Entrevista a cargo de:\s*([^\n\r<]+)/); return m ? m[1].trim() : ''; }

// mapa mail del candidato → datos de su fila, de las búsquedas abiertas
function agCandidatosPorMail_() {
  const mapa = {};
  perfilesActivos_().forEach(p => {
    try {
      const id = p.id || p;
      if (id === 'operaciones') return;
      const sh = getSheet(getConfig(id));
      const n = sh.getLastRow(); if (n < 2) return;
      const d = sh.getRange(1, 1, n, Math.max(sh.getLastColumn(), 22)).getValues();
      const cp = d[0].indexOf('Próximo entrevistador');
      for (let i = 1; i < n; i++) {
        const mail = String(d[i][21] || '').toLowerCase().trim();
        if (!d[i][1] || mail.indexOf('@') < 0) continue;
        mapa[mail] = { perfil: id, hoja: sh.getName(), nombre: String(d[i][1]), fila: i + 1, prox: cp >= 0 ? String(d[i][cp] || '').trim() : '', estado: String(d[i][8] || '') };
      }
    } catch (e) { Logger.log('agCandidatosPorMail_: ' + e); }
  });
  return mapa;
}

function agendaInvitar_() {
  const cal = CalendarApp.getCalendarById(AGENDA_CAL_);
  if (!cal) return;
  const ahora = new Date(), hasta = new Date(ahora.getTime() + 90 * 864e5);
  const eventos = cal.getEvents(ahora, hasta).filter(e => !e.getTag('agd_ok') && !agConDe_(e));
  if (!eventos.length) return;
  const cand = agCandidatosPorMail_(), usuarios = leerUsuarios_().filter(u => u.activo === 'SI');
  const adm = agAdmin_();
  eventos.forEach(e => {
    const inv = e.getGuestList().map(g => String(g.getEmail()).toLowerCase());
    const mailC = inv.find(m => cand[m]);
    if (!mailC) return;                       // no es la reserva de un candidato
    const c = cand[mailC];
    const cuando = agHora_(e.getStartTime());
    const u = c.prox ? usuarios.find(x => agPrimerNombre_(x.nombre) === agPrimerNombre_(c.prox)) : null;
    const aviso = (asunto, cuerpo) => { if (adm) { try { MailApp.sendEmail({ to: adm, subject: asunto, body: cuerpo, name: 'Blumia Consulting' }); } catch (x) { Logger.log('aviso: ' + x); } } };
    if (!u) {
      if (e.getTag('agd_sin')) return;
      aviso('⚠ Blumia: ' + c.nombre + ' reservó una entrevista y no tiene entrevistador asignado',
        c.nombre + ' reservó el ' + cuando + ' (' + e.getTitle() + ').\n\nNo hay nadie asignado como entrevistador, así que nadie recibió el aviso. Asignalo en la ficha del candidato (Próxima entrevista) y se invita solo en los próximos minutos.');
      e.setTag('agd_sin', '1');
      try { registrarHistorial_(c.hoja, c.nombre, 'Agenda', 'Reservó el ' + cuando + ' pero no hay entrevistador asignado: asignalo para que le llegue el aviso', 'sistema'); } catch (x) { }
      return;
    }
    if (!agMailOk_(u.mail)) {
      if (e.getTag('agd_sinmail')) return;
      aviso('⚠ Blumia: falta el mail de ' + u.nombre + ' para avisarle de la entrevista de ' + c.nombre,
        c.nombre + ' reservó el ' + cuando + ' y la entrevista es de ' + u.nombre + ', pero ' + u.nombre + ' no tiene mail cargado en Usuarios (Inicio → ⚙ Administración → Usuarios). Cargalo y se la invita solo en los próximos minutos.');
      e.setTag('agd_sinmail', '1');
      return;
    }
    // invitar: Google le manda el mail con día y hora
    try { e.addGuest(u.mail); } catch (x) { Logger.log('addGuest: ' + x); return; }
    try { e.setDescription((e.getDescription() ? e.getDescription() + '\n\n' : '') + 'Entrevista a cargo de: ' + u.nombre); } catch (x) { }
    e.setTag('agd_ok', '1');
    try { registrarHistorial_(c.hoja, c.nombre, 'Agenda', 'El candidato reservó el ' + cuando + ': se invitó a ' + u.nombre + ' a la entrevista', 'sistema'); } catch (x) { }
  });
}

// unos 30 minutos antes, mail a quien entrevista (entrevistas del calendario de Blumia con «Entrevista a cargo de»)
function agendaRecordar_() {
  const cal = CalendarApp.getCalendarById(AGENDA_CAL_);
  if (!cal) return;
  const ahora = Date.now();
  const eventos = cal.getEvents(new Date(ahora + 20 * 6e4), new Date(ahora + 40 * 6e4)).filter(e => !e.getTag('agd_rec') && agConDe_(e));
  if (!eventos.length) return;
  const usuarios = leerUsuarios_().filter(u => u.activo === 'SI');
  eventos.forEach(e => {
    if (e.isAllDayEvent()) return;
    const u = usuarios.find(x => agPrimerNombre_(x.nombre) === agPrimerNombre_(agConDe_(e)));
    if (!u || !agMailOk_(u.mail)) return;
    const inicio = e.getStartTime(), faltan = Math.max(1, Math.round((inicio.getTime() - ahora) / 6e4));
    const hora = Utilities.formatDate(inicio, 'GMT-3', 'HH:mm');
    const lugar = e.getLocation() ? '\nLugar / link: ' + e.getLocation() : '';
    try {
      MailApp.sendEmail({
        to: u.mail, name: 'Blumia Consulting',
        subject: '⏰ Entrevista en ' + faltan + ' minutos (' + hora + '): ' + e.getTitle(),
        body: 'Hola ' + String(u.nombre).split(' ')[0] + ', en ' + faltan + ' minutos (a las ' + hora + ') tenés esta entrevista:\n\n' + e.getTitle() + lugar + '\n\nEl evento está en el calendario de Blumia.'
      });
      e.setTag('agd_rec', '1');
    } catch (x) { Logger.log('recordatorio: ' + x); }
  });
}
