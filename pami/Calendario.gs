function calendario_() {
  const id = ajustes_().calendario;
  return id ? CalendarApp.getCalendarById(id) : CalendarApp.getDefaultCalendar();
}

function etapaDeTitulo_(titulo) {
  const t = norm_(titulo);
  if (t.indexOf('competencia') > -1) return 'competencias';
  if (t.indexOf('screening') > -1) return 'screening';
  if (t.indexOf('virtual') > -1) return 'pami1';
  if (t.indexOf('presencial') > -1 || t.indexOf('pami') > -1) return 'pami2';
  return '';
}

function syncCalendario() {
  const cal = calendario_();
  if (!cal) return;
  const ahora = new Date();
  const desde = new Date(ahora.getTime() - 180 * 86400000), hasta = new Date(ahora.getTime() + 90 * 86400000);
  const eventos = cal.getEvents(desde, hasta).map(e => ({
    titulo: e.getTitle(), inicio: e.getStartTime(),
    invitados: e.getGuestList().map(g => g.getEmail().toLowerCase())
  }));
  const sh = hoja_('Candidatos');
  const d = sh.getDataRange().getValues();
  const nn = s => norm_(s).replace(/[^a-z0-9]/g, '');
  for (let i = 1; i < d.length; i++) {
    const r = d[i];
    if (!r[COL.nombre - 1]) continue;
    const nombreN = nn(r[COL.nombre - 1]);
    const mail = String(r[COL.email - 1] || '').trim().toLowerCase();
    const suyos = eventos.filter(e => {
      if (mail && e.invitados.indexOf(mail) > -1) return true;
      const m = e.titulo.match(/\(([^)]+)\)/);
      if (!m) return false;
      const evN = nn(m[1]);
      return evN && (evN === nombreN || (evN.length >= 10 && (nombreN.indexOf(evN) > -1 || evN.indexOf(nombreN) > -1)));
    }).filter(e => etapaDeTitulo_(e.titulo));
    const fila = i + 1;

    const hist = suyos.slice().sort((a, b) => b.inicio - a.inicio)
      .map(e => TIPO_ETAPA[etapaDeTitulo_(e.titulo)] + ' — ' + fmt_(e.inicio, 'dd/MM/yyyy HH:mm')).join('\n');
    if (hist !== String(r[COL.entrevistas - 1] || '')) sh.getRange(fila, COL.entrevistas).setValue(hist);

    const futuros = suyos.filter(e => e.inicio >= ahora).sort((a, b) => a.inicio - b.inicio);
    const prox = futuros[0];
    const citaNueva = prox ? fmt_(prox.inicio, 'yyyy-MM-dd HH:mm') + '|' + TIPO_ETAPA[etapaDeTitulo_(prox.titulo)] : '';
    if (citaNueva !== String(r[COL.cita - 1] || '')) sh.getRange(fila, COL.cita).setValue(citaNueva);

    if (prox) {
      const etapa = etapaDeTitulo_(prox.titulo);
      const estado = String(r[COL.estado - 1]).trim();
      const actual = etapaDeEstado_(estado);
      if (actual > -1 && actual <= ETAPAS.indexOf(etapa) && estado !== AGENDADO[etapa]) {
        setEstado_(sh, fila, AGENDADO[etapa], null);
        registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Agenda', TIPO_ETAPA[etapa] + ' agendado: ' + fmt_(prox.inicio, 'dd/MM/yyyy HH:mm'), 'calendario');
      }
    }
  }
}

function agendarCita(k, fila, o) {
  admin_(k);
  const sh = hoja_('Candidatos');
  fila = Number(fila);
  if (!(fila >= 2 && fila <= sh.getLastRow())) throw new Error('Fila inválida');
  if (!AGENDADO[o.etapa]) throw new Error('Tipo de entrevista inválido');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(o.fecha) || !/^\d{2}:\d{2}$/.test(o.hora)) throw new Error('Fecha u hora inválida');
  const dur = Math.max(10, Math.min(240, Number(o.duracion) || 30));
  const inicio = new Date(o.fecha + 'T' + o.hora + ':00-03:00');
  if (isNaN(inicio.getTime())) throw new Error('Fecha u hora inválida');
  const fin = new Date(inicio.getTime() + dur * 60000);
  const r = sh.getRange(fila, 1, 1, HEAD_CAND.length).getValues()[0];
  const cal = calendario_();
  if (!o.forzar) {
    const choque = cal.getEvents(inicio, fin);
    if (choque.length) return { ok: false, conflicto: choque.map(e => e.getTitle()).join(', ') };
  }
  const v = datosComunes_(r);
  const titulo = TIPO_ETAPA[o.etapa] + ' (' + String(r[COL.nombre - 1]).trim() + ')';
  const mail = String(r[COL.email - 1]).trim();
  const opts = { description: 'Búsqueda: ' + v.PUESTO + '\nCandidato/a: ' + r[COL.nombre - 1] + (r[COL.tel - 1] ? ' · ' + r[COL.tel - 1] : '') + (o.nota ? '\n' + String(o.nota).slice(0, 500) : ''), location: String(o.lugar || '').slice(0, 200) };
  if (mailValido_(mail) && o.invitar !== false) { opts.guests = mail; opts.sendInvites = true; }
  cal.createEvent(titulo, inicio, fin, opts);

  const estado = String(r[COL.estado - 1]).trim();
  const actual = etapaDeEstado_(estado);
  if (actual > -1 && actual <= ETAPAS.indexOf(o.etapa)) setEstado_(sh, fila, AGENDADO[o.etapa], null);
  sh.getRange(fila, COL.cita).setValue(fmt_(inicio, 'yyyy-MM-dd HH:mm') + '|' + TIPO_ETAPA[o.etapa]);
  sh.getRange(fila, COL.mov).setValue(new Date());
  registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Agenda', TIPO_ETAPA[o.etapa] + ' agendado: ' + fmt_(inicio, 'dd/MM/yyyy HH:mm') + (o.lugar ? ' · ' + o.lugar : ''), 'tablero');

  let url = '';
  if (r[COL.tel - 1]) {
    const texto = rellenar_(textoPlantilla_('confirmacion-cita'), Object.assign({}, v, { TIPO: TIPO_ETAPA[o.etapa], FECHA: fmt_(inicio, 'dd/MM/yyyy'), HORA: fmt_(inicio, 'HH:mm'), LUGAR: o.lugar || 'A confirmar' }));
    url = 'https://wa.me/' + telWa_(r[COL.tel - 1]) + '?text=' + encodeURIComponent(texto);
  }
  return { ok: true, url: url, ts: fmt_(inicio, 'yyyy-MM-dd HH:mm') };
}
