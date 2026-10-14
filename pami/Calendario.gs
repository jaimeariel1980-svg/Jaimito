function calendario_() {
  const id = ajustes_().calendario;
  return id ? CalendarApp.getCalendarById(id) : CalendarApp.getDefaultCalendar();
}

function calendarioPor_(id) {
  if (!id) return null;
  try { const c = CalendarApp.getCalendarById(id); if (c) return c; } catch (e) { }
  try { return CalendarApp.subscribeToCalendar(id); } catch (e) { return null; }
}

function estadoCalendarios_() {
  const aj = ajustes_(), avisos = [];
  const ahora = new Date(), desde = new Date(ahora.getTime() - 86400000), hasta = new Date(ahora.getTime() + 14 * 86400000);
  const lista = [{ n: 'PAMI 1', id: aj.calendario, cal: () => calendario_() }];
  if (aj.calendario_poli2) lista.push({ n: 'PAMI 2', id: aj.calendario_poli2, cal: () => calendarioPor_(aj.calendario_poli2) });
  lista.forEach(x => {
    let cal = null;
    try { cal = x.cal(); } catch (e) { }
    if (!cal) { avisos.push('Calendario ' + x.n + ': no puedo abrirlo (' + (x.id || 'sin configurar') + '). Compartilo con ' + CUENTA + ' con el permiso "Ver todos los detalles del evento" y aceptá la invitación desde esa cuenta.'); return; }
    let evs = [];
    try { evs = cal.getEvents(desde, hasta); } catch (e) { avisos.push('Calendario ' + x.n + ': no tengo permiso para leer sus eventos. Compartilo con ' + CUENTA + ' con "Ver todos los detalles del evento".'); return; }
    const porMail = estadoPorMail_();
    const titulos = evs.map(e => tituloEvento_(e, porMail));
    if (titulos.length && titulos.every(t => !t || /^(ocupado|busy)$/i.test(t))) { avisos.push('Calendario ' + x.n + ': veo ' + titulos.length + ' turnos pero sin título (está compartido solo como libre/ocupado). Cambiá el permiso a "Ver todos los detalles del evento".'); return; }
    const raros = titulos.filter(t => t && !/^(ocupado|busy)$/i.test(t) && !etapaDeTitulo_(t));
    if (raros.length) avisos.push('Calendario ' + x.n + ': ' + raros.length + ' turno' + (raros.length === 1 ? '' : 's') + ' no se lee' + (raros.length === 1 ? '' : 'n') + ' porque el título no dice Screening, Competencias o Técnica (ej.: "' + raros[0].slice(0, 50) + '"). Agregalo al título de esa agenda en Calendar.');
  });
  return avisos;
}

function calendarios_() {
  const l = [calendario_()];
  const c2 = calendarioPor_(ajustes_().calendario_poli2);
  if (c2) l.push(c2);
  return l;
}

function calendarioDe_(r) {
  const pol = String(r[COL.policlinico - 1] || String(r[COL.polSug - 1] || '').split(' · ')[0]);
  if (/2$/.test(pol)) { const c = calendarioPor_(ajustes_().calendario_poli2); if (c) return c; }
  return calendario_();
}

function etapaDeTitulo_(titulo) {
  const t = norm_(titulo);
  if (t.indexOf('competencia') > -1) return 'competencias';
  if (t.indexOf('screening') > -1) return 'screening';
  if (t.indexOf('tecnica') > -1 || t.indexOf('enfermeria') > -1 || t.indexOf('virtual') > -1 || t.indexOf('presencial') > -1 || t.indexOf('pami') > -1) return 'pami2';
  return '';
}

function tokens_(s) { return norm_(s).replace(/[^a-z0-9 ]/g, ' ').split(/\s+/).filter(t => t.length > 1); }

function mismoNombre_(a, b) {
  const A = tokens_(a), B = tokens_(b);
  if (A.length < 2 || B.length < 2) return false;
  const sub = (X, Y) => X.every(t => Y.indexOf(t) > -1);
  return sub(A, B) || sub(B, A);
}

function eventoEsDe_(ev, cand) {
  if (cand.mail && ev.invitados.indexOf(cand.mail) > -1) return true;
  const m = ev.titulo.match(/\(([^)]+)\)/);
  return !!m && mismoNombre_(m[1], cand.nombre);
}

function estadoPorMail_() {
  const m = {};
  hoja_('Candidatos').getDataRange().getValues().slice(1).forEach(r => { const k = String(r[COL.email - 1] || '').trim().toLowerCase(); if (k) m[k] = r[COL.estado - 1]; });
  return m;
}

function tituloEvento_(e, porMail) {
  const t = String(e.getTitle() || '').trim(), n = norm_(t);
  if (/competencia|screening|tecnica/.test(n)) return t;
  const g = e.getGuestList().map(x => x.getEmail().toLowerCase()).find(x => porMail[x] != null);
  const i = g ? etapaDeEstado_(porMail[g]) : -1;
  return i > 0 ? TIPO_ETAPA[ETAPAS[i]] + ' · ' + t : t;
}

function eventosCal_(desde, hasta) {
  const porMail = estadoPorMail_();
  let evs = [];
  calendarios_().forEach((cal, i) => { try { evs = evs.concat(cal.getEvents(desde, hasta).map(e => ({ e: e, poli: POLICLINICOS[i] || '' }))); } catch (e) { registrarError_('leer calendario', e); } });
  return evs.map(x => ({
    titulo: tituloEvento_(x.e, porMail), inicio: x.e.getStartTime(), invitados: x.e.getGuestList().map(g => g.getEmail().toLowerCase()), poli: x.poli
  })).filter(e => etapaDeTitulo_(e.titulo));
}

function syncCalendario() {
  if (!calendario_()) return;
  const ahora = new Date();
  const eventos = eventosCal_(new Date(ahora.getTime() - 180 * 86400000), new Date(ahora.getTime() + 90 * 86400000));
  const sh = hoja_('Candidatos');
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    const r = d[i];
    if (!r[COL.nombre - 1]) continue;
    const cand = { nombre: r[COL.nombre - 1], mail: String(r[COL.email - 1] || '').trim().toLowerCase() };
    const suyos = eventos.filter(e => eventoEsDe_(e, cand));
    const fila = i + 1;

    const hist = suyos.slice().sort((a, b) => b.inicio - a.inicio)
      .map(e => TIPO_ETAPA[etapaDeTitulo_(e.titulo)] + ' — ' + fmt_(e.inicio, 'dd/MM/yyyy HH:mm')).join('\n');
    if (hist !== String(r[COL.entrevistas - 1] || '')) sh.getRange(fila, COL.entrevistas).setValue(hist);

    const prox = suyos.filter(e => e.inicio >= ahora).sort((a, b) => a.inicio - b.inicio)[0];
    const citaNueva = prox ? fmt_(prox.inicio, 'yyyy-MM-dd HH:mm') + '|' + TIPO_ETAPA[etapaDeTitulo_(prox.titulo)] : '';
    if (citaNueva !== String(r[COL.cita - 1] || '')) sh.getRange(fila, COL.cita).setValue(citaNueva);

    if (prox) {
      const etapa = etapaDeTitulo_(prox.titulo);
      const estado = String(r[COL.estado - 1]).trim();
      const actual = etapaDeEstado_(estado);
      if (actual > -1 && actual <= ETAPAS.indexOf(etapa) && estado !== AGENDADO[etapa]) {
        setEstado_(sh, fila, AGENDADO[etapa], null);
        registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Agenda', TIPO_ETAPA[etapa] + ' agendado: ' + fmt_(prox.inicio, 'dd/MM/yyyy HH:mm') + (prox.titulo ? ' (' + prox.titulo + ')' : ''), 'calendario');
      }
    }
  }
}

function citasSinCandidato_(evs) {
  const ahora = new Date();
  const eventos = evs || eventosCal_(ahora, new Date(ahora.getTime() + 60 * 86400000));
  const d = hoja_('Candidatos').getDataRange().getValues().slice(1).filter(r => r[COL.nombre - 1]).map(r => ({ nombre: r[COL.nombre - 1], mail: String(r[COL.email - 1] || '').trim().toLowerCase() }));
  return eventos.filter(e => !d.some(c => eventoEsDe_(e, c)))
    .sort((a, b) => a.inicio - b.inicio)
    .map(e => ({ titulo: e.titulo, ts: fmt_(e.inicio, 'yyyy-MM-dd HH:mm'), tipo: TIPO_ETAPA[etapaDeTitulo_(e.titulo)], invitados: e.invitados.join(', '), poli: e.poli }));
}

function agendarCita(k, fila, o) {
  escritor_(k);
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
  const cal = calendarioDe_(r);
  if (!o.forzar) {
    const choque = cal.getEvents(inicio, fin);
    if (choque.length) return { ok: false, conflicto: choque.map(e => e.getTitle()).join(', ') };
  }
  const v = datosComunes_(r);
  const titulo = TIPO_ETAPA[o.etapa] + ' (' + String(r[COL.nombre - 1]).trim() + ')';
  const mail = String(r[COL.email - 1]).trim();
  const opts = { description: 'Búsqueda: ' + v.PUESTO + '\nCandidato/a: ' + r[COL.nombre - 1] + (r[COL.tel - 1] ? ' · ' + r[COL.tel - 1] : '') + (o.nota ? '\n' + String(o.nota).slice(0, 500) : ''), location: String(o.lugar || '').slice(0, 200) };
  if (mailValido_(mail) && o.invitar !== false) { opts.guests = mail; opts.sendInvites = true; }
  try { cal.createEvent(titulo, inicio, fin, opts); } catch (e) { calendario_().createEvent(titulo, inicio, fin, opts); }

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
