const PESOS_FINAL_ = { cv: 10, screening: 15, competencias: 35, tecnica: 40 };
const ETAPAS_FINAL_ = [['cv', 'CV'], ['screening', 'Screening'], ['competencias', 'Competencias'], ['tecnica', 'Técnica']];

function pesosFinal_() {
  const aj = ajustes_(), p = {};
  Object.keys(PESOS_FINAL_).forEach(k => {
    const v = parseFloat(String(aj['peso_' + k] || '').replace(',', '.'));
    p[k] = isFinite(v) && v >= 0 ? v : PESOS_FINAL_[k];
  });
  return p;
}

function notaLetra_(n) {
  const v = String(n || '').trim().toUpperCase(), m = { ALTO: 'A', MEDIO: 'B', BAJO: 'C' };
  return m[v] || v;
}

function puntajeEtapa_(nota, score) {
  const s = Number(score);
  if (score !== '' && score != null && isFinite(s) && s >= 0 && s <= 10) return s;
  const L = { A: 9.5, 'B+': 8.25, B: 7.25, C: 4 }[notaLetra_(nota)];
  return L == null ? null : L;
}

function notaFinal_(partes, pesos) {
  let sp = 0, sw = 0;
  ETAPAS_FINAL_.forEach(e => { const x = partes[e[0]]; if (x && x.puntaje != null && pesos[e[0]] > 0) { sp += x.puntaje * pesos[e[0]]; sw += pesos[e[0]]; } });
  return sw ? Math.round(sp / sw * 100) / 100 : null;
}

function partesCand_(ajuste, evs) {
  const nota = (e) => e ? notaLetra_(e.nota) : '';
  const eta = { screening: 'screening', competencias: 'competencias', tecnica: 'pami2' }, out = {};
  out.cv = { nota: notaLetra_(ajuste), puntaje: puntajeEtapa_(ajuste, null) };
  Object.keys(eta).forEach(k => { const e = evs[eta[k]]; out[k] = e ? { nota: nota(e), puntaje: puntajeEtapa_(e.nota, e.score), entrevisto: e.entrevisto, fecha: e.fecha } : null; });
  return out;
}

function evalsPorId_() {
  const ev = hoja_('Evaluaciones'), por = {};
  if (ev && ev.getLastRow() > 1) ev.getRange(2, 1, ev.getLastRow() - 1, 9).getValues().forEach(x => {
    if (!x[1] || !x[3]) return;
    (por[String(x[1])] = por[String(x[1])] || {})[String(x[3])] = { fecha: x[0] instanceof Date ? fmt_(x[0], 'dd/MM/yyyy') : String(x[0] || ''), entrevisto: String(x[4] || ''), nota: String(x[5] || ''), score: x[6], datos: x[8] };
  });
  return por;
}

function jsonSeguro_(s) { try { return typeof s === 'string' ? JSON.parse(s || '{}') : (s || {}); } catch (e) { return {}; } }

function listaNiveles_(comps) {
  const c = (comps || []).filter(x => x.nivel > 0).sort((a, b) => b.nivel - a.nivel);
  return c.length ? c.map(x => x.nombre + ' ' + x.nivel + '/4').join(', ') : '';
}

function textoFinal_(r, evs, partes, fin, letra, pos, total) {
  const S = jsonSeguro_((evs.screening || {}).datos), C = jsonSeguro_((evs.competencias || {}).datos), T = jsonSeguro_((evs.pami2 || {}).datos);
  const sd = S.datos || {}, cm = C.meta || {};
  const g = (...a) => a.map(x => String(x || '').trim()).find(Boolean) || '';
  const nombre = String(r[COL.nombre - 1] || ''), edad = g(sd.edad, cm.edad), form = g(r[COL.formacion - 1], sd.formacion, cm.formacion);
  const loc = String(r[COL.loc - 1] || ''), zona = String(r[COL.zona - 1] || ''), pol = String(r[COL.policlinico - 1] || '');
  const exp = [g(sd.posAct, cm.puestoAct), g(sd.empresaAct, cm.empresaAct)].filter(Boolean).join(' en ');
  const pret = g(r[COL.pret - 1], sd.remPret, cm.remPret), disp = g(sd.disp, cm.disp), turno = String(r[COL.turno - 1] || '').replace(/;\s*/g, ', ');
  const sec = [];
  const perfil = [nombre + (edad ? ' (' + edad + ' años)' : '') + (form ? ', ' + form : '') + (loc ? ', vive en ' + loc : '') + (zona ? ' (' + zona + ')' : '') + '.'];
  if (exp) perfil.push('Actualmente trabaja como ' + exp + '.');
  if (pret) perfil.push('Pretensión salarial: ' + pret + '.');
  if (disp) perfil.push('Disponibilidad: ' + disp + '.');
  if (turno) perfil.push('Turno preferido: ' + turno + '.');
  if (pol) perfil.push('Policlínico asignado: ' + pol + '.');
  sec.push({ t: 'Perfil', p: perfil.join(' ') });
  const cv = String(r[COL.just - 1] || '').replace(/^IA sugiere [A-C+]+:\s*/, '').trim();
  if (cv) sec.push({ t: 'Análisis del CV (' + (partes.cv.nota || '–') + ')', p: cv });
  const cab = (p) => p ? [p.nota ? 'Nota ' + p.nota : '', p.puntaje != null ? p.puntaje.toFixed(1) + '/10' : '', p.entrevisto ? 'entrevistó ' + p.entrevisto : '', p.fecha || ''].filter(Boolean).join(' · ') : '';
  if (partes.screening) sec.push({ t: 'Screening (' + cab(partes.screening) + ')', p: g(S.comentario) || 'Sin comentario cargado.' });
  const bloque = (d, p) => [listaNiveles_(d.comps) ? 'Evaluación: ' + listaNiveles_(d.comps) + '.' : '', d.fortalezas ? 'Fortalezas: ' + d.fortalezas : '', d.areas ? 'A desarrollar: ' + d.areas : '', g(d.recomendacion, d.comentario) ? 'Recomendación: ' + g(d.recomendacion, d.comentario) : ''].filter(Boolean).join('\n') || 'Sin comentarios cargados.';
  if (partes.competencias) sec.push({ t: 'Competencias (' + cab(partes.competencias) + ')', p: bloque(C) });
  if (partes.tecnica) sec.push({ t: 'Entrevista técnica (' + cab(partes.tecnica) + ')', p: bloque(T) });
  sec.push({ t: 'Conclusión', p: nombre + ' obtiene una nota final de ' + fin.toFixed(1) + '/10 (' + letra + ')' + (total > 1 ? ', puesto ' + pos + ' de ' + total + ' entre los candidatos de esta búsqueda con entrevista técnica' : '') + '.' });
  return sec;
}

function getInformeFinal(k, fila) {
  soloAdmin_(k);
  const sh = hoja_('Candidatos');
  fila = Number(fila);
  if (!(fila >= 2 && fila <= sh.getLastRow())) throw new Error('Fila inválida');
  const filas = sh.getRange(2, 1, sh.getLastRow() - 1, HEAD_CAND.length).getValues();
  const r = filas[fila - 2], id = String(r[COL.id - 1]);
  const por = evalsPorId_(), pesos = pesosFinal_(), u = umbrales_();
  const evs = por[id] || {};
  if (!evs.pami2) throw new Error('Todavía no tiene entrevista técnica cargada');
  const partes = partesCand_(r[COL.ajuste - 1], evs), fin = notaFinal_(partes, pesos);
  if (fin == null) throw new Error('No hay notas para calcular');
  const busq = String(r[COL.busq - 1] || ''), rank = [];
  filas.forEach(x => {
    const i = String(x[COL.id - 1]), e = por[i];
    if (!e || !e.pami2 || String(x[COL.busq - 1] || '') !== busq) return;
    const f = notaFinal_(partesCand_(x[COL.ajuste - 1], e), pesos);
    if (f != null) rank.push(f);
  });
  const pos = rank.filter(f => f > fin).length + 1, letra = letraPorScore_(fin, u);
  return JSON.stringify({
    nombre: String(r[COL.nombre - 1] || ''), busq: busq, policlinico: String(r[COL.policlinico - 1] || ''),
    pesos: pesos, partes: partes, final: fin, letra: letra, pos: pos, total: rank.length,
    secciones: textoFinal_(r, evs, partes, fin, letra, pos, rank.length)
  });
}
