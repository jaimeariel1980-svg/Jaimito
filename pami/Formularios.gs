function h_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }

function umbrales_() {
  const aj = ajustes_();
  return { a: Number(aj.umbral_a) || 8.5, bplus: Number(aj.umbral_bplus) || 8, b: Number(aj.umbral_b) || 6.5 };
}

function letraPorScore_(score, u) {
  return score >= u.a ? 'A' : score >= u.bplus ? 'B+' : score >= u.b ? 'B' : 'C';
}

function tablaBusq_(nombre, bid) {
  return hoja_(nombre).getDataRange().getValues().slice(1).filter(x => x[0] === bid && x[1]);
}

function getDatosForm(k, id) {
  admin_(k);
  const sh = hoja_('Candidatos');
  const fila = filaPorId_(sh, id);
  if (!fila) throw new Error('No encontré al candidato');
  const r = sh.getRange(fila, 1, 1, HEAD_CAND.length).getValues()[0];
  const bid = r[COL.busq - 1];
  const b = busquedas_().find(x => x.id === bid) || { titulo: '', ubicacion: '' };
  const aj = ajustes_();
  let resp = [];
  try { resp = JSON.parse(r[COL.resp - 1] || '[]'); } catch (e) { }
  return {
    cand: {
      id: String(id), nombre: r[COL.nombre - 1], email: r[COL.email - 1], tel: String(r[COL.tel - 1] || ''), loc: r[COL.loc - 1], pret: String(r[COL.pret - 1] || ''),
      turno: String(r[COL.turno - 1] || ''), notaCV: r[COL.ajuste - 1], notaScr: r[COL.notaScr - 1], notaComp: r[COL.notaComp - 1], entrev: r[COL.entrev - 1], infScr: String(r[COL.infScr - 1] || ''), resp: resp
    },
    hero: { puesto: b.titulo, ubicacion: b.ubicacion, cliente: aj.nombre_cliente || 'PAMI' },
    evalItems: tablaBusq_('Screening_Eval', bid).map(x => ({ id: String(x[1]), label: String(x[2]), peso: Number(x[3]) || 1, opts: opcionesEval_(x[2]) })),
    valItems: tablaBusq_('Screening_Val', bid).map(x => ({ id: String(x[1]), label: String(x[2]), opts: lista_(x[3], '|') })),
    comps: tablaBusq_('Competencias', bid).map(x => ({ id: String(x[1]), nombre: String(x[2]), descripcion: String(x[3]), peso: Number(x[4]) || 1, preguntas: lista_(x[5], '|') })),
    entrevistadores: lista_(aj.entrevistadores, ','), umbral: umbrales_(), motivos: MOTIVOS_DESCARTE, turnos: TURNOS
  };
}

// Textos del desplegable de "Ajuste al perfil" (de menor a mayor: bajo, medio, alto).
// El puntaje sigue usando bajo/medio/alto; solo cambia lo que se ve.
function opcionesEval_(etiqueta) {
  if (/t[ií]tulo terciario/i.test(String(etiqueta))) return ['Enfermero/a Profesional', 'Lic. Enfermería', 'Lic. + Post Grado'];
  return ['Bajo', 'Medio', 'Alto'];
}

function motivoDescarte_(nota, motivo) {
  if (nota !== 'C') return '';
  const m = motivoNorm_(motivo);
  if (!m) throw new Error('Con C hay que elegir el motivo de descarte');
  return m;
}

function guardarMotivoDescarte_(sh, fila, r, motivo, estado, quien) {
  if (!motivo || !/^Descartad/i.test(String(estado || ''))) return;
  sh.getRange(fila, COL.motivo).clearDataValidations().setValue(motivo);
  registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Motivo de descarte', motivo, quien);
}

function guardarEvaluacion_(sh, fila, r, etapa, p, score, nota, url, datos) {
  hoja_('Evaluaciones').appendRow([new Date(), r[COL.id - 1], r[COL.nombre - 1], etapa, p.entrevistador || '', nota, score, url, JSON.stringify(datos).slice(0, 45000)]);
}

function htmlInforme_(titulo, cand, filasMeta, cuerpo, nota, score, descNota) {
  const col = colorNota_(nota) || ['#9AA7B8', '#FFFFFF'];
  const meta = filasMeta.filter(x => x[1]).map(x => '<tr><td style="padding:4px 8px;color:#5D6B82;width:32%;font-size:11px;text-transform:uppercase;font-weight:bold">' + h_(x[0]) + '</td><td style="padding:4px 8px;font-size:13px">' + h_(x[1]) + '</td></tr>').join('');
  return '<html><head><meta charset="UTF-8"></head><body style="font-family:Arial,Helvetica,sans-serif;color:#192341;margin:0;padding:24px">' +
    '<table style="width:100%;border-bottom:4px solid #00A3E0"><tr><td style="padding-bottom:10px"><img src="' + LOGO_PAMI + '" style="height:46px"></td><td style="text-align:right;font-size:20px;font-weight:bold;color:#034EA2;text-transform:uppercase">' + h_(titulo) + '</td></tr></table>' +
    '<div style="background:#034EA2;color:#fff;padding:10px 14px;margin-top:14px;font-size:18px;font-weight:bold">' + h_(cand.nombre) + '<span style="font-size:12px;font-weight:normal"> · ' + h_(cand.puesto) + '</span></div>' +
    '<table style="width:100%;border-collapse:collapse;margin-top:6px">' + meta + '</table>' + cuerpo +
    '<table style="width:100%;margin-top:18px;border:1.5px solid #DCE3EC"><tr><td style="padding:14px;width:90px;text-align:center"><div style="background:' + col[0] + ';color:' + col[1] + ';font-size:30px;font-weight:bold;padding:10px 0">' + h_(nota) + '</div></td><td style="padding:14px"><div style="font-size:11px;text-transform:uppercase;color:#5D6B82;font-weight:bold">Resultado</div><div style="font-size:14px;margin-top:3px">' + h_(descNota) + '</div></td><td style="padding:14px;text-align:center;width:90px"><div style="font-size:11px;color:#5D6B82;text-transform:uppercase;font-weight:bold">Score</div><div style="font-size:26px;font-weight:bold">' + (score == null ? '—' : score.toFixed(1)) + '</div><div style="font-size:11px;color:#5D6B82">/ 10</div></td></tr></table>' +
    '<div style="margin-top:18px;font-size:10px;color:#5D6B82;text-transform:uppercase;letter-spacing:1px">Confidencial</div></body></html>';
}

function seccion_(t) { return '<div style="margin-top:16px;font-size:11px;font-weight:bold;color:#00A3E0;text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid #DCE3EC;padding-bottom:3px">' + h_(t) + '</div>'; }

const DESC_NOTA = { A: 'A · Avanza. Cumple lo crítico y lo deseable.', 'B+': 'B+ · Avanza. Muy bueno; cumple lo excluyente con una cuestión menor.', B: 'B · Cumple lo crítico con algún flanco; queda a criterio de la etapa.', C: 'C · No avanza. No cumple lo crítico o presenta red flags importantes.' };

function guardarScreening(k, p) {
  escritor_(k);
  const sh = hoja_('Candidatos');
  const fila = filaPorId_(sh, p.id);
  if (!fila) throw new Error('No encontré al candidato');
  const r = sh.getRange(fila, 1, 1, HEAD_CAND.length).getValues()[0];
  const d = getDatosForm(k, p.id);
  const u = d.umbral;
  const ev = p.eval || {}, vl = p.val || {};
  let score = 0, total = 0, hay = false;
  d.evalItems.forEach(it => {
    const x = ev[it.id];
    if (!x) return;
    hay = true; total += it.peso;
    if (x === 'alto') score += it.peso; else if (x === 'medio') score += it.peso * 0.5;
  });
  let norm = null, rechazado = false;
  if (hay) {
    const sal = vl.v1;
    if (sal) { total += 1.5; if (sal === 'si') score += 1.5; else if (sal === 'revisar') score += 0.75; }
    Object.keys(vl).filter(x => x !== 'v1').map(x => vl[x]).filter(Boolean).forEach(v => { total += 0.3; if (v === 'si' || v === 'alto') score += 0.3; else if (v === 'revisar' || v === 'medio') score += 0.15; });
    norm = Math.min(10, Math.round((score / total) * 100) / 10);
    rechazado = sal === 'no';
    if (rechazado) norm = Math.min(norm, 4.9);
  }
  const auto = norm == null ? null : (rechazado ? 'C' : letraPorScore_(norm, u));
  const nota = ['A', 'B+', 'B', 'C'].indexOf(p.resultado) > -1 ? p.resultado : auto;
  if (!nota) throw new Error('Completá la evaluación o elegí una letra de resultado');
  const motivoDesc = motivoDescarte_(nota, p.motivoDescarte);
  const dt = p.datos || {};
  const turno = dt.turno != null ? turnoNorm_(dt.turno) : null;
  const nivel = (v, it) => ({ bajo: it.opts[0], medio: it.opts[1], alto: it.opts[2] }[v] || '—');
  const cuerpo = seccion_('Ajuste al perfil') + '<table style="width:100%;border-collapse:collapse">' + d.evalItems.map(it => '<tr><td style="padding:5px 8px;border-bottom:1px solid #EEF1F5;font-size:13px">' + h_(it.label) + '</td><td style="padding:5px 8px;border-bottom:1px solid #EEF1F5;width:90px;text-align:center;font-weight:bold;font-size:12px;color:' + ({ alto: '#1E8E5A', medio: '#B58900', bajo: '#D9534F' }[ev[it.id]] || '#5D6B82') + '">' + h_(nivel(ev[it.id], it)) + '</td></tr>').join('') + '</table>' +
    seccion_('Validación de condiciones') + '<table style="width:100%;border-collapse:collapse">' + d.valItems.map(it => '<tr><td style="padding:5px 8px;border-bottom:1px solid #EEF1F5;font-size:13px">' + h_(it.label) + '</td><td style="padding:5px 8px;border-bottom:1px solid #EEF1F5;width:90px;text-align:center;font-weight:bold;font-size:12px">' + h_({ si: 'Sí', no: 'No', revisar: 'A revisar', alto: 'Alto', medio: 'Medio', bajo: 'Bajo' }[vl[it.id]] || '—') + '</td></tr>').join('') + '</table>' +
    (p.comentario ? seccion_('Comentario') + '<div style="font-size:13px;white-space:pre-wrap;padding-top:4px">' + h_(p.comentario) + '</div>' : '');
  const fecha = p.fecha || fmt_(new Date(), 'yyyy-MM-dd');
  const html = htmlInforme_('Screening inicial', { nombre: r[COL.nombre - 1], puesto: d.hero.puesto }, [
    ['Entrevistó', p.entrevistador], ['Fecha', fecha], ['Edad', dt.edad], ['Empresa actual', dt.empresaAct], ['Posición actual', dt.posAct], ['Formación', dt.formacion],
    ['Remuneración actual', dt.remAct], ['Remuneración pretendida', dt.remPret], ['Disponibilidad', dt.disp], ['Turno preferido', turno ? turno.replace(/; /g, ', ') : '']
  ], cuerpo, nota, norm, DESC_NOTA[nota]);
  const folder = DriveApp.getFolderById(ajustes_().carpeta_informes);
  const pdf = folder.createFile(Utilities.newBlob(html, 'text/html', 'x.html').getAs('application/pdf').setName('Screening_' + String(r[COL.nombre - 1]).replace(/\s+/g, '_') + '_' + fecha + '.pdf'));
  const url = pdf.getUrl();
  sh.getRange(fila, COL.infScr).setValue(url);
  if (p.entrevistador) sh.getRange(fila, COL.entrev).setValue(p.entrevistador);
  if (dt.remPret && !String(r[COL.pret - 1]).trim()) sh.getRange(fila, COL.pret).setValue(dt.remPret);
  guardarEvaluacion_(sh, fila, r, 'screening', p, norm, nota, url, { eval: ev, val: vl, datos: dt, comentario: p.comentario });
  if (turno !== null) {
    const antesT = String(sh.getRange(fila, COL.turno).getValue() || '');
    sh.getRange(fila, COL.turno).setValue(turno);
    if (antesT !== turno) registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Turno preferido', turno || '(sin dato)', p.entrevistador || 'formulario');
  }
  const estado = aplicarNota_(sh, fila, 'screening', nota, p.entrevistador || 'formulario');
  guardarMotivoDescarte_(sh, fila, r, motivoDesc, estado, p.entrevistador || 'formulario');
  return { ok: true, estado: estado, url: url, nota: nota, score: norm };
}

const NIVELES_COMP = ['', 'No demuestra', 'Señales de desarrollo', 'Demuestra efectividad', 'Demuestra solidez'];
const DESC_AJUSTE = { A: 'A · Alto ajuste al rol. Avanza.', 'B+': 'B+ · Muy buen ajuste al rol. Avanza.', B: 'B · Ajuste con observaciones. Queda en reserva.', C: 'C · No ajusta al rol.' };

function guardarCompetencias(k, p) {
  escritor_(k);
  const sh = hoja_('Candidatos');
  const fila = filaPorId_(sh, p.id);
  if (!fila) throw new Error('No encontré al candidato');
  const r = sh.getRange(fila, 1, 1, HEAD_CAND.length).getValues()[0];
  const d = getDatosForm(k, p.id);
  const comps = (p.comps || []).map(c => ({ nombre: String(c.nombre || '').slice(0, 200), peso: Math.max(0, Number(c.peso) || 0), nivel: Math.max(0, Math.min(4, Number(c.nivel) || 0)), notas: String(c.notas || '').slice(0, 3000) }));
  let pts = 0, max = 0;
  comps.forEach(c => { if (c.nivel > 0 && c.peso > 0) { pts += c.nivel * c.peso; max += 4 * c.peso; } });
  const score = max ? Math.round((pts / max) * 100) / 10 : null;
  const auto = score == null ? null : letraPorScore_(score, d.umbral);
  const nota = ['A', 'B+', 'B', 'C'].indexOf(p.resultado) > -1 ? p.resultado : auto;
  if (!nota) throw new Error('Evaluá al menos una competencia o elegí el ajuste al rol');
  const motivoDesc = motivoDescarte_(nota, p.motivoDescarte);
  const m = p.meta || {};
  const cuerpo = seccion_('Evaluación por competencias') + comps.map(c => '<div style="margin-top:8px;padding:8px 10px;border:1px solid #DCE3EC"><table style="width:100%"><tr><td style="font-weight:bold;font-size:13px;color:#034EA2">' + h_(c.nombre) + ' <span style="font-weight:normal;color:#5D6B82;font-size:11px">(peso ' + c.peso + ')</span></td><td style="text-align:right;font-weight:bold;font-size:12px">' + h_(NIVELES_COMP[c.nivel] || 'Sin evaluar') + '</td></tr></table>' + (c.notas ? '<div style="font-size:12px;white-space:pre-wrap;margin-top:4px">' + h_(c.notas) + '</div>' : '') + '</div>').join('') +
    (p.fortalezas ? seccion_('Fortalezas principales') + '<div style="font-size:13px;white-space:pre-wrap">' + h_(p.fortalezas) + '</div>' : '') +
    (p.areas ? seccion_('Áreas a desarrollar') + '<div style="font-size:13px;white-space:pre-wrap">' + h_(p.areas) + '</div>' : '') +
    (p.recomendacion ? seccion_('Recomendación y justificación') + '<div style="font-size:13px;white-space:pre-wrap">' + h_(p.recomendacion) + '</div>' : '');
  const fecha = p.fecha || fmt_(new Date(), 'yyyy-MM-dd');
  const html = htmlInforme_('Entrevista por competencias', { nombre: r[COL.nombre - 1], puesto: d.hero.puesto }, [
    ['Entrevistó', p.entrevistador], ['Fecha', fecha], ['Edad', m.edad], ['Localidad', m.localidad], ['Situación familiar', m.familia], ['Formación', m.formacion],
    ['Empresa actual', m.empresaAct], ['Puesto actual', m.puestoAct], ['Antigüedad', m.antig], ['Remuneración actual', m.remAct], ['Remuneración pretendida', m.remPret],
    ['Disponibilidad de ingreso', m.disp], ['Motivo de cambio', m.motivo], ['Movilidad', m.movilidad], ['Referencias laborales', m.referencias]
  ], cuerpo, nota, score, DESC_AJUSTE[nota]);
  const folder = DriveApp.getFolderById(ajustes_().carpeta_informes);
  const pdf = folder.createFile(Utilities.newBlob(html, 'text/html', 'x.html').getAs('application/pdf').setName('Competencias_' + String(r[COL.nombre - 1]).replace(/\s+/g, '_') + '_' + fecha + '.pdf'));
  const url = pdf.getUrl();
  sh.getRange(fila, COL.infComp).setValue(url);
  guardarEvaluacion_(sh, fila, r, 'competencias', p, score, nota, url, { comps: comps, meta: m, fortalezas: p.fortalezas, areas: p.areas, recomendacion: p.recomendacion });
  const estado = aplicarNota_(sh, fila, 'competencias', nota, p.entrevistador || 'formulario');
  guardarMotivoDescarte_(sh, fila, r, motivoDesc, estado, p.entrevistador || 'formulario');
  return { ok: true, estado: estado, url: url, nota: nota, score: score };
}
