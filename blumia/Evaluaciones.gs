/* Respuestas de cada formulario de screening / competencias guardadas para poder editarlas después (como en PAMI).
   El estado completo del formulario va en un .json en la carpeta de Drive de la etapa; la pestaña «Evaluaciones»
   guarda [fecha, perfil, candidato, etapa, nota, quién, archivo]. */
const HOJA_EVAL_ = 'Evaluaciones';

function evalNorm_(s) { return String(s || '').toLowerCase().trim().normalize('NFD').replace(/[̀-ͯ]/g, '').replace(/\s+/g, ' '); }

function evalHoja_() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let sh = ss.getSheetByName(HOJA_EVAL_);
  if (!sh) {
    sh = ss.insertSheet(HOJA_EVAL_);
    sh.getRange(1, 1, 1, 7).setValues([['fecha', 'perfil', 'candidato', 'etapa', 'nota', 'quien', 'archivo']]).setFontWeight('bold').setBackground('#243D2C').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function evalFila_(sh, perfilId, nombre, etapa) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const d = sh.getRange(2, 1, n - 1, 7).getValues();
  const k = evalNorm_(nombre);
  for (let i = d.length - 1; i >= 0; i--) if (String(d[i][1]) === String(perfilId) && String(d[i][3]) === etapa && evalNorm_(d[i][2]) === k) return i + 2;
  return 0;
}

// estado de la fila que corresponde a una nota (la misma regla de aplicarEstadoNota)
function estadoPorNota_(etapa, nota) {
  let v = String(nota || '').trim().toUpperCase();
  const m = { 'ALTO': 'A', 'MEDIO': 'B', 'BAJO': 'C' };
  if (m[v]) v = m[v];
  if (['A', 'B+', 'B', 'C'].indexOf(v) === -1) return '';
  if (v === 'C') return { cv: 'Descartado sin screening', screening: 'Descartado post screening', competencias: 'Descartado en competencias', empresa: 'Descartado en empresa' }[etapa];
  if (v === 'B' && etapa !== 'cv' && etapa !== 'screening') return etapa === 'competencias' ? 'Reserva competencias' : 'En reserva';
  return { cv: 'Avanza a screening', screening: 'Avanza a competencias', competencias: 'Avanza a screening empresa', empresa: 'En cliente' }[etapa];
}

// ¿corresponde cambiar el estado al editar una evaluación? Solo si la nota cambió y el estado todavía es el que salía de la nota anterior
function aplicarEdicion_(sheet, fila, etapa, notaAntes, notaNueva) {
  if (String(notaAntes || '').trim() === String(notaNueva || '').trim()) return false;
  if (!String(notaAntes || '').trim()) return true;
  const est = String(sheet.getRange(fila, 9).getValue() || '').trim();
  return est === estadoPorNota_(etapa, notaAntes);
}

function guardarEstadoEval_(cfg, perfilId, nombre, etapa, estadoJSON, nota, quien) {
  if (!estadoJSON) return;
  try {
    const sh = evalHoja_();
    const fila = evalFila_(sh, perfilId, nombre, etapa);
    let id = fila ? String(sh.getRange(fila, 7).getValue() || '') : '';
    if (id) { try { DriveApp.getFileById(id).setContent(String(estadoJSON)); } catch (e) { id = ''; } }
    if (!id) {
      const carpeta = DriveApp.getFolderById(etapa === 'screening' ? cfg.folder_pdf_screening : cfg.folder_pdf_competencias);
      id = carpeta.createFile('Estado_' + etapa + '_' + String(nombre).replace(/\s+/g, '_') + '.json', String(estadoJSON), MimeType.PLAIN_TEXT).getId();
    }
    const vals = [new Date(), perfilId, nombre, etapa, nota || '', quienHist_(quien), id];
    if (fila) sh.getRange(fila, 1, 1, 7).setValues([vals]); else sh.appendRow(vals);
  } catch (e) { console.error('No pude guardar las respuestas de la evaluación: ' + e); }
}

function getEvaluacionPrevia_(perfilId, nombre, etapa) {
  if (['screening', 'competencias'].indexOf(etapa) === -1) throw new Error('Etapa inválida');
  const sh = evalHoja_();
  const fila = evalFila_(sh, perfilId, nombre, etapa);
  if (!fila) return null;
  const r = sh.getRange(fila, 1, 1, 7).getValues()[0];
  let estado = '';
  try { estado = DriveApp.getFileById(String(r[6])).getBlob().getDataAsString(); } catch (e) { return null; }
  return { estado: estado, nota: String(r[4] || ''), quien: String(r[5] || ''), fecha: Utilities.formatDate(new Date(r[0]), 'GMT-3', 'dd/MM/yyyy HH:mm') };
}
