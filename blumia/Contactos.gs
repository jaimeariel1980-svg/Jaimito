const HOJA_REG_CONT_ = 'Registro_Contactos';

function hojaRegCont_() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let sh = ss.getSheetByName(HOJA_REG_CONT_);
  if (!sh) {
    sh = ss.insertSheet(HOJA_REG_CONT_);
    sh.getRange(1, 1, 1, 7).setValues([['fecha', 'perfil', 'fila', 'nombre', 'tel', 'tipo', 'quien']]).setFontWeight('bold').setBackground('#1A6645').setFontColor('#fff');
    sh.setFrozenRows(1);
  }
  return sh;
}

function registrarContacto_(perfilId, fila, nombre, tel, tipo, quien) {
  quien = String(quien || '').trim();
  if (!quien) { try { quien = Session.getActiveUser().getEmail() || ''; } catch (e) { } }
  try { hojaRegCont_().appendRow([new Date(), perfilId, fila, nombre, String(tel || ''), tipo, quien]); } catch (e) { }
}

function registrarContactoManual(perfilId, fila, tipo, quien) {
  const sh = getSheet(getConfig(perfilId));
  registrarContacto_(perfilId, fila, sh.getRange(fila, 2).getValue(), sh.getRange(fila, 3).getValue(), tipo, quien);
  return { ok: true };
}

function getContactosRegistro(perfilId, fila) {
  const nombre = String(getSheet(getConfig(perfilId)).getRange(fila, 2).getValue() || '').trim();
  const sh = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName(HOJA_REG_CONT_);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 7).getValues()
    .filter(r => r[1] === perfilId && String(r[3]).trim() === nombre)
    .map(r => ({ fecha: Utilities.formatDate(new Date(r[0]), 'GMT-3', 'dd/MM/yyyy HH:mm'), tipo: r[5], quien: r[6] }))
    .reverse();
}

function exportarContactosCSV() {
  const OK = ['Contactado', 'Screening agendado', 'Avanza a competencias', 'Competencias agendada', 'Avanza a empresa', 'Entrevista empresa agendada', 'En cliente', 'Contratado'];
  const q = v => '"' + String(v == null ? '' : v).replace(/"/g, '""') + '"';
  const filas = [['Name', 'Given Name', 'Phone 1 - Value', 'E-mail 1 - Value', 'Group Membership'].map(q).join(',')];
  let n = 0;
  listarPerfiles().forEach(p => {
    let cfg, sh;
    try { cfg = getConfig(p.id); sh = getSheet(cfg); } catch (e) { return; }
    if (!sh || sh.getLastRow() < 2) return;
    const grupo = 'Blumia ' + (cfg.empresa ? cfg.empresa + ' - ' : '') + p.nombre;
    sh.getDataRange().getValues().slice(1).forEach(r => {
      const nombre = String(r[1] || '').trim(), tel = String(r[2] || '').replace(/[^0-9+]/g, ''), mail = String(r[21] || '').trim();
      if (!nombre || !tel || OK.indexOf(String(r[8] || '').trim()) === -1) return;
      const nc = p.nombre + ' - ' + nombre;
      filas.push([nc, nc, tel, mail.indexOf('@') > -1 ? mail : '', grupo].map(q).join(','));
      n++;
    });
  });
  return { ok: true, n: n, csv: '﻿' + filas.join('\r\n') };
}
