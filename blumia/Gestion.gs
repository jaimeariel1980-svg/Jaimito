/* Gestión del candidato desde la ficha: observaciones (y, más adelante, edición de datos, reasignar y anular etapa). */
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
