const MOTIVOS_DESCARTE_ = ['Fuera de zona', 'Falta de experiencia en el rubro', 'Falta de experiencia en el puesto', 'Formación no adecuada', 'Remuneración fuera de rango', 'No cumple un requisito excluyente', 'No le interesa', 'Incompatibilidad de modalidad u horarios', 'No respondió / no se presentó', 'Eligió otra oferta'];

function colMotivo_(sh, crear) {
  const h = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0];
  let c = h.indexOf('Motivo de descarte') + 1;
  if (!c && crear) {
    c = Math.max(sh.getLastColumn(), 45) + 1;
    if (sh.getMaxColumns() < c) sh.insertColumnsAfter(sh.getMaxColumns(), c - sh.getMaxColumns());
    sh.getRange(1, c).setValue('Motivo de descarte').setFontWeight('bold');
  }
  return c;
}

// motivos elegidos de la lista (más «Otro» de hasta tres palabras) → texto separado por «; »
function motivoNormBl_(lista, otro) {
  const out = [];
  (lista || []).forEach(m => {
    m = String(m).trim();
    if (MOTIVOS_DESCARTE_.indexOf(m) === -1) throw new Error('Motivo inválido: ' + m);
    if (out.indexOf(m) === -1) out.push(m);
  });
  const o = String(otro || '').replace(/[<>;]/g, '').trim();
  if (o) {
    if (o.split(/\s+/).length > 3) throw new Error('En «Otro» escribí el motivo con máximo tres palabras.');
    out.push('Otro: ' + o);
  }
  if (!out.length) throw new Error('Elegí al menos un motivo de descarte.');
  return out.join('; ');
}

function guardarMotivoDescarte_(perfilId, fila, motivos, otro, quien) {
  QUIEN_ = quien || '';
  const sh = getSheet(getConfig(perfilId));
  if (!/^Descartad/i.test(String(sh.getRange(fila, 9).getValue() || '').trim())) throw new Error('El candidato no está descartado.');
  const texto = motivoNormBl_(motivos, otro);
  sh.getRange(fila, colMotivo_(sh, true)).setNumberFormat('@').setValue(texto);
  registrarHistorial_(sh.getName(), sh.getRange(fila, 2).getValue(), 'Motivo de descarte', texto);
  return { ok: true, motivo: texto };
}

// si el candidato vuelve a una etapa activa, el motivo ya no corresponde
function limpiarMotivoSiNoDescartado_(sh, fila, estado) {
  if (/^Descartad/i.test(String(estado || ''))) return;
  const c = colMotivo_(sh, false);
  if (c && String(sh.getRange(fila, c).getValue() || '').trim()) sh.getRange(fila, c).clearContent();
}
