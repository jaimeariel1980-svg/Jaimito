/* Papelera de candidatos del kanban. «Eliminar» no borra nada: el Estado pasa a «Eliminado|<estado anterior>|<fecha>»
   (el tablero lo esconde y lo muestra en 🗑 Papelera, de donde se restaura). */
function filaCandidato_(sh, fila) {
  fila = Number(fila);
  if (!fila || fila < 2 || fila > sh.getLastRow()) throw new Error('Candidato inválido');
  if (!String(sh.getRange(fila, 2).getValue() || '').trim()) throw new Error('Esa fila no tiene un candidato');
  return fila;
}

function eliminarCandidato(perfilId, fila) {
  const sh = getSheet(getConfig(perfilId));
  fila = filaCandidato_(sh, fila);
  const estado = String(sh.getRange(fila, 9).getValue() || '').trim();
  if (/^Eliminado\|/.test(estado)) return { ok: true };
  const hoy = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'dd/MM/yyyy');
  sh.getRange(fila, 9).setValue('Eliminado|' + estado + '|' + hoy);
  return { ok: true };
}

function restaurarCandidato(perfilId, fila) {
  const sh = getSheet(getConfig(perfilId));
  fila = filaCandidato_(sh, fila);
  const estado = String(sh.getRange(fila, 9).getValue() || '').trim();
  if (!/^Eliminado\|/.test(estado)) return { ok: true };
  sh.getRange(fila, 9).setValue(estado.split('|')[1] || 'Nuevo');
  return { ok: true };
}
