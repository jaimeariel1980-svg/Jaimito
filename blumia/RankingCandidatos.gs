/* Ranking para el cliente (pedido de Ariel 29/10): se arma en la pestaña «Ranking» del tablero, con las mismas fichas que el tablero.
   El orden y la valoración se guardan en las columnas 35 «Orden cliente» y 36 «Valoración cliente» de la hoja de candidatos (las mismas que ya
   usaba «PDF para el cliente»): generarPDFClientes arma el PDF con los informes individuales en ese orden. No hay un segundo informe. */

function rankingGuardar_(perfilId, lista) {
  const cfg = getConfig(perfilId);
  const sh = getSheet(cfg);
  if (sh.getMaxColumns() < 36) sh.insertColumnsAfter(sh.getMaxColumns(), 36 - sh.getMaxColumns());
  if (!sh.getRange(1, 35).getValue()) {
    sh.getRange(1, 35).setValue('Orden cliente');
    sh.getRange(1, 36).setValue('Valoración cliente');
    sh.getRange(1, 35, 1, 2).setFontWeight('bold').setBackground('#0F3F27').setFontColor('#FFFFFF');
  }
  const n = sh.getLastRow() - 1;
  if (n < 1) return { ok: true, n: 0 };
  const VAL = ['Se ajusta', 'Se ajusta con observaciones'];
  const nuevo = [];
  for (let i = 0; i < n; i++) nuevo.push(['', '']);
  let k = 0;
  (Array.isArray(lista) ? lista : []).slice(0, 60).forEach(o => {
    const f = Number(o && o.fila);
    if (!(f >= 2 && f <= n + 1) || nuevo[f - 2][0] !== '') return;
    k++;
    nuevo[f - 2] = [k, VAL.indexOf(String(o.val)) >= 0 ? String(o.val) : VAL[0]];
  });
  sh.getRange(2, 35, n, 2).setValues(nuevo);
  try { registrarHistorial_(sh.getName(), '(búsqueda)', 'Ranking', 'Ranking para el cliente guardado (' + k + ' personas)', QUIEN_); } catch (e) { }
  return { ok: true, n: k };
}
