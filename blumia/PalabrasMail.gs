/* Palabras del asunto del mail para la búsqueda abierta (se ejecuta una vez desde el editor: blumiaCargarPalabrasMail).
   Agrega las palabras a la pestaña Mail_Routing sin borrar las que ya haya. Si hay UNA sola búsqueda abierta, igual todo CV
   va a esa búsqueda (unicaBusquedaAbierta_), así que esto es un refuerzo para cuando haya varias abiertas. */
const PALABRAS_MAIL_ = {
  'director-comercial': ['director', 'director comercial', 'dirección comercial', 'direccion comercial', 'gerente comercial', 'búsqueda director', 'busqueda director',
    'búsqueda comercial', 'busqueda comercial', 'consultoría', 'consultoria', 'energía', 'energia', 'minería', 'mineria', 'proyectos', 'oil & gas', 'oil and gas',
    'curriculum', 'currículum', 'cv', 'postulación', 'postulacion', 'candidato', 'candidata', 'búsqueda laboral', 'busqueda laboral']
};

function blumiaCargarPalabrasMail() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let sh = ss.getSheetByName('Mail_Routing');
  if (!sh) { sh = ss.insertSheet('Mail_Routing'); sh.appendRow(['perfil', 'palabras_asunto']); }
  const out = [];
  perfilesActivos_().forEach(p => {
    const nuevas = (PALABRAS_MAIL_[p.id] || []).slice();
    if (!nuevas.length) return;
    const data = sh.getDataRange().getValues();
    const fila = data.findIndex((r, i) => i > 0 && r[0] === p.id) + 1;
    const previas = fila ? String(data[fila - 1][1] || '').split('|').map(x => x.trim()).filter(Boolean) : [];
    const todas = previas.concat(nuevas.filter(x => previas.map(y => y.toLowerCase()).indexOf(x.toLowerCase()) < 0)).join('|');
    if (fila) sh.getRange(fila, 2).setValue(todas); else sh.appendRow([p.id, todas]);
    out.push(p.id + ': ' + todas.split('|').length + ' palabras');
  });
  Logger.log(out.length ? 'Listo. ' + out.join(' · ') : 'No hay búsquedas abiertas con palabras cargadas en el código.');
}
