/* Nombre de la empresa cliente según el perfil (Config.empresa). Antes estaba «MACS» escrito a mano.
   - empresaInforme_: lo que ve el cliente en los informes (MACS queda como «MACS S.A.»).
   - empresaParaCandidatos_: lo que leen los candidatos en mails y WhatsApp. Si Config tiene la columna opcional
     `empresa_candidatos`, se usa tal cual (ej.: «una empresa cliente del sector Energía y Minería» o el nombre real);
     si no, para una búsqueda que no es de MACS no se revela la empresa: «una empresa cliente del sector <empresa>». */
function empresaInforme_(cfg) {
  const e = (cfg && typeof cfg.empresa === 'string') ? cfg.empresa.trim() : '';
  return (!e || e === 'MACS') ? 'MACS S.A.' : e;
}

function empresaParaCandidatos_(cfg) {
  const x = (cfg && typeof cfg.empresa_candidatos === 'string') ? cfg.empresa_candidatos.trim() : '';
  if (x) return x;
  const e = (cfg && typeof cfg.empresa === 'string') ? cfg.empresa.trim() : '';
  return (!e || e === 'MACS') ? 'MACS' : 'una empresa cliente del sector ' + e;
}

/* Clave de la empresa de un perfil (Config.empresa; sin dato = MACS) y nombre de la hoja de su informe de proceso. */
function empresaKey_(cfg) {
  const e = (cfg && typeof cfg.empresa === 'string') ? cfg.empresa.trim() : '';
  return e || 'MACS';
}

function hojaProceso_(empresa) {
  return (!empresa || empresa === 'MACS') ? 'Informe_Proceso' : 'Informe_Proceso_' + empresa;
}
