/* Copia de seguridad diaria de la planilla de Blumia (toda: candidatos, Config, formularios, historial).
   Se guarda en la carpeta «Copias de seguridad Blumia» de Drive y se conservan las últimas 14. */
const COPIAS_CARPETA_ = 'Copias de seguridad Blumia';
const COPIAS_MAX_ = 14;

function hacerCopiaBlumia_() {
  const it = DriveApp.getFoldersByName(COPIAS_CARPETA_);
  const carpeta = it.hasNext() ? it.next() : DriveApp.createFolder(COPIAS_CARPETA_);
  const fecha = Utilities.formatDate(new Date(), 'America/Argentina/Buenos_Aires', 'yyyy-MM-dd HH:mm');
  DriveApp.getFileById(SHEET_MASTER_ID).makeCopy('Blumia · copia ' + fecha, carpeta);
  // se queda con las últimas COPIAS_MAX_ (solo borra copias hechas por esta función, dentro de su carpeta)
  const copias = [];
  const fs = carpeta.getFiles();
  while (fs.hasNext()) { const f = fs.next(); if (/^Blumia · copia /.test(f.getName())) copias.push(f); }
  copias.sort((a, b) => b.getDateCreated() - a.getDateCreated());
  copias.slice(COPIAS_MAX_).forEach(f => f.setTrashed(true));
  PropertiesService.getScriptProperties().setProperty('ULTIMA_COPIA_BLUMIA', fecha);
  return fecha;
}

function copiaDiaria() {
  try { hacerCopiaBlumia_(); registrarHistorial_('(sistema)', '', 'Copia de seguridad', 'Automática', 'sistema'); } catch (e) { Logger.log('copiaDiaria: ' + e); }
}

// se llama desde el ciclo de 30 min: si todavía no existe el activador diario, lo crea solo
function asegurarCopiaDiaria_() {
  try {
    if (ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'copiaDiaria')) return;
    ScriptApp.newTrigger('copiaDiaria').timeBased().everyDays(1).atHour(3).create();
  } catch (e) { Logger.log('asegurarCopiaDiaria_: ' + e); }
}
