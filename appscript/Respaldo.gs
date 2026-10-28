// Copia de seguridad diaria de la planilla Maestra y registro de errores con aviso diario a la gerencia.
var AZCU_COPIAS_CARPETA_ = 'Copias de seguridad Azcu', AZCU_COPIAS_MAX_ = 14;

function azcuCopiaDiaria(){
  var P = PropertiesService.getScriptProperties();
  if(Date.now() - (+P.getProperty('AZCU_ULTIMA_COPIA') || 0) < 12*3600000) return;
  try{
    var raiz = DriveApp.getFolderById('1LuYwU-I-hE1lEk6rcSblGmMVqzaPUdAh'), it = raiz.getFoldersByName(AZCU_COPIAS_CARPETA_), carpeta = it.hasNext() ? it.next() : raiz.createFolder(AZCU_COPIAS_CARPETA_);
    var fecha = Utilities.formatDate(new Date(), AZCU_TZ, 'yyyy-MM-dd HH:mm');
    DriveApp.getFileById(MAESTRO_ID).makeCopy('Azcu · copia '+fecha, carpeta);
    var copias = [], fs = carpeta.getFiles();
    while(fs.hasNext()){ var f = fs.next(); if(/^Azcu · copia /.test(f.getName())) copias.push(f); }
    copias.sort(function(a,b){ return b.getDateCreated() - a.getDateCreated(); });
    copias.slice(AZCU_COPIAS_MAX_).forEach(function(f){ f.setTrashed(true); });
    P.setProperty('AZCU_ULTIMA_COPIA', String(Date.now()));
    P.setProperty('AZCU_ULTIMA_COPIA_TXT', fecha);
  }catch(e){ azcuErrorLog_('Copia de seguridad', e); }
}

function azcuErrorLog_(origen, e){
  try{
    var msg = String((e && e.message) || e || '').slice(0, 300);
    if(!msg || /PIN|intentos/i.test(msg)) return;
    var ss = SpreadsheetApp.openById(MAESTRO_ID), sh = ss.getSheetByName('Errores') || ss.insertSheet('Errores');
    if(sh.getLastRow()===0) sh.appendRow(['Fecha','Origen','Error']);
    sh.appendRow([new Date(), String(origen||'').slice(0, 60), msg]);
    if(sh.getLastRow() > 600) sh.deleteRows(2, 100);
  }catch(x){}
}

function azcuErroresDiario(){
  try{
    var sh = SpreadsheetApp.openById(MAESTRO_ID).getSheetByName('Errores');
    if(!sh || sh.getLastRow()<2) return;
    var desde = Date.now() - 24*3600000, g = {}, n = 0, v = sh.getRange(2, 1, sh.getLastRow()-1, 3).getValues();
    v.forEach(function(r){ var t = r[0] instanceof Date ? r[0].getTime() : 0; if(t >= desde){ n++; var k = r[1]+': '+String(r[2]).slice(0, 90); g[k] = (g[k]||0)+1; } });
    if(!n) return;
    var top = Object.keys(g).sort(function(a,b){ return g[b]-g[a]; }).slice(0, 2).map(function(k){ return k+(g[k]>1 ? ' (x'+g[k]+')' : ''); }).join(' · ');
    var dest = []; azcuGerentes_().forEach(function(nm){ azcuDestinos_(nm, true).forEach(function(d){ if(dest.indexOf(d)<0) dest.push(d); }); });
    if(dest.length) azcuPush_('⚠️ Azcu tuvo '+n+(n===1 ? ' error' : ' errores')+' en las últimas 24 h', top+'. Detalle en la pestaña Errores.', AZCU_WEB, dest, true, true);
  }catch(e){ Logger.log('azcuErroresDiario: '+e); }
}

// se llama desde azcuCalentar (cada 5 min): crea solo los dos activadores diarios si no existen
function azcuMantenimiento_(){
  var c = CacheService.getScriptCache();
  if(c.get('azmant')) return;
  c.put('azmant', '1', 21600);
  try{
    var hs = ScriptApp.getProjectTriggers().map(function(t){ return t.getHandlerFunction(); });
    if(hs.indexOf('azcuCopiaDiaria')<0) ScriptApp.newTrigger('azcuCopiaDiaria').timeBased().everyDays(1).atHour(3).create();
    if(hs.indexOf('azcuErroresDiario')<0) ScriptApp.newTrigger('azcuErroresDiario').timeBased().everyDays(1).atHour(9).create();
  }catch(e){ Logger.log('azcuMantenimiento_: '+e); }
}
