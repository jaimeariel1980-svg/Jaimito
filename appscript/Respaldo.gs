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

// 👎 del chat: guarda la pregunta y la respuesta (sin teléfonos, mails ni DNI) 30 días para poder arreglar el bot
function azcuFeedback(pin, nombre, pregunta, respuesta){
  azcuPin_(pin);
  function lim(t){ t = String(t||'').replace(/[\w.+-]+@[\w-]+\.[\w.-]+/g, '[mail]').replace(/\+?\d[\d\s().-]{6,}\d/g, '[nº]').replace(/\s+/g, ' ').trim().slice(0, 600); return /^[=+\-@]/.test(t) ? "'"+t : t; }
  var lk = LockService.getScriptLock(); lk.waitLock(10000);
  try{
    var ss = SpreadsheetApp.openById(MAESTRO_ID), sh = ss.getSheetByName('Chat_Fallas') || ss.insertSheet('Chat_Fallas');
    if(sh.getLastRow()===0) sh.appendRow(['Fecha','Quién','Pregunta','Respuesta del bot','Modelo']);
    var lim30 = Date.now() - 30*86400000, v = sh.getLastRow()>1 ? sh.getRange(2, 1, sh.getLastRow()-1, 1).getValues() : [], k = 0;
    while(k<v.length && v[k][0] instanceof Date && v[k][0].getTime() < lim30) k++;
    if(k) sh.deleteRows(2, k);
    sh.appendRow([new Date(), lim(nombre).slice(0, 60), lim(pregunta), lim(respuesta), typeof azcuModelo_==='function' ? azcuModelo_() : '']);
  } finally { lk.releaseLock(); }
  return {ok:true};
}

function azcuErroresDiario(){
  try{
    var ss = SpreadsheetApp.openById(MAESTRO_ID), desde = Date.now() - 24*3600000, g = {}, n = 0, nd = 0, top = '', sh = ss.getSheetByName('Errores'), sf = ss.getSheetByName('Chat_Fallas');
    if(sh && sh.getLastRow()>1){
      sh.getRange(2, 1, sh.getLastRow()-1, 3).getValues().forEach(function(r){ var t = r[0] instanceof Date ? r[0].getTime() : 0; if(t >= desde){ n++; var k = r[1]+': '+String(r[2]).slice(0, 90); g[k] = (g[k]||0)+1; } });
      top = Object.keys(g).sort(function(a,b){ return g[b]-g[a]; }).slice(0, 2).map(function(k){ return k+(g[k]>1 ? ' (x'+g[k]+')' : ''); }).join(' · ');
    }
    if(sf && sf.getLastRow()>1) sf.getRange(2, 1, sf.getLastRow()-1, 1).getValues().forEach(function(r){ if(r[0] instanceof Date && r[0].getTime() >= desde) nd++; });
    if(!n && !nd) return;
    var dest = []; azcuGerentes_().forEach(function(nm){ azcuDestinos_(nm, true).forEach(function(d){ if(dest.indexOf(d)<0) dest.push(d); }); });
    var tit = n ? '⚠️ Azcu tuvo '+n+(n===1 ? ' error' : ' errores')+' en las últimas 24 h' : '👎 El chat de Azcu tuvo '+nd+(nd===1 ? ' respuesta marcada' : ' respuestas marcadas')+' en las últimas 24 h';
    var cuerpo = (n ? top+'. Detalle en la pestaña Errores.' : '')+(n && nd ? ' ' : '')+(nd ? (n ? 'Además ' : '')+nd+' 👎 en el chat (pestaña Chat_Fallas).' : (n ? '' : ''));
    if(dest.length) azcuPush_(tit, cuerpo || 'Detalle en la pestaña Chat_Fallas.', AZCU_WEB, dest, true, true);
  }catch(e){ Logger.log('azcuErroresDiario: '+e); }
}

// resumen semanal de los 👎 del chat (lunes 9 h): cuántas respuestas se marcaron y las primeras preguntas, para pasárselas a Claude
// manda por mail el detalle de los 👎 (ya sin teléfonos ni mails) para que Claude los revise y corrija; destino: propiedad AZCU_FALLAS_MAIL o el mail de la gerencia en la pestaña Usuarios
function azcuFallasMail_(l){
  try{
    var to = String(PropertiesService.getScriptProperties().getProperty('AZCU_FALLAS_MAIL')||'').trim();
    if(!to){ var ms = []; azcuGerentes_().forEach(function(nm){ try{ var u = azcuUsuarioDe_(nm); if(u && u.mail && ms.indexOf(u.mail)<0) ms.push(u.mail); }catch(e){} }); to = ms.join(','); }
    if(!to) return;
    var cuerpo = l.slice(0, 40).map(function(x, i){ return (i+1)+'. '+Utilities.formatDate(x.f, AZCU_TZ, 'dd/MM/yyyy HH:mm')+' · '+x.u+' · '+x.m+'\nPREGUNTA: '+x.p.slice(0, 500)+'\nRESPUESTA: '+x.a.slice(0, 500); }).join('\n\n');
    MailApp.sendEmail({to:to, subject:'[Azcu fallas] '+l.length+(l.length===1 ? ' respuesta marcada' : ' respuestas marcadas')+' del chat esta semana', body:'Respuestas del chat marcadas con 👎 en los últimos 7 días (sin teléfonos ni mails).\n\n'+cuerpo});
  }catch(e){ Logger.log('azcuFallasMail_: '+e); }
}
function azcuFallasSemanal(){
  try{
    var sf = SpreadsheetApp.openById(MAESTRO_ID).getSheetByName('Chat_Fallas');
    if(!sf || sf.getLastRow()<2) return;
    var desde = Date.now() - 7*24*3600000, l = [];
    sf.getRange(2, 1, sf.getLastRow()-1, Math.min(Math.max(sf.getLastColumn(),3),5)).getValues().forEach(function(r){ if(r[0] instanceof Date && r[0].getTime() >= desde) l.push({q:String(r[2]||'').slice(0, 70), u:String(r[1]||''), f:r[0], p:String(r[2]||''), a:String(r[3]||''), m:String(r[4]||'')}); });
    if(!l.length) return;
    azcuFallasMail_(l);
    var top = l.slice(0, 3).map(function(x){ return '«'+x.q+'»'+(x.u ? ' ('+x.u+')' : ''); }).join(' · ');
    var dest = []; azcuGerentes_().forEach(function(nm){ azcuDestinos_(nm, true).forEach(function(d){ if(dest.indexOf(d)<0) dest.push(d); }); });
    if(dest.length) azcuPush_('👎 '+l.length+(l.length===1 ? ' respuesta marcada' : ' respuestas marcadas')+' del chat esta semana', top+(l.length>3 ? ' y '+(l.length-3)+' más' : '')+'. Detalle en la pestaña Chat_Fallas: copiásela a Claude para que las arregle.', AZCU_WEB, dest, true, true);
  }catch(e){ Logger.log('azcuFallasSemanal: '+e); }
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
    if(hs.indexOf('azcuFallasSemanal')<0) ScriptApp.newTrigger('azcuFallasSemanal').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
  }catch(e){ Logger.log('azcuMantenimiento_: '+e); }
}
