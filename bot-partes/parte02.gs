function testBot(){ Logger.log(tg('getMe').getContentText()); Logger.log(tg('getWebhookInfo').getContentText()); }


function crearTriggers(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(['chequeoCada5','alertasDiarias','alertasAutomaticas'].indexOf(t.getHandlerFunction())>=0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('chequeoCada5').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('alertasDiarias').timeBased().everyDays(1).atHour(9).create();
  Logger.log('Triggers creados');
}


function doGet(e){ return HtmlService.createHtmlOutput('Bot interno activo'); }


function doPost(e){
  var ok = HtmlService.createHtmlOutput('ok');
  try{
    var u = JSON.parse(e.postData.contents);
    if(u.update_id){
      var c = CacheService.getScriptCache();
      if(c.get('upd_'+u.update_id)) return ok;
      c.put('upd_'+u.update_id,'1',600);
    }
    if(u.callback_query){ manejarBoton(u.callback_query); return ok; }
    var msg = u.message;
    if(!msg) return ok;
    if(msg.date && (Date.now()/1000 - msg.date) > 120) return ok;
    var chatId = msg.chat.id, texto = msg.text || '';
    if(texto === '/start' || /^men[uú]$/i.test(texto)){ setEstado(chatId, null); mostrarMenu(chatId); return ok; }
    if(msg.voice || msg.audio){ manejarAudio(chatId, msg); return ok; }
    if(texto) manejarTexto(chatId, texto, msg.from);
  }catch(err){ try{ tgSend(CHAT_ID_ALERTAS, '⚠️ Error: '+err.message); }catch(e2){} }
  return ok;
}


function setEstado(chatId, obj){
  var c = CacheService.getScriptCache(), k = 'est_'+chatId;
  if(obj===null) c.remove(k); else c.put(k, JSON.stringify(obj), 3600);
}
