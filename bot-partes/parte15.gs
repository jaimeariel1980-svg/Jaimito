function tgSend(chatId, texto, teclado){
  var p = {chat_id:chatId, text:texto, parse_mode:'Markdown'};
  if(teclado) p.reply_markup = JSON.stringify(teclado);
  var o = {method:'post', contentType:'application/json', muteHttpExceptions:true};
  o.payload = JSON.stringify(p);
  var r = UrlFetchApp.fetch('https://api.telegram.org/bot'+TG()+'/sendMessage', o);
  if(r.getResponseCode()!==200){ delete p.parse_mode; o.payload = JSON.stringify(p); UrlFetchApp.fetch('https://api.telegram.org/bot'+TG()+'/sendMessage', o); }
}
