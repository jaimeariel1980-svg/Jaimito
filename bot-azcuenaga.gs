var MAESTRO_ID = '1O8VKF9ZUC073kGNYIc3SgKPcIrVpnvynpQQ69RAk9W8';
var CHAT_ID_ALERTAS = '5692711315';
var TZ = 'America/Argentina/Buenos_Aires';
var MIN_ANTES = 30;
var DIAS_PAPELES = 10;
var DIAS_REPORTE = 20;
var HOJA = 'Ventas - Base de Datos Maestra';
var C = { COD:0, ACT:2, DUENO:3, TEL:9, MAIL:10, CALLE:43, NUM:44, DOCS:69, PRECIO:108, MONEDA:109, ETAPA:111, TOKKO:119, COMPL:67 };
var DOCS = ['Plano edificación','Plano mensura','Escritura','DNI propietario','CUIL propietario','Título propiedad','TGI Municipal','API Provincial','EPE Luz','Aguas Santafesinas','Litoral Gas'];

function guardarClaves(){
  var p = PropertiesService.getScriptProperties();
  p.setProperty('TG_TOKEN', 'PEGAR_TOKEN_NUEVO');
  p.setProperty('OPENAI_KEY', 'PEGAR_KEY_NUEVA');
  Logger.log('Claves guardadas OK');
}
function TG(){ return PropertiesService.getScriptProperties().getProperty('TG_TOKEN'); }
function OAI(){ return PropertiesService.getScriptProperties().getProperty('OPENAI_KEY'); }
function ss(){ return SpreadsheetApp.openById(MAESTRO_ID); }
function prop(k,v){ var p=PropertiesService.getScriptProperties(); if(v===undefined) return p.getProperty(k); p.setProperty(k,v); }
function fmt(d,f){ return Utilities.formatDate(d,TZ,f); }
function hoy(){ return fmt(new Date(),'dd/MM/yyyy'); }
function tg(m,q){ return UrlFetchApp.fetch('https://api.telegram.org/bot'+TG()+'/'+m+(q||''),{muteHttpExceptions:true}); }

function setWebhookManual(){
  var URL_EXEC = 'https://script.google.com/macros/s/AKfycbx4bNGJvySzo-hhQb5Kv092QdIb28CffgDGV3RMvun0lZ1U7QRKV0RITSMuPPGhHXaF/exec';
  Logger.log(tg('setWebhook','?url='+encodeURIComponent(URL_EXEC)).getContentText());
}
function pararTodo(){ Logger.log(tg('deleteWebhook','?drop_pending_updates=true').getContentText()); }
function testBot(){ Logger.log(tg('getMe').getContentText()); Logger.log(tg('getWebhookInfo').getContentText()); }

function crearTriggers(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    if(['chequeoCada5','alertasDiarias','resumenDia','alertasAutomaticas'].indexOf(t.getHandlerFunction())>=0) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('chequeoCada5').timeBased().everyMinutes(5).create();
  ScriptApp.newTrigger('resumenDia').timeBased().everyDays(1).atHour(8).create();
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
function getEstado(chatId){
  var v = CacheService.getScriptCache().get('est_'+chatId);
  return v ? JSON.parse(v) : null;
}

function mostrarMenu(chatId){
  tgSend(chatId, '🏠 *Bot Azcuénaga*\n\n📝 *Cargar*: registrá una consulta o visita\n📅 *Agendar*: visita con aviso y Calendar\n🔍 *Buscar*: ficha de una propiedad\n📌 *Estado*: cambiá la etapa\n📊 *Reporte*: resumen al propietario\n\n🎤 *Atajo:* mandame un audio o frase con lo que pasó (ej: _"Visité Gutemberg con Juan, le interesó"_) y lo cargo solo.', { inline_keyboard: [
    [{ text:'📝 Cargar visita / consulta', callback_data:'cargar' }],
    [{ text:'📅 Agendar visita',            callback_data:'agendar' }],
    [{ text:'🔍 Buscar propiedad',          callback_data:'buscar' }],
    [{ text:'📌 Cambiar estado',            callback_data:'cambiar_estado' }],
    [{ text:'📊 Reporte a propietario',     callback_data:'reporte' }]
  ]});
}

function listaProps(chatId, texto, prefijo){
  var res = buscarPropiedades(texto);
  if(!res.length){ tgSend(chatId, 'No encontré "'+texto+'". Probá otra parte de la dirección o el código.'); return; }
  var b = res.slice(0,8).map(function(p){ return [{ text:p.dir+' ('+p.codigo+')', callback_data:prefijo+p.codigo }]; });
  b.push([{ text:'✖ Cancelar', callback_data:'cancelar' }]);
  tgSend(chatId, 'Elegí una:', { inline_keyboard:b });
}

var INICIOS = {
  cargar:['buscando_prop','📝 *Cargar visita / consulta*\n\nEscribí parte de la *dirección* o el *código* (ej: "Mendoza" o "MIG-P004").'],
  agendar:['ag_buscando','📅 *Agendar visita*\n\nEscribí parte de la *dirección* o el *código* de la propiedad:'],
  buscar:['buscando_prop_info','🔍 *Buscar propiedad*\n\nEscribí parte de la dirección o el código:'],
  cambiar_estado:['buscando_prop_estado','📌 *Cambiar estado*\n\nEscribí la dirección o código de la propiedad:'],
  reporte:['buscando_prop_reporte','📊 *Reporte a propietario*\n\nEscribí la dirección o código de la propiedad:']
};

function manejarBoton(cq){
  var chatId = cq.message.chat.id, a = cq.data, vend = (cq.from && cq.from.first_name) || '';
  tg('answerCallbackQuery','?callback_query_id='+cq.id);
  var est = getEstado(chatId), p, cod;

  if(INICIOS[a]){ setEstado(chatId, {paso:INICIOS[a][0], vendedor:vend}); tgSend(chatId, INICIOS[a][1]); }
  else if(a.indexOf('prop:')===0){
    cod = a.substring(5); p = buscarPropPorCodigo(cod);
    setEstado(chatId, {paso:'pide_tel', codigo:cod, dir:(p?p.dir:cod), vendedor:vend});
    tgSend(chatId, '✅ Propiedad: *'+(p?p.dir:cod)+'* ('+cod+')\n\nEscribí el *teléfono* del interesado (solo números, ej: 3415702332).');
  }
  else if(a.indexOf('agp:')===0){
    cod = a.substring(4); p = buscarPropPorCodigo(cod);
    setEstado(chatId, {paso:'ag_frase', codigo:cod, dir:(p?p.dir:cod), vendedor:vend});
    tgSend(chatId, '✅ *'+(p?p.dir:cod)+'* ('+cod+')\n\n🎤 Mandame *un audio o una frase* con fecha, hora y cliente.\nEj: _"mañana 17:30 Juan 3415550000"_\n\nO tocá *Paso a paso*.', {inline_keyboard:[
      [{text:'🧭 Paso a paso', callback_data:'ag_pasos'}],
      [{text:'✖ Cancelar', callback_data:'cancelar'}]
    ]});
  }
  else if(a.indexOf('info:')===0){ tgSend(chatId, fichaPropiedad(a.substring(5))); setEstado(chatId, null); }
  else if(a.indexOf('estado_prop:')===0){
    cod = a.substring(12);
    var bts = ['Captación','Publicada','Reserva','Vendida','Suspendida'].map(function(e){ return [{text:e, callback_data:'set_estado:'+cod+':'+e}]; });
    bts.push([{text:'✖ Cancelar', callback_data:'cancelar'}]);
    tgSend(chatId, '¿Nueva etapa para *'+cod+'*?', {inline_keyboard:bts});
  }
  else if(a.indexOf('set_estado:')===0){
    var pt = a.substring(11).split(':');
    cambiarEstado(pt[0], pt[1]); setEstado(chatId, null);
    tgSend(chatId, '✅ *'+pt[0]+'* → '+pt[1]+' actualizado en el maestro.');
  }
  else if(a.indexOf('cand:')===0 && est && est.cands){ var c = est.cands[+a.substring(5)]; seguirConInteresado(chatId, est, c.nombre, c.tel); }
  else if(a==='ag_pasos' && est){ est.paso='ag_fecha'; setEstado(chatId, est); tgSend(chatId, '¿Qué *fecha*? (dd/mm/aaaa, dd/mm, "hoy" o "mañana")'); }
  else if(a==='cand_new' && est){ seguirConInteresado(chatId, est, est.nombreNuevo, ''); }
  else if(a.indexOf('rprop:')===0 && est){
    cod = a.substring(6); p = buscarPropPorCodigo(cod);
    est.codigo = cod; est.dir = p ? p.dir : cod; continuarRapida(chatId, est);
  }
  else boton2(chatId, a, est);
}

function boton2(chatId, a, est){
  var cod;
  if(a.indexOf('rep:')===0){
    cod = a.substring(4);
    var rep = generarReporte(cod);
    setEstado(chatId, {paso:'confirmar_reporte', codigo:cod, reporte:rep.texto, mail:rep.mail, dir:rep.dir});
    tgSend(chatId, rep.texto+'\n\n¿Lo enviamos?', {inline_keyboard:[
      [{text:'✅ Enviar por mail al propietario', callback_data:'enviar_rep'}],
      [{text:'✖ No enviar', callback_data:'cancelar'}]
    ]});
  }
  else if(a==='enviar_rep'){
    if(!est || !est.reporte){ tgSend(chatId,'Se venció la sesión. /start'); return; }
    if(est.mail){
      enviarReporteMail(est.codigo, est.reporte, est.mail);
      tgSend(chatId, '✅ Reporte enviado por mail a '+est.mail+'.'); setEstado(chatId, null);
    } else {
      est.paso = 'pide_mail'; setEstado(chatId, est);
      tgSend(chatId, '⚠️ El propietario no tiene mail cargado.\n\nEscribí el *mail* para enviarle el reporte (se guarda en el maestro):');
    }
  }
  else if(a.indexOf('via:')===0){
    if(!est){ tgSend(chatId,'Se venció la sesión. /start'); return; }
    est.via = a.substring(4); est.paso='pide_resultado'; setEstado(chatId, est);
    var r = ['Solo consulta','Visita','Le interesó','No le interesó','Hizo una propuesta'].map(function(x){ return [{ text:x, callback_data:'res:'+x }]; });
    r.push([{ text:'✖ Cancelar', callback_data:'cancelar' }]);
    tgSend(chatId, '¿Cuál fue el *resultado*?', { inline_keyboard:r });
  }
  else if(a.indexOf('res:')===0){
    if(!est){ tgSend(chatId,'Se venció la sesión. /start'); return; }
    est.resultado = a.substring(4); est.paso='esperando_audio'; setEstado(chatId, est);
    tgSend(chatId, '🎤 Mandame un *audio* contando cómo fue, o escribí el detalle.\n\nSi no tenés nada para agregar, escribí *listo*.');
  }
  else if(a==='conf:1'){
    if(est && est.listo){
      guardarSeguimiento(est);
      tgSend(chatId, '✅ *Guardado* en el seguimiento de *'+est.dir+'*.\n\n/start para el menú.');
      setEstado(chatId, null);
    } else tgSend(chatId, 'Se venció la sesión. /start');
  }
  else if(a==='cancelar'){ setEstado(chatId, null); tgSend(chatId, 'Cancelado. /start para el menú.'); }
}

function manejarTexto(chatId, texto, from){
  var est = getEstado(chatId), t = texto.trim();
  var paso = est && est.paso;
  var bus = { buscando_prop:'prop:', ag_buscando:'agp:', buscando_prop_info:'info:', buscando_prop_estado:'estado_prop:', buscando_prop_reporte:'rep:', rapida_prop_txt:'rprop:' };
  if(bus[paso]){ listaProps(chatId, t, bus[paso]); return; }

  if(paso==='pide_tel' || paso==='ag_tel'){
    if(/[a-zA-ZáéíóúñÁÉÍÓÚÑ]/.test(t)){ est.ag = paso==='ag_tel'; porNombre(chatId, est, t); return; }
    var tel = t.replace(/[^0-9]/g,'');
    if(tel.length < 6){ tgSend(chatId, 'Ese teléfono parece corto. Escribilo de nuevo (solo números).'); return; }
    est.telefono = tel;
    var previo = buscarPersonaPorTel(tel);
    var ag = paso==='ag_tel';
    if(previo){
      est.interesado = previo.nombre;
      tgSend(chatId, '👤 Es *'+previo.nombre+'* — '+previo.cant+' interacción/es registradas.');
      if(ag){ siguienteAg(chatId, est); }
      else { est.paso='pide_via'; setEstado(chatId, est); pedirVia(chatId); }
    } else {
      est.paso = ag ? 'ag_nombre' : 'pide_nombre'; setEstado(chatId, est);
      tgSend(chatId, 'Teléfono nuevo. ¿*Nombre* del interesado?');
    }
    return;
  }
  if(paso==='pide_nombre'){ est.interesado=t; est.paso='pide_via'; setEstado(chatId, est); pedirVia(chatId); return; }
  if(paso==='ag_nombre'){ est.interesado=t; siguienteAg(chatId, est); return; }

  texto2(chatId, t, est, paso, from);
}

function texto2(chatId, t, est, paso, from){
  if(paso==='ag_fecha'){
    var f = parseFechaInput(t);
    if(!f){ tgSend(chatId, 'No entendí la fecha. Probá: 15/10/2026, 15/10, hoy o mañana.'); return; }
    est.fecha = f; siguienteAg(chatId, est); return;
  }
  if(paso==='ag_hora'){
    var m = t.match(/^(\d{1,2})(?:[:.h]?(\d{2}))?$/);
    if(!m || +m[1]>23 || +(m[2]||0)>59){ tgSend(chatId, 'No entendí la hora. Probá: 15:30 o 15.'); return; }
    est.hora = ('0'+m[1]).slice(-2)+':'+(m[2]||'00'); siguienteAg(chatId, est); return;
  }
  if(paso==='ag_mail'){
    if(t!=='-' && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(t)){ tgSend(chatId, 'Ese mail no parece válido. Escribilo de nuevo o "-" para omitir.'); return; }
    est.email = t==='-' ? '' : t;
    if(est.nota!==undefined){ cerrarAgenda(chatId, est); return; }
    est.paso='ag_nota'; setEstado(chatId, est);
    tgSend(chatId, '¿Alguna *nota*? (o escribí "-")'); return;
  }
  if(paso==='ag_frase'){ agendaFrase(chatId, est, t); return; }
  if(paso==='ag_nota'){
    est.nota = t==='-' ? '' : t;
    cerrarAgenda(chatId, est); return;
  }
  texto3(chatId, t, est, paso, from);
}

function cerrarAgenda(chatId, est){
  var calOk = guardarAgenda(est); setEstado(chatId, null);
  tgSend(chatId, '✅ *Visita agendada*\n🏠 '+est.dir+' ('+est.codigo+')\n📅 '+est.fecha+' '+est.hora+'\n👤 '+(est.interesado||'—')+' · '+est.telefono+
      '\n'+(calOk?'🗓 Evento creado en Calendar'+(est.email?' e invitación enviada':'')+'.':'⚠️ No se pudo crear el evento en Calendar.')+'\nTe aviso '+MIN_ANTES+' min antes.\n\n/start para el menú.');
}

function texto3(chatId, t, est, paso, from){
  if(paso==='pide_mail'){
    if(!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(t)){ tgSend(chatId, 'Ese mail no parece válido. Escribilo de nuevo.'); return; }
    guardarMailDueno(est.codigo, t);
    enviarReporteMail(est.codigo, est.reporte, t);
    tgSend(chatId, '✅ Mail guardado en el maestro y reporte enviado a '+t+'.'); setEstado(chatId, null); return;
  }

  if(paso==='esperando_audio'){
    if(/^listo$/i.test(t)) est.observaciones = est.observaciones || '';
    else {
      est.observaciones = t;
      var d = interpretarVisita(t);
      if(d && d.proximo_paso) est.proximo_paso = d.proximo_paso;
    }
    finalizarCarga(chatId, est); return;
  }
  if(t.length>=25 && /\s/.test(t)){ cargaRapida(chatId, t, from && from.first_name); return; }
  tgSend(chatId, 'Para empezar tocá una opción 👇'); mostrarMenu(chatId);
}

function parseFechaInput(t){
  var d = new Date(), s = t.toLowerCase();
  if(s==='hoy') return fmt(d,'dd/MM/yyyy');
  if(/^ma[ñn]ana$/.test(s)) return fmt(new Date(Date.now()+86400000),'dd/MM/yyyy');
  var m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.](\d{2,4}))?$/);
  if(!m) return null;
  var y = m[3] ? (+m[3]<100 ? 2000+ +m[3] : +m[3]) : +fmt(d,'yyyy');
  var dd = +m[1], mm = +m[2];
  var chk = new Date(Date.UTC(y,mm-1,dd));
  if(chk.getUTCMonth()!==mm-1 || chk.getUTCDate()!==dd) return null;
  return ('0'+dd).slice(-2)+'/'+('0'+mm).slice(-2)+'/'+y;
}

function pedirVia(chatId){
  tgSend(chatId, '¿Por qué *vía* fue el contacto?', { inline_keyboard: [
    [{ text:'Visita', callback_data:'via:Visita' }],
    [{ text:'WhatsApp', callback_data:'via:WhatsApp' }],
    [{ text:'Llamada', callback_data:'via:Llamada' }],
    [{ text:'✖ Cancelar', callback_data:'cancelar' }]
  ]});
}

function manejarAudio(chatId, msg){
  var est = getEstado(chatId);
  tgSend(chatId, '🎧 Escuchando...');
  var texto = transcribirAudioTelegram((msg.voice && msg.voice.file_id) || (msg.audio && msg.audio.file_id));
  if(!texto){ tgSend(chatId, 'No pude entender el audio. Probá de nuevo o escribí el detalle.'); return; }
  if(est && est.paso==='ag_frase'){ agendaFrase(chatId, est, texto); return; }
  if(!est || est.paso!=='esperando_audio'){ cargaRapida(chatId, texto, msg.from && msg.from.first_name); return; }
  est.observaciones = texto;
  var d = interpretarVisita(texto);
  if(d && d.proximo_paso) est.proximo_paso = d.proximo_paso;
  if(d && d.observaciones) est.observaciones = d.observaciones;
  finalizarCarga(chatId, est);
}

function interpretarAgenda(texto){
  var hoyIso = fmt(new Date(),'yyyy-MM-dd'), dia = ['domingo','lunes','martes','miércoles','jueves','viernes','sábado'][new Date(Date.now()-3*3600000).getUTCDay()];
  var sys = 'Extraé de lo que dice un vendedor para agendar una visita. Hoy es '+dia+' '+hoyIso+'. Devolvé SOLO JSON: {fecha (AAAA-MM-DD resuelta a fecha real, o vacío), hora (HH:mm 24h o vacío), interesado (nombre o vacío), telefono (solo dígitos o vacío), email (o vacío), nota (o vacío)}. Español rioplatense.';
  try{
    var r = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions',{
      method:'post', contentType:'application/json', headers:{Authorization:'Bearer '+OAI()},
      payload:JSON.stringify({model:'gpt-4o-mini', temperature:0, response_format:{type:'json_object'},
        messages:[{role:'system',content:sys},{role:'user',content:texto}]}),
      muteHttpExceptions:true
    });
    return JSON.parse(JSON.parse(r.getContentText()).choices[0].message.content);
  }catch(e){ return null; }
}
function agendaFrase(chatId, est, texto){
  var d = interpretarAgenda(texto);
  if(!d){ tgSend(chatId, 'No pude interpretarlo. Probá de nuevo o tocá *Paso a paso*.'); return; }
  var m = String(d.fecha||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if(m) est.fecha = m[3]+'/'+m[2]+'/'+m[1];
  var h = String(d.hora||'').match(/^(\d{1,2}):(\d{2})/);
  if(h) est.hora = ('0'+h[1]).slice(-2)+':'+h[2];
  est.interesado = String(d.interesado||'').trim();
  est.telefono = String(d.telefono||'').replace(/[^0-9]/g,'');
  if(est.interesado && !est.telefono){
    var c = buscarPersonasPorNombre(est.interesado);
    if(c.length===1){ est.interesado = c[0].nombre; est.telefono = c[0].tel; }
  }
  if(d.email && /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email)) est.email = d.email;
  if(d.nota) est.nota = String(d.nota);
  siguienteAg(chatId, est);
}
function siguienteAg(chatId, est){
  if(!est.fecha){ est.paso='ag_fecha'; setEstado(chatId, est); tgSend(chatId, '¿Qué *fecha*? (dd/mm/aaaa, dd/mm, "hoy" o "mañana")'); return; }
  if(!est.hora){ est.paso='ag_hora'; setEstado(chatId, est); tgSend(chatId, '📅 '+est.fecha+'\n¿A qué *hora*? (ej: 15:30)'); return; }
  if(!est.interesado && !est.telefono){ est.paso='ag_tel'; setEstado(chatId, est); tgSend(chatId, '📅 '+est.fecha+' '+est.hora+'\n¿*Nombre o teléfono* del cliente?'); return; }
  if(est.email===undefined){ est.paso='ag_mail'; setEstado(chatId, est); tgSend(chatId, '📅 '+est.fecha+' '+est.hora+' · 👤 '+(est.interesado||est.telefono)+'\n¿*Email* del cliente? (opcional, o escribí "-")'); return; }
  if(est.nota===undefined){ est.paso='ag_nota'; setEstado(chatId, est); tgSend(chatId, '¿Alguna *nota*? (o escribí "-")'); return; }
  cerrarAgenda(chatId, est);
}

function sinAcentos(x){ return String(x||'').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'').trim(); }
function buscarPersonasPorNombre(q){
  var v = ss().getSheetByName('Seguimiento').getDataRange().getValues(), n = sinAcentos(q), vistos = {}, out = [];
  if(!n) return out;
  for(var i=v.length-1;i>=1 && out.length<6;i--){
    var nom = String(v[i][3]||'').trim();
    if(!nom || sinAcentos(nom).indexOf(n)<0) continue;
    var tel = String(v[i][4]||'').replace(/[^0-9]/g,''), k = normalizarTel(tel) || sinAcentos(nom);
    if(vistos[k]) continue;
    vistos[k] = 1; out.push({nombre:nom, tel:tel});
  }
  return out;
}
function porNombre(chatId, est, t){
  var c = buscarPersonasPorNombre(t);
  if(!c.length){ seguirConInteresado(chatId, est, t, ''); return; }
  est.cands = c; est.nombreNuevo = t; setEstado(chatId, est);
  var b = c.map(function(p,i){ return [{text:p.nombre+(p.tel?' · '+p.tel:''), callback_data:'cand:'+i}]; });
  b.push([{text:'➕ Nuevo: '+t, callback_data:'cand_new'}]);
  tgSend(chatId, '¿Es alguno de estos?', {inline_keyboard:b});
}
function seguirConInteresado(chatId, est, nombre, tel){
  est.interesado = nombre; est.telefono = tel;
  if(est.ag){ siguienteAg(chatId, est); }
  else { est.paso='pide_via'; setEstado(chatId, est); tgSend(chatId, '👤 *'+nombre+'*'+(tel?' · '+tel:'')); pedirVia(chatId); }
}

function interpretarCargaRapida(texto){
  var sys = 'Sos asistente de una inmobiliaria. De lo que cuenta un vendedor extraé JSON con: propiedad (calle/dirección o código mencionado), interesado (nombre), telefono (solo dígitos o vacío), via (Visita, WhatsApp o Llamada), resultado (uno de: Solo consulta, Visita, Le interesó, No le interesó, Hizo una propuesta), proximo_paso (breve), observaciones (1-2 frases). Campos que no se mencionan: cadena vacía. Devolvé SOLO JSON. Español rioplatense.';
  try{
    var r = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions',{
      method:'post', contentType:'application/json', headers:{Authorization:'Bearer '+OAI()},
      payload:JSON.stringify({model:'gpt-4o-mini', temperature:0.2, response_format:{type:'json_object'},
        messages:[{role:'system',content:sys},{role:'user',content:texto}]}),
      muteHttpExceptions:true
    });
    return JSON.parse(JSON.parse(r.getContentText()).choices[0].message.content);
  }catch(e){ return null; }
}
function cargaRapida(chatId, texto, vend){
  var d = interpretarCargaRapida(texto);
  if(!d){ tgSend(chatId, 'No pude interpretarlo. Probá de nuevo o usá /start.'); return; }
  var vias = ['Visita','WhatsApp','Llamada'], ress = ['Solo consulta','Visita','Le interesó','No le interesó','Hizo una propuesta'];
  var est = { vendedor:vend||'', interesado:String(d.interesado||'').trim(), telefono:String(d.telefono||'').replace(/[^0-9]/g,''),
    via:vias.indexOf(d.via)>=0 ? d.via : 'Visita', resultado:ress.indexOf(d.resultado)>=0 ? d.resultado : 'Solo consulta',
    proximo_paso:d.proximo_paso||'', observaciones:d.observaciones||texto };
  var res = d.propiedad ? buscarPropiedades(d.propiedad) : [];
  if(res.length===1){ est.codigo = res[0].codigo; est.dir = res[0].dir; continuarRapida(chatId, est); return; }
  est.paso = 'rapida_prop_txt'; setEstado(chatId, est);
  if(!res.length){ tgSend(chatId, 'No identifiqué la propiedad. Escribí parte de la *dirección* o el *código*:'); return; }
  var b = res.slice(0,8).map(function(p){ return [{ text:p.dir+' ('+p.codigo+')', callback_data:'rprop:'+p.codigo }]; });
  b.push([{ text:'✖ Cancelar', callback_data:'cancelar' }]);
  tgSend(chatId, '¿Cuál propiedad?', { inline_keyboard:b });
}
function continuarRapida(chatId, est){
  if(est.interesado && !est.telefono){
    var c = buscarPersonasPorNombre(est.interesado);
    if(c.length===1){ est.interesado = c[0].nombre; est.telefono = c[0].tel; }
  }
  finalizarCarga(chatId, est);
}

function finalizarCarga(chatId, est){
  est.listo = true; setEstado(chatId, est);
  tgSend(chatId, '📋 *Reviso antes de guardar:*\n\n'+
    '🏠 Propiedad: '+est.dir+' ('+est.codigo+')\n'+
    '👤 Interesado: '+(est.interesado||'—')+'\n'+
    '📱 Teléfono: '+(est.telefono||'—')+'\n'+
    '📌 Vía: '+(est.via||'—')+'\n'+
    '✅ Resultado: '+(est.resultado||'—')+'\n'+
    '➡️ Próximo paso: '+(est.proximo_paso||'—')+'\n'+
    '📝 Observaciones: '+(est.observaciones||'—')+'\n\n¿Guardo?', { inline_keyboard:[
    [{ text:'✅ Sí, guardar', callback_data:'conf:1' }],
    [{ text:'✖ Cancelar', callback_data:'cancelar' }]
  ]});
}

function dirDe(r){ return (String(r[C.CALLE]||'').trim()+' '+String(r[C.NUM]||'').trim()).trim(); }
function maestro(){ return ss().getSheetByName(HOJA).getDataRange().getValues(); }

function buscarPropiedades(q){
  var v = maestro(), out = []; q = String(q).toLowerCase().trim();
  for(var i=1;i<v.length;i++){
    var cod = String(v[i][C.COD]).trim(), dir = dirDe(v[i]);
    if(cod && (cod.toLowerCase().indexOf(q)>=0 || dir.toLowerCase().indexOf(q)>=0)) out.push({codigo:cod, dir:dir||'(sin dirección)'});
  }
  return out;
}
function buscarPropPorCodigo(cod){
  var v = maestro();
  for(var i=1;i<v.length;i++) if(String(v[i][C.COD]).trim()===cod) return {codigo:cod, dir:dirDe(v[i])||'(sin dirección)'};
  return null;
}
function filaPorCodigo(v, cod){
  for(var i=1;i<v.length;i++) if(String(v[i][C.COD]).trim()===cod) return v[i];
  return null;
}
function docsPendientes(f){
  var p = [];
  for(var d=0;d<DOCS.length;d++) if(String(f[C.DOCS+d]).trim()==='Pendiente') p.push(DOCS[d]);
  return p;
}

function fichaPropiedad(cod){
  var f = filaPorCodigo(maestro(), cod);
  if(!f) return 'No encontré la propiedad '+cod+'.';
  var precio = String(f[C.PRECIO]).trim(), moneda = String(f[C.MONEDA]).trim(), dueno = String(f[C.DUENO]).trim(), tel = String(f[C.TEL]).trim();
  var compl = String(f[C.COMPL]).trim(), tokko = String(f[C.TOKKO]).trim(), pend = docsPendientes(f);
  var seg = ss().getSheetByName('Seguimiento').getDataRange().getValues(), ints = [];
  for(var j=seg.length-1;j>=1 && ints.length<3;j--)
    if(String(seg[j][0]).trim()===cod) ints.push(fechaTxt(seg[j][1])+' · '+seg[j][3]+' · '+seg[j][5]);
  var t = '🏠 *'+dirDe(f)+'* ('+cod+')\n📌 Etapa: '+String(f[C.ETAPA]).trim()+'\n';
  if(precio) t += '💰 '+precio+(moneda?' '+moneda:'')+'\n';
  if(dueno) t += '👤 Dueño: '+dueno+(tel?' · '+tel:'')+'\n';
  if(compl) t += '📊 Completitud: '+compl+'\n';
  if(pend.length) t += '\n⚠️ *Docs pendientes ('+pend.length+'):*\n'+pend.map(function(d){ return '· '+d; }).join('\n')+'\n';
  if(ints.length) t += '\n💬 *Últimas interacciones:*\n'+ints.map(function(i){ return '· '+i; }).join('\n')+'\n';
  return t + (tokko ? '\n🔗 [Ver en Tokko]('+tokko+')' : '\n_(Sin link Tokko cargado)_');
}
function fechaTxt(x){ return x instanceof Date ? fmt(x,'dd/MM/yyyy') : String(x); }

function cambiarEstado(cod, etapa){
  var sh = ss().getSheetByName(HOJA), v = sh.getDataRange().getValues();
  for(var i=1;i<v.length;i++) if(String(v[i][C.COD]).trim()===cod){
    sh.getRange(i+1,C.ETAPA+1).setValue(etapa);
    sh.getRange(i+1,C.ACT+1).setValue(hoy());
    return;
  }
}
function guardarMailDueno(cod, mail){
  var sh = ss().getSheetByName(HOJA), v = sh.getDataRange().getValues();
  for(var i=1;i<v.length;i++) if(String(v[i][C.COD]).trim()===cod){ sh.getRange(i+1,C.MAIL+1).setValue(mail); return; }
}

function parseFH(f, h){
  var fs = f instanceof Date ? fmt(f,'dd/MM/yyyy') : String(f).trim();
  var hs = h instanceof Date ? fmt(h,'HH:mm') : String(h).trim();
  var m = fs.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/), t = hs.match(/^(\d{1,2}):(\d{2})/);
  if(!m || !t) return null;
  return new Date(Date.UTC(+m[3], +m[2]-1, +m[1], +t[1]+3, +t[2]));
}
function guardarAgenda(est){
  var sh = ss().getSheetByName('Agenda'), r = sh.getLastRow()+1, rng = sh.getRange(r,1,1,9);
  rng.setNumberFormat('@');
  rng.setValues([[est.codigo, est.fecha, est.hora, est.interesado||'', est.telefono||'', est.vendedor||'', est.nota||'', 'Pendiente', est.email||'']]);
  try{
    ss().getSheetByName('Seguimiento').appendRow([est.codigo, est.fecha, 'Visita', est.interesado||'', est.telefono||'', 'Visita agendada', 'Visita '+est.fecha+' '+est.hora, '', est.nota||'', est.vendedor||'']);
  }catch(e){ Logger.log(e); }
  try{
    var ini = parseFH(est.fecha, est.hora);
    var o = {description:'Propiedad: '+est.codigo+'\nInteresado: '+(est.interesado||'')+'\nTel: '+(est.telefono||'')+(est.nota?'\n'+est.nota:''), location:est.dir||''};
    if(est.email){ o.guests = est.email; o.sendInvites = true; }
    var ev = CalendarApp.getDefaultCalendar().createEvent('Visita: '+(est.dir||est.codigo)+(est.interesado?' - '+est.interesado:''), ini, new Date(ini.getTime()+3600000), o);
    ev.addPopupReminder(60);
    return true;
  }catch(e){ Logger.log(e); return false; }
}

function chequeoCada5(){
  try{ avisarVisitas(); }catch(e){ Logger.log(e); }
  try{ avisarCambiosEstado(); }catch(e){ Logger.log(e); }
  try{ avisarEncuestas(); }catch(e){ Logger.log(e); }
}

function avisarVisitas(){
  var v = ss().getSheetByName('Seguimiento').getDataRange().getValues(), now = Date.now(), maestroV = null;
  var raw = prop('VISITAS_AVISADAS'), hechas = raw ? JSON.parse(raw) : [], nuevo = false;
  for(var i=1;i<v.length;i++){
    if(String(v[i][5]).trim()!=='Visita agendada') continue;
    var m = String(v[i][6]).match(/^Visita (\d{2}\/\d{2}\/\d{4}) (\d{2}:\d{2})/);
    if(!m) continue;
    var dt = parseFH(m[1], m[2]);
    if(!dt) continue;
    var diff = (dt.getTime()-now)/60000;
    if(diff > MIN_ANTES || diff <= -30) continue;
    var cod = String(v[i][0]).trim(), key = cod+'|'+m[1]+' '+m[2]+'|'+String(v[i][3]).trim();
    if(hechas.indexOf(key)>=0) continue;
    maestroV = maestroV || maestro();
    var f = filaPorCodigo(maestroV, cod);
    tgSend(CHAT_ID_ALERTAS, '⏰ *Visita en '+Math.max(0,Math.round(diff))+' min* ('+m[2]+')\n🏠 '+(f?dirDe(f):cod)+' ('+cod+')\n👤 '+v[i][3]+(v[i][4]?' · '+v[i][4]:'')+
      (v[i][9]?'\n🧑‍💼 '+v[i][9]:'')+(v[i][8]?'\n📝 '+v[i][8]:''));
    hechas.push(key); nuevo = true;
  }
  if(nuevo) prop('VISITAS_AVISADAS', JSON.stringify(hechas.slice(-200)));
}

function avisarCambiosEstado(){
  var v = maestro(), actual = {}, cambios = [];
  for(var i=1;i<v.length;i++){
    var cod = String(v[i][C.COD]).trim();
    if(cod) actual[cod] = String(v[i][C.ETAPA]).trim();
  }
  var raw = prop('SNAP_ESTADOS'), desde = jprop('ESTADO_DESDE'), ahora = Date.now();
  if(raw){
    var previo = JSON.parse(raw);
    for(var k in actual){
      if(previo[k]!==actual[k]) desde[k] = ahora;
      if(previo[k]!==undefined && previo[k]!==actual[k]){
        var f = filaPorCodigo(v, k);
        cambios.push('· '+k+' — '+(f?dirDe(f):'')+': '+(previo[k]||'—')+' → '+(actual[k]||'—'));
      }
    }
  }
  for(var k2 in actual) if(!desde[k2]) desde[k2] = ahora;
  if(cambios.length) tgSend(CHAT_ID_ALERTAS, '📌 *Cambio de estado*\n'+cambios.join('\n'));
  prop('SNAP_ESTADOS', JSON.stringify(actual));
  prop('ESTADO_DESDE', JSON.stringify(desde));
}

function jprop(k){ var r = prop(k); return r ? JSON.parse(r) : {}; }
function numFecha(x){
  if(x instanceof Date) return +fmt(x,'yyyyMMdd');
  var m = String(x).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  return m ? +(m[3]+('0'+m[2]).slice(-2)+('0'+m[1]).slice(-2)) : null;
}
function difDias(a, b){
  function d(n){ return Date.UTC(Math.floor(n/10000), Math.floor(n/100)%100-1, n%100); }
  return Math.round((d(a)-d(b))/86400000);
}
function hoyNum(){ return +fmt(new Date(),'yyyyMMdd'); }
function activas(v){
  var out = [];
  for(var i=1;i<v.length;i++){
    var et = String(v[i][C.ETAPA]).trim();
    if(v[i][C.COD] && (et==='Publicada' || et==='Reserva')) out.push({cod:String(v[i][C.COD]).trim(), dir:dirDe(v[i]), etapa:et});
  }
  return out;
}
function marcarReporte(cod){ var m = jprop('REP_ULT'); m[cod] = Date.now(); prop('REP_ULT', JSON.stringify(m)); }

function alertasDiarias(){
  var p = prop('ULT_PAPELES'), dias = p ? (Date.now()-Number(p))/86400000 : 999;
  if(dias >= DIAS_PAPELES - 0.1){
    var txt = papelesIncompletos();
    if(txt) tgSend(CHAT_ID_ALERTAS, txt);
    prop('ULT_PAPELES', String(Date.now()));
  }
  var mes = fmt(new Date(),'yyyyMM');
  if(+fmt(new Date(),'d')===5 && prop('ULT_REPORTE_MES')!==mes){
    var rm = reportesPendientes(true);
    if(rm) tgSend(CHAT_ID_ALERTAS, '📅 *Hoy es 5:* reportes del mes.\n\n'+rm);
    prop('ULT_REPORTE_MES', mes);
  }
  try{ var r = reportesPendientes(); if(r) tgSend(CHAT_ID_ALERTAS, r); }catch(e){ Logger.log(e); }
  try{ var q = propuestasSinRespuesta(); if(q) tgSend(CHAT_ID_ALERTAS, q); }catch(e){ Logger.log(e); }
  if(fmt(new Date(),'u')==='1'){
    try{ var w = semanal(); if(w) tgSend(CHAT_ID_ALERTAS, w); }catch(e){ Logger.log(e); }
  }
}

function papelesIncompletos(){
  var v = maestro(), out = [];
  for(var i=1;i<v.length;i++){
    var et = String(v[i][C.ETAPA]).trim();
    if(!v[i][C.COD] || !et || et==='Vendida' || et==='Suspendida') continue;
    var pend = docsPendientes(v[i]);
    if(pend.length>=3) out.push('· *'+String(v[i][C.COD]).trim()+'* '+dirDe(v[i])+'\n   Falta: '+pend.slice(0,4).join(', ')+(pend.length>4?' +'+(pend.length-4)+' más':''));
  }
  return out.length ? '⚠️ *Papeles incompletos ('+out.length+') — '+hoy()+'*\n'+out.join('\n') : '';
}

function reportesPendientes(mensual){
  var ult = prop('ULT_AVISO_REP');
  var map = jprop('REP_ULT'), v = maestro(), out = [], ahora = Date.now(), cambio = false;
  var seg = ss().getSheetByName('Seguimiento').getDataRange().getValues(), fechas = {};
  for(var i=1;i<seg.length;i++){
    var c = String(seg[i][0]).trim(), n = numFecha(seg[i][1]);
    if(c && n) (fechas[c] = fechas[c] || []).push(n);
  }
  activas(v).forEach(function(a){
    if(!map[a.cod]){ map[a.cod] = ahora; cambio = true; return; }
    var d = Math.floor((ahora-map[a.cod])/86400000);
    if(!mensual && d<DIAS_REPORTE) return;
    var desde = numFecha(new Date(map[a.cod]));
    var mov = (fechas[a.cod]||[]).filter(function(n){ return n>desde; }).length;
    if(mov) out.push('· *'+a.cod+'* '+a.dir+' — '+mov+' movimiento(s) desde el último reporte (hace '+d+' días)');
  });
  if(cambio) prop('REP_ULT', JSON.stringify(map));
  if(!out.length) return '';
  if(!mensual){ if(ult && (ahora-Number(ult))/86400000 < 3) return ''; prop('ULT_AVISO_REP', String(ahora)); }
  return '📊 *Reportes para enviar ('+out.length+')* — propiedades con movimiento\n'+out.join('\n')+'\n\nUsá "Reporte a propietario".';
}

function ultimasPorPersona(){
  var v = ss().getSheetByName('Seguimiento').getDataRange().getValues(), m = {};
  for(var i=1;i<v.length;i++){
    var cod = String(v[i][0]).trim(); if(!cod) continue;
    m[cod+'|'+(normalizarTel(v[i][4])||sinAcentos(v[i][3]))] = v[i];
  }
  return {filas:v, ult:m};
}

function propuestasSinRespuesta(){
  var d = ultimasPorPersona(), mae = maestro(), ult = {}, out = [], hn = hoyNum();
  for(var k in d.ult) ult[String(d.ult[k][0]).trim()] = d.ult[k];
  for(var cod in ult){
    var r = ult[cod], res = String(r[5]);
    if(!/propuesta/i.test(res) || /acept|rechaz/i.test(res)) continue;
    var f = filaPorCodigo(mae, cod), et = f ? String(f[C.ETAPA]).trim() : '';
    if(et==='Reserva' || et==='Vendida' || et==='Suspendida') continue;
    var n = numFecha(r[1]); if(!n) continue;
    var dd = difDias(hn, n);
    if(dd>=3 && dd<=30 && dd%3===0) out.push('· *'+cod+'* '+(f?dirDe(f):'')+' — '+r[3]+' (hace '+dd+' días)');
  }
  return out.length ? '💬 *Propuestas sin respuesta ('+out.length+')*\n'+out.join('\n') : '';
}

function semanal(){
  var mae = maestro(), d = ultimasPorPersona(), hn = hoyNum(), ultimo = {}, sinMov = [], est = [];
  for(var i=1;i<d.filas.length;i++){
    var cod = String(d.filas[i][0]).trim(), n = numFecha(d.filas[i][1]);
    if(cod && n && (!ultimo[cod] || n>ultimo[cod])) ultimo[cod] = n;
  }
  var desde = jprop('ESTADO_DESDE');
  activas(mae).forEach(function(a){
    var n = ultimo[a.cod], dd = n ? difDias(hn, n) : null;
    if(dd===null || dd>30) sinMov.push('· *'+a.cod+'* '+a.dir+' — '+(dd===null?'sin actividad':'hace '+dd+' días'));
    if(a.etapa==='Reserva' && desde[a.cod]){
      var r = Math.floor((Date.now()-desde[a.cod])/86400000);
      if(r>30) est.push('· *'+a.cod+'* '+a.dir+' — en Reserva hace '+r+' días');
    }
  });
  var t = '';
  if(sinMov.length) t += '🛑 *Sin movimiento +30 días ('+sinMov.length+')*\n'+sinMov.join('\n')+'\n\n';
  if(est.length) t += '⏳ *Reservas estancadas ('+est.length+')*\n'+est.join('\n');
  return t.trim();
}

function resumenDia(){
  var d = ultimasPorPersona(), mae = maestro(), hn = hoyNum(), visitas = [], toca = [];
  for(var i=1;i<d.filas.length;i++){
    var r = d.filas[i];
    if(String(r[5]).trim()!=='Visita agendada') continue;
    var m = String(r[6]).match(/^Visita (\d{2}\/\d{2}\/\d{4}) (\d{2}:\d{2})/);
    if(m && numFecha(m[1])===hn){
      var f = filaPorCodigo(mae, String(r[0]).trim());
      visitas.push({h:m[2], t:'· '+m[2]+' — '+(f?dirDe(f):r[0])+' · '+r[3]+(r[4]?' ('+r[4]+')':'')});
    }
  }
  visitas.sort(function(a,b){ return a.h<b.h?-1:1; });
  for(var k in d.ult){
    var x = d.ult[k], n = numFecha(x[7]);
    if(!n || /cerrar/i.test(String(x[6])) || /no le interes/i.test(String(x[5]))) continue;
    var dd = difDias(hn, n);
    if(dd>=0 && dd<=7){
      var f2 = filaPorCodigo(mae, String(x[0]).trim());
      toca.push('· '+x[3]+(x[4]?' ('+x[4]+')':'')+' — '+(f2?dirDe(f2):x[0])+(x[6]?' · '+x[6]:'')+(dd?' (vencido hace '+dd+'d)':''));
    }
  }
  var t = '';
  if(visitas.length) t += '🗓 *Visitas de hoy ('+visitas.length+')*\n'+visitas.map(function(x){ return x.t; }).join('\n')+'\n\n';
  if(toca.length) t += '📞 *Te toca contactar ('+toca.length+')*\n'+toca.join('\n');
  if(t) tgSend(CHAT_ID_ALERTAS, '☀️ *Buen día — '+hoy()+'*\n\n'+t.trim());
}

function avisarEncuestas(){
  var sh = ss().getSheetByName('Encuestas'); if(!sh) return;
  var n = sh.getLastRow(), ult = prop('ENC_N');
  if(ult===null){ prop('ENC_N', String(n)); return; }
  if(n<=+ult){ if(n<+ult) prop('ENC_N', String(n)); return; }
  var v = sh.getRange(+ult+1, 1, n-+ult, 10).getValues(), mae = maestro();
  v.forEach(function(r){
    var f = filaPorCodigo(mae, String(r[0]).trim());
    tgSend(CHAT_ID_ALERTAS, '⭐ *Nueva encuesta* — '+(f?dirDe(f):r[0])+' ('+r[0]+')\n👤 '+r[2]+' · '+r[3]+'\nNota: *'+r[4]+'/10* · Recomienda: '+(r[8]||'—')+(r[9]?'\n💬 '+r[9]:''));
  });
  prop('ENC_N', String(n));
}

function generarReporte(cod){
  var f = filaPorCodigo(maestro(), cod);
  if(!f) return {texto:'No encontré la propiedad.', mail:'', dir:''};
  var dir = dirDe(f), precio = String(f[C.PRECIO]).trim(), moneda = String(f[C.MONEDA]).trim();
  var seg = ss().getSheetByName('Seguimiento').getDataRange().getValues();
  var n = {consultas:0, visitas:0, propuestas:0}, ultObs = '';
  for(var j=1;j<seg.length;j++){
    if(String(seg[j][0]).trim()!==cod) continue;
    var res = String(seg[j][5]).trim();
    if(res==='Solo consulta') n.consultas++;
    else if(res==='Hizo una propuesta') n.propuestas++;
    else if(/visit|interes/i.test(res) && !/no le interes/i.test(res)) n.visitas++;
    if(seg[j][8]) ultObs = String(seg[j][8]).trim();
  }
  var t = '📊 *Reporte de '+dir+'*\nPeríodo: últimos 30 días al '+hoy()+'\n\n👤 Propietario: '+String(f[C.DUENO]).trim()+'\n📌 Estado: '+String(f[C.ETAPA]).trim()+'\n';
  if(precio) t += '💰 '+precio+(moneda?' '+moneda:'')+'\n';
  t += '\n📈 *Actividad:*\n· Consultas: '+n.consultas+'\n· Visitas / interesados: '+n.visitas+'\n· Propuestas: '+n.propuestas+'\n';
  if(ultObs) t += '\n📝 Última novedad: '+ultObs;
  return {texto:t, mail:String(f[C.MAIL]).trim(), dir:dir};
}
function enviarReporteMail(cod, texto, mail){
  var cuerpo = texto.replace(/\*/g,'').replace(/_/g,'').replace(/\[([^\]]+)\]\([^)]+\)/g,'$1')+'\n\nCualquier consulta estamos a tu disposición.\nAzcuénaga Inmobiliaria';
  GmailApp.sendEmail(mail, 'Reporte de tu propiedad — Azcuénaga Inmobiliaria', cuerpo);
  marcarReporte(cod);
}

function normalizarTel(t){
  var s = String(t).replace(/[^0-9]/g,'');
  return s.replace(/^0/,'').replace(/^(\d{2,4})15/,'$1').replace(/^549/,'').replace(/^54/,'');
}
function buscarPersonaPorTel(tel){
  var v = ss().getSheetByName('Seguimiento').getDataRange().getValues(), o = normalizarTel(tel), nombre = '', cant = 0;
  for(var i=1;i<v.length;i++){
    var t = normalizarTel(v[i][4]);
    if(t && t===o){ cant++; if(v[i][3]) nombre = v[i][3]; }
  }
  return cant>0 ? {nombre:nombre||'(sin nombre)', cant:cant} : null;
}

function guardarSeguimiento(est){
  ss().getSheetByName('Seguimiento').appendRow([
    est.codigo||'', hoy(), est.via||'Visita', est.interesado||'', est.telefono||'', est.resultado||'',
    est.proximo_paso||'', '', est.observaciones||'', est.vendedor||''
  ]);
}

function interpretarVisita(texto){
  var sys = 'Sos asistente de una inmobiliaria. Extraé del audio de un vendedor: proximo_paso (qué hacer después, breve), observaciones (resumen de lo que contó, 1-2 frases). Devolvé SOLO JSON: {proximo_paso, observaciones}. Español rioplatense.';
  try{
    var r = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions',{
      method:'post', contentType:'application/json', headers:{Authorization:'Bearer '+OAI()},
      payload:JSON.stringify({model:'gpt-4o-mini', temperature:0.3, response_format:{type:'json_object'},
        messages:[{role:'system',content:sys},{role:'user',content:texto}]}),
      muteHttpExceptions:true
    });
    return JSON.parse(JSON.parse(r.getContentText()).choices[0].message.content);
  }catch(e){ return null; }
}

function transcribirAudioTelegram(fileId){
  try{
    var path = JSON.parse(tg('getFile','?file_id='+fileId).getContentText()).result.file_path;
    var blob = UrlFetchApp.fetch('https://api.telegram.org/file/bot'+TG()+'/'+path,{muteHttpExceptions:true}).getBlob();
    var b = '----tg'+Date.now();
    var payload = Utilities.newBlob(
      '--'+b+'\r\nContent-Disposition: form-data; name="model"\r\n\r\nwhisper-1\r\n'+
      '--'+b+'\r\nContent-Disposition: form-data; name="language"\r\n\r\nes\r\n'+
      '--'+b+'\r\nContent-Disposition: form-data; name="file"; filename="audio.ogg"\r\nContent-Type: audio/ogg\r\n\r\n'
    ).getBytes().concat(blob.getBytes()).concat(Utilities.newBlob('\r\n--'+b+'--\r\n').getBytes());
    var r = UrlFetchApp.fetch('https://api.openai.com/v1/audio/transcriptions',{
      method:'post', contentType:'multipart/form-data; boundary='+b,
      headers:{Authorization:'Bearer '+OAI()}, payload:payload, muteHttpExceptions:true
    });
    return JSON.parse(r.getContentText()).text || '';
  }catch(err){ return ''; }
}

function tgSend(chatId, texto, teclado){
  var p = {chat_id:chatId, text:texto, parse_mode:'Markdown'};
  if(teclado) p.reply_markup = JSON.stringify(teclado);
  var o = {method:'post', contentType:'application/json', muteHttpExceptions:true};
  o.payload = JSON.stringify(p);
  var r = UrlFetchApp.fetch('https://api.telegram.org/bot'+TG()+'/sendMessage', o);
  if(r.getResponseCode()!==200){ delete p.parse_mode; o.payload = JSON.stringify(p); UrlFetchApp.fetch('https://api.telegram.org/bot'+TG()+'/sendMessage', o); }
}

function testAgenda(){
  guardarAgenda({codigo:'TEST', fecha:'01/01/2030', hora:'10:00', interesado:'Prueba', telefono:'3410000000', vendedor:'Test', nota:'borrar'});
  Logger.log('Fila agregada en Agenda OK');
}

function autorizar(){
  GmailApp.getInboxUnreadCount();
  CalendarApp.getDefaultCalendar().getName();
  ss().getName();
  Logger.log('Permisos OK');
}
