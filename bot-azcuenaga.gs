var MAESTRO_ID = '1O8VKF9ZUC073kGNYIc3SgKPcIrVpnvynpQQ69RAk9W8';
var CHAT_ID_ALERTAS = '5692711315';
var TZ = 'America/Argentina/Buenos_Aires';
var MIN_ANTES = 30;
var DIAS_PAPELES = 10;
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
  }catch(err){}
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
  tgSend(chatId, '🏠 *Bot Azcuénaga* — ¿Qué necesitás?\n\nElegí una opción:', { inline_keyboard: [
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
    setEstado(chatId, {paso:'ag_fecha', codigo:cod, dir:(p?p.dir:cod), vendedor:vend});
    tgSend(chatId, '✅ *'+(p?p.dir:cod)+'* ('+cod+')\n\n¿Qué *fecha*? (dd/mm/aaaa, dd/mm, "hoy" o "mañana")');
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
  else if(a.indexOf('rep:')===0){
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
  var bus = { buscando_prop:'prop:', ag_buscando:'agp:', buscando_prop_info:'info:', buscando_prop_estado:'estado_prop:', buscando_prop_reporte:'rep:' };
  if(bus[paso]){ listaProps(chatId, t, bus[paso]); return; }

  if(paso==='pide_tel' || paso==='ag_tel'){
    var tel = t.replace(/[^0-9]/g,'');
    if(tel.length < 6){ tgSend(chatId, 'Ese teléfono parece corto. Escribilo de nuevo (solo números).'); return; }
    est.telefono = tel;
    var previo = buscarPersonaPorTel(tel);
    var ag = paso==='ag_tel';
    if(previo){
      est.interesado = previo.nombre;
      tgSend(chatId, '👤 Es *'+previo.nombre+'* — '+previo.cant+' interacción/es registradas.');
      if(ag){ est.paso='ag_nota'; setEstado(chatId, est); tgSend(chatId, '¿Alguna *nota*? (o escribí "-")'); }
      else { est.paso='pide_via'; setEstado(chatId, est); pedirVia(chatId); }
    } else {
      est.paso = ag ? 'ag_nombre' : 'pide_nombre'; setEstado(chatId, est);
      tgSend(chatId, 'Teléfono nuevo. ¿*Nombre* del interesado?');
    }
    return;
  }
  if(paso==='pide_nombre'){ est.interesado=t; est.paso='pide_via'; setEstado(chatId, est); pedirVia(chatId); return; }
  if(paso==='ag_nombre'){ est.interesado=t; est.paso='ag_nota'; setEstado(chatId, est); tgSend(chatId, '¿Alguna *nota*? (o escribí "-")'); return; }

  if(paso==='ag_fecha'){
    var f = parseFechaInput(t);
    if(!f){ tgSend(chatId, 'No entendí la fecha. Probá: 15/10/2026, 15/10, hoy o mañana.'); return; }
    est.fecha = f; est.paso='ag_hora'; setEstado(chatId, est);
    tgSend(chatId, '¿A qué *hora*? (ej: 15:30)'); return;
  }
  if(paso==='ag_hora'){
    var m = t.match(/^(\d{1,2})(?:[:.h]?(\d{2}))?$/);
    if(!m || +m[1]>23 || +(m[2]||0)>59){ tgSend(chatId, 'No entendí la hora. Probá: 15:30 o 15.'); return; }
    est.hora = ('0'+m[1]).slice(-2)+':'+(m[2]||'00'); est.paso='ag_tel'; setEstado(chatId, est);
    tgSend(chatId, 'Escribí el *teléfono* del interesado (solo números):'); return;
  }
  if(paso==='ag_nota'){
    est.nota = t==='-' ? '' : t;
    guardarAgenda(est); setEstado(chatId, null);
    tgSend(chatId, '✅ *Visita agendada*\n🏠 '+est.dir+' ('+est.codigo+')\n📅 '+est.fecha+' '+est.hora+'\n👤 '+(est.interesado||'—')+' · '+est.telefono+
      '\n\nTe aviso '+MIN_ANTES+' min antes.'); return;
  }

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
  if(!est || est.paso!=='esperando_audio'){ tgSend(chatId, 'Primero elegí propiedad, teléfono, vía y resultado. /start → Cargar visita.'); return; }
  tgSend(chatId, '🎧 Escuchando...');
  var texto = transcribirAudioTelegram((msg.voice && msg.voice.file_id) || (msg.audio && msg.audio.file_id));
  if(!texto){ tgSend(chatId, 'No pude entender el audio. Probá de nuevo o escribí el detalle.'); return; }
  est.observaciones = texto;
  var d = interpretarVisita(texto);
  if(d && d.proximo_paso) est.proximo_paso = d.proximo_paso;
  if(d && d.observaciones) est.observaciones = d.observaciones;
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
  var sh = ss().getSheetByName('Agenda'), r = sh.getLastRow()+1, rng = sh.getRange(r,1,1,8);
  rng.setNumberFormat('@');
  rng.setValues([[est.codigo, est.fecha, est.hora, est.interesado||'', est.telefono||'', est.vendedor||'', est.nota||'', 'Pendiente']]);
}

function chequeoCada5(){
  try{ avisarVisitas(); }catch(e){ Logger.log(e); }
  try{ avisarCambiosEstado(); }catch(e){ Logger.log(e); }
}

function avisarVisitas(){
  var sh = ss().getSheetByName('Agenda'), v = sh.getDataRange().getValues(), now = Date.now(), maestroV = null;
  for(var i=1;i<v.length;i++){
    var e = String(v[i][7]).trim();
    if(e && e!=='Pendiente') continue;
    var dt = parseFH(v[i][1], v[i][2]);
    if(!dt) continue;
    var diff = (dt.getTime()-now)/60000;
    if(diff > MIN_ANTES) continue;
    if(diff > -30){
      maestroV = maestroV || maestro();
      var cod = String(v[i][0]).trim(), f = filaPorCodigo(maestroV, cod);
      tgSend(CHAT_ID_ALERTAS, '⏰ *Visita en '+Math.max(0,Math.round(diff))+' min* ('+fmt(dt,'HH:mm')+')\n🏠 '+(f?dirDe(f):cod)+' ('+cod+')\n👤 '+v[i][3]+(v[i][4]?' · '+v[i][4]:'')+
        (v[i][5]?'\n🧑‍💼 '+v[i][5]:'')+(v[i][6]?'\n📝 '+v[i][6]:''));
      sh.getRange(i+1,8).setValue('Avisada');
    } else sh.getRange(i+1,8).setValue('Vencida');
  }
}

function avisarCambiosEstado(){
  var v = maestro(), actual = {}, cambios = [];
  for(var i=1;i<v.length;i++){
    var cod = String(v[i][C.COD]).trim();
    if(cod) actual[cod] = String(v[i][C.ETAPA]).trim();
  }
  var raw = prop('SNAP_ESTADOS');
  if(raw){
    var previo = JSON.parse(raw);
    for(var k in actual){
      if(previo[k]!==undefined && previo[k]!==actual[k]){
        var f = filaPorCodigo(v, k);
        cambios.push('· '+k+' — '+(f?dirDe(f):'')+': '+(previo[k]||'—')+' → '+(actual[k]||'—'));
      }
    }
  }
  if(cambios.length) tgSend(CHAT_ID_ALERTAS, '📌 *Cambio de estado*\n'+cambios.join('\n'));
  prop('SNAP_ESTADOS', JSON.stringify(actual));
}

function alertasDiarias(){
  var h = hoy(), p = prop('ULT_PAPELES');
  var dias = p ? (Date.now()-Number(p))/86400000 : 999;
  if(dias >= DIAS_PAPELES - 0.1){
    var txt = papelesIncompletos();
    if(txt) tgSend(CHAT_ID_ALERTAS, txt);
    prop('ULT_PAPELES', String(Date.now()));
  }
  var mes = fmt(new Date(),'yyyyMM');
  if(+fmt(new Date(),'d')===5 && prop('ULT_REPORTE_MES')!==mes){
    tgSend(CHAT_ID_ALERTAS, '📊 *Recordatorio:* hoy 5 — corresponde enviar los reportes a propietarios. Usá "Reporte a propietario".');
    prop('ULT_REPORTE_MES', mes);
  }
}
function papelesIncompletos(){
  var v = maestro(), out = [];
  for(var i=1;i<v.length;i++){
    var et = String(v[i][C.ETAPA]).trim();
    if(!v[i][C.COD] || !et || et==='Vendida' || et==='Suspendida') continue;
    var n = docsPendientes(v[i]).length;
    if(n>=3) out.push('· '+String(v[i][C.COD]).trim()+' — '+dirDe(v[i])+' ('+n+' pend.)');
  }
  return out.length ? '⚠️ *Docs incompletos ('+out.length+') — '+hoy()+'*\n'+out.join('\n') : '';
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
