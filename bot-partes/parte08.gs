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
