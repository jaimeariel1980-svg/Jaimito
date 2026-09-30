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
  else boton2(chatId, a, est);
}
