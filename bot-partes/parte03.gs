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
