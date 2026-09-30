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

  texto2(chatId, t, est, paso);
}
