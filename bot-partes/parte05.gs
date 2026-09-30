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
