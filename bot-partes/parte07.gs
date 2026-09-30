function texto2(chatId, t, est, paso){
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
