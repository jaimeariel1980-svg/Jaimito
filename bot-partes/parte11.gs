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
