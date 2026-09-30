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
