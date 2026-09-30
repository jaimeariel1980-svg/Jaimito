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
