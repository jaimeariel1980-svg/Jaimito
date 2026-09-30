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
