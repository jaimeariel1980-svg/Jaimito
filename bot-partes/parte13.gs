function generarReporte(cod){
  var f = filaPorCodigo(maestro(), cod);
  if(!f) return {texto:'No encontré la propiedad.', mail:'', dir:''};
  var dir = dirDe(f), precio = String(f[C.PRECIO]).trim(), moneda = String(f[C.MONEDA]).trim();
  var seg = ss().getSheetByName('Seguimiento').getDataRange().getValues();
  var n = {consultas:0, visitas:0, propuestas:0}, ultObs = '';
  for(var j=1;j<seg.length;j++){
    if(String(seg[j][0]).trim()!==cod) continue;
    var res = String(seg[j][5]).trim();
    if(res==='Solo consulta') n.consultas++;
    else if(res==='Hizo una propuesta') n.propuestas++;
    else if(/visit|interes/i.test(res) && !/no le interes/i.test(res)) n.visitas++;
    if(seg[j][8]) ultObs = String(seg[j][8]).trim();
  }
  var t = '📊 *Reporte de '+dir+'*\nPeríodo: últimos 30 días al '+hoy()+'\n\n👤 Propietario: '+String(f[C.DUENO]).trim()+'\n📌 Estado: '+String(f[C.ETAPA]).trim()+'\n';
  if(precio) t += '💰 '+precio+(moneda?' '+moneda:'')+'\n';
  t += '\n📈 *Actividad:*\n· Consultas: '+n.consultas+'\n· Visitas / interesados: '+n.visitas+'\n· Propuestas: '+n.propuestas+'\n';
  if(ultObs) t += '\n📝 Última novedad: '+ultObs;
  return {texto:t, mail:String(f[C.MAIL]).trim(), dir:dir};
}

function enviarReporteMail(cod, texto, mail){
  var cuerpo = texto.replace(/\*/g,'').replace(/_/g,'').replace(/\[([^\]]+)\]\([^)]+\)/g,'$1')+'\n\nCualquier consulta estamos a tu disposición.\nAzcuénaga Inmobiliaria';
  GmailApp.sendEmail(mail, 'Reporte de tu propiedad — Azcuénaga Inmobiliaria', cuerpo);
}


function normalizarTel(t){
  var s = String(t).replace(/[^0-9]/g,'');
  return s.replace(/^0/,'').replace(/^(\d{2,4})15/,'$1').replace(/^549/,'').replace(/^54/,'');
}

function buscarPersonaPorTel(tel){
  var v = ss().getSheetByName('Seguimiento').getDataRange().getValues(), o = normalizarTel(tel), nombre = '', cant = 0;
  for(var i=1;i<v.length;i++){
    var t = normalizarTel(v[i][4]);
    if(t && t===o){ cant++; if(v[i][3]) nombre = v[i][3]; }
  }
  return cant>0 ? {nombre:nombre||'(sin nombre)', cant:cant} : null;
}
