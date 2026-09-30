function fichaPropiedad(cod){
  var f = filaPorCodigo(maestro(), cod);
  if(!f) return 'No encontré la propiedad '+cod+'.';
  var precio = String(f[C.PRECIO]).trim(), moneda = String(f[C.MONEDA]).trim(), dueno = String(f[C.DUENO]).trim(), tel = String(f[C.TEL]).trim();
  var compl = String(f[C.COMPL]).trim(), tokko = String(f[C.TOKKO]).trim(), pend = docsPendientes(f);
  var seg = ss().getSheetByName('Seguimiento').getDataRange().getValues(), ints = [];
  for(var j=seg.length-1;j>=1 && ints.length<3;j--)
    if(String(seg[j][0]).trim()===cod) ints.push(fechaTxt(seg[j][1])+' · '+seg[j][3]+' · '+seg[j][5]);
  var t = '🏠 *'+dirDe(f)+'* ('+cod+')\n📌 Etapa: '+String(f[C.ETAPA]).trim()+'\n';
  if(precio) t += '💰 '+precio+(moneda?' '+moneda:'')+'\n';
  if(dueno) t += '👤 Dueño: '+dueno+(tel?' · '+tel:'')+'\n';
  if(compl) t += '📊 Completitud: '+compl+'\n';
  if(pend.length) t += '\n⚠️ *Docs pendientes ('+pend.length+'):*\n'+pend.map(function(d){ return '· '+d; }).join('\n')+'\n';
  if(ints.length) t += '\n💬 *Últimas interacciones:*\n'+ints.map(function(i){ return '· '+i; }).join('\n')+'\n';
  return t + (tokko ? '\n🔗 [Ver en Tokko]('+tokko+')' : '\n_(Sin link Tokko cargado)_');
}

function fechaTxt(x){ return x instanceof Date ? fmt(x,'dd/MM/yyyy') : String(x); }


function cambiarEstado(cod, etapa){
  var sh = ss().getSheetByName(HOJA), v = sh.getDataRange().getValues();
  for(var i=1;i<v.length;i++) if(String(v[i][C.COD]).trim()===cod){
    sh.getRange(i+1,C.ETAPA+1).setValue(etapa);
    sh.getRange(i+1,C.ACT+1).setValue(hoy());
    return;
  }
}

function guardarMailDueno(cod, mail){
  var sh = ss().getSheetByName(HOJA), v = sh.getDataRange().getValues();
  for(var i=1;i<v.length;i++) if(String(v[i][C.COD]).trim()===cod){ sh.getRange(i+1,C.MAIL+1).setValue(mail); return; }
}


function parseFH(f, h){
  var fs = f instanceof Date ? fmt(f,'dd/MM/yyyy') : String(f).trim();
  var hs = h instanceof Date ? fmt(h,'HH:mm') : String(h).trim();
  var m = fs.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/), t = hs.match(/^(\d{1,2}):(\d{2})/);
  if(!m || !t) return null;
  return new Date(Date.UTC(+m[3], +m[2]-1, +m[1], +t[1]+3, +t[2]));
}
