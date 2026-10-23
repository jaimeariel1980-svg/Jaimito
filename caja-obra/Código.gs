// ===== NJ ARQUITECTURA - CAJA DE OBRA - BOT ANITA (multiobra) =====
var SHEET_ID = '1kRqMWKDPwK-4Aj0tDqdanOXodoVCjMjrZyq3Xk-6ruQ';
var EXEC_URL = 'https://script.google.com/macros/s/AKfycbzSA249ZvCJvbUKpllHIpJwBfq7uTl6Lt6CDP8Vh3BjejJ6N7JuASWIP4mN81l96a899g/exec';
var RUBROS = ['Preliminares','Movimiento de suelos','Fundaciones','Estructura H°A°','Mampostería','Aislaciones','Contrapisos','Carpetas','Cubierta','Revoques','Cielorrasos','Revestimientos','Pisos','Zócalos','Aberturas','Marmolería','Herrería y barandas','Vidriería','Pintura','Instalación eléctrica','Instalación sanitaria y agua','Instalación de gas','Ascensor','Aire acondicionado','Limpieza','Higiene y seguridad','Ayuda de gremios','Vereda y exterior','General'];

// ---------- INSTALACION (correr una vez, en este orden) ----------
function resetear(){
  ScriptApp.getProjectTriggers().forEach(function(t){ ScriptApp.deleteTrigger(t); });
  Logger.log(tgApi('setWebhook',{url:EXEC_URL, drop_pending_updates:true}).getContentText());
}
function instalarEstructura(){ armarObras(); armarRubros(); armarResumen(); armarCertificados(); }

// ---------- ESTRUCTURA DEL SHEET ----------
function pintar(c,hr,cols){
  c.setHiddenGridlines(true);
  c.getBandings().forEach(function(b){ b.remove(); });
  var last=Math.min(c.getMaxRows(), Math.max(c.getLastRow()+15, hr+20));
  c.getRange(hr,1,last-hr+1,cols).applyRowBanding(SpreadsheetApp.BandingTheme.PINK,true,false);
  c.getRange(hr,1,1,cols).setBackground('#AD1457').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center').setVerticalAlignment('middle');
  c.setRowHeight(hr,26);
  c.setFrozenRows(hr);
}
function armarObras(){
  var s=ss();
  var mv=s.getSheetByName('Movimientos');
  var headers=mv.getRange(1,1,1,mv.getLastColumn()).getValues()[0];
  if(headers.indexOf('Obra')<0){
    var col=mv.getLastColumn()+1;
    mv.getRange(1,col).setValue('Obra').setFontWeight('bold');
    var last=mv.getLastRow();
    if(last>1) mv.getRange(2,col,last-1,1).setValue('Edificio Roma');
  }
  var old=s.getSheetByName('Obras'); if(old) s.deleteSheet(old);
  var o=s.insertSheet('Obras',0);
  o.getRange('A1:F1').merge().setValue('NJ Arquitectura · Obras').setBackground('#880E4F').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center').setVerticalAlignment('middle');
  o.setRowHeight(1,32);
  o.getRange(3,1,1,6).setValues([['Obra','Dirección','Cobrado','Pagado','Saldo','Estado']]);
  var obras=[['Edificio Roma','San Lorenzo 3211 · Rosario'],['Casa Pellegrini','Pellegrini 850 · Rosario']];
  for(var i=0;i<obras.length;i++){
    var r=i+4;
    o.getRange(r,1,1,2).setValues([obras[i]]);
    o.getRange(r,3).setFormula('=SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;A'+r+';Movimientos!$B:$B;"Ingreso")');
    o.getRange(r,4).setFormula('=SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;A'+r+';Movimientos!$B:$B;"Pago sub")+SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;A'+r+';Movimientos!$B:$B;"Compra")+SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;A'+r+';Movimientos!$B:$B;"Gasto")');
    o.getRange(r,5).setFormula('=C'+r+'-D'+r);
    o.getRange(r,6).setValue('En curso');
  }
  o.getRange(4,3,obras.length,3).setNumberFormat('"$"#,##0');
  o.setColumnWidth(1,150); o.setColumnWidth(2,210); o.setColumnWidth(3,120); o.setColumnWidth(4,120); o.setColumnWidth(5,120);
  pintar(o,3,6);
}
function armarRubros(){
  var s=ss();
  var obras=obrasLista();
  var old=s.getSheetByName('Rubros'); if(old) s.deleteSheet(old);
  var rb=s.insertSheet('Rubros');
  rb.getRange(1,1,1,3).setValues([['Obra','Rubro','Presupuesto']]);
  var rows=[];
  obras.forEach(function(ob){ RUBROS.forEach(function(x){ rows.push([ob,x,'']); }); });
  rb.getRange(2,1,rows.length,3).setValues(rows);
  rb.getRange(2,3,rows.length,1).setNumberFormat('"$"#,##0');
  rb.setColumnWidth(1,150); rb.setColumnWidth(2,200);
  var mv=s.getSheetByName('Movimientos');
  mv.getRange('C2:C').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(RUBROS,true).build());
  pintar(rb,1,3);
}
function armarResumen(){
  var s=ss();
  var old=s.getSheetByName('Resumen'); if(old) s.deleteSheet(old);
  var rs=s.insertSheet('Resumen',1);
  rs.getRange('A1').setValue('Obra:').setFontWeight('bold').setHorizontalAlignment('right');
  rs.getRange('B1').setValue('Edificio Roma').setFontWeight('bold').setBackground('#FCE4EC').setFontColor('#880E4F');
  rs.getRange('B1').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInRange(s.getSheetByName('Obras').getRange('A4:A100'),true).build());
  var hr=3;
  rs.getRange(hr,1,1,6).setValues([['Rubro','Presupuesto','Cobrado','Ejecutado','% Avance','Ganancia']]);
  rs.getRange(hr+1,1,RUBROS.length,1).setValues(RUBROS.map(function(x){return [x];}));
  var data=[];
  for(var i=0;i<RUBROS.length;i++){ var r=hr+1+i;
    data.push([
      '=SUMIFS(Rubros!$C:$C;Rubros!$A:$A;$B$1;Rubros!$B:$B;A'+r+')',
      '=SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$C:$C;A'+r+';Movimientos!$B:$B;"Ingreso")',
      '=SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$C:$C;A'+r+';Movimientos!$B:$B;"Pago sub")+SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$C:$C;A'+r+';Movimientos!$B:$B;"Compra")+SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$C:$C;A'+r+';Movimientos!$B:$B;"Gasto")',
      '=IF(B'+r+'=0;"";D'+r+'/B'+r+')',
      '=C'+r+'-D'+r
    ]);
  }
  rs.getRange(hr+1,2,RUBROS.length,5).setValues(data);
  var t=hr+1+RUBROS.length;
  rs.getRange(t,1,1,6).setValues([['TOTAL','=SUM(B'+(hr+1)+':B'+(t-1)+')','=SUM(C'+(hr+1)+':C'+(t-1)+')','=SUM(D'+(hr+1)+':D'+(t-1)+')','','=C'+t+'-D'+t]]).setFontWeight('bold');
  rs.getRange(hr+1,2,t-hr,3).setNumberFormat('"$"#,##0');
  rs.getRange(hr+1,6,t-hr,1).setNumberFormat('"$"#,##0');
  rs.getRange(hr+1,5,RUBROS.length,1).setNumberFormat('0%');
  pintar(rs,hr,6);
  rs.getRange(t,1,1,6).setBackground('#F48FB1').setFontColor('#4A148C').setFontWeight('bold');
}
function agregarObra(nombre,direccion){
  var s=ss(), o=s.getSheetByName('Obras');
  var r=o.getLastRow()+1;
  o.getRange(r,1,1,2).setValues([[nombre,direccion||'']]);
  o.getRange(r,3).setFormula('=SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;A'+r+';Movimientos!$B:$B;"Ingreso")');
  o.getRange(r,4).setFormula('=SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;A'+r+';Movimientos!$B:$B;"Pago sub")+SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;A'+r+';Movimientos!$B:$B;"Compra")+SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;A'+r+';Movimientos!$B:$B;"Gasto")');
  o.getRange(r,5).setFormula('=C'+r+'-D'+r);
  o.getRange(r,6).setValue('En curso');
  o.getRange(r,3,1,3).setNumberFormat('"$"#,##0');
  var rb=s.getSheetByName('Rubros');
  var rows=RUBROS.map(function(x){ return [nombre,x,'']; });
  rb.getRange(rb.getLastRow()+1,1,rows.length,3).setValues(rows);
  pintar(o,3,6);
}
function armarCertificados(){
  var s=ss();
  var old=s.getSheetByName('Certificados'); if(old) s.deleteSheet(old);
  var c=s.insertSheet('Certificados');
  c.getRange('A1:H1').merge().setValue('NJ Arquitectura · Certificados de Honorarios').setBackground('#880E4F').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(13).setHorizontalAlignment('center').setVerticalAlignment('middle');
  c.getRange('A2:H2').merge().setValue('Dirección Técnica — Reajuste mensual por IPC (inflación acumulada compuesta)').setBackground('#D81B60').setFontColor('#FFFFFF').setFontSize(10).setHorizontalAlignment('center');
  c.setRowHeight(1,30);
  c.getRange('A4').setValue('PARÁMETROS POR OBRA').setFontWeight('bold').setFontColor('#880E4F');
  c.getRange(5,1,1,4).setValues([['Obra','Honorario mensual','Plazo (meses)','Fecha base']]);
  var params=[['Edificio Roma',833333,6,new Date(2026,3,30)],['Casa Pellegrini','','','']];
  c.getRange(6,1,params.length,4).setValues(params);
  c.getRange(6,2,10,1).setNumberFormat('"$"#,##0');
  c.getRange(6,4,10,1).setNumberFormat('dd/mm/yyyy');
  c.getRange('A16').setValue('CERTIFICADOS EMITIDOS').setFontWeight('bold').setFontColor('#880E4F');
  c.getRange(17,1,1,8).setValues([['Obra','N°','Mes','% IPC del mes','Inflación acumulada','Monto reajuste','Honorarios del mes','Fecha']]);
  var certs=[['Edificio Roma',1,'Mayo 2026',0.021,0.021,17500,850833,new Date(2026,5,1)],['Edificio Roma',2,'Junio 2026',0.019,0.040399,33666,866999,new Date(2026,6,1)]];
  c.getRange(18,1,certs.length,8).setValues(certs);
  c.getRange(18,4,certs.length,2).setNumberFormat('0.0###%');
  c.getRange(18,6,certs.length,2).setNumberFormat('"$"#,##0');
  c.getRange(18,8,certs.length,1).setNumberFormat('dd/mm/yyyy');
  c.setHiddenGridlines(true);
  c.getBandings().forEach(function(b){b.remove();});
  c.getRange(5,1,1,4).setBackground('#AD1457').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  c.getRange(6,1,10,4).applyRowBanding(SpreadsheetApp.BandingTheme.PINK,false,false);
  c.getRange(17,1,1,8).setBackground('#AD1457').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  c.getRange(18,1,20,8).applyRowBanding(SpreadsheetApp.BandingTheme.PINK,false,false);
  c.setColumnWidth(1,140); c.setColumnWidth(3,110); c.setColumnWidth(5,150); c.setColumnWidth(7,140); c.setColumnWidth(8,120);
  c.setFrozenRows(17);
}

// ---------- TELEGRAM ----------
function tgToken(){ return PropertiesService.getScriptProperties().getProperty('TG_TOKEN'); }
function oaKey(){ return PropertiesService.getScriptProperties().getProperty('OA_KEY'); }
function tgApi(m,p){ return UrlFetchApp.fetch('https://api.telegram.org/bot'+tgToken()+'/'+m,{method:'post',contentType:'application/json',payload:JSON.stringify(p),muteHttpExceptions:true}); }
function enviar(chatId,texto){ tgApi('sendMessage',{chat_id:chatId,text:texto}); }
function carpeta(){ var it=DriveApp.getFoldersByName('Comprobantes Obra Roma'); return it.hasNext()?it.next():DriveApp.createFolder('Comprobantes Obra Roma'); }

// ---------- MEMORIA / OBRA ACTIVA ----------
function memGet(chatId){ var v=CacheService.getScriptCache().get('mem_'+chatId); return v?JSON.parse(v):[]; }
function memPush(chatId,rol,texto){ var m=memGet(chatId); m.push({r:rol,t:String(texto).slice(0,300)}); if(m.length>8) m=m.slice(-8); CacheService.getScriptCache().put('mem_'+chatId,JSON.stringify(m),1800); }
function obraActiva(chatId){ return CacheService.getScriptCache().get('obra_'+chatId)||'Edificio Roma'; }
function setObra(chatId,obra){ CacheService.getScriptCache().put('obra_'+chatId,obra,21600); }

// ---------- HELPERS ----------
function ss(){ return SpreadsheetApp.openById(SHEET_ID); }
function normal(s){ return (s||'').toString().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,''); }
function money(n){ return '$'+(Number(n)||0).toLocaleString('es-AR'); }
function movimientos(){ var s=ss().getSheetByName('Movimientos'); var last=s.getLastRow(); return last<2?[]:s.getRange(2,1,last-1,11).getValues(); }
function esEgreso(t){ return t==='Pago sub'||t==='Compra'||t==='Gasto'; }
function obrasLista(){ var s=ss().getSheetByName('Obras'); if(!s) return ['Edificio Roma']; var last=s.getLastRow(); if(last<4) return ['Edificio Roma']; return s.getRange(4,1,last-3,1).getValues().map(function(r){return r[0];}).filter(String); }
function matchObra(texto){ var t=normal(texto).trim(); if(!t) return null; var obs=obrasLista(); for(var i=0;i<obs.length;i++){ var n=normal(obs[i]); if(n.indexOf(t)>=0||t.indexOf(n)>=0) return obs[i]; } var pals=t.split(/\s+/).filter(function(w){return w.length>3;}); for(var i=0;i<obs.length;i++){ var n=normal(obs[i]); if(pals.some(function(w){return n.indexOf(w)>=0;})) return obs[i]; } return null; }
function matchRubro(texto){
  var t=normal(texto).trim(); if(!t) return null;
  var SIN={'plomeria':'Instalación sanitaria y agua','plomero':'Instalación sanitaria y agua','sanitaria':'Instalación sanitaria y agua','agua':'Instalación sanitaria y agua','cloacas':'Instalación sanitaria y agua','desague':'Instalación sanitaria y agua','electrica':'Instalación eléctrica','electricidad':'Instalación eléctrica','electricista':'Instalación eléctrica','luz':'Instalación eléctrica','cableado':'Instalación eléctrica','gas':'Instalación de gas','gasista':'Instalación de gas','pintura':'Pintura','pintor':'Pintura','mamposteria':'Mampostería','ladrillos':'Mampostería','pared':'Mampostería','albañil':'Mampostería','albanil':'Mampostería','estructura':'Estructura H°A°','hormigon':'Estructura H°A°','losa':'Estructura H°A°','columna':'Estructura H°A°','hierro':'Estructura H°A°','revoque':'Revoques','revoques':'Revoques','pisos':'Pisos','ceramica':'Pisos','porcelanato':'Pisos','cielorraso':'Cielorrasos','durlock':'Cielorrasos','aberturas':'Aberturas','puerta':'Aberturas','ventana':'Aberturas','ascensor':'Ascensor','aire':'Aire acondicionado','split':'Aire acondicionado','techo':'Cubierta','cubierta':'Cubierta','fundacion':'Fundaciones','fundaciones':'Fundaciones','cimientos':'Fundaciones','pilote':'Fundaciones','excavacion':'Movimiento de suelos','limpieza':'Limpieza','herreria':'Herrería y barandas','baranda':'Herrería y barandas','reja':'Herrería y barandas','vidrio':'Vidriería','vidrios':'Vidriería','vidrieria':'Vidriería','marmol':'Marmolería','marmoleria':'Marmolería','granito':'Marmolería','mesada':'Marmolería','aislacion':'Aislaciones','aislaciones':'Aislaciones','membrana':'Aislaciones','contrapiso':'Contrapisos','contrapisos':'Contrapisos','carpeta':'Carpetas','carpetas':'Carpetas','zocalo':'Zócalos','zocalos':'Zócalos','revestimiento':'Revestimientos','revestimientos':'Revestimientos','azulejo':'Revestimientos','vereda':'Vereda y exterior','exterior':'Vereda y exterior','preliminares':'Preliminares','obrador':'Preliminares','gremios':'Ayuda de gremios','higiene':'Higiene y seguridad','seguridad':'Higiene y seguridad'};
  var pals=t.split(/\s+/);
  for(var j=0;j<pals.length;j++){ if(SIN[pals[j]]) return SIN[pals[j]]; }
  for(var k in SIN){ if(t.indexOf(k)>=0) return SIN[k]; }
  for(var i=0;i<RUBROS.length;i++){ var n=normal(RUBROS[i]); if(n.indexOf(t)>=0||t.indexOf(n)>=0) return RUBROS[i]; }
  var pl=t.split(/\s+/).filter(function(w){return w.length>3;});
  for(var i=0;i<RUBROS.length;i++){ var n=normal(RUBROS[i]); if(pl.some(function(w){return n.indexOf(w)>=0;})) return RUBROS[i]; }
  return null;
}

// ---------- CEREBRO ----------
function llamarIA(messages,jsonMode){
  var body={model:'gpt-4o',temperature:0.3,messages:messages};
  if(jsonMode) body.response_format={type:'json_object'};
  var r=UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions',{method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+oaKey()},payload:JSON.stringify(body),muteHttpExceptions:true});
  return JSON.parse(r.getContentText()).choices[0].message.content;
}
function clasificar(chatId,texto){
  var hoy=Utilities.formatDate(new Date(),'GMT-3','dd/MM/yyyy');
  var mem=memGet(chatId).map(function(x){return (x.r==='u'?'Usuario: ':'Anita: ')+x.t;}).join('\n');
  var sys='Sos Anita, asistente de una obra en Rosario (NJ Arquitectura). Hablás informal, rosarino, canchera pero clara. Hoy es '+hoy+'. Obra activa (solo por defecto al CARGAR): '+obraActiva(chatId)+'. '+
  'Vos manejás TODAS las obras por igual. NUNCA digas que una obra la lleva otro equipo, NUNCA derives a otra persona y NUNCA inventes datos, montos ni nombres. Si preguntan por cualquier obra, usá "consultar" (no "charla").\n'+
  'Devolvé SOLO un JSON con "accion":\n'+
  '- {"accion":"cargar","mov":"Ingreso|Pago sub|Compra|Gasto","rubro":"<de la lista o General>","concepto":"...","proveedor":"...","cantidad":0,"precio_unitario":0,"monto":0,"factura":"Sí|No","obra":"<obra si la nombra, sino vacío>","bonus":0}\n'+
  '- {"accion":"consultar","que":"total_rubro|detalle|ganancia_rubro|ganancia_total|saldo|ultimos","rubro":"<si aplica>","obra":"<obra si la nombra, sino vacío>"}\n'+
  '- {"accion":"corregir","ref":"palabra para encontrar la fila","campo":"monto|rubro|proveedor|concepto|factura","valor":"..."}\n'+
  '- {"accion":"borrar","ref":"palabra para encontrar la fila"}\n'+
  '- {"accion":"charla","respuesta":"tu respuesta canchera y útil"}\n\n'+
  'que: "cuánto gasté/gastamos en X" -> total_rubro. "detalle/listado/desglose de X" -> detalle. "ganancia" -> ganancia_total o ganancia_rubro. "saldo/cuánta plata hay" -> saldo. "últimos movimientos" -> ultimos.\n'+
  'obra: si el mensaje nombra una obra (Roma, Pellegrini, etc.) ponela en "obra"; si no, dejá "".\n'+
  'bonus: SOLO al cargar. Si dice "con X% de bonus" (o margen/diferencial), poné X/100 (ej 0.2). Si no lo menciona, 0.\n'+
  'charla es SOLO para saludos o charla suelta; ahí tampoco inventes datos.\n'+
  'PLATA (crítico): NUNCA calcules. Si hay cantidad y precio ("3 bolsas a 5000") poné cantidad y precio_unitario y monto 0. Si es importe directo ("pagué 8000") poné monto. "5 lucas"=5000, "200 mil"=200000, "2 palos"=2000000.\n'+
  'mov: pago a electricista/plomero/albañil=Pago sub; materiales=Compra; cobro del propietario=Ingreso; suelto (flete,comida)=Gasto.\n'+
  'rubro: SOLO uno exacto de: '+RUBROS.join(', ')+'. Si no se deduce, "General". factura: "con factura"=Sí, "en negro"=No, si no dice "No".';
  var msgs=[{role:'system',content:sys}];
  if(mem) msgs.push({role:'system',content:'Charla reciente:\n'+mem});
  msgs.push({role:'user',content:texto});
  try{ return JSON.parse(llamarIA(msgs,true)); }catch(e){ return {accion:'charla',respuesta:'Perdón, no te entendí, ¿me lo repetís?'}; }
}

// ---------- ENTRADA DE TEXTO ----------
function manejarTexto(chatId,texto){
  memPush(chatId,'u',texto);
  var cache=CacheService.getScriptCache();
  var MENU={keyboard:[[{text:'☰ Menú'}]],resize_keyboard:true,is_persistent:true};
  if(/☰|^\s*(hola|menu|menú|ayuda|opciones|inicio|start|\/start)\b/i.test(texto)){ menuPrincipal(chatId); return; }
  if(/corregir.*certificad|✏️\s*corregir/i.test(texto)){ menuCorregir(chatId); return; }
  var mCC=texto.match(/^✏️\s*(.+)/);
  if(mCC){ cache.put('cc_'+chatId, JSON.stringify({obra:obraActiva(chatId),mes:mCC[1].trim()}),600); tgApi('sendMessage',{chat_id:chatId,text:'¿Cuál es el IPC correcto de '+mCC[1].trim()+'? Escribí el % (ej: 2,3).',reply_markup:MENU}); return; }
  if(/⬅️\s*Volver|^volver/i.test(texto)){ menuPrincipal(chatId); return; }
  if(/🏗\s*Elegir obra|^elegir obra/i.test(texto)){ menuElegirObra(chatId); return; }
  if(/cargar un gasto|💸/i.test(texto) && !/👉/.test(texto)){ menuCargarObra(chatId); return; }
  if(/^💸👉\s*(.+)/.test(texto)){ var oc=matchObra(texto.replace(/^💸👉\s*/,'')); if(oc){ setObra(chatId,oc); tgApi('sendMessage',{chat_id:chatId,text:'📍 Cargando en '+oc+'.\nContame el gasto: por ejemplo "compré 3 bolsas de cemento a 5000" (texto, audio o foto).',reply_markup:MENU}); return; } }
  if(/hacer un certificado|📄/i.test(texto) && !/👉/.test(texto)){ menuCertObra(chatId); return; }
  if(/^📄👉\s*(.+)/.test(texto)){ var oq=matchObra(texto.replace(/^📄👉\s*/,'')); if(oq){ setObra(chatId,oq); tgApi('sendMessage',{chat_id:chatId,text:'📄 Certificado de '+oq+'.\n¿Cuál fue el IPC de este mes? Escribime por ejemplo: "IPC julio 2,5%".',reply_markup:MENU}); return; } }
  if(/^👉\s*(.+)/.test(texto)){ var oe=matchObra(texto.replace(/^👉\s*/,'')); if(oe){ setObra(chatId,oe); tgApi('sendMessage',{chat_id:chatId,text:'📍 Listo, ahora cargás en '+oe+'.'}); menuPrincipal(chatId); return; } }
  var cc=cache.get('cc_'+chatId);
  if(cc){
    var pm=texto.match(/([\d]+(?:[.,]\d+)?)/);
    if(pm){ var d=JSON.parse(cc); cache.remove('cc_'+chatId); corregirCertificado(chatId,d.mes,Number(pm[1].replace(',','.'))/100,d.obra); return; }
    cache.remove('cc_'+chatId);
  }
  var gg=cache.get('gg_'+chatId);
  if(gg){
    if(ggResolver(chatId,texto,JSON.parse(gg))){ cache.remove('gg_'+chatId); return; }
    cache.remove('gg_'+chatId);
  }
  if(/(correg|correcc|corrige|cambi[aá]|\bera\b|estaba)/i.test(texto) && /ipc|certificad|honorar|inflaci/i.test(texto)){
    var mesC=texto.match(/(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)/i);
    var pctC=texto.match(/([\d]+(?:[.,]\d+)?)\s*%/);
    if(mesC&&pctC){
      var obc=matchObra(texto); if(obc) setObra(chatId,obc);
      corregirCertificado(chatId,mesC[1],Number(pctC[1].replace(',','.'))/100); return;
    }
  }
  var ipc=texto.match(/ipc[^\d]*([\d]+(?:[.,]\d+)?)\s*%?/i);
  if(ipc){
    var pct=Number(ipc[1].replace(',','.'))/100;
    var resto=texto.replace(/ipc[^\d]*[\d.,]+\s*%?/i,' ').trim();
    var obm=matchObra(resto); if(obm) setObra(chatId,obm);
    var mm=resto.match(/([a-záéíóúñ]+\s+\d{4})/i);
    nuevoCertificado(chatId,pct,mm?mm[1]:''); return;
  }
  var bm=texto.match(/(\d+(?:[.,]\d+)?)\s*%\s*(?:de\s+)?(?:bonus|margen|diferencial)/i);
  if(bm && /(ponele|pon[eé]|agreg|sum[aá]le|sumale)/i.test(texto) && !/^(compr|pagu|gast|anot|carg)/i.test(texto.trim())){
    var pctb=Number(bm[1].replace(',','.'))/100;
    var refb=texto.replace(/(\d+(?:[.,]\d+)?)\s*%\s*(?:de\s+)?(?:bonus|margen|diferencial)/i,' ').replace(/\b(ponele|pon[eé]|agregale|agreg[aá]|sum[aá]le|a la|al|a el|de|la|el|compra|pago|gasto)\b/gi,' ').trim();
    accionBonus(chatId,pctb,refb); return;
  }
  if(/bonus.*(rubro|por rubro|detalle)|detalle.*bonus/i.test(texto)){ var obb=matchObra(texto); accionBonusRubro(chatId,obb); return; }
  if(/\bpdf\b/i.test(texto)){ var op=matchObra(texto)||obraActiva(chatId); resumenPDF(chatId,op); return; }
  if(/\b(resumen|informe)\b/i.test(texto)){ var obr=matchObra(texto); accionResumen(chatId,obr); return; }
  var nm=texto.match(/nueva obra[:,]?\s+(.+)/i);
  if(nm){
    var partes=nm[1].split(/\s+en\s+/i), nombre=partes[0].trim(), dir=(partes[1]||'').trim();
    if(matchObra(nombre)){ responder(chatId,'Esa obra ya existe.'); return; }
    agregarObra(nombre,dir); setObra(chatId,nombre);
    responder(chatId,'🏗 Creé la obra "'+nombre+'"'+(dir?' ('+dir+')':'')+' y quedó como obra activa.'); return;
  }
  var mo=texto.match(/^\s*(?:obra|cambiar a|cargar en)\s+(.+)/i);
  if(mo){ var ob=matchObra(mo[1]); if(ob){ setObra(chatId,ob); responder(chatId,'📍 Listo, ahora cargás en '+ob+'.'); return; } }
  var pend=cache.get('pr_'+chatId);
  if(pend){
    if(/dejal|dejar|general|as[ií] est/i.test(texto)){ cache.remove('pr_'+chatId); responder(chatId,'Dale, lo dejo en General.'); return; }
    var rb=matchRubro(texto);
    if(rb){ ss().getSheetByName('Movimientos').getRange(Number(pend),3).setValue(rb); cache.remove('pr_'+chatId); responder(chatId,'✅ Listo, lo puse en '+rb+'.'); return; }
    responder(chatId,'Ese rubro no lo reconocí. Probá una palabra (o "dejalo" y queda en General).'); return;
  }
  if(esGastoGeneral(texto) && /\d/.test(texto) && !/(cu[aá]nto|dame|detalle|resumen|ganancia|saldo|bonus|\?)/i.test(texto)){
    var ag=clasificar(chatId,texto);
    if(ag.accion==='cargar'){
      ag.rubro=rubroGeneral(texto);
      cache.put('gg_'+chatId, JSON.stringify(ag), 600);
      preguntarGeneral(chatId);
      return;
    }
  }
  ejecutar(chatId,clasificar(chatId,texto));
}
function responder(chatId,texto){ enviar(chatId,texto); memPush(chatId,'a',texto); }

// ---------- EJECUTOR ----------
function ejecutar(chatId,a){
  if(a.accion==='cargar') return accionCargar(chatId,a);
  if(a.accion==='consultar') return accionConsultar(chatId,a);
  if(a.accion==='corregir') return accionCorregir(chatId,a);
  if(a.accion==='borrar') return accionBorrar(chatId,a);
  responder(chatId,a.respuesta||'Dale, contame el gasto, cobro o consulta.');
}
function accionCargar(chatId,a){
  if(!a._ok){
    var obs=obrasLista().filter(function(o){return o!=='Estudio';});
    if(obs.length>1){
      var kb=obs.map(function(o){ return [{text:'📍 '+o}]; });
      CacheService.getScriptCache().put('gg_'+chatId, JSON.stringify(a), 600);
      tgApi('sendMessage',{chat_id:chatId,text:'¿En qué obra lo cargo?',reply_markup:{keyboard:kb,one_time_keyboard:true,resize_keyboard:true}});
      return;
    }
  }
  var cant=Number(a.cantidad)||0, precio=Number(a.precio_unitario)||0;
  var monto=(cant>0&&precio>0)?cant*precio:(Number(a.monto)||0);
  var rubro=a.rubro||'General', prov=a.proveedor||'', obra=a.obra||obraActiva(chatId);
  var bonus=Number(a.bonus)||0;
  var gbon=Math.round(monto*bonus);
  var dup=false;
  if(monto>0&&prov){ movimientos().forEach(function(r){ if((Number(r[5])||0)===monto&&normal(r[4])===normal(prov)&&r[10]===obra) dup=true; }); }
  var fecha=a.fecha?parseFecha(a.fecha):new Date();
  var mv=ss().getSheetByName('Movimientos');
  mv.appendRow([fecha, a.mov||'', rubro, a.concepto||'', prov, monto, a.factura||'No', a.link||'', 'Telegram', 'Por confirmar', obra, bonus>0?bonus:'', bonus>0?gbon:'']);
  var av=dup?'\n⚠ OJO: ya tenías uno igual ('+money(monto)+' a '+prov+') en '+obra+'.':'';
  var bon=bonus>0?'\n💰 Bonus '+Math.round(bonus*100)+'%: '+money(gbon):'';
  if(rubro==='General'){
    CacheService.getScriptCache().put('pr_'+chatId, String(mv.getLastRow()), 600);
    responder(chatId,'✅ Anoté '+(a.mov||'')+' de '+money(monto)+' en '+obra+'.'+bon+av+'\n¿De qué rubro es? (una palabra: eléctrica, plomería...)');
  }else{
    responder(chatId,'✅ Cargado en '+obra+': '+(a.mov||'')+' · '+rubro+' · '+money(monto)+bon+av);
  }
}
function accionConsultar(chatId,a){
  var rows=movimientos();
  var obra=a.obra?matchObra(a.obra):null;
  if(!obra){
    var set={}; rows.forEach(function(r){ if(r[10]) set[r[10]]=1; });
    var cd=Object.keys(set);
    if(cd.length===1) obra=cd[0];
    else if(cd.length>1){
      var base;
      if(a.que==='detalle') base='detalle de '+(a.rubro||'gastos');
      else if(a.que==='total_rubro') base='cuánto gasté en '+(a.rubro||'');
      else if(a.que==='ganancia_rubro') base='ganancia de '+(a.rubro||'');
      else if(a.que==='ganancia_total') base='ganancia';
      else if(a.que==='saldo') base='saldo';
      else if(a.que==='ultimos') base='últimos movimientos';
      else base='resumen';
      var botones=cd.map(function(o){ return [{text:base+' en '+o}]; });
      tgApi('sendMessage',{chat_id:chatId,text:'¿De qué obra?',reply_markup:{keyboard:botones,one_time_keyboard:true,resize_keyboard:true}});
      return;
    }
    else obra=obraActiva(chatId);
  }
  function egr(rb){ var s=0; rows.forEach(function(r){ if(r[10]===obra&&esEgreso(r[1])&&(!rb||r[2]===rb)) s+=Number(r[5])||0; }); return s; }
  function ing(rb){ var s=0; rows.forEach(function(r){ if(r[10]===obra&&r[1]==='Ingreso'&&(!rb||r[2]===rb)) s+=Number(r[5])||0; }); return s; }
  var rb=a.rubro?matchRubro(a.rubro):null, txt='';
  if(a.que==='detalle'){
    var ls=rows.filter(function(r){ return r[10]===obra&&(!rb||r[2]===rb)&&esEgreso(r[1]); });
    if(!ls.length){ txt='No hay gastos en '+(rb||'ningún rubro')+' de '+obra+'.'; }
    else{
      var det=ls.map(function(r){ var f=(r[0] instanceof Date)?Utilities.formatDate(r[0],'GMT-3','dd/MM'):''; return '• '+f+' · '+money(r[5])+' · '+(r[4]||'s/prov')+(r[3]?' — '+r[3]:''); }).join('\n');
      var tot=ls.reduce(function(ac,r){ return ac+(Number(r[5])||0); },0);
      txt='📋 '+(rb||'Gastos')+' en '+obra+':\n'+det+'\n─────────\nTotal: '+money(tot);
    }
  }
  else if(a.que==='total_rubro'){ txt='En '+obra+', '+(rb||'ese rubro')+': gastaste '+money(egr(rb))+'.'; }
  else if(a.que==='ganancia_rubro'){ txt=obra+' · '+(rb||'ese rubro')+': cobraste '+money(ing(rb))+', gastaste '+money(egr(rb))+', ganancia '+money(ing(rb)-egr(rb))+'.'; }
  else if(a.que==='ganancia_total'){ txt=obra+': ganancia '+money(ing()-egr())+' (cobrado '+money(ing())+' − gastado '+money(egr())+').'; }
  else if(a.que==='saldo'){ txt=obra+': saldo '+money(ing()-egr())+' (entró '+money(ing())+', salió '+money(egr())+').'; }
  else if(a.que==='ultimos'){ var u=rows.filter(function(r){return r[10]===obra;}).slice(-5).reverse().map(function(r){ return '• '+r[1]+' · '+r[2]+' · '+money(r[5])+(r[4]?' ('+r[4]+')':''); }).join('\n'); txt='Últimos de '+obra+':\n'+(u||'nada cargado.'); }
  else { txt='¿Qué querés saber de '+obra+'? Gastado por rubro, detalle, ganancia, saldo o últimos.'; }
  tgApi('sendMessage',{chat_id:chatId,text:txt,reply_markup:{remove_keyboard:true}});
  memPush(chatId,'a',txt);
}
function buscarFila(chatId,ref){
  var mv=ss().getSheetByName('Movimientos'), last=mv.getLastRow(), obra=obraActiva(chatId);
  if(last<2||!ref) return 0; var n=normal(ref);
  for(var f=last; f>=2; f--){ var r=mv.getRange(f,1,1,11).getValues()[0]; if(r[10]===obra&&(normal(r[4]).indexOf(n)>=0||normal(r[3]).indexOf(n)>=0||normal(r[2]).indexOf(n)>=0)) return f; }
  return 0;
}
function accionCorregir(chatId,a){
  var f=buscarFila(chatId,a.ref); if(!f) return responder(chatId,'No encontré esa fila. ¿De qué proveedor o concepto era?');
  var mv=ss().getSheetByName('Movimientos'), col={monto:6,rubro:3,proveedor:5,concepto:4,factura:7}[a.campo];
  if(!col) return responder(chatId,'¿Qué corrijo: monto, rubro, proveedor, concepto o factura?');
  var val=a.valor;
  if(a.campo==='monto') val=Number(String(a.valor).replace(/[^\d]/g,''))||0;
  if(a.campo==='rubro') val=matchRubro(a.valor)||a.valor;
  mv.getRange(f,col).setValue(val);
  responder(chatId,'✅ Corregí: '+a.campo+' → '+(a.campo==='monto'?money(val):val)+'.');
}
function accionBorrar(chatId,a){
  var f=buscarFila(chatId,a.ref); if(!f) return responder(chatId,'No encontré esa fila para borrar. ¿De qué proveedor o concepto era?');
  var mv=ss().getSheetByName('Movimientos'), r=mv.getRange(f,1,1,6).getValues()[0];
  mv.deleteRow(f);
  responder(chatId,'🗑 Borré: '+r[1]+' · '+r[2]+' · '+money(r[5])+(r[4]?' ('+r[4]+')':''));
}
function parseFecha(s){ var m=String(s||'').match(/(\d{1,2})[\/\-\. ]+(\d{1,2})[\/\-\. ]+(\d{2,4})/); if(!m) return new Date(); var y=m[3].length===2?('20'+m[3]):m[3]; var d=new Date(Number(y),Number(m[2])-1,Number(m[1])); return isNaN(d.getTime())?new Date():d; }

// ---------- AUDIO ----------
function manejarAudio(chatId,msg){
  try{
    var fileId=(msg.voice||msg.audio).file_id;
    var path=JSON.parse(tgApi('getFile',{file_id:fileId}).getContentText()).result.file_path;
    var blob=UrlFetchApp.fetch('https://api.telegram.org/file/bot'+tgToken()+'/'+path).getBlob().setName('audio.ogg');
    var r=UrlFetchApp.fetch('https://api.openai.com/v1/audio/transcriptions',{method:'post',headers:{Authorization:'Bearer '+oaKey()},payload:{file:blob,model:'whisper-1',language:'es'},muteHttpExceptions:true});
    var texto=JSON.parse(r.getContentText()).text;
    if(!texto) return responder(chatId,'No escuché bien el audio, ¿lo repetís?');
    manejarTexto(chatId,texto);
  }catch(err){ responder(chatId,'No pude procesar el audio.'); }
}

// ---------- FOTO / PDF ----------
function leerComprobante(chatId,blob,mime,nombre,caption){
  var file=carpeta().createFile(blob.setName(nombre));
  try{ file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(x){}
  var b64=Utilities.base64Encode(blob.getBytes());
  var extra=(caption?('\nEl usuario aclaró: "'+caption+'".'):'');
  var instruc='Es un comprobante de obra (factura, remito, ticket o transferencia).'+extra+' Devolvé SOLO un JSON: {"mov":"Compra|Pago sub|Gasto|Ingreso","rubro":"<de la lista o General>","concepto":"qué es","proveedor":"comercio o emisor","monto":<total solo dígitos>,"fecha":"dd/mm/aaaa si figura, sino vacío","factura":"Sí|No"}. rubro SOLO uno de: '+RUBROS.join(', ')+'. factura A/B/C=Sí, remito/ticket=No. materiales=Compra, servicio=Pago sub. Si no leés el monto, 0.';
  var contenido;
  if(mime.indexOf('image/')===0){ contenido=[{type:'text',text:instruc},{type:'image_url',image_url:{url:'data:'+mime+';base64,'+b64}}]; }
  else { contenido=[{type:'text',text:instruc},{type:'file',file:{filename:/\.pdf$/i.test(nombre)?nombre:(nombre+'.pdf'),file_data:'data:application/pdf;base64,'+b64}}]; }
  var r=UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions',{method:'post',contentType:'application/json',headers:{Authorization:'Bearer '+oaKey()},payload:JSON.stringify({model:'gpt-4o',temperature:0,response_format:{type:'json_object'},messages:[{role:'user',content:contenido}]}),muteHttpExceptions:true});
  var ev={}; try{ ev=JSON.parse(JSON.parse(r.getContentText()).choices[0].message.content); }catch(e){}
  accionCargar(chatId,{mov:ev.mov,rubro:ev.rubro,concepto:ev.concepto,proveedor:ev.proveedor,monto:ev.monto,fecha:ev.fecha,factura:ev.factura,link:file.getUrl()});
}
function manejarFoto(chatId,msg){
  try{
    var fileId=msg.photo[msg.photo.length-1].file_id;
    var path=JSON.parse(tgApi('getFile',{file_id:fileId}).getContentText()).result.file_path;
    var blob=UrlFetchApp.fetch('https://api.telegram.org/file/bot'+tgToken()+'/'+path).getBlob();
    leerComprobante(chatId,blob,'image/jpeg','comprobante_'+Date.now()+'.jpg',msg.caption);
  }catch(err){ responder(chatId,'No pude leer el comprobante.'); }
}
function manejarDocumento(chatId,msg){
  try{
    var doc=msg.document, mime=doc.mime_type||'';
    var path=JSON.parse(tgApi('getFile',{file_id:doc.file_id}).getContentText()).result.file_path;
    var blob=UrlFetchApp.fetch('https://api.telegram.org/file/bot'+tgToken()+'/'+path).getBlob();
    leerComprobante(chatId,blob,mime,doc.file_name||('comprobante_'+Date.now()+'.pdf'),msg.caption);
  }catch(err){ responder(chatId,'No pude leer el documento.'); }
}
function ponerFiltro(){
  var mv=SpreadsheetApp.openById('1kRqMWKDPwK-4Aj0tDqdanOXodoVCjMjrZyq3Xk-6ruQ').getSheetByName('Movimientos');
  if(mv.getFilter()) mv.getFilter().remove();
  mv.getRange(1,1,mv.getLastRow(),mv.getLastColumn()).createFilter();
}
function demoPellegrini(){
  var s=SpreadsheetApp.openById('1kRqMWKDPwK-4Aj0tDqdanOXodoVCjMjrZyq3Xk-6ruQ');
  var mv=s.getSheetByName('Movimientos');
  mv.appendRow([new Date(),'Ingreso','General','Anticipo obra','Propietario Pellegrini',2000000,'Sí','','Manual','Por confirmar','Casa Pellegrini']);
  mv.appendRow([new Date(),'Compra','Mampostería','Ladrillos','Corralón Sur',450000,'No','','Manual','Por confirmar','Casa Pellegrini']);
  var rb=s.getSheetByName('Rubros');
  if(rb.getFilter()) rb.getFilter().remove();
  rb.getRange(1,1,rb.getLastRow(),rb.getLastColumn()).createFilter();
}
function certParams(obra){
  var c=ss().getSheetByName('Certificados');
  var vals=c.getRange(6,1,10,4).getValues();
  var orphan=-1;
  for(var i=0;i<vals.length;i++){
    if(vals[i][0]===obra) return {obra:obra, hon:Number(vals[i][1])||0, plazo:vals[i][2], base:vals[i][3], fila:6+i};
    if(orphan<0 && (vals[i][0]===''||vals[i][0]==null) && (Number(vals[i][1])||0)>0) orphan=i;
  }
  if(orphan>=0) return {obra:obra, hon:Number(vals[orphan][1])||0, plazo:vals[orphan][2], base:vals[orphan][3], fila:6+orphan, huerfana:true};
  return {obra:obra, hon:0, plazo:'', base:'', fila:0};
}
function certPctPrevios(obra){
  var c=ss().getSheetByName('Certificados'), last=c.getLastRow(), arr=[];
  for(var r=18;r<=last;r++){ if(c.getRange(r,1).getValue()===obra){ var v=Number(c.getRange(r,4).getValue()); if(!isNaN(v)&&v>0&&v<1) arr.push(v); } }
  return arr;
}
function pctTxt(v){ var x=(v*100).toFixed(4).replace(/0+$/,'').replace(/\.$/,''); return x.replace('.',',')+' %'; }
function nuevoCertificado(chatId,pct,mesTxt,obraForce){
  var obra=obraForce||obraActiva(chatId);
  if(!obraForce){
    var obs=obrasLista().filter(function(o){return o!=='Estudio';});
    if(obs.length>1){
      var kb=obs.map(function(o){ return [{text:'📍 '+o}]; });
      CacheService.getScriptCache().put('gg_'+chatId, JSON.stringify({_cert:{pct:pct,mesTxt:mesTxt}}), 600);
      tgApi('sendMessage',{chat_id:chatId,text:'¿Certificado de qué obra?',reply_markup:{keyboard:kb,one_time_keyboard:true,resize_keyboard:true}});
      return;
    }
  }
  if(!(pct>0&&pct<1)){ responder(chatId,'Ese % no parece un IPC mensual válido (0 a 100%). Probá tipo "IPC 2,5%" o "IPC julio 2,5%".'); return; }
  var c=ss().getSheetByName('Certificados'), p=certParams(obra);
  if(!p.hon){ responder(chatId,'Falta cargar el honorario mensual de '+obra+' en la pestaña Certificados (Parámetros por obra).'); return; }
  if(p.huerfana&&p.fila){ c.getRange(p.fila,1).setValue(obra); }
  var factor=1; certPctPrevios(obra).forEach(function(x){ factor*=(1+x); }); factor*=(1+pct);
  var acum=factor-1;
  if(!isFinite(acum)||acum>50){ responder(chatId,'Los % anteriores de '+obra+' parecen mal cargados (la acumulada dio un número absurdo). Revisá la columna "% IPC del mes" en Certificados, o corré armarCertificados para reiniciar.'); return; }
  var reaj=Math.round(p.hon*acum), total=p.hon+reaj;
  var last=c.getLastRow(), n=0;
  for(var r=18;r<=last;r++){ if(c.getRange(r,1).getValue()===obra) n++; }
  n=n+1;
  var MESES=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  var bd=p.base?new Date(p.base):new Date();
  var idx=bd.getMonth()+n;
  var etiqueta=mesTxt||(MESES[((idx%12)+12)%12]+' '+(bd.getFullYear()+Math.floor(idx/12)));
  var hoy=new Date(), row=last+1;
  if(c.getRange(18,1).getValue()==='') row=18;
  c.getRange(row,1,1,8).setValues([[obra, n, etiqueta, pct, acum, reaj, total, hoy]]);
  c.getRange(row,4,1,2).setNumberFormat('0.0###%');
  c.getRange(row,6,1,2).setNumberFormat('"$"#,##0');
  c.getRange(row,8).setNumberFormat('dd/mm/yyyy');
  var mv=ss().getSheetByName('Movimientos');
  mv.appendRow([hoy, 'Ingreso', 'Honorarios', 'Honorarios '+etiqueta, 'Cliente', total, 'Sí', '', 'Certificado', 'Por confirmar', obra, '', '']);
  var pdf=certPDF(p,n,etiqueta,pct,acum,reaj,total,hoy);
  tgDoc(chatId,pdf,'📄 Certificado N° '+n+' de '+obra+'. Reenviáselo al cliente.');
  responder(chatId,'✅ '+obra+' · Certificado N° '+n+' ('+etiqueta+') · acumulada '+pctTxt(acum)+' · reajuste '+money(reaj)+' · total '+money(total)+'.\n💵 Cargué el honorario como Ingreso de la obra. Te lo mandé en PDF.');
}
function certPDF(p,n,mes,pct,acum,reaj,total,fecha){
  function m(v){ return '$ '+(Number(v)||0).toLocaleString('es-AR'); }
  var mesTit=mesLindo(mes);
  var f=Utilities.formatDate(fecha,'GMT-3','dd/MM/yyyy');
  var baseF=p.base?Utilities.formatDate(new Date(p.base),'GMT-3','dd/MM/yyyy'):'';
  var plazo=Number(p.plazo)||0, conv=p.hon*plazo;
  var dir='';
  try{
    var ob=ss().getSheetByName('Obras').getDataRange().getValues();
    for(var k=0;k<ob.length;k++){ if(ob[k][0]===p.obra){ dir=ob[k][1]||''; break; } }
  }catch(e){}
  var meses=[];
  try{
    var c=ss().getSheetByName('Certificados').getDataRange().getValues();
    for(var r=0;r<c.length;r++){ if(c[r][0]===p.obra && typeof c[r][3]==='number'){ meses.push({mes:c[r][2], pct:c[r][3]}); } }
  }catch(e){}
  var filasMeses='';
  meses.forEach(function(x){ filasMeses+='<tr><td style="padding:2px 4px">inflación de '+mesLindo(x.mes)+'</td><td style="padding:2px 4px;text-align:right">'+pctTxt(x.pct)+'</td></tr>'; });
  var html='<html><body style="font-family:Georgia,serif;color:#222;padding:40px">'+
   '<div style="text-align:center;border-bottom:3px solid #AD1457;padding-bottom:10px;margin-bottom:20px">'+
   '<div style="font-size:22px;font-weight:bold;color:#AD1457">NJ Arquitectura</div>'+
   '<div style="font-size:13px;color:#666">Dirección Técnica de Obra</div></div>'+
   '<div style="font-size:15px;font-weight:bold">OBRA: '+p.obra+'</div>'+
   '<div style="font-size:12px;color:#666;margin-bottom:14px">'+dir+'</div>'+
   '<div style="font-size:13px;font-weight:bold;text-align:center;margin-bottom:16px;text-transform:uppercase">Certificado de Honorarios por Dirección Técnica de la Estructura</div>'+
   '<table style="width:100%;font-size:13px;border-collapse:collapse">'+
   '<tr><td style="padding:4px">Honorarios convenidos</td><td style="padding:4px;text-align:right">'+m(conv)+'</td></tr>'+
   '<tr><td style="padding:4px">Plazo por los trabajos</td><td style="padding:4px;text-align:right">'+plazo+' meses</td></tr>'+
   '<tr><td style="padding:4px">Honorario mensual</td><td style="padding:4px;text-align:right">'+m(p.hon)+'</td></tr>'+
   '<tr><td style="padding:4px">Fecha base</td><td style="padding:4px;text-align:right">'+baseF+'</td></tr>'+
   '<tr><td style="padding:4px">Reajuste</td><td style="padding:4px;text-align:right">Mensual según inflación</td></tr></table>'+
   '<div style="background:#FCE4EC;border-left:4px solid #AD1457;padding:14px;margin:18px 0">'+
   '<div style="font-size:15px;font-weight:bold;color:#AD1457;margin-bottom:8px">CERTIFICADO N° '+n+' — '+mesTit+'</div>'+
   '<table style="width:100%;font-size:13px">'+
   '<tr><td style="padding:4px">Fecha</td><td style="padding:4px;text-align:right">'+f+'</td></tr></table>'+
   '<div style="font-size:12px;font-weight:bold;margin:8px 0 2px">% de reajuste</div>'+
   '<table style="width:100%;font-size:13px">'+filasMeses+
   '<tr style="font-weight:bold"><td style="padding:4px;border-top:1px solid #AD1457">Inflación acumulada</td><td style="padding:4px;border-top:1px solid #AD1457;text-align:right">'+pctTxt(acum)+'</td></tr>'+
   '</table>'+
   '<table style="width:100%;font-size:13px;margin-top:8px">'+
   '<tr><td style="padding:4px">Monto de reajuste</td><td style="padding:4px;text-align:right">'+m(reaj)+'</td></tr>'+
   '<tr style="font-weight:bold;font-size:14px"><td style="padding:8px 4px;border-top:2px solid #AD1457">Honorarios totales por '+mesTit+'</td><td style="padding:8px 4px;border-top:2px solid #AD1457;text-align:right">'+m(total)+'</td></tr>'+
   '</table></div>'+
   '<div style="font-size:11px;color:#888;margin-top:30px">Generado el '+f+' · S.E.U.O.</div></body></html>';
  return Utilities.newBlob(html,'text/html','c.html').getAs('application/pdf').setName('Certificado_'+n+'_'+String(p.obra).replace(/\s+/g,'_')+'.pdf');
}
function tgDoc(chatId,blob,caption){
  UrlFetchApp.fetch('https://api.telegram.org/bot'+tgToken()+'/sendDocument',{method:'post',payload:{chat_id:String(chatId),document:blob,caption:caption||''},muteHttpExceptions:true});
}
function limpiarTodo(){
  var s=ss();
  var mv=s.getSheetByName('Movimientos');
  var lm=mv.getLastRow();
  if(lm>1) mv.getRange(2,1,lm-1,mv.getLastColumn()).clearContent();
  SpreadsheetApp.flush();
}
function accionResumen(chatId,obra){
  var MENU={keyboard:[[{text:'☰ Menú'}]],resize_keyboard:true,is_persistent:true};
  if(!obra){
    var obs=obrasLista().filter(function(o){return o!=='Estudio';});
    if(obs.length===1) obra=obs[0];
    else if(obs.length>1){ var bt=obs.map(function(o){return [{text:'resumen en '+o}];}); bt.push([{text:'☰ Menú'}]); tgApi('sendMessage',{chat_id:chatId,text:'¿De qué obra?',reply_markup:{keyboard:bt,resize_keyboard:true,is_persistent:true}}); return; }
    else obra=obraActiva(chatId);
  }
  var rows=movimientos();
  var ls=rows.filter(function(r){ return r[10]===obra; });
  var cob=0,pag=0,porRubro={};
  ls.forEach(function(r){
    var mo=Number(r[5])||0, rub=r[2]||'General';
    if(!porRubro[rub]) porRubro[rub]={c:0,g:0};
    if(r[1]==='Ingreso'){ cob+=mo; porRubro[rub].c+=mo; }
    else if(esEgreso(r[1])){ pag+=mo; porRubro[rub].g+=mo; }
  });
  var saldo=cob-pag;
  var t='📊 RESUMEN · '+obra+'\n─────────\n';
  t+='Cobrado:  '+money(cob)+'\n';
  t+='Pagado:   '+money(pag)+'\n';
  t+='Saldo:    '+money(saldo)+'\n';
  var claves=Object.keys(porRubro).filter(function(k){ return porRubro[k].c>0||porRubro[k].g>0; });
  if(claves.length){
    t+='\nPor rubro:\n';
    claves.forEach(function(k){ t+='• '+k+': gastado '+money(porRubro[k].g)+(porRubro[k].c>0?' · cobrado '+money(porRubro[k].c):'')+'\n'; });
  } else { t+='\n(sin movimientos todavía)'; }
  tgApi('sendMessage',{chat_id:chatId,text:t,reply_markup:MENU});
  memPush(chatId,'a','resumen '+obra);
  tgApi('sendMessage',{chat_id:chatId,text:'¿Lo querés en PDF con membrete? Escribime PDF.',reply_markup:MENU});
}
function resumenPDF(chatId,obra){
  var rows=movimientos().filter(function(r){return r[10]===obra;});
  var cob=0,pag=0,porRubro={};
  rows.forEach(function(r){ var mo=Number(r[5])||0,rub=r[2]||'General'; if(!porRubro[rub])porRubro[rub]={c:0,g:0}; if(r[1]==='Ingreso'){cob+=mo;porRubro[rub].c+=mo;} else if(esEgreso(r[1])){pag+=mo;porRubro[rub].g+=mo;} });
  var saldo=cob-pag;
  function m(v){return '$ '+(Number(v)||0).toLocaleString('es-AR');}
  var f=Utilities.formatDate(new Date(),'GMT-3','dd/MM/yyyy');
  var filas='';
  Object.keys(porRubro).forEach(function(k){ if(porRubro[k].c>0||porRubro[k].g>0){ filas+='<tr><td style="padding:6px;border-bottom:1px solid #F3C6D8">'+k+'</td><td style="padding:6px;border-bottom:1px solid #F3C6D8;text-align:right">'+m(porRubro[k].c)+'</td><td style="padding:6px;border-bottom:1px solid #F3C6D8;text-align:right">'+m(porRubro[k].g)+'</td></tr>'; } });
  var html='<html><body style="font-family:Georgia,serif;color:#222;padding:40px">'+
   '<div style="text-align:center;border-bottom:3px solid #AD1457;padding-bottom:10px;margin-bottom:20px"><div style="font-size:22px;font-weight:bold;color:#AD1457">NJ Arquitectura</div><div style="font-size:13px;color:#666">Resumen de Obra</div></div>'+
   '<div style="font-size:16px;font-weight:bold;margin-bottom:2px">'+obra+'</div>'+
   '<div style="font-size:12px;color:#666;margin-bottom:18px">Al '+f+'</div>'+
   '<table style="width:100%;font-size:14px;border-collapse:collapse;margin-bottom:20px">'+
   '<tr><td style="padding:6px">Cobrado</td><td style="padding:6px;text-align:right">'+m(cob)+'</td></tr>'+
   '<tr><td style="padding:6px">Pagado</td><td style="padding:6px;text-align:right">'+m(pag)+'</td></tr>'+
   '<tr style="font-weight:bold;color:#AD1457"><td style="padding:6px;border-top:2px solid #AD1457">Saldo / Ganancia</td><td style="padding:6px;border-top:2px solid #AD1457;text-align:right">'+m(saldo)+'</td></tr>'+
   '</table>'+
   (filas?'<div style="font-size:13px;font-weight:bold;color:#AD1457;margin-bottom:6px">Por rubro</div><table style="width:100%;font-size:13px;border-collapse:collapse"><tr style="background:#AD1457;color:#fff"><th style="padding:6px;text-align:left">Rubro</th><th style="padding:6px;text-align:right">Cobrado</th><th style="padding:6px;text-align:right">Gastado</th></tr>'+filas+'</table>':'')+
   '<div style="font-size:11px;color:#888;margin-top:30px">Generado el '+f+' · NJ Arquitectura</div></body></html>';
  var pdf=Utilities.newBlob(html,'text/html','r.html').getAs('application/pdf').setName('Resumen_'+String(obra).replace(/\s+/g,'_')+'.pdf');
  tgDoc(chatId,pdf,'📄 Resumen de '+obra+'.');
}
function accionBonus(chatId,pct,ref){
  var f=buscarFila(chatId,ref); if(!f) return responder(chatId,'No encontré esa compra/pago. ¿De qué proveedor o concepto era?');
  var mv=ss().getSheetByName('Movimientos');
  var monto=Number(mv.getRange(f,6).getValue())||0;
  var gbon=Math.round(monto*pct);
  mv.getRange(f,12).setValue(pct);
  mv.getRange(f,13).setValue(gbon);
  responder(chatId,'✅ Bonus '+Math.round(pct*100)+'% aplicado. Ganancia: '+money(gbon)+'.');
}
 function agregarDesglose(){
  var rs=ss().getSheetByName('Resumen');
  var last=rs.getLastRow(), totalRow=0;
  for(var r=1;r<=last;r++){ if(rs.getRange(r,1).getValue()==='TOTAL'){ totalRow=r; break; } }
  if(!totalRow) totalRow=last;
  var st=totalRow+2;
  rs.getRange(st,1).setValue('DESGLOSE DE GANANCIA').setFontWeight('bold').setFontColor('#880E4F');
  var filas=[
    ['Caja (anticipos − egresos)','=SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$B:$B;"Ingreso")-SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$B:$B;"Pago sub")-SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$B:$B;"Compra")-SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$B:$B;"Gasto")'],
    ['Bonus (compras/pagos)','=SUMIFS(Movimientos!$M:$M;Movimientos!$K:$K;$B$1)'],
    ['Honorarios (certificados)','=SUMIFS(Certificados!$G:$G;Certificados!$A:$A;$B$1)']
  ];
  rs.getRange(st+1,1,filas.length,2).setValues(filas);
  var gt=st+1+filas.length;
  rs.getRange(gt,1).setValue('GANANCIA TOTAL').setFontWeight('bold');
  rs.getRange(gt,2).setFormula('=B'+(st+2)+'+B'+(st+3));
  rs.getRange(st+1,2,filas.length+1,1).setNumberFormat('"$"#,##0');
  rs.getRange(gt,1,1,2).setBackground('#F48FB1').setFontColor('#4A148C').setFontWeight('bold');
}
function accionBonusRubro(chatId,obra){
  var rows=movimientos();
  if(!obra){
    var set={}; rows.forEach(function(r){ if(r[10]) set[r[10]]=1; });
    var cd=Object.keys(set);
    if(cd.length===1) obra=cd[0];
    else if(cd.length>1){ var bt=cd.map(function(o){return [{text:'bonus por rubro en '+o}];}); tgApi('sendMessage',{chat_id:chatId,text:'¿De qué obra?',reply_markup:{keyboard:bt,one_time_keyboard:true,resize_keyboard:true}}); return; }
    else obra=obraActiva(chatId);
  }
  var porRubro={}, total=0;
  rows.forEach(function(r){ if(r[10]===obra){ var g=Number(r[12])||0; if(g>0){ var rub=r[2]||'General'; porRubro[rub]=(porRubro[rub]||0)+g; total+=g; } } });
  var claves=Object.keys(porRubro);
  var txt='💰 BONUS por rubro · '+obra+'\n─────────\n';
  if(!claves.length){ txt+='(todavía no hay bonus cargado)'; }
  else{ claves.forEach(function(k){ txt+='• '+k+': '+money(porRubro[k])+'\n'; }); txt+='─────────\nTotal bonus: '+money(total); }
  tgApi('sendMessage',{chat_id:chatId,text:txt,reply_markup:{remove_keyboard:true}});
  memPush(chatId,'a','bonus por rubro '+obra);
}
function agregarBonusRubro(){
  var rs=ss().getSheetByName('Resumen');
  var last=rs.getLastRow(), hr=3, totalRow=0;
  for(var r=1;r<=last;r++){ if(rs.getRange(r,1).getValue()==='TOTAL'){ totalRow=r; break; } }
  if(!totalRow) return;
  rs.getRange(hr,7).setValue('Bonus').setBackground('#AD1457').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  var n=totalRow-1-hr;
  var forms=[];
  for(var i=0;i<n;i++){ var row=hr+1+i; forms.push(['=SUMIFS(Movimientos!$M:$M;Movimientos!$K:$K;$B$1;Movimientos!$C:$C;A'+row+')']); }
  rs.getRange(hr+1,7,n,1).setValues(forms).setNumberFormat('"$"#,##0');
  rs.getRange(totalRow,7).setFormula('=SUM(G'+(hr+1)+':G'+(totalRow-1)+')').setNumberFormat('"$"#,##0').setFontWeight('bold');
}
function arreglarFormatoBonus(){
  var mv=ss().getSheetByName('Movimientos');
  var hr=1; for(var r=1;r<=3;r++){ if(mv.getRange(r,1).getValue()==='Fecha'){ hr=r; break; } }
  mv.getRange(hr+1,12,mv.getMaxRows()-hr,1).setNumberFormat('0%');
}
function arreglarBonus(){
  var mv=ss().getSheetByName('Movimientos');
  var hr=2;
  mv.getRange(hr,12).setValue('Bonus %').setFontWeight('bold');
  mv.getRange(hr,13).setValue('Ganancia bonus').setFontWeight('bold');
  mv.getRange(hr+1,12,mv.getMaxRows()-hr,1).setNumberFormat('0%');
  mv.getRange(hr+1,13,mv.getMaxRows()-hr,1).setNumberFormat('"$"#,##0');
  var rs=ss().getSheetByName('Resumen');
  rs.getRange(3,7).setValue('Bonus').setBackground('#AD1457').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
}
function cerrarTabla(){
  var mv=ss().getSheetByName('Movimientos');
  var titRow=1, hr=2, lastCol=13;
  if(mv.getFilter()) mv.getFilter().remove();
  mv.getRange(titRow,1,1,lastCol).breakApart();
  var titulo=mv.getRange(titRow,1).getValue()||'NJ Arquitectura · Movimientos';
  mv.getRange(titRow,1,1,lastCol).merge().setValue(titulo).setBackground('#880E4F').setFontColor('#FFFFFF').setFontWeight('bold').setFontSize(12).setHorizontalAlignment('center').setVerticalAlignment('middle');
  mv.getBandings().forEach(function(b){ b.remove(); });
  var filas=Math.min(mv.getMaxRows(), Math.max(mv.getLastRow()+10, hr+20));
  mv.getRange(hr,1,filas-hr+1,lastCol).applyRowBanding(SpreadsheetApp.BandingTheme.PINK,true,false);
  mv.getRange(hr,1,1,lastCol).setBackground('#AD1457').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  mv.getRange(hr,1,mv.getLastRow()-hr+1,lastCol).createFilter();
}
function datosObras(){
  var rows=movimientos();
  var obras=obrasLista();
  var out=[];
  var dir={};
  var o=ss().getSheetByName('Obras');
  if(o){ var last=o.getLastRow(); if(last>=4){ o.getRange(4,1,last-3,2).getValues().forEach(function(r){ if(r[0]) dir[r[0]]=r[1]||''; }); } }
  obras.forEach(function(ob){
    var cob=0,pag=0,bonus=0;
    rows.forEach(function(r){
      if(r[10]!==ob) return;
      var m=Number(r[5])||0;
      if(r[1]==='Ingreso') cob+=m;
      else if(r[1]==='Pago sub'||r[1]==='Compra'||r[1]==='Gasto') pag+=m;
      bonus+=Number(r[12])||0;
    });
    var hon=0;
    var c=ss().getSheetByName('Certificados');
    if(c){ var lc=c.getLastRow(); for(var i=18;i<=lc;i++){ if(c.getRange(i,1).getValue()===ob) hon+=Number(c.getRange(i,7).getValue())||0; } }
    out.push({obra:ob, direccion:dir[ob]||'', cobrado:cob, pagado:pag, saldo:cob-pag, bonus:bonus, honorarios:hon, ganancia:bonus+hon});
  });
  return out;
}
function testDatosObras(){ Logger.log(JSON.stringify(datosObras(),null,2)); }
function agregarRubroHonorarios(){
  var lista=RUBROS.slice();
  if(lista.indexOf('Honorarios')<0) lista.push('Honorarios');
  var mv=ss().getSheetByName('Movimientos');
  var hr=2;
  mv.getRange(hr+1,3,mv.getMaxRows()-hr,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(lista,true).build()
  );
}
function crearEstudio(){
  if(matchObra('Estudio')){ return; }
  agregarObra('Estudio','Gastos fijos del estudio (sueldos, oficina, etc.)');
}
function borrarEstudio(){
  var s=ss(), o=s.getSheetByName('Obras');
  var last=o.getLastRow();
  for(var r=last;r>=4;r--){ if(o.getRange(r,1).getValue()==='Estudio'){ o.deleteRow(r); } }
  var rb=s.getSheetByName('Rubros'), lr=rb.getLastRow();
  for(var r=lr;r>=2;r--){ if(rb.getRange(r,1).getValue()==='Estudio'){ rb.deleteRow(r); } }
}
function esGastoGeneral(texto){
  return /sueldo|administrativ|alquiler|oficina|contador|servicios|internet|impuesto|monotributo|matr[ií]cula|licencia|software/i.test(texto);
}
function rubroGeneral(texto){
  var t=normal(texto);
  if(/sueldo|administrativ|emplead/.test(t)) return 'Sueldos';
  if(/alquiler|oficina/.test(t)) return 'Oficina';
  if(/contador/.test(t)) return 'Contadora';
  if(/servicio|internet|luz|telefono/.test(t)) return 'Servicios';
  if(/impuesto|monotributo|matricula|afip/.test(t)) return 'Impuestos';
  return 'Sueldos';
}
function repartirGasto(chatId,a){
  var obras=obrasLista().filter(function(o){ return o!=='Estudio'; });
  var monto=(Number(a.cantidad)>0&&Number(a.precio_unitario)>0)?a.cantidad*a.precio_unitario:(Number(a.monto)||0);
  if(!obras.length){ responder(chatId,'No hay obras cargadas para imputar el gasto.'); return; }
  var parte=Math.round(monto/obras.length);
  var mv=ss().getSheetByName('Movimientos'), hoy=new Date();
  obras.forEach(function(o){
    mv.appendRow([hoy,'Gasto',a.rubro||'Sueldos',a.concepto||'Gasto general',a.proveedor||'',parte,'No','','Telegram','Por confirmar',o,'','']);
  });
  var txt = obras.length===1
    ? '✅ Cargué '+money(monto)+' ('+(a.rubro||'Sueldos')+') en '+obras[0]+'.'
    : '✅ Repartí '+money(monto)+' ('+(a.rubro||'Sueldos')+') entre '+obras.length+' obras: '+money(parte)+' en cada una.';
  responder(chatId,txt);
}
function agregarRubrosGenerales(){
  var lista=RUBROS.slice();
  ['Honorarios','Sueldos','Oficina','Servicios','Contadora','Impuestos'].forEach(function(x){ if(lista.indexOf(x)<0) lista.push(x); });
  var mv=ss().getSheetByName('Movimientos');
  mv.getRange(3,3,mv.getMaxRows()-2,1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(lista,true).build()
  );
}
function preguntarGeneral(chatId){
  var obrasR=obrasLista().filter(function(o){return o!=='Estudio';});
  var kb=[[{text:'Repartir entre obras'}]];
  obrasR.forEach(function(o){ kb.push([{text:'Solo en '+o}]); });
  tgApi('sendMessage',{chat_id:chatId,text:'¿Gasto general del estudio (se reparte entre '+obrasR.join(' y ')+') o de una obra sola?',reply_markup:{keyboard:kb,one_time_keyboard:true,resize_keyboard:true}});
}
function ggResolver(chatId,texto,a){
  if(/repartir/i.test(texto)){ repartirGasto(chatId,a); return true; }
  var m=texto.match(/(?:solo en|📍)\s*(.+)/i);
  if(m){ var ob=matchObra(m[1]); if(ob){
    if(a._cert){ nuevoCertificado(chatId,a._cert.pct,a._cert.mesTxt,ob); return true; }
    a.obra=ob; a._ok=true; accionCargar(chatId,a); return true;
  } }
  return false;
}
function limpiarGanancia(){
  var rs=ss().getSheetByName('Resumen');
  var last=rs.getLastRow(), hr=3, totalRow=0;
  for(var r=1;r<=last;r++){ if(rs.getRange(r,1).getValue()==='TOTAL'){ totalRow=r; break; } }
  if(!totalRow) return;
  rs.getRange(hr,6).setValue('Bonus').setBackground('#AD1457').setFontColor('#FFFFFF').setFontWeight('bold').setHorizontalAlignment('center');
  var forms=[];
  for(var row=hr+1; row<totalRow; row++){ forms.push(['=SUMIFS(Movimientos!$M:$M;Movimientos!$K:$K;$B$1;Movimientos!$C:$C;A'+row+')']); }
  rs.getRange(hr+1,6,forms.length,1).setValues(forms).setNumberFormat('"$"#,##0');
  rs.getRange(totalRow,6).setFormula('=SUM(F'+(hr+1)+':F'+(totalRow-1)+')').setNumberFormat('"$"#,##0').setFontWeight('bold');
  if(rs.getRange(hr,7).getValue()==='Bonus'){ rs.getRange(1,7,rs.getMaxRows(),1).clearContent(); }
}
function desgloseNeto(){
  var rs=ss().getSheetByName('Resumen');
  var last=rs.getLastRow(), st=0;
  for(var r=1;r<=last;r++){ if(rs.getRange(r,1).getValue()==='DESGLOSE DE GANANCIA'){ st=r; break; } }
  if(!st){ for(var r=1;r<=last;r++){ if(rs.getRange(r,1).getValue()==='TOTAL'){ st=r+2; break; } } rs.getRange(st,1).setValue('DESGLOSE DE GANANCIA').setFontWeight('bold').setFontColor('#880E4F'); }
  var egG='SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$C:$C;"Sueldos")+SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$C:$C;"Oficina")+SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$C:$C;"Servicios")+SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$C:$C;"Contadora")+SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$C:$C;"Impuestos")';
  var filas=[
    ['Caja (anticipos − egresos)','=SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$B:$B;"Ingreso")-SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$B:$B;"Pago sub")-SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$B:$B;"Compra")-SUMIFS(Movimientos!$F:$F;Movimientos!$K:$K;$B$1;Movimientos!$B:$B;"Gasto")'],
    ['Bonus (compras/pagos)','=SUMIFS(Movimientos!$M:$M;Movimientos!$K:$K;$B$1)'],
    ['Honorarios (certificados)','=SUMIFS(Certificados!$G:$G;Certificados!$A:$A;$B$1)'],
    ['Gastos de estudio (sueldos, oficina...)','=-('+egG+')']
  ];
  rs.getRange(st+1,1,filas.length,2).setValues(filas);
  var gt=st+1+filas.length;
  rs.getRange(gt,1).setValue('GANANCIA NETA').setFontWeight('bold');
  rs.getRange(gt,2).setFormula('=B'+(st+2)+'+B'+(st+3)+'+B'+(st+4));
  rs.getRange(st+1,2,filas.length+1,1).setNumberFormat('"$"#,##0');
  rs.getRange(gt,1,1,2).setBackground('#F48FB1').setFontColor('#4A148C').setFontWeight('bold');
}
function datosObra(ob){
  var rows=movimientos();
  var rub={}, cob=0,pag=0,bonus=0;
  rows.forEach(function(r){
    if(r[10]!==ob) return;
    var m=Number(r[5])||0, rb=r[2]||'General', gb=Number(r[12])||0;
    if(!rub[rb]) rub[rb]={cobrado:0,ejecutado:0,bonus:0};
    if(r[1]==='Ingreso'){ cob+=m; rub[rb].cobrado+=m; }
    else if(r[1]==='Pago sub'||r[1]==='Compra'||r[1]==='Gasto'){ pag+=m; rub[rb].ejecutado+=m; }
    rub[rb].bonus+=gb; bonus+=gb;
  });
  var rubros=[];
  Object.keys(rub).forEach(function(k){ var v=rub[k]; if(v.cobrado||v.ejecutado||v.bonus) rubros.push({rubro:k,cobrado:v.cobrado,ejecutado:v.ejecutado,bonus:v.bonus}); });
  var hon=0;
  var c=ss().getSheetByName('Certificados');
  if(c){ var lc=c.getLastRow(); for(var i=18;i<=lc;i++){ if(c.getRange(i,1).getValue()===ob) hon+=Number(c.getRange(i,7).getValue())||0; } }
  var gastoEst=0;
  ['Sueldos','Oficina','Servicios','Contadora','Impuestos'].forEach(function(g){ if(rub[g]) gastoEst+=rub[g].ejecutado; });
  var ult=rows.filter(function(r){return r[10]===ob;}).slice(-8).reverse().map(function(r){
    var f=(r[0] instanceof Date)?Utilities.formatDate(r[0],'GMT-3','dd/MM'):'';
    return {fecha:f, tipo:r[1], rubro:r[2], monto:Number(r[5])||0, prov:r[4]||''};
  });
  return {obra:ob, cobrado:cob, pagado:pag, saldo:cob-pag, bonus:bonus, honorarios:hon, gastoEstudio:gastoEst, neta:bonus+hon-gastoEst, rubros:rubros, ultimos:ult};
}
function autoReset(){
  try{
    var r=JSON.parse(UrlFetchApp.fetch('https://api.telegram.org/bot'+tgToken()+'/getWebhookInfo',{muteHttpExceptions:true}).getContentText());
    var info=r.result||{};
    var pend=Number(info.pending_update_count)||0;
    var err=info.last_error_message||'';
    var urlMal=(info.url||'').indexOf('/exec')<0;
    if(pend>3 || err || urlMal){
      tgApi('setWebhook',{url:EXEC_URL, drop_pending_updates:true});
    }
  }catch(e){}
}
function instalarAutoReset(){
  ScriptApp.getProjectTriggers().forEach(function(t){ if(t.getHandlerFunction()==='autoReset') ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('autoReset').timeBased().everyMinutes(5).create();
}
function verWebhook(){ Logger.log(UrlFetchApp.fetch('https://api.telegram.org/bot'+tgToken()+'/getWebhookInfo',{muteHttpExceptions:true}).getContentText()); }
function doPost(e){
  try{
    if(!e||!e.postData||!e.postData.contents) return HtmlService.createHtmlOutput('ok');
    var update=JSON.parse(e.postData.contents);
    if(update.update_id!=null){
      var c=CacheService.getScriptCache();
      if(c.get('upd_'+update.update_id)) return HtmlService.createHtmlOutput('ok');
      c.put('upd_'+update.update_id,'1',600);
    }
    var msg=update.message; if(!msg||!msg.chat) return HtmlService.createHtmlOutput('ok');
    var chatId=msg.chat.id;
    if(msg.voice||msg.audio) manejarAudio(chatId,msg);
    else if(msg.photo) manejarFoto(chatId,msg);
    else if(msg.document) manejarDocumento(chatId,msg);
    else if(msg.text) manejarTexto(chatId,msg.text);
  }catch(err){ console.error(err); }
  return HtmlService.createHtmlOutput('ok');
}
function menuPrincipal(chatId){
  var kb=[
    [{text:'💸 Cargar un gasto'}],
    [{text:'📊 Ver un resumen'}],
    [{text:'📄 Hacer un certificado'}],
    [{text:'✏️ Corregir certificado'}],
    [{text:'💰 Bonus por rubro'}],
    [{text:'🏗 Elegir obra'}]
  ];
  tgApi('sendMessage',{chat_id:chatId,text:'¡Hola! Soy Anita 👷‍♀️\n¿Qué querés hacer? Tocá una opción (o escribime directo lo que quieras cargar).',reply_markup:{keyboard:kb,resize_keyboard:true,is_persistent:true}});
}
function menuElegirObra(chatId){
  var obras=obrasLista().filter(function(o){ return o!=='Estudio'; });
  var kb=obras.map(function(o){ return [{text:'👉 '+o}]; });
  kb.push([{text:'⬅️ Volver'}]);
  tgApi('sendMessage',{chat_id:chatId,text:'¿En qué obra querés cargar?',reply_markup:{keyboard:kb,resize_keyboard:true}});
}
function menuCargarObra(chatId){
  var obras=obrasLista().filter(function(o){ return o!=='Estudio'; });
  var kb=obras.map(function(o){ return [{text:'💸👉 '+o}]; });
  kb.push([{text:'⬅️ Volver'}]);
  tgApi('sendMessage',{chat_id:chatId,text:'¿En qué obra vas a cargar el gasto?',reply_markup:{keyboard:kb,resize_keyboard:true}});
}
function menuCertObra(chatId){
  var obras=obrasLista().filter(function(o){ return o!=='Estudio'; });
  var kb=obras.map(function(o){ return [{text:'📄👉 '+o}]; });
  kb.push([{text:'⬅️ Volver'}]);
  tgApi('sendMessage',{chat_id:chatId,text:'¿De qué obra es el certificado?',reply_markup:{keyboard:kb,resize_keyboard:true}});
}
function resetRoma(){ limpiarCertificados('Edificio Roma'); }
function limpiarCertificados(obra){
  var c=ss().getSheetByName('Certificados');
  for(var r=c.getLastRow();r>=18;r--){ if(c.getRange(r,1).getValue()===obra) c.deleteRow(r); }
  var mv=ss().getSheetByName('Movimientos');
  var v=mv.getDataRange().getValues();
  for(var i=v.length-1;i>=1;i--){ if(v[i][1]==='Ingreso'&&v[i][2]==='Honorarios'&&v[i][10]===obra) mv.deleteRow(i+1); }
  SpreadsheetApp.getActive().toast('Certificados de '+obra+' reiniciados');
}
function vaciarMovimientos(){
  var mv=ss().getSheetByName('Movimientos');
  var col=mv.getRange(1,1,mv.getLastRow(),1).getValues();
  var h=0;
  for(var i=0;i<col.length;i++){ if(col[i][0]==='Fecha'){ h=i+1; break; } }
  if(h>0 && mv.getLastRow()>h) mv.getRange(h+1,1,mv.getLastRow()-h,mv.getLastColumn()).clearContent();
  SpreadsheetApp.getActive().toast('Movimientos vaciado — planilla virgen');
}
function dejarVirgen(){
  var S=ss();
  var ru=S.getSheetByName('Rubros');
  var rv=ru.getDataRange().getValues(), rh=-1;
  for(var i=0;i<rv.length;i++){ if(rv[i][0]==='Obra'&&rv[i][1]==='Rubro'){ rh=i; break; } }
  if(rh>=0){
    for(var i=rh+1;i<rv.length;i++){
      if(!rv[i][1]) continue;
      var nom=String(rv[i][1]).trim().toLowerCase();
      if(nom==='plomeria') ru.getRange(i+1,2).setValue('Preliminares');
      else if(nom==='albanileria') ru.getRange(i+1,2).setValue('Movimiento de suelos');
      ru.getRange(i+1,3).clearContent();
    }
  }
  var re=S.getSheetByName('Resumen');
  var ev=re.getDataRange().getValues(), eh=-1;
  for(var i=0;i<ev.length;i++){ if(ev[i][0]==='Rubro'&&ev[i][1]==='Presupuesto'){ eh=i; break; } }
  if(eh>=0){
    for(var i=eh+1;i<ev.length;i++){
      var a=String(ev[i][0]||'').trim();
      if(a===''||a==='DESGLOSE DE GANANCIA') break;
      if(a==='TOTAL') continue;
      re.getRange(i+1,2).setFormula('=SUMIFS(Rubros!C:C,Rubros!A:A,$B$1,Rubros!B:B,A'+(i+1)+')');
    }
  }
  var mv=S.getSheetByName('Movimientos');
  var mc=mv.getRange(1,1,mv.getLastRow(),1).getValues();
  for(var i=0;i<mc.length;i++){ if(mc[i][0]==='Fecha'){ mv.getRange(i+1,11).setValue('Obra'); break; } }
  SpreadsheetApp.getActive().toast('Planilla virgen y ordenada ✓');
}
function repararResumen(){
  var re=ss().getSheetByName('Resumen');
  var ev=re.getDataRange().getValues(), eh=-1;
  for(var i=0;i<ev.length;i++){ if(ev[i][0]==='Rubro'&&ev[i][1]==='Presupuesto'){ eh=i; break; } }
  if(eh<0){ SpreadsheetApp.getActive().toast('No encontré el encabezado del Resumen'); return; }
  for(var i=eh+1;i<ev.length;i++){
    var a=String(ev[i][0]||'').trim();
    if(a===''||a==='DESGLOSE DE GANANCIA') break;
    if(a==='TOTAL') continue;
    re.getRange(i+1,2).setFormula('=SUMIFS(Rubros!C:C;Rubros!A:A;$B$1;Rubros!B:B;A'+(i+1)+')');
  }
  SpreadsheetApp.getActive().toast('Resumen reparado ✓');
}
function corregirCertificado(chatId,mesTxt,nuevoPct,obraForce){
  var obra=obraForce||obraActiva(chatId);
  var c=ss().getSheetByName('Certificados'), last=c.getLastRow(), filas=[];
  for(var r=18;r<=last;r++){
    if(c.getRange(r,1).getValue()===obra){
      filas.push({row:r, n:Number(c.getRange(r,2).getValue())||0, mes:String(c.getRange(r,3).getValue()), pct:Number(c.getRange(r,4).getValue())||0});
    }
  }
  if(!filas.length){ responder(chatId,'No hay certificados de '+obra+' para corregir.'); return; }
  filas.sort(function(a,b){ return a.n-b.n; });
  var idx=-1;
  for(var i=0;i<filas.length;i++){ if(filas[i].mes.toLowerCase().indexOf(mesTxt.toLowerCase())>=0) idx=i; }
  if(idx<0){ responder(chatId,'No encontré un certificado de "'+mesTxt+'" en '+obra+'.'); return; }
  filas[idx].pct=nuevoPct;
  var p=certParams(obra), factor=1, mv=ss().getSheetByName('Movimientos'), mvv=mv.getDataRange().getValues();
  filas.forEach(function(f){
    factor*=(1+f.pct);
    var acum=factor-1, reaj=Math.round(p.hon*acum), total=p.hon+reaj;
    c.getRange(f.row,4).setValue(f.pct);
    c.getRange(f.row,5).setValue(acum);
    c.getRange(f.row,6).setValue(reaj);
    c.getRange(f.row,7).setValue(total);
    for(var j=mvv.length-1;j>=1;j--){
      if(mvv[j][1]==='Ingreso'&&mvv[j][2]==='Honorarios'&&mvv[j][10]===obra&&String(mvv[j][3]).toLowerCase().indexOf(f.mes.toLowerCase())>=0){ mv.getRange(j+1,6).setValue(total); break; }
    }
  });
  var fx=filas[idx];
  responder(chatId,'✅ Corregí '+obra+' · '+fx.mes+' → IPC '+pctTxt(nuevoPct)+'.\nRecalculé en cadena '+filas.length+' certificado(s) y sus honorarios. Si querés el PDF actualizado de alguno, decime "PDF certificado '+fx.mes+'".');
}
function menuCorregir(chatId){
  var obra=obraActiva(chatId);
  var c=ss().getSheetByName('Certificados'), last=c.getLastRow(), kb=[];
  for(var r=18;r<=last;r++){ if(c.getRange(r,1).getValue()===obra){ kb.push([{text:'✏️ '+mesLindo(c.getRange(r,3).getValue())}]); } }
  if(!kb.length){ tgApi('sendMessage',{chat_id:chatId,text:'No hay certificados en '+obra+' para corregir. Si querés otra obra, tocá 🏗 y elegila.',reply_markup:{keyboard:[[{text:'☰ Menú'}]],resize_keyboard:true,is_persistent:true}}); return; }
  kb.push([{text:'⬅️ Volver'},{text:'☰ Menú'}]);
  tgApi('sendMessage',{chat_id:chatId,text:'Corrigiendo certificados de '+obra+'. ¿Qué mes querés corregir?',reply_markup:{keyboard:kb,resize_keyboard:true,is_persistent:true}});
}
function menuElegirObra(chatId){
  var obras=obrasLista().filter(function(o){ return o!=='Estudio'; });
  var kb=obras.map(function(o){ return [{text:'👉 '+o}]; });
  kb.push([{text:'⬅️ Volver'},{text:'☰ Menú'}]);
  tgApi('sendMessage',{chat_id:chatId,text:'¿En qué obra querés trabajar?',reply_markup:{keyboard:kb,resize_keyboard:true,is_persistent:true}});
}
function mesLindo(v){
  var MESES=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];
  if(v instanceof Date) return MESES[v.getMonth()]+' '+v.getFullYear();
  var s=String(v||'');
  var d=new Date(s);
  if(s && !isNaN(d.getTime()) && /\d{4}/.test(s) && /GMT|:\d\d|\bT\d/.test(s)) return MESES[d.getMonth()]+' '+d.getFullYear();
  return s;
}
function arreglarMeses(){
  var c=ss().getSheetByName('Certificados'), last=c.getLastRow(), cambios=0;
  for(var r=18;r<=last;r++){
    var v=c.getRange(r,3).getValue();
    if(v===''||v==null) continue;
    var lindo=mesLindo(v);
    if(lindo!==String(v)){ c.getRange(r,3).setValue(lindo); cambios++; }
  }
  SpreadsheetApp.getActive().toast(cambios+' mes(es) normalizados');
}
function crearFinca(){
  var N='Finca';
  if(!matchObra(N)) agregarObra(N,'');
  var c=ss().getSheetByName('Certificados');
  var v=c.getRange(6,1,10,1).getValues(), f=0;
  for(var i=0;i<v.length;i++){ if(v[i][0]===N){f=6+i;break;} }
  if(!f){ for(var i=0;i<v.length;i++){ if(!v[i][0]){f=6+i;break;} } }
  c.getRange(f,1,1,4).setValues([[N,Math.round(31000000/18),18,new Date(2026,9,23)]]);
  c.getRange(f,2).setNumberFormat('"$"#,##0');
  c.getRange(f,4).setNumberFormat('dd/mm/yyyy');
  SpreadsheetApp.getActive().toast('Finca creada ✓');
}
