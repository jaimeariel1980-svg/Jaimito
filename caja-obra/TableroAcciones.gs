// ===== TABLERO CAJA DE OBRA: acciones (cargar, editar, borrar) =====
// Todas validan la clave en el servidor (propiedad TABLERO_PIN) y usan el mismo Sheet que Anita.

var TB_RUBROS_EXTRA_=['Honorarios','Sueldos','Oficina','Servicios','Contadora','Impuestos'];
var TB_TIPOS_={'Ingreso':1,'Pago sub':1,'Compra':1,'Gasto':1};

function tbAuth_(clave){
  var pin=PropertiesService.getScriptProperties().getProperty('TABLERO_PIN');
  if(!pin||pin==='CAMBIAR_ESTA_CLAVE') throw new Error('SIN_CLAVE');
  var ca=CacheService.getScriptCache(), f=Number(ca.get('tb_fallos'))||0;
  if(f>=20) throw new Error('BLOQUEADO');
  if(String(clave||'')!==pin){ ca.put('tb_fallos',String(f+1),600); throw new Error('CLAVE'); }
}
function tbLock_(fn){
  var l=LockService.getScriptLock(); l.waitLock(15000);
  try{ return fn(); } finally{ l.releaseLock(); }
}
function tbTxt_(v,max){
  v=String(v==null?'':v).replace(/[\r\n]+/g,' ').trim().slice(0,max||120);
  return /^[=+\-@]/.test(v)?"'"+v:v;
}
function tbFecha_(s){
  if(/^\d{4}-\d{2}-\d{2}$/.test(String(s||''))) return Utilities.parseDate(s+' '+Utilities.formatDate(new Date(),'GMT-3','HH:mm'),'GMT-3','yyyy-MM-dd HH:mm');
  return new Date();
}

function tbDatos(clave){ tbAuth_(clave); return tableroDatos_(); }

function tbGuardarMov(clave,m){
  tbAuth_(clave);
  return tbLock_(function(){
    var mv=ss().getSheetByName('Movimientos');
    var obra=String(m.obra||'');
    if(obrasLista().indexOf(obra)<0) throw new Error('Esa obra no existe.');
    var tipo=String(m.tipo||''); if(!TB_TIPOS_[tipo]) throw new Error('Tipo inválido.');
    var monto=Math.round(Number(m.monto)||0);
    if(!(monto>0)) throw new Error('El monto tiene que ser mayor a 0.');
    var rubro=RUBROS.concat(TB_RUBROS_EXTRA_).indexOf(m.rubro)>=0?m.rubro:'General';
    var bonus=(tipo==='Compra'||tipo==='Pago sub')?Math.max(0,Math.min(1,Number(m.bonus)||0)):0;
    var gb=Math.round(monto*bonus), det=tbTxt_(m.concepto,150), prov=tbTxt_(m.proveedor,80);
    var fac=m.factura==='Sí'?'Sí':'No', dup=false;
    if(m.row){
      var row=Number(m.row);
      if(!(row>=2&&row<=mv.getLastRow())) throw new Error('CAMBIO');
      var r=mv.getRange(row,1,1,13).getValues()[0];
      if(tbFp_(r)!==m.fp) throw new Error('CAMBIO');
      var fecha=(r[0] instanceof Date&&Utilities.formatDate(r[0],'GMT-3','yyyy-MM-dd')===m.fecha)?r[0]:tbFecha_(m.fecha);
      mv.getRange(row,1,1,7).setValues([[fecha,tipo,rubro,det,prov,monto,fac]]);
      mv.getRange(row,11,1,3).setValues([[obra,bonus>0?bonus:'',bonus>0?gb:'']]);
    }else{
      if(prov){ var n=normal(prov); movimientos().forEach(function(x){ if((Number(x[5])||0)===monto&&normal(x[4])===n&&x[10]===obra) dup=true; }); }
      mv.appendRow([tbFecha_(m.fecha),tipo,rubro,det,prov,monto,fac,String(m.link||'').slice(0,300),'Tablero','Por confirmar',obra,bonus>0?bonus:'',bonus>0?gb:'']);
    }
    SpreadsheetApp.flush();
    return {dup:dup,datos:tableroDatos_()};
  });
}

function tbBorrarMov(clave,row,fp){
  tbAuth_(clave);
  return tbLock_(function(){
    var s=ss(), mv=s.getSheetByName('Movimientos'); row=Number(row);
    if(!(row>=2&&row<=mv.getLastRow())) throw new Error('CAMBIO');
    var r=mv.getRange(row,1,1,13).getValues()[0];
    if(tbFp_(r)!==fp) throw new Error('CAMBIO');
    var pp=s.getSheetByName('Papelera');
    if(!pp){
      pp=s.insertSheet('Papelera');
      pp.getRange(1,1,1,15).setValues([['Borrado','Por','Fecha','Tipo','Rubro','Detalle','Proveedor','Monto','Factura','Comprobante','Origen','Estado','Obra','Bonus %','Ganancia bonus']]).setFontWeight('bold');
      pp.setFrozenRows(1);
    }
    pp.appendRow([new Date(),'Tablero'].concat(r));
    mv.deleteRow(row);
    SpreadsheetApp.flush();
    return {datos:tableroDatos_(),borrado:{obra:r[10],tipo:r[1],rubro:r[2],concepto:r[3],proveedor:r[4],monto:Number(r[5])||0,factura:r[6],link:r[7],bonus:Number(r[11])||0,fecha:(r[0] instanceof Date)?Utilities.formatDate(r[0],'GMT-3','yyyy-MM-dd'):''}};
  });
}

function tbPresupuestos(clave,obra,mapa){
  tbAuth_(clave);
  return tbLock_(function(){
    obra=String(obra||'');
    if(obrasLista().indexOf(obra)<0) throw new Error('Esa obra no existe.');
    var rb=ss().getSheetByName('Rubros'), last=rb.getLastRow();
    var v=last>=2?rb.getRange(2,1,last-1,3).getValues():[], idx={};
    v.forEach(function(r,i){ if(r[0]===obra) idx[r[1]]=i; });
    var validos=RUBROS.concat(TB_RUBROS_EXTRA_), nuevos=[];
    Object.keys(mapa||{}).forEach(function(k){
      if(validos.indexOf(k)<0) return;
      var n=Math.max(0,Math.round(Number(mapa[k])||0)), val=n>0?n:'';
      if(idx[k]!==undefined) v[idx[k]][2]=val; else if(val!=='') nuevos.push([obra,k,val]);
    });
    if(v.length) rb.getRange(2,3,v.length,1).setValues(v.map(function(r){ return [r[2]]; }));
    if(nuevos.length){ rb.getRange(rb.getLastRow()+1,1,nuevos.length,3).setValues(nuevos); rb.getRange(2,3,rb.getLastRow()-1,1).setNumberFormat('"$"#,##0'); }
    SpreadsheetApp.flush();
    return tableroDatos_();
  });
}

function tbNuevaObra(clave,nombre,direccion){
  tbAuth_(clave);
  return tbLock_(function(){
    nombre=tbTxt_(nombre,60); direccion=tbTxt_(direccion,120);
    if(nombre.length<2) throw new Error('Poné un nombre para la obra.');
    var n=normal(nombre);
    if(obrasLista().some(function(o){ return normal(o)===n; })) throw new Error('Esa obra ya existe.');
    agregarObra(nombre,direccion);
    SpreadsheetApp.flush();
    return tableroDatos_();
  });
}
