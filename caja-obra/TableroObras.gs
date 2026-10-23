// ===== TABLERO CAJA DE OBRA (reemplaza el doGet viejo) =====
// Opcional: clave de acceso -> correr ponerClaveTablero() una vez, y entrar con <url /exec>?clave=XXXX
function ponerClaveTablero(){ PropertiesService.getScriptProperties().setProperty('TABLERO_PIN','CAMBIAR_ESTA_CLAVE'); }

function doGet(e){
  var pin=PropertiesService.getScriptProperties().getProperty('TABLERO_PIN');
  if(pin && String((e&&e.parameter&&e.parameter.clave)||'')!==pin)
    return HtmlService.createHtmlOutput('<meta name="viewport" content="width=device-width,initial-scale=1"><p style="font-family:sans-serif;padding:24px;color:#880E4F">Acceso restringido. Entrá con el link que incluye la clave.</p>');
  var t=HtmlService.createTemplateFromFile('TableroObras');
  t.datos=JSON.stringify(tableroDatos_()).replace(/</g,'\\u003c');
  return t.evaluate().setTitle('NJ Arquitectura · Obras').addMetaTag('viewport','width=device-width, initial-scale=1');
}

function tableroDatos_(){
  var s=ss(), TZ='GMT-3';
  var TIPOS={'Ingreso':1,'Pago sub':1,'Compra':1,'Gasto':1};
  var EST=['Sueldos','Oficina','Servicios','Contadora','Impuestos'];
  function vals(name,r,c,minRow,ncol){
    var sh=s.getSheetByName(name); if(!sh) return [];
    var last=sh.getLastRow(); if(last<minRow) return [];
    return sh.getRange(minRow,1,last-minRow+1,ncol).getValues();
  }
  var mv=vals('Movimientos',0,0,1,13).filter(function(r){ return TIPOS[r[1]]; });
  var ob=vals('Obras',0,0,4,6).filter(function(r){ return r[0]; });
  var rb=vals('Rubros',0,0,2,3);
  var ce=vals('Certificados',0,0,18,8).filter(function(r){ return r[0]; });
  var csh=s.getSheetByName('Certificados');
  var par=csh?csh.getRange(6,1,10,4).getValues():[];
  function f(d,fmt){ return (d instanceof Date)?Utilities.formatDate(d,TZ,fmt):''; }

  var obras=ob.map(function(o){
    var n=o[0], rub={}, orden=[];
    function R(k){ if(!rub[k]){ rub[k]={rubro:k,pres:0,cobrado:0,ejecutado:0,bonus:0}; orden.push(k); } return rub[k]; }
    rb.forEach(function(r){ if(r[0]===n&&r[1]) R(r[1]).pres+=Number(r[2])||0; });
    var cob=0,pag=0,bonus=0,est=0;
    mv.forEach(function(r){
      if(r[10]!==n) return;
      var mo=Number(r[5])||0, k=r[2]||'General', g=Number(r[12])||0, x=R(k);
      if(r[1]==='Ingreso'){ cob+=mo; x.cobrado+=mo; } else { pag+=mo; x.ejecutado+=mo; if(EST.indexOf(k)>=0) est+=mo; }
      x.bonus+=g; bonus+=g;
    });
    var certs=ce.filter(function(r){ return r[0]===n; }).map(function(r){
      return {n:Number(r[1])||0, mes:mesLindo(r[2]), pct:Number(r[3])||0, acum:Number(r[4])||0, reaj:Number(r[5])||0, total:Number(r[6])||0, fecha:f(r[7],'dd/MM/yy')};
    });
    var hon=certs.reduce(function(a,c){ return a+c.total; },0);
    var p=null;
    for(var i=0;i<par.length;i++){ if(par[i][0]===n){ p={mensual:Number(par[i][1])||0, plazo:Number(par[i][2])||0, base:f(par[i][3],'dd/MM/yyyy')}; break; } }
    var ahora=new Date(), MS=['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'], serie=[];
    for(var q=5;q>=0;q--){ var dm=new Date(ahora.getFullYear(),ahora.getMonth()-q,1); serie.push({k:dm.getFullYear()*12+dm.getMonth(),l:MS[dm.getMonth()],c:0,p:0}); }
    mv.forEach(function(r){
      if(r[10]!==n||!(r[0] instanceof Date)) return;
      var k=r[0].getFullYear()*12+r[0].getMonth(), mo=Number(r[5])||0;
      serie.forEach(function(x){ if(x.k===k){ if(r[1]==='Ingreso') x.c+=mo; else x.p+=mo; } });
    });
    var pres=0, ejeP=0;
    orden.forEach(function(k){ if(rub[k].pres>0){ pres+=rub[k].pres; ejeP+=rub[k].ejecutado; } });
    return {
      obra:n, dir:o[1]||'', estado:o[5]||'', cobrado:cob, pagado:pag, saldo:cob-pag, bonus:bonus, honorarios:hon, gastoEstudio:est,
      neta:bonus+hon-est, serie:serie, pres:pres, avance:pres>0?ejeP/pres:null, certs:certs, params:p,
      rubros:orden.map(function(k){ return rub[k]; }).filter(function(x){ return x.pres||x.cobrado||x.ejecutado||x.bonus; })
    };
  });
  var movs=mv.slice(-400).reverse().map(function(r){
    return {f:f(r[0],'dd/MM/yy HH:mm'), tipo:r[1], rubro:r[2]||'', det:r[3]||'', prov:r[4]||'', monto:Number(r[5])||0, fac:r[6]||'', link:r[7]||'', obra:r[10]||''};
  });
  return {obras:obras, movs:movs, hoy:Utilities.formatDate(new Date(),TZ,'dd/MM/yyyy HH:mm')};
}
