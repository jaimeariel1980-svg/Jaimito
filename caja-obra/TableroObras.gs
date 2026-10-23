// ===== TABLERO CAJA DE OBRA (reemplaza el doGet viejo) =====
// Opcional: clave de acceso -> correr ponerClaveTablero() una vez, y entrar con <url /exec>?clave=XXXX
function ponerClaveTablero(){ PropertiesService.getScriptProperties().setProperty('TABLERO_PIN','CAMBIAR_ESTA_CLAVE'); }

function doGet(e){
  var pin=PropertiesService.getScriptProperties().getProperty('TABLERO_PIN');
  if(pin && String((e&&e.parameter&&e.parameter.clave)||'')!==pin)
    return HtmlService.createHtmlOutput('<meta name="viewport" content="width=device-width,initial-scale=1"><p style="font-family:sans-serif;padding:24px;color:#880E4F">Acceso restringido. Entrá con el link que incluye la clave.</p>');
  var json=JSON.stringify(tableroDatos_()).replace(/</g,'\\u003c');
  return HtmlService.createHtmlOutput(TABLERO_HTML_.replace('__DATA__',function(){return json;}))
    .setTitle('NJ Arquitectura · Obras').addMetaTag('viewport','width=device-width, initial-scale=1');
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
    var pres=0, ejeP=0;
    orden.forEach(function(k){ if(rub[k].pres>0){ pres+=rub[k].pres; ejeP+=rub[k].ejecutado; } });
    return {
      obra:n, dir:o[1]||'', estado:o[5]||'', cobrado:cob, pagado:pag, saldo:cob-pag, bonus:bonus, honorarios:hon, gastoEstudio:est,
      neta:bonus+hon-est, pres:pres, avance:pres>0?ejeP/pres:null, certs:certs, params:p,
      rubros:orden.map(function(k){ return rub[k]; }).filter(function(x){ return x.pres||x.cobrado||x.ejecutado||x.bonus; })
    };
  });
  var movs=mv.slice(-400).reverse().map(function(r){
    return {f:f(r[0],'dd/MM/yy HH:mm'), tipo:r[1], rubro:r[2]||'', det:r[3]||'', prov:r[4]||'', monto:Number(r[5])||0, fac:r[6]||'', link:r[7]||'', obra:r[10]||''};
  });
  return {obras:obras, movs:movs, hoy:Utilities.formatDate(new Date(),TZ,'dd/MM/yyyy HH:mm')};
}

var TABLERO_HTML_ = '<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>NJ Arquitectura · Obras</title><style>'
+':root{--p:#AD1457;--pd:#880E4F;--bg:#FFF5F8;--bd:#F3C6D8;--mu:#9C6079;--tx:#3A0A24;--ok:#1A6645;--ng:#B00020}'
+'*{box-sizing:border-box}body{margin:0;font-family:-apple-system,Segoe UI,Roboto,sans-serif;background:var(--bg);color:var(--tx)}'
+'.hdr{background:var(--p);color:#fff;padding:14px 16px;font-size:18px;font-weight:800;display:flex;align-items:center;gap:10px}.hdr span{flex:1}'
+'.hdr small{font-weight:400;font-size:11px;opacity:.85}.btn{background:rgba(255,255,255,.92);color:var(--p);border:0;border-radius:6px;padding:6px 10px;font-weight:800;font-size:13px;cursor:pointer}'
+'.kpis{display:grid;grid-template-columns:repeat(auto-fit,minmax(130px,1fr));gap:10px;padding:12px 16px 0}'
+'.kpi{background:#fff;border:1px solid var(--bd);border-radius:10px;padding:10px 12px}.kpi i{display:block;font-style:normal;font-size:11px;color:var(--mu)}.kpi b{font-size:17px}'
+'.tabs{display:flex;gap:6px;padding:12px 16px 0;position:sticky;top:0;background:var(--bg);z-index:2}'
+'.tab{flex:1;text-align:center;padding:9px 6px;border:1px solid var(--bd);border-radius:8px;background:#fff;color:var(--pd);font-weight:700;font-size:13px;cursor:pointer}.tab.on{background:var(--p);color:#fff;border-color:var(--p)}'
+'.grid{display:grid;grid-template-columns:repeat(auto-fit,minmax(270px,1fr));gap:14px;padding:14px 16px}'
+'.card{background:#fff;border:1px solid var(--bd);border-left:5px solid #E91E63;border-radius:0 12px 12px 0;padding:14px}.card.c{cursor:pointer}'
+'.nom{font-size:17px;font-weight:800;color:var(--pd)}.dir{font-size:12px;color:var(--mu);margin-bottom:10px}'
+'.row{display:flex;justify-content:space-between;font-size:14px;margin:4px 0}.row span{color:var(--mu)}.sal b{color:var(--ok)}.neg{color:var(--ng)}'
+'.gan{display:flex;justify-content:space-between;font-size:16px;font-weight:800;color:var(--p);border-top:2px solid var(--bd);margin-top:8px;padding-top:8px}'
+'.bar{height:8px;background:#F8DCE7;border-radius:5px;overflow:hidden;margin:4px 0}.bar i{display:block;height:100%;background:#E91E63}.bar.x i{background:var(--ng)}'
+'.sub{font-size:11px;color:var(--mu);margin-top:4px}.pill{display:inline-block;font-size:11px;background:#FCE4EC;color:var(--pd);border-radius:10px;padding:2px 8px;margin-left:6px;font-weight:700}'
+'.wrap{padding:14px 16px;max-width:900px}.box{background:#fff;border:1px solid var(--bd);border-radius:10px;padding:12px;margin-bottom:14px;overflow-x:auto}.box h3{margin:0 0 8px;color:var(--pd);font-size:14px}'
+'.tbl{width:100%;border-collapse:collapse;font-size:13px}.tbl th{background:var(--p);color:#fff;font-size:12px;padding:7px;text-align:left}.tbl td{padding:7px;border-bottom:1px solid #F3D9E4;vertical-align:top}.tbl td.n,.tbl th.n{text-align:right;white-space:nowrap}'
+'.dl{display:flex;justify-content:space-between;font-size:14px;margin:4px 0}.dl.tot{font-weight:800;color:var(--p);border-top:2px solid var(--bd);margin-top:8px;padding-top:8px}'
+'.flt{display:flex;gap:8px;flex-wrap:wrap;padding:14px 16px 0}.flt select,.flt input{padding:8px;border:1px solid var(--bd);border-radius:8px;font-size:14px;background:#fff;flex:1;min-width:130px}'
+'.vacio{padding:24px;text-align:center;color:var(--mu)}a{color:var(--p)}'
+'</style></head><body><div class="hdr"><span id="ttl">NJ Arquitectura &middot; Obras<br><small id="upd"></small></span><button class="btn" onclick="location.reload()">&#8635;</button></div>'
+'<div class="kpis" id="kpis"></div><div class="tabs" id="tabs"></div><div id="app"></div>'
+'<script>var D=__DATA__;'
+'var T="obras",SEL=-1,FO=-1,FT="",FQ="";'
+'function m(n){n=Math.round(Number(n)||0);return (n<0?"-$":"$")+Math.abs(n).toLocaleString("es-AR");}'
+'function esc(s){return String(s==null?"":s).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");}'
+'function pc(x){return x==null?"":Math.round(x*100)+"%";}'
+'function pct4(x){return (Math.round(x*10000)/100).toString().replace(".",",")+" %";}'
+'function bar(x){if(x==null)return "";var w=Math.min(100,Math.max(0,x*100));return "<div class=\\"bar"+(x>1?" x":"")+"\\"><i style=\\"width:"+w+"%\\"></i></div>";}'
+'function sum(k){return D.obras.reduce(function(a,o){return a+(o[k]||0);},0);}'
+'function kpis(){var k=[["Cobrado",sum("cobrado")],["Pagado",sum("pagado")],["Saldo",sum("saldo")],["Ganancia neta",sum("neta")]];'
+'document.getElementById("kpis").innerHTML=k.map(function(x){return "<div class=\\"kpi\\"><i>"+x[0]+"</i><b"+(x[1]<0?" class=\\"neg\\"":"")+">"+m(x[1])+"</b></div>";}).join("");}'
+'function tabs(){var t=[["obras","Obras"],["certs","Certificados"],["movs","Movimientos"]];'
+'document.getElementById("tabs").innerHTML=t.map(function(x){return "<div class=\\"tab"+(T===x[0]?" on":"")+"\\" onclick=\\"go(\x27"+x[0]+"\x27)\\">"+x[1]+"</div>";}).join("");}'
+'function go(t){T=t;SEL=-1;draw();window.scrollTo(0,0);}'
+'function ver(i){SEL=i;draw();window.scrollTo(0,0);}'
+'function draw(){tabs();var a=document.getElementById("app");'
+'if(T==="obras")a.innerHTML=SEL>=0?vDetalle(D.obras[SEL]):vObras();'
+'else if(T==="certs")a.innerHTML=vCerts();else{a.innerHTML=vMovs();lstMovs();}}'
+'function vObras(){if(!D.obras.length)return "<div class=\\"vacio\\">Todavia no hay obras.</div>";'
+'return "<div class=\\"grid\\">"+D.obras.map(function(o,i){var cp=o.params&&o.params.plazo?" &middot; cert. "+o.certs.length+"/"+o.params.plazo:"";'
+'return "<div class=\\"card c\\" onclick=\\"ver("+i+")\\"><div class=\\"nom\\">"+esc(o.obra)+"<span class=\\"pill\\">"+esc(o.estado)+"</span></div><div class=\\"dir\\">"+esc(o.dir)+"</div>"'
+'+"<div class=\\"row\\"><span>Cobrado</span><b>"+m(o.cobrado)+"</b></div><div class=\\"row\\"><span>Pagado</span><b>"+m(o.pagado)+"</b></div>"'
+'+"<div class=\\"row sal\\"><span>Saldo</span><b>"+m(o.saldo)+"</b></div>"'
+'+(o.avance==null?"":"<div class=\\"row\\"><span>Avance vs presupuesto</span><b>"+pc(o.avance)+"</b></div>"+bar(o.avance))'
+'+"<div class=\\"gan\\"><span>Ganancia neta</span><b>"+m(o.neta)+"</b></div><div class=\\"sub\\">Toca para ver el detalle"+cp+"</div></div>";}).join("")+"</div>";}'
+'function vDetalle(o){var h="<div class=\\"wrap\\"><button class=\\"btn\\" style=\\"background:var(--p);color:#fff;margin-bottom:12px\\" onclick=\\"ver(-1)\\">&larr; Obras</button>";'
+'h+="<div class=\\"nom\\" style=\\"font-size:20px\\">"+esc(o.obra)+"</div><div class=\\"dir\\">"+esc(o.dir)+"</div>";'
+'h+="<div class=\\"box\\"><h3>Desglose de ganancia</h3>";'
+'h+="<div class=\\"dl\\"><span>Caja (cobrado - pagado)</span><b>"+m(o.saldo)+"</b></div><div class=\\"dl\\"><span>Bonus</span><b>"+m(o.bonus)+"</b></div>";'
+'h+="<div class=\\"dl\\"><span>Honorarios</span><b>"+m(o.honorarios)+"</b></div><div class=\\"dl\\"><span>Gastos de estudio</span><b class=\\"neg\\">-"+m(o.gastoEstudio)+"</b></div>";'
+'h+="<div class=\\"dl tot\\"><span>Ganancia neta</span><b>"+m(o.neta)+"</b></div></div>";'
+'h+="<div class=\\"box\\"><h3>Rubros</h3><table class=\\"tbl\\"><tr><th>Rubro</th><th class=\\"n\\">Presup.</th><th class=\\"n\\">Cobrado</th><th class=\\"n\\">Ejecutado</th><th class=\\"n\\">Avance</th><th class=\\"n\\">Bonus</th></tr>";'
+'if(!o.rubros.length)h+="<tr><td colspan=6 class=vacio>Sin movimientos ni presupuesto.</td></tr>";'
+'o.rubros.forEach(function(r){var av=r.pres>0?r.ejecutado/r.pres:null;h+="<tr><td>"+esc(r.rubro)+"</td><td class=n>"+(r.pres?m(r.pres):"")+"</td><td class=n>"+m(r.cobrado)+"</td><td class=n>"+m(r.ejecutado)+"</td><td class=n>"+(av==null?"":pc(av)+bar(av))+"</td><td class=n>"+m(r.bonus)+"</td></tr>";});'
+'h+="</table></div>";h+=boxCerts(o);'
+'var u=D.movs.filter(function(x){return x.obra===o.obra;}).slice(0,10);'
+'h+="<div class=\\"box\\"><h3>Ultimos movimientos</h3>"+tblMovs(u)+"</div></div>";return h;}'
+'function boxCerts(o){var p=o.params,h="<div class=\\"box\\"><h3>Certificados de honorarios</h3>";'
+'if(!p||!p.mensual)return h+"<div class=\\"sub\\">Falta cargar el honorario mensual en la pestana Certificados.</div></div>";'
+'h+="<div class=\\"dl\\"><span>Honorario mensual</span><b>"+m(p.mensual)+"</b></div><div class=\\"dl\\"><span>Plazo</span><b>"+p.plazo+" meses</b></div><div class=\\"dl\\"><span>Fecha base</span><b>"+esc(p.base)+"</b></div>";'
+'h+="<div class=\\"dl\\"><span>Emitidos</span><b>"+o.certs.length+" / "+p.plazo+"</b></div>"+bar(p.plazo?o.certs.length/p.plazo:null);'
+'if(o.certs.length)h+="<table class=\\"tbl\\" style=\\"margin-top:8px\\"><tr><th>N&deg;</th><th>Mes</th><th class=\\"n\\">IPC</th><th class=\\"n\\">Acum.</th><th class=\\"n\\">Total</th></tr>"+o.certs.map(function(c){return "<tr><td>"+c.n+"</td><td>"+esc(c.mes)+"</td><td class=n>"+pct4(c.pct)+"</td><td class=n>"+pct4(c.acum)+"</td><td class=n>"+m(c.total)+"</td></tr>";}).join("")+"</table>";'
+'return h+"</div>";}'
+'function vCerts(){if(!D.obras.length)return "<div class=\\"vacio\\">Todavia no hay obras.</div>";'
+'return "<div class=\\"wrap\\">"+D.obras.map(function(o){return "<div class=\\"nom\\" style=\\"margin:6px 0\\">"+esc(o.obra)+"</div>"+boxCerts(o);}).join("")+"</div>";}'
+'function tblMovs(l){if(!l.length)return "<div class=\\"vacio\\">Sin movimientos.</div>";'
+'return "<table class=\\"tbl\\"><tr><th>Fecha</th><th>Tipo</th><th>Rubro / detalle</th><th class=\\"n\\">Monto</th></tr>"+l.map(function(x){'
+'var lk=x.link?" <a href=\\""+esc(x.link)+"\\" target=\\"_blank\\" rel=\\"noopener\\">&#128206;</a>":"";'
+'return "<tr><td>"+esc(x.f)+"</td><td>"+esc(x.tipo)+"</td><td>"+esc(x.rubro)+"<div class=\\"sub\\">"+esc(x.det)+(x.prov?" &middot; "+esc(x.prov):"")+lk+"</div></td><td class=n"+(x.tipo==="Ingreso"?"":" style=\\"color:var(--ng)\\"")+">"+(x.tipo==="Ingreso"?"":"-")+m(x.monto)+"</td></tr>";}).join("")+"</table>";}'
+'function vMovs(){return "<div class=\\"flt\\"><select onchange=\\"FO=+this.value;lstMovs()\\"><option value=\\"-1\\">Todas las obras</option>"+D.obras.map(function(o,i){return "<option value=\\""+i+"\\""+(FO===i?" selected":"")+">"+esc(o.obra)+"</option>";}).join("")+"</select>"'
+'+"<select onchange=\\"FT=this.value;lstMovs()\\"><option value=\\"\\">Todos los tipos</option>"+["Ingreso","Pago sub","Compra","Gasto"].map(function(t){return "<option"+(FT===t?" selected":"")+">"+t+"</option>";}).join("")+"</select>"'
+'+"<input placeholder=\\"Buscar proveedor, detalle, rubro...\\" value=\\""+esc(FQ)+"\\" oninput=\\"FQ=this.value;lstMovs()\\"></div><div class=\\"wrap\\"><div class=\\"box\\" id=\\"lm\\"></div></div>";}'
+'function lstMovs(){var q=FQ.toLowerCase(),on=FO>=0?D.obras[FO].obra:"";var l=D.movs.filter(function(x){return (!on||x.obra===on)&&(!FT||x.tipo===FT)&&(!q||(x.prov+" "+x.det+" "+x.rubro).toLowerCase().indexOf(q)>=0);});'
+'var t=l.reduce(function(a,x){return a+(x.tipo==="Ingreso"?x.monto:-x.monto);},0);'
+'document.getElementById("lm").innerHTML="<div class=\\"sub\\" style=\\"margin-bottom:6px\\">"+l.length+" movimientos &middot; neto "+m(t)+" (ultimos 400)</div>"+tblMovs(l.slice(0,150));}'
+'document.getElementById("upd").textContent="Actualizado "+D.hoy;kpis();draw();'
+'<\/script></body></html>';
