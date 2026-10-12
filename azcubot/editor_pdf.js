/*ED_PDF_INI*/
var EdPDF=(function(){
  var SRC={pdflib:'https://cdnjs.cloudflare.com/ajax/libs/pdf-lib/1.17.1/pdf-lib.min.js',pdfjs:'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js',worker:'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js'};
  var CARGAS={},S=null,UID=0,COL=['#2b6cb0','#2f855a','#b7791f','#9b2c2c','#6b46c1','#2c7a7b','#b83280','#4a5568','#c05621','#2a4365','#276749','#97266d','#744210','#553c9a'];
  function esc(s){return String(s==null?'':s).replace(/[&<>"]/g,function(c){return {'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c];});}
  function lib(src,ok){if(ok())return Promise.resolve();return CARGAS[src]||(CARGAS[src]=new Promise(function(res,rej){var s=document.createElement('script');s.src=src;s.onload=function(){res();};s.onerror=function(){delete CARGAS[src];rej(new Error('No pude cargar el editor de PDF. Revisá la conexión.'));};document.head.appendChild(s);}));}
  function cargar(){
    return Promise.all([lib(SRC.pdflib,function(){return window.PDFLib;}),lib(SRC.pdfjs,function(){return window.pdfjsLib;})]).then(function(){
      if(pdfjsLib.GlobalWorkerOptions.workerSrc)return;
      return fetch(SRC.worker).then(function(r){return r.text();}).then(function(t){pdfjsLib.GlobalWorkerOptions.workerSrc=URL.createObjectURL(new Blob([t],{type:'application/javascript'}));});
    });
  }
  function estilos(){
    if(document.getElementById('edp-css'))return;
    var c=document.createElement('style');c.id='edp-css';
    c.textContent='#edp{position:fixed;inset:0;z-index:99990;background:#eef1f6;display:flex;flex-direction:column;font-family:Segoe UI,Helvetica,Arial,sans-serif;color:#1f2a44}'+
    '#edp .h{background:#22386a;color:#fff;padding:10px 14px;display:flex;gap:8px;align-items:center;flex-wrap:wrap}#edp .h b{font-size:16px;flex:1;min-width:160px}'+
    '#edp .tb{background:#fff;border-bottom:1px solid #dfe3ea;padding:8px 14px;display:flex;gap:6px;align-items:center;flex-wrap:wrap}'+
    '#edp button,#edp select{font:600 13px inherit;font-family:inherit;border:1px solid #c9d0dc;background:#fff;color:#22386a;border-radius:8px;padding:7px 11px;cursor:pointer}'+
    '#edp button.p{background:#b5673e;border-color:#b5673e;color:#fff}#edp .h button{border-color:rgba(255,255,255,.4);background:rgba(255,255,255,.12);color:#fff}#edp .h button.p{background:#b5673e;border-color:#b5673e}#edp button:disabled{opacity:.5;cursor:default}'+
    '#edp .g{flex:1;overflow:auto;padding:12px;display:grid;grid-template-columns:repeat(auto-fill,minmax(150px,1fr));gap:10px;align-content:start}'+
    '#edp .c{position:relative;background:#fff;border:2px solid #cbd2de;border-radius:10px;padding:6px;cursor:grab;user-select:none}#edp .c.s{box-shadow:0 0 0 3px #f5b041}#edp .c.o{outline:3px dashed #22386a}'+
    '#edp .im{height:118px;display:flex;align-items:center;justify-content:center;overflow:hidden;background:#f3f5f9;border-radius:6px}#edp .im img{max-width:100%;max-height:118px;transition:transform .15s}#edp .im i{font-size:11px;color:#8a93a5;font-style:normal}'+
    '#edp .n{position:absolute;top:3px;left:5px;background:#22386a;color:#fff;border-radius:8px;font-size:10px;padding:0 6px}#edp .k{position:absolute;top:3px;right:5px}'+
    '#edp .a{display:flex;gap:2px;margin-top:5px;justify-content:space-between}#edp .a button{padding:2px 0;font-size:12px;border-radius:6px;flex:1}'+
    '#edp .t{font-size:10.5px;font-weight:700;color:#fff;border-radius:6px;padding:1px 5px;margin-top:4px;text-align:center;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}'+
    '#edp .em{grid-column:1/-1;text-align:center;color:#6b7280;padding:50px 10px;border:2px dashed #b9c1d1;border-radius:14px;background:#fff}'+
    '#edp .rs{padding:6px 14px;font-size:12.5px;color:#4a5568;background:#fff;border-top:1px solid #dfe3ea}'+
    '#edp .big{display:flex;flex-direction:column;gap:4px;align-items:center;width:210px;padding:16px 12px;font-size:28px;border-radius:14px}#edp .big b{font-size:15px}#edp .big small{font-size:12px;color:#6b7280;font-weight:400;font-style:normal}#edp .rg input{width:58px;padding:6px;border:1px solid #c9d0dc;border-radius:8px}'+
    '#edp .tb select{max-width:340px;min-width:200px}'+
    '#edp .hp{background:#fff8e1;border-bottom:1px solid #f0dca0;padding:8px 14px;font-size:13px;color:#5a4a10;line-height:1.45}#edp .im{cursor:zoom-in}#edp .k{width:22px;height:22px;cursor:pointer}#edp .sel{background:#e8f0fe;border-bottom:1px solid #c7d7f5;padding:8px 14px;display:flex;gap:8px;align-items:center;flex-wrap:wrap;font-size:13px}'+
    '#edpv{position:fixed;inset:0;z-index:99995;background:rgba(10,20,40,.85);display:flex;flex-direction:column;align-items:center;padding:10px;gap:8px}#edpv .vb{display:flex;gap:6px;align-items:center;flex-wrap:wrap;color:#fff;font-size:13px}#edpv .vb button{font:600 13px inherit;font-family:inherit;border:1px solid rgba(255,255,255,.4);background:rgba(255,255,255,.14);color:#fff;border-radius:8px;padding:7px 12px;cursor:pointer}#edpv .vc{flex:1;overflow:auto;display:flex;align-items:flex-start;justify-content:center;width:100%}#edpv canvas,#edpv img{max-width:100%;background:#fff;border-radius:6px}';
    document.head.appendChild(c);
  }
  function nuevaId(){return 'p'+(++UID);}
  function setMsg(t){var m=document.getElementById('edp-msg');if(m)m.textContent=t||'';}
  function jpegDe(f){return new Promise(function(ok,ko){var u=URL.createObjectURL(f),im=new Image();im.onload=function(){var k=Math.min(1,1600/Math.max(im.naturalWidth,im.naturalHeight)),c=document.createElement('canvas');c.width=Math.round(im.naturalWidth*k);c.height=Math.round(im.naturalHeight*k);var x=c.getContext('2d');x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);x.drawImage(im,0,0,c.width,c.height);URL.revokeObjectURL(u);var th=document.createElement('canvas');var kk=150/Math.max(c.width,c.height);th.width=Math.round(c.width*kk);th.height=Math.round(c.height*kk);th.getContext('2d').drawImage(c,0,0,th.width,th.height);c.toBlob(function(b){b.arrayBuffer().then(function(a){ok({b:new Uint8Array(a),w:c.width,h:c.height,th:th.toDataURL('image/jpeg',0.6),blob:b});});},'image/jpeg',0.78);};im.onerror=function(){URL.revokeObjectURL(u);ko(new Error('No pude leer «'+f.name+'».'));};im.src=u;});}
  async function agregar(files){
    var nuevos=[];
    try{await S.listo;}catch(e){return;}
    for(var f of files){
      if(!S)return;
      try{
        if(f.type==='application/pdf'||/\.pdf$/i.test(f.name)){
          var ab=await f.arrayBuffer(),js=await pdfjsLib.getDocument({data:ab.slice(0)}).promise,fu={t:'pdf',ab:ab,js:js,name:f.name};S.fu.push(fu);
          for(var i=0;i<js.numPages;i++){var pg={id:nuevaId(),f:fu,i:i,rot:0,box:S.modo==='unir'&&S.dest?S.dest:(S.o.base||''),th:''};S.pg.push(pg);nuevos.push(pg);}
        }else if(/^image\//.test(f.type)){
          var j=await jpegDe(f),fi={t:'img',j:j,name:f.name};S.fu.push(fi);var p2={id:nuevaId(),f:fi,i:0,rot:0,box:S.modo==='unir'&&S.dest?S.dest:(S.o.base||''),th:j.th};S.pg.push(p2);
        }else{alert('«'+f.name+'»: formato no soportado (usá PDF, JPG o PNG).');}
      }catch(e){alert('No pude leer «'+f.name+'»: '+(e&&e.message?e.message:e));}
    }
    render();
    for(var q of nuevos){await miniatura(q);}
  }
  async function miniatura(q){
    if(!S||q.th)return;
    try{
      var page=await q.f.js.getPage(q.i+1),v0=page.getViewport({scale:1}),v=page.getViewport({scale:150/Math.max(v0.width,v0.height)}),c=document.createElement('canvas');
      c.width=Math.round(v.width);c.height=Math.round(v.height);await page.render({canvasContext:c.getContext('2d'),viewport:v}).promise;q.th=c.toDataURL('image/jpeg',0.6);
      var im=document.querySelector('#edp [data-im="'+q.id+'"]');if(im)im.innerHTML='<img src="'+q.th+'" style="transform:rotate('+q.rot+'deg)">';
    }catch(e){}
  }
  async function construir(lista){
    var D=PDFLib.PDFDocument,out=await D.create();
    for(var q of lista){
      var pg;
      if(q.f.t==='pdf'){
        if(!q.f.lib)q.f.lib=await D.load(q.f.ab,{ignoreEncryption:true});
        pg=(await out.copyPages(q.f.lib,[q.i]))[0];out.addPage(pg);
      }else{
        var im=await out.embedJpg(q.f.j.b),k=Math.min(1,842/Math.max(q.f.j.w,q.f.j.h));
        pg=out.addPage([q.f.j.w*k,q.f.j.h*k]);pg.drawImage(im,{x:0,y:0,width:q.f.j.w*k,height:q.f.j.h*k});
      }
      if(q.rot)pg.setRotation(PDFLib.degrees(((pg.getRotation().angle||0)+q.rot)%360));
    }
    return out.save();
  }
  var LIM=5e6;
  async function compactarBytes(bytes,progreso){
    await cargar();
    var js=await pdfjsLib.getDocument({data:bytes.slice()}).promise,n=js.numPages,lado=n>20?1600:2000,D=PDFLib.PDFDocument,out=await D.create();
    for(var i=1;i<=n;i++){
      if(progreso)progreso('Optimizando el PDF… página '+i+' de '+n);
      var page=await js.getPage(i),v1=page.getViewport({scale:1}),k=lado/Math.max(v1.width,v1.height),v=page.getViewport({scale:Math.min(k,3)}),c=document.createElement('canvas');
      c.width=Math.round(v.width);c.height=Math.round(v.height);var x=c.getContext('2d');x.fillStyle='#fff';x.fillRect(0,0,c.width,c.height);
      await page.render({canvasContext:x,viewport:v}).promise;
      var blob=await new Promise(function(r){c.toBlob(r,'image/jpeg',0.72);}),im=await out.embedJpg(new Uint8Array(await blob.arrayBuffer())),pg=out.addPage([v1.width,v1.height]);
      pg.drawImage(im,{x:0,y:0,width:v1.width,height:v1.height});c.width=c.height=1;
    }
    var res=await out.save();return res.length<bytes.length?res:bytes;
  }
  async function compactarArchivo(f,progreso){
    if(!f||f.size<=LIM)return f;
    var by=await compactarBytes(new Uint8Array(await f.arrayBuffer()),progreso);
    return by.length<f.size?new File([by],f.name,{type:'application/pdf'}):f;
  }
  function u8b64(by){var s='';for(var i=0;i<by.length;i+=32768)s+=String.fromCharCode.apply(null,by.subarray(i,i+32768));return btoa(s);}
  function color(key){var i=S.slots.findIndex(function(s){return s.key===key;});return i<0?'#8a93a5':COL[i%COL.length];}
  function labelDe(key){var s=S.slots.filter(function(x){return x.key===key;})[0];return s?s.label:'';}
  function render(){
    if(!S)return;
    var g=document.getElementById('edp-g'),tb=document.getElementById('edp-tb'),rs=document.getElementById('edp-rs');
    var nSel=S.sel.size,con=S.slots.length>0;
    var base=S.o.base||'';
    var opts=S.slots.filter(function(x){return x.key!==base;}).map(function(x){return '<option value="'+esc(x.key)+'">'+esc(x.label)+'</option>';}).join('');
    var etq=base?'Sacar a otro documento…':'Mover a…';
    var ext=(base?'<option value="'+esc(base)+'">↩ Dejar en «'+esc(labelDe(base))+'»</option>':'<option value="__none">Sin asignar (no se guarda)</option>');
    var ayuda=base?'Seleccioná las páginas que querés <b>sacar</b> (tildá el cuadradito de cada una; tocá la miniatura para verla en grande) y elegí a qué documento van en <b>«Sacar a otro documento…»</b>. Las demás <b>quedan en «'+esc(labelDe(base))+'»</b>.':
      (con&&S.modo==='unir'?'<b>Unir:</b> elegí abajo el documento donde se guarda y agregá todas las fotos y PDFs que quieras con <b>+ Archivos</b>. Se guardan juntos en un solo PDF.':
      (con?'<b>Cómo se usa:</b> 1) tildá las páginas de un mismo documento (tocá la miniatura para verlas en grande), 2) elegí a cuál documento van en <b>«Mover a…»</b>, 3) repetí con los demás y tocá <b>Aplicar</b>. Las páginas «Sin asignar» no se guardan.':
      'Reordená arrastrando, rotá o quitá páginas. Tocá la miniatura para verla en grande. Todo se guarda como <b>un solo PDF</b>.'));
    document.getElementById('edp-help').innerHTML=S.pg.length?ayuda:'';
    document.getElementById('edp-help').style.display=S.pg.length?'':'none';
    tb.innerHTML='<button data-x="all">'+(nSel===S.pg.length&&nSel?'Quitar selección':'Seleccionar todo')+'</button>'+
      '<span style="font-size:12px;color:#6b7280">'+(nSel?'<b>'+nSel+' seleccionada(s):</b>':'Ninguna seleccionada')+'</span>'+
      (con?'<select data-x="asig" '+(nSel?'':'disabled')+'><option value="">'+etq+'</option>'+opts+ext+'</select>':'')+
      '<button data-x="rot" '+(nSel?'':'disabled')+'>↻ Rotar</button><button data-x="del" '+(nSel?'':'disabled')+'>🗑 Quitar</button>'+
      (con&&S.modo==='unir'?'<div style="width:100%;margin-top:4px"><b style="font-size:13px">🔗 Unir todo en:</b> <select data-x="dest"><option value="">Elegí el documento…</option>'+S.slots.map(function(x){return '<option value="'+esc(x.key)+'"'+(S.dest===x.key?' selected':'')+'>'+esc(x.label)+'</option>';}).join('')+'</select></div>':'')+
      (con&&!base&&S.modo!=='unir'?'<div class="rg" style="width:100%;margin-top:4px;display:flex;gap:6px;align-items:center;flex-wrap:wrap"><b style="font-size:13px">✂️ Separar por páginas:</b> del <input type="number" min="1" max="'+S.pg.length+'" data-x="r1" value="1"> al <input type="number" min="1" max="'+S.pg.length+'" data-x="r2" value="'+S.pg.length+'"> es <select data-x="rasig"><option value="">Elegí el documento…</option>'+S.slots.map(function(x){return '<option value="'+esc(x.key)+'">'+esc(x.label)+'</option>';}).join('')+'</select></div>':'');
    if(!S.pg.length&&(S.modo==='unir'||S.modo==='separar')){g.innerHTML='<div class="em">'+(S.modo==='unir'?'Elegí arriba en qué documento se guarda y tocá <b>+ Archivos</b>. Agregá todas las fotos y PDFs que quieras.':'Tocá <b>+ Archivos</b> y elegí el PDF que querés separar.')+'<br><br><button data-x="add">+ Archivos</button> <button data-x="m-ini">Volver</button></div>';rs.textContent='';return;}
    if(!S.pg.length){
      var bt='<button class="big" data-x="m-unir">🔗<b>Unir</b><small>Juntá fotos y PDFs en un solo '+(con?'documento':'PDF')+'</small></button>'+(con?'<button class="big" data-x="m-sep">✂️<b>Separar</b><small>Dividí un PDF en varios documentos (a mano o con IA)</small></button>':'')+'<button class="big" data-x="m-ord">📑<b>Ordenar / editar</b><small>Reordená, rotá o quitá páginas</small></button>';
      g.innerHTML='<div class="em"><div style="font-weight:700;font-size:16px;margin-bottom:14px;color:#22386a">¿Qué querés hacer?</div><div style="display:flex;gap:12px;justify-content:center;flex-wrap:wrap;margin-bottom:14px">'+bt+'</div><small>También podés arrastrar PDFs y fotos acá. Se pueden mezclar.</small></div>';rs.textContent='';return;}
    g.innerHTML=S.pg.map(function(q,n){
      return '<div class="c'+(S.sel.has(q.id)?' s':'')+'" draggable="true" data-id="'+q.id+'" style="border-color:'+(con&&q.box?color(q.box):'#cbd2de')+'">'+
        '<span class="n">'+(n+1)+'</span><input type="checkbox" class="k" data-k="'+q.id+'" '+(S.sel.has(q.id)?'checked':'')+'>'+
        '<div class="im" data-im="'+q.id+'" data-v="'+q.id+'">'+(q.th?'<img src="'+q.th+'" style="transform:rotate('+q.rot+'deg)">':'<i>cargando…</i>')+'</div>'+
        (con?'<div class="t" style="background:'+color(q.box)+'">'+esc(q.box?(q.box===base?'Queda: ':'→ ')+labelDe(q.box):'Sin asignar')+'</div>':'')+
        '<div class="a"><button data-m="-1" title="Mover antes">◀</button><button data-v="'+q.id+'" title="Ver en grande">🔍</button><button data-r="'+q.id+'" title="Rotar">↻</button><button data-d="'+q.id+'" title="Quitar">🗑</button><button data-m="1" title="Mover después">▶</button></div></div>';
    }).join('');
    if(con){
      var cnt={},sin=0;S.pg.forEach(function(q){if(q.box)cnt[q.box]=(cnt[q.box]||0)+1;else sin++;});
      rs.innerHTML=S.slots.filter(function(s){return cnt[s.key];}).map(function(s){return '<b style="color:'+color(s.key)+'">'+esc(s.label)+(s.key===base?' (queda)':'')+'</b>: '+cnt[s.key]+' pág.';}).join(' · ')+(sin?(cnt&&Object.keys(cnt).length?' · ':'')+'<b>Sin asignar</b>: '+sin+' (no se guardan)':'');
    }else rs.textContent=S.pg.length+' página(s) → se guardan como un solo PDF';
    if(con&&S.modo==='unir'&&!S.dest)rs.innerHTML='Elegí arriba en qué documento se guarda la unión.';
  }
  function mover(ids,antesDe){
    var set=new Set(ids),mov=S.pg.filter(function(q){return set.has(q.id);}),rest=S.pg.filter(function(q){return !set.has(q.id);});
    var at=antesDe?rest.findIndex(function(q){return q.id===antesDe;}):rest.length;if(at<0)at=rest.length;
    S.pg=rest.slice(0,at).concat(mov,rest.slice(at));
  }
  async function ver(id){
    var i=S.pg.findIndex(function(x){return x.id===id;});if(i<0)return;var q=S.pg[i];
    var o=document.getElementById('edpv');
    if(!o){o=document.createElement('div');o.id='edpv';document.body.appendChild(o);
      o.addEventListener('click',function(e){var b=e.target.closest('button');if(!b){return;}var x=b.dataset.vx,k=S.pg.findIndex(function(z){return z.id===o.dataset.id;});
        if(x==='x'){cerrarVista();return;}
        if(x==='prev'&&k>0){ver(S.pg[k-1].id);return;}
        if(x==='next'&&k<S.pg.length-1){ver(S.pg[k+1].id);return;}
        if(x==='sel'){var id3=S.pg[k].id;if(S.sel.has(id3))S.sel.delete(id3);else S.sel.add(id3);render();ver(id3);return;}
        if(x==='rot'){S.pg[k].rot=(S.pg[k].rot+90)%360;render();ver(S.pg[k].id);return;}
      });
      o._kd=function(e){if(!document.getElementById('edpv'))return;var k=S.pg.findIndex(function(z){return z.id===o.dataset.id;});if(e.key==='Escape')cerrarVista();else if(e.key==='ArrowLeft'&&k>0)ver(S.pg[k-1].id);else if(e.key==='ArrowRight'&&k<S.pg.length-1)ver(S.pg[k+1].id);};
      document.addEventListener('keydown',o._kd);
    }
    o.dataset.id=id;
    o.innerHTML='<div class="vb"><button data-vx="prev"'+(i===0?' disabled':'')+'>◀ Anterior</button><b>Página '+(i+1)+' de '+S.pg.length+'</b><button data-vx="next"'+(i===S.pg.length-1?' disabled':'')+'>Siguiente ▶</button><button data-vx="sel">'+(S.sel.has(id)?'☑ Seleccionada':'☐ Seleccionar esta página')+'</button><button data-vx="rot">↻ Rotar</button><button data-vx="x">✕ Cerrar</button></div><div class="vc" id="edpvc"><span style="color:#fff">Cargando…</span></div>';
    var cont=document.getElementById('edpvc');
    try{
      if(q.f.t==='img'){cont.innerHTML='<img src="'+URL.createObjectURL(q.f.j.blob)+'" style="transform:rotate('+q.rot+'deg);max-height:calc(100vh - 70px)">';return;}
      var page=await q.f.js.getPage(q.i+1),w=Math.min(window.innerWidth-30,900),v1=page.getViewport({scale:1,rotation:(page.rotate+q.rot)%360}),v=page.getViewport({scale:Math.min(w/v1.width,2),rotation:(page.rotate+q.rot)%360}),c=document.createElement('canvas');c.width=Math.round(v.width);c.height=Math.round(v.height);
      await page.render({canvasContext:c.getContext('2d'),viewport:v}).promise;
      if(document.getElementById('edpv')&&o.dataset.id===id){cont.innerHTML='';cont.appendChild(c);}
    }catch(e){cont.textContent='No pude mostrar la página.';}
  }
  function cerrarVista(){var v=document.getElementById('edpv');if(v){if(v._kd)document.removeEventListener('keydown',v._kd);v.remove();}}
  async function ia(){
    if(!S.pg.length){alert('Primero agregá archivos.');return;}
    var n=S.pg.length,T=8,btn=document.querySelector('#edp [data-x="ia"]');if(btn)btn.disabled=true;
    try{
      var tipos=[];
      for(var i=0;i<n;i+=T){
        setMsg('La IA está leyendo las páginas '+(i+1)+' a '+Math.min(i+T,n)+' de '+n+'…');
        var ch=S.pg.slice(i,i+T),by=await construir(ch);
        var r=await new Promise(function(ok,ko){google.script.run.withSuccessHandler(ok).withFailureHandler(function(e){ko(new Error(e&&e.message?e.message:e));}).clasificarPaginasLegajo({base64:u8b64(by),mimeType:'application/pdf',paginas:ch.length});});
        if(!r||!r.ok)throw new Error('La IA no pudo leer las páginas '+(i+1)+' a '+Math.min(i+T,n)+(r&&r.motivo?' ('+r.motivo+')':'')+'.');
        tipos=tipos.concat(r.tipos);
      }
      for(var k=1;k<n-1;k++){if(tipos[k]==='otro'&&tipos[k-1]===tipos[k+1]&&tipos[k-1]!=='otro')tipos[k]=tipos[k-1];}
      S.pg.forEach(function(q,idx){var s=S.slots.filter(function(x){return x.cls===tipos[idx];})[0];q.box=s?s.key:'';});
      var orden=S.slots.map(function(x){return x.key;});
      S.pg=S.pg.map(function(q,i){return {q:q,i:i,k:q.box?orden.indexOf(q.box):999};}).sort(function(a,b){return a.k-b.k||a.i-b.i;}).map(function(z){return z.q;});
      setMsg('Listo: la IA separó y juntó las páginas de cada documento. Revisá y corregí lo que haga falta.');render();
    }catch(e){setMsg('');alert('No pude separar con IA: '+(e&&e.message?e.message:e));}
    if(btn)btn.disabled=false;
  }
  async function aplicar(){
    var con=S.slots.length>0,res={};
    if(!S.pg.length){alert('No hay páginas.');return;}
    if(con){
      var sin=S.pg.filter(function(q){return !q.box;}).length;
      S.slots.forEach(function(s){var l=S.pg.filter(function(q){return q.box===s.key;});if(l.length)res[s.key]=l;});
      if(!Object.keys(res).length){alert('Ninguna página está asignada a un documento. Seleccioná páginas y usá «Mover a…» o «Separar y unir con IA».');return;}
      if(sin&&!confirm(sin+' página(s) sin asignar no se van a guardar. ¿Seguir?'))return;
    }else res._=S.pg.slice();
    var bt=document.querySelector('#edp [data-x="ok"]');if(bt)bt.disabled=true;
    try{
      setMsg('Armando los PDF…');var out={};
      for(var k of Object.keys(res)){var by=await construir(res[k]);if(by.length>LIM)by=await compactarBytes(by,setMsg);out[k]=new File([by],(k==='_'?'documento':k.replace(/[^\w\-]+/g,'_'))+'.pdf',{type:'application/pdf'});}
      await S.o.alAplicar(out,{base:S.o.base||''});cerrar();
    }catch(e){setMsg('');if(bt)bt.disabled=false;if(!(e&&e.silencio))alert('No se pudo aplicar: '+(e&&e.message?e.message:e));}
  }
  async function descargar(){
    var con=S.slots.length>0,res={};
    if(!S.pg.length){alert('No hay páginas.');return;}
    if(con){S.slots.forEach(function(s){var l=S.pg.filter(function(q){return q.box===s.key;});if(l.length)res[s.key]=l;});if(!Object.keys(res).length){alert('Asigná las páginas a un documento para descargarlas.');return;}}else res._=S.pg.slice();
    try{
      setMsg('Armando los PDF…');
      for(var k of Object.keys(res)){var by=await construir(res[k]);if(by.length>LIM)by=await compactarBytes(by,setMsg);var a=document.createElement('a');a.href=URL.createObjectURL(new Blob([by],{type:'application/pdf'}));a.download=(k==='_'?'documento':(labelDe(k)||k))+'.pdf';document.body.appendChild(a);a.click();a.remove();}
      setMsg('');
    }catch(e){setMsg('');alert('No pude armar el PDF: '+(e&&e.message?e.message:e));}
  }
  function cerrar(){var e=document.getElementById('edp');if(e)e.remove();cerrarVista();S=null;}
  function abrir(o){
    estilos();cerrar();
    S={o:o,fu:[],pg:[],sel:new Set(),slots:o.slots||[],drag:null};
    var el=document.createElement('div');el.id='edp';
    el.innerHTML='<div class="h"><b>'+esc(o.titulo||'Editor de PDF')+'</b><span id="edp-msg" style="font-size:12px;opacity:.9"></span>'+
      '<button data-x="add">+ Archivos</button>'+(S.slots.length?'<button data-x="ia" title="La IA separa las páginas por documento y junta las de cada uno">✨ Separar y unir con IA</button>':'')+'<button data-x="dl" title="Descargar el resultado en tu equipo">⬇ Descargar</button><button data-x="cancel">Cancelar</button><button class="p" data-x="ok">'+esc(o.textoOk||'Aplicar')+'</button></div>'+
      '<div class="hp" id="edp-help" style="display:none"></div><div class="tb" id="edp-tb"></div><div class="g" id="edp-g"></div><div class="rs" id="edp-rs"></div>'+
      '<input type="file" id="edp-f" accept=".pdf,application/pdf,image/*" multiple style="display:none">';
    document.body.appendChild(el);
    var fi=document.getElementById('edp-f');
    fi.onchange=function(){var fs=Array.from(fi.files||[]);fi.value='';if(fs.length)agregar(fs);};
    el.addEventListener('click',function(e){
      var t=e.target,b=t.closest('button,select'),x=b&&b.dataset.x;
      if(t.dataset&&t.dataset.k){var id=t.dataset.k;if(S.sel.has(id))S.sel.delete(id);else S.sel.add(id);render();return;}
      if(x==='add'){fi.click();return;}
      if(x==='m-unir'){S.modo='unir';render();fi.click();return;}
      if(x==='m-sep'){S.modo='separar';render();fi.click();return;}
      if(x==='m-ord'){S.modo=null;fi.click();return;}
      if(x==='dl'){descargar();return;}
      if(x==='m-ini'){S.modo=null;render();return;}
      if(x==='cancel'){if(!S.pg.length||confirm('¿Cerrar sin guardar?'))cerrar();return;}
      if(x==='ok'){aplicar();return;}
      if(x==='ia'){ia();return;}
      if(x==='all'){if(S.sel.size===S.pg.length)S.sel.clear();else S.pg.forEach(function(q){S.sel.add(q.id);});render();return;}
      if(x==='rot'){S.pg.forEach(function(q){if(S.sel.has(q.id))q.rot=(q.rot+90)%360;});render();return;}
      if(x==='del'){if(confirm('¿Quitar '+S.sel.size+' página(s)?')){S.pg=S.pg.filter(function(q){return !S.sel.has(q.id);});S.sel.clear();render();}return;}
      var c=t.closest('.c');if(!c)return;var id2=c.dataset.id;
      if(b&&b.dataset.d){S.pg=S.pg.filter(function(q){return q.id!==id2;});S.sel.delete(id2);render();return;}
      if(b&&b.dataset.r){var q=S.pg.filter(function(z){return z.id===id2;})[0];q.rot=(q.rot+90)%360;render();return;}
      if(b&&b.dataset.v){ver(id2);return;}
      if(!b&&t.closest('.im')){ver(id2);return;}
      if(b&&b.dataset.m){var i=S.pg.findIndex(function(z){return z.id===id2;}),j=i+(+b.dataset.m);if(j<0||j>=S.pg.length)return;var tmp=S.pg[i];S.pg[i]=S.pg[j];S.pg[j]=tmp;render();return;}
      if(S.sel.has(id2))S.sel.delete(id2);else S.sel.add(id2);render();
    });
    el.addEventListener('change',function(e){
      var t=e.target;
      if(t.dataset&&t.dataset.x==='dest'){S.dest=t.value;if(S.dest)S.pg.forEach(function(q){q.box=S.dest;});render();return;}
      if(t.dataset&&t.dataset.x==='rasig'){var v2=t.value;if(!v2)return;var r1=+document.querySelector('#edp [data-x="r1"]').value||1,r2=+document.querySelector('#edp [data-x="r2"]').value||S.pg.length;if(r1>r2){var tt=r1;r1=r2;r2=tt;}S.pg.forEach(function(q,i){if(i+1>=r1&&i+1<=r2)q.box=v2;});render();return;}
      if(t.dataset&&t.dataset.x==='asig'){var v=t.value;if(!v)return;S.pg.forEach(function(q){if(S.sel.has(q.id))q.box=v==='__none'?'':v;});S.sel.clear();render();}
    });
    el.addEventListener('dragstart',function(e){var c=e.target.closest&&e.target.closest('.c');if(!c)return;S.drag=c.dataset.id;try{e.dataTransfer.setData('text/plain',S.drag);e.dataTransfer.effectAllowed='move';}catch(x){}});
    el.addEventListener('dragover',function(e){e.preventDefault();var c=e.target.closest&&e.target.closest('.c');el.querySelectorAll('.c.o').forEach(function(z){z.classList.remove('o');});if(c&&S.drag)c.classList.add('o');});
    el.addEventListener('drop',function(e){
      e.preventDefault();var fs=e.dataTransfer&&e.dataTransfer.files;
      if(fs&&fs.length&&!S.drag){agregar(Array.from(fs));return;}
      var c=e.target.closest&&e.target.closest('.c'),ids=S.sel.has(S.drag)?S.pg.filter(function(q){return S.sel.has(q.id);}).map(function(q){return q.id;}):[S.drag];
      if(S.drag&&(!c||ids.indexOf(c.dataset.id)<0))mover(ids,c?c.dataset.id:null);
      S.drag=null;render();
    });
    el.addEventListener('dragend',function(){S.drag=null;el.querySelectorAll('.c.o').forEach(function(z){z.classList.remove('o');});});
    render();setMsg('Cargando el editor…');
    S.listo=cargar();
    return S.listo.then(function(){setMsg('');return agregar(o.files||[]);}).catch(function(e){alert(e.message);cerrar();});
  }
  return {abrir:abrir,compactarArchivo:compactarArchivo};
})();
/*ED_PDF_FIN*/
