(function(){
var API='__API__';
function ls(k,v){ try{ if(v===undefined) return localStorage.getItem(k); localStorage.setItem(k,v); }catch(e){ return null; } }
function uid(){ return 'b'+Date.now().toString(36)+Math.random().toString(36).slice(2,10)+Math.random().toString(36).slice(2,6); }
function transport(fn, args, ok, onErr){
  var id = uid(), done = false, got = false, fbOn = false, pollN = 0, t0 = Date.now(), pollT, scr, json = JSON.stringify({id:id, fn:fn, args:args});
  var ifr = document.createElement('iframe'); ifr.name = 'tbf_'+id; ifr.style.display = 'none';
  var form = document.createElement('form'); form.method = 'POST'; form.action = API; form.target = ifr.name; form.style.display = 'none';
  var inp = document.createElement('input'); inp.type = 'hidden'; inp.name = 'payload'; inp.value = json; form.appendChild(inp);
  function finish(){ done = true; clearTimeout(pollT); if(scr) scr.remove(); setTimeout(function(){ ifr.remove(); form.remove(); }, 300); }
  function fin(d){ if(done || !d || d.ok===undefined) return false; if(fbOn) ls('azcu_via','anon'); else ls('azcu_via',''); finish(); if(d.ok) ok(d.data); else onErr(new Error(d.error || 'error')); return true; }
  function jsonp(extra, anon){
    var cb = 'tbcb'+Math.random().toString(36).slice(2,9), s = document.createElement('script');
    window[cb] = function(d){ delete window[cb]; s.remove(); fin(d); };
    if(anon) s.crossOrigin = 'anonymous';
    s.src = API+'?azr='+id+extra+'&cb='+cb+'&_='+Date.now();
    s.onerror = function(){ delete window[cb]; s.remove(); };
    document.head.appendChild(s);
  }
  function fallback(){
    if(done || got || fbOn) return; fbOn = true;
    var pl; try{ pl = btoa(unescape(encodeURIComponent(json))); }catch(e){ pl = ''; }
    try{ fetch(API, {method:'POST', mode:'no-cors', credentials:'omit', headers:{'Content-Type':'text/plain'}, body:json}); }catch(e){}
    if(pl && pl.length <= 6500){ jsonp('&azc='+encodeURIComponent(pl), false); setTimeout(function(){ if(!done) jsonp('&azc='+encodeURIComponent(pl), true); }, ls('azcu_via')==='anon' ? 0 : 3000); }
  }
  function poll(){
    if(done) return;
    if(!got && !fbOn && Date.now()-t0 > 9000) fallback();
    if(Date.now()-t0 > 330000){ finish(); onErr(new Error('El servidor no respondió. Revisá tu conexión y probá de nuevo.')); return; }
    var cb = 'tbpl'+Math.random().toString(36).slice(2,9), wd, s0;
    window[cb] = function(d){ delete window[cb]; clearTimeout(wd); if(s0) s0.remove(); if(done) return; if(d) got = true; if(d && d.ok!==undefined) fin(d); else pollT = setTimeout(poll, d && d.progress ? 400 : 450); };
    scr = s0 = document.createElement('script'); if(fbOn && (pollN++ % 2)) scr.crossOrigin = 'anonymous';
    scr.src = API+'?azr='+id+'&cb='+cb+'&_='+Date.now();
    wd = setTimeout(function(){ if(window[cb]){ delete window[cb]; s0.remove(); if(!done) pollT = setTimeout(poll, 600); } }, 4000);
    scr.onerror = function(){ clearTimeout(wd); delete window[cb]; s0.remove(); if(!done) pollT = setTimeout(poll, 2000); };
    document.head.appendChild(scr);
  }
  (document.body || document.documentElement).appendChild(ifr); (document.body || document.documentElement).appendChild(form);
  form.submit(); pollT = setTimeout(poll, 500); if(ls('azcu_via')==='anon') setTimeout(fallback, 300);
}
var PUBLIC = ['obtenerDatosOFallar','extraerDatosDNI','extraerDatosPropiedad','extraerDatosEscritura','precargarDatos','guardarEncuesta'];
var pinQ = null;
function askPin(msg, cb){
  if(pinQ){ pinQ.push(cb); return; } pinQ = [cb];
  var ov = document.createElement('div');
  ov.style.cssText = 'position:fixed;inset:0;background:rgba(15,27,51,.55);z-index:2147483000;display:flex;align-items:center;justify-content:center;padding:20px;font-family:Inter,Segoe UI,sans-serif';
  ov.innerHTML = '<div style="background:#fff;border-radius:20px;padding:22px;max-width:340px;width:100%;box-shadow:0 20px 50px rgba(0,0,0,.3)"><div style="font:800 20px Archivo,Inter,sans-serif;color:#0a3d91;margin-bottom:6px">Ingresá el PIN</div><div style="font-size:14px;color:'+(msg ? '#d1495b' : '#6b7280')+';margin-bottom:12px">'+(msg || 'Te lo da quien administra Azcu. Se guarda en este dispositivo.')+'</div><input id="tbpin" type="password" inputmode="numeric" autocomplete="off" placeholder="PIN" style="width:100%;box-sizing:border-box;border:1.5px solid #e6e9f2;border-radius:14px;padding:14px;font:600 18px Inter,sans-serif;margin-bottom:10px"><button id="tbok" style="width:100%;border:0;border-radius:999px;background:#0a3d91;color:#fff;font:700 15px Inter,sans-serif;padding:14px;margin-bottom:8px;cursor:pointer">Entrar</button><button id="tbno" style="width:100%;border:0;border-radius:999px;background:#eef1f8;color:#0f1b33;font:600 14px Inter,sans-serif;padding:12px;cursor:pointer">Cancelar</button></div>';
  document.body.appendChild(ov);
  var i = ov.querySelector('#tbpin'); setTimeout(function(){ i.focus(); }, 100);
  function close(v){ ov.remove(); var q = pinQ; pinQ = null; q.forEach(function(f){ f(v); }); }
  ov.querySelector('#tbok').onclick = function(){ var v = (i.value||'').trim(); if(v) close(v); };
  ov.querySelector('#tbno').onclick = function(){ close(null); };
  i.onkeydown = function(e){ if(e.key==='Enter'){ var v = (i.value||'').trim(); if(v) close(v); } };
}
function send(fn, args, ok, fail, retried){
  if(fn==='getAppUrl' || fn==='getEncuestaUrl'){ setTimeout(function(){ ok(location.origin+'/go.html'); }, 0); return; }
  var az = /^azcu/.test(fn) || PUBLIC.indexOf(fn) >= 0, a = az ? args : [ls('azcu_pin')||''].concat(args);
  transport(fn, a, ok, function(e){
    var m = (e && e.message) || String(e);
    if(!az && /PIN/.test(m) && !retried){
      askPin(/incorrecto/i.test(m) && ls('azcu_pin') ? 'Ese PIN no es correcto, probá de nuevo.' : '', function(p){
        if(p===null){ fail(new Error('Falta el PIN para usar el tablero.')); return; }
        ls('azcu_pin', p); send(fn, args, ok, fail, false);
      });
      return;
    }
    fail(e instanceof Error ? e : new Error(m));
  });
}
function mk(h){
  return new Proxy({}, {get:function(t, k){
    if(k==='withSuccessHandler') return function(f){ return mk({ok:f, fail:h.fail}); };
    if(k==='withFailureHandler') return function(f){ return mk({ok:h.ok, fail:f}); };
    if(k==='withUserObject') return function(){ return mk(h); };
    return function(){
      var args = Array.prototype.slice.call(arguments);
      send(k, args, h.ok || function(){}, h.fail || function(e){ try{ console.error(e); alert('No se pudo completar la operación: '+((e && e.message) || e)); }catch(x){} });
    };
  }});
}
window.google = window.google || {};
window.google.script = {run: mk({}), host:{close:function(){}, setHeight:function(){}, setWidth:function(){}}};
})();
