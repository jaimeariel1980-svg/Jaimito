EXEC = "https://script.google.com/macros/s/AKfycbw82JqJx-WAAah3rd9V1U-KObnkWLmMKIASwDhrHK9CX4fkDGJPiO2fAnmMjofsPiHBFg/exec"
w = open('azcubot-widget.html', encoding='utf8').read()
head = '''<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#0a3d91">
<meta name="mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="Azcu">
<title>Azcu — Azcuénaga</title>
<link rel="manifest" href="manifest.webmanifest">
<link rel="icon" href="icons/icon-192.png">
<link rel="apple-touch-icon" href="icons/icon-192.png">
<script src="https://cdn.onesignal.com/sdks/web/v16/OneSignalSDK.page.js" defer></script>
<style>html,body{margin:0;height:100%;background:#0a3d91;overscroll-behavior:none}</style>
</head><body>
<script>
/* ===== CONFIGURACIÓN — editar estas dos líneas ===== */
window.AZCU_API = "'''+EXEC+'''";      // URL /exec del despliegue de Apps Script del tablero
window.AZCU_TABLERO = "tablero.html";
window.AZCU_APP = true;
window.AZCU_ONESIGNAL = "26a05ded-18d0-4349-972b-b30e5e614805";
window.OneSignalDeferred = window.OneSignalDeferred || [];
OneSignalDeferred.push(function(OneSignal){ return Promise.resolve(OneSignal.init({appId: "26a05ded-18d0-4349-972b-b30e5e614805"})).catch(function(e){ window.__osErr = String((e && e.message) || e); }); });

</script>
'''
open('netlify/index.html', 'w', encoding='utf8').write(head + w + '\n</body></html>\n')

b = open('tablero_bridge.js', encoding='utf8').read().replace('__API__', EXEC)
t = open('Tablero.html', encoding='utf8').read()
t = t.replace('<head>', '<head>\n<script>' + b + '</script>', 1)
open('netlify/tablero.html', 'w', encoding='utf8').write(t)

import re
BR = open('tablero_bridge.js', encoding='utf8').read().replace('__API__', EXEC)
def hospedar(src, dst, fn=None):
    h = open('paginas/'+src, encoding='utf8').read()
    if fn: h = fn(h)
    h = re.sub(r'<base target="_top">\s*', '', h)
    m = re.search(r'<head[^>]*>', h, re.I)
    h = h[:m.end()] + '\n<script>' + BR + '</script>' + h[m.end():]
    open('netlify/'+dst, 'w', encoding='utf8').write(h)

def f_precarga(h):
    m = re.search(r"<\? if \(codigoPrecarga\) \{ \?>(.*?)<\? \} else \{ \?>(.*?)<\? \} \?>", h, re.S)
    assert m, 'precarga: no encontré el bloque de retomar'
    A = m.group(1).replace('<?= codigoPrecarga ?>', '<span id="cod-txt"></span>')
    B = m.group(2)
    h = h[:m.start()] + '<div id="bn-ret" style="display:none">' + A + '</div><div id="bn-new">' + B + '</div>' + h[m.end():]
    h = h.replace("<?= codigoPrecarga || '' ?>", '')
    h = h.replace('</body>', """<script>
(function(){
  var c = new URLSearchParams(location.search).get('precarga') || '';
  if(c){ document.getElementById('codigoPrecarga').value = c; document.getElementById('cod-txt').textContent = c; document.getElementById('bn-ret').style.display = ''; document.getElementById('bn-new').style.display = 'none'; }
})();
</script>
</body>""")
    assert '<?' not in h
    return h
def f_encuesta(h):
    h = h.replace("'<?= codigo ?>'", "(new URLSearchParams(location.search).get('prop') || '')")
    assert '<?' not in h
    return h
hospedar('Precarga.html', 'precarga.html', f_precarga)
hospedar('Encuesta.html', 'encuesta.html', f_encuesta)
hospedar('Documentos.html', 'documentos.html')
hospedar('ProcedimientoVentas.html', 'procedimiento.html')
GO = """<!DOCTYPE html><html lang="es"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Azcuénaga</title></head><body style="font-family:sans-serif;text-align:center;padding:40px;color:#0a3d91">Abriendo…
<script>
(function(){
  var q = new URLSearchParams(location.search), app = q.get('app') || '', to = 'precarga.html';
  if(app === 'encuesta') to = 'encuesta.html?prop=' + encodeURIComponent(q.get('prop') || '');
  else if(app === 'documentos') to = 'documentos.html' + (q.get('prop') ? '?prop=' + encodeURIComponent(q.get('prop')) : '');
  else if(app === 'procedimiento') to = 'procedimiento.html';
  else if(app === 'tablero') to = 'tablero.html';
  else if(q.get('precarga')) to = 'precarga.html?precarga=' + encodeURIComponent(q.get('precarga'));
  location.replace(to);
})();
</script></body></html>"""
open('netlify/go.html', 'w', encoding='utf8').write(GO)
# enlaces internos del tablero hospedado
tb = open('netlify/tablero.html', encoding='utf8').read()
tb = tb.replace(EXEC + '?app=procedimiento', 'go.html?app=procedimiento').replace('href="' + EXEC + '"', 'href="go.html"')
open('netlify/tablero.html', 'w', encoding='utf8').write(tb)
