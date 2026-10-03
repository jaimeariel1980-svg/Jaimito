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
