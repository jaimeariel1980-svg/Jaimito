w = open('azcubot-widget.html', encoding='utf8').read()
head = '''<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1, viewport-fit=cover">
<meta name="theme-color" content="#0a3d91">
<meta name="apple-mobile-web-app-capable" content="yes">
<title>AzcuBot — Azcuénaga</title>
<style>html,body{margin:0;height:100%;background:#0a3d91;overscroll-behavior:none}</style>
</head><body>
<script>var LOGO="<?!= logo ?>";window.AZCU_APP=true;window.AZCU_TABLERO="<?!= tablero ?>";</script>
'''
open('AzcuBotApp.html', 'w', encoding='utf8').write(head + w + '\n</body></html>\n')
