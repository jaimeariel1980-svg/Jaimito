# AzcuBot

## A) Servidor (Apps Script del tablero)
1. Archivo nuevo `AzcuBot.gs` → pegar `AzcuBot.gs`.
2. Agregar una función nueva (no existe `doPost` en el tablero):
   `function doPost(e){ return azcuApi(e); }`
3. Propiedades del script: `OPENAI_API_KEY` (ya existe) y `AZCU_PIN` (PIN del equipo).
4. Implementar → Administrar implementaciones → Nueva versión. Acceso: "Cualquier persona".

## B) App de celular (Netlify)
1. Subir la carpeta `netlify/` (o `azcubot-netlify.zip` descomprimido) a https://app.netlify.com/drop
2. En el celular abrir la URL de Netlify → menú del navegador → "Agregar a pantalla de inicio".
3. La primera vez pide nombre y el PIN.

## C) Burbuja en el tablero de escritorio (opcional)
Pegar `azcubot-widget.html` antes de `</body>` en Tablero.html.

`build_netlify.py` regenera `netlify/index.html` a partir de `azcubot-widget.html`.
