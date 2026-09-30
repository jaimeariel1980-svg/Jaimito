# AzcuBot — asistente del tablero

1. **Apps Script del tablero** → nuevo archivo `AzcuBot.gs` → pegar `AzcuBot.gs`.
2. **Configuración del proyecto → Propiedades del script**:
   - `OPENAI_API_KEY` = clave NUEVA (revocar las anteriores).
   - `AZCU_PIN` = un PIN (recomendado; se pide una vez por navegador).
3. **Tablero.html** → pegar `azcubot-widget.html` justo antes de `</body>`.
4. Implementar → Administrar implementaciones → Nueva versión.
5. Probar: abrir el tablero (Ctrl+F5) → burbuja abajo a la derecha. Función de prueba en el editor: `azcuProbar`.

El widget usa la variable `LOGO` del tablero para el ícono (logo + ojitos).
Cambios de datos: siempre pasan por una tarjeta "Confirmá esta acción".
