# Empleos PAMI

Proyecto independiente (no comparte nada con MACS).

1. Crear un Google Sheet nuevo (ej. "Empleos PAMI") → Extensiones → Apps Script.
2. Pegar `Codigo.gs`, `Web.html`, `Tablero.html` (Archivo → HTML) y el contenido de `appsscript.json` (Configuración del proyecto → "Mostrar appsscript.json").
3. Configuración del proyecto → Propiedades de la secuencia de comandos → agregar `OPENAI_API_KEY`.
4. Ejecutar `setup()` y aceptar permisos.
5. Implementar → Nueva implementación → Aplicación web (ejecutar como: yo; acceso: cualquier persona).
6. Ejecutar `mostrarLinks()` y leer el registro.

Todo se edita desde el Sheet: Busquedas, Preguntas, Ajustes.
