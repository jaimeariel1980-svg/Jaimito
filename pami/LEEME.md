# Empleos PAMI

Proyecto independiente (no comparte nada con MACS). Todo vive en la cuenta dotacionpp1@gmail.com.

1. Con la sesión de dotacionpp1@gmail.com: crear un Google Sheet nuevo (ej. "Empleos PAMI") → Extensiones → Apps Script.
2. Pegar `Codigo.gs`, `Web.html`, `Tablero.html` (Archivo → HTML) y el contenido de `appsscript.json` (Configuración del proyecto → "Mostrar appsscript.json").
3. Configuración del proyecto → Propiedades de la secuencia de comandos → agregar `OPENAI_API_KEY`.
4. Ejecutar `setup()` y aceptar permisos.
5. Implementar → Nueva implementación → Aplicación web (ejecutar como: yo; acceso: cualquier persona).
6. Ejecutar `mostrarLinks()` y leer el registro.

Entradas de CVs: web (form), correo (adjuntos PDF/Word a la bandeja de entrada) y carpeta Drive "Para procesar". Un disparador cada 10 minutos procesa correo y carpeta.
Todo se edita desde el Sheet: Busquedas, Preguntas, Ajustes.
