# Empleos PAMI

Proyecto independiente (no comparte nada con MACS). Todo vive en la cuenta dotacionpp1@gmail.com.

## Archivos
- Apps Script: `Codigo.gs`, `Proceso.gs`, `Calendario.gs`, `Formularios.gs`, `Logo.gs`
- HTML: `Web`, `Tablero`, `Screening`, `Competencias`, `Estilos`
- `appsscript.json` (Configuración del proyecto → "Mostrar appsscript.json")

## Puesta en marcha
1. Con la sesión de dotacionpp1@gmail.com: Google Sheet nuevo → Extensiones → Apps Script.
2. Pegar todos los archivos.
3. Propiedades de la secuencia de comandos → `OPENAI_API_KEY`.
4. Ejecutar `setup()` (aceptar permisos: Drive, Gmail, Calendar).
5. Implementar → Aplicación web (ejecutar como: yo; acceso: cualquier persona).
6. Ejecutar `mostrarLinks()`.

Si ya estaba instalado: reemplazar archivos, volver a ejecutar `setup()` (agrega columnas/pestañas nuevas sin borrar nada) y publicar una versión nueva.

## Flujo de estados
CV → Screening → Competencias → Entrevista PAMI virtual → Entrevista PAMI presencial → Contratado.
Nota A / B+ avanzan, B avanza en CV y screening (en competencias y PAMI queda en reserva), C descarta en esa etapa.

Todo se edita desde el Sheet: Busquedas, Preguntas, Screening_Eval, Screening_Val, Competencias, Plantillas, Ajustes.
