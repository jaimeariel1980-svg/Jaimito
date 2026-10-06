# Empleos PAMI

Proyecto independiente (no comparte nada con MACS). Todo vive en la cuenta dotacionpp1@gmail.com.

## Archivos
- Apps Script: `Codigo.gs`, `Proceso.gs`, `Calendario.gs`, `Formularios.gs`, `Logo.gs`
- HTML: `Web`, `Tablero`, `Screening`, `Competencias`, `Estilos`, `FormScreening`, `FormCompetencias`
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

## Agenda de citas (Google Calendar)
1. En Google Calendar de dotacionpp1 → Crear → "Agenda de citas" (una por tipo de entrevista).
2. El título de la agenda tiene que incluir la palabra: `Screening`, `Competencias`, `PAMI virtual` o `PAMI presencial`.
3. Copiar el link público de cada agenda en la pestaña Ajustes: `link_agenda_screening`, `link_agenda_competencias`, `link_agenda_pami1`, `link_agenda_pami2`.
4. Los mensajes de WhatsApp (pestaña Plantillas) mandan ese link. Cuando el candidato reserva, el sistema lo detecta solo (cada 10 min) por su mail o por el nombre que figura entre paréntesis en el evento, marca "agendado" y muestra la cita en la tarjeta y en Agenda.
5. Los turnos reservados que no coinciden con ningún candidato aparecen en Agenda como "sin candidato asociado".

## Si una página no abre
Ejecutar `verificarInstalacion()`. Las páginas muestran el error real. Si la URL termina en /dev, usar la /exec y pegarla en Ajustes → url_web.
