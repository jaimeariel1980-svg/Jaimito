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
CV → Screening → Competencias → Entrevista técnica (División de Enfermería) → Contratado.
Nota A / B+ avanzan, B avanza en CV y screening (en competencias y PAMI queda en reserva), C descarta en esa etapa.

Todo se edita desde el Sheet: Busquedas, Preguntas, Screening_Eval, Screening_Val, Competencias, Plantillas, Ajustes.

## Agenda de citas (Google Calendar)
1. En Google Calendar de dotacionpp1 → Crear → "Agenda de citas" (una por tipo de entrevista).
2. El título de la agenda tiene que incluir la palabra: `Screening`, `Competencias` o `Técnica`.
3. Copiar el link público de cada agenda en la pestaña Ajustes: `link_agenda_screening`, `link_agenda_competencias`, `link_agenda_pami2` (esta última es la de la entrevista técnica).
4. Los mensajes de WhatsApp (pestaña Plantillas) mandan ese link. Cuando el candidato reserva, el sistema lo detecta solo (cada 10 min) por su mail o por el nombre que figura entre paréntesis en el evento, pasa a "Screening agendado" (o Competencias/Técnica) y muestra día y hora en la tarjeta, la ficha y Agenda. Mientras tanto la tarjeta dice "Contactado para screening · sin agendar" (en amarillo) y Agenda los lista en "Contactados · todavía sin agendar".
5. Los turnos reservados que no coinciden con ningún candidato aparecen en Agenda como "sin candidato asociado".

## Si una página no abre
Ejecutar `verificarInstalacion()`. Las páginas muestran el error real. Si la URL termina en /dev, usar la /exec y pegarla en Ajustes → url_web.

## Dos policlínicos (PAMI 1 y PAMI 2)
- PAMI 1 usa el calendario de dotacionpp1; PAMI 2, el de dotacionpp2 (Ajustes → `calendario_poli2`).
- dotacionpp2 le comparte su calendario a dotacionpp1 con permiso "Ver todos los detalles de los eventos" (o "Hacer cambios en eventos" para que el tablero cree citas ahí). El sistema lee los dos calendarios y asocia cada turno al candidato.
- Links de agenda por policlínico: PAMI 1 usa `link_agenda_screening`, `link_agenda_competencias`, `link_agenda_pami2`; PAMI 2 usa las mismas claves con sufijo `_poli2`. Si falta el de PAMI 2, se usa el de PAMI 1.
- El mensaje de WhatsApp elige el link según el policlínico del candidato (el confirmado, o el sugerido).
- Los títulos de las agendas siguen llevando `Screening`, `Competencias` o `Técnica`.
