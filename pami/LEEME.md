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
3. Copiar el link público de cada agenda en la pestaña Ajustes: `link_agenda_screening`, `link_agenda_competencias`, `link_agenda_tecnica` (entrevista técnica con los jefes de Enfermería).
4. Los mensajes de WhatsApp (pestaña Plantillas) mandan ese link. Cuando el candidato reserva, el sistema lo detecta solo (cada 10 min) por su mail o por el nombre que figura entre paréntesis en el evento, pasa a "Screening agendado" (o Competencias/Técnica) y muestra día y hora en la tarjeta, la ficha y Agenda. Mientras tanto la tarjeta dice "Contactado para screening · sin agendar" (en amarillo) y Agenda los lista en "Contactados · todavía sin agendar".
5. Los turnos reservados que no coinciden con ningún candidato aparecen en Agenda como "sin candidato asociado".

## Si una página no abre
Ejecutar `verificarInstalacion()`. Las páginas muestran el error real. Si la URL termina en /dev, usar la /exec y pegarla en Ajustes → url_web.

## Dos policlínicos (PAMI 1 y PAMI 2)
- PAMI 1 usa el calendario de dotacionpp1; PAMI 2, el de dotacionpp2 (Ajustes → `calendario_poli2`).
- dotacionpp2 le comparte su calendario a dotacionpp1 con permiso "Ver todos los detalles de los eventos" (o "Hacer cambios en eventos" para que el tablero cree citas ahí). El sistema lee los dos calendarios y asocia cada turno al candidato.
- Links de agenda por policlínico: PAMI 1 usa `link_agenda_screening`, `link_agenda_competencias`, `link_agenda_tecnica`; PAMI 2 usa las mismas claves con sufijo `_poli2`. Si a un candidato de PAMI 2 le falta su link, el tablero avisa (no manda el de PAMI 1). Las claves viejas `link_agenda_pami2(_poli2)` se renombraron solas a `link_agenda_tecnica(_poli2)`.
- El mensaje de WhatsApp elige el link según el policlínico del candidato (el confirmado, o el sugerido).
- Los títulos de las agendas siguen llevando `Screening`, `Competencias` o `Técnica`.
- El calendario de PAMI 2 (dotacionpp2) tiene que estar compartido con dotacionpp1 con el permiso **«Ver todos los detalles del evento»** y la invitación aceptada. Si no, la Agenda muestra un cartel rojo explicando qué falta (el sistema también intenta suscribirse solo).

## Documentación para el ingreso
- Formulario público: `<url /exec>?docs=1` (PAMI 1) y `<url /exec>?docs=2` (PAMI 2). `mostrarLinks()` los muestra.
- Pide apellido, nombre, CUIL/CUIT (se valida) y cinco archivos (PDF o imagen, hasta 10 MB c/u): DNI (admite varios, frente y dorso), CV, constancia de CUIL, título de Enfermería y matrícula vigente.
- Cada persona tiene su carpeta `Apellido_Nombre_CUIL` y los archivos se nombran `Apellido_Nombre_DNI`, `_CV`, `_CUIL`, `_Titulo_Enfermeria`, `_Matricula` (si se reenvía, se agrega `_v2`).
- Además se arma `Apellido_Nombre_Completo.pdf` con todo junto en ese orden (las fotos pasan a páginas A4). Lo arma el navegador del candidato con pdf-lib; si no se puede (PDF protegido o dañado, sin conexión a la librería), ese archivo se saltea y los sueltos quedan igual.
- Dónde se guarda (pestaña Ajustes):
  - PAMI 1 → `carpeta_docs` (se crea sola en el Drive de dotacionpp1: "Empleos PAMI - Documentación ingreso").
  - PAMI 2 → `carpeta_docs_poli2` (cargada: carpeta "Documentación" de dotacionpp2, ID `1C0HSI69hoCWhVN2UW0aBG8lTs5hV5ZGO`). Para cambiarla: en el Drive de **dotacionpp2** crear una carpeta, compartirla con dotacionpp1 como **Editor** y pegar su ID (lo que sigue a `/folders/` en el link). Mientras esté vacío, lo de PAMI 2 se guarda en la subcarpeta "PAMI 2 (falta carpeta de dotacionpp2)" dentro de la de PAMI 1.
- Tablero → WhatsApp → "📄 Pedido de documentación": manda el link del policlínico del candidato (confirmado o sugerido) con su ID, así la documentación queda asociada a la ficha (columna Documentación y botón "📂 Documentación"). Si llega por el link general, se asocia por DNI (sale del CUIL).
- Todo envío queda en la pestaña **Documentacion** y llega un mail a `correo_aviso`. Texto del WhatsApp: plantilla `pedido-documentacion`.

## URL pública
`urlBase_()` usa Ajustes → `url_web` si no termina en `/dev`; si no, la `/exec` que devuelve Apps Script o la última `/exec` por la que entró alguien (propiedad `URL_EXEC`). Los links que se mandan a candidatos nunca deberían ser `/dev`.

## Agenda de entrevista técnica (jefes de Enfermería)
- Los turnos se leen de: el calendario principal de dotacionpp1 y de dotacionpp2, cualquier calendario de la lista de dotacionpp1 cuyo nombre diga "Técnica" o "Enfermería" (o "Competencias"), y los calendarios cargados en Ajustes `calendario_tecnica` (PAMI 1) y `calendario_tecnica_poli2` (PAMI 2): ID del calendario (Calendar → Configuración del calendario → "ID del calendario") o el mail de la cuenta.
- Ese calendario tiene que estar compartido con dotacionpp1 con "Ver todos los detalles del evento".
- En esos calendarios propios, todo turno cuenta como entrevista técnica aunque el título no lo diga.
