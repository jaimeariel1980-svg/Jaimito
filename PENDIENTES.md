# Pendientes y observaciones (lo que se detectó y no está resuelto)

Lo mantiene Claude. Prioridad: 🔴 urgente · 🟠 importante · 🟡 cuando se pueda. «Quién» = quién lo hace.

## Seguridad (auditoría 23/10 de Azcu; se hizo solo lo que no cambia el uso)
- 🔴 **Revocar el token de Telegram que se pegó en el chat** (BotFather → Revoke) y cargar el nuevo en `TELEGRAM_TOKEN`; después ejecutar `azcuTgActivar`. — *Ariel*
- 🔴 **Datos de propietarios sin PIN**: `obtenerDatosOFallar` devuelve la fila completa y `precargarDatos` con código existente pisa campos; hay que devolver solo lo necesario, sacar `diagnosticoBusquedaWeb` y pedir un token largo por ficha para retomar. Cambia cómo se retoma una precarga: hablarlo antes. — *Claude, con OK de Ariel*
- 🟠 **`google.script.run` desde el `/exec`** esquiva el PIN (probable, a confirmar con `google.script.run.obtenerPropiedades()` en la consola de `<exec>?app=tablero`). Arreglo: que el `/exec` no sirva HTML y/o chequear PIN adentro de cada función. — *Claude, con OK*
- 🟠 **PIN único para todos**: pasar a PIN por persona (token), también resuelve el nombre «autodeclarado» y el PIN en la URL. Obliga a todos a cambiar: planificar. — *Ariel + Claude*
- 🟠 **Encuesta pública**: token único por venta, un solo voto, validar nota. — *Claude*
- 🟠 **Topes en funciones públicas** (OpenAI, Drive, mails). — *Claude*
- 🟡 CSP y `X-Frame-Options` (`_headers` en Cloudflare, sacar `ALLOWALL`), escapar mails HTML, mensajes de error genéricos, `drive.file` en vez de `drive`. — *Claude*
- 🟡 Auditar también `pami/`, `pami-real/`, `liquidaciones/`, `caja-obra/` (NJ Arquitectura: `TABLERO_PIN`, bot Anita). — *Claude*

## Blumia (alta 26/10)
- 🔴 **Rotar las claves de OpenAI y Anthropic de Blumia**: estaban escritas en `Código.gs` y quedaron en la rama `sync-blumia` y en el historial. Crear nuevas, cargarlas en las propiedades `OPENAI_API_KEY` y `ANTHROPIC_API_KEY` y revocar las viejas. — *Ariel*
- 🟠 **Borrar la rama `sync-blumia`** (tiene las claves viejas) una vez integrado el código. — *Claude, con OK de Ariel*
- 🟠 **Web `/exec` pública** (kanban e internas sin clave): dejar abierto solo el formulario de postulación y proteger el resto con clave o login. Cambia el uso: hablarlo antes. — *Claude + Ariel*
- 🟠 **Perfil `director-comercial` incompleto**: faltan las agendas de Google Calendar (screening, competencias, empresa), `link_flyer` y `link_perfil`, y sumarlo a las listas de perfiles escritas a mano en `Código.gs` (sincronización de calendario, estados, links). Después, unificar esas listas en una sola leída de Config para que cada búsqueda nueva sea una sola función. — *Claude + Ariel*
- 🟡 Revisar con Ariel las competencias, criterios de screening y preguntas de postulación que se redactaron para `director-comercial`. — *Ariel*
- 🔴 **Blumia llegó al tope de 200 versiones de Apps Script** (26/10): el código se sube pero la web (`/exec`) no se actualiza y sigue con una versión vieja, que todavía trae las claves viejas de OpenAI/Anthropic ya revocadas (la IA de la web falla hasta actualizar). Hay que borrar versiones viejas desde el editor (Historial/Versiones del proyecto) y volver a publicar (Actions → Apps Script → `blumia` → `publicar`). Después, cuidar que cada merge a `blumia/` gasta una versión: agrupar cambios. El workflow ahora falla en rojo si no puede actualizar la web. — *Ariel (borrar) + Claude (publicar)*
- ✅ **Clave de PDFShift** (26/10): sacada del código; ahora se lee de la propiedad `PDFSHIFT_KEY`. Ariel decidió no rotarla (la clave vieja queda en el historial de git y en un chat): no volver a mencionárselo.
- ✅ **Avisos 🔔 del kanban** (26/10): campanita con pendientes y novedades, popups y refresco cada 3 min (probado en navegador; falta verlo con datos reales). Falta de PAMI: «anular etapa». — *Claude*
- ✅ **Registro de contactos y exportación a Contactos de Google** (26/10): log por candidato en la ficha y botón 📇 que baja un CSV para importar (sin permisos nuevos). Falta probarlo con datos reales. — *Ariel*
- ✅ **Workflow**: no detectaba cambios cuando solo cambiaba `Código.gs` (git escapaba la tilde); arreglado con `core.quotepath=false`.
- 🟡 Los textos de Plantillas dicen «Soy Ariel» fijo en algunos mensajes (p. ej. contacto-competencias): cambiarlo a [ENTREVISTADOR] desde ⚙ Configurar → Mensajes para que cada persona firme con su nombre. — *Ariel / Flavia*
- 🟠 **Tarjeta de Blumia vs PAMI** (27/10): PAMI muestra motivo de descarte (lista obligatoria con la C), formación, recomendado/espontáneo, días sin movimiento, «falta contactar», «descarte enviado», «ya se postuló antes» y faltantes; Blumia muestra nombre, zona, pretensión, cita y marca de contacto. Falta el **motivo de descarte** (columna nueva, selección al descartar, se ve en la tarjeta y en el Informe) y emparejar los datos de la tarjeta. — *Claude (esperando el «dale» de Ariel)*
- 🟡 Registrar los cambios de estado (fecha y quién) para poder medir tiempos por etapa en el 📊 Resumen. — *Claude*
- ✅ **Modo edición de los formularios** (26/10, v1): Flavia edita screening y competencias desde la página (PIN). Falta v2: botón «deshacer» desde la página (hoy el respaldo está en `Historial_Formularios` y se restaura a mano), editar la hoja de portada del perfil (hero) desde competencias y probarlo con ella. — *Claude / Ariel*
- ✅ **Informe de proceso por empresa** (26/10): cada empresa arma el suyo; MACS sin cambios.
- ✅ **Lectura de CVs con Claude** (26/10) para las búsquedas que no son de MACS; MACS sigue con ChatGPT mini (se cambia por perfil con la columna `cv_ia`).
- 🟡 Revisar con un informe y una entrevista reales cómo quedan los textos con Opus 5.5 / Sonnet 5.5 y si hay timeouts (Apps Script puede cortar esperas largas). — *Ariel*
- 🟠 **Probar el autocompletado de competencias con una entrevista real** (hecho el 26/10, probado solo con simulaciones): revisar que la evidencia cuente el ejemplo concreto y ajustar el prompt (`promptAnalisisEntrevista_`). El análisis usa Claude: ver el costo por entrevista. — *Ariel + Claude*
- 🟠 **Ideas de Ariel (26/10), orden sugerido**: (1) ~~autocompletado de **Competencias** como PAMI~~ (hecho; falta probarlo con una entrevista real); en **Screening** (llamada telefónica, no se graba) solo «Redactar con IA» y «Dictar» las notas; (2) poder **reabrir y corregir** screenings y competencias ya guardados; (3) **avisos, novedades y agenda** igual que PAMI; (4) **tarjetas borrables** (con papelera, como Azcu); (5) **contactos y registro de contactos** igual que PAMI; (6) inventario de qué más de PAMI se puede migrar. — *Claude + Ariel*
- 🟠 **Screening desde el celular** (a veces se hace por teléfono y en la calle, sin compu): versión móvil del formulario de screening con «Dictar» por voz y «Redactar con IA» sobre las notas; después evaluar un bot (Telegram, estilo Azcu) para cargar comentarios y resultados por chat. La llamada no se graba: se dicta o escribe lo que cuenta la persona. — *Claude + Ariel*
- 🟡 Compartir el perfil del puesto de Director Comercial solo con Flavia, Ariel y Analía (hace falta pasar sus mails). — *Ariel*
- 🟡 `Postulacion.html` pesa 650 KB (tipografías incrustadas): revisar si la página carga lento. — *Claude*
- 🟡 Cada `bajar` crea `sync-blumia`: si vuelve a haber claves en el código, quedan en esa rama. — *Claude*

## Empleos PAMI (27/10)
- ✅ **Candidatos con C ocultos para los jefes** (27/10, v162): un administrador los habilita desde la ficha (columna `Habilitado técnica`). Falta probarlo con datos reales. — *Ariel*
- ✅ **Segunda pasada del resumen para la técnica** (27/10, v162, `verificarResumen_`): falta ver con textos reales si corrige bien. — *Ariel*
- ✅ **Reprogramaciones, citas desaparecidas y contactos repetidos** (27/10, v162, `hechosProceso_`) llegan solos al resumen. No detecta un «no se presentó» como tal: solo lo que registra el sistema. — *Claude*
- 🟡 **Respaldo de la grabación**: el audio de la entrevista solo vive en la ventanita; si se cierra antes de terminar se pierde. Evaluar guardar el audio en Drive de forma temporal (decisión de Ariel por privacidad). — *Ariel + Claude*

## Operativo
- 🔴 **Apagar el bot viejo de Telegram (piloto)**: revocar su token en BotFather, borrar los activadores y archivar el proyecto de Apps Script, revocar su clave de OpenAI. Después, sacar del repo `bot-azcuenaga.gs` y los `.py` de la raíz. — *Ariel (cuentas) + Claude (repo)*
- 🟠 **Sin pruebas antes de publicar**: cualquiera con escritura en la rama principal publica a producción y el workflow no revisa sintaxis. Proponer chequeo automático de sintaxis y proteger la rama. — *Claude + Ariel (GitHub)*
- 🟠 **Sin copia de seguridad de la planilla Maestra** (la Papelera cubre borrados, no una planilla rota). Copia automática semanal a Drive. — *Claude*
- 🟠 **Dos copias de cada pantalla** (`appscript/` y `azcubot/`) que ya se desfasaron; unificar la fuente. Mientras tanto, cambiar siempre las dos. — *Claude*
- 🟡 Copias viejas sin uso dentro de `azcubot/` (`AzcuBot.gs`, `nueva_cuenta/`): revisar y borrar. — *Claude*
- 🟡 Avisos: decidir en una semana si Telegram pasa a ser el canal principal (hoy salen por OneSignal y Telegram a la vez, con posibles duplicados). Para escalar a clientes: evaluar WhatsApp (API de Meta o ManyChat) y, más adelante, app nativa. — *Ariel + Claude*
- 🟡 Cargar el mail de cada persona en Usuarios solo si se vuelve a activar el envío por mail. — *Ariel*

## Auditorías que faltan
- 🟠 **Funcional** (flujos de punta a punta, errores, diferencias entre copias), **diseño y uso en celular**, **rendimiento** y **estructura del código**. — *Claude*

## Producto (ideas hablando con Ariel)
- 🟡 Guía corta «qué se hace con Azcu y qué con el Tablero» para el equipo (Azcu = cargar al paso, sobre todo en el celular; Tablero = ver y gestionar, sobre todo en la compu). — *Claude, si Ariel quiere*

## Auditorías del 24/10 (funcional, diseño/uso, rendimiento) — hallazgos agrupados en tandas
Tanda 1 — ✅ HECHA el 24/10 (encuesta en celular con viewport/charset/title/lang; formularios del Tablero en una columna en celular; memoria de avisos y cola en trozos `azcuPropGet_/azcuPropSet_`; festejo de Reserva en el widget; `azcuTokens_` duplicada → `azcuTokensNom_`; promedios de encuesta; texto de borrar propiedad). Quedó sin hacer, por riesgo/valor bajo: unificar «Hizo propuesta» / «Hizo una propuesta» (hoy se tolera por regex).
Tanda 2 — corrigen datos (cambian algunos números o comportamientos, avisar):
- 🔴 **«Guardar cambios» de la ficha pisa lo que hicieron otros** y puede dejar documentos en «Pendiente» perdiendo el link (`Tablero.html:1500`, `Codigo.gs:2162`); **retomar una precarga** también pisa los `Doc:` subidos desde Tablero/Azcu (`_guardarFila`). Mandar solo lo tocado y no pisar links.
- 🟠 **Borrar una visita agendada no saca el evento del Calendar**, y el Tablero borra por número de fila sin validar (puede borrar otra); cancelar busca eventos por nombre: guardar `eventId`.
- 🟠 **Visitas infladas** (una agendada + su resultado cuentan 2; el informe al propietario usa ese número).
- 🟠 «Propuesta aceptada» puede devolver a Reserva una propiedad ya **Vendida**; el chat avisa «marcó como VENDIDA» aunque no cambie la etapa.
- 🟡 Fichas EXT duplicadas por «Av.»/mayúsculas (`_ndir_` vs `_dirClave_`); visitas huérfanas al borrar una propiedad; fichas EXT «reviven» al borrarlas; moneda del precio no se guarda; casi sin candados en escrituras simultáneas.
Tanda 3 — cambian cómo se usa (hablar antes):
- 🟠 **Ficha:** subir «Cargar interacción», «Agendar visita» y «qué falta» arriba; hoy están al final de 2 pantallas.
- 🟠 **Un solo vocabulario** («Cargar interacción» / «Cargar resultado» / «Visita externa» vs «Otra inmobiliaria») en Tablero y Azcu.
- 🟡 Contraste/tamaño de letra y botones de 44 px; burbuja de Azcu tapa botones; ✕ de borrar tarjeta chico y pegado; accesibilidad (cerrar con Escape, etiquetas); ocultar barra de documentación en fichas de otras inmobiliarias; «Cambiar etapa» se aplica solo sin confirmar; paso a paso de Azcu rígido; pantalla de inicio de Azcu larga (no se tocó por pedido de Ariel).
Rendimiento (a medida que crezca la base):
- 🟠 El paquete de datos crece sin techo (hoy 1–2 MB, ~8 MB con 500 propiedades): mandar solo lo liviano y pedir el detalle al abrir la ficha; leer el caché de una vez (`getAll`); `azcuCalentar` cada 10 min y solo si cambió; candado de «recalculando»; unir pedidos de apertura. Correr `azcuMedir` y ver cuántas propiedades/filas hay hoy.
Limpieza:
- 🟡 Borrar copias viejas sin uso: `azcubot/AzcuBot.gs`, `azcubot/Codigo.gs`, `azcubot/nueva_cuenta/`, zips, `LEEME.md` y `DIAGNOSTICO.md` viejos de `azcubot/`, capturas sueltas, ~570 líneas de utilidades de migración en `Codigo.gs`.
- 🟠 **Funciones peligrosas si se corren a mano desde el editor:** `corregirMigracion` (borra filas `MIG-` sin Papelera), `limpiarInteraccionesYComentarios`, `recalcularEstadosDocumentacion`: renombrarlas con `_` o borrarlas.
- 🟡 `appscript/Tablero.html` desfasada respecto de `azcubot/Tablero.html` (sin notch iPhone, sin `azcuQuien`…).
