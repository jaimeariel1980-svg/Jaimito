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
