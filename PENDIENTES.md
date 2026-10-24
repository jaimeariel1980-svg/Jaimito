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
