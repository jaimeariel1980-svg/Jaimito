# Proyecto Azcu / Empleos PAMI — guía para quien trabaje en el repo

Dueño del proyecto: Ariel (no programador). Responder en español rioplatense, corto y directo. Él no aprueba ni mergea: Claude mergea sus propios cambios (ver "Cómo se publica"); nunca pegar código a mano en Apps Script.

## Cómo asesorar (regla pedida por Ariel, vale para todo chat de Claude en este repo, también el de Ever)
Ariel y Ever no ven el código: **Claude sabe más de lo técnico y tiene que guiar, no solo obedecer.** Decir que sí a todo hace perder tiempo.
- Antes de ejecutar un pedido de diseño o de lógica, dar **opinión propia y breve**: qué recomienda, qué riesgos o contradicciones ve (datos duplicados, porcentajes que se distorsionan, cosas que se rompen en otras pantallas) y si hay una alternativa más simple.
- Si el pedido tiene un problema, **decirlo y contradecir con fundamentos** aunque sea incómodo; proponer un camino recomendado en vez de un listado de opciones.
- Avisar de los problemas que nadie pidió mirar (ej.: dos clientes visitando la misma casa generan dos fichas).
- No preguntar lo que se puede decidir con criterio; preguntar solo lo que cambia el resultado y es del negocio.
- Si el pedido es claro y chico, hacerlo sin vueltas. Si es grande, proponer el orden de trabajo.
- Ariel puede escribir «objetá» o «contra-argumentá»; sin pedirlo, igual hay que opinar.

## Qué hay en el repo
| Carpeta | Qué es | Dónde se publica |
|---|---|---|
| `appscript/` | **Azcu oficial** (cuenta azcuenaga): `Azcubot.gs` (agente, avisos, usuarios), `Codigo.gs` (planilla Maestra, precarga, documentos) y las pantallas HTML | Apps Script de azcuenaga (clasp + GitHub Actions) |
| `pami/` | **Empleos PAMI** (cuenta dotacionpp1): reclutamiento, tablero, screening, competencias | Apps Script de dotacionpp1 |
| `azcubot/` | **Web/PWA** `azcubot.pages.dev` (celulares). Fuentes: `azcubot-widget.html`, `Tablero.html`, `tablero_bridge.js`, `paginas/`. Salida: `azcubot/netlify/` | Cloudflare Pages |
| raíz (`*.py`, `bot-azcuenaga.gs`, `job_profiles/`) | Restos de proyectos viejos (Telegram, reclutamiento en Python). **No tocar** | — |

## Cómo se publica (todo con Merge)
- **Autorización del dueño (Ariel):** confía en Claude y en Ever. Claude **mergea sus propios pull requests sin pedir aprobación** después de validar sintaxis y revisar el resultado de Actions/Cloudflare; Ever mergea los suyos. Excepciones: pedir confirmación antes de borrar datos, cambiar permisos o hacer cambios destructivos.
- Trabajar en una rama, abrir pull request hacia la rama principal `claude/recruitment-automation-google-b1Ric` (no hay `main`).
- Al hacer Merge: `.github/workflows/apps-script.yml` publica `appscript/` y/o `pami/` con clasp, y Cloudflare publica la web (rama de producción = la principal).
- Antes de publicar, el workflow compara el Apps Script real contra lo último publicado (`.github/scripts/guardia.py`). Si alguien lo tocó a mano, **frena** y deja la copia real en la rama `sync-azcu` / `sync-pami`. Integrar esos cambios al repo antes de volver a publicar.
- Acción manual: Actions → Apps Script → Run workflow (`bajar` copia el real al repo; `publicar` fuerza).
- Secretos/variables de GitHub ya cargados: `CLASPRC_AZCU`, `CLASPRC_PAMI`, `DEPLOYMENT_AZCU`, `DEPLOYMENT_PAMI`. No imprimirlos ni pedirlos.
- Los `.json` están ignorados por git salvo `appscript/*.json` y `pami/*.json`.

## Reglas de trabajo
1. **Nunca editar Apps Script directo.** Todo por el repo.
2. Después de tocar `azcubot-widget.html`, `Tablero.html`, `tablero_bridge.js` o `paginas/`: `cd azcubot && python3 build_netlify.py` y commitear `azcubot/netlify/`. La línea 1 de `build_netlify.py` (`EXEC`) apunta al `/exec` de Azcu oficial.
3. Validar sintaxis antes de subir: `node -e "new Function(require('fs').readFileSync('appscript/Azcubot.gs','utf8'))"` (idem `Codigo.gs`) y `node --check` sobre los scripts de los HTML.
3b. El editor de PDF (miniaturas, separar con IA, unir, ordenar) vive en `azcubot/editor_pdf.js` y se copia dentro de `Precarga.html` y `Tablero.html` (ambas copias). Tras editarlo: `cd azcubot && python3 inyectar_editor.py && python3 build_netlify.py`.
4. Las pantallas de `appscript/*.html` (Apps Script) y `azcubot/paginas/*.html` (web hospedada) son **dos copias**; mantenerlas coherentes (la hospedada usa `tablero_bridge.js` para emular `google.script.run`).
5. Mantener cambios mínimos, sin rediseños no pedidos. No sacar funciones existentes.
6. Nada de claves en el repo (OpenAI, OneSignal, PIN).

## Azcu (appscript/)
- Planilla Maestra: `1EwTv5hpxMCDHi4K47eCrhoKxxnk-FStulSeiEf6myl0` (pestañas Seguimiento, Movimientos, Encuestas, Consultas, Recordatorios, **Usuarios**). Carpeta raíz de Drive: `1LuYwU-I-hE1lEk6rcSblGmMVqzaPUdAh`. Script ID: `14UlZOapJSJGC7mmPG3ViPopmfBSGj4D5ffuDO--LxOcPvpaykviTzkrq`.
- Propiedades del script (no están en el repo): `AZCU_PIN`, `OPENAI_API_KEY`, `ONESIGNAL_KEY`, `AZCU_GERENCIA`, `AZCU_GERENTE` (respaldo), `AZCU_TODOS` (modo prueba), `ULTIMO_CORRELATIVO_<año>`.
- Códigos de precarga `PRE-AAAA-NNNN`. Columnas del Maestro por nombre de encabezado (si falta una, `crearColumnasFaltantes`).
- Transporte PWA ↔ Apps Script: pedidos por `fetch` no-cors (POST) + JSONP, con camino GET directo para pedidos chicos (`azc` en base64); deduplicación `azs_<id>`. Los pedidos arrancan **después del evento `load`** para que la pestaña no quede "cargando". Fuentes y OneSignal también se piden después de cargar.
- **Visitas** (tarjetas «Visitas hoy/semana» y botón 🗓 en el Tablero; `visitasTb`): salen de la pestaña Seguimiento. Las visitas a propiedades de **otras inmobiliarias** no son propiedades: se guardan como filas de Seguimiento con código `EXT-…` y columnas 12–16 (Inmobiliaria, Dirección, Link ficha, Precio, Colega) vía `registrarVisitaExterna`; no entran en Propiedades ni KPIs. Chat: `proponer_visita_externa`, `visitas` y `proponer_cancelar_visita`. Cancelar una visita (`cancelarVisita`) no borra nada: la fila queda con resultado «Visita cancelada» y el motivo en Observaciones, y se saca del Calendar; no cuenta como visita realizada.
- **Documentos obligatorios** (`DOC_OBLIG_` en `Codigo.gs`; `req` por propiedad, espejo `DOC_OBLIG` en `Tablero.html`): EPE Luz, Litoral Gas, Aguas Santafesinas, TGI Municipal, API Provincial, Escritura, Autorización de venta, DNI propietario, y Expensas solo si es Departamento. El resto es opcional (gris, no cuenta como falta ni en los %) salvo que se «exija» en esa propiedad (columna `Docs exigidos (extra)`, se crea sola). Completa = 100 %, Parcial = 50–99 %, Inicial = menos de 50 % de los obligatorios.
- **Papelera** (pestaña `Papelera`, botón 🗑 en el Tablero; `papeleraTb`, `restaurarPapelera`): al borrar una propiedad, una interacción/visita o un documento se guarda una copia 30 días (`_aPapelera_` en `Codigo.gs`) y se puede restaurar desde el Tablero o por chat (`papelera`, `proponer_restaurar`). Todo borrado nuevo debe pasar por `_aPapelera_`. Se purga sola al usarla.
- **Ficha**: dirección y datos técnicos van en `<details>` plegados (`detSecHtml`); el bot muestra la ficha resumida y ofrece «Ver datos técnicos» / «Ver historial» (`azcuFmtFicha_(id, 'tec')`).
- Horario de silencio de avisos: viernes 18 h → lunes 8 h (cola `AZCU_COLA`).

## Usuarios y avisos
- Pestaña **Usuarios** (se crea con `azcuCrearUsuarios`; `azcuActualizarListaAvisos` actualiza las opciones): nombre y apellido, rol, activo, avisos (`Todos` / `Solo gerencia` / `Solo mis propiedades` / `Ninguno`) y nombres anteriores.
- La app llama a `azcuQuien` al abrir y reemplaza el nombre viejo por el oficial; el identificador de OneSignal es el slug del nombre (`azcuSlug_`: minúsculas, sin tildes, `_`).
- `azcuPush_` aplica la pestaña Usuarios; "Azcu terminó ✅" y los recordatorios ⏰ son personales y no se difunden.
- Probar avisos: `azcuProbarGerentes` (muestra la respuesta cruda de OneSignal por nombre), `azcuProbarPush`, `azcuDiagnostico`.
- Cada persona tiene que instalar la PWA (iPhone: Safari → Compartir → Agregar a pantalla de inicio, iOS ≥ 16.4) y tocar "Activar avisos".

## Empleos PAMI (pami/)
- Script ID `1-qPLIJcxSKR9Xk3V_EV-mAdOOwHHN-W1b6MUhDygLI7ZM5AHVKQJmgfq`; cuenta `dotacionpp1@gmail.com`. Ver `pami/LEEME.md` (puesta en marcha, estados, agenda de citas). Tras cambios que agreguen columnas/pestañas: ejecutar `setup()` una vez; `verificarInstalacion()` revisa todo.
- Tablero privado: `<url /exec>?tablero=<ADMIN_KEY>` (la clave no está en el repo; sale de `mostrarLinks()`).

## Herramientas útiles ya hechas
- `azcubot/netlify/diag.html` (en la web: `/diag.html`): mide latencia real del transporte desde una compu.
- `azcuMedir` (Apps Script): tiempos del servidor.
- `azcubot/PUBLICACION.md`: resumen del circuito de publicación.
