# Sistema de selección multi-perfil (MACS) — contexto técnico completo

Sistema de reclutamiento construido en Google Apps Script + Sheets + Drive + Gmail + Calendar + OpenAI (GPT-4o-mini), operado por Blumia Consulting (Ariel, no programador — Claude escribe todo el código). Nació como sistema de una sola búsqueda, se migró a motor multi-perfil (un código, N búsquedas) y luego a multi-empresa. Construido a lo largo de ~5 sesiones: "PLANTILLA AUTOMATIZACIÓN SELECCIÓN" (fundacional) → "Multiperfil 1" → "Multiperfil 4" (saltea numeración de sesión, no de versión del sistema) → sesiones de banco de talento / comunicaciones (ver `macs-recruitment-system.md`).

## 1. Entornos / IDs (OJO: dos sistemas distintos, no confundir)

**Producción original (una sola búsqueda, Supervisor de Operaciones, MACS/Ternium San Nicolás):**
- Sheet ID: `1kDYhElzoRhGwgyBkfEA4EdE4R3zAfWuXe3Thb5FcX00`
- Deploy URL termina en `AKfycbwz...`
- No modificar estructuralmente.

**Sistema multi-perfil (el vigente, el que escaló a multi-empresa):**
- Sheet maestro ID: `151V8cc-w00tP9C6mSxVaKNkRxgKXgUDI4L0KWXeNJGI`
- Deploy URL activa termina en `AKfycbytekPKN31bZLxczDVknMr433K9KflzmrR-dCWsC1hLlx2MO7tbw6SrMjFXa1PnucQo/exec` (verificar vigencia — históricamente hubo 3 URLs distintas pisadas en columnas, ver §8)
- London Supply Group (2da empresa, simulada): Sheet `1Ej-ZBvl...NZxI`

OPENAI_API_KEY vive en `PropertiesService.getScriptProperties()` — **pendiente**: en el código actual sigue hardcodeada en varios lugares (deuda de seguridad, nunca resuelta).

## 2. Carpeta Drive raíz

`ATRACCIÓN Y ELECCIÓN` (ID `18bhn7cD2PkcK_SldSfKHrjiieHvO3tYH`). Cada perfil tiene 4 subcarpetas: madre, recibidos (CVs nuevos), pdf_screening, pdf_competencias. **El script usa el ID de carpeta, no el nombre** — los nombres de carpeta pueden no coincidir con el nombre del puesto sin que rompa nada.

## 3. Pestaña `Config` — el corazón del motor

Columnas (en este orden, la columna `empresa` se agregó después, siempre AL FINAL para no correr nada):

```
perfil_id | nombre_perfil | sheet_id | hoja_nombre | folder_madre | folder_recibidos |
folder_pdf_screening | folder_pdf_competencias | prompt_extra | empresa (col 10/J, agregada después)
```

**Regla de oro:** una columna nueva SIEMPRE va al final. Insertarla en el medio corre las demás y rompe cualquier `getRange(fila, N)` hardcodeado en el resto del código (pasó con `prompt_extra` en col 9).

`getConfig(perfil_id)` lee toda la fila y arma el objeto de config para ese perfil; `getConfig` recorre headers dinámicamente (no columnas fijas) para poder leer `empresa` sin romper nada preexistente.

Agregar un perfil nuevo = una fila en Config (+ filas espejo en las demás pestañas de perfil) + carpetas en Drive. **Nunca se toca código para sumar un perfil.**

## 4. Pestañas del Sheet maestro

- `Config` — ver §3
- `Perfiles_Hero` — datos de cabecera por perfil: `perfil_id, empresa, posicion, ubicacion, entrevistadoras, criterio_avanza, criterio_reserva`
- `Perfiles_Screening_Eval` — criterios de evaluación del screening telefónico, por perfil
- `Perfiles_Screening_Val` — condiciones de validación (checks tipo v1, v2…vN por perfil; reorganizada con `unificarScreeningVal_MACS` para que cada perfil quede agrupado y renumerado limpio)
- `Perfiles_Competencias` — competencias STAR + preguntas, por perfil, con pesos editables (pills)
- `Postulacion_Preguntas` — preguntas de screening del formulario público, configurables por perfil (patrón: todo desde el Sheet, sin tocar código)
- `Mail_Routing` — keywords de asunto → perfil, para el router de mails
- `Plantillas` (comunicaciones WhatsApp) — columnas `clave, canal, link_calendar, texto`, 8 plantillas con huecos `[NOMBRE][PUESTO][EMPRESA][ENTREVISTADOR][LINK][MOTIVO TÉCNICO]`
- `Banco` — banco de talento curado: `fecha_alta, nombre, categoria, seniority, zona, formacion, pretension, tel, mail, cv, resumen, origen_perfil, origen_empresa, origen_fila, nota_banco, screening_pdf, competencias_pdf, informe, postulacion, nota_cv, nota_scr, nota_comp, estado_origen`
- Una pestaña de candidatos POR PERFIL (ej. `Hoja 1` para operaciones, `RRHH` para coord-rrhh, etc.) — **los candidatos de distintos perfiles/empresas nunca se mezclan en la misma pestaña**, esa es la decisión de aislamiento de riesgo tomada explícitamente (ver §9)

## 5. Perfiles cargados (todos cliente MACS, Pilar Bs.As., salvo aclaración)

- `operaciones` — Supervisor de Operaciones (búsqueda original, zona Ramallo/Villa Constitución/San Nicolás, rubro medioambiental/limpieza industrial, liderazgo +3 años) — **esta es la de Ternium/San Nicolás, corrida en el sistema de producción original, no en el multiperfil**
- `coord-rrhh` — Coordinador RRHH (el "molde" que se usó para clonar los 3 siguientes)
- `analista-datos` — Analista de Datos, reporta a Dir. Ingeniería y Proyectos, Cs. Económicas, 3 años excluyente, Excel avanzado + BI/SQL/DAX intermedio + IA generativa, inglés, $2.5M, KPIs/dashboards/análisis económico-financiero
- `analista-compras` — Analista de Compras y Contrataciones, reporta a Dir. Admón. y Gestión, Ing. Industrial/Admón., 5 años excluyente, compras punta a punta + importación/comex, Excel avanzado, ERP excluyente, inglés, $3.2M, OSDE 210, presencial 8-17
- `analista-sgi` — Analista de Sistema de Gestión (Calidad/Seguridad/MA), reporta a Dir. Admón. y Gestión, Ing. Industrial/Seguridad/Ambiental, ISO 9001/45001/14001 excluyente, **exp. cambiada de 3 a 5 años excluyente durante la carga**, Excel avanzado, ERP intermedio. Red flag crítico: cliente quiere perfil analista, no auditor puro ni supervisor de seguridad — anotado explícitamente para no confundir el casting
- `london-gerente-rrhh` — 2da empresa (London Supply Group), reusa infraestructura/molde de `coord-rrhh`

Funciones de carga: `cargarPerfil_AnalistaDatos()`, `cargarPerfil_AnalistaCompras()`, `cargarPerfil_AnalistaSGI()` — todas con la misma firma que el molde `cargarPerfil_CoordRRHH()`, escriben en Config + Perfiles_Hero + Postulacion_Preguntas con `prompt_extra` como template string.

## 6. Flujo de datos (intake de CVs)

```
CV llega por mail a empleos@blumiaconsulting.com (o se tira a mano a carpeta "CVs nuevos" del perfil)
        ↓
procesarMailsRouter(): cascada — 1) keywords de asunto vs tabla Mail_Routing
                                   2) si no matchea, clasificación GPT (primeros 2000 caracteres del mail)
                                   3) si tampoco, va a hoja "Sin clasificar"
        ↓
procesarTodo_v2 (dispara procesarMailsRouter 1 vez, luego procesarDriveManual(cfg) por cada perfil activo leído de Config)
        ↓
procesarConChatGPT(file, cfg, textoYaExtraido?) — 3er parámetro opcional: si viene texto ya extraído
   (del router), SALTEA la conversión de Drive; si no viene, extrae como siempre (retrocompatible)
        ↓
esDuplicado(datos, sheet) — chequea si el candidato ya está cargado
        ↓
insertarFilaLibre(sheet, datos, url, fecha, fuente) — inserta en la primera fila libre de la pestaña del perfil
        ↓
CV se mueve de "CVs nuevos" a "CVs procesados"
```

`extraerTextoCV_` usa `Drive.Files.create` (API v3) — **nunca** `Drive.Files.insert` (v2), no existe en este proyecto y fue un bug recurrente al clonar código.

## 7. Scoring / umbrales

- Escala: **A / B+ / B / C** (NO "Alto/Medio/Bajo" — prompts viejos usan esa escala vieja, `migrarNotasViejas()` las mapea automáticamente al abrir)
- Auto-selección por puntaje con thresholds: **A ≥ 8.5, B ≥ 6.5, C < 6.5**, con override manual + botón de auto-reset
- `patchAntiDescarte_MACS` — parche aplicado a los 4 perfiles MACS con reglas anti-descarte:
  - Falta de experiencia en industria/rubro específico NUNCA justifica sola un C → pasa a B + nota "REVISAR:"
  - Candidato sobrecalificado → B (nunca B+ ni A) + nota "REVISAR: perfil por encima del puesto"
- Competencias: pesos editables por pill, tanto en el form de producción (Monday) como en `Index.html` del multiperfil

## 8. doGet / routing de las webapps

Ruteo por querystring: `?perfil=X` → screening (`Formulario.html`); `?tipo=competencias&perfil=X` → competencias (`Index.html`); `?tipo=kanban` → Kanban; `?tipo=banco` → banco de talento; `?tipo=postulacion` → form público (mid-development, ver §12).

Bug recurrente resuelto: un `doGet` corto en `Screening.gs` con condición `if (tipo === "competencias" || nombre)` pisaba al `doGet` correcto de `Codigo.gs`/`Código.gs`, haciendo que todos los links abrieran competencias. **Patrón general del proyecto: funciones duplicadas en distintos .gs se pisan en silencio** (pasó entre Código.gs/Screening.gs/Competencias.gs varias veces) — siempre buscar y borrar duplicados antes de debuguear otra cosa.

`PERFIL_ID` y `nombre`/`entrevistador` NO se leen de `window.location.search` dentro de `Formulario.html`/`Index.html` — falla en el sandbox iframe de Apps Script. Se inyectan server-side desde `doGet` con template scriptlets: `<?= perfil ?>`, `<?= candidato ?>`, `<?= entrevistador ?>`.

Columnas K y O de cada pestaña de perfil tienen HYPERLINK a las webapps (screening / competencias) — hubo 3 URLs de deploy distintas pisadas ahí por error; `arreglarTodosLosLinks()` las reescribió las 100 filas × 5 perfiles usando **punto y coma** (no coma) como separador de fórmula porque el Sheet está en locale español.

`procesarScreening` — bug resuelto: comparación de string plana en vez de NFD-normalizada (insensible a tildes), hacía que no encontrara candidatos con acentos. Fix: normalizar NFD tanto el nombre buscado como los valores del Sheet.

## 9. Multi-empresa — decisión de arquitectura (tomada, no teórica)

Dos opciones evaluadas:
- **A — un Sheet por empresa**: aislamiento total, pero requiere centralizar `SHEET_MASTER_ID` (hoy hardcodeado en todos lados)
- **B — un solo Sheet con columna/prefijo empresa**: un deploy, pero riesgo de mezclar candidatos de clientes distintos si se filtra mal

**Decisión final (con 2-3 clientes en el horizonte cercano):** un solo Sheet, pero la separación es **por pestaña, nunca por columna dentro de la pestaña de candidatos**. Esto es lo central: evita el riesgo de mostrarle a un cliente un candidato de otro por un filtro mal puesto. Lo que se tocó para esto:
1. `Config` +1 columna `empresa` (al final)
2. `getConfig` devuelve también `empresa`
3. Nav del Kanban: selector de empresa arriba, filtra qué perfiles del nav muestra
4. Funciones que barrían "los perfiles fijos" (calendar sync, `procesarTodo`, descartes) → pasaron a iterar por empresa activa en vez de lista fija de 4

MACS sigue siendo la empresa por defecto de las 5 búsquedas originales; London Supply es la 2da, simulada sobre el molde de `coord-rrhh`.

## 10. Kanban — módulos implementados

- Drag-and-drop, sync de Calendar, modal de candidato, estados por etapa (incluye etapas "empresa" = entrevista con el cliente)
- Selector de empresa en nav (optgroup), header muestra "Puesto · Empresa"
- Sync de Calendar matchea candidatos por nombre normalizado (2+ palabras de 4+ letras) o por username de email (ignorando dominio), para tolerar mismatches entre Sheet y Calendar
- Triggers activos: `procesarTodo_v2` cada 15 min, `procesarPendientes` aparte (forms), `onEditDescartado` on edit, `syncCalendarConfirmados` cada 15 min

### Comunicaciones WhatsApp
- `armarComunicacion(perfilId, fila, tipo)` — rellena huecos de la plantilla (`[NOMBRE][PUESTO][EMPRESA][ENTREVISTADOR][LINK][MOTIVO TÉCNICO]`), arma link `wa.me`
- Botón WhatsApp en la ficha → menú `menuWA`, 6 opciones (4 contactos: screening/competencias/empresa-virtual/empresa-presencial + 2 descartes)
- Marca "Contactado": `armarComunicacion` escribe en **col41** `"etapa · fecha"` al enviar un contacto (no descarte); `doGetKanban` expone el campo `contactado`; tarjeta muestra "✓ Contact. <etapa> · <fecha>" en verde sin mover de columna
- Botón "✕ borrar" junto a "Último contacto" en la ficha → `borrarMarcaContacto()` limpia col41
- Redacción de los textos la definió Ari (cálida pero profesional, sin amiguismo) — respetar estilo, solo corregir ortografía/gramática si lo pide

### Notas de empresa (etapas con el cliente)
- `notaEmpresa()` guarda A/B/C en **col44** (screening empresa) y **col45** (presencial empresa); crea columnas si faltan (`insertColumnsAfter`) + `clearDataValidations`
- `doGetKanban` lee `notaScrEmp` (col43) y `notaPresEmp` (col44, base-0 — ojo con el off-by-one entre lo que escribe notaEmpresa y lo que lee doGetKanban, documentado así en la fuente, no es un typo mío)
- Ficha muestra 4 notas: screening / competencias / screening empresa / presencial empresa
- Indicador de la tarjeta = etapa más avanzada: `notaPresEmp || notaScrEmp || notaComp || notaScr || nota`
- Candidatos previos al cambio migrados con `rellenarNotasEmpViejas` (deduce A/B/C desde estado: Contratado/Avanza→A, Reserva→B, Descartado→C)

### Accesos rápidos en el nav
- "📋 Perfil del puesto" → PDF (columna `link_perfil` en Config)
- "📢 Flyer" → columna `link_flyer` en Config
- "🔗 Link postulación" → prompt con URL del form para copiar
- Botón volver desde banco: "← Volver al proceso" — **evitar la palabra "Kanban" en la UI**, Ari no la usa como terminología de selección

### Banco de talento
- Curado a mano (botón en la ficha, Ari elige a quién manda — NO automático)
- Clasificación al enviar: categoría amplia (RRHH/Gestión de Personas, Datos/BI, Compras/Abastecimiento, Calidad/SGI, Administración/Finanzas, Operaciones/Producción, Comercial/Ventas, Otros) + seniority (Jr, Ssr, Sr, Coordinador, Supervisor, Jefatura/Gerencia)
- `enviarABanco`: copia la ficha entera con links HYPERLINK, avisa si ya está
- `doGetBanco`, `borrarDelBanco` (verifica nombre antes de borrar)
- Vista independiente (violeta), tarjetas con filtros (nombre/categoría/seniority/zona/formación-experiencia/pretensión), listas fijas con cantidad entre paréntesis
- `enviarABusqueda(filaBanco, perfilDestino)`: copia nombre/tel/mail/zona/formación/resumen/CV/nota_cv a la hoja destino, entra con estado "Avanza a screening" y fuente "Banco de talento", arrastra fórmulas de columnas 11/21/31 desde la fila de arriba

## 11. Formularios (candidato / entrevistador)

- `index.html` (form de competencias): botones "Ajuste al Rol" → Alto ajuste / Ajuste con observaciones / No ajusta (mismas clases `adj-alto/mod/bajo`, no rompe guardado). Campo "Formación" agregado en meta-strip (input vacío `d-formacion`, se completa a mano en la entrevista, NO se carga del Sheet, ubicado entre Situación familiar y Empresa actual)
- `Formulario.html` (screening): mismo campo "Formación" agregado debajo de Empresa/Posición Actual
- `eval-competencias-blumia.html`: metodología STAR, acordeón, webhook a Google Sheets, voz a texto vía Web Speech API

## 12. Pendiente — form público de postulación (`Postulacion.html`)

A mitad de desarrollo, quedó frenado. Diseño decidido:
- Candidato sube CV + responde 3-5 preguntas de screening configurables por perfil (`Postulacion_Preguntas`, mismo patrón que el resto: todo desde Sheet, sin tocar código)
- El CV se procesa con ChatGPT al toque, el prompt recibe CV + respuestas declaradas (para que `nivel_ajuste` cruce ambas fuentes)
- Escribe directo a la pestaña del perfil vía `appendRow` (no Google Forms — generaría pestañas nuevas y rompería la estructura de columnas existente)
- Falta: handler de carga de CV, lógica `appendRow`, prompt GPT de fit, ruta `tipo=postulacion` en `doGet`

## 13. Deuda técnica / bugs conocidos (sin resolver)

- **Alta prioridad, nunca resuelto:** API keys hardcodeadas en el código fuente → hay que moverlas a `PropertiesService`
- `esDuplicado` (Código.gs): dos `Logger.log` referencian variable indefinida `buscado` → puede romper el procesamiento automático de CVs. Borrar esas dos líneas.
- Migración evaluada y no hecha: pasar a **Claude Code + clasp** para editar el proyecto Apps Script directamente en vez de pegar código a mano en el editor web — ahora que estás en Claude Code esto probablemente ya no aplica, pero es contexto de por qué el código históricamente se pegó a mano
- Proyecto dado por "finalizado" en julio 2026, con las dos deudas de arriba explícitamente sin hacer

## 14. Quirks de plataforma (aplican a todo Apps Script de Blumia, no solo MACS)

- WeasyPrint (si se genera PDF con esto en vez de Playwright/Chromium): nada de `display:flex`/`grid` en contenido paginado → usar `display:block` + `column-count`. `break-inside:avoid` solo en items, nunca en contenedores padre
- Drive API: siempre `Drive.Files.create` (v3), nunca `Drive.Files.insert` (v2)
- Funciones duplicadas en distintos `.gs` se pisan en silencio, sin error — siempre revisar antes de debuguear
- Gmail search: `newer_than:Nd` más confiable que `after:unix_timestamp`
- Sheet en locale español → separador de fórmula es **punto y coma**, no coma

## 15. Estilo de trabajo de Ari (para Claude Code)

- No programa — espera código completo y funcional, nunca fragmentos/diffs
- Cambios quirúrgicos, nada tocado en producción sin confirmación explícita
- Explicaciones breves, en español rioplatense informal
- Comunica mucho por voz-a-texto → mensajes con artefactos de transcripción, typos, cosas cortadas
- Se frustra si se le repite la misma pregunta diagnóstica más de una vez — leer el historial antes de preguntar
