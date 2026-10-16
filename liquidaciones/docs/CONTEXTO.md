# CONTEXTO — Sistema de Liquidaciones · Policlínico I (PAMI I, Rosario)
### Documento de traspaso para construir las liquidaciones en Claude Code

> Objetivo: sobre la **planilla de horarios** (ya construida, formato nuevo), calcular tres liquidaciones:
> **(1) Horas extra · (2) Nocturnidad · (3) Feriados**.
> El ideal es que cubra **todo el hospital**, no solo enfermería.

---

## 0. Entorno y forma de trabajo

- Todo corre en **Google Apps Script**. Sheets es la base de datos.
- El usuario (**Ari**) **no programa**: pega el código a mano. Hablá claro, en español, y pasá el código listo para pegar.
- El proyecto de la planilla de horarios es **bound** a ese Sheet (tiene `onOpen`, menús, `SpreadsheetApp.getActiveSpreadsheet()`).
- Las liquidaciones pueden ir en ese mismo proyecto bound, o en uno standalone que abra los IDs. Definir según convenga.

---

## 1. FUENTE PRINCIPAL — Planilla de horarios (formato nuevo, ya armado)

**Planilla OCTUBRE 2026:** ID `1mxpaPAUGhe-QEtc_asX_xX4bLaq9ujFrgXBlMDmSpOE`

Una **pestaña por servicio**. Pestañas actuales: `GUARDIA`, `Consultorios`, `2 piso`, `3 piso`, `4 piso`, `Terapia`, `SUP`, `SUP SADOFE`, y `LISTAS` (validación). **SUP y SUP SADOFE son supervisores → NO cobran extra, excluir.**

### Estructura de cada pestaña
- **Fila 1**: título `HORARIOS · {SERVICIO} · MM/AAAA`. A la derecha, leyenda de colores (Cambio=rojo, Extra=amarillo, Licencia=gris). *(Nota: alguna pestaña puede no tener la fila de título; detectar la fila de encabezado buscando "Legajo"/los números de día.)*
- **Fila 2 (encabezado)**: `Legajo | Apellido | Tipo | 1 | 2 | … | 31`
- Columnas (1-based): **Legajo = col 1 (A)**, **Apellido = col 2 (B)**, **Tipo = col 3 (C)**, **días 1..31 = cols 4..34 (D..AH)**.
- **Cada agente ocupa 4 filas**, distinguidas por la columna **Tipo**:
  1. `Habitual` — el turno de cada día (legajo y apellido solo van en esta fila).
  2. `Cambio` — cambios de turno de ese día (si hay).
  3. `Extra` — franjas de hora extra de ese día (si hay).
  4. `Licencias` — licencias/novedades de ese día (si hay).
- **Findes y feriados** vienen con fondo gris `#f3f3f3`; los feriados llevan `FDO` en la fila Habitual. Los francos llevan `F`.

> **Ventaja del formato nuevo:** el tipo (habitual/cambio/extra/licencia) está **explícito en la columna C** — ya NO hay que leer colores ni filas sombra como en las grillas viejas. Esto simplifica muchísimo la lectura.

### Valores que aparecen en las celdas
- **Turnos** (formato `H/H`, inicio/fin en horas): `6/13`, `6/14`, `23/6`, `12/19`, `17/24`, `18/06`, `6/18`, `7/14`, `0/6`, `6/12`, `7/12`, `13/17`, `19/23`, `14/22`, `11/19`, `10/18`, `13/18`, `14/19`. A veces con cero adelante (`23/06`) o minutos (`13/17:30`, `7/17:30`). El `18/06`/`6/18` = turnos de 12 h de noche/día.
- **F** = franco · **FDO** = franco por feriado.
- **Extras** (fila Extra): franjas de **duración variable** (2, 3, 5, 6 h), texto libre: `13/17`, `19/23`, `7/12`, `0/6`, `00/06`, `19/22`, `13/17:30`, `7/17:30`.
- **Cambios** (fila Cambio): un turno nuevo (`6/13`, `17/24`, `12/19`) que reemplaza al habitual de ese día; o un código (`ACA`, `LG`, `FC`, `HD`); o un cambio de horario en texto `C/H 7 a.m. a 2 p.m.`
- **Licencias** (fila Licencias): `LAO`, `PM`, `TP`, `FC`, `Enfermedad`, `ACA`, y versiones con vencimiento: `13 LAO al 29/10/2026`, `19 Maternidad al 09/11/2026`, `11 ACA al 02/10/2026`. (El número inicial es el día de inicio; "al DD/MM/AAAA" es el fin.)
- **Pestaña LISTAS** (validación de desplegables): turnos válidos (`7/14,12/19,6/12,12/18,18/24,0/6,23/6,17/24,18/6,6/18,13/17,19/23,7/12,14/19,13/18`) y licencias (`FRANCO,COMPENSATORIO,MATERNIDAD,EXAMEN,ART,DUELO,PARTE MEDICO,LAO`).

### Turno efectivo de un día (para liquidar)
Para cada agente/día: partir del **Habitual**; si hay **Cambio** (turno) ese día, ese reemplaza. La **Extra** es adicional. La **Licencia** marca ausencia (no trabajó). Los **turnos partidos** de noche siguen apareciendo: `18/06` ya viene entero; si vieras `18/24` + `0/6` en días consecutivos, unir a `18/06` y suprimir la cola.

---

## 2. FUENTE — Fichadas del reloj (para validar y para nocturnidad)

**Fichadas AGOSTO (referencia de estructura):** ID `1caGLHLl1faw6X61Cs3BQh4-i97qizmLX2Gpf1tkTliY` *(usar el archivo de fichadas del mes que se liquide).*

- **Una pestaña por día** del mes, nombradas `"1"`..`"31"`.
- Encabezado en fila 1. Marcas del reloj en columnas **FICHA_1..FICHA_6 (índices 7..12, 0-based)**.
- **Leer con `getDisplayValues()`** — el valor crudo está corrido por zona horaria (GMT) y fecha base 1899.
- Un mismo legajo puede tener **varias filas el mismo día** (una por locación); quedarse con la fila que tiene **más marcas**.
- La **salida de un turno noche** queda en la pestaña del **día siguiente** (las marcas dentro de un día están ordenadas por hora, mezclan salida ~06:00 del turno de anoche con entrada ~23:00 del de hoy).
- Columnas útiles: Legajo (0), Agente/nombre (1), Locación (2).

---

## 3. FUENTE — Excel GRH (autorización mensual de RRHH, para el consolidado de extras)

**ID** `1DFt_hweW5ahkaMIh454rEG0UuBwOO_yOUhBksZxOtfU` · hoja **"Planillas Sectores"**.
- Sección **"División Enfermería"** (toda la enfermería junta). Por agente: **LEGAJO (col 0), HS 50% (col 7), HS 100% (col 8)**. Filtrar filas donde AREA contenga "ENFERMER".
- Autorización típica: **10 al 50% y 6 al 100%** (total 16); algunos 12/0.
- (El "Sector UTI" del Excel son médicos, otro circuito, no se toca.)

---

## 4. REGLAS — LIQUIDACIÓN DE HORAS EXTRA

- **Extra = lo trabajado por encima del turno.** En el formato nuevo, la extra **ya viene declarada** en la fila `Extra` (franja H/H). La **fichada valida** que se hizo.
- **Hachazo / mínimo**: piso de **1 hora** (menos de 1 h no se paga ni se lista). De ahí en bloques de **30 min** (1, 1.5, 2, 2.5…), descartando lo que sobra. El reparto 50/100 también en bloques de 30.
- **Recargos**:
  - L–V: **50%**, salvo nocturno **22:00–06:00 = 100%**.
  - Sábado: **50% hasta 13:00**, **100% desde 13:00**.
  - Domingo y **feriado: 100%**.
  *(El reporte actual usa un criterio simple 06–22=50% / 22–06=100%; el definitivo es el de arriba, con sábado/domingo/feriado.)*
- **Autorización JEFATURA** (la fila `Extra` de la planilla) = **techo diario**. Se paga lo **realizado** = solapamiento entre la franja declarada y lo fichado (`realIni = max(declIni, entrada)`, `realFin = min(declFin, salida)`), con tope en lo declarado. Sin declaración pero fichó de más ⇒ *"extra no autorizada – revisar"* (no se paga; la declaración autoriza).
- **Autorización GRH** (Excel, mensual) = **manda** sobre el total del mes. Consolidado por agente:
  - `exceso50 = max(0, Real50 − Aut50)`; `sobra100 = max(0, Aut100 − Real100)`; `reconversión = min(exceso50, sobra100)` (100→50, más barato, permitido; **nunca 50→100**).
  - `endoso = max(0, TotalReal − TotalAut)` — el endoso pasa **siempre el TOTAL de las horas de un agente a otro** (nunca parcial, nunca de uno a varios).
  - **Tope 30 hs/mes DURO** (50+100 juntas); lo que exceda no se paga, ni con endoso. No se recorta primero 50 ni 100: el límite es sobre el total.
  - La alerta puede ser reconversión sola, endoso solo, o ambos.
- **Exclusiones**: supervisores (SUP / SUP SADOFE → jornada 8 h, nunca extra). SADOFE en sus días propios (finde/feriado, régimen de 12 h). Feriado no genera extra (se paga como feriado, concepto aparte).
- **SADOFE**: trabajan solo findes/feriados. Si un SADOFE ficha un día de semana sin turno, esa jornada es extra.
- **Identidad**: cruce por legajo; verificación de apellido tolerante (similitud de bigramas, umbral ~0.6) para perdonar tipeos (YACUZZI≈YACCUZZI) y apellidos partidos (DELLABARCA≈"BARCA DELLA"), pero cazar legajos mal cargados.
- **Salida sugerida**: una planilla **LIQUIDACION**, una fila por extra, con: Servicio, Legajo, Apellido, Día, Turno, Entrada, Salida, Aut. Jefatura (declarada), Franja realizada, Hs 50%, Hs 100%, Estado (`AUTORIZADA`/`A REVISAR`/`REVISAR IDENTIDAD`), Condición (SADOFE/noche/feriado), y columnas mensuales GRH (Aut 50, Aut 100, Total real mes, A pagar tope 30, Alerta). Ordenar por apellido + día.

**Estado actual del código**: `generarReporteExtras()` ya recorre todas las pestañas, lee las filas `Extra`, y arma "Reporte Extras Global" con hs50/hs100 por hora (criterio simple 6–22 / 22–6). **Falta**: hachazo, recargos finos (sáb/dom/feriado), validación contra fichada, y el consolidado mensual GRH (reconversión/endoso/tope).

---

## 5. REGLAS — LIQUIDACIÓN DE NOCTURNIDAD

- Se pagan las **horas trabajadas en 21:00–06:00**. **No hay 50/100**, solo la cantidad de horas por agente. Salida: Legajo, Apellido y Nombre, Hs nocturnas, Alerta.
- Se calcula sobre el **turno programado** (Habitual con Cambio aplicado), **NO** sobre lo fichado de más: **las extras NO cuentan** para nocturnidad (ya se pagan como extra; sería doble pago).
- **OMISIÓN = no se paga**: exige **entrada Y salida** fichadas. Si entró y no marcó salida ⇒ omisión (no paga, marcar). Si **no fichó nada ⇒ no vino** (se ignora, no es omisión).
- Cuenta lo estrictamente fichado dentro del turno que cae en 21–06.
- **Validado en agosto** para los 6 servicios con horario (ej. BUTTORI 62315 ≈ 47.9 h).
- **Para servicios sin turno fijo detectable** (guardias de 24/12 h en otros sectores): lo que NO cambia por agente es la **duración de jornada** (7/12/24 h); con duración + entrada fichada + verificar salida se arma la ventana [ent, ent+D] y se cuenta 21–06. Ejemplo Corizzo (24 h) = 72 h/mes. *(Con la planilla nueva, si el Habitual está completo, se usa el turno directamente; la duración es el plan B.)*

---

## 6. REGLAS — LIQUIDACIÓN DE FERIADOS (a construir)

- Las horas trabajadas en **feriado** se pagan como **feriado** (concepto aparte, NO como extra).
- Los feriados del mes ya se detectan al generar la planilla (Google Calendar `es.ar#holiday`, con confirmación manual del usuario) y quedan marcados con `FDO` + fondo gris. 
- Falta definir con el usuario el detalle del cálculo (qué se paga, recargo, cómo se cruza con fichada) y construirlo.

---

## 7. Estado del código actual (proyecto bound a la planilla)

Variables: `CARPETA_DESTINO_ID`, `URL_SERVIDOR_RRHH` (web app de RRHH que sirve licencias). Funciones:
- **`onOpen()`** — arma el menú.
- **`crearArchivoMesSiguiente()`** — genera la planilla del mes siguiente: copia el archivo, lee el **patrón semanal** del Habitual (turno por día de semana del mes viejo) y lo **expande** al mes nuevo, detecta feriados (Calendar + confirmación), pinta findes/feriados gris `#f3f3f3` y pone `FDO`, limpia Cambio/Extra/Licencias. (Antepone `'` a los turnos con `/` para que no se interpreten como fecha.)
- **`traerLicenciasDelReloj()`** — consulta `URL_SERVIDOR_RRHH?action=getLicencias&mes=MM/AAAA`, y pega las novedades por legajo: los `C/H` (cambios de horario) en la fila `Cambio`, el resto en la fila `Licencias`, con su color.
- **`generarReporteExtras()`** — recorre todas las pestañas, lee filas `Extra`, calcula hs50/hs100 (criterio simple) y arma "Reporte Extras Global".

---

## 8. Lo que falta construir (prioridad para Claude Code)

1. **Liquidación de extras completa** sobre la planilla nueva: tomar las filas `Extra`, aplicar hachazo (piso 1 h, bloques 30), recargos finos (L-V/sáb/dom/feriado, nocturno 100%), **validar contra fichada** (realizado = solapamiento declarada ∩ fichada), y el **consolidado mensual GRH** (reconversión / endoso / tope 30). Salida `LIQUIDACION`.
2. **Liquidación de nocturnidad**: sobre el turno (Habitual+Cambio), exigiendo entrada y salida fichadas (omisión no paga), contando 21–06, sin las extras. Salida `NOCTURNIDAD`.
3. **Liquidación de feriados**: definir reglas y construir.
4. **Escalar a todo el hospital** (sumar sectores a la planilla; la nómina tiene 757 personas en ~60 sectores).
5. (Opcional) **Cara web / roles**: supervisor carga → jefe valida → RRHH revalida; carga guiada de extras tipo formulario.

---

## 9. Convenciones y gotchas

- Turnos con `/` se guardan como texto (anteponer `'`) para que Sheets no los tome como fecha/división.
- Turno que cruza medianoche: `cruza = (fin ≤ ini) || (fin === 24)`; la salida cae en el día siguiente.
- Detectar la fila de encabezado por la fila con más números 1..31 (no todas las pestañas tienen título).
- Días del mes 1..31 → columnas D..AH (4..34). El agente son 4 filas consecutivas; el legajo/apellido solo en la fila `Habitual`.
- Códigos no-turno (F, FDO, LAO, PM, TP, FC, ACA, LG, HD, Enfermedad, Maternidad…) no son turnos; no generan extra por sí solos.
- Al leer fichadas: `getDisplayValues`, quedarse con la fila de más marcas por legajo/día, salida nocturna en el día siguiente.
