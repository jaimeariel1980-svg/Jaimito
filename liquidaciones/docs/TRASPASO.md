# TRASPASO — Sistema de Liquidaciones · Policlínico I (PAMI I, Rosario)

> Documento para retomar el proyecto en un chat nuevo sin perder contexto.
> Owner de los archivos: **dotacionpp1@gmail.com**

---

## 1. Objetivo del proyecto

Sistema en **Google Apps Script** que cruza los **horarios programados** contra las **fichadas reales del reloj** y calcula automáticamente las liquidaciones del personal:

- **Horas extra** (con recargos 50% / 100%, autorización de jefatura y de RRHH).
- **Nocturnidad** (horas trabajadas en 21:00–06:00).
- **Feriados** (todavía sin construir).

**El ideal es que cubra TODO EL HOSPITAL, no solo enfermería.** Enfermería fue el piloto. La nómina ya tiene los 757 empleados del Poli en ~60 sectores.

---

## 2. Cómo trabajar con el usuario (Ari)

- **No programa**: edita el código a mano, copiando y pegando en el editor de Apps Script.
- Habla en **español rioplatense**, a menudo por voz a texto (hay typos).
- **NO pasarle el archivo de código completo cada vez** (le quema tokens): pasar **solo la función que cambia**. El archivo completo solo para emergencias (cuando algo se rompió y no vale la pena cazar qué).
- Al pegar reemplazos parciales a veces pierde funciones o deja llaves sueltas → errores "X is not defined" o "Unexpected token }". Solución recurrente: pasarle el archivo completo verificado.
- Respuestas **concisas**. Valora que se **piense y agregue valor**, no que se copie.
- Prefiere validar las reglas con **casos reales** antes de codear lo delicado.

---

## 3. Arquitectura

- Proyecto **Apps Script standalone** (no bound a ningún Sheet): abre las planillas por ID con `SpreadsheetApp.openById`.
- **Google Sheets como base de datos.**
- Los `onEdit`/`onOpen` NO funcionan en standalone; para autopintado se usa **formato condicional** (no triggers).
- Hay un **caché** por ejecución (`_cacheHorario`, `_cacheFich`, `_ssFich`) para leer cada archivo una sola vez y no pasar el límite de 6 minutos.
- El código fuente vive en el proyecto Apps Script del usuario (~26 funciones).

---

## 4. IDs de archivos (AGOSTO 2026)

| Qué | ID |
|---|---|
| **Fichadas** tarjeta reloj (todo el hospital) | `1caGLHLl1faw6X61Cs3BQh4-i97qizmLX2Gpf1tkTliY` |
| **Base de resultados** (pestañas LIQUIDACION, NOCTURNIDAD) | `1dL4DgJTfYbHA7Ra-BulukXHbWFUUxUGPkqVvemOhrsQ` |
| **Excel GRH** (autorizaciones RRHH) | `1DFt_hweW5ahkaMIh454rEG0UuBwOO_yOUhBksZxOtfU` |
| **Nómina** del Poli (Sheet) | `1s3AajJLZJIDWVhqLFiyequaRt-OaZxtssLzheaFSz1Q` |
| Horario **GUARDIA** | `1MhWVXYYoQp2Emk80oPV3zxL35xQhjQDOHBACqCdBF7E` |
| Horario **UTI** | `1mOGXPznXM2LeG6nSZD1VkZUtZ6CX5CetCMsh4Z8wvMo` |
| Horario **2 PISO** | `1rw52e_e60A8qkDPCrbx1sacJ7MhSblJqZjOEnrL4s7I` |
| Horario **3 PISO** | `1BetkVTw-2uSzPwn_SewJIFpItp6NGr5tsFAUW95DcAA` |
| Horario **4 PISO** | `1XGqMDNAWHn91YgSyXigC9JjUN6hvuGnXYR0Fbyz0zlY` |
| Horario **CONSULTORIOS** | `14LtIqcWS_d-Pvd6B_zw3a1Fj2r_NxqXx8jIUAHcOqNk` |

**Servicios que liquidan (6):** GUARDIA, UTI, 2 PISO, 3 PISO, 4 PISO, CONSULTORIOS.
**Supervisión y Sup SADOFE quedan AFUERA** (son jefes, no cobran extra).

---

## 5. Estructura de los datos

### Fichadas (tarjeta reloj)
Una pestaña por día ("1".."31"). Encabezado en fila 1. Marcas del reloj en columnas FICHA_1..FICHA_6 (índices 7–12). **Leer con `getDisplayValues`** (el valor crudo está corrido por zona horaria). Un legajo puede tener varias filas el mismo día (por locación); quedarse con la de más marcas. La salida de un turno noche queda en la pestaña del **día siguiente**.

### Grillas de horario (una por servicio)
Matriz: agentes en filas, días 1..31 en columnas. Encabezado detectado por la fila con más números 1..31. Cada agente ocupa **2+ filas**: la principal (legajo col A, apellido col B, turnos por día) y "filas sombra" debajo (cargo, cambios, extras, segundos turnos).
- Turno formato `H/H`: `6/13`, `23/6`, `17/24`, `0/6`, etc.
- **Colores = datos**: ROJO = cambio de turno; AMARILLO = extra autorizada por jefatura. Se distinguen por **familia de color** (no hex exacto): rojo = R alto, G y B bajos; amarillo = R y G altos, B bajo.
- Códigos no-turno (F, FDO, LAO, PE, ACA, CA, TM/TT…) se ignoran para el cálculo.

### Excel GRH (autorizaciones mensuales de RRHH)
Hoja **"Planillas Sectores"**. Sección **"División Enfermería"** arranca fila 101. Columnas: LEGAJO (col 0), HS 50% (col 7), HS 100% (col 8). Mayoría autorizada **10 al 50% y 6 al 100%**; algunos 12/0. (El "Sector UTI" del Excel son médicos, otro circuito, no se toca.)

### Nómina (todo el Poli)
757 personas, ~60 sectores. Fila 0 = números, fila 1 = encabezado, datos desde fila 2. Columnas (0-based): LEGAJO 0, APELLIDO 1, NOMBRE 2, CARGO 3, Especialidad 4, **AREA 5 (= sector, la que vale)**, CONDICIÓN 6; turnos habituales por día de semana LU-DO = índices 22–28; REGIMEN 29. Sectores grandes: Enfermería 135, Área ambulatoria 52, Guardia 50, Docencia 44, Clínica Quirúrgica 42, Diagnóstico Imágenes 34, Mucamas 30, Clínica Médica 28, UTI 18. Los turnos en la nómina vienen con formato desparejo (`06-13`, `6-13`, `06 - 13`, `07,30-14,30`=7:30-14:30, `17-00`/`17 - 24`=17/24, `F`=franco, `SADOFE`, `11 A 18`, `8-8`) → normalizar al parsear.

---

## 6. Funciones del proyecto (≈26)

`liquidarMes` (motor de extras → pestaña LIQUIDACION), `procesarDia`, `leerTurnosServicio`, `leerFichadasDia`, `leerFichadasDiaSeguro`, `_leerHorarioCrudo` (caché de lectura + detección de encabezado), `leerAutorizadoGRH`, `escribir`, `nocturnidad`, `minEnFranjaNoche`, `parseTurno`, `marcaCerca`, `marcaEnVentana`, `horaToMin`, `minToHora`, `horaRango`, `esFeriado`, `esNoLaborable`, `esHoraAl100`, `calcularExtra`, `esRojo`, `esAmarillo`, `normaliza`, `apellidoCoincide`, `dice` (similitud de bigramas), `unificarHorarios`, `crearPlantilla`.
Globales: `FERIADOS=["7-9","8-17"]`, `NO_LABORABLES=["7-10"]`, `TURNOS_VALIDOS`, `LICENCIAS`, cachés.

---

## 7. Reglas — HORAS EXTRA (motor completo y validado)

- **Extra = trabajado real − turno**, medido por fichada.
- **Hachazo**: piso de **1 hora** (menos no cuenta ni se lista); de ahí bloques de 30 min (1, 1.5, 2…). El reparto 50/100 también en bloques de 30.
- **Recargos**: L-V 50%, salvo nocturno 22:00–06:00 = 100%; sábado 50% hasta 13:00 y 100% después; domingo y feriado 100%.
- **Autorización JEFATURA** (amarillo en la grilla) = techo diario. Se paga lo **realizado** (solapamiento fichada ∩ franja autorizada) con tope en lo autorizado. Sin amarillo pero fichó de más = "extra NO autorizada - revisar" (no se paga).
- **Autorización GRH** (mensual, del Excel) = manda. Consolidado por agente:
  - `exceso50 = max(0, Real50 − Aut50)`; `sobra100 = max(0, Aut100 − Real100)`; `reconversión = min(exceso50, sobra100)` (100→50, más barato, permitido; **nunca 50→100**).
  - `endoso = max(0, TotalReal − TotalAut)` — el endoso pasa **siempre el TOTAL de las hs de un agente a otro** (nunca una parte, nunca de uno a varios).
  - **Tope 30 hs/mes DURO** (50+100 juntas); si supera, se paga 30 y lo demás no se paga. No se recorta primero una ni otra.
  - La alerta puede ser reconversión sola, endoso solo o ambos.
- **Exclusiones**: supervisores (jornada 8h, nunca extra; marcados "SUP" en fila sombra); SADOFE en sus días propios (finde/feriado); feriado no genera extra (se paga como feriado, concepto aparte).
- **SADOFE**: se detectan solos (tienen turnos solo en findes/feriados/no-laborables y ninguno en semana). Si un SADOFE ficha un día de semana sin turno, esa jornada es extra.
- **Identidad**: cruce por legajo + `apellidoCoincide()` tolerante (bigramas, umbral 0.6). Perdona tipeos (YACUZZI≈YACCUZZI) y apellidos partidos (DELLABARCA≈"BARCA DELLA"), pero caza legajos mal cargados (LOMBARDO ≠ Cecilia Alonso).
- **Turnos partidos**: `18/24` (fila) + `0/6` (fila del día siguiente) = un turno `18/06`. `leerTurnosServicio` los une y suprime la cola `0/6` para no duplicar.
- **Salida**: pestaña única **LIQUIDACION**, 18 columnas: Servicio, Legajo, Apellido, Día, Turno, Entrada, Salida, Aut. Jefatura, Franja realizada, Hs 50%, Hs 100%, Estado, Condición, Aut GRH 50%, Aut GRH 100%, Total real mes, A pagar (tope 30), Alerta mensual. Estados: **AUTORIZADA / A REVISAR / REVISAR IDENTIDAD**. Condición: SADOFE, noche, no laborable, realizó menos. Ordenada por apellido + día.
- En agosto solo GUARDIA tiene amarillo cargado; el resto sale "no autorizada" (correcto).

---

## 8. Reglas — NOCTURNIDAD

- Se pagan las horas trabajadas en la franja **21:00–06:00**. No hay 50/100, solo la cantidad de horas por agente. Salida: pestaña NOCTURNIDAD (Legajo, Apellido y Nombre, Hs nocturnas, Alerta).
- **Las extras NO cuentan** para nocturnidad (ya se pagan como extra; sería doble pago). Por eso se calcula sobre el **turno programado**, no sobre lo fichado de más.
- **OMISIÓN = no se paga**: hay que verificar que fichó **entrada Y salida**. Si entró y no marcó salida → omisión (no paga, se marca). Si **no fichó nada → no vino** (se ignora, no es omisión).
- Cuenta lo estrictamente fichado dentro del turno en la franja 21–06 (`minEnFranjaNoche`).
- **Implementado y funcionando para los 6 servicios con horario.** Validado: BUTTORI (62315) ~47.9h.
- **Para el resto del hospital (sin horario)**: se necesita la **duración de jornada** de cada agente (7 / 12 / 24 hs — la duración NO cambia; el horario rota). Modelo acordado: con duración D + entrada fichada + verificar salida → ventana [ent, ent+D], contar 21–06. Ejemplo Corizzo (guardias 24h) = 72h/mes. **Pendiente: conseguir de RRHH la tabla legajo → duración.**

---

## 9. Reglas — FERIADOS (pendiente, sin construir)

Tercera liquidación. Pagar el trabajo en feriados (concepto aparte de la extra). El motor ya sabe qué días son feriados (`FERIADOS`) y quién trabajó (fichada). Falta definir reglas y construir.

---

## 10. Sub-proyecto: PLANILLA MODELO de horarios

Objetivo: una planilla modelo "pro" para todo el Poli, **base única de las 3 liquidaciones**.

Decisiones tomadas:
- **Personal** de la **nómina** (legajo, nombre, sector) — sin errores de legajo.
- **Horario habitual** de las **grillas actuales** (tienen el horario real del mes; la nómina tiene el patrón teórico). El habitual de las grillas solo existe para los 6 servicios convertidos.
- Cada agente en **3 filas**: Habitual / Cambio / Extra. Autopintado por **formato condicional**: Cambio→rojo, Extra→amarillo, Licencia→gris, Habitual→blanco.
- **Cambios y extras se cargan como texto libre** (`13/17`), porque son de duración variable (2, 3, 5, 6 hs). El desplegable de turnos fijos es solo ayuda para el habitual.
- **Selector de sector** (desplegable) — fundamental.
- `LICENCIAS = ["FRANCO","COMPENSATORIO","MATERNIDAD","EXAMEN","ART","DUELO","PARTE MEDICO","LAO"]` (todas mismo color).
- La carga guiada de extras tipo formulario (desde/hasta cada media hora) va en el HTML, no en el Sheet.
- Permisos: un archivo por servicio (permisos Drive reales) + unificada de RRHH de solo lectura. **Sheets NO permite permisos de visualización por pestaña** (por eso archivos separados; IMPORTRANGE tampoco sirve porque no trae colores).

Estado: `crearPlantilla` existe y precarga desde la grilla por color. **Próximo paso: rehacerla para cruzar NÓMINA (personal por sector) + GRILLA (horario), arrancando por GUARDIA de prueba.**

---

## 11. Estado actual

- ✅ **Motor de extras**: completo y validado (6 servicios + consolidado GRH con reconversión/endoso/tope).
- ✅ **Nocturnidad**: funcionando para los 6 servicios con horario (parcial mientras el mes esté en curso; las fichadas de agosto llegan hasta ~día 14).
- ✅ Turnos partidos, optimización de tiempo, cruce de identidad tolerante.
- 🔧 **Plantilla modelo**: en construcción (falta cruzar nómina + grilla, arrancando por Guardia).

---

## 12. Pendientes

1. **Terminar la plantilla modelo** (nómina + grilla, empezar por Guardia, luego replicar).
2. **Liquidación de feriados** (construir).
3. **Nocturnidad del resto del hospital** (tabla legajo → duración de jornada de RRHH).
4. **Escalar a todo el Poli** (sumar grillas del resto de los servicios).
5. **Cara linda (HTML)**: subsistema con roles (supervisor carga → jefe valida → RRHH revalida) y carga guiada de extras tipo formulario.
6. **Cierre de agosto**: correr la nocturnidad completa cuando estén todas las fichadas del mes.
