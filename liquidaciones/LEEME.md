# Liquidaciones PAMI I (horas extra · nocturnidad · feriados)
Apps Script `Liquidaciones PAMI I` (cuenta dotacionpp1). Base de resultados: Sheet `1MErbdbNaRVZULcVQmWdyD9SRT4iVwBK7BckfTEqNgLQ`.
Reglas completas: `docs/CONTEXTO.md`, `docs/TRASPASO.md`. Motor viejo de agosto (referencia): `docs/motor_agosto_referencia.txt`.

## Uso (cada mes)
1. Abrir el Sheet → menú **Liquidaciones → Preparar CONFIG** (crea la hoja CONFIG; no pisa lo ya cargado).
2. Completar en CONFIG: mes, año, `id_horarios` (planilla del mes); `id_fichadas` puede quedar vacío (busca solo "Tarjeta Reloj por persona - <Mes> <Año>" en el Drive), `id_grh` (Excel GRH, opcional), feriados/no laborables si hace falta.
3. **Liquidaciones → Liquidar horas extra** → hoja `LIQUIDACION`.

## Qué hace `liquidarExtras`
- Lee cada pestaña de la planilla de horarios (4 filas por agente: Habitual/Cambio/Extra/Licencias). Excluye las de `hojas_excluidas` (SUP, SUP SADOFE, LISTAS).
- Extra declarada (fila Extra) = techo. Se paga el solapamiento con lo fichado, hachazo (piso 1 h, bloques de 30 min) y recargos (L-V 50%, 22–06 100%, sáb desde 13 h 100%, dom y feriado 100%).
- Fichadas: se muestran tal cual la tarjeta reloj (con segundos); columna "Fichadas del día" con todas las marcas. Jornada = de la primera a la última fichada (las intermedias no cuentan: nunca se trabaja cortado). Una sola fichada = omisión, nunca hay extra.
- Extra declarada que cae dentro del turno (o en la cola de una noche del día anterior) se descuenta.
- **Sin datos del reloj (#N/A en la tarjeta)**: si el agente no figura en la pestaña de un día de la tarjeta reloj (en la Tarjeta sale #N/A; pasa sobre todo con la salida del sábado de un 23 a 06), no es omisión: la fila sale «SIN DATOS DEL RELOJ… no se paga hasta tener el registro» en celeste, en horas extra, nocturnidad y feriados. Si el legajo no figura en ningún día del mes, sale un solo aviso por agente («revisar el legajo»). En la validación del jefe hay la opción «El reloj tiene el registro: pagar la franja».
- Extra de madrugada (0/6) cargada en un día pero hecha la noche siguiente (entró 23:23 del día y salió 6:01 del otro): si no hay fichadas para las 00–06 de ese día y sí para las del día siguiente (sin otra 0/x cargada ahí), se paga esa noche y la fila lo aclara («revisar en qué día va cargada»).
- Tolerancia de minutos: si la entrada o la salida difiere hasta 15' de la franja autorizada, se paga la franja entera (ej. 19:03 a una extra 19/23 = 4 h).
- Horario distinto al programado: si la franja cae dentro del turno de la planilla pero lo fichado está corrido más de 1 h (sin Cambio cargado), se toma el turno real desde la entrada fichada con la misma duración y se paga la franja. Fila en naranja claro «HORARIO DISTINTO…»: cargar el cambio en la planilla.
- Franjas y cambios se leen con «:» o «.» (13.30/17) y en formato «C/H 12 p.m. a 7 p.m.».
- Fin y principio de mes: el sistema busca solo la tarjeta reloj del mes anterior y del siguiente (mismo nombre, otro mes) para la salida del último día y la entrada del día 1. Si no la encuentra, avisa al terminar.
- Feriado = día con `FDO` en la planilla o en CONFIG `feriados_extra`.
- Fichó de más sin extra declarada → fila `A REVISAR` (no se paga). Declarada sin fichada → `A REVISAR`. Apellido que no coincide con la fichada → `REVISAR IDENTIDAD`.
- Consolidado mensual GRH: reconversión, endoso y tope (CONFIG `tope_extra`, 30).
- Reconversión: es **todo el bloque o nada**. Si hizo más al 50 de lo autorizado y no usó nada del 100, se reconvierten todas las horas del 100 al 50 (ej. expte 10/6, hizo 16 al 50 → reconversión 6). Si usó aunque sea una hora del 100, no se reconvierte: todo el excedente va como endoso (pedir a otro agente). Igual al revés (50 → 100).
- Cambios de turno sin informar: se pagan igual y la fila sale marcada «HORARIO DISTINTO AL PROGRAMADO… (cambio de turno sin informar en la planilla)». Incluye a quien no fichó su turno pero sí la extra.
- **Validación del jefe de Enfermería** (`Validacion.gs`), **un solo paso**: al tocar «Liquidar horas extra» se arma en la planilla de horarios la pestaña **«Extras - Validación del jefe»** (no hace falta generar el «Reporte Extras Global»; el nombre empieza con «Extras -» para que el script de Enfermería no la lea). Una fila por cada extra cargada en la fila «Extra» de la planilla (horas calculadas bien, con medias horas) con: Liquidación 50/100 por fichadas, Diferencia, Fichadas del día, Qué se encontró, **Decisión del jefe** (lista) y Comentario. Sin color = coincide; rojo = se paga menos que lo cargado; amarillo = más; verde = ya decidido. Debajo, «EXTRAS NO INFORMADAS» (naranja). A la derecha, el cuadro «CONFORMIDAD DEL JEFE DE ENFERMERÍA» (nombre y «Sí»).
  - Decisiones: extras cargadas con diferencia → *Conforme con la liquidación* / *Omisión justificada: pagar la franja* (n.º de nota en el comentario) / *Corregir la planilla de horarios*; extras no informadas → *Autorizo la extra no informada* / *No autorizo*.
  - Se aplican en la liquidación siguiente: lo pagado por decisión del jefe sale en verde en LIQUIDACION («AUTORIZADA por el jefe…») y cuenta en el expediente.
  - **ACTA_CONFORMIDAD** (planilla de liquidación): estado, totales y cada caso con su decisión. La conformidad queda registrada cuando no quedan decisiones pendientes y el jefe puso su nombre y «Sí»; si aparecen casos nuevos vuelve a «pendiente».
  - Las decisiones también se guardan en la pestaña oculta `_VALID_JEFE`.
- **Extras no informadas** (sin franja en la planilla): con turno, cuando lo fichado supera la duración del turno en 1 h o más (si solo está corrido es cambio de turno, no extra); sin turno (franco, licencia cargada o SADOFE en día de semana), toda la jornada fichada. Salen «A REVISAR» y se pagan solo si el jefe las autoriza. Una licencia escrita como cambio («C/H 7 a.m. a 2 p.m.») se toma como turno.
- Tests del cálculo: `node liquidaciones/docs/test_liquidaciones.js`.

## Resumen del mes
Las columnas del mes (GRH, total, a pagar, alerta) se llenan una sola vez por agente en LIQUIDACION. La pestaña `RESUMEN` tiene una fila por agente: hechas 50/100, autorizado GRH, reconversión, endoso, excede tope, a pagar, filas a revisar; incluye a quien pidió horas en GRH y no tiene extras autorizadas.

## Cambios en las extras
Cada corrida guarda una foto de las extras cargadas (pestaña oculta `_FOTO_EXTRAS`) y anota en `CAMBIOS` lo agregado, quitado o modificado desde la corrida anterior del mismo mes; en rojo los cambios sobre días ya pasados.

## Nocturnidad (menú 3)
- Por agente y día con turno (Habitual/Cambio, licencias vigentes no) que tenga horas entre 21 y 6: se cuentan los minutos **fichados dentro del turno** que caen en 21–06. Las extras no cuentan (se recorta al turno).
- Hace falta entrada y salida: con una sola marca es **omisión** (no se paga, en rojo). Sin ninguna fichada = no vino (no aparece).
- 18/24 + 0/6 del día siguiente = una sola jornada. La salida se busca en la pestaña del día siguiente.
- Salida: `NOCTURNIDAD` (una fila por agente: hs a pagar, jornadas, omisiones) y `NOCTURNIDAD_DETALLE` (una fila por jornada con fichadas tal cual la tarjeta).

## Feriados (menú 4)
- Feriado = día con FDO en la planilla o en CONFIG `feriados_extra`. Se pagan las **horas trabajadas** dentro del día feriado (0 a 24), monto fijo por hora (sin 50/100).
- Con turno: horas fichadas dentro del turno (un 23/6 de la víspera suma de 0 a 6; uno del feriado, de 23 a 24). Lo que excede el turno es extra (si está autorizada) y se liquida en horas extra.
- Con FDO/franco que vino igual: cobra feriado por las horas fichadas. Si además tiene extra cargada ese día → REVISAR (para no pagar dos veces).
- **No cobran feriado**: supervisores (pestañas SUP / SUP SADOFE) ni SADOFE de línea (se detectan solos: todos sus turnos del mes son fin de semana/feriado, al menos 3). Nocturnidad sí incluye supervisores y SADOFE.
- Las extras hechas en feriado se pagan siempre al 100%.
- Omisión (falta entrada o salida) = no se paga. Salida: `FERIADOS` y `FERIADOS_DETALLE`.

## Días no laborables (fijos 2026: 13/11 y 7/12; se pueden sumar en CONFIG `no_laborables`)
- En horas extra, la jornada completa fichada se paga como extra: 50% hasta las 13 y 100% después (y de noche).

## SADOFE en día de semana
- La salida del turno nocturno del fin de semana (ej. 6:08 del lunes) y la entrada a una extra del día siguiente (ej. 23:57 para un 0/6) no cuentan como jornada del día hábil.
- Si un SADOFE ficha un día hábil sin turno: lo fichado es extra. Con franja declarada → AUTORIZADA; sin franja → A REVISAR (extra NO autorizada).

## Pendiente
—
