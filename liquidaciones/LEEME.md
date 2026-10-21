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
- Feriado = día con `FDO` en la planilla o en CONFIG `feriados_extra`.
- Fichó de más sin extra declarada → fila `A REVISAR` (no se paga). Declarada sin fichada → `A REVISAR`. Apellido que no coincide con la fichada → `REVISAR IDENTIDAD`.
- Consolidado mensual GRH: reconversión, endoso y tope (CONFIG `tope_extra`, 30).
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
- SADOFE cobran. Supervisores incluidos (nocturnidad y feriados solo excluyen LISTAS).
- Omisión (falta entrada o salida) = no se paga. Salida: `FERIADOS` y `FERIADOS_DETALLE`.

## Días no laborables (CONFIG `no_laborables`)
- En horas extra, la jornada completa fichada se paga como extra: 50% hasta las 13 y 100% después (y de noche).

## Pendiente
SADOFE en día de semana como extra.
