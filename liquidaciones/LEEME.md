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
- Feriado = día con `FDO` en la planilla o en CONFIG `feriados_extra`.
- Fichó de más sin extra declarada → fila `A REVISAR` (no se paga). Declarada sin fichada → `A REVISAR`. Apellido que no coincide con la fichada → `REVISAR IDENTIDAD`.
- Consolidado mensual GRH: reconversión, endoso y tope (CONFIG `tope_extra`, 30).
- Tests del cálculo: `node liquidaciones/docs/test_liquidaciones.js`.

## Pendiente
Nocturnidad, feriados, SADOFE en día de semana.
