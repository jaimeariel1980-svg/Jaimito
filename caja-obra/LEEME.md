# Caja de obra · NJ Arquitectura (bot Anita)
- `Código.gs`: bot de Telegram + planilla (sin claves; viven en las propiedades del script `TG_TOKEN` y `OA_KEY`).
- `TableroObras.gs` (pantalla + lectura), `TableroAcciones.gs` (cargar/editar/borrar movimientos, presupuesto, obra nueva) y `TableroVista.html` (interfaz). Un `.gs` y un `.html` no pueden llamarse igual en Apps Script.
- Clave del tablero: propiedad del script `TABLERO_PIN` (Configuración del proyecto → Propiedades de la secuencia de comandos). Sin ella el tablero no muestra datos. Se valida en el servidor en cada acción.
- Borrar un movimiento lo copia a la pestaña `Papelera`.
- Pendiente: emitir certificados (IPC) desde el tablero; hoy solo por Anita.

## Regla de diseño (pedida por Ariel)
Todo número o tarjeta del tablero se puede tocar y lleva al detalle (movimientos filtrados, certificados, ficha de la obra). Si se agrega un número nuevo, tiene que ser tocable.
- Cobrado / Pagado / Saldo / Bonus → Movimientos filtrados. Gastos de estudio → movimientos de esos rubros. Honorarios → Certificados. Ganancia neta total → desglose por obra. Rubro → sus movimientos.
- Certificados: se puede borrar solo el último de cada obra (copia en `Papelera certificados`, y el ingreso que generó a `Papelera`).
