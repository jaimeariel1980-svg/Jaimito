# Caja de obra · NJ Arquitectura (bot Anita)
- `Código.gs`: bot de Telegram + planilla (sin claves; viven en las propiedades del script `TG_TOKEN` y `OA_KEY`).
- `TableroObras.gs` (pantalla + lectura), `TableroAcciones.gs` (cargar/editar/borrar movimientos, presupuesto, obra nueva) y `TableroVista.html` (interfaz). Un `.gs` y un `.html` no pueden llamarse igual en Apps Script.
- Clave del tablero: propiedad del script `TABLERO_PIN` (Configuración del proyecto → Propiedades de la secuencia de comandos). Sin ella el tablero no muestra datos. Se valida en el servidor en cada acción.
- Borrar un movimiento lo copia a la pestaña `Papelera`.
- Pendiente: emitir certificados (IPC) desde el tablero; hoy solo por Anita.
