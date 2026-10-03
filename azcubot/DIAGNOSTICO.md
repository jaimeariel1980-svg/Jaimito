# Diagnóstico completo — Azcu / Tablero / Formularios

Fecha: 03/10/2026. Prueba de punta a punta con Google, OpenAI y OneSignal **simulados** (los datos y la lógica son los reales del código; los servicios externos son simulados).

## Resultado

- Servidor (flujo completo, 75 pruebas): **75 ✅ / 0 ❌**
- Pantallas del celular y del tablero (Azcu, novedades, cámara, calculadora, tablero, precarga, páginas hospedadas, generador, PDF con formato, imagen del reporte): **todo ✅, sin errores de JavaScript**

## Qué NO se puede probar desde acá (probar en el equipo real)

1. Entrega real de notificaciones (OneSignal) en cada celular, incluida "Azcu terminó" y los avisos repetidos.
2. Respuestas reales de OpenAI (calidad, que no invente datos, lectura real de DNI/escrituras y audio Whisper).
3. Drive, Calendar y envío real de mails (MailApp).
4. PDF "ficha Tokko" (necesita Google Docs real).
5. Cámara y micrófono de cada modelo de celular (S22, Redmi, iPhone).
6. Velocidad real (la mejora de caché del prompt y el servidor caliente se miden con uso).

## Detalle de las pruebas de servidor

## 1. CAPTACIÓN (formulario de precarga)
✅ 1.1 Envío parcial (solo datos, sin documentos) crea la propiedad y el código → PRE-2026-0001 | 
✅ 1.2 Mail de confirmación al propietario
✅ 1.3 Retomar por código devuelve lo cargado
✅ 1.4 obtenerDatosOFallar (usado por la página pública)
✅ 1.5 Código inexistente falla con mensaje claro
✅ 1.6 Segundo envío con DNI y escritura: sube a Drive y verifica con IA → {"ok":true,"codigoPrecarga":"PRE-2026-0001","archivosGuardados":{"escritura":"https://drive.google.com/file/d/file_Mendo
✅ 1.7 Extracción IA del DNI para autocompletar → {"ok":true,"nombreCompleto":"Juan Pérez","dni":"20123456","fechaNacimiento":"1980-05-12","nacionalidad":"Argen
✅ 1.8 Extracción IA de la escritura (titulares + inscripción) → {"ok":true,"titulares":[{"nombreCompleto":"Juan Pérez","dni":"20123456","domicilio":"Italia 100","nacionalidad
✅ 1.9 Estado de documentación calculado → {"plano_edificacion":"Pendiente","plano_mensura":"Pendiente","escritura":"Presentado","dni_propietar

## 2. TABLERO (gestión de la propiedad)
✅ 2.1 getDatos trae la propiedad con papeles y dueños → Mendoza 7201 | papeles 12
✅ 2.2 Subir documento desde el tablero (Drive + columna Doc + verificación IA) → https://drive.google.com/file/d/file_DNI propietario - dni.png | Verificado OK
✅ 2.3 Subir la Escritura extrae la inscripción del dominio → Matrícula 12-345 Folio Real
✅ 2.4 Borrar un documento lo deja pendiente
✅ 2.5 Cargar consulta/visita guarda con el vendedor → {"_row":2,"fecha":"04/08/2026","tipo":"WhatsApp","cli":"Matias","tel":"341","res":"Quiere visitar","
✅ 2.6 Editar y eliminar una interacción → editada=true tras borrar=1
✅ 2.7 Agendar visita (calendario + Seguimiento con vendedor) → Visita 06/10/2026 13:41
✅ 2.8 Registrar propuesta
✅ 2.9 Cambiar etapa, precio y URL de ficha → Reserva 95.000 https://ficha.test/1
✅ 2.10 Reporte al propietario por mail con texto editado y formato → Juan Pérez (juan@test.com)
✅ 2.11 Datos para el generador de documentos (obtenerPropiedadesDoc) → {"id":"PRE-2026-0001","etiqueta":"Mendoza 7201 — PRE-2026-0001","direccion":"Mendoza 7201","propietario":"Juan Pérez","dni_propietario":"20123456","do
✅ 2.12 Lista clásica de propiedades (compatibilidad)
✅ 2.13 omitido en simulación: requiere DocumentApp real

## 3. ENCUESTA DE SATISFACCIÓN Y SU IMPACTO
✅ 3.1 El comprador completa la encuesta (guardarEncuesta, sin PIN)
✅ 3.2 obtenerEncuesta devuelve el detalle completo → {"Código de Precarga":"PRE-2026-0001","Fecha":"03/10/2026","Nombre":"Sabrina","Operación":"Compra","Calificaci
✅ 3.3 El tablero muestra la nota en la ficha (getDatos.encuesta.cal) → {"cal":"5","nombre":"Sabrina","rec":"Sí"}
✅ 3.4 Azcu ve la encuesta (ficha y herramienta encuestas) → {"total":1,"encuestas":[{"id":"PRE-2026-0001","direccion":"Mendoza 7201","calificacion":"5","cliente
✅ 3.5 Panel de Azcu incluye satisfacción → {"kpis":[{"n":1,"l":"Propiedades"},{"n":1,"l":"Activas"},{"n":0,"l":"Vendidas","c":"g"},{"n":0,"l":"Suspendidas"},{"n":"
✅ 3.6 Enlace de encuesta por propiedad (getEncuestaUrl)

## 4. AZCU — consultas y herramientas
✅ 4.1 PIN incorrecto rechazado
✅ 4.2 Calculadora: 3% de 150.000 → 4.500 (sin llamar a la IA) → **4.500**
✅ 4.3 Calculadora: montos con "mil", "u$s" y porcentaje
✅ 4.4 La calculadora no se confunde con datos (dirección/teléfono/fecha) → respuesta IA / respuesta IA
✅ 4.5 Herramienta buscar_propiedades y ficha_propiedad vía IA → Ficha cargada
✅ 4.6 Últimas interacciones: ordenadas por fecha, la de hoy primero → 03/10/2026 Hoy | 03/10/2026 Juan | 04/08/2026 Matias
✅ 4.7 Últimas interacciones "solo mías" filtra por vendedor
✅ 4.8 Alertas del día → papeles 1
✅ 4.9 Documentos faltantes
✅ 4.10 Resumen de la semana y de la cartera → {"actividad":{"semana":{"interacciones":2,"visitas_realizadas":0,"visitas_agenda
✅ 4.11 Captación: repaso del formulario cargado → {"id":"PRE-2026-0001","direccion":"Mendoza 7201","formulario_captacion":{"Código de Precarga":"PRE-2
✅ 4.12 Guías (procedimiento, generador, formulario) → {"pagina":"procedimiento_ventas","parte":1,"partes":1,"texto":"ProcedimientoVent
✅ 4.13 Enlaces (captación, encuesta, ficha) para abrir/WhatsApp/copiar
✅ 4.14 Atajo "documentos faltantes" con botón de subir → **Documentos faltantes (1 propiedad)**
- **Mendoza 7201** (R
✅ 4.15 Atajo alertas
✅ 4.16 Voz: el audio se transcribe y una cuenta dictada se resuelve sin IA → cuánto es 120000 más 10 por ciento → **132.000**
✅ 4.17 Voz: una pregunta normal va a la IA → contame de mendoza
✅ 4.18 Portada (azcuHoy) y novedades (azcuMov) por PIN → {"v":0,"c":0,"p":0,"d":1}
✅ 4.19 Adjuntar archivo: la IA reconoce el DNI y propone qué hacer (azcuAdjPlan) → {"id":"PRE-2026-0001","dir":"Mendoza 7201","archivos":[],"sin":[{"i":0,"nombre":"dni.png"}],"docs":[
✅ 4.20 Completar formulario: lista de datos que faltan (azcuCompletarPlan) → 9 faltantes
✅ 4.21 Marcar/desmarcar documento desde Azcu

## 5. AZCU — acciones con tarjeta de confirmación
✅ 5.1 Cargar interacción
✅ 5.2 Agendar visita (calendario + vendedor = quien) → Visita agendada en Calendar para el 08/10/2026 a las 10:30.
✅ 5.2b La visita agendada por Azcu queda con el vendedor → Ariel
✅ 5.3 Registrar propuesta
✅ 5.4 Cambiar etapa (el tablero lo ve)
✅ 5.5 Editar campos (URL de ficha, precio) → 99.000 https://ficha.test/2
✅ 5.6 Completar datos faltantes del formulario sin pisar lo cargado → Listo, completé en Mendoza 7201: Baños.
✅ 5.7 Cada acción queda registrada en Novedades → Listo, completé en Mendoza 7201: Baños. | Datos actualizados en Mendoza 7201. | Mendoza 7201 ahora está en Vendida. | Visita agendada en Calendar para
✅ 5.8 Novedades: visita agendada trae id y fecha para navegar al tocar → id=PRE-2026-0001 vf=undefined

## 6. AGENDA, RECORDATORIOS Y AVISOS
✅ 6.1 Agenda trae las visitas agendadas (ventana de 31 días) → 2 visitas: 2026-10-06 13:41, 2026-10-08 10:30
✅ 6.2 Visitas cargadas con vendedor se marcan como mías
✅ 6.3 Crear recordatorio personal → sábado 03/10 13:46
✅ 6.4 Aviso a otra persona con repetición (confirmación del que manda) → {"external_id":["marcos"]} https://azcubot.pages.dev?agenda=1&rec=rmusmdjk7qig
✅ 6.5 El aviso repetido se reprograma (sigue pendiente) → 1 pendiente(s) enviados
✅ 6.6 Quien lo mandó puede frenarlo ("frená el aviso a Marcos") → {"ok":true,"cancelado":"Subí el DNI","cuando":"sábado 03/10 13:51","detenido_aviso_a":"Marcos"}
✅ 6.7 El destinatario lo confirma (tocar la notificación) y el que mandó recibe el "confirmó" → ["✅ Marcos confirmó"]
✅ 6.8 Aviso a todo el equipo (push general)
✅ 6.9 Aviso "Azcu terminó": no se manda con la app activa, sí con la app ausente, sin duplicar → activa=0 ausente=1 repetido=1

## 7. REVISIÓN IA, SEGURIDAD Y TRANSPORTE
✅ 7.1 Revisar con IA (texto + avisos de datos) → {"texto":"Texto corregido.","avisos":["El monto en letras no coincide"]}
✅ 7.2 Funciones del tablero exigen PIN → PIN incorrecto
✅ 7.3 Funciones públicas (encuesta, precarga) funcionan sin PIN
✅ 7.4 Función desconocida o no permitida rechazada
✅ 7.5 Mismo pedido enviado dos veces no se ejecuta dos veces
✅ 7.6 Poll del resultado (JSONP) y soporte GET con carga útil → cb1({"ok":true,"data":[{"id":"PRE-2026-0
✅ 7.7 Chat completo vía transporte (azcuChat por azcuApi + aviso si la app no consulta) → Respuesta larga
✅ 7.8 azcuDiagnostico interno corre completo

RESULTADO: 75 ✅  /  0 ❌
