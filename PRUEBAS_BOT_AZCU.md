# Pruebas del bot de Azcu (set fijo)

Armado el 28/10 a partir de la auditoría de producto. Sirve para medir el bot **antes y después de cada cambio** en `azcuSistema_`, `azcuChat`, las herramientas o el modelo (`AZCU_MODELO`).

**Corredor (hecho el 28/10, `appscript/PruebasBot.gs`):** en el editor de Apps Script de Azcu, cargar la propiedad del script `AZCU_PRUEBAS = SI` y ejecutar `azcuPruebasBot`. Reemplaza solo las direcciones por propiedades reales de la cartera, llama a `azcuChat` con el historial y el contexto de cada frase (sin confirmar tarjetas), califica ✅ / 🟡 / ❌ con chequeos automáticos y escribe todo en la pestaña `Pruebas_Bot` de la planilla Maestra (frase, respuesta, tarjetas, llamadas, segundos, motivo y un renglón RESUMEN). Si pasa de ~5 minutos se pausa y hay que volver a ejecutarla. Se apaga sola (`AZCU_PRUEBAS = NO`) al terminar: así nadie puede dispararla desde el `/exec`. No corre la 29 (crea un recordatorio) ni la 43 (genera un PDF): probarlas a mano. Los chequeos son automáticos y conservadores: un ❌ hay que mirarlo, puede ser un falso negativo.

Cómo usarlo:
- Correrlo con la clave real, **fuera del `/exec`**. Por ejemplo, una función que se ejecuta desde el editor y que falla si el usuario no es el dueño del script; nunca una función pública que `google.script.run` pueda llamar.
- No confirma tarjetas: solo mira qué responde el bot y qué tarjeta propone.
- Las direcciones son ejemplos: reemplazarlas por propiedades reales de la cartera antes de correrlo (una calle con dos unidades, una calle con «Bv.» o «Av.», dos propiedades en la misma calle).
- Anotar por frase: ✅ bien, 🟡 bien pero con texto robótico o largo, ❌ mal. La meta es 0 ❌ y como máximo 5 🟡.

Ficha abierta = `ctx.fichaId`. «Historial» = mensajes anteriores de la misma charla.

| # | Historial / contexto | Frase del usuario | Qué tiene que pasar |
|---|---|---|---|
| 1 | — | hola! | Saludo corto y cálido, ofrece ayuda. Sin listas ni menú largo. |
| 2 | — | gracias genia | Respuesta de una línea. No llama herramientas. |
| 3 | — | cuánto sale Riobamba 323? | Da el precio de esa propiedad (dato real). No muestra la ficha completa. |
| 4 | — | precio de Forest 5302 (no existe; sí Forest 1234 y Sucre 5302) | «No encuentro Forest 5302», ofrece las parecidas. No inventa precio. |
| 5 | — | contame de Santa Fe 4500 (dos unidades) | Pregunta cuál unidad (piso/dpto), nombrándolas. |
| 6 | — | qué precio tienen Pellegrini 2200 y Córdoba 1015? | Da los dos precios. **No** pregunta piso ni dpto. |
| 7 | — | Oroño 1500 qué etapa tiene? (en la planilla: «Bv. Oroño 1500») | La encuentra y responde la etapa. |
| 8 | — | tenés algo en Riobamba? | Lista las propiedades de esa calle. |
| 9 | Ficha abierta: 27 de Febrero 1450 | y en Riobamba? | Habla de Riobamba, no de la ficha abierta. |
| 10 | «contame de Riobamba 323» → ficha | y el dueño? | Nombre del dueño de Riobamba 323. |
| 11 | «contame de Riobamba 323» → ficha | en la ficha tampoco está el tel? | Consulta y responde el teléfono de Riobamba 323 (o dice que no está cargado). No muestra otra ficha. |
| 12 | Se mostraron las 2 unidades de Santa Fe 4500 | y la otra? | Muestra la unidad que no se mostró. |
| 13 | — | la otra | Pregunta a cuál se refiere (no hay contexto). |
| 14 | Ficha abierta: Mendoza 7201 | mandame la ubicación | Botones de ubicación de Mendoza 7201. |
| 15 | Ficha abierta: Mendoza 7201 | mandame la ubicación de Riobamba 323 | Ubicación de Riobamba 323, no de la ficha abierta. |
| 16 | — | visita en Mendoza 7201, Marcos Tejo, le interesó, quiere volver con la esposa | Tarjeta directa: Visita · Le interesó · observaciones con lo que contó. |
| 17 | — | vino una pareja a ver Sucre 5302 | Pregunta solo lo que falta (nombre y resultado) en una línea. |
| 18 | Tarjeta pendiente: visita Mendoza 7201, ofreció 50000 | no, 54000 | Tarjeta nueva con 54.000, misma propiedad. No muestra la ficha. |
| 19 | Tarjeta pendiente | dejalo, mejor no | No arma otra tarjeta; dice que no confirme la anterior. |
| 20 | — | agendá visita mañana a las 15 en Córdoba 1015 con Ana López | Tarjeta de agendar con fecha AAAA-MM-DD correcta. |
| 21 | — | el 2/11 a las 15 visita en Córdoba 1015 con Ana | Toma 2/11 como 2 de noviembre a las 15 h, no como número de calle. |
| 22 | — | fui con un cliente a ver una casa de la inmobiliaria Bertollo en Pérez Bulnes 6637, le gustó | Visita externa (`proponer_visita_externa`), sin buscar en la cartera. |
| 23 | — | Ana Gómez hizo una oferta de 90 mil por Mendoza 7201 | Tarjeta de propuesta con el monto. |
| 24 | — | rechazaron la propuesta de Forest 1234 | Pide el motivo antes de armar la tarjeta. |
| 25 | — | aceptaron la propuesta de Forest 1234 | Tarjeta; aclara que pasa a Reserva. |
| 26 | — | pasá Sucre 5302 a publicada | Tarjeta de cambio de etapa. |
| 27 | — | bajale el precio a Forest 1234 a 62 mil dólares | Tarjeta de editar precio con USD 62.000. |
| 28 | — | se cayó la visita de mañana en Córdoba 1015 | La ubica y arma la tarjeta de cancelar. No borra. |
| 29 | — | recordame mañana a las 9 llamar a Juan | Recordatorio personal; aclara «es solo para vos». |
| 30 | — | avisá que mañana no hay reunión | Pregunta si es para vos, para una persona o para todo el equipo. |
| 31 | — | decile a Marcos que mañana vaya a Sucre 5302 a las 10 | Tarjeta de aviso a Marcos con el texto y la hora. |
| 32 | — | qué tengo hoy? | Agenda del día (o «no tenés nada»). |
| 33 | — | cómo viene la semana? | Resumen de la semana con datos reales. |
| 34 | — | qué papeles faltan en toda la cartera? | Lista COMPLETA por propiedad, sin «y N más» y sin cortarse a la mitad. |
| 35 | — | pasame las consultas de WhatsApp del finde | Todas, con fecha, nombre, teléfono, propiedad y mensaje; cuántas sin atender. |
| 36 | — | qué está atrasado? | Atrasos agrupados por vendedor. |
| 37 | — | cuánto es la comisión del 4% sobre 95 mil dólares? | **USD 3.800** (sin mandarlo a usar la calculadora). |
| 38 | — | 120 mil menos 5% | 114.000 |
| 39 | — | llamé al 341-5551234 y no atendió | No hace una resta. Pregunta de qué propiedad y quién, o arma la interacción «No responde». |
| 40 | — | 341-5551234 | No responde «-5.550.893». Pregunta qué hace con ese teléfono. |
| 41 | — | mandale la encuesta al cliente de San Martín 1500 | Botones de encuesta para esa propiedad (WhatsApp / copiar). |
| 42 | — | pasame el formulario de captación para mandarle a un dueño | Link del formulario de captación nuevo, sin pedir propiedad. |
| 43 | — | haceme el pdf de la ficha de Mendoza 7201 | PDF de esa ficha. |
| 44 | — | qué es mejor, crédito hipotecario UVA o tradicional? | Máximo 3 líneas, sin listas, ofrece ampliar. |
| 45 | — | redactame un mail para el dueño de Forest 1234 contándole que bajamos el precio | Mail completo, con el dato real de la propiedad. |
| 46 | — | recuperá la visita que borré ayer | Mira la Papelera y arma la tarjeta de restaurar. |
| 47 | — | dame info de mendosa 7201 | Tolera el error de tipeo y encuentra Mendoza 7201 (o pregunta «¿Mendoza 7201?»). |
| 48 | — | che q onda la de forest? | Entiende que es Forest 1234 (si hay una sola en esa calle). |
| 49 | — | cargame que hoy vinieron 3 personas a ver Sucre 5302: Ana (le gustó), Pedro (no le interesó, caro) y Luis (va a hacer una oferta) | Tres tarjetas (o una por persona) con resultado y observaciones correctos para cada uno. |
| 50 | — | (dictado por voz) ehh visita en mendoza al setenta y dos cero uno con marcos tejo le encantó | Entiende Mendoza 7201 y arma la tarjeta. |

## Medir también
- **Tiempo de respuesta** de cada frase (meta: menos de 8 s en las de consulta).
- **Cantidad de llamadas al modelo** por frase (meta: 1–2; más de 4 es una señal de que se enreda).
- **Tono**: que no repita siempre la misma muletilla («¡Dale!», «che») y que no conteste con listas cuando la charla es corta.
