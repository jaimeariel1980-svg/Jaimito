# Auditoría: Empleos PAMI (la «biblia») vs Empleos Blumia

Fecha: 27/10/2026 · Para: Ariel · Hecha leyendo el código real de `pami/` (Tablero.html 1.178 líneas, Proceso.gs, Codigo.gs, Formularios.gs, Calendario.gs, Contactos.gs, Final.gs, Documentacion.gs, Web.html, LEEME.md) y de `blumia/`.

Criterio: se deja afuera todo lo propio de PAMI que acá no aplica (sindicatos, policlínicos, domicilio y policlínico sugerido, turnos, habilitación para la técnica, jefes de Enfermería, resumen ET, documentación de ingreso). Cada punto tiene un código (V1, F3…) para que puedas decir «hacé V1 a V6».

Leyenda: ✅ ya está · 🟡 parcial · ❌ falta · 🚫 no aplica. Prioridad: 🔴 alta · 🟠 media · 🟡 baja. Esfuerzo: S (horas) · M (un día) · L (varios días).

---

## 1. Resumen en 10 líneas

1. Blumia ya tiene lo más pesado de PAMI: ingreso por mail/Drive/postulación, IA para CVs, screening y competencias con formulario, WhatsApp con plantillas, avisos 🔔, papelera, historial, motivo de descarte, descartes masivos, agenda sincronizada con Calendar, banco de talento.
2. Blumia tiene cosas que PAMI no: multiempresa y multiperfil, grabación y análisis de entrevistas con IA, dictado, mensajes por búsqueda, ⚙ Configurar, editor de formularios, Resumen con realidad salarial, informe para la empresa por búsqueda.
3. **Lo que más se nota es lo visual**: PAMI tiene tarjetas grandes con información en capas, columnas de ancho fijo con encabezado de color, barra de indicadores clickeables que filtran, filtros potentes y una ficha ordenada en alertas / paneles / calificaciones / gestión / historial. Blumia tiene tarjetas chicas y planas.
4. **Lo más importante que le falta a Blumia no es visual: es seguridad y control.** El kanban está abierto a cualquiera que tenga el link (sin usuario ni PIN), no hay copia de seguridad automática y no hay registro de accesos. PAMI tiene usuarios con PIN y roles, copia diaria y auditoría (ver S1 a S4).
5. También faltan datos que alimentan la tarjeta: fecha del último movimiento, «recomendado / espontáneo», «ya se postuló antes», edición de los datos del candidato y reasignar de búsqueda.
6. Plan sugerido en 4 etapas (sección 7): visual del tablero → datos y filtros → agenda y citas desde el tablero → seguridad y control.

---

## 2. Visual: cómo es PAMI y cómo quedaría Blumia

### 2.1 Qué tiene PAMI (medido en el código)
| Elemento | PAMI | Blumia hoy |
|---|---|---|
| Encabezado | Banda grande (≈104 px) con logo, nombre del producto, versión, botones de navegación con **ícono en cuadrado de color** (🔔 naranja, ↻ verde, 📅 azul…), contador de candidatos en píldora, botón rojo «Link postulación» y avatar de usuario | Barra verde chica; ahora ordenada con menús (🔔, selector, 👤, 3 menús) |
| Barra de indicadores (KPI) | 8 tarjetas clickeables **que filtran el tablero**: Todos / policlínicos / En proceso / Citas hoy / Nuevos (7 días) / Contratados, con color por tipo y marca de «filtro activo» | Resumen separado (página aparte); el tablero no tiene indicadores arriba |
| Filtros | Búsqueda, nombre/DNI, zona, formación, experiencia (palabra clave), estado, pretensión hasta $, recomendados / ya se postularon antes / espontáneos, «Limpiar» y contador de resultados | Nombre, zona, formación, experiencia, estado, pretensión, «Limpiar» |
| Columnas | Ancho fijo 250 px, **encabezado de color** (azul; rojo para descartados, dorado para reservas, verde para contratado) con **contador en píldora**, scroll propio por columna, alto = pantalla completa, barra de scroll gruesa | Columnas sin color propio, contador simple |
| Orden | Primero las citas más próximas, después las sin cita por último movimiento | Orden de la planilla |
| Tarjeta | Borde izquierdo de **6 px por nota** (A verde, B+ verde claro, B amarillo, C rojo), nombre grande (15 px, negrita), **insignia de nota con la etapa** arriba a la derecha, línea de formación + etiqueta «★ recomendado / Espontáneo», líneas con ícono separadas por una raya fina: 📍 ubicación + zona, 📅 etiquetas de cita y contacto (con ✕ para borrar), ⚑ alertas (motivo de descarte, ya se postuló antes, sin movimiento, faltantes, analizando CV…) | Tarjeta chica: nota, nombre, zona + pretensión, cita, marca de contacto, formación (nuevo), motivo y avisos (nuevo) |
| Ficha | Modal ancho: **alertas en franja de color arriba**, encabezado con una **fila de etapas** (cada etapa con su letra y subrayado de color), botonera ordenada (formularios, ver postulación, CV, WhatsApp ▾, mail), **paneles** (Contacto / Postulación / Análisis), «Análisis del CV», **Calificaciones** por etapa con botones, Entrevistas (calendario), **Gestión** (reasignar búsqueda, motivos de descarte, observaciones, Guardar) e Historial | Lista de datos en una columna, botones de nota, historial desplegable |
| Estados vacíos y errores | Mensajes explicando qué hacer | Mínimos |

### 2.2 Cómo se vería en Blumia (mismo formato, estilo Blumia)
- Paleta Blumia (verde `#1A6645`, dorado `#F5C842`, coral), misma tipografía de la marca; formas y proporciones de PAMI.
- **V1** 🔴 S/M · **Tarjeta grande en capas**: borde izquierdo de color por nota; nombre 15 px; insignia de nota con la etapa («CV / SCR / COMP / EMP»); línea de formación; líneas con ícono: 📍 zona y etiqueta «en zona / fuera de zona», 💰 pretensión (verde / ámbar / rojo según la banda de la búsqueda), 📅 cita (verde, sólida si es hoy), 💬 contacto con ✕, ⚑ alertas.
- **V2** 🔴 S · **Columnas con encabezado de color** y contador en píldora (descartados en coral, reservas en dorado, contratados en verde), ancho fijo, scroll propio y alto de pantalla completa.
- **V3** 🔴 M · **Barra de indicadores clickeables arriba del tablero** (como PAMI): En proceso · Citas hoy · Nuevos 7 días · Falta contactar · Sin movimiento · Reservas · Descartados; al tocar uno se filtra el tablero y se marca como activo. (El Resumen sigue siendo la página de métricas del informe.)
- **V4** 🟠 S · **Encabezado grande** con íconos de color en cuadrado y contador «N candidatos»; el botón «Link postulación» destacado.
- **V5** 🟠 M · **Ficha ordenada como la de PAMI**: franja de alertas arriba, fila de etapas con letras de colores, botonera uniforme, paneles Contacto / Postulación / Análisis, Calificaciones, Gestión, Historial.
- **V6** 🟠 S · **Orden inteligente**: citas próximas primero, después por último movimiento (necesita T3).
- **V7** 🟡 S · Mensajes de estado vacío y de error entendibles; tooltips en las etiquetas.
- **V8** 🟡 S · Mobile: columnas en carrusel, KPI en 2 columnas (PAMI las pasa a 4 por fila en pantallas chicas).

---

## 3. Funcionalidad: comparación por área

### 3.1 Ingreso y análisis de CVs
| Cód. | PAMI | Blumia | Prio / esf. |
|---|---|---|---|
| I1 | Ingreso por mail, carpeta de Drive y formulario público | ✅ igual (mail con router por asunto/IA, Drive, postulación) | — |
| I2 | **Estado del análisis en la tarjeta** («Analizando CV…», «¿No es un CV?», «Revisar CV») y botón «Re-analizar» | 🟡 el análisis ocurre pero la tarjeta no lo muestra ni se puede reintentar | 🟠 S |
| I3 | **«Ya se postuló antes»** (mismo DNI/mail/teléfono/nombre en otra búsqueda) con alerta en tarjeta y ficha, y limpieza de duplicados | 🟡 solo evita duplicados dentro de la misma búsqueda; el Banco cubre reutilización manual | 🟠 M |
| I4 | **Recomendado por / Espontáneo** (quién lo recomendó) con filtro | ❌ | 🟡 S (útil si Blumia recibe referidos) |
| I5 | **Fuente visible** («Llegó por») en la ficha | 🟡 se guarda; la ficha no la muestra en Postulación | 🟡 S |
| I6 | Ciclo automático cada 10 min | 🟡 cada 30 min | 🟡 S (cambiar frecuencia) |
| I7 | Portada pública con **varias búsquedas abiertas** (`Web.html`) | 🟡 un link de postulación por búsqueda; no hay portada | 🟡 M (opcional) |
| I8 | Mails de error y aviso de postulación fallida a un correo de aviso | ✅ (registra en `Errores` y avisa) | — |

### 3.2 Tablero y filtros
| Cód. | PAMI | Blumia | Prio / esf. |
|---|---|---|---|
| T1 | KPI clickeables que filtran (en proceso, citas hoy, nuevos 7 días, contratados, «falta contactar», «sin movimiento») | ❌ (V3) | 🔴 M |
| T2 | Filtro **recomendados / ya se postularon / espontáneos** | ❌ (depende de I3/I4) | 🟡 S |
| T3 | **Fecha de último movimiento** por candidato y alerta «N días sin movimiento» configurable | ❌ (hoy solo tenemos `Historial`; falta guardar la fecha en la fila) | 🔴 M (columna «Último movimiento» que escribe el historial) |
| T4 | Arrastrar y soltar para mover de etapa | ✅ | — |
| T5 | Desplegables de **Reservas** y **Descartados** por tipo | ✅ | — |
| T6 | Contador «N de M candidatos» con filtros activos | 🟡 (contadorF) | 🟡 S |
| T7 | Botón «↻ Actualizar» y actualización automática cada 3 min con aviso | ✅ | — |
| T8 | Vista de solo lectura | 🚫 | — |

### 3.3 Ficha del candidato
| Cód. | PAMI | Blumia | Prio / esf. |
|---|---|---|---|
| F1 | **Editar los datos del candidato** (nombre, teléfono, mail, DNI, formación, pretensión, observaciones) y regenerar la ficha | ❌ hoy hay que ir al Sheet | 🔴 M |
| F2 | **Reasignar a otra búsqueda** desde la ficha | ❌ (el Banco copia, no mueve) | 🟠 M |
| F3 | **Observaciones** (campo libre) con historial del cambio | 🟡 existe la columna de comentarios pero no se edita desde la ficha | 🟠 S |
| F4 | **Anular una etapa** (con confirmación; la evaluación queda como «anulada», se conserva el PDF y el estado vuelve a «Avanza a …») | ❌ | 🟠 M |
| F5 | Alertas en franja arriba de la ficha (ya se postuló antes, motivo, recomendado, cita, IA, días sin movimiento, etapa faltante) | 🟡 algunas aparecen como etiquetas de la tarjeta | 🟠 S (con V5) |
| F6 | Fila de etapas con la nota de cada una | 🟡 notas sueltas en filas | 🟠 S (con V5) |
| F7 | Ver el CV dentro de la ficha (visor) | 🟡 abre el archivo de Drive en otra pestaña | 🟡 M |
| F8 | Respuestas del formulario de postulación en la ficha | ✅ («Ver postulación») | — |
| F9 | Entrevistas del calendario en la ficha | ✅ (campo «Historial entrevistas») | — |
| F10 | Motivos de descarte múltiples | ✅ (nuevo) | — |
| F11 | Informe final del candidato (todas las etapas, nota ponderada) | ✅ «Generar informe» (con IA, Opus) | — |

### 3.4 Contacto, mensajes y agenda
| Cód. | PAMI | Blumia | Prio / esf. |
|---|---|---|---|
| C1 | WhatsApp con plantillas por etapa, link de agenda correcto, firma de quien envía | ✅ (más: plantillas por búsqueda y firma con el 👤) | — |
| C2 | Marca de contacto con ✕ en la tarjeta y registro | ✅ | — |
| C3 | **Agendar la cita desde el tablero** (elige fecha, hora y duración y crea el evento en Calendar y en la ficha) | ❌ hoy el candidato reserva por link o se carga en Calendar a mano | 🟠 M/L |
| C4 | Vista **📅 Agenda** (citas próximas por día, turnos sin candidato asociado, exportar) | ❌ (lo cubre un aviso y el Resumen, no una vista) | 🟠 M |
| C5 | **Asociar a mano un turno** del calendario con un candidato | ❌ | 🟡 M |
| C6 | **Recordatorios de cita** (aviso 30 min antes) | 🟡 el aviso «entrevista en 30 min» existe en el tablero, no sale a mail/celular | 🟡 M |
| C7 | Contactos de Google automáticos (API) | 🟡 CSV manual (por no pedir un permiso nuevo) | 🟡 M (pide re-autorizar) |
| C8 | Pedido de documentación para el ingreso | 🚫 | — |

### 3.5 Descartes y reservas
| Cód. | PAMI | Blumia | Prio / esf. |
|---|---|---|---|
| D1 | Motivos de descarte obligatorios al poner C | 🟡 se pide al descartar pero se puede «Después» | 🟡 S (decidir si se vuelve obligatorio) |
| D2 | Mail de descarte masivo con vista previa de cantidades | ✅ | — |
| D3 | Tarjeta marca «Descarte enviado» | ✅ (nuevo) | — |
| D4 | Reservas con acordeón por tipo y alerta «en reserva hace N días» | 🟡 acordeón sí, alerta no | 🟡 S |

### 3.6 Avisos
| Cód. | PAMI | Blumia | Prio / esf. |
|---|---|---|---|
| A1 | Pendientes + novedades + campanita + popups, refresco cada 3 min | ✅ | — |
| A2 | Urgentes **una vez por día** y avisos del navegador aunque la pestaña esté minimizada | ❌ | 🟡 M |
| A3 | Más tipos de pendiente: etapa incompleta, en reserva hace N días, sin movimiento | 🟡 | 🟠 S (con F1) |

### 3.7 Historial, auditoría y seguridad
| Cód. | PAMI | Blumia | Prio / esf. |
|---|---|---|---|
| H1 | Historial de quién toca cada cosa | ✅ (nuevo, desde hoy) | — |
| S1 | **Usuarios con nombre y PIN, roles** (administrador, escritor, solo lectura) y sesión recordada | ❌ **el kanban y las páginas internas están abiertos a cualquiera con el link** | 🔴 L (hay que hablarlo: cambia el uso) |
| S2 | **Panel «Control»**: registro de accesos y acciones | ❌ | 🟠 M (depende de S1) |
| S3 | **Copia de seguridad diaria** automática de la planilla | ❌ | 🔴 M |
| S4 | PIN para cambiar datos sensibles | 🟡 solo en perfil del puesto, IA y configuración | 🟠 M |
| S5 | `verificarInstalacion()` que revisa todo | 🟡 `probarClaves` solo revisa las claves | 🟡 S |

### 3.8 Informes
| Cód. | PAMI | Blumia | Prio / esf. |
|---|---|---|---|
| R1 | «📊 Informe» de dirección (resumen del proceso) | ✅ más completo (Informe de proceso por empresa + Resumen) | — |
| R2 | «🔗 Enlaces» (panel de links útiles copiables: postulación, perfil, flyer, agendas) | 🟡 están repartidos en los menús | 🟡 S |
| R3 | Perfil y flyer visibles | ✅ | — |

---

## 4. Lo que Blumia tiene y PAMI no (para no perderlo)
Multiempresa y multiperfil · grabación de la entrevista de Meet con análisis por IA · dictado y «Redactar con IA» · ⚙ Configurar (requisitos de la IA, preguntas, mensajes por búsqueda, links, banda salarial) · editor de formularios · Resumen con realidad salarial y embudo igual al informe · informe de proceso por empresa · banco de talento · exportación de contactos.

## 5. Hallazgos de riesgo en Blumia que conviene saber
1. **Sin login:** cualquiera con el link del kanban ve datos personales (teléfono, mail, CV). Pendiente de antes; PAMI lo resuelve con usuarios y PIN (S1).
2. **Sin copia de seguridad automática** (S3). Un error en el Sheet no se puede deshacer salvo por el historial de versiones de Google.
3. **Casilla `empleos@` llena** (hosting) mientras tanto: no entran CVs por mail.
4. **Límite de 200 versiones** de Apps Script: cada publicación gasta una; conviene agrupar los cambios.
5. Los motivos de descarte y el historial **recién empiezan hoy**: lo anterior no tiene datos.

## 6. Cosas de PAMI que NO conviene copiar
Sindicatos y cupos · policlínicos y su domicilio sugerido · turno preferido · habilitar a un candidato con C para la técnica · vista de jefes de Enfermería y su resumen ET · documentación de ingreso (por ahora).

---

## 7. Plan sugerido (en este orden)

| Etapa | Qué incluye | Esfuerzo | Por qué primero |
|---|---|---|---|
| **1 · Visual del tablero** | V1, V2, V3, V4, V6 (parcial), T6 | 2 días | Es lo que más pediste y no depende de datos nuevos |
| **2 · Datos que alimentan la tarjeta** | T3 (fecha de último movimiento) → A3, V6 · I2 · I3 · F3 · F5/F6/V5 (ficha ordenada) | 3 días | Sin la fecha de movimiento no hay «sin movimiento» ni orden inteligente |
| **3 · Edición y agenda** | F1 (editar candidato) · F2 · F4 · C3 · C4 · C6 | 4 días | Quita idas al Sheet y arma la agenda |
| **4 · Seguridad y control** | S3 (copia diaria) primero, después S1 / S2 / S4 | 3–4 días | Cambia el uso (login): se habla antes |

Mi recomendación: arrancar por la **Etapa 1** y, en paralelo, hacer **S3 (copia de seguridad)** porque es chico y protege todo lo demás.

Para pedir algo: «hacé la etapa 1», o por código («hacé V1, V2 y T3»).


---
## Estado de avance (27/10)
- ✅ Etapa 1 hecha: V1, V2, V3, V4 (píldora de candidatos y botón de postulación), V6, T6.
- ✅ Etapa 2 hecha: T3 (último movimiento desde el Historial), I2 (etiquetas de análisis), I3 (ya se postuló antes), F3 (observaciones), F5/F6/V5 (ficha ordenada), A3 (sin movimiento y reservas viejas).
- ✅ S3 hecho: copia de seguridad diaria + botón «Copia de seguridad ahora».
- ✅ Etapa 3 hecha: F1 (editar datos), F2 (pasar de búsqueda), F4 (anular etapa), C3 (agendar desde el tablero) y C4 (pestaña Agenda). Sin hacer C6 (recordatorios automáticos). ✅ S1/S2/S4 hechos el 27/10 (login con nombre y PIN, roles Administrador/Usuario, pestaña Accesos).
- ✅ 27/10 (auditoría de producto): hechos I4 (campo «Llegó por» editable, etiqueta ⭐ Recomendado y filtro Origen), F7 (CV: quedó un solo «📄 Ver CV» en otra pestaña, también dentro de los formularios), C6 (recordatorio de Google Calendar 30 min antes), A2 (avisos del navegador con la pestaña abierta; urgentes una vez por día) y R2 (🔗 Enlaces en 📚 Búsqueda). I7 ya estaba (pantalla de Inicio).
