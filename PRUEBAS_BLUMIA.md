# Pruebas de Blumia

## 1. Prueba automática del servidor (5 minutos, la corre Claude o Ariel desde el editor)
1. Apps Script de Blumia → ⚙ Configuración del proyecto → Propiedades del script → agregar `BLUMIA_PRUEBAS` = `SI`.
2. Editor → elegir la función `blumiaPruebaCompleta` → ▶ Ejecutar (la primera vez pide permisos).
3. Resultado: pestaña `Pruebas_Blumia` de la planilla (un renglón por paso, ✅/❌ y el error) y el registro de ejecución.
4. La prueba crea la búsqueda «PRUEBA AUTOMÁTICA …», pasa un candidato falso por todo el circuito (postulación, duplicado, CV con IA, estados, contacto, observaciones, agenda, screening y competencias con PDF, informe, flyer, descarte con motivo, papelera, historial, 5 páginas) y **borra todo al final**. La invitación de la cita le llega al mail de quien la corre. No se manda ningún mail a candidatos.
5. Si se corta a la mitad: ejecutar `blumiaLimpiarPruebas` (borra restos de búsquedas de prueba).
6. Correrla después de cada cambio grande y antes de mostrarle Blumia a alguien.

## 2. Lista a mano (10 minutos, con los ojos; lo que el servidor no ve)
Usar una búsqueda real o la de prueba.
- [ ] Ingreso con nombre y apellido; el Inicio carga y muestra las búsquedas.
- [ ] Tablero: carga en menos de 8 s; las tarjetas se arrastran; los filtros funcionan; la ficha abre.
- [ ] Ficha: Contacto ▾ (WhatsApp y Mail abren), Ver CV abre el CV, Agendar crea la cita y aparece en la Agenda.
- [ ] Screening: abrir el formulario, completarlo, «Guardar entrevista» → aparece la nota y el PDF. **Abrir el PDF y mirar la letra (Pangram) y el logo.**
- [ ] Competencias: igual. **Mirar la letra del PDF.**
- [ ] Postulación pública (abrir el link en una ventana privada): se ve el logo y las preguntas; enviar con un CV; llega a la pestaña.
- [ ] Flyer: ⚙ Configurar → Flyer y textos: subir foto, generar textos, «Guardar en Drive»; luego 📚 Búsqueda → Ver y enviar flyer muestra la imagen y **💬 Enviar por WhatsApp abre WhatsApp** (probar en el celular).
- [ ] Celular (390 px): tablero, ficha, Resumen, formularios y el flyer se ven bien, sin cortes ni textos encimados.
- [ ] Resumen: las etapas son Análisis de CV · Screening · Competencias · Entrevista empresa · En reserva · Descartados; todo clickeable.
- [ ] ⏱ Velocidad (⚙ de Administración): el tablero abre en ~5–6 s.
- [ ] Descartar con motivo y mandar el mail de descarte **solo a un candidato de prueba con tu mail**.
