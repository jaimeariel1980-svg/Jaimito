function cargarPerfil_EnfermeroPAMI() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const PERFIL = 'enfermero-pami';
  const NOMBRE = 'Enfermero PAMI';
  const EMPRESA = 'PAMI';
  const ROOT_ID = '18bhn7cD2PkcK_SldSfKHrjiieHvO3tYH';
  const ENTREVISTADORES = ['Ariel', 'Marcos', 'Ever'];

  const promptExtra =
`PERFIL: Enfermero/a de piso (Internación general / Guardia). Cliente: PAMI (INSSJP), Policlínicos PAMI I y II, Rosario y zona. Reporta a Supervisor de Enfermería.
MISIÓN: Realizar prácticas de enfermería a pacientes internados en salas generales / guardia / cuidados críticos, garantizando la cobertura de sus necesidades asistenciales dentro de las competencias de la Enfermería Profesional.
FUNCIONES: cuidados sustitutorios y compensatorios; controles de rutina y estado general del paciente; administración de fármacos y procedimientos terapéuticos según indicación médica; seguridad del paciente; asistencia en higiene personal; registro detallado de medicación e insumos; trabajo con médicos, camilleros, mucamos y farmacia; contacto con familiares.

FORMACIÓN (EXCLUYENTE): Enfermería con título TERCIARIO o UNIVERSITARIO COMPLETO (Enfermero/a Profesional o Licenciado/a en Enfermería). Matrícula profesional habilitante vigente. AUXILIAR DE ENFERMERÍA NO ES VÁLIDO: título de auxiliar o curso sin título terciario/universitario completo → NO cumple lo excluyente.
EXPERIENCIA (NO EXCLUYENTE): 3 años en prácticas de enfermería, preferentemente atención a la tercera edad / gerontología. Menos de 3 años NO descarta por sí solo.
CONOCIMIENTOS: prácticas de enfermería, administración de medicación, bioseguridad, registro clínico. Informática moderada (sistemas corporativos, correo). Idioma: no aplica (no evaluar Excel, Power BI ni inglés).
ZONA: Rosario y zona (Policlínicos PAMI). Valorar cercanía y disponibilidad real de traslado.
MODALIDAD: turnos ROTATIVOS. Lunes a viernes turnos de 7 hs; sábados, domingos y feriados turnos de 12 hs. Disponibilidad para rotar es requisito.
COMPETENCIAS CLAVE: compromiso con el paciente/afiliado, escucha empática y buen trato a personas mayores, flexibilidad y adaptabilidad, calidad y seguridad del paciente, trabajo en equipo, compromiso con el aprendizaje.

BANDERA ROJA PRINCIPAL: ALTA ROTACIÓN laboral.
CONTINUIDAD LABORAL (criterio importante para Blumia): analizá el historial. VALORÁ permanencia y crecimiento dentro de cada institución. PENALIZÁ rotación alta con saltos cortos y frecuentes SIN justificación razonable (cierre de la institución, fin de contrato, mudanza, reemplazos). Otras banderas rojas: conflictos reiterados con superiores o equipo, falta de matrícula, antecedentes de maltrato o negligencia con pacientes.
NO uses la edad ni el género del candidato como criterio de evaluación.

EVALUACIÓN DE BORDERLINE — "incluir y señalar": NO descartes mecánicamente a quien no cumple el 100%. Si le falta algo no excluyente (años de experiencia, geriatría) pero tiene FACTORES COMPENSATORIOS (título completo, buena continuidad, formación continua, cursos de RCP/seguridad del paciente), clasificalo como MEDIO y marcá en el resumen "REVISAR:" explicando QUÉ le falta y QUÉ lo compensa. Ante la duda, INCLUIR y señalar.

REGLAS ANTI-DESCARTE (CRÍTICAS):
1. La falta de experiencia en el ámbito específico (geriatría, internación, policlínico) NUNCA es por sí sola motivo de C si tiene el título excluyente y experiencia real en enfermería: baja a B con nota "REVISAR: falta [X], pero compensa con [Y]". C SOLO para: no tener título terciario/universitario completo (ej. auxiliar), no tener experiencia ni formación real en enfermería, o banderas rojas claras.
2. SOBRECALIFICACIÓN: si excede claramente el puesto (ej. jefatura/supervisión, posgrado de gestión) asignale B con nota "REVISAR: perfil por encima del puesto, evaluar encaje y expectativa salarial". Nunca B+ ni A.

NIVEL_AJUSTE:
- Alto: título terciario/universitario completo + matrícula + 3+ años de enfermería (ideal geriatría) + buena continuidad laboral + zona Rosario OK.
- Medio: cumple el título pero le falta algo no excluyente (menos de 3 años, sin geriatría, zona dudosa) O borderline con factores compensatorios → "REVISAR:" en el resumen.
- Bajo: auxiliar o sin título habilitante completo, sin experiencia real en enfermería, alta rotación injustificada marcada, o banderas rojas.`;

  // ---------- 1) Drive + Config ----------
  const config = ss.getSheetByName('Config');
  const cfgData = config.getDataRange().getValues();
  const cfgHeaders = cfgData[0];
  let filaCfg = -1;
  for (let i = 1; i < cfgData.length; i++) if (cfgData[i][0] === PERFIL) { filaCfg = i + 1; break; }

  let folders;
  if (filaCfg === -1) {
    const carpetaPerfil = DriveApp.getFolderById(ROOT_ID).createFolder(NOMBRE);
    folders = [
      carpetaPerfil.createFolder('CVs nuevos').getId(),
      carpetaPerfil.createFolder('CVs procesados').getId(),
      carpetaPerfil.createFolder('Screening PDFs').getId(),
      carpetaPerfil.createFolder('Competencias PDFs').getId()
    ];
    Logger.log('✅ Carpetas Drive: ' + carpetaPerfil.getUrl());
    const fila = new Array(cfgHeaders.length).fill('');
    const set = (h, v) => { const ix = cfgHeaders.indexOf(h); if (ix > -1) fila[ix] = v; };
    set('perfil_id', PERFIL); set('nombre_perfil', NOMBRE); set('sheet_id', SHEET_MASTER_ID); set('hoja_nombre', NOMBRE);
    set('folder_madre', folders[0]); set('folder_recibidos', folders[1]);
    set('folder_pdf_screening', folders[2]); set('folder_pdf_competencias', folders[3]);
    set('prompt_extra', promptExtra); set('empresa', EMPRESA);
    config.appendRow(fila);
    filaCfg = config.getLastRow();
  } else {
    config.getRange(filaCfg, cfgHeaders.indexOf('prompt_extra') + 1).setValue(promptExtra);
    config.getRange(filaCfg, cfgHeaders.indexOf('empresa') + 1).setValue(EMPRESA);
  }
  Logger.log('✅ Config cargada');

  // ---------- 2) Pestaña de candidatos (clon de estructura de coord-rrhh) ----------
  let hoja = ss.getSheetByName(NOMBRE);
  if (!hoja) {
    const modelo = ss.getSheetByName(getConfig('coord-rrhh').hoja_nombre);
    hoja = ss.insertSheet(NOMBRE);
    const lastCol = modelo.getLastColumn();
    modelo.getRange(1, 1, 1, lastCol).copyTo(hoja.getRange(1, 1));
    for (let c = 1; c <= lastCol; c++) hoja.setColumnWidth(c, modelo.getColumnWidth(c));
    hoja.setFrozenRows(1);
    [9, 12, 16].forEach(col => {
      const regla = modelo.getRange(2, col).getDataValidation();
      if (regla) hoja.getRange(2, col, 100, 1).setDataValidation(regla);
    });
    const reglaEntrev = SpreadsheetApp.newDataValidation().requireValueInList(ENTREVISTADORES, true).build();
    hoja.getRange(2, 14, 100, 1).setDataValidation(reglaEntrev);
    hoja.getRange(2, 18, 100, 1).setDataValidation(reglaEntrev);
    Logger.log('✅ Pestaña "' + NOMBRE + '" creada');
  }
  for (let i = 2; i <= 101; i++) {
    hoja.getRange(i, 11).setFormula('=IF(B' + i + '<>"";HYPERLINK("' + URL_DEPLOY + '?perfil=' + PERFIL + '&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(N' + i + ');"Iniciar Screening");"")');
    hoja.getRange(i, 15).setFormula('=IF(B' + i + '<>"";HYPERLINK("' + URL_DEPLOY + '?tipo=competencias&perfil=' + PERFIL + '&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(R' + i + ');"Evaluar competencias");"")');
  }
  Logger.log('✅ Links de screening/competencias');

  const reemplazar = (nombreHoja, filas) => {
    const sh = ss.getSheetByName(nombreHoja);
    const d = sh.getDataRange().getValues();
    for (let i = d.length - 1; i >= 1; i--) if (d[i][0] === PERFIL) sh.deleteRow(i + 1);
    if (filas.length) sh.getRange(sh.getLastRow() + 1, 1, filas.length, filas[0].length).setValues(filas);
    Logger.log('✅ ' + nombreHoja + ': ' + filas.length + ' filas');
  };

  // ---------- 3) Perfiles_Hero ----------
  const hero = ss.getSheetByName('Perfiles_Hero');
  let hHeaders = hero.getRange(1, 1, 1, hero.getLastColumn()).getValues()[0];
  if (hHeaders.indexOf('form_conocimientos') === -1) {
    hero.getRange(1, hHeaders.length + 1).setValue('form_conocimientos');
    hHeaders.push('form_conocimientos');
  }
  const heroVals = {
    perfil_id: PERFIL, empresa: EMPRESA, posicion: 'Enfermero/a de piso', ubicacion: 'Rosario y zona',
    instancia: 'Screening · 1ª entrevista', entrevistadoras: ENTREVISTADORES.join(', '),
    criterio_avanza: '8.5', criterio_reserva: '6.5', form_conocimientos: 'NO'
  };
  const heroFila = hHeaders.map(h => heroVals[h] != null ? heroVals[h] : '');
  const heroData = hero.getDataRange().getValues();
  let filaHero = -1;
  for (let i = 1; i < heroData.length; i++) if (heroData[i][0] === PERFIL) { filaHero = i + 1; break; }
  if (filaHero === -1) hero.appendRow(heroFila);
  else hero.getRange(filaHero, 1, 1, heroFila.length).setValues([heroFila]);
  Logger.log('✅ Perfiles_Hero');

  // ---------- 4) Screening: evaluación (id, etiqueta, peso) ----------
  reemplazar('Perfiles_Screening_Eval', [
    [PERFIL, 'e1', 'Experiencia en prácticas de enfermería (asistencia, medicación, registro)', 3],
    [PERFIL, 'e2', 'Experiencia con adultos mayores / gerontología', 2],
    [PERFIL, 'e3', 'Título terciario/universitario completo y matrícula habilitante', 3],
    [PERFIL, 'e4', 'Conocimiento de bioseguridad y seguridad del paciente', 1],
    [PERFIL, 'e5', 'Continuidad y estabilidad laboral', 2]
  ]);

  // ---------- 5) Screening: condiciones (v1 = remuneración, la usa el cálculo) ----------
  reemplazar('Perfiles_Screening_Val', [
    [PERFIL, 'v1', 'Remuneración pretendida compatible', 'Sí|No|A revisar'],
    [PERFIL, 'v2', 'Disponibilidad para turnos rotativos (L-V 7 hs; sáb/dom/feriados 12 hs)', 'Sí|No'],
    [PERFIL, 'v3', 'Vive en Rosario o zona / traslado resuelto', 'Sí|No'],
    [PERFIL, 'v4', 'Preaviso razonable', 'Sí|No'],
    [PERFIL, 'v5', 'Interés real en la posición', 'Alto|Medio|Bajo'],
    [PERFIL, 'v6', 'Título de Enfermero/a Profesional o Lic. (no auxiliar)', 'Sí|No'],
    [PERFIL, 'v7', 'Matrícula profesional vigente', 'Sí|No|A revisar'],
    [PERFIL, 'v8', 'Sin rotación alta injustificada', 'Sí|No|A revisar']
  ]);

  // ---------- 6) Competencias (id, nombre, descripción, peso, preguntas) ----------
  reemplazar('Perfiles_Competencias', [
    [PERFIL, 'c1', 'Compromiso con el paciente y el afiliado', 'Poner al paciente y su familia en el centro de la atención, sosteniendo calidad y calidez aun con alta carga de trabajo. Evaluar evidencia de cuidado más allá de lo estrictamente indicado.', 2,
      'Contame de un paciente internado al que le dedicaste atención especial más allá de la rutina. ¿Qué hiciste y qué pasó?|Dame un ejemplo de un familiar angustiado o reclamando. ¿Cómo lo acompañaste?|Describí una vez que tuviste que priorizar entre varios pacientes a la vez. ¿Cómo decidiste?'],
    [PERFIL, 'c2', 'Escucha empática y buen trato', 'Comunicarse con respeto, paciencia y empatía con personas mayores y sus familias, incluso con pacientes desorientados o difíciles. Evaluar evidencia de buen trato y manejo de situaciones sensibles.', 2,
      'Contame de un paciente mayor desorientado o poco colaborador. ¿Cómo lo abordaste?|Dame un ejemplo de una noticia difícil que tuviste que acompañar con un paciente o su familia. ¿Qué hiciste?|Describí una situación donde tu forma de comunicarte cambió la evolución o el ánimo de un paciente.'],
    [PERFIL, 'c3', 'Flexibilidad y adaptabilidad', 'Adaptarse a turnos rotativos, guardias, cambios de sala y situaciones imprevistas manteniendo el rendimiento. Evaluar evidencia de disponibilidad y respuesta ante imprevistos.', 2,
      'Contame de una guardia o turno donde todo se complicó (faltantes de personal, urgencias). ¿Cómo te organizaste?|Dame un ejemplo de un cambio de sala, de turno o de procedimiento que tuviste que absorber rápido. ¿Cómo lo manejaste?|Describí cómo sostenés tu rendimiento con turnos rotativos y fines de semana.'],
    [PERFIL, 'c4', 'Calidad y seguridad del paciente', 'Aplicar protocolos, bioseguridad y registro riguroso de medicación e insumos; detectar riesgos y mejorar procesos. Evaluar evidencia de rigor y mejora continua.', 2,
      'Contame de un error o casi-error en medicación o procedimiento que detectaste. ¿Qué hiciste y qué cambió después?|Dame un ejemplo de una mejora que propusiste o implementaste en tu sala o servicio.|Describí cómo llevás el registro de medicación e insumos y cómo evitás omisiones.'],
    [PERFIL, 'c5', 'Trabajo en equipo y colaboración', 'Articular con médicos, camilleros, mucamos, farmacia y otros enfermeros para sostener la continuidad del cuidado. Evaluar evidencia de colaboración y comunicación en el cambio de turno.', 1,
      'Contame de una situación donde dependías de otro sector (farmacia, médicos, camilleros) para atender a un paciente. ¿Cómo lo resolviste?|Dame un ejemplo de un desacuerdo con un compañero o un médico. ¿Cómo lo manejaste?|Describí cómo hacés el pase de turno para asegurar la continuidad del cuidado.']
  ]);

  // ---------- 7) Postulacion_Preguntas (id, etiqueta, tipo, opciones, requerido) ----------
  reemplazar('Postulacion_Preguntas', [
    [PERFIL, 'p1', '¿Tenés título terciario o universitario de Enfermería completo?', 'select', 'Enfermero/a Profesional|Licenciado/a en Enfermería|Auxiliar de Enfermería|En curso', 'SI'],
    [PERFIL, 'p2', '¿Tenés matrícula profesional vigente?', 'select', 'Sí|En trámite|No', 'SI'],
    [PERFIL, 'p3', '¿Cuántos años de experiencia tenés en prácticas de enfermería?', 'select', 'Menos de 1 año|1 a 3 años|3 a 5 años|Más de 5 años', 'SI'],
    [PERFIL, 'p4', '¿Tenés experiencia con adultos mayores (geriatría, internación, domiciliaria)?', 'select', 'Sí, mucha|Sí, alguna|No', 'SI'],
    [PERFIL, 'p5', '¿Tenés disponibilidad para turnos rotativos (L-V de 7 hs y sábados, domingos y feriados de 12 hs)?', 'select', 'Sí|No|A confirmar', 'SI']
  ]);

  // ---------- 8) Mail_Routing ----------
  const mr = ss.getSheetByName('Mail_Routing');
  const mrData = mr.getDataRange().getValues();
  const claves = 'enfermero|enfermera|enfermería|enfermeria|pami|policlínico|policlinico';
  let filaMr = -1;
  for (let i = 1; i < mrData.length; i++) if (mrData[i][0] === PERFIL) { filaMr = i + 1; break; }
  if (filaMr === -1) mr.appendRow([PERFIL, claves]); else mr.getRange(filaMr, 2).setValue(claves);
  Logger.log('✅ Mail_Routing');

  Logger.log('LISTO — ' + PERFIL + ' configurado. Link postulación: ' + URL_DEPLOY + '?tipo=postulacion&perfil=' + PERFIL);
}
