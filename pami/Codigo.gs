const CUENTA = 'dotacionpp1@gmail.com';
const HEAD_CAND = ['Fecha', 'Nombre', 'Teléfono', 'DNI', 'Email', 'Localidad', 'Búsqueda', 'Pretensión', 'Nivel de Ajuste', 'Justificación', 'Resumen', 'Requisitos', 'Link CV', 'Estado', 'Entrevistó', 'Nota screening', 'Nota competencias', 'Recomendó', 'Recomendado por', 'Comentarios', 'Respuestas', 'Procesado', 'Intentos', 'Fuente', 'ID', 'Nota PAMI virtual', 'Nota PAMI presencial', 'Cita', 'Entrevistas (calendario)', 'Contactado', 'Descarte enviado', 'Informe screening', 'Informe competencias', 'Último movimiento'];
const COL = { fecha: 1, nombre: 2, tel: 3, dni: 4, email: 5, loc: 6, busq: 7, pret: 8, ajuste: 9, just: 10, resumen: 11, req: 12, cv: 13, estado: 14, entrev: 15, notaScr: 16, notaComp: 17, recTipo: 18, recPor: 19, coment: 20, resp: 21, proc: 22, intentos: 23, fuente: 24, id: 25, notaP1: 26, notaP2: 27, cita: 28, entrevistas: 29, contactado: 30, descarte: 31, infScr: 32, infComp: 33, mov: 34 };
const NOTAS = ['', 'A', 'B+', 'B', 'C'];
const HEAD_BUSQ = ['id', 'titulo', 'ubicacion', 'modalidad', 'descripcion', 'requisitos', 'estado', 'prompt_ia', 'recomendacion', 'palabras_clave', 'link_perfil'];
const HEAD_PREG = ['busqueda_id', 'id', 'etiqueta', 'tipo', 'opciones', 'requerido'];
const HEAD_EVAL = ['busqueda_id', 'id_eval', 'etiqueta', 'peso'];
const HEAD_VAL = ['busqueda_id', 'id_val', 'etiqueta', 'opciones'];
const HEAD_COMP = ['busqueda_id', 'id_comp', 'nombre_comp', 'descripcion', 'peso', 'preguntas'];
const HEAD_PLANT = ['clave', 'canal', 'texto'];
const HEAD_HIST = ['Fecha', 'ID', 'Candidato', 'Tipo', 'Detalle', 'Quién'];
const HEAD_EVALS = ['Fecha', 'ID', 'Candidato', 'Etapa', 'Entrevistó', 'Nota', 'Score', 'Informe', 'Datos'];
const ESTADOS_COL = [
  ['Análisis de CV', ['Nuevo', 'Contactado', 'Pendiente', 'Revisar'], 'Nuevo'],
  ['Screening', ['Avanza a screening', 'Screening agendado'], 'Avanza a screening'],
  ['Competencias', ['Avanza a competencias', 'Competencias agendada'], 'Avanza a competencias'],
  ['Entrevista PAMI virtual', ['Avanza a PAMI virtual', 'PAMI virtual agendada'], 'Avanza a PAMI virtual'],
  ['Entrevista PAMI presencial', ['Avanza a PAMI presencial', 'PAMI presencial agendada', 'En PAMI'], 'Avanza a PAMI presencial'],
  ['Contratado', ['Contratado'], 'Contratado']
];
const ESTADOS_DESC = ['Descartado sin screening', 'Descartado post screening', 'Descartado en competencias', 'Descartado en PAMI virtual', 'Descartado en PAMI presencial', 'Descartado Fuera de zona', 'Descartado NO le interesa'];
const ESTADOS_RES = ['En reserva', 'Reserva competencias', 'Reserva PAMI virtual', 'Reserva PAMI presencial'];
const ETAPAS = ['cv', 'screening', 'competencias', 'pami1', 'pami2'];
const NOTA_COL = { cv: COL.ajuste, screening: COL.notaScr, competencias: COL.notaComp, pami1: COL.notaP1, pami2: COL.notaP2 };
const AGENDADO = { screening: 'Screening agendado', competencias: 'Competencias agendada', pami1: 'PAMI virtual agendada', pami2: 'PAMI presencial agendada' };
const TIPO_ETAPA = { screening: 'Screening', competencias: 'Competencias', pami1: 'PAMI virtual', pami2: 'PAMI presencial' };

function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }
function ss_() { return SpreadsheetApp.openById(prop_('SHEET_ID')); }
function hoja_(n) { return ss_().getSheetByName(n); }

function ajustes_() {
  const o = {};
  hoja_('Ajustes').getDataRange().getValues().slice(1).forEach(r => { if (r[0]) o[String(r[0]).trim()] = String(r[1] == null ? '' : r[1]).trim(); });
  return o;
}
function lista_(txt, sep) { return String(txt || '').split(sep).map(s => s.trim()).filter(Boolean); }

function nuevoId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 10); }

function asegurarEncabezados_(sh, headers) {
  const n = sh.getLastColumn();
  const actual = n ? sh.getRange(1, 1, 1, n).getValues()[0].map(String) : [];
  if (!actual.some(x => x.trim())) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#034EA2').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
    return;
  }
  let c = actual.length;
  headers.forEach(h => {
    if (actual.indexOf(h) === -1) {
      c++;
      sh.getRange(1, c).setValue(h).setFontWeight('bold').setBackground('#034EA2').setFontColor('#FFFFFF');
    }
  });
}

function asegurarAjustes_(pares) {
  const sh = hoja_('Ajustes');
  const claves = sh.getDataRange().getValues().map(r => String(r[0]).trim());
  const faltan = pares.filter(p => claves.indexOf(p[0]) === -1);
  if (faltan.length) sh.getRange(sh.getLastRow() + 1, 1, faltan.length, 2).setValues(faltan);
}

function setup() {
  const yo = String(Session.getEffectiveUser().getEmail()).toLowerCase();
  if (yo !== CUENTA) throw new Error('Tenés que estar logueado como ' + CUENTA + ' (ahora estás como ' + yo + '). Creá el Sheet y el script desde esa cuenta.');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  props.setProperty('SHEET_ID', ss.getId());
  if (!props.getProperty('ADMIN_KEY')) props.setProperty('ADMIN_KEY', Utilities.getUuid().replace(/-/g, ''));

  const tab = (nombre, headers) => {
    let sh = ss.getSheetByName(nombre);
    const nuevo = !sh;
    if (!sh) sh = ss.insertSheet(nombre);
    asegurarEncabezados_(sh, headers);
    return { sh: sh, nuevo: nuevo };
  };

  const aj = tab('Ajustes', ['clave', 'valor']);
  if (aj.nuevo) {
    aj.sh.getRange(2, 1, 9, 2).setValues([
      ['nombre_sitio', 'Empleos PAMI'],
      ['color', '#034EA2'],
      ['correo_aviso', CUENTA],
      ['entrevistadores', 'Ariel, Marcos, Ever'],
      ['estados', 'Nuevo|Contactado'],
      ['carpeta_cvs', ''],
      ['carpeta_subir', ''],
      ['autorespuesta', 'NO'],
      ['texto_autorespuesta', 'Hola [NOMBRE], recibimos tu CV para la búsqueda de [PUESTO]. Para avanzar, completá tu postulación acá: [LINK]\n¡Gracias!']
    ]);
    aj.sh.setColumnWidth(1, 160); aj.sh.setColumnWidth(2, 520);
  }
  SpreadsheetApp.flush();
  asegurarAjustes_([
    ['nombre_cliente', 'PAMI'], ['firma', 'Equipo de Selección'], ['asunto_mail', 'Proceso de selección · PAMI'],
    ['umbral_a', '8.5'], ['umbral_bplus', '8'], ['umbral_b', '6.5'], ['dias_alerta', '7'],
    ['calendario', ''], ['link_agenda', ''], ['carpeta_informes', ''],
    ['duracion_screening', '30'], ['duracion_competencias', '60'], ['duracion_pami1', '30'], ['duracion_pami2', '60']
  ]);
  SpreadsheetApp.flush();
  const cfg = ajustes_();
  const hojaAj = ss.getSheetByName('Ajustes');
  const filaAj = k => hojaAj.getDataRange().getValues().findIndex(r => r[0] === k) + 1;
  const carpeta = (nombre, padre) => {
    const it = padre ? padre.getFoldersByName(nombre) : DriveApp.getFoldersByName(nombre);
    return it.hasNext() ? it.next() : (padre ? padre.createFolder(nombre) : DriveApp.createFolder(nombre));
  };
  if (!cfg.carpeta_cvs) {
    cfg.carpeta_cvs = carpeta('Empleos PAMI - CVs').getId();
    hojaAj.getRange(filaAj('carpeta_cvs'), 2).setValue(cfg.carpeta_cvs);
  }
  if (!cfg.carpeta_subir) {
    const id = carpeta('Para procesar', DriveApp.getFolderById(cfg.carpeta_cvs)).getId();
    hojaAj.getRange(filaAj('carpeta_subir'), 2).setValue(id);
  }
  if (!cfg.carpeta_informes) {
    const id = carpeta('Informes', DriveApp.getFolderById(cfg.carpeta_cvs)).getId();
    hojaAj.getRange(filaAj('carpeta_informes'), 2).setValue(id);
  }
  SpreadsheetApp.flush();

  const bq = tab('Busquedas', HEAD_BUSQ);
  if (bq.nuevo) {
    bq.sh.getRange(2, 1, 1, 10).setValues([[
      'enfermero', 'Enfermero/a de piso', 'Rosario y zona · Policlínicos PAMI I y II',
      'Turnos rotativos: lunes a viernes de 7 hs; sábados, domingos y feriados de 12 hs',
      'Realizar prácticas de enfermería a pacientes internados en salas generales, guardia y cuidados críticos, garantizando la cobertura de sus necesidades asistenciales. Reporta al Supervisor de Enfermería.\n\nFunciones:\n- Brindar cuidados a los pacientes internados, bajo indicación médica.\n- Realizar controles de rutina y del estado general del paciente.\n- Administrar fármacos y realizar procedimientos terapéuticos según indicación médica.\n- Garantizar la seguridad del paciente durante su permanencia.\n- Asistir en la higiene personal de quienes lo requieran.\n- Registrar la medicación suministrada y los insumos utilizados.\n- Trabajar junto a médicos, camilleros, mucamas y farmacia, y atender a los familiares.',
      'Título terciario o universitario de Enfermería completo (no se admite auxiliar).\nMatrícula profesional habilitante vigente.\nExperiencia de 3 años en prácticas de enfermería, preferentemente con adultos mayores (no excluyente).\nDisponibilidad para turnos rotativos.\nResidir en Rosario o zona.',
      'ABIERTA',
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
- Bajo: auxiliar o sin título habilitante completo, sin experiencia real en enfermería, alta rotación injustificada marcada, o banderas rojas.`,
      'Sindicato:NORTE;UTI;UTERA;APPAMIA;ATE;UPCN;SUTEPA|Director|Empleado',
      'enfermero|enfermera|enfermería|enfermeria|policlínico|policlinico'
    ]]);
    bq.sh.setColumnWidth(5, 400); bq.sh.setColumnWidth(6, 400);
    bq.sh.getRange(2, 7, 200, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['ABIERTA', 'CERRADA'], true).build());
  }

  const pr = tab('Preguntas', HEAD_PREG);
  if (pr.nuevo) {
    pr.sh.getRange(2, 1, 6, 6).setValues([
      ['enfermero', 'p1', '¿Tenés título terciario o universitario de Enfermería completo?', 'select', 'Enfermero/a Profesional|Licenciado/a en Enfermería|Auxiliar de Enfermería|En curso', 'SI'],
      ['enfermero', 'p2', '¿Tenés matrícula profesional vigente?', 'select', 'Sí|En trámite|No', 'SI'],
      ['enfermero', 'p3', '¿Cuántos años de experiencia tenés en prácticas de enfermería?', 'select', 'Menos de 1 año|1 a 3 años|3 a 5 años|Más de 5 años', 'SI'],
      ['enfermero', 'p4', '¿Tenés experiencia con adultos mayores (geriatría, internación, domiciliaria)?', 'select', 'Sí, mucha|Sí, alguna|No', 'SI'],
      ['enfermero', 'p5', '¿Tenés disponibilidad para turnos rotativos (L-V de 7 hs; sábados, domingos y feriados de 12 hs)?', 'select', 'Sí|No|A confirmar', 'SI'],
      ['enfermero', 'p6', '¿Conocés a alguien que trabaje en PAMI? Si es así, ¿quién?', 'texto', '', 'NO']
    ]);
    pr.sh.setColumnWidth(3, 500); pr.sh.setColumnWidth(5, 360);
  }

  const ev = tab('Screening_Eval', HEAD_EVAL);
  if (ev.nuevo) { const v = [['enfermero', 'e1', 'Experiencia en prácticas de enfermería (asistencia, medicación, registro)', 3],
    ['enfermero', 'e2', 'Experiencia con adultos mayores / gerontología', 2],
    ['enfermero', 'e3', 'Título terciario/universitario completo y matrícula habilitante', 3],
    ['enfermero', 'e4', 'Conocimiento de bioseguridad y seguridad del paciente', 1],
    ['enfermero', 'e5', 'Continuidad y estabilidad laboral', 2]]; ev.sh.getRange(2, 1, v.length, 4).setValues(v); ev.sh.setColumnWidth(3, 480); }
  const vl = tab('Screening_Val', HEAD_VAL);
  if (vl.nuevo) { const v = [['enfermero', 'v1', 'Remuneración pretendida compatible', 'Sí|No|A revisar'],
    ['enfermero', 'v2', 'Disponibilidad para turnos rotativos (L-V 7 hs; sáb/dom/feriados 12 hs)', 'Sí|No'],
    ['enfermero', 'v3', 'Vive en Rosario o zona / traslado resuelto', 'Sí|No'],
    ['enfermero', 'v4', 'Preaviso razonable', 'Sí|No'],
    ['enfermero', 'v5', 'Interés real en la posición', 'Alto|Medio|Bajo'],
    ['enfermero', 'v6', 'Título de Enfermero/a Profesional o Lic. (no auxiliar)', 'Sí|No'],
    ['enfermero', 'v7', 'Matrícula profesional vigente', 'Sí|No|A revisar'],
    ['enfermero', 'v8', 'Sin rotación alta injustificada', 'Sí|No|A revisar']]; vl.sh.getRange(2, 1, v.length, 4).setValues(v); vl.sh.setColumnWidth(3, 480); }
  const cp = tab('Competencias', HEAD_COMP);
  if (cp.nuevo) { const v = [['enfermero', 'c1', 'Compromiso con el paciente y el afiliado', 'Poner al paciente y su familia en el centro de la atención, sosteniendo calidad y calidez aun con alta carga de trabajo. Evaluar evidencia de cuidado más allá de lo estrictamente indicado.', 2,
      'Contame de un paciente internado al que le dedicaste atención especial más allá de la rutina. ¿Qué hiciste y qué pasó?|Dame un ejemplo de un familiar angustiado o reclamando. ¿Cómo lo acompañaste?|Describí una vez que tuviste que priorizar entre varios pacientes a la vez. ¿Cómo decidiste?'],
    ['enfermero', 'c2', 'Escucha empática y buen trato', 'Comunicarse con respeto, paciencia y empatía con personas mayores y sus familias, incluso con pacientes desorientados o difíciles. Evaluar evidencia de buen trato y manejo de situaciones sensibles.', 2,
      'Contame de un paciente mayor desorientado o poco colaborador. ¿Cómo lo abordaste?|Dame un ejemplo de una noticia difícil que tuviste que acompañar con un paciente o su familia. ¿Qué hiciste?|Describí una situación donde tu forma de comunicarte cambió la evolución o el ánimo de un paciente.'],
    ['enfermero', 'c3', 'Flexibilidad y adaptabilidad', 'Adaptarse a turnos rotativos, guardias, cambios de sala y situaciones imprevistas manteniendo el rendimiento. Evaluar evidencia de disponibilidad y respuesta ante imprevistos.', 2,
      'Contame de una guardia o turno donde todo se complicó (faltantes de personal, urgencias). ¿Cómo te organizaste?|Dame un ejemplo de un cambio de sala, de turno o de procedimiento que tuviste que absorber rápido. ¿Cómo lo manejaste?|Describí cómo sostenés tu rendimiento con turnos rotativos y fines de semana.'],
    ['enfermero', 'c4', 'Calidad y seguridad del paciente', 'Aplicar protocolos, bioseguridad y registro riguroso de medicación e insumos; detectar riesgos y mejorar procesos. Evaluar evidencia de rigor y mejora continua.', 2,
      'Contame de un error o casi-error en medicación o procedimiento que detectaste. ¿Qué hiciste y qué cambió después?|Dame un ejemplo de una mejora que propusiste o implementaste en tu sala o servicio.|Describí cómo llevás el registro de medicación e insumos y cómo evitás omisiones.'],
    ['enfermero', 'c5', 'Trabajo en equipo y colaboración', 'Articular con médicos, camilleros, mucamos, farmacia y otros enfermeros para sostener la continuidad del cuidado. Evaluar evidencia de colaboración y comunicación en el cambio de turno.', 1,
      'Contame de una situación donde dependías de otro sector (farmacia, médicos, camilleros) para atender a un paciente. ¿Cómo lo resolviste?|Dame un ejemplo de un desacuerdo con un compañero o un médico. ¿Cómo lo manejaste?|Describí cómo hacés el pase de turno para asegurar la continuidad del cuidado.']]; cp.sh.getRange(2, 1, v.length, 6).setValues(v); cp.sh.setColumnWidth(4, 420); cp.sh.setColumnWidth(6, 520); }

  const pl = tab('Plantillas', HEAD_PLANT);
  if (pl.nuevo) { const v = plantillasIniciales_(); pl.sh.getRange(2, 1, v.length, 3).setValues(v); pl.sh.setColumnWidth(3, 620); pl.sh.getRange(2, 3, v.length, 1).setWrap(true); }

  const cd = tab('Candidatos', HEAD_CAND);
  migrarCandidatos_();
  const n = 1000;
  const val = lista => SpreadsheetApp.newDataValidation().requireValueInList(lista, true).build();
  cd.sh.getRange(2, COL.estado, n, 1).setDataValidation(val(todosEstados_()));
  cd.sh.getRange(2, COL.entrev, n, 1).setDataValidation(val(lista_(ajustes_().entrevistadores || '', ',')));
  [COL.ajuste, COL.notaScr, COL.notaComp, COL.notaP1, COL.notaP2].forEach(c => cd.sh.getRange(2, c, n, 1).setDataValidation(val(['A', 'B+', 'B', 'C'])));
  if (cd.nuevo) { cd.sh.setColumnWidth(COL.nombre, 200); cd.sh.setColumnWidth(COL.resumen, 360); }

  tab('Historial', HEAD_HIST);
  tab('Evaluaciones', HEAD_EVALS);
  tab('Errores', ['Fecha', 'Contexto', 'Detalle']);

  const trg = ScriptApp.getProjectTriggers().map(t => t.getHandlerFunction());
  if (trg.indexOf('ciclo') === -1) ScriptApp.newTrigger('ciclo').timeBased().everyMinutes(10).create();
  if (trg.indexOf('alEditar') === -1) ScriptApp.newTrigger('alEditar').forSpreadsheet(ss).onEdit().create();
  const def = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);
  Logger.log('✅ Listo. Cuenta: ' + CUENTA + '. Ahora: 1) cargá OPENAI_API_KEY en Propiedades del script  2) Implementar como aplicación web  3) ejecutá mostrarLinks()');
}

function migrarCandidatos_() {
  const sh = hoja_('Candidatos');
  const n = sh.getLastRow() - 1;
  if (n < 1) return;
  const ids = sh.getRange(2, COL.id, n, 1).getValues();
  const nom = sh.getRange(2, COL.nombre, n, 1).getValues();
  const fec = sh.getRange(2, COL.fecha, n, 1).getValues();
  const mov = sh.getRange(2, COL.mov, n, 1).getValues();
  let cambio = false;
  for (let i = 0; i < n; i++) {
    if (!String(nom[i][0]).trim()) continue;
    if (!String(ids[i][0]).trim()) { ids[i][0] = nuevoId_(); cambio = true; }
    if (!mov[i][0]) { mov[i][0] = fec[i][0] || new Date(); cambio = true; }
  }
  if (cambio) { sh.getRange(2, COL.id, n, 1).setValues(ids); sh.getRange(2, COL.mov, n, 1).setValues(mov); }
}

function todosEstados_() {
  let l = [];
  ESTADOS_COL.forEach(c => { l = l.concat(c[1]); });
  return l.concat(ESTADOS_RES, ESTADOS_DESC);
}

function mostrarLinks() {
  const url = ScriptApp.getService().getUrl();
  if (!url) { Logger.log('Primero implementá la aplicación web (Implementar → Nueva implementación → Aplicación web).'); return; }
  const bs = hoja_('Busquedas').getDataRange().getValues().slice(1).filter(r => r[0]);
  Logger.log('WEB PÚBLICA (portada): ' + url);
  bs.forEach(r => Logger.log('  Postulación "' + r[1] + '": ' + url + '?b=' + r[0]));
  Logger.log('TABLERO (privado, no lo compartas): ' + url + '?tablero=' + prop_('ADMIN_KEY'));
}

function salida_(t, titulo) {
  return t.evaluate().setTitle(titulo).addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function incluir_(n) { return HtmlService.createHtmlOutputFromFile(n).getContent(); }

function doGet(e) {
  const p = (e && e.parameter) || {};
  const aj = ajustes_();
  const seguro = s => (/^[A-Za-z0-9_-]{1,64}$/.test(String(s || '')) ? String(s) : '');
  const url = ScriptApp.getService().getUrl();
  const privado = archivo => {
    if (seguro(p.k || p.tablero) !== prop_('ADMIN_KEY')) return null;
    const t = HtmlService.createTemplateFromFile(archivo);
    t.clave = seguro(p.k || p.tablero); t.sitio = aj.nombre_sitio || 'Empleos'; t.url = url;
    t.cid = seguro(p.c); t.etapa = seguro(p.etapa); t.entre = String(p.e || '').slice(0, 60);
    return t;
  };
  if (p.tablero || p.tipo === 'screening' || p.tipo === 'competencias') {
    const archivo = p.tablero ? 'Tablero' : (p.tipo === 'screening' ? 'Screening' : 'Competencias');
    const t = privado(archivo);
    if (!t) return HtmlService.createHtmlOutput('Acceso no autorizado');
    return salida_(t, (p.tablero ? 'Tablero' : archivo) + ' · ' + t.sitio);
  }
  const t = HtmlService.createTemplateFromFile('Web');
  t.url = url; t.bid = seguro(p.b); t.sitio = aj.nombre_sitio || 'Empleos';
  return salida_(t, t.sitio);
}

function busquedas_() {
  return hoja_('Busquedas').getDataRange().getValues().slice(1).filter(r => r[0]).map(r => {
    const o = {};
    HEAD_BUSQ.forEach((h, i) => o[h] = String(r[i] == null ? '' : r[i]));
    return o;
  });
}

function getPortada() {
  return busquedas_().filter(b => b.estado.toUpperCase() === 'ABIERTA')
    .map(b => ({ id: b.id, titulo: b.titulo, ubicacion: b.ubicacion, modalidad: b.modalidad }));
}

function getBusquedaPublica(id) {
  const b = busquedas_().find(x => x.id === id);
  if (!b) return null;
  const preguntas = hoja_('Preguntas').getDataRange().getValues().slice(1)
    .filter(r => r[0] === id && r[1])
    .map(r => ({ id: String(r[1]), etiqueta: String(r[2] || ''), tipo: String(r[3] || 'texto').toLowerCase(), opciones: String(r[4] || ''), requerido: String(r[5] || '').toUpperCase() === 'SI' }));
  return { id: b.id, titulo: b.titulo, ubicacion: b.ubicacion, modalidad: b.modalidad, descripcion: b.descripcion, requisitos: b.requisitos, abierta: b.estado.toUpperCase() === 'ABIERTA', preguntas: preguntas };
}

function norm_(s) { return String(s || '').toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, ''); }

function remitente_(from) {
  const m = String(from).match(/^\s*"?([^"<]*?)"?\s*<([^>]+)>\s*$/);
  const email = (m ? m[2] : String(from)).trim().toLowerCase();
  let nombre = m ? m[1].trim() : '';
  if (!nombre || nombre.indexOf('@') > -1) nombre = email.split('@')[0].replace(/[._\d]+/g, ' ').trim().replace(/\b\w/g, c => c.toUpperCase());
  return { nombre: nombre || 'Sin nombre', email: email };
}

function buscarBusqueda_(texto) {
  const abiertas = busquedas_().filter(b => b.estado.toUpperCase() === 'ABIERTA');
  const t = norm_(texto);
  const hit = abiertas.find(b => lista_(b.palabras_clave, '|').some(k => t.indexOf(norm_(k)) > -1));
  if (hit) return hit.id;
  return abiertas.length === 1 ? abiertas[0].id : '';
}

function agregarCandidato_(o) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = hoja_('Candidatos');
    const datos = sh.getDataRange().getValues();
    const dni = String(o.dni || '').replace(/\D/g, '');
    const email = String(o.email || '').trim().toLowerCase();
    for (let i = 1; i < datos.length; i++) {
      if (datos[i][COL.busq - 1] !== (o.busq || '')) continue;
      const dniE = String(datos[i][COL.dni - 1] || '').replace(/\D/g, '');
      const mailE = String(datos[i][COL.email - 1] || '').trim().toLowerCase();
      if ((dni && dniE === dni) || (email && mailE === email)) return 'DUPLICADO';
    }
    const file = o.archivo();
    const row = new Array(HEAD_CAND.length).fill('');
    row[COL.fecha - 1] = new Date(); row[COL.nombre - 1] = o.nombre; row[COL.tel - 1] = o.tel || '';
    row[COL.dni - 1] = dni; row[COL.email - 1] = email; row[COL.loc - 1] = o.loc || '';
    row[COL.busq - 1] = o.busq || ''; row[COL.pret - 1] = o.pret || '';
    row[COL.resumen - 1] = 'Procesando…'; row[COL.cv - 1] = file.getUrl();
    row[COL.estado - 1] = 'Nuevo'; row[COL.resp - 1] = JSON.stringify(o.resp || []);
    row[COL.proc - 1] = 'Pendiente'; row[COL.intentos - 1] = 0; row[COL.fuente - 1] = o.fuente;
    row[COL.id - 1] = nuevoId_(); row[COL.mov - 1] = new Date();
    sh.appendRow(row);
    registrarHistorial_(row[COL.id - 1], o.nombre, 'Ingreso', 'Llegó por ' + o.fuente, 'sistema');
    return sh.getLastRow();
  } finally { lock.releaseLock(); }
}

function recibirPostulacion(p) {
  try {
    const b = getBusquedaPublica(p && p.b);
    if (!b || !b.abierta) return 'CERRADA';
    const d = p.datos || {};
    const dni = String(d.dni || '').replace(/\D/g, '');
    const email = String(d.email || '').trim().toLowerCase();
    if (!String(d.nombre || '').trim() || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || !String(d.telefono || '').trim() || dni.length < 6 || !String(d.localidad || '').trim()) throw new Error('Datos incompletos');
    const decl = p.declaradas || {};
    b.preguntas.forEach(q => { if (q.requerido && !String(decl[q.id] || '').trim()) throw new Error('Falta responder: ' + q.etiqueta); });
    if (!p.cvBase64 || String(p.cvBase64).length > 14000000) throw new Error('CV ausente o muy pesado');
    const ext = (String(p.cvNombre || '').split('.').pop() || 'pdf').toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 5);
    if (['pdf', 'doc', 'docx'].indexOf(ext) === -1) throw new Error('Formato de CV no válido');
    const nombre = String(d.nombre).trim();
    const resp = b.preguntas.filter(q => String(decl[q.id] || '').trim()).map(q => ({ q: q.etiqueta, a: String(decl[q.id]).trim() }));
    const mime = ext === 'pdf' ? 'application/pdf' : (ext === 'doc' ? 'application/msword' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');

    const fila = agregarCandidato_({
      busq: b.id, nombre: nombre, email: email, tel: String(d.telefono).trim(), dni: dni, loc: String(d.localidad).trim(),
      pret: String(d.pretension || '').trim(), fuente: 'Web', resp: resp,
      archivo: () => DriveApp.getFolderById(ajustes_().carpeta_cvs).createFile(
        Utilities.newBlob(Utilities.base64Decode(p.cvBase64), mime, nombre.replace(/[^\wÁ-ú ]/g, '').trim().replace(/\s+/g, '_') + '_' + b.id + '.' + ext))
    });
    if (fila === 'DUPLICADO') return 'DUPLICADO';

    try { analizarFila_(fila); } catch (e) { registrarError_('analizar fila ' + fila, e); }
    avisar_(b.titulo, nombre, 'la web');
    return 'OK';
  } catch (err) {
    registrarError_('recibirPostulacion', err);
    throw new Error('No se pudo procesar la postulación');
  }
}

function avisar_(titulo, nombre, via) {
  try {
    const to = ajustes_().correo_aviso;
    if (to) MailApp.sendEmail(to, 'Nueva postulación · ' + (titulo || 'sin búsqueda'), nombre + ' llegó por ' + via + '. Mirala en el tablero.');
  } catch (e) { }
}

function registrarError_(ctx, e) {
  try { hoja_('Errores').appendRow([new Date(), ctx, String(e && e.stack || e)]); } catch (x) { }
}

function textoCV_(url) {
  const m = String(url).match(/[-\w]{25,}/);
  if (!m) return '';
  const file = DriveApp.getFileById(m[0]);
  try {
    if (file.getMimeType() === 'application/vnd.google-apps.document') return DocumentApp.openById(file.getId()).getBody().getText();
    const tmp = Drive.Files.create({ name: 'Temp_CV', mimeType: MimeType.GOOGLE_DOCS }, file.getBlob());
    const texto = DocumentApp.openById(tmp.id).getBody().getText();
    Drive.Files.remove(tmp.id);
    return texto;
  } catch (e) { registrarError_('extraer texto CV', e); return ''; }
}

function analizarFila_(fila) {
  const sh = hoja_('Candidatos');
  const r = sh.getRange(fila, 1, 1, HEAD_CAND.length).getValues()[0];
  const intentos = (Number(r[COL.intentos - 1]) || 0) + 1;
  sh.getRange(fila, COL.intentos).setValue(intentos);
  const revisar = msg => { sh.getRange(fila, COL.just, 1, 3).setValues([[msg, 'Sin análisis automático', '']]); sh.getRange(fila, COL.proc).setValue('Revisar'); };
  try {
    const b = busquedas_().find(x => x.id === r[COL.busq - 1]);
    if (!b) { revisar('Sin búsqueda asignada. Elegila en la ficha y re-analizá.'); return; }
    const texto = textoCV_(r[COL.cv - 1]);
    if (!texto || texto.trim().length < 50) { revisar('No se pudo leer el CV (¿escaneado o imagen?). Revisar a mano.'); return; }
    let decl = [];
    try { decl = JSON.parse(r[COL.resp - 1] || '[]'); } catch (e) { }
    const bloque = ['Nombre: ' + r[COL.nombre - 1], 'Localidad: ' + (r[COL.loc - 1] || 'no indicada'), 'Pretensión salarial: ' + (r[COL.pret - 1] || 'no indicó')]
      .concat(decl.length ? decl.map(x => x.q + ': ' + x.a) : ['(No completó el formulario: evaluá solo con el CV y marcá como "REVISAR" lo que no se pueda verificar, por ejemplo la matrícula.)']).join('\n');
    const prompt = `Sos un analista senior de selección de personal. Analizá esta postulación para la búsqueda "${b.titulo}".

${b.prompt_ia}

Tenés DOS fuentes: (A) los datos DECLARADOS por el candidato y (B) el texto del CV. Cruzalas. Si declara algo que el CV no respalda o contradice, reflejalo en la justificación.

=== DECLARADOS (no inventar) ===
${bloque}
=== FIN DECLARADOS ===

Devolvé SOLO JSON válido con este formato:
{"es_cv":true,"nombre":"Nombre y apellido real del candidato","telefono":"","dni":"solo números","localidad":"","nivel_ajuste":"A/B+/B/C","justificacion":"una sola oración corta con el porqué de la nota","requisitos":{"formacion":"","matricula":"","experiencia":"","zona":"","continuidad":""},"resumen_experiencia":"3-4 oraciones narrativas sobre trayectoria, atributos y actitud, sin repetir el checklist"}
Devolvé es_cv:false SOLO si el texto claramente NO es un currículum (por ejemplo una factura o publicidad). Ante la duda, es_cv:true.
REGLAS DE LA NOTA:
- A: cumple todo, perfil potente sin ruido.
- B+: muy bueno, cumple lo excluyente con una cuestión menor.
- B: cumple lo mínimo con dudas relevantes, o borderline con factores compensatorios (reserva).
- C: no cumple un requisito excluyente o tiene una bandera roja clara.
- zona: "Zona OK" o "Fuera" según la localidad declarada o la del CV.`;
    const resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      headers: { Authorization: 'Bearer ' + prop_('OPENAI_API_KEY') },
      payload: JSON.stringify({ model: 'gpt-4o-mini', response_format: { type: 'json_object' }, messages: [{ role: 'system', content: prompt }, { role: 'user', content: 'CV:\n' + texto.slice(0, 15000) }] })
    });
    const parsed = JSON.parse(resp.getContentText());
    if (parsed.error) throw new Error('OpenAI: ' + parsed.error.message);
    const c = JSON.parse(parsed.choices[0].message.content);
    if (c.es_cv === false) {
      sh.getRange(fila, COL.just, 1, 3).setValues([['La IA no reconoció el adjunto como un CV. Revisalo a mano.', 'Sin análisis automático', '']]);
      sh.getRange(fila, COL.proc).setValue('NoCV');
      return;
    }
    if (String(r[COL.fuente - 1]) !== 'Web') {
      const fill = (col, v) => { v = String(v || '').trim(); if (v && !String(r[col - 1] || '').trim()) sh.getRange(fila, col).setValue(v); };
      const nom = String(c.nombre || '').trim();
      if (nom.split(/\s+/).length >= 2) sh.getRange(fila, COL.nombre).setValue(nom);
      fill(COL.tel, c.telefono);
      const dni = String(c.dni || '').replace(/\D/g, '');
      if (dni.length >= 6) fill(COL.dni, dni);
      fill(COL.loc, c.localidad);
    }
    const req = c.requisitos || {};
    const reqTxt = Object.keys(req).map(k => k.charAt(0).toUpperCase() + k.slice(1) + ': ' + (req[k] || 'N/D')).join(' | ');
    const nota = ['A', 'B+', 'B', 'C'].indexOf(String(c.nivel_ajuste)) > -1 ? String(c.nivel_ajuste) : 'C';
    sh.getRange(fila, COL.ajuste, 1, 4).setValues([[nota, 'IA sugiere ' + nota + ': ' + String(c.justificacion || ''), String(c.resumen_experiencia || ''), reqTxt]]);
    registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Análisis CV', 'La IA sugiere ' + nota, 'sistema');
    sh.getRange(fila, COL.proc).setValue('SI');
  } catch (e) {
    registrarError_('analizar fila ' + fila, e);
    sh.getRange(fila, COL.proc).setValue(intentos >= 3 ? 'Error' : 'Pendiente');
    if (intentos >= 3) sh.getRange(fila, COL.resumen).setValue('No se pudo analizar el CV automáticamente.');
  }
}

function procesarPendientes() {
  const sh = hoja_('Candidatos');
  const d = sh.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < d.length && n < 5; i++) {
    if (d[i][COL.proc - 1] === 'Pendiente' && d[i][COL.nombre - 1]) { analizarFila_(i + 1); n++; }
  }
}

function textoAuto_(aj, nombre, busq) {
  const b = busquedas_().find(x => x.id === busq);
  const link = ScriptApp.getService().getUrl() + (busq ? '?b=' + busq : '');
  return String(aj.texto_autorespuesta || '')
    .replace(/\[NOMBRE\]/g, String(nombre).split(' ')[0]).replace(/\[PUESTO\]/g, b ? b.titulo : 'nuestras búsquedas').replace(/\[LINK\]/g, link);
}

function procesarCorreos() {
  const lbl = GmailApp.getUserLabelByName('PAMI_procesado') || GmailApp.createLabel('PAMI_procesado');
  const aj = ajustes_();
  const hilos = GmailApp.search('in:inbox has:attachment -label:PAMI_procesado -from:me newer_than:14d', 0, 15);
  let n = 0;
  for (const th of hilos) {
    if (n >= 5) break;
    for (const m of th.getMessages()) {
      const rem = remitente_(m.getFrom());
      if (rem.email === CUENTA) continue;
      const adj = m.getAttachments({ includeInlineImages: false }).filter(a => /\.(pdf|docx?)$/i.test(a.getName()) && a.getSize() <= 8 * 1024 * 1024);
      if (!adj.length) continue;
      const cv = adj.find(a => /cv|curric/i.test(a.getName())) || adj[0];
      const busq = buscarBusqueda_(m.getSubject() + ' ' + m.getPlainBody().slice(0, 1500));
      const ext = cv.getName().split('.').pop().toLowerCase();
      const fila = agregarCandidato_({
        busq: busq, nombre: rem.nombre, email: rem.email, fuente: 'Correo', resp: [],
        archivo: () => DriveApp.getFolderById(aj.carpeta_cvs).createFile(cv.copyBlob().setName(rem.nombre.replace(/\s+/g, '_') + '_' + Utilities.formatDate(new Date(), 'GMT-3', 'yyyyMMdd_HHmm') + '.' + ext))
      });
      if (fila !== 'DUPLICADO') {
        try { analizarFila_(fila); } catch (e) { registrarError_('analizar fila ' + fila, e); }
        if (hoja_('Candidatos').getRange(fila, COL.proc).getValue() !== 'NoCV') {
          if (String(aj.autorespuesta).toUpperCase() === 'SI') { try { m.reply(textoAuto_(aj, rem.nombre, busq)); } catch (e) { registrarError_('autorespuesta', e); } }
          avisar_((busquedas_().find(x => x.id === busq) || {}).titulo, rem.nombre, 'correo');
        }
        n++;
      }
      break;
    }
    th.addLabel(lbl);
  }
}

function procesarCarpeta() {
  const aj = ajustes_();
  if (!aj.carpeta_subir || !aj.carpeta_cvs) return;
  const destino = DriveApp.getFolderById(aj.carpeta_cvs);
  const it = DriveApp.getFolderById(aj.carpeta_subir).getFiles();
  let n = 0;
  while (it.hasNext() && n < 5) {
    const f = it.next();
    if (!/\.(pdf|docx?)$/i.test(f.getName()) && f.getMimeType() !== 'application/vnd.google-apps.document') continue;
    const nombre = f.getName().replace(/\.[^.]+$/, '').replace(/[_\-]+/g, ' ').replace(/\b(cv|curriculum|curr[ií]culum|vitae)\b/ig, '').replace(/\s+/g, ' ').trim() || 'Sin nombre';
    const busq = buscarBusqueda_(f.getName());
    const fila = agregarCandidato_({ busq: busq, nombre: nombre, fuente: 'Carpeta', resp: [], archivo: () => { f.moveTo(destino); return f; } });
    if (fila === 'DUPLICADO') continue;
    try { analizarFila_(fila); } catch (e) { registrarError_('analizar fila ' + fila, e); }
    n++;
  }
}

function ciclo() {
  const lock = LockService.getUserLock();
  if (!lock.tryLock(1000)) return;
  try {
    [procesarCorreos, procesarCarpeta, procesarPendientes, syncCalendario].forEach(fn => { try { fn(); } catch (e) { registrarError_(fn.name, e); } });
  } finally { lock.releaseLock(); }
}
