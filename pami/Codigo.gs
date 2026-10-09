const CUENTA = 'dotacionpp1@gmail.com';
const HEAD_CAND = ['Fecha', 'Nombre', 'Teléfono', 'DNI', 'Email', 'Localidad', 'Búsqueda', 'Pretensión', 'Nivel de Ajuste', 'Justificación', 'Resumen', 'Requisitos', 'Link CV', 'Estado', 'Entrevistó', 'Nota screening', 'Nota competencias', 'Recomendó', 'Recomendado por', 'Comentarios', 'Respuestas', 'Procesado', 'Intentos', 'Fuente', 'ID', 'Nota PAMI virtual', 'Nota entrevista técnica', 'Cita', 'Entrevistas (calendario)', 'Contactado', 'Descarte enviado', 'Informe screening', 'Informe competencias', 'Último movimiento', 'Zona', 'Formación', 'Postulación', 'Policlínico', 'Domicilio', 'Policlínico sugerido', 'Domicilio revisado', 'Motivo de descarte'];
const COL = { fecha: 1, nombre: 2, tel: 3, dni: 4, email: 5, loc: 6, busq: 7, pret: 8, ajuste: 9, just: 10, resumen: 11, req: 12, cv: 13, estado: 14, entrev: 15, notaScr: 16, notaComp: 17, recTipo: 18, recPor: 19, coment: 20, resp: 21, proc: 22, intentos: 23, fuente: 24, id: 25, notaP1: 26, notaP2: 27, cita: 28, entrevistas: 29, contactado: 30, descarte: 31, infScr: 32, infComp: 33, mov: 34, zona: 35, formacion: 36, postulacion: 37, policlinico: 38, domicilio: 39, polSug: 40, domRev: 41, motivo: 42 };
const MOTIVOS_DESCARTE = ['Fuera de zona', 'Fuera de edad', 'Falta de experiencia', 'Formación no adecuada', 'Remuneración fuera de rango'];
const POLICLINICOS = ['Policlínico PAMI 1', 'Policlínico PAMI 2'];
const NOTAS = ['', 'A', 'B+', 'B', 'C'];
const HEAD_BUSQ = ['id', 'titulo', 'ubicacion', 'modalidad', 'descripcion', 'requisitos', 'estado', 'prompt_ia', 'recomendacion', 'palabras_clave', 'link_perfil', 'link_flyer'];
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
  ['Entrevista técnica', ['Avanza a entrevista técnica', 'Entrevista técnica agendada'], 'Avanza a entrevista técnica'],
  ['Contratado', ['Contratado'], 'Contratado']
];
const ESTADOS_DESC = ['Descartado sin screening', 'Descartado post screening', 'Descartado en competencias', 'Descartado en entrevista técnica', 'Descartado Fuera de zona', 'Descartado NO le interesa'];
const ESTADOS_RES = ['En reserva', 'Reserva competencias', 'Reserva entrevista técnica'];
const ETAPAS = ['cv', 'screening', 'competencias', 'pami2'];
const NOTA_COL = { cv: COL.ajuste, screening: COL.notaScr, competencias: COL.notaComp, pami2: COL.notaP2 };
const AGENDADO = { screening: 'Screening agendado', competencias: 'Competencias agendada', pami2: 'Entrevista técnica agendada' };
const TIPO_ETAPA = { screening: 'Screening', competencias: 'Competencias', pami2: 'Entrevista técnica' };

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
    ['nombre_cliente', 'PAMI'], ['firma', 'Equipo de Atracción y Elección'], ['asunto_mail', 'Proceso de selección · PAMI'],
    ['umbral_a', '8.5'], ['umbral_bplus', '8'], ['umbral_b', '6.5'], ['dias_alerta', '7'],
    ['calendario', ''], ['url_web', ''], ['link_agenda', ''], ['link_agenda_screening', ''], ['link_agenda_competencias', ''], ['link_agenda_pami2', ''], ['carpeta_informes', ''], ['calendario_poli2', 'dotacionpp2@gmail.com'], ['link_agenda_screening_poli2', ''], ['link_agenda_competencias_poli2', ''], ['link_agenda_pami2_poli2', ''],
    ['duracion_screening', '30'], ['duracion_competencias', '60'], ['duracion_pami2', '60'],
    ['direccion_policlinico_1', 'Sarmiento 373, Rosario, Santa Fe, Argentina'], ['direccion_policlinico_2', 'Olive 1159, Rosario, Santa Fe, Argentina']
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

  if (!pl.nuevo) {
    const nuevas = plantillasIniciales_();
    const filasPl = pl.sh.getDataRange().getValues();
    filasPl.forEach((f, i) => {
      if (i === 0 || !/^contacto-/.test(String(f[0]))) return;
      if (/te quedan cómodos\? \[LINK\]|para agendarla\. \[LINK\]/.test(String(f[2]))) {
        const n = nuevas.find(x => x[0] === f[0]);
        if (n) pl.sh.getRange(i + 1, 3).setValue(n[2]);
      }
    });
  }

  const cd = tab('Candidatos', HEAD_CAND);
  migrarCandidatos_();
  const n = 1000;
  const val = lista => SpreadsheetApp.newDataValidation().requireValueInList(lista, true).build();
  cd.sh.getRange(2, COL.estado, n, 1).setDataValidation(val(todosEstados_()));
  cd.sh.getRange(2, COL.policlinico, n, 1).setDataValidation(val(POLICLINICOS));
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

const ESQUEMA = '18';

function actualizarEsquema_() {
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('ESQUEMA') === ESQUEMA) return;
  const sh = hoja_('Candidatos');
  const pasos = [
    () => [[COL.policlinico, 'Policlínico'], [COL.domicilio, 'Domicilio'], [COL.polSug, 'Policlínico sugerido'], [COL.domRev, 'Domicilio revisado'], [COL.motivo, 'Motivo de descarte'], [COL.notaP2, 'Nota entrevista técnica']].forEach(x => {
      const c = sh.getRange(1, x[0]);
      const v = String(c.getValue()).trim();
      if (!v || (x[0] === COL.notaP2 && v === 'Nota PAMI presencial')) c.setValue(x[1]).setFontWeight('bold').setBackground('#034EA2').setFontColor('#FFFFFF');
    }),
    () => {
      const ss = ss_();
      let u = ss.getSheetByName('Usuarios');
      if (!u) u = ss.insertSheet('Usuarios');
      asegurarEncabezados_(u, ['Nombre', 'PIN', 'Activo', 'Notas', 'Rol']);
      u.getRange(2, 4).setValue('Los usuarios los da de alta el administrador desde el tablero (avatar → Usuarios). Activo = NO bloquea el acceso.');
      u.getRange(2, 3, 50, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['SI', 'NO'], true).build());
      CacheService.getScriptCache().remove('USUARIOS');
    },
    () => {
      const ss = ss_();
      let a = ss.getSheetByName('Accesos');
      if (!a) a = ss.insertSheet('Accesos');
      asegurarEncabezados_(a, ['Fecha', 'Usuario', 'Acción', 'Detalle']);
      if (!ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'copiaDiaria')) ScriptApp.newTrigger('copiaDiaria').timeBased().everyDays(1).atHour(3).create();
    },
    () => asegurarAjustes_([['zona_max_km', '40']]),
    () => repararLocalidades_(),
    () => asegurarAjustes_([['sesion_inactividad_horas', '8'], ['sesion_recordar_dias', '30']]),
    () => asegurarAjustes_([['calendario_poli2', 'dotacionpp2@gmail.com'], ['link_agenda_screening_poli2', ''], ['link_agenda_competencias_poli2', ''], ['link_agenda_pami2_poli2', '']]),
    () => asegurarAjustes_([['direccion_policlinico_1', 'Sarmiento 373, Rosario, Santa Fe, Argentina'], ['direccion_policlinico_2', 'Olive 1159, Rosario, Santa Fe, Argentina']]),
    () => {
      const pl = hoja_('Plantillas'), filas = pl.getDataRange().getValues();
      plantillasIniciales_().forEach(t => {
        const i = filas.findIndex(f => f[0] === t[0]);
        if (i === -1) { pl.appendRow(t); pl.getRange(pl.getLastRow(), 3).setWrap(true); }
        else if (t[0] === 'contacto-screening' && PropertiesService.getScriptProperties().getProperty('PLANT_SCR') !== '2') {
          pl.appendRow(['contacto-screening-anterior', 'WhatsApp', filas[i][2]]);
          pl.getRange(i + 1, 3).setValue(t[2]);
        }
      });
      PropertiesService.getScriptProperties().setProperty('PLANT_SCR', '2');
      const todas = pl.getDataRange().getValues();
      todas.forEach((f, i) => { if (i > 0 && /equipo de Selección/.test(String(f[2]))) pl.getRange(i + 1, 3).setValue(String(f[2]).replace(/equipo de Selección/g, 'equipo de Atracción y Elección')); });
      const aj = hoja_('Ajustes'), da = aj.getDataRange().getValues();
      da.forEach((f, i) => { if (String(f[0]).trim() === 'firma' && String(f[1]).trim() === 'Equipo de Selección') aj.getRange(i + 1, 2).setValue('Equipo de Atracción y Elección'); });
    },
    () => {
      const n = Math.max(1, Math.min(1000, sh.getMaxRows() - 1));
      const val = lista => SpreadsheetApp.newDataValidation().requireValueInList(lista, true).build();
      sh.getRange(2, COL.estado, n, 1).setDataValidation(val(todosEstados_()));
      sh.getRange(2, COL.policlinico, n, 1).setDataValidation(val(POLICLINICOS));
      sh.getRange(2, COL.motivo, n, 1).setDataValidation(val(MOTIVOS_DESCARTE));
    },
    () => {
      const sh = hoja_('Ajustes'), d = sh.getDataRange().getValues();
      const i = d.findIndex(r => String(r[0]).trim() === 'link_agenda_screening');
      if (i > 0 && !String(d[i][1] || '').trim()) sh.getRange(i + 1, 2).setValue('https://calendar.app.google/5681GAipvQ1VFkWf6');
      const j = d.findIndex(r => String(r[0]).trim() === 'link_agenda_screening_poli2');
      if (j > 0 && !String(d[j][1] || '').trim()) sh.getRange(j + 1, 2).setValue('https://calendar.app.google/fgkSfx6QzC7JnzuA6');
    }
  ];
  let ok = true;
  pasos.forEach(f => { try { f(); } catch (e) { ok = false; registrarError_('actualizarEsquema', e); } });
  if (ok) props.setProperty('ESQUEMA', ESQUEMA);
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
  const url = urlBase_();
  if (!url) { Logger.log('Primero implementá la aplicación web (Implementar → Nueva implementación → Aplicación web).'); return; }
  const bs = hoja_('Busquedas').getDataRange().getValues().slice(1).filter(r => r[0]);
  Logger.log('WEB PÚBLICA (portada): ' + url);
  bs.forEach(r => Logger.log('  Postulación "' + r[1] + '": ' + url + '?b=' + r[0]));
  Logger.log('TABLERO (privado, no lo compartas): ' + url + '?tablero=' + prop_('ADMIN_KEY'));
}

function salida_(t, titulo) {
  return t.evaluate().setTitle(titulo).addMetaTag('viewport', 'width=device-width, initial-scale=1').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function urlBase_() {
  const manual = (ajustes_().url_web || '').trim();
  return manual || ScriptApp.getService().getUrl();
}

function incluir_(n) { return HtmlService.createHtmlOutputFromFile(n).getContent(); }

function doGet(e) {
  try {
    return doGet_(e);
  } catch (err) {
    registrarError_('doGet', err);
    return HtmlService.createHtmlOutput('<div style="font-family:Arial;padding:30px;max-width:640px"><h2 style="color:#034EA2">No se pudo abrir la página</h2><p>' + h_(err && err.message ? err.message : err) + '</p><p style="color:#5D6B82;font-size:13px">Si sos quien administra el sistema: ejecutá <b>verificarInstalacion()</b> en Apps Script para ver qué falta.</p></div>');
  }
}

function doGet_(e) {
  const p = (e && e.parameter) || {};
  const aj = ajustes_();
  const seguro = s => (/^[A-Za-z0-9_-]{1,64}$/.test(String(s || '')) ? String(s) : '');
  const url = urlBase_();
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
    return salida_(t, (p.tablero ? 'Atracción y Elección' : archivo) + ' · ' + t.sitio);
  }
  const t = HtmlService.createTemplateFromFile('Web');
  t.url = url; t.bid = seguro(p.b); t.sitio = aj.nombre_sitio || 'Empleos';
  return salida_(t, t.sitio);
}

function verificarInstalacion() {
  const faltan = [];
  ['Web', 'Tablero', 'Screening', 'Competencias', 'Estilos', 'FormScreening', 'FormCompetencias'].forEach(n => { try { HtmlService.createHtmlOutputFromFile(n); } catch (e) { faltan.push('archivo HTML "' + n + '"'); } });
  ['OPENAI_API_KEY', 'ADMIN_KEY', 'SHEET_ID'].forEach(k => { if (!prop_(k)) faltan.push('propiedad ' + k); });
  ['Candidatos', 'Busquedas', 'Preguntas', 'Screening_Eval', 'Screening_Val', 'Competencias', 'Plantillas', 'Historial', 'Evaluaciones', 'Ajustes'].forEach(n => { try { if (!hoja_(n)) faltan.push('pestaña ' + n); } catch (e) { faltan.push('pestaña ' + n); } });
  Logger.log(faltan.length ? '❌ Faltan: ' + faltan.join(', ') : '✅ Todo instalado');
  const url = urlBase_();
  Logger.log('URL base: ' + (url || '(sin implementar)'));
  if (url && /\/dev$/.test(url)) Logger.log('⚠️ Esa URL termina en /dev (es la de prueba, solo funciona para el dueño). Copiá la que termina en /exec (Implementar → Administrar implementaciones) y pegala en Ajustes → url_web.');
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
    row[COL.dni - 1] = dni; row[COL.email - 1] = email; row[COL.loc - 1] = o.loc || ''; row[COL.domicilio - 1] = o.dom || '';
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
    if (!String(d.nombre || '').trim() || !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email) || !String(d.telefono || '').trim() || dni.length < 6 || !String(d.localidad || '').trim() || !String(d.domicilio || '').trim()) throw new Error('Datos incompletos');
    const decl = p.declaradas || {};
    b.preguntas.forEach(q => { if (q.requerido && !String(decl[q.id] || '').trim()) throw new Error('Falta responder: ' + q.etiqueta); });
    if (!p.cvBase64 || String(p.cvBase64).length > 14000000) throw new Error('CV ausente o muy pesado');
    const ext = (String(p.cvNombre || '').split('.').pop() || 'pdf').toLowerCase().replace(/[^a-z0-9]/g, '').slice(0, 5);
    if (['pdf', 'doc', 'docx'].indexOf(ext) === -1) throw new Error('Formato de CV no válido');
    const nombre = String(d.nombre).trim();
    const resp = b.preguntas.filter(q => String(decl[q.id] || '').trim()).map(q => ({ q: q.etiqueta, a: String(decl[q.id]).trim() }));
    const mime = ext === 'pdf' ? 'application/pdf' : (ext === 'doc' ? 'application/msword' : 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');

    const fila = agregarCandidato_({
      busq: b.id, nombre: nombre, email: email, tel: String(d.telefono).trim(), dni: dni, loc: String(d.localidad).trim(), dom: String(d.domicilio).trim().slice(0, 200),
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
    const bloque = ['Nombre: ' + r[COL.nombre - 1], 'Localidad: ' + (r[COL.loc - 1] || 'no indicada'), 'Domicilio: ' + (r[COL.domicilio - 1] || 'no indicado'), 'Pretensión salarial: ' + (r[COL.pret - 1] || 'no indicó')]
      .concat(decl.length ? decl.map(x => x.q + ': ' + x.a) : ['(No completó el formulario: evaluá solo con el CV y marcá como "REVISAR" lo que no se pueda verificar, por ejemplo la matrícula.)']).join('\n');
    const prompt = `Sos un analista senior de selección de personal. Analizá esta postulación para la búsqueda "${b.titulo}".

${b.prompt_ia}

Tenés DOS fuentes: (A) los datos DECLARADOS por el candidato y (B) el texto del CV. Cruzalas. Si declara algo que el CV no respalda o contradice, reflejalo en la justificación.

=== DECLARADOS (no inventar) ===
${bloque}
=== FIN DECLARADOS ===

Devolvé SOLO JSON válido con este formato:
{"es_cv":true,"nombre":"Nombre y apellido real del candidato","telefono":"","dni":"solo números","localidad":"SOLO la ciudad o localidad (ej. Rosario, Funes); nunca la calle ni el número. Vacío si el CV no la menciona","domicilio":"domicilio de residencia del candidato: el que figura como Domicilio o Dirección en sus datos personales (calle y número). NO uses direcciones de instituciones, empleadores ni lugares de estudio. Vacío si no figura","nivel_ajuste":"A/B+/B/C","justificacion":"una sola oración corta con el porqué de la nota","formacion_titulo":"título o formación principal, textual (ej. Enfermero/a Profesional)","requisitos":{"formacion":"","matricula":"","experiencia":"","zona":"","continuidad":""},"resumen_experiencia":"3-4 oraciones narrativas sobre trayectoria, atributos y actitud, sin repetir el checklist"}
Devolvé es_cv:false SOLO si el texto claramente NO es un currículum (por ejemplo una factura o publicidad). Ante la duda, es_cv:true.
REGLAS DE LA NOTA:
- A: cumple todo, perfil potente sin ruido.
- B+: muy bueno, cumple lo excluyente con una cuestión menor.
- B: cumple lo mínimo con dudas relevantes, o borderline con factores compensatorios (reserva).
- C: no cumple un requisito excluyente o tiene una bandera roja clara.
- zona: si no se indica ninguna localidad (solo calle y número), asumí que vive en Rosario: "Zona OK". Usá "Fuera" solo si la localidad indicada queda lejos de Rosario y su zona.`;
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
    const fill = (col, v) => { v = String(v || '').trim(); if (v && !String(r[col - 1] || '').trim()) sh.getRange(fila, col).setValue(v); };
    fill(COL.domicilio, domicilioPorEtiqueta_(texto) || c.domicilio);
    if (String(r[COL.fuente - 1]) !== 'Web') {
      const nom = String(c.nombre || '').trim();
      if (nom.split(/\s+/).length >= 2) sh.getRange(fila, COL.nombre).setValue(nom);
      fill(COL.tel, c.telefono);
      const dni = String(c.dni || '').replace(/\D/g, '');
      if (dni.length >= 6) fill(COL.dni, dni);
      const locIA = limpiaLoc_(c.localidad);
      if (locIA) fill(COL.loc, locIA);
    }
    let asumida = false;
    if (!limpiaLoc_(sh.getRange(fila, COL.loc).getValue())) { sh.getRange(fila, COL.loc).setValue('Rosario'); asumida = true; }
    const req = c.requisitos || {};
    const zonaTxt = asumida ? 'Zona OK' : (/fuera/i.test(String(req.zona)) ? 'Fuera' : 'Zona OK');
    const declForm = (decl.find(x => /t[ií]tulo|formaci/i.test(x.q)) || {}).a;
    if (zonaTxt) sh.getRange(fila, COL.zona).setValue(zonaTxt);
    const forma = String(declForm || c.formacion_titulo || '').trim();
    if (forma) sh.getRange(fila, COL.formacion).setValue(forma);
    const reqTxt = Object.keys(req).map(k => k.charAt(0).toUpperCase() + k.slice(1) + ': ' + (req[k] || 'N/D')).join(' | ');
    const nota = ['A', 'B+', 'B', 'C'].indexOf(String(c.nivel_ajuste)) > -1 ? String(c.nivel_ajuste) : 'C';
    sh.getRange(fila, COL.ajuste, 1, 4).setValues([[nota, 'IA sugiere ' + nota + ': ' + String(c.justificacion || ''), String(c.resumen_experiencia || ''), reqTxt]]);
    registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Análisis CV', 'La IA sugiere ' + nota, 'sistema');
    sh.getRange(fila, COL.proc).setValue('SI');
    try { sugerirPoliclinico_(fila); } catch (e2) { registrarError_('sugerir policlínico fila ' + fila, e2); }
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
    [actualizarEsquema_, procesarCorreos, procesarCarpeta, procesarPendientes, completarDomicilios_, verificarDomicilios_, sugerirPendientes_, syncCalendario].forEach(fn => { try { fn(); } catch (e) { registrarError_(fn.name, e); } });
  } finally { lock.releaseLock(); }
}


function geo_(q) {
  const res = Maps.newGeocoder().setRegion('ar').setLanguage('es').geocode(q);
  if (res.status === 'OK' && res.results.length) {
    const g = res.results[0].geometry;
    return { lat: g.location.lat, lng: g.location.lng, aprox: g.location_type === 'APPROXIMATE' };
  }
  if (res.status === 'ZERO_RESULTS') return null;
  throw new Error('Geocoder: ' + res.status);
}

function kmEntre_(a, b) {
  const rad = x => x * Math.PI / 180, dLat = rad(b.lat - a.lat), dLng = rad(b.lng - a.lng);
  const h = Math.sin(dLat / 2) ** 2 + Math.cos(rad(a.lat)) * Math.cos(rad(b.lat)) * Math.sin(dLng / 2) ** 2;
  return 6371 * 2 * Math.asin(Math.sqrt(h));
}

function puntosPoli_() {
  const aj = ajustes_();
  const dirs = [aj.direccion_policlinico_1, aj.direccion_policlinico_2];
  const props = PropertiesService.getScriptProperties();
  let c = null;
  try { c = JSON.parse(props.getProperty('GEO_POLI') || 'null'); } catch (e) { }
  if (c && c.dirs && c.dirs[0] === dirs[0] && c.dirs[1] === dirs[1]) return c.pts;
  const pts = dirs.map(d => { const g = geo_(d); if (!g) throw new Error('No se encontró la dirección del policlínico: ' + d); return g; });
  props.setProperty('GEO_POLI', JSON.stringify({ dirs: dirs, pts: pts }));
  return pts;
}

function domicilioPorEtiqueta_(texto) {
  const m = String(texto || '').replace(/\r/g, '\n').match(/(?:domicilio|direcci[oó]n|dir\.|residencia|vive en)\s*(?:real|actual|particular)?\s*[:\-–]\s*([^\n|]{4,90})/i);
  if (!m) return '';
  const v = m[1].replace(/\s+(fecha|tel|cel|tel[eé]fono|e-?mail|correo|d\.?n\.?i|mat|estado civil|nacionalidad)\b.*$/i, '').replace(/[\s.,;]+$/, '').trim();
  return v.length >= 4 ? v.slice(0, 120) : '';
}

function verificarDomicilios_() {
  const sh = hoja_('Candidatos');
  const d = sh.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < d.length && n < 4; i++) {
    const r = d[i];
    const dom = String(r[COL.domicilio - 1] || '').trim(), rev = String(r[COL.domRev - 1] || '');
    if (!r[COL.nombre - 1] || !r[COL.cv - 1] || !dom || rev === 'Verificado' || rev === 'Manual') continue;
    const fila = i + 1;
    if (String(r[COL.fuente - 1]) === 'Web') { sh.getRange(fila, COL.domRev).setValue('Verificado'); continue; }
    n++;
    const texto = textoCV_(r[COL.cv - 1]);
    if (!texto || texto.trim().length < 50) {
      const k = (Number(rev) || 0) + 1;
      sh.getRange(fila, COL.domRev).setValue(k >= 3 ? 'Verificado' : String(k));
      continue;
    }
    const et = domicilioPorEtiqueta_(texto);
    const lim = x => norm_(x).replace(/[^a-z0-9]/g, '');
    if (et && lim(et) !== lim(dom)) {
      sh.getRange(fila, COL.domicilio).setValue(et);
      sh.getRange(fila, COL.polSug).clearContent();
      registrarHistorial_(r[COL.id - 1], r[COL.nombre - 1], 'Domicilio', 'Corregido leyendo el CV: ' + dom + ' → ' + et, 'sistema');
    }
    sh.getRange(fila, COL.domRev).setValue('Verificado');
  }
}

function limpiaLoc_(t) {
  t = String(t || '').trim();
  return (!t || /\d/.test(t) || t.length > 40) ? '' : t;
}

function repararLocalidades_() {
  const sh = hoja_('Candidatos');
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    const r = d[i];
    if (!r[COL.nombre - 1]) continue;
    const loc = String(r[COL.loc - 1] || '').trim();
    const z0 = String(r[COL.zona - 1] || '').trim();
    if (z0 && z0 !== 'Zona OK' && z0 !== 'Fuera') sh.getRange(i + 1, COL.zona).setValue('Zona OK');
    if (limpiaLoc_(loc)) continue;
    const fila = i + 1;
    if (loc && !String(r[COL.domicilio - 1] || '').trim()) sh.getRange(fila, COL.domicilio).setValue(loc);
    sh.getRange(fila, COL.loc).setValue('Rosario');
    sh.getRange(fila, COL.polSug).clearContent();
    if (String(r[COL.zona - 1] || '').trim() === 'Fuera') sh.getRange(fila, COL.zona).setValue('Zona OK');
  }
}

function sugerirPoliclinico_(fila) {
  const sh = hoja_('Candidatos');
  const dom = String(sh.getRange(fila, COL.domicilio).getValue() || '').trim();
  if (!dom) return;
  const loc = limpiaLoc_(sh.getRange(fila, COL.loc).getValue()) || 'Rosario';
  const q = dom + (loc && dom.toLowerCase().indexOf(loc.toLowerCase()) === -1 ? ', ' + loc : '') + ', Santa Fe, Argentina';
  const g = geo_(q);
  const celda = sh.getRange(fila, COL.polSug);
  if (!g) { celda.setValue('No se pudo ubicar la dirección'); return; }
  const pts = puntosPoli_();
  const d = pts.map(p => kmEntre_(g, p));
  const i = d[0] <= d[1] ? 0 : 1;
  if (!g.aprox) sh.getRange(fila, COL.zona).setValue(d[i] <= (Number(ajustes_().zona_max_km) || 40) ? 'Zona OK' : 'Fuera');
  const f = x => (g.aprox ? '≈' : '') + (Math.round(x * 10) / 10).toString().replace('.', ',') + ' km';
  celda.setValue(POLICLINICOS[i] + ' · ' + f(d[i]) + ' (' + POLICLINICOS[1 - i].replace('Policlínico ', '') + ': ' + f(d[1 - i]) + ')');
  registrarHistorial_(sh.getRange(fila, COL.id).getValue(), sh.getRange(fila, COL.nombre).getValue(), 'Policlínico sugerido', POLICLINICOS[i] + ' (' + f(d[i]) + ')', 'sistema');
}

function sugerirPendientes_() {
  const sh = hoja_('Candidatos');
  const d = sh.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < d.length && n < 8; i++) {
    if (!d[i][COL.nombre - 1] || !String(d[i][COL.domicilio - 1] || '').trim() || String(d[i][COL.polSug - 1] || '').trim()) continue;
    try { sugerirPoliclinico_(i + 1); } catch (e) { registrarError_('sugerir policlínico fila ' + (i + 1), e); break; }
    n++;
  }
}


function domicilioDeCV_(texto) {
  const resp = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
    method: 'post', contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + prop_('OPENAI_API_KEY') },
    payload: JSON.stringify({ model: 'gpt-4o-mini', response_format: { type: 'json_object' }, messages: [
      { role: 'system', content: 'Del texto de un currículum extraé únicamente el domicilio de residencia del candidato (calle y número), el que figura como Domicilio o Dirección en sus datos personales. NO uses direcciones de instituciones, empleadores ni lugares de estudio. Devolvé SOLO JSON {"domicilio":"..."} con cadena vacía si el CV no lo indica. No inventes nada.' },
      { role: 'user', content: String(texto).slice(0, 6000) }] })
  });
  const parsed = JSON.parse(resp.getContentText());
  if (parsed.error) throw new Error('OpenAI: ' + parsed.error.message);
  return String(JSON.parse(parsed.choices[0].message.content).domicilio || '').trim().slice(0, 200);
}

function completarDomicilios_() {
  const sh = hoja_('Candidatos');
  const d = sh.getDataRange().getValues();
  let n = 0;
  for (let i = 1; i < d.length && n < 4; i++) {
    const r = d[i];
    if (!r[COL.nombre - 1] || !r[COL.cv - 1] || String(r[COL.domicilio - 1] || '').trim()) continue;
    const rev = String(r[COL.domRev - 1] || '');
    if (rev === 'OK' || rev === 'Sin dato') continue;
    const fila = i + 1;
    n++;
    const texto = textoCV_(r[COL.cv - 1]);
    if (!texto || texto.trim().length < 50) {
      const k = (Number(rev) || 0) + 1;
      sh.getRange(fila, COL.domRev).setValue(k >= 3 ? 'Sin dato' : String(k));
      continue;
    }
    const dom = domicilioPorEtiqueta_(texto) || domicilioDeCV_(texto);
    if (!dom) { sh.getRange(fila, COL.domRev).setValue('Sin dato'); continue; }
    sh.getRange(fila, COL.domicilio).setValue(dom);
    sh.getRange(fila, COL.domRev).setValue('OK');
    try { sugerirPoliclinico_(fila); } catch (e) { registrarError_('sugerir policlínico fila ' + fila, e); }
  }
}
