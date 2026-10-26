/* ============================================================
   PERFIL: Director/a Comercial de Proyectos de Consultoría (director-comercial)
   Empresa: Energía y Minería — CABA
   Correr UNA vez: crearPerfilDirectorComercial()
   Crea carpetas, pestaña de candidatos, Config, Perfiles_Hero,
   criterios de screening, competencias y preguntas de postulación.
   Si el perfil ya existe en Config, no hace nada.
   ============================================================ */

function crearPerfilDirectorComercial() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const carpetaMadre = DriveApp.getFolderById('18bhn7cD2PkcK_SldSfKHrjiieHvO3tYH');
  const p = {
    id: 'director-comercial',
    nombre: 'Director Comercial de Proyectos de Consultoría',
    hoja: 'Director Comercial',
    empresa: 'Energía y Minería',
    ubicacion: 'CABA'
  };

  const config = ss.getSheetByName('Config');
  const cfgRows = config.getDataRange().getValues();
  if (cfgRows.some(r => r[0] === p.id)) { Logger.log('Ya existe ' + p.id + ' en Config: no se hace nada.'); return; }

  const promptExtra =
`PERFIL: Director/a Comercial de Proyectos de Consultoría – Energía & Minería. Cliente: consultora (sector Energía y Minería), CABA. Reporta al CEO. Tiempo completo, esquema híbrido 3x2, viajes comerciales locales eventuales.
MISIÓN: liderar la estrategia de desarrollo de negocios y captación de clientes de las líneas de consultoría, enfocado en Energía (Oil & Gas, Energías Renovables, Generación/Distribución Eléctrica) y Minería. Abrir nuevas cuentas, gestionar relaciones a largo plazo a nivel C-Level y estructurar propuestas comerciales complejas a medida. Cumple objetivos anuales de ventas y margen de rentabilidad de la práctica sectorial.
RESPONSABILIDADES CLAVE: prospección y cierre de oportunidades de consultoría (estratégica, operativa, técnica o regulatoria); gestión de cuentas clave y venta cruzada; liderazgo de licitaciones RFP/RFQ (estructura de costos y propuestas técnico-comerciales); análisis de tendencias del sector (transición energética, automatización de minas, regulaciones ambientales).

EXPERIENCIA (EXCLUYENTE): mínimo 8 años en desarrollo de negocios, ventas B2B de ciclo largo o consultoría senior, atendiendo de manera directa a empresas del sector minero y/o energético.
NETWORK (EXCLUYENTE): red de contactos activa y comprobable con tomadores de decisiones (CEOs, COOs, VPs de Operaciones, Directores de Abastecimiento).
FORMACIÓN: profesional en Ingeniería (Industrial, Minas, Petróleo, Energía), Administración de Empresas, Economía o carreras afines. MBA o especialización en gestión de recursos naturales suma mucho (no excluyente).
CONOCIMIENTOS CLAVE: procesos operativos, sustentabilidad (ESG), normativa regulatoria y ciclos de inversión de energía y minería. Armado de propuestas y licitaciones complejas.
IDIOMAS: inglés avanzado/bilingüe (EXCLUYENTE). Portugués deseable (no excluyente).
ZONA: CABA, híbrido 3x2, con posibilidad de viajes comerciales locales.

BANDERAS ROJAS: sin experiencia comprobable en el sector minero/energético; trayectoria sin cierre de negocios de consultoría ni ventas B2B complejas; sin red a nivel C-Level; alta rotación sin justificación; logros comerciales sin cifras ni cuentas concretas.

CONTINUIDAD LABORAL (criterio importante para Blumia): analizá el historial. VALORÁ permanencia y crecimiento dentro de cada empresa. PENALIZÁ rotación alta con saltos cortos y frecuentes (job hopping) SIN justificación razonable (cierres, fin de obra/proyecto, mudanzas). La estabilidad y la progresión interna suman fuerte.

EVALUACIÓN DE BORDERLINE — "incluir y señalar": NO descartes mecánicamente a quien no cumple el 100%. Si a un candidato le falta algún requisito pero presenta FACTORES COMPENSATORIOS (trayectoria sólida, continuidad laboral ejemplar, formación destacada, experiencia transferible, logros relevantes), clasificalo como MEDIO y en el resumen marcá explícitamente "REVISAR:" explicando QUÉ le falta y QUÉ lo compensa, para que el selector decida. Ante la duda, INCLUIR y señalar, no filtrar. Reservá BAJO para quienes claramente no tienen experiencia comercial de consultoría o ventas B2B complejas, no llegan a la base mínima, o caen en banderas rojas claras.

NIVEL_AJUSTE:
- Alto: 8+ años en desarrollo de negocios / ventas B2B de ciclo largo / consultoría senior con clientes de energía y/o minería + red C-Level + inglés avanzado, con buena continuidad laboral.
- Medio: cumple lo básico pero le falta algo (años, sector, red o inglés) O es un borderline con factores compensatorios → marcar "REVISAR:" en el resumen.
- Bajo: sin experiencia comercial real en consultoría o ventas B2B complejas, sin vínculo con el sector, rotación injustificada marcada, o banderas rojas.`;

  // 1) Carpetas Drive
  const itC = carpetaMadre.getFoldersByName(p.empresa);
  const carpetaEmpresa = itC.hasNext() ? itC.next() : carpetaMadre.createFolder(p.empresa);
  const carpetaPerfil = carpetaEmpresa.createFolder(p.nombre);
  const folderMadre = carpetaPerfil.createFolder('CVs nuevos');
  const folderRecibidos = carpetaPerfil.createFolder('CVs procesados');
  const folderScreening = carpetaPerfil.createFolder('Screening PDFs');
  const folderCompetencias = carpetaPerfil.createFolder('Competencias PDFs');
  Logger.log('Carpetas Drive creadas: ' + carpetaPerfil.getUrl());

  // 2) Pestaña de candidatos
  let hoja = ss.getSheetByName(p.hoja);
  if (!hoja) {
    hoja = ss.insertSheet(p.hoja);
    const headers = ['Fecha','Nombre','Teléfono','DNI','Nivel de Ajuste','Requisitos','Resumen Experiencia','Link CV','Estado','Fuente','Screening','Impresión','Detalle','Entrevistó','Entrevista competencias','Impresión','Detalle','Entrevistó','Entrevista Compañia','Impresión','Contactar','Correo'];
    hoja.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1A6645').setFontColor('#FFFFFF');
    hoja.setFrozenRows(1);
    hoja.autoResizeColumns(1, headers.length);
  }

  // 3) Config (por nombre de columna; "empresa" agrupa el perfil en el kanban)
  const H = cfgRows[0];
  const valores = {
    perfil_id: p.id, nombre_perfil: p.nombre, nombre: p.nombre, sheet_id: SHEET_MASTER_ID, hoja_nombre: p.hoja,
    folder_madre: folderMadre.getId(), folder_recibidos: folderRecibidos.getId(),
    folder_pdf_screening: folderScreening.getId(), folder_pdf_competencias: folderCompetencias.getId(),
    prompt_extra: promptExtra, empresa: p.empresa
  };
  if (H.indexOf('empresa') === -1) Logger.log('⚠️ Config no tiene columna "empresa": el perfil no se va a agrupar bien en el kanban.');
  else if (H.indexOf('empresa') !== 9) Logger.log('⚠️ La columna "empresa" de Config no es la J: revisar empresasYPerfiles().');
  config.appendRow(H.map(h => valores[h] !== undefined ? valores[h] : ''));

  // 4) Perfiles_Hero
  ss.getSheetByName('Perfiles_Hero').appendRow([p.id, p.empresa, p.nombre, p.ubicacion, 'Screening · 1ª entrevista', 'Ariel, Analía, Flavia, Marcos', '8.5', '6.5']);

  // 5) Screening: criterios con peso
  const ev = ss.getSheetByName('Perfiles_Screening_Eval');
  [['Experiencia en desarrollo de negocios B2B de ciclo largo / consultoría senior', 3],
   ['Experiencia directa con empresas de energía y/o minería', 2],
   ['Red de contactos a nivel C-Level', 2],
   ['Liderazgo de licitaciones y propuestas complejas (RFP/RFQ)', 1],
   ['Formación e inglés', 1]]
  .forEach((e, i) => ev.appendRow([p.id, 'e' + (i + 1), e[0], e[1]]));

  // 6) Screening: validaciones
  const va = ss.getSheetByName('Perfiles_Screening_Val');
  [['Remuneración pretendida compatible', 'Sí|No|A revisar'],
   ['Disponibilidad para viajes comerciales locales', 'Sí|No'],
   ['Modalidad híbrida 3x2 en CABA', 'Sí|No'],
   ['Preaviso razonable', 'Sí|No'],
   ['Interés real en la posición', 'Alto|Medio|Bajo'],
   ['Inglés avanzado / bilingüe', 'Avanzado|Intermedio|Básico']]
  .forEach((v, i) => va.appendRow([p.id, 'v' + (i + 1), v[0], v[1]]));

  // 7) Competencias (con preguntas STAR)
  const co = ss.getSheetByName('Perfiles_Competencias');
  [['Desarrollo de negocios', 'Prospecta y abre cuentas nuevas en ciclos de venta largos y complejos.', 2,
    'Contame cómo abriste una cuenta nueva grande: qué hiciste y qué resultado tuvo|Describí una oportunidad que parecía perdida y cómo la encaraste|¿Cómo armás tu pipeline y qué hacés cuando se vacía?'],
   ['Relación C-Level y cuentas clave', 'Construye y sostiene vínculos de largo plazo con decisores y hace venta cruzada.', 2,
    'Contame una relación con un CEO/COO que construiste a lo largo de los años|Dame un ejemplo de venta cruzada en una cuenta existente|¿Cómo recuperaste una cuenta en riesgo?'],
   ['Negociación y propuestas / licitaciones', 'Estructura costos y propuestas técnico-comerciales y negocia condiciones.', 2,
    'Contame una licitación grande (RFP/RFQ) que lideraste y cómo estructuraste la propuesta|Describí una negociación difícil de precio o alcance y cómo cerró|¿Cómo coordinás a los equipos técnicos para una propuesta?'],
   ['Visión estratégica del sector', 'Lee tendencias de energía y minería y adapta la oferta de servicios.', 1,
    'Contame cómo detectaste una tendencia del sector y la convertiste en oferta|¿Qué cambio regulatorio o de mercado hoy impacta más a tus clientes?|Dame un ejemplo de un servicio nuevo que ayudaste a lanzar'],
   ['Orientación a resultados', 'Se compromete con objetivos de ventas y margen y los cumple.', 1,
    'Contame un año en el que superaste tu objetivo: cómo lo lograste|Describí un objetivo que no alcanzaste y qué aprendiste|¿Cómo cuidás el margen sin perder la venta?']]
  .forEach((c, i) => co.appendRow([p.id, 'c' + (i + 1), c[0], c[1], c[2], c[3]]));

  // 8) Preguntas de postulación
  const pp = ss.getSheetByName('Postulacion_Preguntas');
  const preguntas = [
    [p.id, 'p1', '¿Cuántos años de experiencia tenés en desarrollo de negocios, ventas B2B de ciclo largo o consultoría senior?', 'select', 'Menos de 5|5 a 7|8 a 12|Más de 12', 'SI'],
    [p.id, 'p2', '¿Trabajaste directamente con empresas del sector minero y/o energético? Indicá cuáles', 'texto', '', 'SI'],
    [p.id, 'p3', '¿Tenés red de contactos activa con tomadores de decisiones (CEO, COO, VP de Operaciones, Director de Abastecimiento)?', 'select', 'Sí, amplia|Sí, parcial|No', 'SI'],
    [p.id, 'p4', '¿Participaste en licitaciones (RFP/RFQ)?', 'select', 'Sí, liderándolas|Sí, participé|No', 'SI'],
    [p.id, 'p5', 'Nivel de inglés', 'select', 'Básico|Intermedio|Avanzado|Bilingüe', 'SI']
  ];
  pp.getRange(pp.getLastRow() + 1, 1, preguntas.length, 6).setValues(preguntas);

  formatearHojaDirectorComercial();

  Logger.log('✅ Perfil ' + p.id + ' creado (empresa: ' + p.empresa + ').');
  Logger.log('Kanban: ' + URL_DEPLOY + '?tipo=kanban&perfil=' + p.id);
  Logger.log('Postulación: ' + URL_DEPLOY + '?tipo=postulacion&perfil=' + p.id);
}

/* Formato, desplegables y links de la pestaña "Director Comercial".
   Se llama sola desde crearPerfilDirectorComercial; se puede correr aparte. */
function formatearHojaDirectorComercial() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  const p = { id: 'director-comercial', hoja: 'Director Comercial' };
  const hoja = ss.getSheetByName(p.hoja);
  if (!hoja) { Logger.log('No existe la pestaña ' + p.hoja); return; }
  const modelo = ss.getSheetByName('Supervisor de Operaciones');
  if (modelo) for (let c = 1; c <= Math.min(modelo.getLastColumn(), hoja.getMaxColumns()); c++) hoja.setColumnWidth(c, modelo.getColumnWidth(c));
  const lista = l => SpreadsheetApp.newDataValidation().requireValueInList(l, true).setAllowInvalid(true).build();
  const entrevistadores = ['Ariel', 'Marcos', 'Ever', 'Analía', 'Flavia'];
  hoja.getRange(2, 9, 250, 1).setDataValidation(lista([
    'Pendiente','Nuevo','Contactado','Avanza a screening','Screening agendado','Avanza a competencias','Competencias agendada',
    'Avanza a screening empresa','Screening empresa agendado','Avanza a presencial empresa','Entrevista empresa agendada',
    'En cliente','Contratado','En reserva','Descartado sin screening','Descartado post screening','Descartado en competencias',
    'Descartado en screening empresa','Descartado presencial empresa','Descartado Fuera de zona','Descartado NO le interesa']));
  hoja.getRange(2, 12, 250, 1).setDataValidation(lista(['A', 'B', 'C']));
  hoja.getRange(2, 16, 250, 1).setDataValidation(lista(['A', 'B', 'C']));
  hoja.getRange(2, 14, 250, 1).setDataValidation(lista(entrevistadores));
  hoja.getRange(2, 18, 250, 1).setDataValidation(lista(entrevistadores));
  const fScr = [], fComp = [];
  for (let i = 2; i <= 250; i++) {
    fScr.push(['=IF(B' + i + '<>"",HYPERLINK("' + URL_DEPLOY + '?perfil=' + p.id + '&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(N' + i + '),"Iniciar Screening"),"")']);
    fComp.push(['=IF(B' + i + '<>"",HYPERLINK("' + URL_DEPLOY + '?tipo=competencias&perfil=' + p.id + '&nombre="&ENCODEURL(B' + i + ')&"&entrevistador="&ENCODEURL(R' + i + '),"Evaluar competencias"),"")']);
  }
  hoja.getRange(2, 11, fScr.length, 1).setFormulas(fScr);
  hoja.getRange(2, 15, fComp.length, 1).setFormulas(fComp);
  Logger.log('✅ Formato y links de ' + p.hoja + ' listos.');
}

/* Vincula el relevamiento HTML editable (RelevamientoDirectorComercial.html) al botón "Perfil del puesto" del kanban.
   El botón pide el PIN (propiedad del script BLUMIA_PIN) y abre el HTML con un acceso temporal.
   Si había un PDF o Doc vinculado antes, lo manda a la papelera.
   Correr: vincularPerfilHtmlDirectorComercial() */
function vincularPerfilHtmlDirectorComercial() {
  const config = SpreadsheetApp.openById(SHEET_MASTER_ID).getSheetByName('Config');
  const rows = config.getDataRange().getValues();
  const H = rows[0];
  const f = rows.findIndex(r => r[0] === 'director-comercial');
  if (f === -1) { Logger.log('No existe director-comercial en Config'); return; }
  const col = n => { let c = H.indexOf(n); if (c === -1) { c = H.length; H.push(n); config.getRange(1, c + 1).setValue(n); } return c + 1; };
  const cLink = col('link_perfil');
  const viejo = String(rows[f][cLink - 1] || '').match(/[-\w]{25,}/);
  if (viejo) { try { DriveApp.getFileById(viejo[0]).setTrashed(true); } catch (e) { } }
  config.getRange(f + 1, cLink).setValue('');
  config.getRange(f + 1, col('perfil_html')).setValue('RelevamientoDirectorComercial');
  config.getRange(f + 1, col('perfil_protegido')).setValue('SI');
  Logger.log('✅ Perfil del puesto (HTML editable) vinculado al kanban, protegido con PIN.');
  if (!PropertiesService.getScriptProperties().getProperty('BLUMIA_PIN')) Logger.log('⚠️ Falta cargar la propiedad del script BLUMIA_PIN (el PIN).');
}
