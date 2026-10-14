// Documentación para el ingreso: formulario público (?docs=1 o ?docs=2, opcional &c=<ID candidato>).
// Cada archivo se guarda en una carpeta por persona dentro de la carpeta del policlínico:
//   PAMI 1 → Ajustes "carpeta_docs" (Drive de dotacionpp1, se crea sola)
//   PAMI 2 → Ajustes "carpeta_docs_poli2" (carpeta del Drive de dotacionpp2 compartida con dotacionpp1 como Editor)
// Nombres normalizados: Apellido_Nombre_DNI, Apellido_Nombre_CV, etc.

const DOCS_TIPOS = [
  { id: 'dni', etiqueta: 'DNI (frente y dorso)', sufijo: 'DNI' },
  { id: 'cv', etiqueta: 'Currículum Vitae', sufijo: 'CV' },
  { id: 'cuil', etiqueta: 'Constancia de CUIL', sufijo: 'CUIL' },
  { id: 'titulo', etiqueta: 'Título de Enfermería', sufijo: 'Titulo_Enfermeria' },
  { id: 'matricula', etiqueta: 'Constancia de Matrícula de Enfermería (vigente y actualizada)', sufijo: 'Matricula' }
];
const DOCS_EXT = { pdf: 'application/pdf', jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', heic: 'image/heic', webp: 'image/webp' };
const HEAD_DOCS = ['Fecha', 'Policlínico', 'Apellido', 'Nombre', 'CUIL', 'ID candidato', 'Candidato', 'Carpeta', 'Archivos'];

function docsPol_(p) { return String(p) === '2' ? '2' : String(p) === '1' ? '1' : ''; }

function docsPalabra_(t) {
  return String(t || '').normalize('NFC').replace(/[^A-Za-zÁÉÍÓÚÜÑáéíóúüñ' -]/g, '').replace(/\s+/g, ' ').trim()
    .split(' ').map(w => w ? w.charAt(0).toUpperCase() + w.slice(1).toLowerCase() : '').join('_').slice(0, 60);
}

function docsCuilValido_(c) {
  const d = String(c || '').replace(/\D/g, '');
  if (d.length !== 11 || ['20', '23', '24', '25', '26', '27', '30', '33', '34'].indexOf(d.slice(0, 2)) === -1) return false;
  const pesos = [5, 4, 3, 2, 7, 6, 5, 4, 3, 2];
  const r = pesos.reduce((s, p, i) => s + p * Number(d[i]), 0) % 11;
  if (r === 1) return true; // casos especiales (prefijo 23): no rechazar
  return Number(d[10]) === (r === 0 ? 0 : 11 - r);
}

function docsDatos_(p) {
  const pol = docsPol_(p && p.pol);
  const ap = docsPalabra_(p && p.apellido), no = docsPalabra_(p && p.nombre);
  const cuil = String(p && p.cuil || '').replace(/\D/g, '');
  if (!pol) throw new Error('Policlínico inválido');
  if (!ap || !no) throw new Error('Completá apellido y nombre');
  if (!docsCuilValido_(cuil)) throw new Error('CUIL / CUIT inválido');
  return { pol: pol, ap: ap, no: no, cuil: cuil, cuilTxt: cuil.slice(0, 2) + '-' + cuil.slice(2, 10) + '-' + cuil.slice(10) };
}

function docsCarpetaRaiz_(pol) {
  const aj = ajustes_();
  const guardar = (clave, id) => {
    asegurarAjustes_([[clave, '']]);
    const sh = hoja_('Ajustes'), d = sh.getDataRange().getValues();
    const i = d.findIndex(r => String(r[0]).trim() === clave);
    if (i > 0) sh.getRange(i + 1, 2).setValue(id);
  };
  let raiz = null;
  if (aj.carpeta_docs) try { raiz = DriveApp.getFolderById(aj.carpeta_docs); } catch (e) { registrarError_('carpeta_docs', e); }
  if (!raiz) { raiz = DriveApp.createFolder('Empleos PAMI - Documentación ingreso'); guardar('carpeta_docs', raiz.getId()); }
  if (pol === '1') return raiz;
  if (aj.carpeta_docs_poli2) {
    try { return DriveApp.getFolderById(aj.carpeta_docs_poli2); } catch (e) { registrarError_('carpeta_docs_poli2 (¿está compartida con dotacionpp1 como Editor?)', e); }
  }
  // Sin carpeta de dotacionpp2 configurada: se guarda aparte en el Drive de PAMI 1 para no perder nada.
  const it = raiz.getFoldersByName('PAMI 2 (falta carpeta de dotacionpp2)');
  return it.hasNext() ? it.next() : raiz.createFolder('PAMI 2 (falta carpeta de dotacionpp2)');
}

function docsCarpetaPersona_(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const raiz = docsCarpetaRaiz_(d.pol);
    const nombre = d.ap + '_' + d.no + '_' + d.cuil;
    const it = raiz.getFoldersByName(nombre);
    return it.hasNext() ? it.next() : raiz.createFolder(nombre);
  } finally { lock.releaseLock(); }
}

// Sube un archivo. p: {pol, apellido, nombre, cuil, tipo, n, total, archivo, base64}
function docsSubirArchivo(p) {
  try {
    const d = docsDatos_(p);
    const t = DOCS_TIPOS.find(x => x.id === String(p.tipo));
    if (!t) throw new Error('Tipo de documento inválido');
    const ext = (String(p.archivo || '').split('.').pop() || '').toLowerCase().replace(/[^a-z0-9]/g, '');
    if (!DOCS_EXT[ext]) throw new Error('Formato no permitido: subí PDF o JPG');
    if (!p.base64 || String(p.base64).length > 14000000) throw new Error('Archivo vacío o de más de 10 MB');
    const carpeta = docsCarpetaPersona_(d);
    const total = Number(p.total) || 1, n = Number(p.n) || 1;
    let base = d.ap + '_' + d.no + '_' + t.sufijo + (total > 1 ? '_' + n : '');
    let nombre = base + '.' + ext, k = 2;
    while (carpeta.getFilesByName(nombre).hasNext()) nombre = base + '_v' + (k++) + '.' + ext;
    const f = carpeta.createFile(Utilities.newBlob(Utilities.base64Decode(p.base64), DOCS_EXT[ext], nombre));
    return { ok: true, nombre: nombre, id: f.getId() };
  } catch (err) {
    registrarError_('docsSubirArchivo', err);
    return { ok: false, error: String(err && err.message || err) };
  }
}

// Cierra el envío: registra en la pestaña Documentacion, en la ficha del candidato y avisa por mail.
function docsFinalizar(p) {
  try {
    const d = docsDatos_(p);
    const carpeta = docsCarpetaPersona_(d);
    const archivos = (Array.isArray(p.archivos) ? p.archivos : []).map(String).slice(0, 40);
    const sh = hoja_('Candidatos');
    let fila = 0;
    const cid = /^[A-Za-z0-9_-]{1,64}$/.test(String(p.cid || '')) ? String(p.cid) : '';
    if (cid) fila = filaPorId_(sh, cid);
    if (!fila) {
      const dni = d.cuil.slice(2, 10).replace(/^0+/, '');
      const n = sh.getLastRow() - 1;
      if (n > 0) {
        const v = sh.getRange(2, COL.dni, n, 1).getValues();
        const i = v.findIndex(r => String(r[0]).replace(/\D/g, '').replace(/^0+/, '') === dni);
        if (i > -1) fila = i + 2;
      }
    }
    let idCand = '', nomCand = '';
    if (fila) {
      const r = sh.getRange(fila, 1, 1, HEAD_CAND.length).getValues()[0];
      idCand = String(r[COL.id - 1]); nomCand = String(r[COL.nombre - 1]);
      sh.getRange(fila, COL.docs).setValue(carpeta.getUrl());
      sh.getRange(fila, COL.mov).setValue(new Date());
      registrarHistorial_(idCand, nomCand, 'Documentación', 'Envió documentación para el ingreso (' + archivos.length + ' archivos)', 'candidato');
    }
    const ss = ss_();
    let h = ss.getSheetByName('Documentacion');
    if (!h) h = ss.insertSheet('Documentacion');
    asegurarEncabezados_(h, HEAD_DOCS);
    h.appendRow([new Date(), POLICLINICOS[Number(d.pol) - 1], d.ap.replace(/_/g, ' '), d.no.replace(/_/g, ' '), d.cuilTxt, idCand, nomCand || '(sin asociar)', carpeta.getUrl(), archivos.join(', ')]);
    try {
      const to = ajustes_().correo_aviso;
      if (to) MailApp.sendEmail(to, 'Documentación de ingreso · ' + d.ap.replace(/_/g, ' ') + ', ' + d.no.replace(/_/g, ' '),
        'Llegó la documentación de ' + d.ap.replace(/_/g, ' ') + ', ' + d.no.replace(/_/g, ' ') + ' (CUIL ' + d.cuilTxt + ') para el ' + POLICLINICOS[Number(d.pol) - 1] + '.\n' +
        (nomCand ? 'Candidato: ' + nomCand + '\n' : 'No se encontró el candidato en la planilla (revisar pestaña Documentacion).\n') + 'Carpeta: ' + carpeta.getUrl() + '\nArchivos: ' + archivos.join(', '));
    } catch (e) { }
    return { ok: true };
  } catch (err) {
    registrarError_('docsFinalizar', err);
    return { ok: false, error: String(err && err.message || err) };
  }
}

function docsTipos() { return DOCS_TIPOS.map(t => ({ id: t.id, etiqueta: t.etiqueta })); }

function linkDocs_(pol, cid) {
  return urlBase_() + '?docs=' + (docsPol_(pol) || '1') + (cid ? '&c=' + encodeURIComponent(cid) : '');
}
