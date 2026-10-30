/* Pestañas del sheet (pedido de Ariel 30/10): dejar visible en la planilla solo la pestaña de una búsqueda y ocultar todas las demás
   (para quien prefiera trabajar directo en el sheet). Ocultar es solo comodidad, no seguridad: cualquiera con acceso al archivo puede
   mostrarlas de nuevo (Ver → Hojas ocultas). El sistema sigue funcionando igual: encuentra las pestañas por nombre, estén ocultas o no.
   Solo administradores (registroApi_). */

function pestanasSheetSolo_(perfilId) {
  const cfg = getConfig(perfilId);
  const ss = SpreadsheetApp.openById(cfg.sheet_id);
  const dejar = ss.getSheetByName(cfg.hoja_nombre);
  if (!dejar) throw new Error('No encontré la pestaña de esa búsqueda en el sheet.');
  if (dejar.isSheetHidden()) dejar.showSheet();
  ss.setActiveSheet(dejar); // la pestaña activa no se puede ocultar
  let n = 0;
  ss.getSheets().forEach(s => { if (s.getSheetId() !== dejar.getSheetId() && !s.isSheetHidden()) { s.hideSheet(); n++; } });
  return { ok: true, ocultas: n, visible: dejar.getName() };
}

function pestanasSheetTodas_() {
  const ss = SpreadsheetApp.openById(SHEET_MASTER_ID);
  let n = 0;
  ss.getSheets().forEach(s => { if (s.isSheetHidden()) { s.showSheet(); n++; } });
  return { ok: true, mostradas: n };
}
