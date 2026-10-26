function GUARDADO_FINAL_SUPERVISOR(data) {
  // ANCLAJE A PLANILLA SUPERVISOR
  const ID_PLANILLA = "1kDYhElzoRhGwgyBkfEA4EdE4R3zAfWuXe3Thb5FcX00";
  const ss = SpreadsheetApp.openById(ID_PLANILLA);
  const sheet = ss.getSheetByName("Control Busqueda");
  const FOLDER_ID = "144hG9GoqWjcjop4pZyd0DzAPA6_cX1nu"; 

  let nombreBuscado = data.nombre ? data.nombre.trim() : "";
  
  // Buscador que ignora mayúsculas y espacios
  const valores = sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getValues();
const normalizar = s => String(s || '').toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/\s+/g, ' ').trim();
const buscado = normalizar(nombreBuscado);
let fila = -1;
for (let i = 0; i < valores.length; i++) {
  if (normalizar(valores[i][0]) === buscado) { fila = i + 2; break; }
}
if (fila === -1) {
  throw new Error("No encontré a '" + nombreBuscado + "' en col B.");
}

  // Columna L (12) - Resultado
  const celdaL = sheet.getRange(fila, 12);
  celdaL.setDataValidation(null);
  celdaL.setValue(data.resultado);
  
  let color = "#ffffff"; 
  if (data.resultado === "A") color = "#E6F2EC"; 
  if (data.resultado === "B") color = "#FEF3D8"; 
  if (data.resultado === "C") color = "#FDEAE4"; 
  celdaL.setBackground(color).setFontWeight("bold").setHorizontalAlignment("center");

  // Columna N (14) - Entrevistador
  sheet.getRange(fila, 14).setValue(data.entrevistador); 
  
  try {
    const folder = DriveApp.getFolderById(FOLDER_ID);
    const nombreArchivo = "Informe_Screening_" + nombreBuscado.replace(/\s+/g, '_');
    const blob = Utilities.newBlob(data.htmlFull, "text/html", nombreArchivo + ".html");
    const pdfFile = folder.createFile(blob.getAs("application/pdf"));
    
    // Columna M (13) - Link al PDF
    const celdaM = sheet.getRange(fila, 13);
    celdaM.clearFormat(); 
    celdaM.setFormula('=HYPERLINK("' + pdfFile.getUrl() + '"; "Ver detalle")');
    celdaM.setFontLine("underline").setFontColor("#000000"); 
    
    return "Éxito total";
  } catch(e) {
    return "Excel OK, pero falló PDF: " + e.toString();
  }
}

