// ARCHIVO: Competencias.gs (o Código.gs)
// AVISO: Borrá cualquier función doGet(e) que haya estado acá.

// 1. Genera los links en la Columna O
function llenarColumnaO() {
  const sheet = SpreadsheetApp.openById(SHEET_DESTINO_ID).getActiveSheet();
  const data = sheet.getDataRange().getValues();
  
  for(let i=1; i<data.length; i++) {
    const nombre = data[i][1];        // Col B
    const colO = data[i][14];          // Col O
    const entrevistador = data[i][17]; // Col R
    
    if(nombre && !colO) {
      const url = URL_FORM_COMPETENCIAS + '?tipo=competencias&nombre=' + encodeURIComponent(nombre) + '&entrevistador=' + encodeURIComponent(entrevistador);
      sheet.getRange(i+1, 15).setFormula('=HYPERLINK("' + url + '";"Evaluar competencias")');
    }
  }
}
