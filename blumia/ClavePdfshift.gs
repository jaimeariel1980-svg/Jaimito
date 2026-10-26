/* Clave de PDFShift («PDF cliente»): vive en la propiedad del script PDFSHIFT_KEY, no en el código. */
function pdfshiftKey_() {
  const k = PropertiesService.getScriptProperties().getProperty('PDFSHIFT_KEY');
  if (!k) throw new Error('Falta cargar la propiedad PDFSHIFT_KEY en Apps Script (Configuración del proyecto → Propiedades del script).');
  return k;
}
