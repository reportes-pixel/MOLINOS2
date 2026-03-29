// ==============================================================================
// 00_WEBAPP.gs - MOTOR DEL PORTAL WEB INDEPENDIENTE (VERSIÓN IFRAME)
// ==============================================================================

function doGet(e) {
  // Si el sistema pide un formulario específico (Ej. la URL termina en ?v=Form_Pagos)
  if (e.parameter.v) {
    try {
      return HtmlService.createTemplateFromFile(e.parameter.v)
          .evaluate()
          .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); // Permite incrustarlo en el iFrame
    } catch(err) {
      return HtmlService.createHtmlOutput('<b>Error:</b> Módulo no encontrado.');
    }
  }
  
  // Si no pide nada, carga el Index principal (Login y Dashboard)
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('Portal Molinos Real II')
      .setFaviconUrl('https://cdn-icons-png.flaticon.com/512/3135/3135673.png')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Esta función le pasa la URL de la Web App al Index.html para saber dónde apuntar los iFrames
function getAppUrl() {
  return ScriptApp.getService().getUrl();
}