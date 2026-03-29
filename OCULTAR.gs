/**
 * Oculta o Muestra las hojas sensibles del sistema.
 * Solo accesible para el rol ADMIN.
 */
function toggleSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // Lista exacta de tus hojas sensibles
  const hojasSensibles = [
    "CONFIGURACION", 
    "REGISTRO_PAGOS", 
    "CARGOS_Y_DEUDAS", 
    "EGRESOS", 
    "SALDOS_A_FAVOR", 
    "RENTAS_ESTACIONAMIENTO",
    "DEUDAS VIEJAS",
    "USER" 
  ];

  // Revisamos el estado de la primera hoja de la lista para decidir qué hacer
  const primeraHoja = ss.getSheetByName(hojasSensibles[0]);
  if (!primeraHoja) {
    ui.alert("Error", "No se encontró la hoja de CONFIGURACION para validar el estado.", ui.ButtonSet.OK);
    return;
  }

  const estanOcultas = primeraHoja.isSheetHidden();

  if (estanOcultas) {
    // VAMOS A MOSTRARLAS
    hojasSensibles.forEach(nombre => {
      const hoja = ss.getSheetByName(nombre);
      if (hoja) {
        hoja.showSheet();
      }
    });
    // CORRECCIÓN: toast pertenece a la Hoja de Cálculo, no a la UI
    ss.toast("Bases de datos visibles para edición.", "🔒 MODO ADMIN", 3);
  } else {
    // VAMOS A OCULTARLAS
    hojasSensibles.forEach(nombre => {
      const hoja = ss.getSheetByName(nombre);
      // Nunca ocultamos todas, siempre debe quedar una visible (ej. UNIDADES)
      if (hoja && ss.getSheets().filter(s => !s.isSheetHidden()).length > 1) {
        hoja.hideSheet();
      }
    });
    // CORRECCIÓN: toast pertenece a la Hoja de Cálculo, no a la UI
    ss.toast("Bases de datos protegidas y ocultas.", "🔒 SISTEMA BLINDADO", 3);
  }
}