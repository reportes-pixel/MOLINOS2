// ==============================================================================
// GESTOR MANUAL DE CARGOS Y CONVENIOS (EDICIÓN Y BORRADO)
// ==============================================================================

function abrirForm_EditorCargos() {
  const html = HtmlService.createTemplateFromFile('Form_EditorCargos')
      .evaluate()
      .setWidth(1000)
      .setHeight(700);
  SpreadsheetApp.getUi().showModalDialog(html, '🛠️ Gestor Manual de Cargos y Convenios');
}

// 1. Obtener lista de unidades para el selector
function getUnidadesParaEditor() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("UNIDADES");
  const data = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().flat().filter(String);
  return [...new Set(data)].sort();
}

// 2. Obtener todos los cargos de una unidad específica
function getCargosUnidad(unidad) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("CARGOS_Y_DEUDAS");
  const data = sh.getDataRange().getValues().slice(1);
  
  const result = [];
  data.forEach((r) => {
    if (String(r[1]).trim().toUpperCase() === unidad.toUpperCase()) {
      result.push({
        idCargo: r[0],
        fecha: Utilities.formatDate(new Date(r[3]), Session.getScriptTimeZone(), "dd/MM/yyyy"),
        concepto: r[2],
        monto: r[4],
        estado: String(r[5]).toUpperCase()
      });
    }
  });
  
  // Invertimos para ver los más recientes arriba
  return result.reverse();
}

// 3. Guardar cambios en un cargo específico
function guardarEdicionCargo(idCargo, nuevoConcepto, nuevoMonto) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("CARGOS_Y_DEUDAS");
  const data = sh.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0] === idCargo) {
       sh.getRange(i + 1, 3).setValue(nuevoConcepto); // Columna C: Concepto
       sh.getRange(i + 1, 5).setValue(Number(nuevoMonto)); // Columna E: Monto Base
       return {success: true};
    }
  }
  return {success: false, message: "Cargo no encontrado en la base de datos."};
}

// 4. Eliminar un cargo (Ej. borrar un recargo)
function eliminarCargoManual(idCargo) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("CARGOS_Y_DEUDAS");
  const data = sh.getDataRange().getValues();
  
  for(let i = 1; i < data.length; i++) {
    if(data[i][0] === idCargo) {
       sh.deleteRow(i + 1);
       return {success: true};
    }
  }
  return {success: false, message: "Cargo no encontrado."};
}