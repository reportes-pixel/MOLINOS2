/**
 * OCULTA SIEMPRE las hojas sensibles
 * Esta función es para TRIGGERS (onOpen / time-driven)
 */
function ocultarHojasSensibles() {

  const NOMBRES_HOJAS_SENSIBLES = [
    "CONFIG_MULTAS",
    "UNIDADES",
    "CONFIGURACION",
    "USUARIOS",
    "SALDOS_A_FAVOR",
    "CARGOS_Y_DEUDAS",
    "REGISTRO_PAGOS",
    "EGRESOS"
  ];

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  NOMBRES_HOJAS_SENSIBLES.forEach(nombre => {
    const hoja = ss.getSheetByName(nombre);
    if (hoja && !hoja.isSheetHidden()) {
      hoja.hideSheet();
    }
  });
}
