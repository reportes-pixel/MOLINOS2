/**
 * GESTIÓN DE CONFIGURACIÓN Y UNIDADES CON AUDITORÍA
 */



function enviarAlertaJavier(autor, mod, detalle) {
  const body = `LOG DE CAMBIOS\nAutor: ${autor}\nMódulo: ${mod}\nDetalle: ${detalle}\nFecha: ${new Date()}`;
  MailApp.sendEmail("JAVIER.PUENTE.MX@GMAIL.COM", "⚠️ Alerta de Cambio en Sistema", body);
}





function obtenerDatosMaestros() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const confSh = ss.getSheetByName("CONFIGURACION");
  const unitSh = ss.getSheetByName("UNIDADES");

  const globalesRaw = confSh.getRange("A1:B6").getValues();
  const globales = {};
  globalesRaw.forEach(r => { if(r[0]) globales[r[0]] = r[1]; });

  const excepciones = confSh.getRange("I2:K" + Math.max(confSh.getLastRow(), 2)).getValues()
    .filter(r => r[0] !== "").map(r => ({ unidad: r[0], cuota: r[1], pierde: r[2] }));

  const catalogo = unitSh.getRange("A2:E" + Math.max(unitSh.getLastRow(), 2)).getValues()
    .filter(r => r[0] !== "").map(r => ({ id: r[0], depto: r[1], prop: r[2], rent: r[3], mail: r[4] }));

  return { globales, excepciones, catalogo };
}


function actualizarUnidad_Final(u) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName("UNIDADES");
  const data = sh.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == u.id) {
      sh.getRange(i + 1, 2, 1, 4).setValues([[u.depto, u.prop, u.rent, u.mail]]);
      enviarAlertaJavier("ADMIN_SESION", "CATÁLOGO", `Actualizó datos de ${u.id}`);
      return { success: true };
    }
  }
}




function guardarTodo_Final(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const confSh = ss.getSheetByName("CONFIGURACION");

  // 1. Guardar Globales (A1:B6)
  const filasGlobales = [
    ["MENSUALIDAD_BASE", payload.globales.MENSUALIDAD_BASE],
    ["TASA_PRONTO_PAGO", payload.globales.TASA_PRONTO_PAGO],
    ["TASA_RECARGO", payload.globales.TASA_RECARGO],
    ["DIA_LIMITE_PP", payload.globales.DIA_LIMITE_PP],
    ["DIA_LIMITE_NORMAL", payload.globales.DIA_LIMITE_NORMAL],
    ["MENSUALIDAD_PRONTO_PAGO", payload.globales.MENSUALIDAD_PRONTO_PAGO]
  ];
  confSh.getRange(1, 1, 6, 2).setValues(filasGlobales);

  // 2. Guardar Montos Especiales (I2:K)
  const ultimaFila = confSh.getLastRow();
  if (ultimaFila >= 2) confSh.getRange("I2:K" + ultimaFila).clearContent();

  if (payload.excepciones.length > 0) {
    const filasEx = payload.excepciones.map(e => [e.unidad, e.cuota, e.pierde]);
    confSh.getRange(2, 9, filasEx.length, 3).setValues(filasEx);
  }

  enviarAlertaJavier("ADMIN_SESION", "FINANZAS", "Cambio masivo guardado.");
  return { success: true };
}