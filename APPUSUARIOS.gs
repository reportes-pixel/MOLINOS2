/**
 * VALIDADOR DE LOGIN PARA RESIDENTES
 */
function validarLoginResidente(idUnidad, pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("UNIDADES");
  const data = sheet.getDataRange().getValues().slice(1);
  
  // Limpiamos el ID que viene del selector para comparar
  const target = String(idUnidad).replace("-","").trim().toUpperCase();
  
  const usuario = data.find(row => 
    String(row[0]).replace("-","").trim().toUpperCase() === target && 
    String(row[5]).trim() === String(pin).trim()
  );
  
  if (usuario) {
    return { 
      success: true, 
      idUnidad: idUnidad, // Mantenemos el ID original para la búsqueda de datos
      nombre: usuario[2] || usuario[3] || "Residente" 
    };
  } else {
    return { success: false, message: "Unidad o PIN incorrectos." };
  }
}




/**
 * Obtiene la lista de departamentos para el portal de residentes
 * Lee la hoja UNIDADES, columnas A y B
 */
function getListaUnidadesPortal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("UNIDADES");
  if (!sheet) return [];
  
  const data = sheet.getDataRange().getValues().slice(1); // Omitir encabezados
  
  // Creamos la lista con ID (Col A) y Nombre (Col B)
  return data.map(row => {
    return {
      id: String(row[0]).trim(),
      departamento: String(row[1]).trim()
    };
  }).filter(u => u.id !== ""); // Filtramos filas vacías
}




/**
 * MOTOR DE DATOS DEL PORTAL (Conectado a la función Maestra)
 */
function getDatosPortalResidente(idUnidad) {
  // Limpiamos el ID para evitar errores de guiones o espacios
  const targetID = String(idUnidad).replace("-", "").trim().toUpperCase();

  // LLAMAMOS DIRECTO A TU FUNCIÓN ORIGINAL PARA GARANTIZAR 100% DE IGUALDAD
  const resultadoMaestro = generarEstadoCuentaWebApp(targetID, false);

  if (!resultadoMaestro.success || !resultadoMaestro.detalles || resultadoMaestro.detalles.length === 0) {
    return { success: false, message: "Error al procesar el estado de cuenta." };
  }

  // Extraemos la información procesada por tu lógica perfecta
  const data = resultadoMaestro.detalles[0];

  // Buscamos el nombre del propietario
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const unidadesSheet = ss.getSheetByName("UNIDADES");
  let nombreVecino = "Residente";
  if (unidadesSheet) {
    const uData = unidadesSheet.getDataRange().getValues();
    const filaU = uData.find(r => String(r[0]).replace("-","").trim().toUpperCase() === targetID);
    if (filaU) nombreVecino = filaU[2] || filaU[3] || "Residente";
  }

  // Devolvemos el paquete de datos en dos formatos: uno para celular y otro para PDF
  return {
    success: true,
    nombre: nombreVecino,
    idUnidad: targetID,
    resumen: {
      totalCargos: data.totalCargos,
      totalPagos: data.totalPagos,
      saldoGuardado: data.saldoAFavor, // El sobrante de Hoja Pagos
      totalAPagar: data.deudaNeta
    },
    historial_pdf: data.historial, // Para imprimir idéntico al admin (Cronológico normal)
    historial_movil: [...data.historial].reverse() // Para el celular (Recientes arriba)
  };
}



/**
 * ESTO VA EN UN ARCHIVO .GS
 * Genera el PDF real usando el motor de Google.
 */
function generarPDFServidor(idUnidad, htmlTabla) {
  const nombreArchivo = "Estado_Cuenta_" + idUnidad + ".pdf";
  
  // Estilos obligatorios para que el PDF se vea bien
  const estilos = `
    <style>
      body { font-family: Arial, sans-serif; padding: 20px; color: #111827; }
      h1 { font-size: 24px; text-transform: uppercase; margin-bottom: 5px; }
      h2 { font-size: 14px; color: #4b5563; text-transform: uppercase; margin-bottom: 20px; }
      .tabla-contable { width: 100%; border-collapse: collapse; font-size: 10px; }
      .tabla-contable th, .tabla-contable td { border: 1px solid #9ca3af; padding: 8px; text-align: left; }
      .tabla-contable th { background-color: #f3f4f6; font-weight: bold; }
      .resumen { margin-bottom: 20px; width: 100%; }
      .resumen td { border: 1px solid #d1d5db; padding: 10px; text-align: center; background: #f9fafb; }
    </style>
  `;

  const htmlFinal = estilos + htmlTabla;
  
  // Convertimos a PDF
  const blob = HtmlService.createHtmlOutput(htmlFinal)
    .getAs('application/pdf')
    .setName(nombreArchivo);
  
  // Enviamos los bytes codificados al cliente
  return Utilities.base64Encode(blob.getBytes());
}


/**
 * Obtiene la lista de unidades con saldos pendientes para el Portal de Residentes
 */
function obtenerRelacionAdeudos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
  if (!sheet) return { success: false, message: "Base de datos no encontrada." };

  const data = sheet.getDataRange().getValues();
  const resumenDeudores = {};

  for (let i = 1; i < data.length; i++) {
    const unidad = String(data[i][1]).trim();
    const monto = Number(data[i][4]);
    const estado = String(data[i][5]).toUpperCase().trim();

    if (estado === "PENDIENTE" && unidad) {
      if (!resumenDeudores[unidad]) {
        resumenDeudores[unidad] = { unidad: unidad, cargos: 0, total: 0 };
      }
      resumenDeudores[unidad].cargos += 1;
      resumenDeudores[unidad].total += monto;
    }
  }

  const listaOrdenada = Object.values(resumenDeudores).sort((a, b) => b.total - a.total);
  return { success: true, lista: listaOrdenada };
}