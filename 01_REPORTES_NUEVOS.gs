////////////////////////////////////////////////////////////////////////

// ==============================================================================
// MÓDULO FINANCIERO MAESTRO (INTEGRADO PARA EXCEL Y WEB APP)
// ==============================================================================

function procesarPeticionFinancieroMaster(strInicio, strFin) {
  const partesIni = strInicio.split("-");
  const partesFin = strFin.split("-");
  
  const fechaInicio = new Date(partesIni[0], parseInt(partesIni[1]) - 1, 1);
  const fechaFin = new Date(partesFin[0], parseInt(partesFin[1]), 0); 
  fechaFin.setHours(23, 59, 59, 999);

  const nombresMeses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  
  let periodoTexto = "";
  if (strInicio === strFin) {
    periodoTexto = `${nombresMeses[fechaInicio.getMonth()]} ${fechaInicio.getFullYear()}`;
  } else {
    periodoTexto = `${nombresMeses[fechaInicio.getMonth()]} ${fechaInicio.getFullYear()} a ${nombresMeses[fechaFin.getMonth()]} ${fechaFin.getFullYear()}`;
  }

  // Devolvemos el resultado directo a la Web App
  return construirExcelFinancieroMaster(fechaInicio, fechaFin, periodoTexto);
}

function construirExcelFinancieroMaster(fechaInicio, fechaFin, periodoTexto) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
  const pagosSheet = ss.getSheetByName("REGISTRO_PAGOS");
  const egresosSheet = ss.getSheetByName("EGRESOS");

  if (!cargosSheet || !pagosSheet || !egresosSheet) {
      return { success: false, message: "Error: Faltan hojas base en el documento." };
  }

  let totalCargosGenerados = 0, totalCargosPagados = 0, totalMontoRecibido = 0, totalMontoEgreso = 0;
  const pagosPorConcepto = {}; const egresosPorCategoria = {};
  
  // ARREGLOS PARA LAS LISTAS DETALLADAS
  const listaDetalladaPagos = [];
  const listaDetalladaEgresos = [];

  // 1. PROCESAR CARGOS
  cargosSheet.getDataRange().getValues().slice(1).forEach(row => {
      const mesCorte = new Date(row[3]), monto = Number(row[4]) || 0;
      if (mesCorte >= fechaInicio && mesCorte <= fechaFin) {
          totalCargosGenerados += monto;
          if (String(row[5]).toUpperCase() === "PAGADO") totalCargosPagados += monto;
      }
  });

  // 2. PROCESAR PAGOS Y LLENAR LISTA DETALLADA
  pagosSheet.getDataRange().getValues().slice(1).forEach(row => {
      const fechaPago = new Date(row[1]);
      const depto = String(row[2] || "N/A"); 
      const montoRecibido = Number(row[4]) || 0; 
      const concepto = String(row[9] || "Sin Concepto"); 

      if (fechaPago >= fechaInicio && fechaPago <= fechaFin) {
          totalMontoRecibido += montoRecibido;
          pagosPorConcepto[concepto] = (pagosPorConcepto[concepto] || 0) + montoRecibido;
          listaDetalladaPagos.push([fechaPago, depto, concepto, montoRecibido]);
      }
  });

  // 3. PROCESAR EGRESOS Y LLENAR LISTA DETALLADA
  egresosSheet.getDataRange().getValues().slice(1).forEach(row => {
      const fechaRegistro = new Date(row[1]);
      const categoria = String(row[3] || "Sin Categoría").trim(); 
      const proveedor = String(row[4] || "Sin Proveedor").trim(); 
      const montoEgreso = Number(row[5]) || 0; 

      if (fechaRegistro >= fechaInicio && fechaRegistro <= fechaFin) {
          totalMontoEgreso += montoEgreso;
          egresosPorCategoria[categoria] = (egresosPorCategoria[categoria] || 0) + montoEgreso;
          listaDetalladaEgresos.push([fechaRegistro, categoria, proveedor, montoEgreso]);
      }
  });

  // ORDENAR LAS LISTAS CRONOLÓGICAMENTE
  listaDetalladaPagos.sort((a, b) => a[0] - b[0]);
  listaDetalladaEgresos.sort((a, b) => a[0] - b[0]);

  // ==============================================================================
  // CREACIÓN DEL EXCEL ORIGINAL (BLOQUES A, B, C, D, E)
  // ==============================================================================
  const nombreHoja = "REPORTE_FINANCIERO_COMPLETO";
  let reporteSheet = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);
  reporteSheet.clear(); // Limpiamos todo (datos y colores viejos)

  let row = 1;
  reporteSheet.getRange(`A${row}`).setValue("REPORTE FINANCIERO EXTENDIDO Y DETALLADO").setFontWeight('bold');
  reporteSheet.getRange(`A${row+1}`).setValue(`Periodo: ${periodoTexto}`).setFontWeight('bold');
  row += 3;

  // BLOQUE A: RESUMEN
  const datosResumen = [
      ["--- RESUMEN DE CARGOS Y SALDOS DEVENGADOS ---", "", "", ""],
      ["Total Cargos Generados", totalCargosGenerados, "", ""],
      ["Total Cargos Pagados", totalCargosPagados, "", ""],
      ["Total Cargos PENDIENTES", totalCargosGenerados - totalCargosPagados, "", ""],
      ["", "", "", ""],
      ["--- SALDO NETO DE EFECTIVO (Flujo de Caja) ---", "", "", ""],
      ["Total Recibido en Pagos (INGRESOS)", totalMontoRecibido, "", ""],
      ["Total Egresos Registrados (GASTOS)", totalMontoEgreso, "", ""],
      ["SALDO NETO DE EFECTIVO", totalMontoRecibido - totalMontoEgreso, "", ""]
  ];
  
  reporteSheet.getRange(row, 1, datosResumen.length, 4).setValues(datosResumen);
  reporteSheet.getRange(`A${row}`).setFontWeight('bold').setBackground("#D9EAD3");
  reporteSheet.getRange(`A${row+5}`).setFontWeight('bold').setBackground("#B4C6E7");
  reporteSheet.getRange(`A${row+8}:B${row+8}`).setFontWeight('bold').setBackground("#FFE599");
  reporteSheet.getRange(row + 1, 2, 8, 1).setNumberFormat("$#,##0.00");
  row += datosResumen.length + 1;

  // BLOQUE B: AGRUPADO INGRESOS
  reporteSheet.getRange(`A${row}`).setValue("DETALLE DE INGRESOS (Agrupado)").setFontWeight('bold').setBackground("#B6D7A8"); row++;
  const detallePagos = [["CONCEPTO DE PAGO", "MONTO TOTAL"]];
  Object.entries(pagosPorConcepto).forEach(([k, v]) => detallePagos.push([k, v]));
  if (detallePagos.length > 1) {
      reporteSheet.getRange(row, 1, detallePagos.length, 2).setValues(detallePagos);
      reporteSheet.getRange(row, 2, detallePagos.length, 1).setNumberFormat("$#,##0.00");
      reporteSheet.getRange(row, 1, 1, 2).setFontWeight('bold');
  }
  row += detallePagos.length + 1;

  // BLOQUE C: AGRUPADO EGRESOS
  reporteSheet.getRange(`A${row}`).setValue("DETALLE DE EGRESOS (Agrupado)").setFontWeight('bold').setBackground("#F4CCCC"); row++;
  const detalleEgresos = [["CATEGORÍA", "MONTO TOTAL"]];
  Object.entries(egresosPorCategoria).forEach(([k, v]) => detalleEgresos.push([k, v]));
  if (detalleEgresos.length > 1) {
      reporteSheet.getRange(row, 1, detalleEgresos.length, 2).setValues(detalleEgresos);
      reporteSheet.getRange(row, 2, detalleEgresos.length, 1).setNumberFormat("$#,##0.00");
      reporteSheet.getRange(row, 1, 1, 2).setFontWeight('bold');
  }
  row += detalleEgresos.length + 2;

  // BLOQUE D: LISTA DETALLADA DE PAGOS
  reporteSheet.getRange(`A${row}`).setValue("LISTA DETALLADA DE PAGOS RECIBIDOS (Ingresos del Periodo)").setFontWeight('bold').setBackground("#CFE2F3"); row++;
  if (listaDetalladaPagos.length > 0) {
      const headersPagos = ["FECHA DE PAGO", "DEPARTAMENTO", "CONCEPTO PAGADO", "MONTO RECIBIDO"];
      reporteSheet.getRange(row, 1, 1, 4).setValues([headersPagos]).setFontWeight('bold').setBackground("#EAEAEA");
      row++;
      
      reporteSheet.getRange(row, 1, listaDetalladaPagos.length, 4).setValues(listaDetalladaPagos);
      reporteSheet.getRange(row, 1, listaDetalladaPagos.length, 1).setNumberFormat("dd/MM/yyyy");
      reporteSheet.getRange(row, 4, listaDetalladaPagos.length, 1).setNumberFormat("$#,##0.00");
      
      row += listaDetalladaPagos.length;
      reporteSheet.getRange(row, 3).setValue("TOTAL PAGOS DEL PERIODO:").setFontWeight('bold');
      reporteSheet.getRange(row, 4).setValue(totalMontoRecibido).setFontWeight('bold').setNumberFormat("$#,##0.00").setBackground("#D9EAD3");
  } else {
      reporteSheet.getRange(`A${row}`).setValue("No se encontraron registros individuales de pagos en este periodo.");
  }
  row += 3;

  // BLOQUE E: LISTA DETALLADA DE EGRESOS
  reporteSheet.getRange(`A${row}`).setValue("LISTA DETALLADA DE EGRESOS (Gastos del Periodo)").setFontWeight('bold').setBackground("#F4CCCC"); row++;
  if (listaDetalladaEgresos.length > 0) {
      const headersEgresos = ["FECHA DE REGISTRO", "CONCEPTO / CATEGORÍA", "PROVEEDOR", "MONTO GASTADO"];
      reporteSheet.getRange(row, 1, 1, 4).setValues([headersEgresos]).setFontWeight('bold').setBackground("#EAEAEA");
      row++;
      
      reporteSheet.getRange(row, 1, listaDetalladaEgresos.length, 4).setValues(listaDetalladaEgresos);
      reporteSheet.getRange(row, 1, listaDetalladaEgresos.length, 1).setNumberFormat("dd/MM/yyyy");
      reporteSheet.getRange(row, 4, listaDetalladaEgresos.length, 1).setNumberFormat("$#,##0.00");
      
      row += listaDetalladaEgresos.length;
      reporteSheet.getRange(row, 3).setValue("TOTAL EGRESOS DEL PERIODO:").setFontWeight('bold');
      reporteSheet.getRange(row, 4).setValue(totalMontoEgreso).setFontWeight('bold').setNumberFormat("$#,##0.00").setBackground("#F4CCCC");
  } else {
      reporteSheet.getRange(`A${row}`).setValue("No se encontraron registros individuales de egresos en este periodo.");
  }

  reporteSheet.autoResizeColumns(1, 4);
  
  // Opcional: Activamos la hoja en el Excel para que la vea el usuario de escritorio
  try { reporteSheet.activate(); } catch(e) {}

  // ==============================================================================
  // ENVIAR TODO A LA WEB APP PARA EL PDF MAESTRO
  // ==============================================================================
  
  // Transformamos las fechas reales del servidor a texto DD/MM/YYYY para la pantalla
  const detallePagosFront = listaDetalladaPagos.map(fila => [
    Utilities.formatDate(fila[0], Session.getScriptTimeZone(), "dd/MM/yyyy"), 
    fila[1], fila[2], fila[3]
  ]);
  
  const detalleEgresosFront = listaDetalladaEgresos.map(fila => [
    Utilities.formatDate(fila[0], Session.getScriptTimeZone(), "dd/MM/yyyy"), 
    fila[1], fila[2], fila[3]
  ]);

  return {
      success: true,
      periodoTexto: periodoTexto,
      cargos: {
          generados: totalCargosGenerados,
          pagados: totalCargosPagados,
          pendientes: totalCargosGenerados - totalCargosPagados
      },
      flujo: {
          ingresos: totalMontoRecibido,
          egresos: totalMontoEgreso,
          saldoNeto: totalMontoRecibido - totalMontoEgreso
      },
      agrupadoIngresos: Object.entries(pagosPorConcepto).map(([k, v]) => ({concepto: k, monto: v})),
      agrupadoEgresos: Object.entries(egresosPorCategoria).map(([k, v]) => ({categoria: k, monto: v})),
      listaPagos: detallePagosFront,
      listaEgresos: detalleEgresosFront
  };
}



function obtenerDatosCorteWebApp(fechaBusqueda, capturistaFiltro) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("REGISTRO_PAGOS");
  const data = sheet.getDataRange().getValues();
  const resultados = [];
  const datosParaExcel = [];
  let totalMonto = 0;

  // Convertir fecha de YYYY-MM-DD a DD/MM/YYYY para comparar con la BD
  const partes = fechaBusqueda.split('-');
  const fechaFormateada = `${partes[2]}/${partes[1]}/${partes[0]}`;

  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue; 
    
    let fechaCelda = Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), "dd/MM/yyyy");
    let capturistaCelda = data[i][3];
    
    if (fechaCelda === fechaFormateada) {
      if (capturistaFiltro && capturistaFiltro !== "TODOS" && capturistaCelda !== capturistaFiltro) continue;
      
      let monto = parseFloat(data[i][4]) || 0;
      let concepto = data[i][9] || "Pago registrado";

      resultados.push({
        folio: data[i][0],
        unidad: data[i][2],
        capturista: capturistaCelda,
        monto: monto,
        concepto: concepto
      });

      datosParaExcel.push([data[i][0], fechaCelda, data[i][2], capturistaCelda, monto, concepto]);
      totalMonto += monto;
    }
  }

  // ==========================================================
  // LA MAGIA DEL REPORTE ÚNICO (BORRA Y SOBREESCRIBE)
  // ==========================================================
  let hojaCorte = ss.getSheetByName("REPORTE_CORTE");
  if (!hojaCorte) {
    hojaCorte = ss.insertSheet("REPORTE_CORTE"); // Si no existe, la crea la primera vez
  } else {
    hojaCorte.clear(); // Si ya existe, la limpia por completo
  }

  // Escribimos el título
  hojaCorte.getRange("A1").setValue(`CORTE DIARIO: ${fechaFormateada} - Filtro: ${capturistaFiltro || "TODOS"}`).setFontWeight("bold").setFontSize(14);
  
  if (datosParaExcel.length > 0) {
    const encabezados = ["Folio", "Fecha", "Unidad", "Capturista", "Monto", "Concepto"];
    const matriz = [encabezados].concat(datosParaExcel);
    
    // Pegamos todos los datos
    hojaCorte.getRange(3, 1, matriz.length, matriz[0].length).setValues(matriz);
    
    // Formato de tabla
    hojaCorte.getRange(3, 1, 1, matriz[0].length).setFontWeight("bold").setBackground("#4f46e5").setFontColor("white");
    hojaCorte.getRange(matriz.length + 4, 4).setValue("TOTAL RECAUDADO:").setFontWeight("bold");
    hojaCorte.getRange(matriz.length + 4, 5).setValue(totalMonto).setFontWeight("bold").setNumberFormat("$#,##0.00");
    hojaCorte.autoResizeColumns(1, 6);
  } else {
    hojaCorte.getRange("A3").setValue("No hay ingresos registrados para esta fecha y filtro.");
  }
  // ==========================================================

  return {
    success: true,
    fecha: fechaFormateada,
    pagos: resultados,
    total: totalMonto
  };
}


function obtenerDatosCortePeriodoWebApp(fechaInicio, fechaFin, capturistaFiltro) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("REGISTRO_PAGOS");
  const data = sheet.getDataRange().getValues();
  const resultados = [];
  const datosParaExcel = [];
  let totalMonto = 0;

  const start = new Date(fechaInicio + "T00:00:00").getTime();
  const end = new Date(fechaFin + "T23:59:59").getTime();
  
  const fInicioStr = Utilities.formatDate(new Date(fechaInicio + "T00:00:00"), Session.getScriptTimeZone(), "dd/MM/yyyy");
  const fFinStr = Utilities.formatDate(new Date(fechaFin + "T23:59:59"), Session.getScriptTimeZone(), "dd/MM/yyyy");

  for (let i = 1; i < data.length; i++) {
    if (!data[i][1]) continue; 
    
    let dateObj = new Date(data[i][1]);
    let time = dateObj.getTime();
    let capturistaCelda = data[i][3];
    
    if (time >= start && time <= end) {
      if (capturistaFiltro && capturistaFiltro !== "TODOS" && capturistaCelda !== capturistaFiltro) continue;
      
      let montoPago = parseFloat(data[i][4]) || 0;
      let fechaStr = Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "dd/MM/yyyy");
      let conceptoPago = data[i][9] || "Pago registrado";

      resultados.push({
        fechaStr: fechaStr,
        folio: data[i][0],
        unidad: data[i][2],
        capturista: capturistaCelda,
        monto: montoPago,
        concepto: conceptoPago
      });
      
      datosParaExcel.push([data[i][0], fechaStr, data[i][2], capturistaCelda, montoPago, conceptoPago]);
      totalMonto += montoPago;
    }
  }

  // ==========================================================
  // LA MAGIA DEL REPORTE ÚNICO (BORRA Y SOBREESCRIBE)
  // ==========================================================
  let hojaCorte = ss.getSheetByName("REPORTE_CORTE");
  if (!hojaCorte) {
    hojaCorte = ss.insertSheet("REPORTE_CORTE"); 
  } else {
    hojaCorte.clear(); 
  }

  hojaCorte.getRange("A1").setValue(`CORTE POR PERÍODO: Del ${fInicioStr} al ${fFinStr} - Filtro: ${capturistaFiltro || "TODOS"}`).setFontWeight("bold").setFontSize(14);
  
  if (datosParaExcel.length > 0) {
    const encabezados = ["Folio", "Fecha", "Unidad", "Capturista", "Monto", "Concepto"];
    const matriz = [encabezados].concat(datosParaExcel);
    
    hojaCorte.getRange(3, 1, matriz.length, matriz[0].length).setValues(matriz);
    hojaCorte.getRange(3, 1, 1, matriz[0].length).setFontWeight("bold").setBackground("#4f46e5").setFontColor("white");
    hojaCorte.getRange(matriz.length + 4, 4).setValue("TOTAL RECAUDADO:").setFontWeight("bold");
    hojaCorte.getRange(matriz.length + 4, 5).setValue(totalMonto).setFontWeight("bold").setNumberFormat("$#,##0.00");
    hojaCorte.autoResizeColumns(1, 6);
  } else {
    hojaCorte.getRange("A3").setValue("No hay ingresos registrados en este rango y filtro.");
  }
  // ==========================================================

  return {
    success: true,
    fechaInicioStr: fInicioStr,
    fechaFinStr: fFinStr,
    pagos: resultados,
    total: totalMonto
  };
}




function obtenerReporteFinancieroWebApp(mesStr) { 
  // mesStr viene en formato "YYYY-MM"
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Calcular Ingresos (REGISTRO_PAGOS)
  const sheetPagos = ss.getSheetByName("REGISTRO_PAGOS");
  const dataPagos = sheetPagos ? sheetPagos.getDataRange().getValues() : [];
  let totalIngresos = 0;
  let ingresosDetalle = [];

  for(let i = 1; i < dataPagos.length; i++) {
    if(!dataPagos[i][1]) continue;
    let fecha = new Date(dataPagos[i][1]);
    let mesPago = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "yyyy-MM");
    if(mesPago === mesStr) {
      let monto = parseFloat(dataPagos[i][4]) || 0;
      totalIngresos += monto;
      // Guardamos: [Folio, Fecha, Unidad, Monto, Concepto]
      ingresosDetalle.push([dataPagos[i][0], Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy"), dataPagos[i][2], monto, dataPagos[i][9]]);
    }
  }

  // 2. Calcular Egresos (EGRESOS)
  const sheetEgresos = ss.getSheetByName("EGRESOS");
  const dataEgresos = sheetEgresos ? sheetEgresos.getDataRange().getValues() : [];
  let totalEgresos = 0;
  let egresosDetalle = [];

  for(let i = 1; i < dataEgresos.length; i++) {
    if(!dataEgresos[i][0]) continue; // Asumiendo que col A es ID Egreso
    let mesEgreso = dataEgresos[i][1]; // Asumiendo que col B es el mes YYYY-MM
    if(mesEgreso === mesStr) {
      let monto = parseFloat(dataEgresos[i][5]) || 0; // Monto en col F
      totalEgresos += monto;
      // Guardamos: [Folio, Fecha, Proveedor, Monto, Concepto]
      egresosDetalle.push([dataEgresos[i][0], dataEgresos[i][6], dataEgresos[i][4], monto, dataEgresos[i][3]]); 
    }
  }

  const saldoNeto = totalIngresos - totalEgresos;

  // ==========================================================
  // 3. MAGIA EXCEL: Pestaña Única "REPORTE_FINANCIERO"
  // ==========================================================
  let hojaReporte = ss.getSheetByName("REPORTE_FINANCIERO");
  if(!hojaReporte) {
    hojaReporte = ss.insertSheet("REPORTE_FINANCIERO");
  } else {
    hojaReporte.clear();
  }

  hojaReporte.getRange("A1").setValue("REPORTE FINANCIERO MASTER: " + mesStr).setFontWeight("bold").setFontSize(14);

  // Cuadro de Resumen
  hojaReporte.getRange("A3:B3").merge().setValue("RESUMEN DEL MES").setFontWeight("bold").setBackground("#1e293b").setFontColor("white");
  hojaReporte.getRange("A4").setValue("Total Ingresos:");
  hojaReporte.getRange("B4").setValue(totalIngresos).setNumberFormat("$#,##0.00").setFontColor("green").setFontWeight("bold");
  hojaReporte.getRange("A5").setValue("Total Egresos:");
  hojaReporte.getRange("B5").setValue(totalEgresos).setNumberFormat("$#,##0.00").setFontColor("red").setFontWeight("bold");
  hojaReporte.getRange("A6").setValue("SALDO NETO:");
  hojaReporte.getRange("B6").setValue(saldoNeto).setNumberFormat("$#,##0.00").setFontWeight("bold").setBackground(saldoNeto >= 0 ? "#dcfce7" : "#fee2e2");

  // Detalle de Ingresos
  let filaActual = 9;
  hojaReporte.getRange(filaActual, 1, 1, 5).merge().setValue("DETALLE DE INGRESOS").setFontWeight("bold");
  filaActual++;
  if(ingresosDetalle.length > 0) {
    hojaReporte.getRange(filaActual, 1, 1, 5).setValues([["Folio", "Fecha", "Unidad", "Monto", "Concepto"]]).setBackground("#10b981").setFontColor("white").setFontWeight("bold");
    filaActual++;
    hojaReporte.getRange(filaActual, 1, ingresosDetalle.length, 5).setValues(ingresosDetalle);
    filaActual += ingresosDetalle.length;
  } else {
    hojaReporte.getRange(filaActual, 1).setValue("Sin ingresos registrados en este mes.");
    filaActual++;
  }

  filaActual += 2;
  // Detalle de Egresos
  hojaReporte.getRange(filaActual, 1, 1, 5).merge().setValue("DETALLE DE EGRESOS").setFontWeight("bold");
  filaActual++;
  if(egresosDetalle.length > 0) {
    hojaReporte.getRange(filaActual, 1, 1, 5).setValues([["Folio", "Fecha", "Proveedor", "Monto", "Concepto"]]).setBackground("#ef4444").setFontColor("white").setFontWeight("bold");
    filaActual++;
    hojaReporte.getRange(filaActual, 1, egresosDetalle.length, 5).setValues(egresosDetalle);
  } else {
    hojaReporte.getRange(filaActual, 1).setValue("Sin egresos registrados en este mes.");
  }
  
  hojaReporte.autoResizeColumns(1, 5);
  // ==========================================================

  return {
    success: true,
    mes: mesStr,
    ingresos: totalIngresos,
    egresos: totalEgresos,
    saldo: saldoNeto
  };
}





// ==============================================================================
// MÓDULO: ESTADO DE CUENTA (PDF WEB APP + EXCEL AUDITORÍA)
// ==============================================================================

function obtenerUnidadesParaEstadoCuenta() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let unidades = new Set();
  
  // Buscar en Cargos
  const cargos = ss.getSheetByName("CARGOS_Y_DEUDAS");
  if(cargos) {
    cargos.getRange(2, 2, cargos.getLastRow(), 1).getValues().forEach(r => {
      if(r[0]) unidades.add(String(r[0]).trim().toUpperCase());
    });
  }
  // Buscar en Pagos
  const pagos = ss.getSheetByName("REGISTRO_PAGOS");
  if(pagos) {
    pagos.getRange(2, 3, pagos.getLastRow(), 1).getValues().forEach(r => {
      if(r[0]) unidades.add(String(r[0]).trim().toUpperCase());
    });
  }
  
  const arrayUnidades = Array.from(unidades).sort();
  return arrayUnidades;
}

function generarEstadoCuentaWebApp(targetUnit, crearExcel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
  const pagosSheet = ss.getSheetByName("REGISTRO_PAGOS");
  
  if (!cargosSheet || !pagosSheet) return { success: false, message: "Faltan hojas base en el documento." };
  
  const cargos = cargosSheet.getDataRange().getValues().slice(1);
  const pagos = pagosSheet.getDataRange().getValues().slice(1);
  let transacciones = [];
  
  cargos.forEach(row => {
      const idUnidad = String(row[1]).trim().toUpperCase();
      if (targetUnit !== 'TODOS' && idUnidad !== targetUnit) return;
      transacciones.push({
          unidad: idUnidad, fecha: new Date(row[3]), concepto: String(row[2]),
          estado: String(row[5] || "PENDIENTE").toUpperCase(), cargo: Number(row[4]) || 0,
          pago: 0, tipo: 'CARGO', id: row[0]
      });
  });
  
  pagos.forEach(row => {
      const idUnidad = String(row[2]).trim().toUpperCase(); 
      if (targetUnit !== 'TODOS' && idUnidad !== targetUnit) return;
      const montoAplicado = Number(row[5]) || 0; 
      if (montoAplicado === 0) return; 

      transacciones.push({
          unidad: idUnidad, fecha: new Date(row[1]), concepto: `PAGO: ${String(row[9] || "Sin concepto")}`,
          estado: "PAGO", cargo: 0, pago: montoAplicado, tipo: 'PAGO', id: row[0]
      });
  });

  if (transacciones.length === 0) return { success: false, message: `No hay transacciones registradas para ${targetUnit}.` };
  
  transacciones.sort((a, b) => a.fecha - b.fecha);
  
  let saldoAcumulado = 0; 
  let totalCargos = 0;
  let totalPagos = 0;
  
  const datosParaFront = [];
  const datosParaExcel = [];
  
  transacciones.forEach(t => {
      saldoAcumulado = saldoAcumulado + t.cargo - t.pago; 
      totalCargos += t.cargo;
      totalPagos += t.pago;
      
      // Datos formateados para el HTML (PDF)
      datosParaFront.push([
          t.unidad, 
          Utilities.formatDate(t.fecha, Session.getScriptTimeZone(), "dd/MM/yyyy"), 
          t.concepto, 
          t.cargo, 
          t.pago, 
          saldoAcumulado
      ]);

      // Datos crudos para el Excel (fechas reales y números reales)
      datosParaExcel.push([t.unidad, t.fecha, t.concepto, t.estado, t.cargo, t.pago, saldoAcumulado]);
  });

  // ==============================================================================
  // CREAR PESTAÑA DE EXCEL (SOLO SI EL USUARIO LO PIDIÓ)
  // ==============================================================================
  if (crearExcel) {
      const nombreHoja = targetUnit === 'TODOS' ? `EDO_CTA_GLOBAL` : `EDO_CTA_${targetUnit}`;
      let reporteSheet = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);
      reporteSheet.clear();
      
      const headers = ["UNIDAD", "FECHA", "CONCEPTO", "ESTADO", "CARGO", "PAGO", "SALDO_ACUMULADO"];
      reporteSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#1e293b").setFontColor("white");
      reporteSheet.getRange(2, 1, datosParaExcel.length, headers.length).setValues(datosParaExcel);
      reporteSheet.getRange(2, 5, datosParaExcel.length, 3).setNumberFormat("$#,##0.00");
      reporteSheet.getRange(2, 2, datosParaExcel.length, 1).setNumberFormat("dd/MM/yyyy");
      reporteSheet.autoResizeColumns(1, headers.length);
  }

  return {
      success: true,
      unidad: targetUnit,
      totalCargos: totalCargos,
      totalPagos: totalPagos,
      saldoActual: saldoAcumulado,
      historial: datosParaFront
  };
}



// ==============================================================================
// MÓDULO: REPORTE DE DEUDORES (RANKING MOROSIDAD)
// ==============================================================================

function generarReporteDeudoresWebApp(crearExcel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
  const pagosSheet = ss.getSheetByName("REGISTRO_PAGOS");
  const unidadesSheet = ss.getSheetByName("UNIDADES");

  if (!cargosSheet || !pagosSheet || !unidadesSheet) {
      return { success: false, message: "Error: Faltan hojas base (CARGOS, PAGOS o UNIDADES)." };
  }

  // Obtenemos todas las unidades y armamos el diccionario inicial
  const units = unidadesSheet.getRange(2, 1, unidadesSheet.getLastRow() - 1, 1).getValues().flat().map(String);
  const saldos = units.reduce((acc, id) => { 
      if(id.trim() !== "") acc[id] = { totalCargos: 0, totalPagos: 0, saldo: 0, nombre: id }; 
      return acc; 
  }, {});

  const cargos = cargosSheet.getDataRange().getValues().slice(1);
  const pagos = pagosSheet.getDataRange().getValues().slice(1);

  // 1. Sumar Cargos (Solo Mensualidades y Multas según tu código original)
  cargos.forEach(row => {
      const idUnidad = String(row[1]).trim();
      const concepto = row[2] ? row[2].toString() : '';
      const monto = Number(row[4]) || 0;
      const isDebtConcept = concepto.includes('Mensualidad') || concepto.includes('Multa:'); 
      if (saldos[idUnidad] && isDebtConcept) {
          saldos[idUnidad].totalCargos += monto;
      }
  });

  // 2. Sumar Pagos
  pagos.forEach(row => {
      const idUnidad = String(row[2]).trim();
      const montoAplicadoNeto = Number(row[5]) || 0; 
      if (saldos[idUnidad]) {
          saldos[idUnidad].totalPagos += montoAplicadoNeto;
      }
  });

  // 3. Calcular Saldos, Filtrar Morosos y Ordenar (Ranking)
  let totalCarteraVencida = 0;
  const resultados = Object.values(saldos)
      .map(item => { 
          item.saldo = item.totalCargos - item.totalPagos; 
          return item; 
      })
      .filter(item => item.saldo > 0)
      .sort((a, b) => b.saldo - a.saldo);

  if (resultados.length === 0) {
      return { success: false, message: "¡Excelentes noticias! No hay deudas pendientes en el condominio." };
  }

  // Preparamos datos para la Web App y sumamos el gran total
  const datosParaFront = resultados.map((r, index) => {
      totalCarteraVencida += r.saldo;
      return [index + 1, r.nombre, r.saldo];
  });

  // ==============================================================================
  // CREACIÓN DE EXCEL (SOLO SI SE SOLICITÓ)
  // ==============================================================================
  if (crearExcel) {
      const nombreHoja = "REPORTE_DEUDORES_RANKING";
      let reporteSheet = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);
      reporteSheet.clear();

      const headers = ["RANKING", "ID_UNIDAD", "TOTAL_DEUDA_ACTUAL (Mensualidad/Multa)"];
      
      reporteSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#b91c1c").setFontColor("white");
      reporteSheet.getRange(2, 1, datosParaFront.length, headers.length).setValues(datosParaFront);
      reporteSheet.getRange(2, 3, datosParaFront.length, 1).setNumberFormat("$#,##0.00");
      reporteSheet.autoResizeColumns(1, headers.length);
  }

  return {
      success: true,
      totalDeudores: resultados.length,
      carteraVencida: totalCarteraVencida,
      ranking: datosParaFront
  };
}



// ==============================================================================
// MÓDULO: MENSUALIDADES VENCIDAS (ALERTA DE COBRANZA)
// ==============================================================================

function generarReporteVencidasWebApp(fechaConsultaStr, crearExcel) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
  
  // Asumimos que la función getConfig() existe en tu código original
  const config = typeof getConfig === "function" ? getConfig() : { DIA_LIMITE_NORMAL: 10 }; 
  const diaLimitePago = Number(config.DIA_LIMITE_NORMAL) || 10;

  if (!cargosSheet) return { success: false, message: "Error: Falta la hoja CARGOS_Y_DEUDAS." };

  // Convertir fecha string (YYYY-MM-DD) a objeto Date a las 00:00:00
  const parts = fechaConsultaStr.split('-');
  const fechaHoy = new Date(parts[0], parts[1] - 1, parts[2]);
  
  const cargosData = cargosSheet.getDataRange().getValues().slice(1);
  const resultadosVencidos = [];
  let montoTotalVencido = 0;

  cargosData.forEach(row => {
      const idUnidad = String(row[1]).trim();
      const concepto = row[2] ? row[2].toString() : '';
      const mesCorte = new Date(row[3]);
      const montoBase = Number(row[4]) || 0;
      const estado = row[5] ? row[5].toString().toUpperCase() : '';
      
      // Filtro estricto: Solo Pendientes y que sean Mensualidades
      if (estado !== "PENDIENTE" || !concepto.includes('Mensualidad')) return;
      if (!mesCorte || isNaN(mesCorte.getTime())) return;

      // Calcular la fecha límite real de ese cargo
      const fechaLimitePago = new Date(mesCorte.getFullYear(), mesCorte.getMonth(), diaLimitePago);
      
      // Si el día de consulta ya pasó la fecha límite, ¡A la lista negra!
      if (fechaHoy > fechaLimitePago) {
          resultadosVencidos.push([idUnidad, concepto, mesCorte, montoBase]);
          montoTotalVencido += montoBase;
      }
  });

  if (resultadosVencidos.length === 0) {
      return { success: false, message: "No se encontraron mensualidades vencidas a la fecha seleccionada." };
  }

  // Ordenar la lista: Primero por Unidad (alfabético) y luego por Fecha
  resultadosVencidos.sort((a, b) => {
      if (a[0] < b[0]) return -1;
      if (a[0] > b[0]) return 1;
      return a[2] - b[2];
  });

  // Preparar datos formateados para mostrar y exportar
  const datosFormateados = resultadosVencidos.map(r => [
      r[0], 
      r[1], 
      Utilities.formatDate(r[2], Session.getScriptTimeZone(), "MMMM yyyy").toUpperCase(), 
      r[3]
  ]);

  // ==============================================================================
  // CREAR PESTAÑA EXCEL (OPCIONAL)
  // ==============================================================================
  if (crearExcel) {
      const nombreHoja = "REPORTE_VENCIDOS_MENSUAL";
      let reporteSheet = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);
      reporteSheet.clear();

      reporteSheet.getRange('A1').setValue("REPORTE DE MENSUALIDADES VENCIDAS").setFontWeight("bold").setFontSize(14);
      reporteSheet.getRange('A2').setValue(`Fecha de Consulta: ${Utilities.formatDate(fechaHoy, Session.getScriptTimeZone(), "dd/MM/yyyy")}`).setFontWeight('bold');
      
      const headers = ["ID_UNIDAD", "CONCEPTO_MENSUALIDAD", "MES_CORTE", "MONTO_VENCIDO"];
      reporteSheet.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#f59e0b").setFontColor("white");
      
      reporteSheet.getRange(5, 1, datosFormateados.length, headers.length).setValues(datosFormateados);
      reporteSheet.getRange(5, 4, datosFormateados.length, 1).setNumberFormat("$#,##0.00");
      
      const tRow = datosFormateados.length + 5;
      reporteSheet.getRange(tRow, 3).setValue("TOTAL VENCIDO:").setFontWeight("bold");
      reporteSheet.getRange(tRow, 4).setValue(montoTotalVencido).setNumberFormat("$#,##0.00").setFontWeight("bold").setBackground("#fef3c7");

      reporteSheet.autoResizeColumns(1, headers.length);
  }

  return {
      success: true,
      fechaConsulta: Utilities.formatDate(fechaHoy, Session.getScriptTimeZone(), "dd/MM/yyyy"),
      diaLimiteConfigurado: diaLimitePago,
      totalRecibos: datosFormateados.length,
      montoTotal: montoTotalVencido,
      lista: datosFormateados
  };
}




/**
 * REPORTE: Deuda Real (Integrado para Web App y Sincronización Externa)
 * Mantiene tu lógica original intacta.
 */
function generarReporteDeudaRealWebApp(sincronizarExterno) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ID_ARCHIVO_EXTERNO = "1yKAExPx4FbqIEj10t-ZZZFgp_ZIsodt3a5xPxCzsx8Y"; 

  const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
  const pagosSheet = ss.getSheetByName("REGISTRO_PAGOS");
  const saldosAFavorSheet = ss.getSheetByName("SALDOS_A_FAVOR");

  if (!cargosSheet || !pagosSheet || !saldosAFavorSheet) {
    return { success: false, message: "Faltan hojas necesarias (CARGOS, PAGOS o SALDOS_A_FAVOR)." };
  }

  // 1. Cargar Saldos a Favor
  const anticipos = {};
  const anticiposData = saldosAFavorSheet.getDataRange().getValues().slice(1);
  anticiposData.forEach(row => {
    const id = String(row[0]); 
    const monto = Number(row[1]) || 0; 
    if (monto > 0) anticipos[id] = monto;
  });

  // 2. Lógica Fecha Actual
  const hoy = new Date();
  const mesActual = hoy.getMonth();
  const anioActual = hoy.getFullYear();

  // 3. Procesar Cargos (Tu lógica original de exclusión del mes actual)
  const cargosData = cargosSheet.getDataRange().getValues().slice(1);
  const deudasPorUnidad = {};

  cargosData.forEach(row => {
    const id = String(row[1]);
    const concepto = String(row[2]);
    const fechaCorte = new Date(row[3]);
    const monto = Number(row[4]) || 0;
    const estado = String(row[5]);

    let esMesActual = false;
    if (!isNaN(fechaCorte.getTime())) {
      esMesActual = (fechaCorte.getMonth() === mesActual && fechaCorte.getFullYear() === anioActual);
    }

    if (estado !== "Pagado" && !esMesActual) {
      if (!deudasPorUnidad[id]) deudasPorUnidad[id] = [];
      deudasPorUnidad[id].push({concepto: concepto, monto: monto});
    }
  });

  // 4. Construir Reporte (Tu matriz original)
  const filasReporte = [];
  const rankingReal = [];
  const desgloseParaPDF = []; // Solo para la pantalla
  let granTotalCondominio = 0;
  const unidades = Object.keys(deudasPorUnidad).sort();

  unidades.forEach(id => {
    let sumaCargos = 0;
    filasReporte.push([`Departamento ${id}`, ""]);
    
    const detallesUnidad = [];
    deudasPorUnidad[id].forEach(item => {
      filasReporte.push([item.concepto, item.monto]);
      sumaCargos += item.monto;
      detallesUnidad.push({concepto: item.concepto, monto: item.monto});
    });

    const saldoAFavor = anticipos[id] || 0;
    const totalReal = Math.max(0, sumaCargos - saldoAFavor);

    if (saldoAFavor > 0) {
      filasReporte.push(["Subtotal Cargos Pendientes", sumaCargos]);
      filasReporte.push(["(-) SALDO A FAVOR DISPONIBLE", -saldoAFavor]);
    }

    filasReporte.push([`TOTAL REAL A PAGAR ${id}`, totalReal]);
    filasReporte.push(["", ""]);

    if (totalReal > 0) {
      granTotalCondominio += totalReal;
      rankingReal.push({id: id, total: totalReal});
    }

    // Guardamos para el PDF
    desgloseParaPDF.push({
      id: id,
      subtotal: sumaCargos,
      anticipo: saldoAFavor,
      totalReal: totalReal,
      detalles: detallesUnidad
    });
  });

  filasReporte.push(["---------------------------------------", ""]);
  filasReporte.push(["GRAN TOTAL RECUPERABLE", granTotalCondominio]);
  filasReporte.push(["---------------------------------------", ""]);

  rankingReal.sort((a, b) => b.total - a.total);
  filasReporte.push(["", ""]);
  filasReporte.push(["RANKING DE DEUDORES (MAYOR A MENOR)", ""]);
  rankingReal.forEach(item => {
    filasReporte.push([`Depto ${item.id}`, item.total]);
  });

  // 5. Función de Escritura (Tu lógica original)
  const escribirEnHoja = (targetSS, nombre) => {
    let sheet = targetSS.getSheetByName(nombre) || targetSS.insertSheet(nombre);
    sheet.clear();
    sheet.getRange(1, 1, filasReporte.length, 2).setValues(filasReporte);
    sheet.setColumnWidth(1, 350);
    sheet.getRange(1, 2, filasReporte.length, 1).setNumberFormat("$#,##0.00");
    
    for (let i = 0; i < filasReporte.length; i++) {
      let t = String(filasReporte[i][0]);
      if (t.startsWith("Departamento")) sheet.getRange(i+1, 1, 1, 2).setFontWeight("bold").setBackground("#D9EAD3");
      if (t.startsWith("TOTAL REAL")) sheet.getRange(i+1, 1, 1, 2).setFontWeight("bold").setBackground("#FFF2CC");
      if (t.includes("(-) SALDO")) sheet.getRange(i+1, 1, 1, 2).setFontColor("green").setFontStyle("italic");
    }
  };

  // Ejecución
  escribirEnHoja(ss, "REPORTE_DEUDA_REAL");

  let msgExterno = "Sincronización no solicitada.";
  if (sincronizarExterno) {
    try {
      const ssExterno = SpreadsheetApp.openById(ID_ARCHIVO_EXTERNO);
      escribirEnHoja(ssExterno, "REPORTE_DEUDA_REAL");
      msgExterno = "Sincronización con Comité Exitosa.";
    } catch (e) {
      msgExterno = "Error en archivo externo: " + e.message;
    }
  }

  return {
    success: true,
    granTotal: granTotalCondominio,
    ranking: rankingReal,
    desglose: desgloseParaPDF,
    msgExterno: msgExterno
  };
}


