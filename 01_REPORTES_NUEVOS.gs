// ==============================================================================
// MÓDULO FINANCIERO MAESTRO (CON ESTILOS Y ESTRUCTURA ORIGINAL)
// ==============================================================================

function procesarPeticionFinancieroMaster(strInicio, strFin) {
  // Convertimos las fechas del HTML (Ej: "2026-03") a fechas de Apps Script
  const partesIni = strInicio.split("-");
  const partesFin = strFin.split("-");
  
  // Fecha inicio: Día 1 del mes de inicio
  const fechaInicio = new Date(partesIni[0], parseInt(partesIni[1]) - 1, 1);
  // Fecha fin: Último día del mes de fin
  const fechaFin = new Date(partesFin[0], parseInt(partesFin[1]), 0); 
  fechaFin.setHours(23, 59, 59, 999);

  const nombresMeses = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
  
  // Armamos tu título dinámico
  let periodoTexto = "";
  if (strInicio === strFin) {
    periodoTexto = `${nombresMeses[fechaInicio.getMonth()]} ${fechaInicio.getFullYear()}`;
  } else {
    periodoTexto = `${nombresMeses[fechaInicio.getMonth()]} ${fechaInicio.getFullYear()} a ${nombresMeses[fechaFin.getMonth()]} ${fechaFin.getFullYear()}`;
  }

  construirExcelFinancieroMaster(fechaInicio, fechaFin, periodoTexto);
}

function construirExcelFinancieroMaster(fechaInicio, fechaFin, periodoTexto) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
  const pagosSheet = ss.getSheetByName("REGISTRO_PAGOS");
  const egresosSheet = ss.getSheetByName("EGRESOS");

  if (!cargosSheet || !pagosSheet || !egresosSheet) return ui.alert("Error", "Faltan hojas.", ui.ButtonSet.OK);

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

  // PREPARAR HOJA
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
  reporteSheet.activate();
  
  ui.alert("Éxito", `El Reporte Financiero para ${periodoTexto} se generó con éxito.`, ui.ButtonSet.OK);
}