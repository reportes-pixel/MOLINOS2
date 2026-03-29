// ==============================================================================
// MODULO DE REPORTES Y ESTADOS DE CUENTA
// ==============================================================================

/**
 * 1. REPORTE: Saldo Detallado por Concepto y Fechas
 * Muestra cargos generados y cuánto se ha pagado de ellos en un periodo.
 */
function reporteSaldoDetalladoPorConcepto_Fechas() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    const pagosSheet = ss.getSheetByName("REGISTRO_PAGOS");

    if (!cargosSheet || !pagosSheet) return ui.alert("Error", "Faltan hojas.", ui.ButtonSet.OK);

    const prompt = ui.prompt("Reporte Detallado", "Rango de fechas (DD/MM/AAAA - DD/MM/AAAA):", ui.ButtonSet.OK_CANCEL);
    if (prompt.getSelectedButton() !== ui.Button.OK) return;

    const dates = prompt.getResponseText().split('-').map(s => s.trim());
    if (dates.length !== 2) return ui.alert("Error", "Formato incorrecto.", ui.ButtonSet.OK);

    const fechaInicio = new Date(dates[0].split('/').reverse().join('/'));
    const fechaFin = new Date(dates[1].split('/').reverse().join('/'));
    if (isNaN(fechaInicio) || isNaN(fechaFin)) return ui.alert("Error", "Fechas no válidas.", ui.ButtonSet.OK);
    
    fechaFin.setHours(23, 59, 59, 999); 

    const cargosData = cargosSheet.getDataRange().getValues().slice(1);
    const pagosData = pagosSheet.getDataRange().getValues().slice(1);
    const resultados = [];
    let totalGlobalCargos = 0, totalGlobalPagado = 0;

    const getPagoYConceptoDeCargo = (idCargo, fechaPagoMin, fechaPagoMax) => {
        let montoPagado = 0, conceptoPago = "PENDIENTE";
        pagosData.forEach(pagoRow => {
            const idCargoCubierto = String(pagoRow[8]);
            const fechaPago = new Date(pagoRow[1]);
            if (fechaPago >= fechaPagoMin && fechaPago <= fechaPagoMax && idCargoCubierto.includes(idCargo)) {
                montoPagado = Number(pagoRow[5]) || 0;
                conceptoPago = pagoRow[9] || "Pago sin Concepto";
            }
        });
        return { pagado: montoPagado, conceptoPago: conceptoPago };
    };
    
    cargosData.forEach(row => {
        const idCargo = String(row[0]), idUnidad = String(row[1]), conceptoCargo = String(row[2]);
        const mesCorte = new Date(row[3]);
        const montoBase = Number(row[4]) || 0;
        const estado = row[5] ? row[5].toString().toUpperCase() : '';

        if (conceptoCargo.includes('Recargo')) return;
        if (mesCorte >= fechaInicio && mesCorte <= fechaFin) {
            const pagoInfo = getPagoYConceptoDeCargo(idCargo, fechaInicio, fechaFin);
            let totalPagosAplicados = pagoInfo.pagado, conceptoPagado = pagoInfo.conceptoPago;
            let saldoPendiente = montoBase - totalPagosAplicados;

            if (estado === "PENDIENTE") {
                 totalPagosAplicados = 0; saldoPendiente = montoBase; conceptoPagado = "PENDIENTE";
            }
            if (montoBase > 0 || totalPagosAplicados > 0) {
                resultados.push([idUnidad, conceptoCargo, montoBase, totalPagosAplicados, conceptoPagado, saldoPendiente]);
                totalGlobalCargos += montoBase; totalGlobalPagado += totalPagosAplicados;
            }
        }
    });

    if (resultados.length === 0) return ui.alert("Aviso", "No hay datos para esas fechas.", ui.ButtonSet.OK);
    
    const nombreHoja = "REPORTE_SALDO_X_CONCEPTO";
    let reporteSheet = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);
    reporteSheet.clearContents();

    reporteSheet.getRange('A1').setValue("REPORTE DETALLADO POR CONCEPTO Y UNIDAD");
    reporteSheet.getRange('A2').setValue(`Periodo: ${dates[0]} al ${dates[1]}`).setFontWeight('bold');
    
    const headers = ["ID_UNIDAD", "CONCEPTO_CARGO", "TOTAL_CARGOS", "TOTAL_PAGOS_APLICADOS", "CONCEPTO_PAGADO", "SALDO_PENDIENTE"];
    reporteSheet.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    reporteSheet.getRange(5, 1, resultados.length, resultados[0].length).setValues(resultados);
    reporteSheet.getRange(5, 3, resultados.length, 4).setNumberFormat("$#,##0.00");
    
    const tRow = resultados.length + 5;
    reporteSheet.getRange(tRow, 1).setValue("TOTAL GLOBAL:").setFontWeight("bold");
    reporteSheet.getRange(tRow, 3).setValue(totalGlobalCargos).setNumberFormat("$#,##0.00").setFontWeight("bold");
    reporteSheet.getRange(tRow, 4).setValue(totalGlobalPagado).setNumberFormat("$#,##0.00").setFontWeight("bold");
    reporteSheet.getRange(tRow, 6).setValue(totalGlobalCargos - totalGlobalPagado).setNumberFormat("$#,##0.00").setFontWeight("bold").setBackground("#FFE599");

    reporteSheet.autoResizeColumns(1, headers.length);
    reporteSheet.activate();
}


/**
 * 2. REPORTE: Total Deudores (Ranking)
 * Lista los deudores ordenados por saldo pendiente (Mensualidades y Multas).
 */
function reporteTotalDeudores() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    const pagosSheet = ss.getSheetByName("REGISTRO_PAGOS");
    const unidadesSheet = ss.getSheetByName("UNIDADES");

    if (!cargosSheet || !pagosSheet || !unidadesSheet) return ui.alert("Error", "Faltan hojas.", ui.ButtonSet.OK);

    const units = unidadesSheet.getRange(2, 1, unidadesSheet.getLastRow() - 1, 1).getValues().flat().map(String);
    const cargos = cargosSheet.getDataRange().getValues().slice(1);
    const pagos = pagosSheet.getDataRange().getValues().slice(1);

    const saldos = units.reduce((acc, id) => { acc[id] = { totalCargos: 0, totalPagos: 0, saldo: 0, nombre: id }; return acc; }, {});

    cargos.forEach(row => {
        const idUnidad = String(row[1]), concepto = row[2] ? row[2].toString() : '', monto = Number(row[4]) || 0;
        const isDebtConcept = concepto.includes('Mensualidad') || concepto.includes('Multa:'); 
        if (saldos[idUnidad] && isDebtConcept) saldos[idUnidad].totalCargos += monto;
    });

    pagos.forEach(row => {
        const idUnidad = String(row[2]), montoAplicadoNeto = Number(row[5]) || 0; 
        if (saldos[idUnidad]) saldos[idUnidad].totalPagos += montoAplicadoNeto;
    });

    const resultados = Object.values(saldos)
        .map(item => { item.saldo = item.totalCargos - item.totalPagos; return item; })
        .filter(item => item.saldo > 0)
        .sort((a, b) => b.saldo - a.saldo);

    if (resultados.length === 0) return ui.alert("Aviso", "No hay deudas pendientes.", ui.ButtonSet.OK);

    const nombreHoja = "REPORTE_DEUDORES_RANKING";
    let reporteSheet = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);
    reporteSheet.clearContents();

    const headers = ["RANKING", "ID_UNIDAD", "TOTAL_DEUDA_ACTUAL (Mensualidad/Multa)"];
    const datos = resultados.map((r, index) => [index + 1, r.nombre, r.saldo]);

    reporteSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    reporteSheet.getRange(2, 1, datos.length, datos[0].length).setValues(datos);
    reporteSheet.getRange(2, 3, datos.length, 1).setNumberFormat("$#,##0.00");
    reporteSheet.autoResizeColumns(1, headers.length);
    reporteSheet.activate();
}


/**
 * 3. REPORTE: Mensualidades Vencidas
 * Lista mensualidades pendientes que ya pasaron su fecha límite de pago.
 */
function reporteMensualidadesVencidas_CORREGIDO() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    const config = getConfig(); 

    if (!cargosSheet || !config) return ui.alert("Error", "Faltan hojas.", ui.ButtonSet.OK);

    const diaLimitePago = Number(config.DIA_LIMITE_NORMAL) || 10;
    const prompt = ui.prompt("Vencidas", `Día límite es el ${diaLimitePago}.\nIntroduce Fecha de Consulta (DD/MM/AAAA):`, ui.ButtonSet.OK_CANCEL);
    if (prompt.getSelectedButton() !== ui.Button.OK) return;
    
    const parts = prompt.getResponseText().split('/');
    if (parts.length !== 3) return ui.alert("Error", "Formato inválido.", ui.ButtonSet.OK);
    
    const fechaConsulta = new Date(parts[2], parts[1] - 1, parts[0]);
    if (isNaN(fechaConsulta)) return ui.alert("Error", "Fecha no válida.", ui.ButtonSet.OK);
    
    const fechaHoy = new Date(fechaConsulta.getFullYear(), fechaConsulta.getMonth(), fechaConsulta.getDate());
    const lastRow = cargosSheet.getLastRow();
    if (lastRow < 2) return;
    
    const cargosData = cargosSheet.getRange(2, 2, lastRow - 1, 5).getValues(); 
    const resultadosVencidos = [];
    let montoTotalVencido = 0;

    cargosData.forEach(row => {
        const idUnidad = String(row[0]), concepto = row[1] ? row[1].toString() : '';
        const mesCorte = new Date(row[2]);
        const montoBase = Number(row[3]) || 0;
        const estado = row[4] ? row[4].toString().toUpperCase() : '';
        
        if (estado !== "PENDIENTE" || !concepto.includes('Mensualidad')) return;
        if (!mesCorte || isNaN(mesCorte.getTime())) return;

        const fechaLimitePago = new Date(mesCorte.getFullYear(), mesCorte.getMonth(), diaLimitePago);
        if (fechaHoy > fechaLimitePago) {
            resultadosVencidos.push([idUnidad, concepto, mesCorte.toLocaleDateString('es-ES'), montoBase]);
            montoTotalVencido += montoBase;
        }
    });

    const nombreHoja = "REPORTE_VENCIDOS_MENSUAL";
    let reporteSheet = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);
    reporteSheet.clearContents();

    reporteSheet.getRange('A1').setValue("REPORTE DE MENSUALIDADES VENCIDAS");
    reporteSheet.getRange('A2').setValue(`Fecha Consulta: ${fechaConsulta.toLocaleDateString()}`).setFontWeight('bold');
    
    if (resultadosVencidos.length === 0) return ui.alert("Aviso", "No hay mensualidades vencidas.", ui.ButtonSet.OK);
    
    const headers = ["ID_UNIDAD", "CONCEPTO_MENSUALIDAD", "MES_CORTE", "MONTO_VENCIDO"];
    reporteSheet.getRange(4, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#F4CCCC");
    reporteSheet.getRange(5, 1, resultadosVencidos.length, resultadosVencidos[0].length).setValues(resultadosVencidos);
    reporteSheet.getRange(5, 4, resultadosVencidos.length, 1).setNumberFormat("$#,##0.00");
    
    const tRow = resultadosVencidos.length + 5;
    reporteSheet.getRange(tRow, 3).setValue("TOTAL VENCIDO:").setFontWeight("bold");
    reporteSheet.getRange(tRow, 4).setValue(montoTotalVencido).setNumberFormat("$#,##0.00").setFontWeight("bold").setBackground("#E06666");

    reporteSheet.autoResizeColumns(1, headers.length);
    reporteSheet.activate();
}


/**
 * 4. REPORTE: Estado de Cuenta SEPARADO
 * Muestra el historial cronológico de un departamento.
 */
function reporteEstadoDeCuenta_SEPARADO() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    const pagosSheet = ss.getSheetByName("REGISTRO_PAGOS");
    
    if (!cargosSheet || !pagosSheet) return ui.alert("Error", "Faltan hojas.", ui.ButtonSet.OK);
    
    const prompt = ui.prompt("Estado de Cuenta", "ID Unidad (ej: 17A) o 'TODOS':", ui.ButtonSet.OK_CANCEL);
    if (prompt.getSelectedButton() !== ui.Button.OK) return;
    const targetUnit = prompt.getResponseText().trim().toUpperCase();

    const cargos = cargosSheet.getDataRange().getValues().slice(1);
    const pagos = pagosSheet.getDataRange().getValues().slice(1);
    let transacciones = [];
    
    cargos.forEach(row => {
        const idUnidad = String(row[1]);
        if (targetUnit !== 'TODOS' && idUnidad !== targetUnit) return;
        transacciones.push({
            unidad: idUnidad, fecha: new Date(row[3]), concepto: String(row[2]),
            estado: String(row[5] || "PENDIENTE").toUpperCase(), cargo: Number(row[4]) || 0,
            pago: 0, tipo: 'CARGO', id: row[0]
        });
    });
    
    pagos.forEach(row => {
        const idUnidad = String(row[2]); 
        if (targetUnit !== 'TODOS' && idUnidad !== targetUnit) return;
        const montoAplicado = Number(row[5]) || 0; 
        if (montoAplicado === 0) return; 

        transacciones.push({
            unidad: idUnidad, fecha: new Date(row[1]), concepto: `PAGO: ${String(row[9] || "Sin concepto")}`,
            estado: "PAGO", cargo: 0, pago: montoAplicado, tipo: 'PAGO', id: row[0]
        });
    });

    if (transacciones.length === 0) return ui.alert("Aviso", "No hay transacciones.", ui.ButtonSet.OK);
    
    transacciones.sort((a, b) => a.fecha - b.fecha);
    
    const nombreHoja = `EDO_CTA_${targetUnit}`;
    let reporteSheet = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);
    reporteSheet.clearContents();
    
    const headers = ["UNIDAD", "FECHA", "CONCEPTO", "ESTADO", "CARGO", "PAGO", "SALDO_ACUMULADO"];
    let saldoAcumulado = 0; const datosFinales = [];
    
    transacciones.forEach(t => {
        saldoAcumulado = saldoAcumulado + t.cargo - t.pago; 
        datosFinales.push([t.unidad, t.fecha, t.concepto, t.estado, t.cargo, t.pago, saldoAcumulado]);
    });

    reporteSheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold");
    reporteSheet.getRange(2, 1, datosFinales.length, datosFinales[0].length).setValues(datosFinales);
    reporteSheet.getRange(2, 5, datosFinales.length, 3).setNumberFormat("$#,##0.00");
    reporteSheet.getRange(2, 2, datosFinales.length, 1).setNumberFormat("dd/MM/yyyy");
    reporteSheet.autoResizeColumns(1, headers.length);
    reporteSheet.activate();
}


/**
 * 5. REPORTE: Financiero Resumido
 * ⭐️ CORREGIDO PARA EGRESOS (Monto en índice 5)
 */
function generarReporteFinanciero() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    const pagosSheet = ss.getSheetByName("REGISTRO_PAGOS");
    const egresosSheet = ss.getSheetByName("EGRESOS");

    if (!cargosSheet || !pagosSheet || !egresosSheet) return ui.alert("Error", "Faltan hojas.", ui.ButtonSet.OK);

    const prompt = ui.prompt("Reporte Financiero", "Fechas (DD/MM/AAAA - DD/MM/AAAA):", ui.ButtonSet.OK_CANCEL);
    if (prompt.getSelectedButton() !== ui.Button.OK) return;

    const dates = prompt.getResponseText().split('-').map(s => s.trim());
    const fechaInicio = new Date(dates[0].split('/').reverse().join('/'));
    const fechaFin = new Date(dates[1].split('/').reverse().join('/'));
    fechaFin.setHours(23, 59, 59, 999);

    let totalCargosGenerados = 0, totalCargosPagados = 0, totalMontoRecibido = 0, totalMontoEgreso = 0;

    const cargosData = cargosSheet.getDataRange().getValues().slice(1);
    cargosData.forEach(row => {
        const mesCorte = new Date(row[3]), monto = Number(row[4]) || 0, estado = String(row[5]).toUpperCase();
        if (mesCorte >= fechaInicio && mesCorte <= fechaFin) {
            totalCargosGenerados += monto;
            if (estado === "PAGADO") totalCargosPagados += monto;
        }
    });

    const pagosData = pagosSheet.getDataRange().getValues().slice(1);
    pagosData.forEach(row => {
        const fechaPago = new Date(row[1]), montoRecibido = Number(row[4]) || 0;
        if (fechaPago >= fechaInicio && fechaPago <= fechaFin) totalMontoRecibido += montoRecibido;
    });

    const egresosData = egresosSheet.getDataRange().getValues().slice(1);
    egresosData.forEach(row => {
        const fechaRegistro = new Date(row[1]);
        const montoEgreso = Number(row[5]) || 0; // ⭐️ CORRECCIÓN EGRESOS: Col F
        if (fechaRegistro >= fechaInicio && fechaRegistro <= fechaFin) totalMontoEgreso += montoEgreso;
    });

    const nombreHoja = "REPORTE_FINANCIERO_RESUMEN";
    let reporteSheet = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);
    reporteSheet.clearContents();

    const datos = [
        ["--- DINERO ESPERADO (CARGOS) ---", ""],
        ["Total de Cargos Generados", totalCargosGenerados],
        ["Total de Cargos Pagados", totalCargosPagados],
        ["Total de Cargos PENDIENTES", totalCargosGenerados - totalCargosPagados],
        ["", ""],
        ["--- DINERO COBRADO (PAGOS) ---", ""],
        ["Total Recibido en Pagos", totalMontoRecibido],
        ["", ""],
        ["--- DINERO GASTADO (EGRESOS) ---", ""],
        ["Total Egresos Registrados", totalMontoEgreso],
        ["", ""],
        ["*===================================*", ""],
        ["SALDO NETO DE EFECTIVO", totalMontoRecibido - totalMontoEgreso]
    ];

    reporteSheet.getRange('A1').setValue("REPORTE FINANCIERO RESUMIDO").setFontWeight('bold');
    reporteSheet.getRange('A2').setValue(`Periodo: ${dates[0]} al ${dates[1]}`).setFontWeight('bold');
    reporteSheet.getRange(4, 1, datos.length, 2).setValues(datos);
    
    reporteSheet.getRange('A4').setFontWeight('bold').setBackground("#D9EAD3");
    reporteSheet.getRange('A8').setFontWeight('bold').setBackground("#B4C6E7");
    reporteSheet.getRange('A11').setFontWeight('bold').setBackground("#F4CCCC");
    reporteSheet.getRange('A14:B14').setFontWeight('bold').setBackground("#FFE599");
    reporteSheet.getRange(5, 2, 9, 1).setNumberFormat("$#,##0.00");
    reporteSheet.autoResizeColumns(1, 2);
    reporteSheet.activate();
}


/**
 * 6. REPORTE: Financiero Extendido
 * ⭐️ CORREGIDO PARA EGRESOS (Monto en índice 5, Concepto en índice 3)
 */
function generarReporteFinanciero_EXTENDIDO() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    const pagosSheet = ss.getSheetByName("REGISTRO_PAGOS");
    const egresosSheet = ss.getSheetByName("EGRESOS");

    if (!cargosSheet || !pagosSheet || !egresosSheet) return ui.alert("Error", "Faltan hojas.", ui.ButtonSet.OK);

    const prompt = ui.prompt("Reporte Extendido", "Fechas (DD/MM/AAAA - DD/MM/AAAA):", ui.ButtonSet.OK_CANCEL);
    if (prompt.getSelectedButton() !== ui.Button.OK) return;

    const dates = prompt.getResponseText().split('-').map(s => s.trim());
    const fechaInicio = new Date(dates[0].split('/').reverse().join('/'));
    const fechaFin = new Date(dates[1].split('/').reverse().join('/'));
    fechaFin.setHours(23, 59, 59, 999);

    let totalCargosGenerados = 0, totalCargosPagados = 0, totalMontoRecibido = 0, totalMontoEgreso = 0;
    const pagosPorConcepto = {}; const egresosPorCategoria = {};

    cargosSheet.getDataRange().getValues().slice(1).forEach(row => {
        const mesCorte = new Date(row[3]), monto = Number(row[4]) || 0;
        if (mesCorte >= fechaInicio && mesCorte <= fechaFin) {
            totalCargosGenerados += monto;
            if (String(row[5]).toUpperCase() === "PAGADO") totalCargosPagados += monto;
        }
    });

    pagosSheet.getDataRange().getValues().slice(1).forEach(row => {
        const fechaPago = new Date(row[1]), montoRecibido = Number(row[4]) || 0, concepto = String(row[9] || "Sin Concepto");
        if (fechaPago >= fechaInicio && fechaPago <= fechaFin) {
            totalMontoRecibido += montoRecibido;
            pagosPorConcepto[concepto] = (pagosPorConcepto[concepto] || 0) + montoRecibido;
        }
    });

    egresosSheet.getDataRange().getValues().slice(1).forEach(row => {
        const fechaRegistro = new Date(row[1]);
        const categoria = String(row[3] || "Sin Categoría").trim(); // ⭐️ CORRECCIÓN EGRESOS: Concepto Col D
        const montoEgreso = Number(row[5]) || 0; // ⭐️ CORRECCIÓN EGRESOS: Monto Col F

        if (fechaRegistro >= fechaInicio && fechaRegistro <= fechaFin) {
            totalMontoEgreso += montoEgreso;
            egresosPorCategoria[categoria] = (egresosPorCategoria[categoria] || 0) + montoEgreso;
        }
    });

    const nombreHoja = "REPORTE_FINANCIERO_EXTENDIDO";
    let reporteSheet = ss.getSheetByName(nombreHoja) || ss.insertSheet(nombreHoja);
    reporteSheet.clearContents();

    let row = 1;
    reporteSheet.getRange(`A${row}`).setValue("REPORTE FINANCIERO EXTENDIDO").setFontWeight('bold');
    reporteSheet.getRange(`A${row+1}`).setValue(`Periodo: ${dates[0]} al ${dates[1]}`).setFontWeight('bold');
    row += 3;

    const datosResumen = [
        ["--- RESUMEN DE CARGOS Y SALDOS DEVENGADOS ---", ""],
        ["Total Cargos Generados", totalCargosGenerados],
        ["Total Cargos Pagados", totalCargosPagados],
        ["Total Cargos PENDIENTES", totalCargosGenerados - totalCargosPagados],
        ["", ""],
        ["--- SALDO NETO DE EFECTIVO (Flujo de Caja) ---", ""],
        ["Total Recibido en Pagos (INGRESOS)", totalMontoRecibido],
        ["Total Egresos Registrados (GASTOS)", totalMontoEgreso],
        ["SALDO NETO DE EFECTIVO", totalMontoRecibido - totalMontoEgreso]
    ];
    
    reporteSheet.getRange(row, 1, datosResumen.length, 2).setValues(datosResumen);
    reporteSheet.getRange(`A${row}`).setFontWeight('bold').setBackground("#D9EAD3");
    reporteSheet.getRange(`A${row+5}`).setFontWeight('bold').setBackground("#B4C6E7");
    reporteSheet.getRange(`A${row+8}:B${row+8}`).setFontWeight('bold').setBackground("#FFE599");
    reporteSheet.getRange(row + 1, 2, 8, 1).setNumberFormat("$#,##0.00");
    row += datosResumen.length + 1;

    // Detalle Pagos
    reporteSheet.getRange(`A${row}`).setValue("DETALLE DE INGRESOS").setFontWeight('bold').setBackground("#B6D7A8"); row++;
    const detallePagos = [["CONCEPTO DE PAGO", "MONTO TOTAL"]];
    Object.entries(pagosPorConcepto).forEach(([k, v]) => detallePagos.push([k, v]));
    if (detallePagos.length > 1) {
        reporteSheet.getRange(row, 1, detallePagos.length, 2).setValues(detallePagos);
        reporteSheet.getRange(row, 2, detallePagos.length, 1).setNumberFormat("$#,##0.00");
        reporteSheet.getRange(row, 1, 1, 2).setFontWeight('bold');
    }
    row += detallePagos.length + 1;

    // Detalle Egresos
    reporteSheet.getRange(`A${row}`).setValue("DETALLE DE EGRESOS").setFontWeight('bold').setBackground("#F4CCCC"); row++;
    const detalleEgresos = [["CATEGORÍA", "MONTO TOTAL"]];
    Object.entries(egresosPorCategoria).forEach(([k, v]) => detalleEgresos.push([k, v]));
    if (detalleEgresos.length > 1) {
        reporteSheet.getRange(row, 1, detalleEgresos.length, 2).setValues(detalleEgresos);
        reporteSheet.getRange(row, 2, detalleEgresos.length, 1).setNumberFormat("$#,##0.00");
        reporteSheet.getRange(row, 1, 1, 2).setFontWeight('bold');
    }

    reporteSheet.autoResizeColumns(1, 2);
    reporteSheet.activate();
}


/**
 * 7. REPORTE: Deuda Real (Incluye Saldos a favor cruzados)
 */
 function reporteDeudaRealCorregido() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  // --- CONFIGURACIÓN ARCHIVO EXTERNO ---
  const ID_ARCHIVO_EXTERNO = "1yKAExPx4FbqIEj10t-ZZZFgp_ZIsodt3a5xPxCzsx8Y"; 
  // -------------------------------------

  const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
  const pagosSheet = ss.getSheetByName("REGISTRO_PAGOS");
  const saldosAFavorSheet = ss.getSheetByName("SALDOS_A_FAVOR");

  if (!cargosSheet || !pagosSheet || !saldosAFavorSheet) {
    ui.alert("Error", "Faltan hojas necesarias en el archivo origen.", ui.ButtonSet.OK);
    return;
  }

  // 1. Cargar Saldos a Favor
  const anticipos = {};
  const anticiposData = saldosAFavorSheet.getDataRange().getValues().slice(1);
  
  anticiposData.forEach(row => {
    const id = String(row[0]); 
    const monto = Number(row[1]) || 0; 
    if (monto > 0) anticipos[id] = monto;
  });

  // 2. Cargar Pagos realizados (Lectura original)
  const pagosData = pagosSheet.getDataRange().getValues().slice(1);

  // --- LÓGICA: FECHA ACTUAL ---
  const hoy = new Date();
  const mesActual = hoy.getMonth();
  const anioActual = hoy.getFullYear();
  // ----------------------------

  // 3. Procesar Cargos (CON LOS ÍNDICES CORRECTOS)
  const cargosData = cargosSheet.getDataRange().getValues().slice(1);
  const deudasPorUnidad = {};

  cargosData.forEach(row => {
    const id = String(row[1]);         // Columna B: ID_UNIDAD
    const concepto = String(row[2]);   // Columna C: CONCEPTO
    const fechaCorte = new Date(row[3]); // Columna D: MES_CORTE
    const monto = Number(row[4]) || 0; // Columna E: MONTO_BASE
    const estado = String(row[5]);     // Columna F: ESTADO

    let esMesActual = false;
    
    // Validamos que fechaCorte sea una fecha real antes de comprobar
    if (!isNaN(fechaCorte.getTime())) {
      esMesActual = (fechaCorte.getMonth() === mesActual && fechaCorte.getFullYear() === anioActual);
    }

    // Solo se suma si NO está pagado y NO corresponde al mes y año actual
    if (estado !== "Pagado" && !esMesActual) {
      if (!deudasPorUnidad[id]) deudasPorUnidad[id] = [];
      deudasPorUnidad[id].push({concepto: concepto, monto: monto});
    }
  });

  // 4. Construir Reporte (Matriz de datos)
  const filasReporte = [];
  const rankingReal = [];
  let granTotalCondominio = 0;
  const unidades = Object.keys(deudasPorUnidad).sort();

  unidades.forEach(id => {
    let sumaCargos = 0;
    filasReporte.push([`Departamento ${id}`, ""]);
    
    deudasPorUnidad[id].forEach(item => {
      filasReporte.push([item.concepto, item.monto]);
      sumaCargos += item.monto;
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
  });

  filasReporte.push(["---------------------------------------", ""]);
  filasReporte.push(["GRAN TOTAL RECUPERABLE", granTotalCondominio]);
  filasReporte.push(["---------------------------------------", ""]);
  filasReporte.push(["", ""]);

  rankingReal.sort((a, b) => b.total - a.total);
  filasReporte.push(["RANKING DE DEUDORES (MAYOR A MENOR)", ""]);
  rankingReal.forEach(item => {
    filasReporte.push([`Depto ${item.id}`, item.total]);
  });

  // 5. FUNCIÓN INTERNA PARA ESCRIBIR Y DAR FORMATO (INTACTA)
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

  // 6. Ejecutar escritura en ambos archivos
  const nombreHojaReporte = "REPORTE_DEUDA_REAL";
  
  // Actualizar en este archivo
  escribirEnHoja(ss, nombreHojaReporte);

  // Actualizar en el archivo externo
  try {
    const ssExterno = SpreadsheetApp.openById(ID_ARCHIVO_EXTERNO);
    escribirEnHoja(ssExterno, nombreHojaReporte);
    ui.alert("Éxito", "Reporte actualizado en ambos archivos.", ui.ButtonSet.OK);
  } catch (e) {
    ui.alert("Error de acceso", "No se pudo actualizar el archivo externo. Revisa el ID y los permisos. " + e.message, ui.ButtonSet.OK);
  }
}


/**
 * ==============================================================================
 * HERRAMIENTAS DE CORTE Y TICKETS
 * ==============================================================================
 */

function generarCortePagos() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRegistro = ss.getSheetByName("REGISTRO_PAGOS");
  const ui = SpreadsheetApp.getUi();
  
  const respuesta = ui.prompt("Corte", "Fecha (DD/MM/YYYY):", ui.ButtonSet.OK_CANCEL);
  if (respuesta.getSelectedButton() !== ui.Button.OK) return;

  const datos = hojaRegistro.getDataRange().getValues();
  let resultados = [datos[0]]; let totalMonto = 0;

  for (let i = 1; i < datos.length; i++) {
    let fechaCelda = Utilities.formatDate(new Date(datos[i][1]), Session.getScriptTimeZone(), "dd/MM/yyyy");
    if (fechaCelda === respuesta.getResponseText()) {
      resultados.push(datos[i]); totalMonto += parseFloat(datos[i][4]) || 0;
    }
  }

  if (resultados.length <= 1) return ui.alert("No hay pagos para esa fecha.");

  let hojaCorte = ss.getSheetByName("Corte_del_Día") || ss.insertSheet("Corte_del_Día");
  hojaCorte.clear();
  hojaCorte.getRange(1, 1, resultados.length, resultados[0].length).setValues(resultados);
  hojaCorte.getRange(resultados.length + 2, 4).setValue("TOTAL RECAUDADO:");
  hojaCorte.getRange(resultados.length + 2, 5).setValue(totalMonto).setFontWeight("bold").setNumberFormat("$#,##0.00");
  hojaCorte.activate();
}

function generarCorteVariosDias() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaRegistro = ss.getSheetByName("REGISTRO_PAGOS");
  const ui = SpreadsheetApp.getUi();
  
  const rInicio = ui.prompt("Corte Rango", "INICIO (DD/MM/YYYY):", ui.ButtonSet.OK_CANCEL);
  if (rInicio.getSelectedButton() !== ui.Button.OK) return;
  const rFin = ui.prompt("Corte Rango", "FIN (DD/MM/YYYY):", ui.ButtonSet.OK_CANCEL);
  if (rFin.getSelectedButton() !== ui.Button.OK) return;

  const fInicio = parseFecha(rInicio.getResponseText()), fFin = parseFecha(rFin.getResponseText());
  if (!fInicio || !fFin) return ui.alert("Formato inválido.");

  const datos = hojaRegistro.getDataRange().getValues();
  let resultados = [datos[0]]; let totalMonto = 0;

  for (let i = 1; i < datos.length; i++) {
    let d = new Date(datos[i][1]);
    if (d >= fInicio && d <= fFin) { resultados.push(datos[i]); totalMonto += parseFloat(datos[i][4]) || 0; }
  }

  if (resultados.length <= 1) return ui.alert("No hay pagos.");

  let hojaCorte = ss.getSheetByName("Corte_Rango") || ss.insertSheet("Corte_Rango");
  hojaCorte.clear();
  hojaCorte.getRange(1, 1, resultados.length, resultados[0].length).setValues(resultados);
  hojaCorte.getRange(resultados.length + 2, 4).setValue("TOTAL:");
  hojaCorte.getRange(resultados.length + 2, 5).setValue(totalMonto).setFontWeight("bold").setNumberFormat("$#,##0.00");
  hojaCorte.activate();
}

function parseFecha(str) {
  const partes = str.split('/');
  return partes.length === 3 ? new Date(partes[2], partes[1] - 1, partes[0]) : null;
}

function lanzarBuscadorTicket() {
  const ui = SpreadsheetApp.getUi();
  const r = ui.prompt('Reimpresión', 'ID Pago (ej. RGP-...):', ui.ButtonSet.OK_CANCEL);
  if (r.getSelectedButton() == ui.Button.OK) showTicketDialog(r.getResponseText().toUpperCase().trim());
}

function obtenerDatosCorteWebApp(fechaBusqueda, capturistaFiltro) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("REGISTRO_PAGOS");
  const data = sheet.getDataRange().getValues();
  const resultados = [];
  let totalMonto = 0;

  // Convertir fecha de YYYY-MM-DD a DD/MM/YYYY para comparar con la BD
  const partes = fechaBusqueda.split('-');
  const fechaFormateada = `${partes[2]}/${partes[1]}/${partes[0]}`;

  for (let i = 1; i < data.length; i++) {
    // Evitamos filas vacías
    if (!data[i][1]) continue; 
    
    let fechaCelda = Utilities.formatDate(new Date(data[i][1]), Session.getScriptTimeZone(), "dd/MM/yyyy");
    let capturistaCelda = data[i][3];
    
    if (fechaCelda === fechaFormateada) {
      // Si el filtro es MIO, solo trae los del usuario activo
      if (capturistaFiltro && capturistaFiltro !== "TODOS" && capturistaCelda !== capturistaFiltro) continue;
      
      resultados.push({
        folio: data[i][0],
        unidad: data[i][2],
        capturista: capturistaCelda,
        monto: parseFloat(data[i][4]) || 0,
        concepto: data[i][9] || "Pago registrado"
      });
      totalMonto += parseFloat(data[i][4]) || 0;
    }
  }

  return {
    success: true,
    fecha: fechaFormateada,
    pagos: resultados,
    total: totalMonto
  };
}

