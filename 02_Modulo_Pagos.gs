// ==============================================================================
// 02_MODULO_PAGOS.gs - CASCADA OBLIGATORIA (FIFO), ABONOS PARCIALES Y CANCELACIONES
// ==============================================================================

/**
 * Procesa el pago aplicando el dinero OBLIGATORIAMENTE a la deuda más antigua primero.
 * Si el dinero no alcanza para cubrirla, genera un "Saldo Restante" (Abono).
 */
function processPayment(paymentData) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPagos = ss.getSheetByName("REGISTRO_PAGOS");
    const sheetCargos = ss.getSheetByName("CARGOS_Y_DEUDAS");
    
    if (!sheetPagos || !sheetCargos) {
        return { success: false, message: "Error: Faltan hojas 'REGISTRO_PAGOS' o 'CARGOS_Y_DEUDAS'." };
    }
    
    let { idUnidad, montoRecibido, fechaPagoStr, capturista, deudasAplicadas } = paymentData;
    montoRecibido = parseFloat(montoRecibido);
    const fechaPago = new Date(fechaPagoStr); 
    fechaPago.setMinutes(fechaPago.getMinutes() + fechaPago.getTimezoneOffset());
    
    // 1. OBTENER SALDO A FAVOR PREVIO (La Bolsa de Dinero Total)
    let saldoAFavorInicial = getUnitAnticipo(idUnidad); 
    let bolsaDeDinero = montoRecibido + saldoAFavorInicial; 
    
    let montoTotalAplicadoADeuda_Neto = 0; 
    let conceptosCubiertosIds = [];
    
    // 2. BUSCAR Y ORDENAR TODAS LAS DEUDAS DE LA UNIDAD (FIFO)
    const cargosData = sheetCargos.getDataRange().getValues();
    const idCargoCol = 0;   // A
    const conceptoCol = 2;  // C
    const fechaCorteCol = 3;// D
    const montoBaseCol = 4; // E
    const estadoCol = 5;    // F

    let deudasAProcesar = [];

    deudasAplicadas.forEach(deudaHtml => {
        for (let i = 1; i < cargosData.length; i++) {
            if (cargosData[i][idCargoCol] === deudaHtml.idCargo && String(cargosData[i][estadoCol]).toUpperCase() === "PENDIENTE") {
                deudasAProcesar.push({
                    rowIndex: i + 1, 
                    idCargo: cargosData[i][idCargoCol],
                    concepto: cargosData[i][conceptoCol],
                    montoBase: Number(cargosData[i][montoBaseCol]),
                    fechaCorteDate: new Date(cargosData[i][fechaCorteCol]),
                    filaCompleta: cargosData[i] 
                });
                break;
            }
        }
    });

    // Ordenar estrictamente de la más antigua a la más nueva
    deudasAProcesar.sort((a, b) => a.fechaCorteDate.getTime() - b.fechaCorteDate.getTime());

    let remanentesAGenerar = []; 

    // 3. LA CASCADA (VACIAR LA BOLSA DE DINERO EN ORDEN)
    for (let d of deudasAProcesar) {
        if (bolsaDeDinero <= 0.001) break; 

        if (bolsaDeDinero >= d.montoBase) {
            // PAGO COMPLETO
            sheetCargos.getRange(d.rowIndex, estadoCol + 1).setValue("Pagado");
            
            bolsaDeDinero -= d.montoBase;
            montoTotalAplicadoADeuda_Neto += d.montoBase;
            conceptosCubiertosIds.push(d.idCargo);
            
        } else {
            // ABONO PARCIAL OBLIGATORIO
            let abono = bolsaDeDinero;
            let saldoRestante = d.montoBase - abono;
            
            sheetCargos.getRange(d.rowIndex, montoBaseCol + 1).setValue(abono);
            sheetCargos.getRange(d.rowIndex, estadoCol + 1).setValue("Pagado");
            
            const nuevoIdRemanente = 'CGO-' + Utilities.getUuid().substring(0, 8).toUpperCase();
            let filaClonada = [...d.filaCompleta];
            
            filaClonada[idCargoCol] = nuevoIdRemanente;
            filaClonada[conceptoCol] = `Saldo Restante: ${d.concepto}`; 
            filaClonada[montoBaseCol] = saldoRestante; 
            filaClonada[estadoCol] = "Pendiente";
            filaClonada[6] = ""; 
            filaClonada[7] = ""; 
            
            remanentesAGenerar.push(filaClonada);
            
            montoTotalAplicadoADeuda_Neto += abono;
            conceptosCubiertosIds.push(d.idCargo);
            bolsaDeDinero = 0; 
        }
    }

    // 4. INYECTAR LAS DEUDAS PARTIDAS 
    if (remanentesAGenerar.length > 0) {
        const lastRow = sheetCargos.getLastRow();
        sheetCargos.getRange(lastRow + 1, 1, remanentesAGenerar.length, remanentesAGenerar[0].length).setValues(remanentesAGenerar);
    }

    // 5. ACTUALIZAR SALDO A FAVOR
    let nuevoSaldoAFavor = bolsaDeDinero; 
    if (nuevoSaldoAFavor < 0.01) nuevoSaldoAFavor = 0; 
    
    updateAnticipo(idUnidad, nuevoSaldoAFavor);
    let anticipoGenerado = nuevoSaldoAFavor > saldoAFavorInicial ? nuevoSaldoAFavor - saldoAFavorInicial : 0;
    
    // 6. REGISTRAR PAGO Y GENERAR TICKET
    let cargoIdsPagados = conceptosCubiertosIds.join(', ');
    let conceptoDescriptivo = conceptosCubiertosIds.length > 0 ? getChargeConcept(conceptosCubiertosIds[0]) : 'ANTICIPO / SALDO A FAVOR';
    
    if(conceptosCubiertosIds.length > 1) {
        conceptoDescriptivo += " y otros cargos...";
    }

    const idPago = 'RGP-' + Utilities.getUuid().substring(0, 8).toUpperCase(); 

    const nuevaFilaPago = [
        idPago,                        
        fechaPago,                     
        idUnidad,                      
        capturista,                    
        montoRecibido,                 
        montoTotalAplicadoADeuda_Neto, 
        anticipoGenerado,              
        nuevoSaldoAFavor,              
        cargoIdsPagados,               
        conceptoDescriptivo            
    ];
    
    sheetPagos.appendRow(nuevaFilaPago);
    
    return { 
        success: true, 
        message: `Abono/Pago registrado con éxito.`,
        anticipoGenerado: anticipoGenerado, 
        nuevoSaldoAFavor: nuevoSaldoAFavor,
        idPago: idPago 
    };
}

function getChargeConcept(chargeId) {
    const sheetCargos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CARGOS_Y_DEUDAS");
    if (!sheetCargos) return "ERROR";
    const data = sheetCargos.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(chargeId)) return data[i][2] || "Concepto no especificado";
    }
    return `ID no encontrado`;
}

function getUnitAnticipo(idUnidad) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SALDOS_A_FAVOR");
    if (!sheet) return 0;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(idUnidad)) {
            return parseFloat(data[i][1]) || 0;
        }
    }
    return 0;
}

function updateAnticipo(idUnidad, nuevoSaldo) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("SALDOS_A_FAVOR");
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    let found = false;
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(idUnidad)) {
            sheet.getRange(i + 1, 2).setValue(nuevoSaldo);
            found = true;
            break;
        }
    }
    if (!found) {
        sheet.appendRow([idUnidad, nuevoSaldo]);
    }
}

// -----------------------------------------------------------------------------------
// MÓDULO DE ANULACIÓN BLINDADO
// -----------------------------------------------------------------------------------
// -----------------------------------------------------------------------------------
// MÓDULO DE ANULACIÓN BLINDADO (Actualizado para MOLINOS - Soporta Adelantos)
// -----------------------------------------------------------------------------------
function cancelPayment(idPago) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPagos = ss.getSheetByName("REGISTRO_PAGOS");
    const sheetCargos = ss.getSheetByName("CARGOS_Y_DEUDAS");
    
    const pagosData = sheetPagos.getDataRange().getValues();
    let filaPago = -1;
    let datosPago = null;

    // 1. Buscar el pago en REGISTRO_PAGOS
    for (let i = 1; i < pagosData.length; i++) {
        if (pagosData[i][0] === idPago) {
            filaPago = i + 1;
            datosPago = {
                idUnidad: pagosData[i][2],
                montoRecibido: parseFloat(pagosData[i][4]) || 0,
                montoAplicado: parseFloat(pagosData[i][5]) || 0,
                idsCargos: pagosData[i][8],
                conceptoAnulado: pagosData[i][9]
            };
            break;
        }
    }

    if (filaPago === -1) return { success: false, message: "ID de pago no encontrado." };
    if (String(datosPago.conceptoAnulado).includes("ANULADO")) return { success: false, message: "Este pago ya está anulado." };

    try {
        const cargosData = sheetCargos.getDataRange().getValues();
        const idsA_Revertir = String(datosPago.idsCargos).split(',').map(id => id.trim());
        
        // ⭐️ NUEVO: Determinar si este pago fue un "Adelanto" revisando el concepto del ticket
        const esAdelanto = String(datosPago.conceptoAnulado).toUpperCase().includes("ADELANTO");

        idsA_Revertir.forEach(idBuscado => {
            if(!idBuscado) return;
            
            for (let j = 1; j < cargosData.length; j++) {
                let idActual = String(cargosData[j][0]);
                
                if (idActual === idBuscado) {
                    
                    // ⭐️ NUEVO: Lógica de anulación para ADELANTOS
                    if (esAdelanto) {
                        // Si es un adelanto, la deuda original NO existía, así que BORRAMOS la fila completa.
                        sheetCargos.getRange(j + 1, 1, 1, 8).clearContent();
                    } 
                    // Lógica de anulación NORMAL (Pagos regulares y Abonos Parciales)
                    else {
                        let esAbonoParcial = false;
                        let filaHija = -1;
                        let conceptoBuscado = String(cargosData[j][2]);
                        
                        for (let k = 1; k < cargosData.length; k++) {
                            if (String(cargosData[k][1]) === String(datosPago.idUnidad) && 
                                String(cargosData[k][2]) === ("Saldo Restante: " + conceptoBuscado)) {
                                filaHija = k + 1;
                                esAbonoParcial = true;
                                break;
                            }
                        }

                        if (esAbonoParcial) {
                            let montoAbonado = parseFloat(cargosData[j][4]);
                            let montoRemanente = parseFloat(cargosData[filaHija - 1][4]);
                            sheetCargos.getRange(j + 1, 5).setValue(montoAbonado + montoRemanente); 
                            sheetCargos.getRange(filaHija, 1, 1, 8).clearContent(); 
                        }
                        
                        sheetCargos.getRange(j + 1, 6).setValue("Pendiente"); 
                    }
                }
            }
        });

        // 2. Restaurar el Saldo a Favor
        let saldoActual = getUnitAnticipo(datosPago.idUnidad);
        let impactoNeto = datosPago.montoRecibido - datosPago.montoAplicado;
        updateAnticipo(datosPago.idUnidad, saldoActual - impactoNeto);

        // 3. Marcar el ticket como ANULADO en REGISTRO_PAGOS
        sheetPagos.getRange(filaPago, 5, 1, 3).setValues([[0, 0, 0]]);
        sheetPagos.getRange(filaPago, 10).setValue(`ANULADO: ${datosPago.conceptoAnulado}`);
        sheetPagos.getRange(filaPago, 1, 1, 10).setBackground("#f8d7da");

        let mensajeExito = esAdelanto ? 
            `Anulación de Adelanto exitosa. Se eliminaron los cargos futuros generados.` : 
            `Anulación exitosa. Se ha restaurado la deuda original a estado Pendiente.`;

        return { success: true, message: mensajeExito };

    } catch (e) {
        return { success: false, message: "Error al anular: " + e.message };
    }
}
// -----------------------------------------------------------------------------------
// FUNCIONES DE ADMINISTRADOR PARA HTML
// -----------------------------------------------------------------------------------
function ADMIN_ejecutarAnulacionConClave(idPago, passwordIntroducida) {
    const CLAVE_AUTORIZADA = "Super25"; 
    if (passwordIntroducida !== CLAVE_AUTORIZADA) {
        return { success: false, message: "⚠️ CLAVE DE ADMINISTRADOR INCORRECTA." };
    }
    return cancelPayment(idPago);
}

function ADMIN_obtenerPagosPorUnidad(idUnidad) {
    const sheetPagos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("REGISTRO_PAGOS");
    const data = sheetPagos.getDataRange().getValues();
    return data.slice(1)
        .filter(row => String(row[2]) === String(idUnidad) && !String(row[9]).includes("ANULADO"))
        .map(row => ({
            idPago: row[0],
            fecha: Utilities.formatDate(new Date(row[1]), "GMT-6", "dd/MM/yyyy HH:mm"),
            monto: row[4]
        })).reverse(); 
}