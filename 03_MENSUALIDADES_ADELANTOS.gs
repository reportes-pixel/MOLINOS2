// ==============================================================================
// 03_MENSUALIDADES_ADELANTOS.gs - PROYECTO MOLINOS
// ==============================================================================

/**
 * 1. Generar Mensualidades MOLINOS (Smart)
 * Incluye: Mensualidad Base/PP + Rentas de Estacionamiento + Saldo a Favor
 */
/**
 * 1. Generar Mensualidades MOLINOS (Cargos Independientes)
 * Separa la Mensualidad de las Rentas de Estacionamiento en filas distintas.
 */
function generarCargosMensuales() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const unidadesSheet = ss.getSheetByName("UNIDADES");
    const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    const saldosSheet = ss.getSheetByName("SALDOS_A_FAVOR"); 
    const rentasSheet = ss.getSheetByName("RENTAS_ESTACIONAMIENTO");
    const config = getConfig();
    const ui = SpreadsheetApp.getUi();

    if (!unidadesSheet || !cargosSheet || !config || !saldosSheet || !rentasSheet) {
        ui.alert("Error", "Faltan hojas críticas para el proceso.", ui.ButtonSet.OK);
        return;
    }

    const montoNormalBase = Number(config.MENSUALIDAD_BASE) || 0;
    const montoPPBase = Number(config.MENSUALIDA_PRONTO_PAGO) || 0;
    const hoy = new Date();

    const prompt = ui.prompt("Generar Cargos MOLINOS", "¿Mes/Año? (Ej: 03/2026):", ui.ButtonSet.OK_CANCEL);
    if (prompt.getSelectedButton() !== ui.Button.OK) return;

    const input = prompt.getResponseText().trim();
    const parts = input.split('/');
    const mesCorte = new Date(parseInt(parts[1]), parseInt(parts[0]) - 1, 1);
    const txtMes = mesCorte.toLocaleString('es-ES', { month: 'short', year: 'numeric' });

    // 1. CARGAR DATOS
    const idsUnidades = unidadesSheet.getRange(2, 1, unidadesSheet.getLastRow() - 1, 1).getValues().flat();
    const cargosExistentes = cargosSheet.getDataRange().getValues().slice(1);
    const rentasData = rentasSheet.getDataRange().getValues().slice(1);
    const saldosData = saldosSheet.getDataRange().getValues().slice(1);

    const mapaSaldos = {};
    saldosData.forEach((row, i) => { 
        if(row[0]) mapaSaldos[String(row[0])] = { monto: Number(row[1]) || 0, fila: i + 2 }; 
    });

    const nuevosCargos = [];
    let contadorAbonos = 0;

    // 2. PROCESAR UNIDADES
    idsUnidades.forEach(id => {
        if(!id) return;
        id = String(id);

        // --- A. GENERAR CARGO DE MENSUALIDAD ---
        const yaExisteM = cargosExistentes.some(row => {
            const cFecha = new Date(row[3]);
            return String(row[1]) === id && String(row[2]).includes('Mensualidad') && 
                   cFecha.getFullYear() === mesCorte.getFullYear() && cFecha.getMonth() === mesCorte.getMonth();
        });

        if (!yaExisteM) {
            const tieneDeuda = cargosExistentes.some(row => String(row[1]) === id && String(row[5]).toUpperCase() === "PENDIENTE");
            let montoM = tieneDeuda ? montoNormalBase : montoPPBase;
            let conceptoM = `Mensualidad ${txtMes}${tieneDeuda ? "" : " PP"}`;
            
            // Aplicar saldo a favor a la mensualidad primero
            let {estado, montoFinal, pagoRef, saldoRestante} = aplicarSaldo(id, montoM, mapaSaldos, saldosSheet);
            if(pagoRef !== "") { contadorAbonos++; conceptoM += ` (${pagoRef})`; }
            
            nuevosCargos.push(['CGO-' + Utilities.getUuid().substring(0, 8).toUpperCase(), id, conceptoM, mesCorte, montoFinal, estado, pagoRef, ""]);
        }

        // --- B. GENERAR CARGOS DE ESTACIONAMIENTO (FILAS EXTRAS) ---
        rentasData.forEach(renta => {
            const idURenta = String(renta[1]);
            const montoR = Number(renta[4]) || 0;
            const estadoR = String(renta[5]).toUpperCase();
            const idRentaDoc = String(renta[0]);

            if (idURenta === id && estadoR === "ACTIVO") {
                // Checar si ya se generó este estacionamiento este mes
                const yaExisteE = cargosExistentes.some(row => {
                    const cFecha = new Date(row[3]);
                    return String(row[1]) === id && String(row[2]).includes(idRentaDoc) && 
                           cFecha.getFullYear() === mesCorte.getFullYear() && cFecha.getMonth() === mesCorte.getMonth();
                });

                if (!yaExisteE) {
                    let conceptoE = `Renta Estac. ${txtMes} (Ref: ${idRentaDoc})`;
                    let {estado, montoFinal, pagoRef} = aplicarSaldo(id, montoR, mapaSaldos, saldosSheet);
                    if(pagoRef !== "") { contadorAbonos++; conceptoE += ` (${pagoRef})`; }

                    nuevosCargos.push(['CGO-' + Utilities.getUuid().substring(0, 8).toUpperCase(), id, conceptoE, mesCorte, montoFinal, estado, pagoRef, ""]);
                }
            }
        });
    });

    // 3. GUARDAR
    if (nuevosCargos.length > 0) {
        cargosSheet.getRange(cargosSheet.getLastRow() + 1, 1, nuevosCargos.length, 8).setValues(nuevosCargos);
        ui.alert("Éxito", `Se generaron ${nuevosCargos.length} cargos independientes.\nAbonos automáticos: ${contadorAbonos}`, ui.ButtonSet.OK);
    } else {
        ui.alert("Aviso", "No hay nuevos cargos por generar.", ui.ButtonSet.OK);
    }
}

/**
 * Función auxiliar para aplicar saldo a favor de forma secuencial
 */
function aplicarSaldo(id, montoACobrar, mapaSaldos, saldosSheet) {
    let resultado = { estado: "Pendiente", montoFinal: montoACobrar, pagoRef: "" };
    
    if (mapaSaldos[id] && mapaSaldos[id].monto > 0) {
        let disponible = mapaSaldos[id].monto;
        
        if (disponible >= montoACobrar) {
            mapaSaldos[id].monto -= montoACobrar;
            resultado.estado = "Pagado";
            resultado.montoFinal = montoACobrar;
            resultado.pagoRef = "SALDO-TOTAL";
        } else {
            mapaSaldos[id].monto = 0;
            resultado.estado = "Pendiente";
            resultado.montoFinal = montoACobrar - disponible;
            resultado.pagoRef = `ABONO-$${disponible.toFixed(2)}`;
        }
        // Actualizar la celda físicamente
        saldosSheet.getRange(mapaSaldos[id].fila, 2).setValue(mapaSaldos[id].monto);
    }
    return resultado;
}


/**
 * 2. Corregir Monto Vencido (Fin de Pronto Pago)
 */
function corregirMontoVencido() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    const config = getConfig(); 

    if (!cargosSheet || !config) {
        ui.alert("Error", "Faltan hojas necesarias.", ui.ButtonSet.OK);
        return;
    }

    const montoBaseNormal = Number(config.MENSUALIDAD_BASE) || 0; 
    const montoBasePP = Number(config.MENSUALIDA_PRONTO_PAGO) || 0; 
    
    const prompt = ui.prompt("Corregir Montos Vencidos (PP)", "Introduce MES/AÑO (Ej: 02/2026):", ui.ButtonSet.OK_CANCEL);
    if (prompt.getSelectedButton() !== ui.Button.OK) return;
    
    const input = prompt.getResponseText().trim();
    const [mesStr, anioStr] = input.split('/');
    const mesObjetivo = parseInt(mesStr) - 1; 
    const anioObjetivo = parseInt(anioStr);
    
    const lastRow = cargosSheet.getLastRow();
    if (lastRow < 2) return;
    
    const cargosRange = cargosSheet.getRange(2, 1, lastRow - 1, 6);
    const cargosData = cargosRange.getValues(); 
    let corregidos = 0;

    cargosData.forEach((row, index) => {
        const concepto = String(row[2]);
        const fecha = new Date(row[3]);
        const estado = String(row[5]).toUpperCase();
        const rowIndex = index + 2;

        if (concepto.includes(' PP') && fecha.getMonth() === mesObjetivo && fecha.getFullYear() === anioObjetivo && estado === "PENDIENTE") {
            // Actualizar Monto y quitar ' PP' del concepto
            cargosSheet.getRange(rowIndex, 5).setValue(montoBaseNormal);
            cargosSheet.getRange(rowIndex, 3).setValue(concepto.replace(' PP', '').trim());
            corregidos++;
        }
    });

    ui.alert("Éxito", `Se corrigieron ${corregidos} cargos vencidos.`, ui.ButtonSet.OK);
}

/**
 * 3. Generar Recargos por Mora
 */
function generarRecargosPorMora() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    const config = getConfig(); 
    
    if (!cargosSheet || !config) return ui.alert("Error", "Faltan hojas.", ui.ButtonSet.OK);

    const tasaRecargo = Number(config.TASA_RECARGO) || 0; 
    const fechaHoy = new Date();

    const prompt = ui.prompt("Aplicar Recargos", "MES/AÑO de la deuda a penalizar (Ej: 03/2026):", ui.ButtonSet.OK_CANCEL);
    if (prompt.getSelectedButton() !== ui.Button.OK) return;

    const input = prompt.getResponseText().trim();
    const partes = input.split('/');
    const mesObj = parseInt(partes[0]) - 1;
    const anioObj = parseInt(partes[1]);
    
    const data = cargosSheet.getRange(2, 1, cargosSheet.getLastRow() - 1, 8).getValues(); 
    const nuevosRecargos = [];

    data.forEach((row, index) => {
        const idU = row[1];
        const concepto = String(row[2]);
        const fechaCorte = new Date(row[3]);
        const montoBase = Number(row[4]);
        const estado = String(row[5]).toUpperCase();
        const tieneRecargo = row[7];

        if (estado === "PENDIENTE" && !tieneRecargo && fechaCorte.getMonth() === mesObj && fechaCorte.getFullYear() === anioObj) {
            const montoRecargo = Math.round((montoBase * tasaRecargo) * 100) / 100;
            if (montoRecargo > 0) {
                const idR = 'R-' + Utilities.getUuid().substring(0, 8).toUpperCase();
                nuevosRecargos.push([idR, idU, `Recargo Mora (${input})`, fechaHoy, montoRecargo, "Pendiente", "", ""]);
                cargosSheet.getRange(index + 2, 8).setValue(idR); // Marcar que ya tiene recargo
            }
        }
    });

    if (nuevosRecargos.length > 0) {
        cargosSheet.getRange(cargosSheet.getLastRow() + 1, 1, nuevosRecargos.length, 8).setValues(nuevosRecargos);
        ui.alert("Éxito", `Se aplicaron ${nuevosRecargos.length} recargos.`, ui.ButtonSet.OK);
    } else {
        ui.alert("Aviso", "No hay cargos pendientes para recargo en ese mes.", ui.ButtonSet.OK);
    }
}


/**
 * 4. Procesar Adelanto de Mensualidades
 * Genera N meses futuros de CUOTA DE MANTENIMIENTO a precio de Pronto Pago
 * y los marca como Pagados inmediatamente, generando un ticket unificado.
 */
function procesarAdelantoMensualidades(data) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
        const pagosSheet = ss.getSheetByName("REGISTRO_PAGOS");
        const config = getConfig();

        if (!cargosSheet || !pagosSheet || !config) {
            return { success: false, message: "Error: Faltan hojas de base de datos o CONFIGURACION." };
        }

        // ⭐️ REGLA DE ORO: Siempre a precio de Pronto Pago
        const montoPPBase = Number(config.MENSUALIDA_PRONTO_PAGO) || 0;
        if (montoPPBase <= 0) {
            return { success: false, message: "El monto de Pronto Pago no está configurado correctamente." };
        }

        const { idUnidad, capturista, numMeses, mesInicioStr } = data;
        const numMesesInt = parseInt(numMeses);
        
        // Parsear la fecha de inicio enviada desde el HTML (ej. "2026-04-01")
        const partesFecha = mesInicioStr.split('-');
        const anioInicio = parseInt(partesFecha[0]);
        const mesInicio = parseInt(partesFecha[1]) - 1; // En JavaScript los meses van del 0 al 11
        
        let fechaActual = new Date(anioInicio, mesInicio, 1);
        const fechaRegistro = new Date();
        const idPagoUnificado = 'RGP-' + Utilities.getUuid().substring(0, 8).toUpperCase();

        let nuevosCargos = [];
        let cargosIdsPagados = [];
        let totalCobrado = 0;

        // Generar los cargos futuros
        for (let i = 0; i < numMesesInt; i++) {
            let mesCorte = new Date(fechaActual.getFullYear(), fechaActual.getMonth() + i, 1);
            let txtMes = mesCorte.toLocaleString('es-ES', { month: 'short', year: 'numeric' });
            
            let idCargo = 'CGO-' + Utilities.getUuid().substring(0, 8).toUpperCase();
            let concepto = `Mensualidad ${txtMes} PP (Adelanto)`;

            // [ID, Unidad, Concepto, Fecha, Monto, Estado, ID_Pago, ID_Recargo]
            nuevosCargos.push([idCargo, idUnidad, concepto, mesCorte, montoPPBase, "Pagado", idPagoUnificado, ""]);
            
            cargosIdsPagados.push(idCargo);
            totalCobrado += montoPPBase;
        }

        // 1. Inyectar los nuevos cargos ya marcados como "Pagados"
        cargosSheet.getRange(cargosSheet.getLastRow() + 1, 1, nuevosCargos.length, 8).setValues(nuevosCargos);

        // 2. Registrar el pago unificado en la hoja REGISTRO_PAGOS
        let txtPrimerMes = nuevosCargos[0][2].replace('Mensualidad ', '').replace(' PP (Adelanto)', '');
        let txtUltimoMes = nuevosCargos[nuevosCargos.length - 1][2].replace('Mensualidad ', '').replace(' PP (Adelanto)', '');
        let mesesAbarcados = numMesesInt > 1 ? `${txtPrimerMes} a ${txtUltimoMes}` : txtPrimerMes;
        
        // Para obtener el saldo a favor actual sin modificarlo
        let saldoAFavorActual = 0;
        try { saldoAFavorActual = getUnitAnticipo(idUnidad); } catch(e) {} 
        
        const nuevaFilaPago = [
            idPagoUnificado,               // A: ID_PAGO
            fechaRegistro,                 // B: FECHA_PAGO
            idUnidad,                      // C: ID_UNIDAD
            capturista,                    // D: CAPTURISTA
            totalCobrado,                  // E: MONTO_RECIBIDO
            totalCobrado,                  // F: MONTO_APLICADO_DEUDA
            0,                             // G: ANTICIPO_GENERADO
            saldoAFavorActual,             // H: SALDO_A_FAVOR_FINAL (No cambia)
            cargosIdsPagados.join(', '),   // I: ID_CARGO_CUBIERTOS
            `Adelanto de ${numMesesInt} mes(es) de mantenimiento (${mesesAbarcados})` // J: CONCEPTO_PAGADO
        ];

        pagosSheet.appendRow(nuevaFilaPago);

        return { 
            success: true, 
            message: `Adelanto por $${totalCobrado.toFixed(2)} registrado exitosamente.`,
            idPago: idPagoUnificado 
        };

    } catch (e) {
        return { success: false, message: "Error interno: " + e.message };
    }
}

/**
 * 5. Obtener Mes de Inicio para Adelantos (Calendario Inteligente)
 * Busca la última mensualidad cobrada a la unidad y devuelve el mes siguiente.
 */
function obtenerMesInicioAdelanto(idUnidad) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    
    if (!cargosSheet) return null;

    const data = cargosSheet.getDataRange().getValues();
    let maxDate = new Date(2000, 0, 1); // Fecha muy antigua de referencia
    let found = false;

    // Buscar el cargo de "Mensualidad" más reciente de esta unidad
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][1]) === String(idUnidad) && String(data[i][2]).includes("Mensualidad")) {
            let fechaCorte = new Date(data[i][3]);
            if (fechaCorte > maxDate) {
                maxDate = fechaCorte;
                found = true;
            }
        }
    }

    let nextMonth;
    if (found) {
        // Si encontramos pagos anteriores, sugerimos el mes SIGUIENTE
        nextMonth = new Date(maxDate.getFullYear(), maxDate.getMonth() + 1, 1);
    } else {
        // Si es una unidad totalmente nueva sin historial, sugerimos el mes ACTUAL
        const today = new Date();
        nextMonth = new Date(today.getFullYear(), today.getMonth(), 1);
    }

    // Devolver formato YYYY-MM que requiere el input de HTML
    const mm = String(nextMonth.getMonth() + 1).padStart(2, '0');
    const yyyy = nextMonth.getFullYear();
    
    return `${yyyy}-${mm}`;
}

/**
 * 4. Procesar Adelanto de Mensualidades
 * Genera N meses futuros de CUOTA DE MANTENIMIENTO a precio de Pronto Pago
 * y los marca como Pagados inmediatamente, generando un ticket unificado.
 */
function procesarAdelantoMensualidades(data) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
        const pagosSheet = ss.getSheetByName("REGISTRO_PAGOS");
        const config = getConfig();

        if (!cargosSheet || !pagosSheet || !config) {
            return { success: false, message: "Error: Faltan hojas de base de datos o CONFIGURACION." };
        }

        const montoPPBase = Number(config.MENSUALIDA_PRONTO_PAGO) || 0;
        if (montoPPBase <= 0) {
            return { success: false, message: "El monto de Pronto Pago no está configurado." };
        }

        const { idUnidad, capturista, numMeses, mesInicioStr } = data;
        const numMesesInt = parseInt(numMeses);
        
        const partesFecha = mesInicioStr.split('-');
        const anioInicio = parseInt(partesFecha[0]);
        const mesInicio = parseInt(partesFecha[1]) - 1; 
        
        let fechaActual = new Date(anioInicio, mesInicio, 1);
        const fechaRegistro = new Date();
        const idPagoUnificado = 'RGP-' + Utilities.getUuid().substring(0, 8).toUpperCase();

        let nuevosCargos = [];
        let cargosIdsPagados = [];
        let totalCobrado = 0;

        for (let i = 0; i < numMesesInt; i++) {
            let mesCorte = new Date(fechaActual.getFullYear(), fechaActual.getMonth() + i, 1);
            let txtMes = mesCorte.toLocaleString('es-ES', { month: 'short', year: 'numeric' });
            
            let idCargo = 'CGO-' + Utilities.getUuid().substring(0, 8).toUpperCase();
            let concepto = `Mensualidad ${txtMes} PP (Adelanto)`;

            nuevosCargos.push([idCargo, idUnidad, concepto, mesCorte, montoPPBase, "Pagado", idPagoUnificado, ""]);
            
            cargosIdsPagados.push(idCargo);
            totalCobrado += montoPPBase;
        }

        cargosSheet.getRange(cargosSheet.getLastRow() + 1, 1, nuevosCargos.length, 8).setValues(nuevosCargos);

        let txtPrimerMes = nuevosCargos[0][2].replace('Mensualidad ', '').replace(' PP (Adelanto)', '');
        let txtUltimoMes = nuevosCargos[nuevosCargos.length - 1][2].replace('Mensualidad ', '').replace(' PP (Adelanto)', '');
        let mesesAbarcados = numMesesInt > 1 ? `${txtPrimerMes} a ${txtUltimoMes}` : txtPrimerMes;
        
        let saldoAFavorActual = 0;
        try { saldoAFavorActual = getUnitAnticipo(idUnidad); } catch(e) {} 
        
        const nuevaFilaPago = [
            idPagoUnificado,               
            fechaRegistro,                 
            idUnidad,                      
            capturista,                    
            totalCobrado,                  
            totalCobrado,                  
            0,                             
            saldoAFavorActual,             
            cargosIdsPagados.join(', '),   
            `Adelanto de ${numMesesInt} mes(es) de mantenimiento (${mesesAbarcados})` 
        ];

        pagosSheet.appendRow(nuevaFilaPago);

        return { 
            success: true, 
            message: `Adelanto por $${totalCobrado.toFixed(2)} registrado exitosamente.`,
            idPago: idPagoUnificado 
        };

    } catch (e) {
        return { success: false, message: "Error interno: " + e.message };
    }
}

