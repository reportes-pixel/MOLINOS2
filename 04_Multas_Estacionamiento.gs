// ==============================================================================
// 04_MULTAS_ESTACIONAMIENTO.gs - MULTAS Y SUSCRIPCIONES DE CAJONES
// ==============================================================================

function processFineRegistration(fineData) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
        const { idUnidad, idMulta, capturista, fechaAplicacionStr } = fineData; 

        const fineConfigResult = getFineConfig();
        if (fineConfigResult.error) return { success: false, message: fineConfigResult.error };
        
        const fineConfig = fineConfigResult.fines.find(f => f.idMulta === idMulta);
        if (!fineConfig) return { success: false, message: `La multa ${idMulta} no existe.` };
        
        const fechaAplicacion = new Date(fechaAplicacionStr);
        const idCargo = 'M-' + Utilities.getUuid().substring(0, 8).toUpperCase(); 

        cargosSheet.appendRow([
            idCargo, idUnidad.toUpperCase(), `Multa: ${fineConfig.concepto}`, 
            fechaAplicacion, fineConfig.montoBase, "Pendiente", "", ""
        ]);
        
        return { success: true, message: `Multa registrada para la unidad ${idUnidad}.` };
    } catch (e) {
        return { success: false, message: `Error: ${e.message}` };
    }
}

// ⭐️ NUEVO: CONTRATOS DE ESTACIONAMIENTO
function procesarRentaEstacionamiento(data) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetRentas = ss.getSheetByName("RENTAS_ESTACIONAMIENTO");
    
    // Si la hoja no existe, la crea invisiblemente
    if (!sheetRentas) {
        sheetRentas = ss.insertSheet("RENTAS_ESTACIONAMIENTO");
        sheetRentas.appendRow(["ID_RENTA", "ID_UNIDAD", "FECHA_INICIO", "FECHA_TERMINO", "MONTO_MENSUAL", "ESTADO", "CAPTURISTA"]);
    }

    const idRenta = 'EST-' + Utilities.getUuid().substring(0, 6).toUpperCase();
    
    sheetRentas.appendRow([
        idRenta, 
        data.idUnidad, 
        new Date(data.fechaInicio), 
        new Date(data.fechaTermino), 
        parseFloat(data.monto), 
        "ACTIVO",
        data.capturista
    ]);

    return { success: true, message: `Contrato de estacionamiento registrado exitosamente.` };
}

// ⭐️ NUEVO: GENERADOR AUTOMÁTICO DE ESTACIONAMIENTO (Se conecta con la mensualidad)
function generarCargosEstacionamiento(fechaCorte) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetRentas = ss.getSheetByName("RENTAS_ESTACIONAMIENTO");
    const sheetCargos = ss.getSheetByName("CARGOS_Y_DEUDAS");
    
    if(!sheetRentas || !sheetCargos) return;

    const rentas = sheetRentas.getDataRange().getValues().slice(1);
    const cargosExistentes = sheetCargos.getDataRange().getValues().slice(1);
    const txtMes = fechaCorte.toLocaleString('es-ES', { month: 'short', year: 'numeric' });
    const nuevosCargos = [];

    rentas.forEach(r => {
        if(r[5] !== "ACTIVO") return; // Solo contratos activos
        
        let fInicio = new Date(r[2]);
        let fFin = new Date(r[3]);
        
        // Verificamos si el mes de corte cae dentro del contrato de renta
        if(fechaCorte >= fInicio && fechaCorte <= fFin) {
            let concepto = `Renta Estacionamiento ${txtMes}`;
            
            // Verificamos que no se haya cobrado ya este mes
            let yaExiste = cargosExistentes.some(c => String(c[1]) === String(r[1]) && String(c[2]) === concepto);
            
            if(!yaExiste) {
                let idCargo = 'CGO-' + Utilities.getUuid().substring(0, 8).toUpperCase();
                nuevosCargos.push([idCargo, r[1], concepto, fechaCorte, parseFloat(r[4]), "Pendiente", "", ""]);
            }
        }
    });

    if(nuevosCargos.length > 0) {
        sheetCargos.getRange(sheetCargos.getLastRow() + 1, 1, nuevosCargos.length, 8).setValues(nuevosCargos);
    }
}

function guardarCargoExtraordinario(data) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
        
        if (!cargosSheet) return { success: false, message: "No se encontró la hoja CARGOS_Y_DEUDAS." };
        
        const idCargo = 'CGO-EXT-' + Utilities.getUuid().substring(0, 6).toUpperCase();
        const fechaHoy = new Date();
        const montoNum = parseFloat(data.monto);
        
        cargosSheet.appendRow([
            idCargo, 
            data.idUnidad, 
            data.concepto, 
            fechaHoy, 
            montoNum, 
            "Pendiente", 
            "", 
            ""
        ]);
        
        return { success: true, message: "Cargo extraordinario registrado exitosamente." };
    } catch (e) {
        return { success: false, message: "Error: " + e.message };
    }
}
