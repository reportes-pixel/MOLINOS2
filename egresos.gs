

/**
 * Busca un egreso por su ID para mandarlo al formulario HTML
 */
function buscarEgresoPorId(idEgreso) {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("EGRESOS");
    if (!sheet) return { success: false, message: "No existe la base de EGRESOS." };
    
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        if (data[i][0] === idEgreso.trim().toUpperCase()) {
            
            // Formatear fecha para el input type="date" (YYYY-MM-DD)
            let d = data[i][6];
            let fEmision = (d instanceof Date) ? new Date(d.getTime() - d.getTimezoneOffset() * 60000).toISOString().substring(0, 10) : "";
            
            // Formatear mes para el input type="month" (YYYY-MM)
            let m = data[i][2] ? data[i][2].toString() : "";
            let mStr = m.includes('/') ? m.split('/')[1] + '-' + m.split('/')[0] : "";

            return {
                success: true,
                fila: i + 1,
                mesEgresoStr: mStr,
                concepto: data[i][3],
                proveedor: data[i][4],
                monto: data[i][5],
                fechaEmisionStr: fEmision,
                capturista: data[i][8]
            };
        }
    }
    return { success: false, message: "No se encontró ningún egreso con ese Folio." };
}

/**
 * Procesa el registro NUEVO o la EDICIÓN de un Egreso.
 */
function processExternalDebtRegistration(debtData) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("EGRESOS");
    
    // Si no existe, la crea con 10 columnas (La 10 es CLAVE_CONTROL)
    if (!sheet) {
        sheet = ss.insertSheet("EGRESOS");
        sheet.getRange(1, 1, 1, 10).setValues([
            ["ID_EGRESO", "FECHA_REGISTRO", "MES_EGRESO", "CONCEPTO", "PROVEEDOR", "MONTO", "FECHA_EMISION", "ESTADO", "CAPTURISTA", "CONTROL_EDICION"]
        ]);
    }

    const { modo, idEgreso, mesEgresoStr, concepto, proveedor, monto, fechaEmisionStr, capturista, claveControl } = debtData;
    
    const montoNum = parseFloat(monto);
    if (isNaN(montoNum) || montoNum <= 0) return { success: false, message: "El monto debe ser positivo." };
    
    const mesPartes = mesEgresoStr.split('-');
    const mesFormateado = mesPartes.length === 2 ? `${mesPartes[1]}/${mesPartes[0]}` : mesEgresoStr;

    let fechaEmision = new Date(fechaEmisionStr);
    fechaEmision = new Date(fechaEmision.getTime() + fechaEmision.getTimezoneOffset() * 60000);

    const fechaAccion = new Date();

    if (modo === 'editar') {
        // --- MODO EDICIÓN ---
        if (!idEgreso || !claveControl) return { success: false, message: "Falta el Folio o la Clave de Control para editar." };
        
        const data = sheet.getDataRange().getValues();
        let filaEncontrada = -1;
        
        for (let i = 1; i < data.length; i++) {
            if (data[i][0] === idEgreso.trim().toUpperCase()) {
                filaEncontrada = i + 1;
                break;
            }
        }
        
        if (filaEncontrada === -1) return { success: false, message: "El folio no existe." };

        // Sobreescribir las celdas de la fila encontrada
        sheet.getRange(filaEncontrada, 3).setValue(mesFormateado);
        sheet.getRange(filaEncontrada, 4).setValue(concepto);
        sheet.getRange(filaEncontrada, 5).setValue(proveedor);
        sheet.getRange(filaEncontrada, 6).setValue(montoNum);
        sheet.getRange(filaEncontrada, 7).setValue(fechaEmision);
        sheet.getRange(filaEncontrada, 9).setValue(capturista);
        
        // ⭐️ Auditoría en la columna 10
        const rastro = `Editado el ${fechaAccion.toLocaleDateString()} con clave: ${claveControl}`;
        sheet.getRange(filaEncontrada, 10).setValue(rastro);

        return { success: true, message: `El egreso ${idEgreso} fue actualizado correctamente.` };

    } else {
        // --- MODO NUEVO ---
        const nuevoId = 'EGR-' + Utilities.getUuid().substring(0, 8).toUpperCase(); 
        
        const nuevoEgreso = [
            nuevoId, fechaAccion, mesFormateado, concepto, proveedor, 
            montoNum, fechaEmision, "Registrado", capturista, 
            claveControl ? `Autorizado con clave: ${claveControl}` : "Registro Original" // Columna 10
        ];
        
        sheet.appendRow(nuevoEgreso);
        return { success: true, message: `Egreso registrado. Folio: ${nuevoId}` };
    }
}