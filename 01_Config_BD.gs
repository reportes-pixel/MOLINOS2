// ==============================================================================
// 01_CONFIG_BD.gs - LECTURA DE CONFIGURACIONES, LISTAS Y SALDOS
// ==============================================================================

function getUnitList() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("UNIDADES");
    if (!sheet) return { error: "Hoja 'UNIDADES' no encontrada.", units: [] };

    const data = sheet.getDataRange().getValues();
    const units = [];
    for (let i = 1; i < data.length; i++) {
        if (data[i][0] && data[i][1] && data[i][2]) {
            units.push({
                id: data[i][0],
                departamento: data[i][1],
                propietario: data[i][2],
                email: data[i][4] || "" // <--- ¡AQUÍ ESTÁ LA MAGIA! Ahora sí manda el correo
            });
        }
    }
    return { units: units };
}

function getIdsUnidadesParaForm() {
    const result = getUnitList();
    if (result.error) return [];
    return result.units.map(unit => unit.id);
}

function getUnitDetails(idUnidad) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("UNIDADES");
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(idUnidad)) {
            return {
                id: data[i][0],
                departamento: data[i][1], 
                propietario: data[i][2],
                email: data[i][4] || "" // COLUMNA E (Índice 4): EMAIL DE LA UNIDAD (NUEVO)
            };
        }
    }
    return null;
}

// ⭐️ NUEVA FUNCIÓN: Guarda el correo electrónico cuando lo piden desde el ticket
function guardarEmailUnidad(idUnidad, email) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("UNIDADES");
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(idUnidad)) {
            // Columna E = 5
            sheet.getRange(i + 1, 5).setValue(email);
            return { success: true };
        }
    }
    return { success: false, message: "Unidad no encontrada." };
}

function getCapturistaList() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const usuariosSheet = ss.getSheetByName("USUARIOS");
        if (!usuariosSheet) {
            const userEmail = Session.getActiveUser().getEmail();
            const userName = userEmail ? userEmail.substring(0, userEmail.indexOf('@')) : 'Administrador';
            return { capturistas: [userName] };
        }
        const lastRow = usuariosSheet.getLastRow();
        if (lastRow < 2) return { capturistas: [] };
        
        const capturistas = usuariosSheet.getRange(2, 1, lastRow - 1, 1).getValues().flat().filter(String);
        return { capturistas: capturistas };
    } catch (e) {
        return { error: `Error: ${e.message}` };
    }
}

function getConfig() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName("CONFIGURACION");
    if (!configSheet) return null;

    const fullData = configSheet.getDataRange().getValues();
    const config = {};
    if (fullData.length === 0) return {};

    fullData.forEach(row => {
        const clave = row[0]; 
        const valor = row[1];
        const claveStr = clave ? String(clave).toString().trim() : '';
        
        if (claveStr && claveStr !== "") {
            const numValue = Number(valor);
            if (!isNaN(numValue) && typeof valor === 'number') {
                config[claveStr] = valor;
            } else if (!isNaN(numValue) && String(valor).trim() !== "") {
                config[claveStr] = numValue;
            } else {
                config[claveStr] = valor;
            }
        }
    });
    return config;
}

function getFineConfig() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("CONFIG_MULTAS");
    if (!sheet) return { error: "Hoja 'CONFIG_MULTAS' no encontrada.", fines: [] };

    const data = sheet.getDataRange().getValues();
    const fines = [];
    for (let i = 1; i < data.length; i++) {
        fines.push({
            idMulta: data[i][0],
            concepto: data[i][1],
            montoBase: parseFloat(data[i][2]) || 0,
            montoRecargo: parseFloat(data[i][3]) || 0,
            diaLimitePP: parseInt(data[i][4]) || 0,
            tasaProntoPago: parseFloat(data[i][5]) || 0 // Agregado si usas Tasa en Config Multas
        });
    }
    return { fines: fines };
}

function getUnitAnticipo(idUnidad) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("SALDOS_A_FAVOR"); 
    if (!sheet) return 0; 
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) { 
        if (String(data[i][0]) === String(idUnidad)) {
            return parseFloat(data[i][1]) || 0; 
        }
    }
    return 0; 
}

function updateAnticipo(idUnidad, monto) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("SALDOS_A_FAVOR");
    if (!sheet) return;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][0]) === String(idUnidad)) {
            sheet.getRange(i + 1, 2).setValue(monto);
            return;
        }
    }
    sheet.appendRow([idUnidad, monto]);
}

function getDebtsForUnit(idUnidad) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    if (!sheet) return { filteredDeudas: [], totalPendiente: 0 };

    const data = sheet.getDataRange().getValues();
    let filteredDeudas = [];
    let totalPendiente = 0;
    
    const COLUMN_ID_CARGO = 0, COLUMN_ID_UNIDAD = 1, COLUMN_CONCEPTO = 2;
    const COLUMN_MES_CORTE = 3, COLUMN_MONTO_BASE = 4, COLUMN_ESTADO = 5;

    for (let i = 1; i < data.length; i++) {
        const row = data[i];
        if (String(row[COLUMN_ID_UNIDAD]) === String(idUnidad) && String(row[COLUMN_ESTADO]).toUpperCase() === "PENDIENTE") {
            const montoBase = parseFloat(row[COLUMN_MONTO_BASE]) || 0;
            totalPendiente += montoBase;
            
            filteredDeudas.push({
                idCargo: row[COLUMN_ID_CARGO],
                concepto: row[COLUMN_CONCEPTO],
                montoBase: montoBase,
                mesCorte: Utilities.formatDate(new Date(row[COLUMN_MES_CORTE]), ss.getSpreadsheetTimeZone(), "yyyy-MM-dd"), 
                isFine: String(row[COLUMN_CONCEPTO]).includes("Multa")
            });
        }
    }
    return { filteredDeudas, totalPendiente };
}

function getUnitDebt(idUnidad) {
    const unitDetails = getUnitDetails(idUnidad);
    if (!unitDetails) return { error: "Unidad no encontrada." };

    const { filteredDeudas, totalPendiente } = getDebtsForUnit(idUnidad); 
    const saldoAFavor = getUnitAnticipo(idUnidad); 
    
    let montoNetoPendiente = totalPendiente - saldoAFavor;
    if (montoNetoPendiente < 0) montoNetoPendiente = 0;

    return {
        departamento: unitDetails.departamento,
        propietario: unitDetails.propietario,
        deudas: filteredDeudas, 
        totalPendiente: totalPendiente, 
        saldoAFavor: saldoAFavor,       
        montoNetoPendiente: montoNetoPendiente 
    };
}