// ==============================================================================
// 05_INTERESES_MORA.gs - CÁLCULO DE INTERÉS DEL 10% CON MEMORIA
// ==============================================================================
/**
 * Genera el cargo de Interés por Mora dinámico
 * Toma la tasa de la pestaña CONFIGURACION (TASA_INTERES_MORA)
 */
function generarInteresesMora() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();
    const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    const config = getConfig(); // ⭐️ Llamamos a tu configuración general
    
    if (!cargosSheet || !config) {
        ui.alert("Error", "Faltan hojas 'CARGOS_Y_DEUDAS' o 'CONFIGURACION'.", ui.ButtonSet.OK);
        return;
    }

    // ⭐️ Jalamos la variable de la hoja. Si por algún error la borran, usa 0.10 por seguridad.
    const tasaInteres = Number(config.TASA_INTERES_MORA) || 0.10; 
    const porcentajeMostrar = (tasaInteres * 100).toFixed(0); // Convierte 0.10 en "10" para los textos

    if (tasaInteres <= 0) {
        ui.alert("Advertencia", "El valor de TASA_INTERES_MORA en CONFIGURACION es 0 o no válido.", ui.ButtonSet.OK);
        return;
    }

    const fechaHoy = new Date();
    const prompt = ui.prompt(
        `Generar Intereses Moratorios (${porcentajeMostrar}%)`,
        `Introduce el MES/AÑO actual (Ej: ${fechaHoy.getMonth() + 1}/${fechaHoy.getFullYear()}).\n\n⚠️ El interés SOLO se aplicará a deudas de meses anteriores y NO se cobrará interés sobre recargos.`,
        ui.ButtonSet.OK_CANCEL
    );

    if (prompt.getSelectedButton() !== ui.Button.OK) return;
    
    const inputMesAnio = prompt.getResponseText().trim();
    const partes = inputMesAnio.split('/');
    if (partes.length !== 2) {
        ui.alert("Error", "Formato inválido. Use MM/AAAA.", ui.ButtonSet.OK);
        return;
    }

    const mesActual = parseInt(partes[0]) - 1;
    const anioActual = parseInt(partes[1]);
    const fechaFrontera = new Date(anioActual, mesActual, 1); 

    const lastRow = cargosSheet.getLastRow();
    if (lastRow < 2) return;
    
    const data = cargosSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const nuevosCargos = [];
    let cargosGenerados = 0;

    data.forEach((row) => {
        const idUnidad = row[1];
        const conceptoOriginal = row[2] ? row[2].toString() : ''; 
        const fechaCargo = new Date(row[3]);
        const montoBase = Number(row[4]) || 0;
        const estado = row[5] ? row[5].toString().toUpperCase() : '';

        // Filtramos pagados y meses actuales
        if (estado !== "PENDIENTE" || fechaCargo >= fechaFrontera) return;
        
        // Evitamos cobrar interés sobre otro interés o recargo
        if (conceptoOriginal.includes("Interés") || conceptoOriginal.includes("Recargo")) return;

        const montoInteres = Math.round((montoBase * tasaInteres) * 100) / 100;
        
        if (montoInteres > 0) {
            const idInteres = 'INT-' + Utilities.getUuid().substring(0, 6).toUpperCase();
            
            // ⭐️ El concepto ahora es dinámico e imprime el % real de tu configuración
            const conceptoAmigable = `Interés ${porcentajeMostrar}% [${inputMesAnio}] s/ ${conceptoOriginal}`;

            nuevosCargos.push([
                idInteres, idUnidad, conceptoAmigable, fechaHoy, 
                montoInteres, "Pendiente", "", ""
            ]);
            cargosGenerados++;
        }
    });

    if (nuevosCargos.length > 0) {
        cargosSheet.getRange(cargosSheet.getLastRow() + 1, 1, nuevosCargos.length, 8).setValues(nuevosCargos);
        ui.alert("Éxito", `Se generaron ${cargosGenerados} cargos de interés limpios y legibles al ${porcentajeMostrar}%.`, ui.ButtonSet.OK);
    } else {
        ui.alert("Aviso", "No se encontraron deudas viejas que califiquen para aplicar el interés.", ui.ButtonSet.OK);
    }
}