// ==============================================================================
// 04_GESTOR_CARGOS_V2.gs - FUNCIONES CLONADAS SILENCIOSAS (BACKEND HTML)
// ==============================================================================

// Función para abrir la nueva ventana desde el Menú
function accesoGestorCargosMaestro() {
  const html = HtmlService.createTemplateFromFile('Form_GestorCargos')
      .evaluate()
      .setWidth(700)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Gestor Maestro de Cargos');
}

// ------------------------------------------------------------------------------
// 1. GENERAR MENSUALIDADES V2 (¡CON LÓGICA VIP / INSEN!)
// ------------------------------------------------------------------------------
function generarCargosMensuales_V2(mesStr) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const unidadesSheet = ss.getSheetByName("UNIDADES");
    const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    const saldosSheet = ss.getSheetByName("SALDOS_A_FAVOR"); 
    const rentasSheet = ss.getSheetByName("RENTAS_ESTACIONAMIENTO");
    const configSheet = ss.getSheetByName("CONFIGURACION");
    const config = getConfig();

    if (!unidadesSheet || !cargosSheet || !config || !saldosSheet || !rentasSheet || !configSheet) {
        throw new Error("Faltan hojas críticas para el proceso.");
    }

    const montoNormalBase = Number(config.MENSUALIDAD_BASE) || 0;
    const montoPPBase = Number(config.MENSUALIDA_PRONTO_PAGO) || 0;
    
    // Convertir el string del HTML ("YYYY-MM") a Fecha
    const partes = mesStr.split('-');
    const mesCorte = new Date(parseInt(partes[0]), parseInt(partes[1]) - 1, 1);
    const txtMes = mesCorte.toLocaleString('es-ES', { month: 'short', year: 'numeric' });

    // 1. CARGAR DATOS
    const idsUnidades = unidadesSheet.getRange(2, 1, unidadesSheet.getLastRow() - 1, 1).getValues().flat();
    const cargosExistentes = cargosSheet.getDataRange().getValues().slice(1);
    const rentasData = rentasSheet.getDataRange().getValues().slice(1);
    const saldosData = saldosSheet.getDataRange().getValues().slice(1);

    // --- MAGIA VIP: LEER COLUMNAS I, J, K (Índices 9, 10, 11 en notación A1) ---
    const vipData = configSheet.getRange("I2:K" + Math.max(configSheet.getLastRow(), 2)).getValues();
    const mapaVIP = {};
    vipData.forEach(row => {
        const depto = String(row[0]).trim();
        if(depto) {
            mapaVIP[depto] = {
                cuota: Number(row[1]) || 0,
                pierdeSiDebe: String(row[2]).trim().toUpperCase() === "SI"
            };
        }
    });

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
            let sufijo = tieneDeuda ? "" : " PP";
            
            // --- EVALUACIÓN VIP / INSEN ---
            if (mapaVIP[id]) {
                const vip = mapaVIP[id];
                if (tieneDeuda && vip.pierdeSiDebe) {
                    // Castigado: Pierde beneficio
                    montoM = montoNormalBase;
                    sufijo = "";
                } else {
                    // Salvado: Mantiene cuota especial
                    montoM = vip.cuota;
                    sufijo = " (Especial)"; // Diferenciador visual
                }
            }

            let conceptoM = `Mensualidad ${txtMes}${sufijo}`;
            
            // Aplicar saldo a favor
            let {estado, montoFinal, pagoRef, saldoRestante} = aplicarSaldo(id, montoM, mapaSaldos, saldosSheet);
            if(pagoRef !== "") { contadorAbonos++; conceptoM += ` (${pagoRef})`; }
            
            nuevosCargos.push(['CGO-' + Utilities.getUuid().substring(0, 8).toUpperCase(), id, conceptoM, mesCorte, montoFinal, estado, pagoRef, ""]);
        }

        // --- B. GENERAR CARGOS DE ESTACIONAMIENTO ---
        rentasData.forEach(renta => {
            const idURenta = String(renta[1]);
            const montoR = Number(renta[4]) || 0;
            const estadoR = String(renta[5]).toUpperCase();
            const idRentaDoc = String(renta[0]);

            if (idURenta === id && estadoR === "ACTIVO") {
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
        return `Se generaron ${nuevosCargos.length} cargos.\nAbonos automáticos aplicados: ${contadorAbonos}`;
    } else {
        return "No hay nuevos cargos pendientes por generar en este mes.";
    }
}

// ------------------------------------------------------------------------------
// 2. CORREGIR MONTO VENCIDO V2
// ------------------------------------------------------------------------------
function corregirMontoVencido_V2(mesStr) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    const config = getConfig(); 

    if (!cargosSheet || !config) throw new Error("Faltan hojas necesarias.");

    const montoBaseNormal = Number(config.MENSUALIDAD_BASE) || 0; 
    
    const partes = mesStr.split('-');
    const anioObjetivo = parseInt(partes[0]);
    const mesObjetivo = parseInt(partes[1]) - 1; 
    
    const lastRow = cargosSheet.getLastRow();
    if (lastRow < 2) return "Base de datos vacía.";
    
    const cargosData = cargosSheet.getRange(2, 1, lastRow - 1, 6).getValues(); 
    let corregidos = 0;

    cargosData.forEach((row, index) => {
        const concepto = String(row[2]);
        const fecha = new Date(row[3]);
        const estado = String(row[5]).toUpperCase();
        const rowIndex = index + 2;

        if (concepto.includes(' PP') && fecha.getMonth() === mesObjetivo && fecha.getFullYear() === anioObjetivo && estado === "PENDIENTE") {
            cargosSheet.getRange(rowIndex, 5).setValue(montoBaseNormal);
            cargosSheet.getRange(rowIndex, 3).setValue(concepto.replace(' PP', '').trim());
            corregidos++;
        }
    });

    return `Se corrigieron y penalizaron ${corregidos} cargos vencidos a tarifa normal.`;
}

// ------------------------------------------------------------------------------
// 3. GENERAR RECARGOS POR MORA V2
// ------------------------------------------------------------------------------
function generarRecargosPorMora_V2(mesStr) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    const config = getConfig(); 
    
    if (!cargosSheet || !config) throw new Error("Faltan hojas.");

    const tasaRecargo = Number(config.TASA_RECARGO) || 0; 
    const fechaHoy = new Date();

    const partes = mesStr.split('-');
    const anioObj = parseInt(partes[0]);
    const mesObj = parseInt(partes[1]) - 1;
    
    const data = cargosSheet.getRange(2, 1, cargosSheet.getLastRow() - 1, 8).getValues(); 
    const nuevosRecargos = [];

    data.forEach((row, index) => {
        const idU = row[1];
        const fechaCorte = new Date(row[3]);
        const montoBase = Number(row[4]);
        const estado = String(row[5]).toUpperCase();
        const tieneRecargo = row[7];

        if (estado === "PENDIENTE" && !tieneRecargo && fechaCorte.getMonth() === mesObj && fechaCorte.getFullYear() === anioObj) {
            const montoRecargo = Math.round((montoBase * tasaRecargo) * 100) / 100;
            if (montoRecargo > 0) {
                const idR = 'R-' + Utilities.getUuid().substring(0, 8).toUpperCase();
                // Ponemos el mesStr legible Ej. "03/2026"
                let mesLegible = `${partes[1]}/${partes[0]}`; 
                nuevosRecargos.push([idR, idU, `Recargo Mora (${mesLegible})`, fechaHoy, montoRecargo, "Pendiente", "", ""]);
                cargosSheet.getRange(index + 2, 8).setValue(idR); 
            }
        }
    });

    if (nuevosRecargos.length > 0) {
        cargosSheet.getRange(cargosSheet.getLastRow() + 1, 1, nuevosRecargos.length, 8).setValues(nuevosRecargos);
        return `Se generaron y aplicaron ${nuevosRecargos.length} recargos con éxito.`;
    } else {
        return "No se encontraron deudas que requirieran recargo en ese mes.";
    }
}

// ------------------------------------------------------------------------------
// 4. GENERAR INTERESES MORA V2
// ------------------------------------------------------------------------------
function generarInteresesMora_V2(mesStr) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const cargosSheet = ss.getSheetByName("CARGOS_Y_DEUDAS");
    const config = getConfig(); 
    
    if (!cargosSheet || !config) throw new Error("Faltan hojas.");

    const tasaInteres = Number(config.TASA_INTERES_MORA) || 0.10; 
    const porcentajeMostrar = (tasaInteres * 100).toFixed(0); 

    if (tasaInteres <= 0) throw new Error("La Tasa de Interés en configuración es 0.");

    const fechaHoy = new Date();
    const partes = mesStr.split('-');
    const anioActual = parseInt(partes[0]);
    const mesActual = parseInt(partes[1]) - 1;
    
    // Frontera: Todo lo que sea MENOR a este mes, lleva interés.
    const fechaFrontera = new Date(anioActual, mesActual, 1); 

    const lastRow = cargosSheet.getLastRow();
    if (lastRow < 2) return "No hay datos.";
    
    const data = cargosSheet.getRange(2, 1, lastRow - 1, 8).getValues();
    const nuevosCargos = [];

    data.forEach((row) => {
        const idUnidad = row[1];
        const conceptoOriginal = row[2] ? row[2].toString() : ''; 
        const fechaCargo = new Date(row[3]);
        const montoBase = Number(row[4]) || 0;
        const estado = row[5] ? row[5].toString().toUpperCase() : '';

        if (estado !== "PENDIENTE" || fechaCargo >= fechaFrontera) return;
        if (conceptoOriginal.includes("Interés") || conceptoOriginal.includes("Recargo")) return;

        const montoInteres = Math.round((montoBase * tasaInteres) * 100) / 100;
        
        if (montoInteres > 0) {
            const idInteres = 'INT-' + Utilities.getUuid().substring(0, 6).toUpperCase();
            let mesLegible = `${partes[1]}/${partes[0]}`;
            const conceptoAmigable = `Interés ${porcentajeMostrar}% [${mesLegible}] s/ ${conceptoOriginal}`;

            nuevosCargos.push([
                idInteres, idUnidad, conceptoAmigable, fechaHoy, 
                montoInteres, "Pendiente", "", ""
            ]);
        }
    });

    if (nuevosCargos.length > 0) {
        cargosSheet.getRange(cargosSheet.getLastRow() + 1, 1, nuevosCargos.length, 8).setValues(nuevosCargos);
        return `Se generaron ${nuevosCargos.length} cargos de interés moratorio (${porcentajeMostrar}%).`;
    } else {
        return "No hay deudas anteriores vencidas que califiquen para el interés.";
    }
}