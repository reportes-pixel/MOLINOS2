// ==============================================================================
// 08_COMUNICACIONES.gs - TICKET HTML Y ENVÍO POR CORREO
// ==============================================================================

function showTicketDialog(idPago) {
    const result = generatePaymentTicketHtml(idPago);
    if (result.error) {
        SpreadsheetApp.getUi().alert("Error al generar el ticket: " + result.error);
        return;
    }
    const htmlOutput = HtmlService.createHtmlOutput(result.html).setWidth(380).setHeight(550);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Ticket de Pago');
}

function reimprimirTicket(idPago) {
    const result = generatePaymentTicketHtml(idPago); 
    if (!result.success) {
        Browser.msgBox('Error al generar ticket', result.error, Browser.Buttons.OK);
        return;
    }
    const htmlOutput = HtmlService.createHtmlOutput(result.html).setWidth(380).setHeight(550);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Ticket de Pago ' + idPago);
}

function closeTicketAndReopenForm() {
    showPaymentForm(); 
}

// ⭐️ NUEVO: MOTOR DE ENVÍO DE CORREO
function enviarTicketPorCorreo(idPago, emailDestino) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheetPagos = ss.getSheetByName("REGISTRO_PAGOS");
        const data = sheetPagos.getDataRange().getValues();
        let idUnidad = "";
        
        for (let i = data.length - 1; i >= 1; i--) {
            if (String(data[i][0]) === String(idPago)) {
                idUnidad = String(data[i][2]).trim();
                break;
            }
        }
        
        if(idUnidad !== "") {
            guardarEmailUnidad(idUnidad, emailDestino); // Lo guarda en UNIDADES para la próxima vez
        }

        // Generamos el HTML versión "Solo Lectura" (sin botones)
        const result = generatePaymentTicketHtml(idPago, true); 
        if (!result.success) return { success: false, message: result.error };

        MailApp.sendEmail({
            to: emailDestino,
            subject: "Comprobante de Pago - Privada Molinos Real II",
            htmlBody: result.html
        });

        return { success: true, message: "¡Ticket enviado exitosamente a " + emailDestino + "!" };
    } catch (e) {
        return { success: false, message: "Error al enviar correo: " + e.toString() };
    }
}

// ⭐️ EL TICKET ACTUALIZADO (Acepta modo Correo y modo Impresión)
function generatePaymentTicketHtml(idPago, esParaCorreo = false) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPagos = ss.getSheetByName("REGISTRO_PAGOS");
    if (!sheetPagos) return { success: false, error: "Hoja no encontrada." };

    const data = sheetPagos.getDataRange().getValues();
    let paymentRecord = null;
    
    for (let i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]) === String(idPago)) {
            paymentRecord = {
                idPago: data[i][0],
                fechaPago: Utilities.formatDate(new Date(data[i][1]), ss.getSpreadsheetTimeZone(), "dd/MM/yyyy HH:mm:ss"),
                idUnidad: String(data[i][2]).trim(),
                capturista: data[i][3],
                montoRecibido: parseFloat(data[i][4]) || 0,
                montoAplicado: parseFloat(data[i][5]) || 0,
                saldoFinal: parseFloat(data[i][7]) || 0, 
                cargosIds: data[i][8] 
            };
            break;
        }
    }
    
    if (!paymentRecord) return { success: false, error: `No se encontró el pago.` };

    const unitDetails = getUnitDetails(paymentRecord.idUnidad); 
    const nombreDepto = unitDetails ? unitDetails.departamento : "N/A";
    const nombrePropietario = unitDetails ? unitDetails.propietario : "N/A";
    const emailRegistrado = unitDetails ? unitDetails.email : "";

    const adeudos = getDebtsForUnit(paymentRecord.idUnidad);
    
    let deudaRestanteHtml = "";
    if (adeudos.filteredDeudas.length > 0) {
        deudaRestanteHtml = adeudos.filteredDeudas.map(d => `<p style="margin-bottom: 4px;"><span>${d.concepto}</span> <span style="float: right;">$${d.montoBase.toFixed(2)}</span></p>`).join('');
    } else {
        deudaRestanteHtml = `<p class="center" style="font-style: italic; margin-bottom: 5px;">Sin adeudos pendientes</p>`;
    }

    const conceptosPagadosHtml = getChargeConceptsByIds(paymentRecord.cargosIds);

    // ⭐️ BLOQUE DE BOTONES (Se oculta si se envía por correo)
    let controlesHtml = ``;
    if (!esParaCorreo) {
        controlesHtml = `
        <div class="print-controls" style="padding-top: 15px; display:flex; flex-direction:column; gap:8px;">
            <button class="control-button" onclick="window.print();" style="background-color: #4CAF50; padding:10px;">🖨️ Imprimir Ticket</button>
            <button class="control-button" onclick="pedirYEnviarCorreo('${idPago}', '${emailRegistrado}')" style="background-color: #1a73e8; padding:10px;">✉️ Enviar por Correo</button>
            <button class="control-button" onclick="google.script.run.closeTicketAndReopenForm(); window.close();" style="background-color: #f44336; padding:10px;">Cerrar y Volver</button>
            <div id="email-status" style="text-align:center; font-size:12px; color:#1a73e8; font-weight:bold;"></div>
        </div>
        <script>
            function pedirYEnviarCorreo(id, correoExistente) {
                let email = prompt("Ingrese el correo electrónico para enviar el ticket:", correoExistente);
                if(email && email.includes('@')) {
                    document.getElementById('email-status').innerText = "Enviando correo...";
                    google.script.run.withSuccessHandler(res => {
                        alert(res.message);
                        document.getElementById('email-status').innerText = res.success ? "¡Enviado!" : "";
                    }).enviarTicketPorCorreo(id, email);
                } else if(email) { alert("Correo no válido."); }
            }
        </script>`;
    }

    const html = `
    <!DOCTYPE html>
    <html>
    <head>
        <base target="_top">
        <style>
            body { width: 170px; margin: 0 auto; padding: 2px; font-family: monospace; font-size: 8pt; line-height: 1.2; color: #000; }
            .ticket { width: 100%; margin: 0; }
            h3, p { margin: 0; padding: 1px 0; }
            .divider { border-top: 1px dashed #000; margin: 4px 0; }
            .strong { font-weight: bold; }
            .center { text-align: center; }
            .control-button { width: 100%; color: white; border: none; cursor: pointer; font-weight: bold; border-radius: 4px;}
            @media print { .print-controls { display: none; } @page { size: 48mm auto; margin: 0 !important; } }
        </style>
    </head>
    <body>
        <div class="ticket">
            <h3 class="center">COMPROBANTE DE PAGO</h3>
            <p class="center strong">MESA DIRECTIVA</p>
            <p class="center strong">PRIVADA MOLINOS REAL II</p>
            <div class="divider"></div>
            
            <p><span class="strong">Fecha:</span> <span style="float: right;">${paymentRecord.fechaPago}</span></p>
            <p><span class="strong">Unidad:</span> ${paymentRecord.idUnidad}</p>
            <p><span class="strong">Depto:</span> ${nombreDepto}</p>
            <p><span class="strong">Propietario:</span> ${nombrePropietario}</p>
            <p><span class="strong">Capturista:</span> ${paymentRecord.capturista}</p>
            
            <div class="divider"></div>
            <p class="strong" style="margin-bottom: 5px; text-decoration: underline;">Conceptos Pagados/Abonados:</p>
            <div style="font-size: 8pt; margin-bottom: 5px;">
                ${conceptosPagadosHtml}
            </div>
            
            <div class="divider"></div>
            <p><span class="strong">Monto Recibido:</span> <span style="float: right;">$${paymentRecord.montoRecibido.toFixed(2)}</span></p>
            <p style="margin-bottom: 3px;"><span class="strong">Monto Aplicado:</span> <span style="float: right;">$${paymentRecord.montoAplicado.toFixed(2)}</span></p>
            
            <div class="divider"></div>
            <p class="strong" style="margin-bottom: 5px; text-decoration: underline;">Deuda Restante de la Unidad:</p>
            <div style="font-size: 8pt;">
                ${deudaRestanteHtml}
            </div>
            <p class="strong" style="margin-top: 6px; border-top: 1px solid #000; padding-top: 4px;">
                Total Pendiente: <span style="float: right;">$${adeudos.totalPendiente.toFixed(2)}</span>
            </p>
            
            <div class="divider"></div>
            <div style="margin-top: 10px; margin-bottom: 10px;">
                <p class="strong">SALDO A FAVOR (SIN APLICAR):</p>
                <p style="text-align: right; font-size: 11pt; font-weight: bold;">$${paymentRecord.saldoFinal.toFixed(2)}</p>
            </div>
            <div class="divider"></div>
            
            <p><span class="strong">ID Transacción:</span> ${paymentRecord.idPago}</p>
            <p class="center" style="margin-top: 15px;">¡Gracias por su pago!</p>
        </div>
        ${controlesHtml}
    </body>
    </html>
    `;
    
    return { success: true, html: html };
}

function getChargeConceptsByIds(cargoIdsString) {
    if (!cargoIdsString || cargoIdsString === '') return `<p style="text-align: center; font-style: italic;">ANTICIPO o Sin Deudas</p>`;
    const sheetCargos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("CARGOS_Y_DEUDAS");
    if (!sheetCargos) return `<p style="color: red;">Error: Hoja Cargos no encontrada.</p>`;

    const cargoIds = cargoIdsString.split(',').map(id => id.trim()).filter(Boolean);
    const data = sheetCargos.getDataRange().getValues();
    const chargeLines = [];

    cargoIds.forEach(idBuscado => {
        for (let i = 1; i < data.length; i++) {
            if (String(data[i][0]) === idBuscado) {
                let concepto = data[i][2] || `Cargo sin concepto`;
                const monto = parseFloat(data[i][4]) || 0;
                
                if (concepto.includes('Mensualidad')) concepto = concepto.replace('Mensualidad', 'Mens.');
                else if (concepto.includes('Recargo Mora')) concepto = concepto.replace('Recargo Mora', 'Recargo');
                else if (concepto.includes('Deuda') || concepto.includes('Extraordinario')) concepto = "Extr. (" + concepto.split(' ')[1] + ")"; 
                else if (concepto.includes('Multa:')) concepto = concepto.replace(':', '');
                
                if (concepto.length > 20) concepto = concepto.substring(0, 18) + '...';
                
                chargeLines.push(`<p style="margin: 0; padding: 1px 0;">${concepto}<span style="float: right;">$${monto.toFixed(2)}</span></p>`);
                return;
            }
        }
    });
    return chargeLines.join('');
}