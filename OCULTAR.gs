/**
 * Oculta o muestra las hojas de configuración y datos maestros.
 * Requiere una contraseña para su ejecución.
 */
function toggleSheets() {
    const PASSWORD_REQUERIDA = "Super25";
    
    // Lista de las hojas sensibles que queremos proteger
    const NOMBRES_HOJAS_SENSIBLES = [
        "CONFIG_MULTAS",
        "UNIDADES",
        "CONFIGURACION",
        "USUARIOS",
        "SALDOS_A_FAVOR",
        "CARGOS_Y_DEUDAS",
        "REGISTRO_PAGOS",
        "EGRESOS"
    ];

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const ui = SpreadsheetApp.getUi();

    // 1. Pedir Contraseña
    const prompt = ui.prompt(
        "Acceso Restringido",
        "Por favor, introduce la contraseña para ocultar/mostrar las hojas sensibles.",
        ui.ButtonSet.OK_CANCEL
    );

    if (prompt.getSelectedButton() !== ui.Button.OK) {
        // El usuario canceló
        return;
    }

    const passwordIngresada = prompt.getResponseText();

    // 2. Verificar Contraseña
    if (passwordIngresada !== PASSWORD_REQUERIDA) {
        ui.alert("Acceso Denegado", "Contraseña incorrecta. No se realizaron cambios.", ui.ButtonSet.OK);
        return;
    }

    // 3. Determinar la acción (Mostrar u Ocultar)
    // Buscamos la primera hoja sensible. Si está oculta, el objetivo será MOSTRAR.
    let targetSheet = ss.getSheetByName(NOMBRES_HOJAS_SENSIBLES[0]);
    
    // Si la hoja no existe o está oculta, el objetivo es MOSTRAR
    const isHidden = targetSheet && targetSheet.isSheetHidden(); 
    const accion = isHidden ? "mostrando" : "ocultando";

    let contador = 0;

    // 4. Iterar y aplicar la acción
    NOMBRES_HOJAS_SENSIBLES.forEach(sheetName => {
        let sheet = ss.getSheetByName(sheetName);
        if (sheet) {
            if (isHidden) {
                // Si estaban ocultas, las mostramos
                sheet.showSheet();
            } else {
                // Si estaban visibles, las ocultamos
                sheet.hideSheet();
            }
            contador++;
        }
    });

    // 5. Mostrar Resultado
    ui.alert("Éxito", `Se han terminado de ${accion} ${contador} hojas.`, ui.ButtonSet.OK);
}