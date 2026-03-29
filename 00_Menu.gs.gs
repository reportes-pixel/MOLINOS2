// ==============================================================================
// 00_MENU.gs - FUNCIONES DE ACTIVACIÓN, MENÚ Y SEGURIDAD
// ==============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu('💳 Administración Condominal')
    // ===== MÓDULOS BÁSICOS =====
    .addItem('Registrar Pago de Cuotas', 'accesoRegistrarPagos')
    .addItem('Adelantar Mensualidades 🌟', 'accesoRegistrarAdelanto') // NUEVO
    .addItem('Registrar Multa a Unidad', 'accesoRegistrarMulta')
    .addItem('Registrar Egresos 🔒', 'accesoRegistrarEgresos')
    .addSeparator()

    // ===== SUBMENÚ: GENERACIÓN DE CARGOS (SOLO ADMIN) =====
    .addSubMenu(
      ui.createMenu('Generación de Cargos 🔒')
        .addItem('Generar Mensualidades (Smart PP)', 'accesoGenerarCargosMensuales')
        .addItem('Corregir Monto Vencido (Días 6-10)', 'accesoCorregirMontoVencido')
        .addItem('Aplicar Recargos por Mora', 'accesoGenerarRecargos')
        .addItem('Generar Intereses por Atraso (10%) ⚠️', 'accesoGenerarIntereses') // NUEVO
    )
    .addSeparator()

    // ===== SUBMENÚ: REPORTES =====
    .addSubMenu(
      ui.createMenu('Reportes 📊 🔒')
        .addItem('1. Reporte Financiero (Flujo Extendido)', 'accesoReporteFinanciero')
        .addItem('2. Saldo Detallado (Por Concepto/Fechas)', 'accesoReporteSaldoDetallado')
        .addItem('3. Estado de Cuenta (Cargos/Pagos Detallado)', 'accesoEstadoCuenta')
        .addItem('4. Mensualidades Vencidas (Lista)', 'accesoMensualidadesVencidas')
        .addSeparator()
        .addItem('5. Total de Deudores (Ranking)', 'accesoTotalDeudores')
        .addItem('6. Reporte Financiero (Flujo Extendido)', 'accesoReporteFinanciero2')
    )
    .addSeparator()

    // ===== SUBMENÚ: HERRAMIENTAS Y EXTRAS =====
    .addSubMenu(
      ui.createMenu('Deudas Extra y Estacionamiento')
        .addItem('Generar Cargo Extraordinario', 'mostrarFormularioCargoExtra')
        .addItem('Gestionar Estacionamientos 🚗', 'mostrarFormularioEstacionamiento') // NUEVO
    )
    .addSeparator()

    // ===== ADMINISTRACIÓN =====
    .addSubMenu(
      ui.createMenu('🔒 Herramientas de Administrador')
        .addItem('Ocultar/Mostrar Hojas Sensibles', 'toggleSheets')
        .addItem('Anular Pago (Reversión)', 'abrirCanceladorAdmin')
    )
    .addToUi();
}

// --- SISTEMA DE ROLES Y PERMISOS ---
function obtenerRolUsuario() {
  const ui = SpreadsheetApp.getUi();
  const CREDENCIALES = {
    "Super25": "ADMIN",
    "CP25A": "CONTADOR",
    "KING25": "PRESIDENTE",
    "PAGOS": "CAPTURISTA"
  };

  const prompt = ui.prompt("Acceso al Sistema", "Introduce tu contraseña", ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return null;
  return CREDENCIALES[prompt.getResponseText().trim()] || null;
}

function validarPermiso(rolesPermitidos) {
  const rol = obtenerRolUsuario();
  if (!rol) {
    SpreadsheetApp.getUi().alert("Acceso Denegado", "Credenciales incorrectas.", SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
  if (!rolesPermitidos.includes(rol)) {
    SpreadsheetApp.getUi().alert("Acceso Restringido", "No tienes permiso para acceder a este módulo.", SpreadsheetApp.getUi().ButtonSet.OK);
    return false;
  }
  return true;
}

// --- ENLACES A FORMULARIOS (CON PERMISOS) ---
function accesoRegistrarPagos() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE", "CAPTURISTA"])) showPaymentForm(); }
function accesoRegistrarAdelanto() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE", "CAPTURISTA"])) showAdelantoForm(); }
function accesoRegistrarMulta() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE"])) showFineForm(); }
function accesoRegistrarEgresos() { if (validarPermiso(["ADMIN", "CONTADOR"])) showExternalDebtForm(); }

function accesoGenerarCargosMensuales() { if (validarPermiso(["ADMIN"])) generarCargosMensuales(); }
function accesoCorregirMontoVencido() { if (validarPermiso(["ADMIN"])) corregirMontoVencido(); }
function accesoGenerarRecargos() { if (validarPermiso(["ADMIN"])) generarRecargosPorMora(); }
function accesoGenerarIntereses() { if (validarPermiso(["ADMIN"])) generarInteresesMora(); } // NUEVO

function accesoReporteFinanciero() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE"])) generarReporteFinanciero_EXTENDIDO(); }
function accesoReporteSaldoDetallado() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE"])) reporteSaldoDetalladoPorConcepto_Fechas(); }
function accesoEstadoCuenta() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE"])) reporteEstadoDeCuenta_SEPARADO(); }
function accesoMensualidadesVencidas() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE"])) reporteMensualidadesVencidas_CORREGIDO(); }
function accesoTotalDeudores() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE"])) reporteTotalDeudores(); }
function accesoReporteFinanciero2() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE"])) generarReporteFinanciero_CON_DETALLE(); }

// --- APERTURA DE DIÁLOGOS HTML ---
function showPaymentForm() {
  const htmlOutput = HtmlService.createTemplateFromFile('Form_Pagos').evaluate().setWidth(500).setHeight(650);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Registro de Pagos de Unidad');
}
function showFineForm() {
  const htmlOutput = HtmlService.createTemplateFromFile('Form_Multas').evaluate().setWidth(450).setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Registro de Multa a Unidad');
}
function showExternalDebtForm() {
  const htmlOutput = HtmlService.createTemplateFromFile('Form_Egresos').evaluate().setWidth(550).setHeight(480);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Registro de Egresos'); 
}
function mostrarFormularioCargoExtra() {
  const htmlOutput = HtmlService.createTemplateFromFile('Form_CargosExtra').evaluate().setWidth(400).setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Generar Cargo Extraordinario');
}
function abrirCanceladorAdmin() {
  const htmlOutput = HtmlService.createHtmlOutputFromFile('Form_Cancelacion').setWidth(450).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, "Panel de Control Administrativo");
}
function showAdelantoForm() {
  const htmlOutput = HtmlService.createTemplateFromFile('Form_Adelantos').evaluate().setWidth(500).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Adelanto de Mensualidades');
}
function mostrarFormularioEstacionamiento() {
  const htmlOutput = HtmlService.createTemplateFromFile('Form_Estacionamiento').evaluate().setWidth(450).setHeight(450);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Suscripción de Estacionamiento');
}