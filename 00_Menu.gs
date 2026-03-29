// ==============================================================================
// 00_MENU.gs - FUNCIONES DE ACTIVACIÓN, MENÚ Y SEGURIDAD
// ==============================================================================

function onOpen() {
  const ui = SpreadsheetApp.getUi();

  // 1. Crea el menú súper limpio como "Plan B"
  ui.createMenu('🌟 SISTEMA MOLINOS')
    .addItem('Abrir Panel de Control 🚀', 'abrirSidebarPrincipal')
    .addToUi();

  // 2. ¡AQUÍ ESTÁ LA MAGIA! Abre el panel lateral automáticamente al cargar
  try {
    abrirSidebarPrincipal();
  } catch (e) {
    // Si por temas de permisos Google tarda en reaccionar, no marca error.
  }
}


function abrirSidebarPrincipal() {
  const html = HtmlService.createTemplateFromFile('Sidebar')
      .evaluate()
      .setTitle('Sistema Molinos Real II')
      .setWidth(300); // El Sidebar tiene un ancho fijo estándar
      
  SpreadsheetApp.getUi().showSidebar(html);
}

// --- SISTEMA DE ROLES Y PERMISOS (CON CACHÉ INTEGRADO) ---
function obtenerRolUsuario() {
  const cache = CacheService.getUserCache();
  const rolGuardado = cache.get('rolActivo');
  
  // Si ya se logueó desde el Sidebar (o antes en el menú), lo deja pasar directo
  if (rolGuardado) return rolGuardado;

  const ui = SpreadsheetApp.getUi();
  const CREDENCIALES = {
    "Super25": "ADMIN",
    "CP25A": "CONTADOR",
    "KING25": "PRESIDENTE",
    "PAGOS": "CAPTURISTA"
  };

  const prompt = ui.prompt("Acceso al Sistema", "Introduce tu contraseña", ui.ButtonSet.OK_CANCEL);
  if (prompt.getSelectedButton() !== ui.Button.OK) return null;
  
  const rol = CREDENCIALES[prompt.getResponseText().trim()] || null;
  
  // Si puso bien la contraseña, guarda la sesión por 6 horas (21600 segundos)
  if (rol) cache.put('rolActivo', rol, 21600);
  
  return rol;
}

// ESTA ES LA FUNCIÓN QUE SE CONECTA CON TU SIDEBAR AL HACER CLIC EN "INICIAR SESIÓN"
function validarLoginLogin(user, pass) {
  const CREDENCIALES = {
    "Super25": "ADMIN",
    "CP25A": "CONTADOR",
    "KING25": "PRESIDENTE",
    "PAGOS": "CAPTURISTA"
  };
  
  const rol = CREDENCIALES[pass];
  if (rol) {
    // Al entrar desde el Sidebar, también guarda el rol en la caché para el Menú Superior
    CacheService.getUserCache().put('rolActivo', rol, 21600); 
    return { success: true, nombre: user, rol: rol };
  } else {
    return { success: false, message: "Contraseña / PIN incorrecto" };
  }
}

// (Opcional) Esta función la puedes llamar al cerrar sesión en el Sidebar
function limpiarCacheSesion() {
  CacheService.getUserCache().remove('rolActivo');
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

// --- ENLACES A FORMULARIOS Y REPORTES (CON PERMISOS CORREGIDOS) ---

// 1. COBRANZA
function accesoRegistrarPagos() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE", "CAPTURISTA"])) showPaymentForm(); }
function accesoRegistrarAdelanto() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE", "CAPTURISTA"])) showAdelantoForm(); }
function accesoCorteDiario() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE", "CAPTURISTA"])) generarCortePagos(); }
function accesoEstadoCuenta() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE", "CAPTURISTA"])) reporteEstadoDeCuenta_SEPARADO(); }

// 2. REPORTES
function accesoReporteSaldoDetallado() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE"])) reporteSaldoDetalladoPorConcepto_Fechas(); }
function accesoTotalDeudores() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE"])) reporteTotalDeudores(); }
function accesoMensualidadesVencidas() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE"])) reporteMensualidadesVencidas_CORREGIDO(); }
function accesoReporteDeudaReal() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE"])) reporteDeudaRealCorregido(); }
function accesoCortePeriodo() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE"])) generarCorteVariosDias(); }

// 3. GENERACIÓN DE CARGOS Y CONTROL
function accesoGenerarCargosMensuales() { if (validarPermiso(["ADMIN"])) generarCargosMensuales(); }
function accesoCorregirMontoVencido() { if (validarPermiso(["ADMIN", "CONTADOR"])) corregirMontoVencido(); }
function accesoGenerarRecargos() { if (validarPermiso(["ADMIN", "CONTADOR"])) generarRecargosPorMora(); }
function accesoGenerarIntereses() { if (validarPermiso(["ADMIN", "CONTADOR"])) generarInteresesMora(); }
function accesoRegistrarMulta() { if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE"])) showFineForm(); }
function accesoCargoExtra() { if (validarPermiso(["ADMIN", "CONTADOR"])) mostrarFormularioCargoExtra(); }
function accesoEstacionamiento() { if (validarPermiso(["ADMIN", "CONTADOR"])) mostrarFormularioEstacionamiento(); }

// 4. ADMINISTRACIÓN Y SEGURIDAD
function accesoRegistrarEgresos() { if (validarPermiso(["ADMIN", "CONTADOR"])) showExternalDebtForm(); }
function accesoCancelarPago() { if (validarPermiso(["ADMIN", "CONTADOR"])) abrirCanceladorAdmin(); }
function accesoToggleBD() { if (validarPermiso(["ADMIN"])) toggleSheets(); }
function accesoGestionUsuarios() { if (validarPermiso(["ADMIN"])) mostrarGestorUsuarios(); }


// ==============================================================================
// APERTURA DE DIÁLOGOS HTML (¡INTACTOS Y COMPLETOS!)
// ==============================================================================

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
  const htmlOutput = HtmlService.createTemplateFromFile('CargoExtraordinario').evaluate().setWidth(400).setHeight(300);
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

// NUEVO: APERTURA DEL REPORTE FINANCIERO MAESTRO
function accesoGeneradorFinancieroMaster() {
  if (validarPermiso(["ADMIN", "CONTADOR", "PRESIDENTE"])) {
    const htmlOutput = HtmlService.createTemplateFromFile('Form_ReporteFinanciero')
        .evaluate()
        .setWidth(450)
        .setHeight(480);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Generador Financiero Maestro');
  }
} // <--- ESTA LLAVE FALTABA EN TU CÓDIGO ORIGINAL

// --- APERTURA DEL NUEVO GESTOR MAESTRO DE CARGOS ---
function accesoGestorCargosMaestro() {
  const htmlOutput = HtmlService.createTemplateFromFile('Form_GestorCargos')
      .evaluate()
      .setWidth(700)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Gestor Maestro de Cargos');
}