// ==============================================================================
// 10_LOGIN_USUARIOS.gs - GESTIÓN DE SEGURIDAD, ROLES Y ACCESOS
// ==============================================================================

/**
 * Función de inicialización. Ejecútala UNA VEZ manualmente para crear la hoja.
 */
function inicializarBaseUsuarios() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheetUser = ss.getSheetByName("USER");
    
    if (!sheetUser) {
        sheetUser = ss.insertSheet("USER");
        // Definir columnas (A-E)
        const headers = ["ID_USUARIO", "NOMBRE_MOSTRAR", "PASSWORD", "ROL", "ESTADO"];
        sheetUser.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight("bold").setBackground("#d9ead3");
        
        // Crear el Super Administrador inicial (Obligatorio para no quedar fuera)
        const idAdmin = 'USR-' + Utilities.getUuid().substring(0, 6).toUpperCase();
        sheetUser.appendRow([idAdmin, "Super Admin", "Super25", "ADMIN", "ACTIVO"]);
        
        // Congelar la fila superior y ajustar columnas
        sheetUser.setFrozenRows(1);
        sheetUser.autoResizeColumns(1, headers.length);
        
        SpreadsheetApp.getUi().alert("Éxito", "Hoja 'USER' creada y Super Admin configurado con la clave 'Super25'.", SpreadsheetApp.getUi().ButtonSet.OK);
    } else {
        SpreadsheetApp.getUi().alert("Aviso", "La hoja 'USER' ya existe. No se hicieron cambios.", SpreadsheetApp.getUi().ButtonSet.OK);
    }
}

/**
 * Función que usará el Sidebar y la WebApp para validar credenciales.
 * Retorna el perfil completo del usuario si tiene éxito.
 */
function validarLoginLogin(nombreUsuario, password) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetUser = ss.getSheetByName("USER");
    
    if (!sheetUser) return { success: false, message: "Error interno: Base de usuarios no encontrada." };
    
    // Normalizar entradas para evitar errores por espacios o mayúsculas en el nombre
    const userClean = String(nombreUsuario).trim().toLowerCase();
    const passClean = String(password).trim();
    
    const data = sheetUser.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
        const dbName = String(data[i][1]).trim().toLowerCase(); // Col B: NOMBRE_MOSTRAR
        const dbPass = String(data[i][2]).trim();               // Col C: PASSWORD
        const dbEstado = String(data[i][4]).trim().toUpperCase(); // Col E: ESTADO
        
        // Si coinciden nombre y contraseña
        if (dbName === userClean && dbPass === passClean) {
            // Verificar si la cuenta está suspendida
            if (dbEstado !== "ACTIVO") {
                return { success: false, message: "Tu cuenta ha sido desactivada. Contacta al administrador." };
            }
            
            // Login exitoso: Devolvemos sus datos (sin la contraseña, por seguridad)
            return {
                success: true,
                idUsuario: data[i][0],
                nombre: data[i][1],
                rol: String(data[i][3]).toUpperCase(),
                message: "¡Bienvenido, " + data[i][1] + "!"
            };
        }
    }
    
    // Si terminó el ciclo y no encontró coincidencias
    return { success: false, message: "Usuario o contraseña incorrectos." };
}

/**
 * Utilidad rápida para el formulario: Obtiene solo los nombres de los usuarios ACTIVOS
 * (Para que los seleccionen en un Dropdown en lugar de escribirlos)
 */
function obtenerNombresUsuariosActivos() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USER");
    if (!sheet) return [];
    
    const data = sheet.getDataRange().getValues();
    const activos = [];
    
    for (let i = 1; i < data.length; i++) {
        if (String(data[i][4]).trim().toUpperCase() === "ACTIVO") {
            activos.push(data[i][1]); // Guardamos el nombre
        }
    }
    return activos.sort(); // Los devolvemos ordenados alfabéticamente
}

// --- FUNCIONES PARA EL GESTOR DE USUARIOS ---

function mostrarGestorUsuarios() {
    const htmlOutput = HtmlService.createTemplateFromFile('Form_Usuarios').evaluate().setWidth(600).setHeight(550);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Gestor de Usuarios y Permisos');
}

function getUsuariosData() {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USER");
    if (!sheet) return { success: false, message: "La hoja USER no existe." };
    
    const data = sheet.getDataRange().getValues();
    const users = [];
    
    // Saltamos la fila 1 (encabezados)
    for (let i = 1; i < data.length; i++) {
        if(data[i][0]) {
            users.push({
                id: data[i][0],
                nombre: data[i][1],
                password: data[i][2],
                rol: data[i][3],
                estado: data[i][4]
            });
        }
    }
    return { success: true, users: users };
}

function guardarUsuario(userData) {
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("USER");
        const data = sheet.getDataRange().getValues();
        
        // MODO EDICIÓN
        if (userData.id) {
            for (let i = 1; i < data.length; i++) {
                if (data[i][0] === userData.id) {
                    sheet.getRange(i + 1, 2).setValue(userData.nombre);
                    sheet.getRange(i + 1, 3).setValue(userData.password);
                    sheet.getRange(i + 1, 4).setValue(userData.rol);
                    sheet.getRange(i + 1, 5).setValue(userData.estado);
                    return { success: true, message: "Usuario actualizado correctamente." };
                }
            }
        }
        
        // MODO NUEVO
        const newId = 'USR-' + Utilities.getUuid().substring(0, 6).toUpperCase();
        sheet.appendRow([newId, userData.nombre, userData.password, userData.rol, userData.estado]);
        return { success: true, message: "Nuevo usuario creado con éxito." };
        
    } catch (e) {
        return { success: false, message: "Error al guardar: " + e.message };
    }
}