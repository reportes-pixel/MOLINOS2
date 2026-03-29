function copiarHojasEspecificas() {
  // 1. IDs de las hojas (extraídos de tus links)
  const idOrigen = "1EvZ6564r1CjZMUXQ_80dlTU1idoI7LzrqjEQWyYM-Pg";
  const idDestino = "1lMDk3HJa7paWjHz8tNj3R6wcBnAWpdQp-NE8GGNLb9A";
  
  // 2. Lista de hojas a copiar
  const nombresHojas = [
    "CONFIG_MULTAS",
    "UNIDADES",
    "CONFIGURACION",
    "USUARIOS",
    "SALDOS_A_FAVOR",
    "CARGOS_Y_DEUDAS",
    "REGISTRO_PAGOS",
    "EGRESOS"
  ];
  
  const ssOrigen = SpreadsheetApp.openById(idOrigen);
  const ssDestino = SpreadsheetApp.openById(idDestino);
  
  nombresHojas.forEach(nombre => {
    const hojaCopiada = ssOrigen.getSheetByName(nombre);
    
    if (hojaCopiada) {
      // Si la hoja ya existe en el destino, le ponemos un nombre temporal o la borramos
      const hojaVieja = ssDestino.getSheetByName(nombre);
      if (hojaVieja) {
        ssDestino.deleteSheet(hojaVieja);
      }
      
      // Copiar la hoja al destino
      const nuevaHoja = hojaCopiada.copyTo(ssDestino);
      nuevaHoja.setName(nombre); // Quitar el "Copia de..."
      
      console.log("Copiada con éxito: " + nombre);
    } else {
      console.warn("No se encontró la hoja: " + nombre);
    }
  });
  
  SpreadsheetApp.getUi().alert("¡Proceso completado con éxito!");
}