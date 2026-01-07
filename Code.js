const SHEET_MAESTRA_NAME = "Maestra";
const SHEET_INVENTARIO_NAME = "Inventario";
const SHEET_SOLICITUDES_NAME = "Solicitud";

/**
 * Función principal para desplegar la aplicación web.
 */
function doGet() {
  const htmlTemplate = HtmlService.createTemplateFromFile('Index');
  const output = htmlTemplate.evaluate()
    .setTitle('Solicitud de Uniformes')
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  // CRUCIAL para el responsive en móvil
  output.addMetaTag('viewport', 'width=device-width, initial-scale=1'); 
  
  return output;
}

/**
 * Función de 'login' que valida si la cédula existe en la Maestra.
 * @param {string} cedula - Cédula ingresada por el usuario.
 * @returns {object|null} - Objeto con datos del colaborador o null si no existe.
 */
function validateLogin(cedula) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_MAESTRA_NAME);
  const data = sheet.getRange("A2:Z" + sheet.getLastRow()).getValues(); // Asume Cédula, Nombre, Sexo están en A, B, C

  const userData = data.find(row => String(row[0]) === String(cedula)); // [0]=Cedula, [1]=Nombre, [2]=Sexo
  
  if (userData) {
    return {
      cedula: userData[0],
      nombre: userData[1],
      sexo: userData[2],
      cargo: userData[6],
      ceco: userData[10],
      rol: userData[12]
      // Agrega aquí los datos necesarios para el historial o validaciones posteriores
    };
  }
  return null;
}

/**
 * Obtiene el historial de solicitudes y normaliza los datos.
 * @param {string} cedula - Cédula del colaborador.
 * @returns {Array<Array<string>>} - Array de solicitudes con todos los valores como strings.
 */
function getSolicitudes(cedula) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_SOLICITUDES_NAME); 
    
    if (!sheet || sheet.getLastRow() <= 1) { 
      return []; 
    }

    const data = sheet.getDataRange().getValues(); 
    const CEDULA_COL_INDEX = 2; // Asume columna C (indice 2)
    
    const solicitudesFiltradas = data.filter((row, index) => {
      if (index === 0) return false; 
      return String(row[CEDULA_COL_INDEX]) === String(cedula);
    });

    // *** PASO CRÍTICO: Normalizar los datos a strings ***
    const solicitudesNormalizadas = solicitudesFiltradas.map(row => {
      return row.map((cell, index) => {
        if (cell instanceof Date) {
          // Normaliza el objeto Date a un string ISO simple (el cliente lo volverá a formatear)
          return cell.toISOString();
        }
        if (cell === null || cell === undefined) {
          return ""; // Reemplaza null/undefined con string vacío
        }
        return String(cell); // Convierte todo lo demás (incluidos números) a string
      });
    });
    
    // Logger.log('Datos normalizados y listos para enviar al cliente: ' + solicitudesNormalizadas);
    return solicitudesNormalizadas; 

  } catch (e) {
    Logger.log("Error catastrófico en getSolicitudes: " + e.toString());
    return []; 
  }
}


/**
 * Obtiene la lista dinámica de uniformes basada en el sexo y unisex.
 * @param {string} sexo - Sexo del colaborador ('M' o 'F').
 * @returns {Array<string>} - Lista de nombres de uniformes disponibles.
 */
function getUniformesBySexo(sexo, rol) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_INVENTARIO_NAME);
  const data = sheet.getDataRange().getValues(); 
  
  // Asume que la hoja de Inventario tiene (Columna Uniforme | Columna Sexo Uniforme)
  const uniformes = data.filter((row, index) => {
    if (index === 0) return false; // Ignorar encabezado
    const uniformeSexo = String(row[1]).toUpperCase(); // Asume Sexo Uniforme en columna B
    const uniformeRol = String(row[4]).toUpperCase(); // Asume Rol Uniforme en columna E
    return ((uniformeSexo == 'UNISEX' || uniformeSexo === sexo.toUpperCase())&& uniformeRol == rol.toUpperCase());
  }).map(row => row[2]); // Asume Nombre Uniforme en columna A
  Logger.log(uniformes)
  return [...new Set(uniformes)]; // Elimina duplicados
}

/**
 * Obtiene las tallas disponibles para un uniforme específico.
 * @param {string} nombreUniforme - Nombre del uniforme.
 * @returns {Array<string>} - Lista de tallas disponibles.
 */
function getTallasByUniforme(nombreUniforme) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_INVENTARIO_NAME);
  const data = sheet.getDataRange().getValues();
  
  // Asume (Columna Uniforme | Columna Talla | Columna Stock)
  const tallas = data.filter((row, index) => {
    if (index === 0) return false;
    // Asume Uniforme en Columna A, Stock en Columna C
    return String(row[2]) === nombreUniforme; // Solo si hay stock
  }).map(row => row[3]); // Asume Talla en Columna C
  
  return [...new Set(tallas)]; 
}

/**
 * Obtiene el código del uniforme y la talla seleccionada.
 * @param {string} nombreUniforme - Nombre del uniforme.
 * @param {string} talla - Talla seleccionada.
 * @returns {string} - Código del uniforme.
 */
function getCodigoUniforme(nombreUniforme, talla) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_INVENTARIO_NAME);
  const data = sheet.getDataRange().getValues();

  // Asume (Columna Uniforme | Columna Talla | Columna Código)
  const rowData = data.find((row, index) => {
    if (index === 0) return false;
    // Asume Uniforme en Columna A, Talla en Columna C, Código en Columna D
    return String(row[2]) === nombreUniforme && String(row[3]) === talla;
  });

  return rowData ? rowData[0] : 'N/A';
}


/**
 * Guarda la nueva solicitud en la hoja de Solicitudes.
 * @param {object} formData - Datos del formulario.
 */
function saveSolicitud(formData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_SOLICITUDES_NAME);
  
  // Los datos de la maestra se traen al inicio. El usuario solo selecciona Uniforme y Talla.
  // Campos a guardar: Marca temporal-Correo-Cedula-Nombre-Sexo-ID Registro-Código-Uniforme-Talla
  const newRow = [
    new Date(), // Marca temporal (now)
    formData.correo || '', // Si tienes el correo en la Maestra, tráelo también.
    formData.cedula,
    formData.nombre,
    formData.sexo,
    formData.cargo,
    formData.ceco,
    formData.codigo,
    formData.uniforme,
    formData.talla
  ];
  Logger.log(newRow)

  sheet.appendRow(newRow);
  return 'Solicitud guardada con éxito.';
}

/**
 * Función auxiliar para incluir archivos HTML (CSS, JavaScript).
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}