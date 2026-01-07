const HOJA_MAESTRA = "Maestra";
const HOJA_INVENTARIO = "Inventario";
const HOJA_SOLICITUD = "Solicitud";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Gestión de Uniformes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(nombreArchivo) {
  return HtmlService.createHtmlOutputFromFile(nombreArchivo).getContent();
}

function validarIngreso(cedula) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_MAESTRA);
  const datos = hoja.getRange("A2:M" + hoja.getLastRow()).getValues();
  const usuario = datos.find(f => String(f[0]) === String(cedula));
  
  if (usuario) {
    return {
      cedula: usuario[0], nombre: usuario[1], sexo: usuario[2],
      cargo: usuario[6], ceco: usuario[10], rol: usuario[12]
    };
  }
  return null;
}

function obtenerUniformesPorRol(sexo, rol) {
  const datos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_INVENTARIO).getDataRange().getValues();
  const prendas = datos.slice(1).filter(f => {
    const s = String(f[1]).toUpperCase();
    const r = String(f[4]).toUpperCase();
    return (s === 'UNISEX' || s === sexo.toUpperCase()) && r === rol.toUpperCase();
  }).map(f => f[2]);
  return [...new Set(prendas)];
}

function obtenerTallasPorUniforme(prenda) {
  const datos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_INVENTARIO).getDataRange().getValues();
  return datos.slice(1).filter(f => f[2] === prenda).map(f => f[3]);
}

function obtenerCodigoUniforme(prenda, talla) {
  const datos = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_INVENTARIO).getDataRange().getValues();
  const fila = datos.find(f => f[2] === prenda && f[3] === talla);
  return fila ? fila[0] : 'N/A';
}

function guardarPedidoMasivo(carrito, user) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_SOLICITUD);
  const fecha = new Date();
  const filas = carrito.map(item => [
    fecha, "", user.cedula, user.nombre, user.sexo, user.cargo, user.ceco, item.codigo, item.prenda, item.talla
  ]);
  hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 10).setValues(filas);
  return "✅ ¡Pedido enviado con éxito!";
}

function obtenerHistorial(cedula) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_SOLICITUD);
  if (hoja.getLastRow() < 2) return [];
  const datos = hoja.getDataRange().getValues();
  return datos.slice(1).filter(f => String(f[2]) === String(cedula))
    .map(f => [f[0] instanceof Date ? f[0].toISOString() : f[0], f[3], f[8], f[9]]);
}