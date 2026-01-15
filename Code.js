const HOJA_MAESTRA = "Maestra";
const HOJA_INVENTARIO = "Inventario";
const HOJA_SOLICITUD = "Solicitud";

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Solicitud de Uniformes')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(nombreArchivo) {
  return HtmlService.createHtmlOutputFromFile(nombreArchivo).getContent();
}

function validarIngreso(cedula) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaM = ss.getSheetByName(HOJA_MAESTRA);
  const hojaS = ss.getSheetByName(HOJA_SOLICITUD);
  const datosM = hojaM.getRange("A2:M" + hojaM.getLastRow()).getValues();
  const usuario = datosM.find(f => String(f[0]) === String(cedula));

  if (!usuario) return null;

  let yaTieneSolicitud = false;
  if (hojaS.getLastRow() >= 2) {
    const datosS = hojaS.getRange("C2:C" + hojaS.getLastRow()).getValues();
    yaTieneSolicitud = datosS.some(fila => String(fila[0]) === String(cedula));
  }

  return {
    cedula: usuario[0],
    nombre: usuario[1],
    sexo: usuario[2],
    cargo: usuario[6],
    ceco: usuario[10],
    rol: usuario[12],
    haSolicitado: yaTieneSolicitud
  };
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
  const fila = datos.find(f => f[2] == prenda && f[3] == talla);
  console.log(fila)
  Logger.log(prenda)
  Logger.log(talla)
  return fila ? fila[0] : 'N/A';
}

function guardarPedidoMasivo(carrito, user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName(HOJA_SOLICITUD);
  const fecha = new Date();
  const correoUsuario = Session.getActiveUser().getEmail();

  const filas = carrito.map(item => [
    fecha,
    correoUsuario,
    user.cedula,
    user.nombre,
    user.sexo,
    user.cargo,
    user.ceco,
    item.codigo,
    item.prenda,
    item.talla
  ]);

  hoja.getRange(hoja.getLastRow() + 1, 1, filas.length, 10).setValues(filas);

  enviarCorreoConfirmacion(carrito, user, correoUsuario, fecha);

  return "✅ ¡Pedido enviado con éxito! Se ha enviado un resumen a: " + correoUsuario;
}

function enviarCorreoConfirmacion(carrito, user, email, fecha) {
  const fechaTxt = Utilities.formatDate(fecha, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm");

  let tablaArticulos = `
    <table style="border-collapse: collapse; width: 100%; font-family: sans-serif;">
      <thead>
        <tr style="background-color: #007bff; color: white;">
          <th style="padding: 10px; border: 1px solid #ddd;">Prenda</th>
          <th style="padding: 10px; border: 1px solid #ddd;">Talla</th>
          <th style="padding: 10px; border: 1px solid #ddd;">Código</th>
        </tr>
      </thead>
      <tbody>`;

  carrito.forEach(item => {
    tablaArticulos += `
      <tr>
        <td style="padding: 10px; border: 1px solid #ddd;">${item.prenda}</td>
        <td style="padding: 10px; border: 1px solid #ddd; text-align: center;">${item.talla}</td>
        <td style="padding: 10px; border: 1px solid #ddd; text-align: center;">${item.codigo}</td>
      </tr>`;
  });

  tablaArticulos += `</tbody></table>`;

  const cuerpoHtml = `
    <div style="font-family: sans-serif; color: #333; max-width: 600px; border: 1px solid #eee; padding: 20px; border-radius: 10px;">
      <h2 style="color: #00A441;">Confirmación de Solicitud de Uniformes</h2>
      <p>Hola <strong>${user.nombre}</strong>,</p>
      <p>Hemos registrado correctamente tu solicitud realizada el <strong>${fechaTxt}</strong>. A continuación, el detalle de tus artículos:</p>
      
      ${tablaArticulos}
      
      <p style="margin-top: 20px;"><strong>Datos del colaborador:</strong><br>
      Cédula: ${user.cedula}<br>
      Cargo: ${user.cargo}<br>
      CECO: ${user.ceco}</p>
      
      <hr style="border: 0; border-top: 1px solid #eee; margin: 20px 0;">
      <p style="font-size: 12px; color: #777; text-align: center;">
        Este es un mensaje automático del Sistema de Gestión de Uniformes.
      </p>
    </div>
  `;

  try {
    GmailApp.sendEmail(email, "Confirmación de Solicitud de Uniformes - " + user.nombre, "", {
      htmlBody: cuerpoHtml,
      name: "Uniformes CNCH"
    });
  } catch (error) {
    console.error("Error al enviar email con GmailApp: " + error.toString());
  }
}

function obtenerHistorial(cedula) {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_SOLICITUD);
  if (hoja.getLastRow() < 2) return [];
  const datos = hoja.getDataRange().getValues();
  return datos.slice(1).filter(f => String(f[2]) === String(cedula))
    .map(f => [f[0] instanceof Date ? f[0].toISOString() : f[0], f[3], f[8], f[9]]);
}