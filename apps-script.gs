// ============================================================
// PEDIDO DE LIBROS 2026 — Google Apps Script
// ============================================================
// INSTRUCCIONES INICIALES:
// 1. Abre script.google.com → Nuevo proyecto
// 2. Pega todo este código (reemplaza lo que hay)
// 3. Clic en "Implementar" → "Nueva implementación"
//    - Tipo: Aplicación web
//    - Ejecutar como: Yo (tu cuenta)
//    - Quién tiene acceso: Cualquier usuario
// 4. Copia la URL → pégala en index.html Y en admin.html
//    donde dice REPLACE_WITH_YOUR_APPS_SCRIPT_URL
//
// CUANDO ACTUALICES EL CÓDIGO:
//   Implementar → Administrar implementaciones → Editar (lápiz)
//   → Versión: Nueva versión → Implementar
// ============================================================

const EMAIL_NOTIFICACION = 'info@smartlabd.com';
const NOMBRE_HOJA        = 'Pedidos';
const CLAVE_ADMIN        = 'libros2026'; // Cambia esto por tu contraseña

// ------------------------------------------------------------
// GET — panel admin o ping
// ------------------------------------------------------------
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;

  // Ping de verificación
  if (action === 'ping') {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, message: 'Apps Script activo' }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // Panel admin: devuelve datos consolidados
  if (action === 'getOrders') {
    const clave = e.parameter.clave || '';
    if (clave !== CLAVE_ADMIN) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: 'Acceso no autorizado' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    try {
      const result = obtenerDatos();
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, ...result }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: 'Acción no reconocida' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ------------------------------------------------------------
// POST — recibe pedidos del formulario
// ------------------------------------------------------------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    guardarPedido(data);
    enviarNotificacion(data);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ------------------------------------------------------------
// GUARDAR PEDIDO
// ------------------------------------------------------------
function guardarPedido(data) {
  const ss  = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(NOMBRE_HOJA);

  if (!sheet) {
    sheet = ss.insertSheet(NOMBRE_HOJA);
    sheet.appendRow(['Fecha','Familia','Teléfono','Ref','Grado','Descripción','Cantidad','Precio USD','Subtotal USD']);
    sheet.getRange(1,1,1,9).setFontWeight('bold').setBackground('#1a2744').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 9, 120);
    sheet.setColumnWidth(6, 320);
  }

  data.libros.forEach(libro => {
    sheet.appendRow([
      data.fecha, data.nombre, data.telefono,
      libro.ref, libro.grade, libro.desc,
      libro.qty, parseFloat(libro.usd), parseFloat(libro.subtotal)
    ]);
  });

  // Fila de total de familia
  const totalRow = ['', '★ ' + data.nombre, data.telefono, '', '', 'TOTAL FAMILIA', '', '', parseFloat(data.total)];
  sheet.appendRow(totalRow);
  const lr = sheet.getLastRow();
  sheet.getRange(lr, 1, 1, 9).setBackground('#f5e6c8').setFontWeight('bold');
}

// ------------------------------------------------------------
// LEER DATOS PARA EL PANEL ADMIN
// ------------------------------------------------------------
function obtenerDatos() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(NOMBRE_HOJA);
  if (!sheet || sheet.getLastRow() < 2) {
    return { familias: [], libros: [], totalGlobal: 0 };
  }

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();

  const familiasMap = {};   // nombre -> { nombre, telefono, fecha, libros[], total }
  const librosMap   = {};   // ref    -> { ref, grade, desc, qty, total }

  rows.forEach(row => {
    const [fecha, familia, tel, ref, grado, desc, qty, precio, subtotal] = row;

    if (!ref || ref === '') return; // fila de total, skip

    // Por familia
    if (!familiasMap[familia]) {
      familiasMap[familia] = { nombre: familia, telefono: tel, fecha: fecha, libros: [], total: 0 };
    }
    familiasMap[familia].libros.push({ ref, grado, desc, qty: parseInt(qty), usd: parseFloat(precio), subtotal: parseFloat(subtotal) });
    familiasMap[familia].total += parseFloat(subtotal);

    // Por libro
    if (!librosMap[ref]) librosMap[ref] = { ref, grade: grado, desc, qty: 0, total: 0 };
    librosMap[ref].qty   += parseInt(qty);
    librosMap[ref].total += parseFloat(subtotal);
  });

  const gradeOrder = ['Preescolar','Grado 1','Grado 2','Grado 3','Grado 4','Grado 5','Grado 6','Grado 7','Grado 8','Grado 9'];
  const librosArr = Object.values(librosMap).sort((a,b) => gradeOrder.indexOf(a.grade) - gradeOrder.indexOf(b.grade));
  const familiasArr = Object.values(familiasMap).sort((a,b) => b.total - a.total);
  const totalGlobal = librosArr.reduce((s, l) => s + l.total, 0);

  return { familias: familiasArr, libros: librosArr, totalGlobal: parseFloat(totalGlobal.toFixed(2)) };
}

// ------------------------------------------------------------
// EMAIL DE NOTIFICACIÓN (admin) + COPIA A LA FAMILIA
// ------------------------------------------------------------
function enviarNotificacion(data) {
  const lista = data.libros.map(l =>
    `  • [${l.grade}] ${l.desc} (Ref: ${l.ref}) — ${l.qty} × $${l.usd} = $${l.subtotal}`
  ).join('\n');

  const cuerpoAdmin = `Nuevo pedido de libros — ${data.fecha}

FAMILIA: ${data.nombre}
TELÉFONO: ${data.telefono}
EMAIL: ${data.email}

LIBROS:
${lista}

TOTAL: $${data.total} USD`;

  // Email al admin
  MailApp.sendEmail({
    to: EMAIL_NOTIFICACION,
    subject: `📚 Nuevo pedido — ${data.nombre} ($${data.total} USD)`,
    body: cuerpoAdmin
  });

  // Copia a la familia si tienen email
  if (data.email) {
    const cuerpoFamilia = `Hola ${data.nombre},

Tu pedido de libros fue recibido exitosamente. Aquí está el resumen:

LIBROS SELECCIONADOS:
${lista}

TOTAL A PAGAR: $${data.total} USD

Fecha: ${data.fecha}

Nos pondremos en contacto contigo pronto para coordinar la entrega y el pago.

Gracias,
El equipo`;

    MailApp.sendEmail({
      to: data.email,
      subject: `📚 Confirmación de tu pedido de libros — $${data.total} USD`,
      body: cuerpoFamilia
    });
  }
}

// ------------------------------------------------------------
// TEST MANUAL
// ------------------------------------------------------------
function testPedido() {
  guardarPedido({
    nombre: 'Familia Test', telefono: '+58 414 1234567',
    fecha: new Date().toLocaleString('es-VE'), total: '23.55',
    libros: [
      { ref: '20511', grade: 'Grado 1', desc: 'Estudios Sociales Alumno', qty: 1, usd: 11.71, subtotal: '11.71' },
      { ref: '20319', grade: 'Grado 1', desc: 'Matemáticas Guía de maestro', qty: 1, usd: 11.84, subtotal: '11.84' }
    ]
  });
  Logger.log('Test OK — revisa la hoja Pedidos');
}

function testEmail() {
  enviarNotificacion({
    nombre: 'Familia Test', telefono: '+58 414 0000000',
    fecha: new Date().toLocaleString('es-VE'), total: '23.55',
    libros: [{ ref: '20511', grade: 'Grado 1', desc: 'Estudios Sociales Alumno', qty: 2, usd: 11.71, subtotal: '23.42' }]
  });
  Logger.log('Email enviado a ' + EMAIL_NOTIFICACION);
}
