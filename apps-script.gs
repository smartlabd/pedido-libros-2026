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

const EMAIL_NOTIFICACION = 'directorclassical@gmail.com';
const NOMBRE_HOJA        = 'Pedidos';
const CLAVE_ADMIN        = 'libros2026'; // Cambia esto por tu contraseña
const ESTADOS_VALIDOS    = ['Pendiente', 'Pagado', 'Entregado'];
const ESTADO_POR_DEFECTO = 'Pendiente';
const COL_ESTADO         = 10; // columna J

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

  // Eliminar item de un pedido
  if (action === 'deleteItem') {
    const clave = e.parameter.clave || '';
    if (clave !== CLAVE_ADMIN) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: 'Acceso no autorizado' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    try {
      const familia  = e.parameter.familia;
      const ref      = e.parameter.ref;
      const email    = e.parameter.email;
      const telefono = e.parameter.telefono;
      const result   = eliminarItem(familia, ref, email, telefono);
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, ...result }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Cambiar estado de una o varias familias (individual o en lote)
  if (action === 'setEstado') {
    const clave = e.parameter.clave || '';
    if (clave !== CLAVE_ADMIN) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: 'Acceso no autorizado' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    try {
      const result = setEstadoFamilias(e.parameter.familias || '[]', e.parameter.estado || '');
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, ...result }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }

  // Borrar en lote todos los pedidos de una o varias familias
  if (action === 'deleteFamilias') {
    const clave = e.parameter.clave || '';
    if (clave !== CLAVE_ADMIN) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: 'Acceso no autorizado' }))
        .setMimeType(ContentService.MimeType.JSON);
    }
    try {
      const result = eliminarFamilias(e.parameter.familias || '[]');
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, ...result }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch(err) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
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
  const ss  = SpreadsheetApp.openById('1eZDkyjuHYCA0XByAxlx2hUcPetneWrJTl_I74rQP7bU');
  let sheet = ss.getSheetByName(NOMBRE_HOJA);

  if (!sheet) {
    sheet = ss.insertSheet(NOMBRE_HOJA);
    sheet.appendRow(['Fecha','Familia','Teléfono','Ref','Grado','Descripción','Cantidad','Precio USD','Subtotal USD','Estado']);
    sheet.getRange(1,1,1,10).setFontWeight('bold').setBackground('#1a2744').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, 10, 120);
    sheet.setColumnWidth(6, 320);
  }
  asegurarColumnaEstado(sheet);

  data.libros.forEach(libro => {
    sheet.appendRow([
      data.fecha, data.nombre, "'" + data.telefono,
      libro.ref, libro.grade, libro.desc,
      libro.qty, parseFloat(libro.usd), parseFloat(libro.subtotal), ESTADO_POR_DEFECTO
    ]);
  });

  // Fila de total de familia
  const totalRow = ['', '★ ' + data.nombre, "'" + data.telefono, '', '', 'TOTAL FAMILIA', '', '', parseFloat(data.total), ESTADO_POR_DEFECTO];
  sheet.appendRow(totalRow);
  const lr = sheet.getLastRow();
  sheet.getRange(lr, 1, 1, 10).setBackground('#f5e6c8').setFontWeight('bold');
}

// ------------------------------------------------------------
// ASEGURAR COLUMNA "Estado" (migración suave de hojas viejas)
// ------------------------------------------------------------
function asegurarColumnaEstado(sheet) {
  const encabezado = String(sheet.getRange(1, COL_ESTADO).getValue()).trim();
  if (encabezado !== 'Estado') {
    sheet.getRange(1, COL_ESTADO).setValue('Estado')
      .setFontWeight('bold').setBackground('#1a2744').setFontColor('#ffffff');
  }
}

// ------------------------------------------------------------
// LEER DATOS PARA EL PANEL ADMIN
// ------------------------------------------------------------
function obtenerDatos() {
  const ss    = SpreadsheetApp.openById('1eZDkyjuHYCA0XByAxlx2hUcPetneWrJTl_I74rQP7bU');
  const sheet = ss.getSheetByName(NOMBRE_HOJA);
  if (!sheet || sheet.getLastRow() < 2) {
    return { familias: [], libros: [], totalGlobal: 0 };
  }
  asegurarColumnaEstado(sheet);

  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues();

  const familiasMap = {};   // nombre -> { nombre, telefono, fecha, estado, libros[], total }
  const librosMap   = {};   // ref    -> { ref, grade, desc, qty, total }

  rows.forEach(row => {
    const [fecha, familia, tel, ref, grado, desc, qty, precio, subtotal, estado] = row;

    if (!ref || ref === '') return; // fila de total, skip

    // Por familia
    if (!familiasMap[familia]) {
      familiasMap[familia] = {
        nombre: familia, telefono: tel, fecha: fecha,
        estado: (estado && String(estado).trim()) || ESTADO_POR_DEFECTO,
        libros: [], total: 0
      };
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

  // ── Email al director ──
  const cuerpoAdmin =
    'Se ha recibido un nuevo pedido de libros.\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'DATOS DE LA FAMILIA\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'Nombre:    ' + data.nombre + '\n' +
    'Teléfono:  ' + data.telefono + '\n' +
    'Email:     ' + (data.email || '—') + '\n' +
    'Fecha:     ' + data.fecha + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'PEDIDO DE ' + data.nombre.toUpperCase() + '\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    lista + '\n\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
    'TOTAL: $' + data.total + ' USD\n' +
    '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━';

  MailApp.sendEmail({
    to: EMAIL_NOTIFICACION,
    subject: '📚 Pedido de ' + data.nombre + ' — $' + data.total + ' USD',
    body: cuerpoAdmin
  });

  // ── Copia a la familia ──
  if (data.email) {
    var primerNombre = data.nombre.split(' ')[0];
    const cuerpoFamilia =
      'Hola ' + primerNombre + ',\n\n' +
      'Tu pedido de libros fue recibido exitosamente. Aquí está el resumen:\n\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      'TU PEDIDO\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      lista + '\n\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
      'TOTAL A PAGAR: $' + data.total + ' USD\n' +
      '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n\n' +
      'Fecha: ' + data.fecha + '\n\n' +
      'Pronto nos pondremos en contacto contigo para coordinar la entrega y el pago.\n\n' +
      'Gracias,\nEl equipo';

    MailApp.sendEmail({
      to: data.email,
      subject: '📚 Confirmación de tu pedido de libros — $' + data.total + ' USD',
      body: cuerpoFamilia
    });
  }
}

// ------------------------------------------------------------
// ELIMINAR ITEM DE UN PEDIDO
// ------------------------------------------------------------
function eliminarItem(familia, ref, emailFamilia, telefonoFamilia) {
  const ss    = SpreadsheetApp.openById('1eZDkyjuHYCA0XByAxlx2hUcPetneWrJTl_I74rQP7bU');
  const sheet = ss.getSheetByName(NOMBRE_HOJA);
  if (!sheet) throw new Error('Hoja no encontrada');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No hay datos');

  const data = sheet.getRange(2, 1, lastRow - 1, 9).getValues();

  // Find the row to delete: matching familia name AND ref
  let deletedDesc = '';
  let deletedQty  = 0;
  let deletedSub  = 0;
  let rowToDelete = -1;

  for (let i = 0; i < data.length; i++) {
    const [fecha, fam, tel, rowRef, grado, desc, qty, precio, subtotal] = data[i];
    if (String(fam).trim() === String(familia).trim() && String(rowRef).trim() === String(ref).trim()) {
      rowToDelete = i + 2; // +2 because data starts at row 2
      deletedDesc = desc;
      deletedQty  = qty;
      deletedSub  = parseFloat(subtotal);
      break;
    }
  }

  if (rowToDelete === -1) throw new Error('Item no encontrado en la hoja');

  // Delete the row
  sheet.deleteRow(rowToDelete);

  // Update the TOTAL FAMILIA row for this family
  // Re-read after deletion
  const newLastRow = sheet.getLastRow();
  if (newLastRow >= 2) {
    const newData = sheet.getRange(2, 1, newLastRow - 1, 9).getValues();
    let newTotal = 0;
    let totalRowIdx = -1;

    for (let i = 0; i < newData.length; i++) {
      const [, fam, , rowRef, , , , , subtotal] = newData[i];
      if (String(fam).includes('★') && String(fam).includes(familia)) {
        totalRowIdx = i + 2;
      } else if (String(fam).trim() === String(familia).trim() && rowRef) {
        newTotal += parseFloat(subtotal) || 0;
      }
    }

    // Update total row
    if (totalRowIdx !== -1) {
      sheet.getRange(totalRowIdx, 9).setValue(newTotal);
    }
  }

  // Send email to family notifying the change
  if (emailFamilia) {
    const primerNombre = familia.split(' ')[0];
    MailApp.sendEmail({
      to: emailFamilia,
      subject: '📚 Tu pedido de libros fue modificado',
      body:
        'Hola ' + primerNombre + ',\n\n' +
        'El administrador realizó un cambio en tu pedido de libros.\n\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
        'ITEM ELIMINADO\n' +
        '━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n' +
        '  • ' + deletedDesc + '\n' +
        '    Cantidad: ' + deletedQty + ' · Subtotal: $' + deletedSub.toFixed(2) + ' USD\n\n' +
        'Si tienes alguna pregunta, por favor contáctanos.\n\n' +
        'Gracias,\nEl equipo'
    });
  }

  // Also notify admin
  MailApp.sendEmail({
    to: EMAIL_NOTIFICACION,
    subject: '🗑️ Item eliminado del pedido de ' + familia,
    body:
      'Se eliminó un item del pedido de ' + familia + '.\n\n' +
      'Item: ' + deletedDesc + '\n' +
      'Cantidad: ' + deletedQty + '\n' +
      'Subtotal eliminado: $' + deletedSub.toFixed(2) + ' USD\n\n' +
      'Teléfono familia: ' + (telefonoFamilia || '—') + '\n' +
      'Email familia: ' + (emailFamilia || '—')
  });

  return { deleted: deletedDesc };
}


// ------------------------------------------------------------
// CAMBIAR ESTADO DE UNA O VARIAS FAMILIAS
// familiasJson: JSON string con array de nombres de familia
// ------------------------------------------------------------
function setEstadoFamilias(familiasJson, estado) {
  if (ESTADOS_VALIDOS.indexOf(estado) === -1) {
    throw new Error('Estado inválido: ' + estado);
  }
  let familias;
  try { familias = JSON.parse(familiasJson); }
  catch(e) { throw new Error('Lista de familias inválida'); }
  if (!Array.isArray(familias) || familias.length === 0) {
    throw new Error('Sin familias para actualizar');
  }

  const ss    = SpreadsheetApp.openById('1eZDkyjuHYCA0XByAxlx2hUcPetneWrJTl_I74rQP7bU');
  const sheet = ss.getSheetByName(NOMBRE_HOJA);
  if (!sheet) throw new Error('Hoja no encontrada');
  asegurarColumnaEstado(sheet);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No hay datos');

  const objetivos  = familias.map(f => String(f).trim());
  const nFilas     = lastRow - 1;
  const columnaFam = sheet.getRange(2, 2, nFilas, 1).getValues();
  const columnaEst = sheet.getRange(2, COL_ESTADO, nFilas, 1).getValues();

  let actualizadas = 0;
  for (let i = 0; i < columnaFam.length; i++) {
    if (objetivos.indexOf(nombreFamiliaPlano_(columnaFam[i][0])) !== -1) {
      columnaEst[i][0] = estado;
      actualizadas++;
    }
  }
  sheet.getRange(2, COL_ESTADO, nFilas, 1).setValues(columnaEst);

  MailApp.sendEmail({
    to: EMAIL_NOTIFICACION,
    subject: '✏️ Estado → ' + estado + ' · ' + familias.length + ' familia(s)',
    body:
      'Se cambió el estado a "' + estado + '" para:\n\n' +
      familias.map(f => '  • ' + f).join('\n') + '\n\n' +
      'Filas actualizadas en la hoja: ' + actualizadas
  });

  return { estado: estado, familias: familias, filasActualizadas: actualizadas };
}

// ------------------------------------------------------------
// BORRAR EN LOTE TODOS LOS PEDIDOS DE VARIAS FAMILIAS
// (elimina las filas de libros y la fila ★ TOTAL de cada familia)
// ------------------------------------------------------------
function eliminarFamilias(familiasJson) {
  let familias;
  try { familias = JSON.parse(familiasJson); }
  catch(e) { throw new Error('Lista de familias inválida'); }
  if (!Array.isArray(familias) || familias.length === 0) {
    throw new Error('Sin familias para eliminar');
  }

  const ss    = SpreadsheetApp.openById('1eZDkyjuHYCA0XByAxlx2hUcPetneWrJTl_I74rQP7bU');
  const sheet = ss.getSheetByName(NOMBRE_HOJA);
  if (!sheet) throw new Error('Hoja no encontrada');

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error('No hay datos');

  const objetivos  = familias.map(f => String(f).trim());
  const columnaFam = sheet.getRange(2, 2, lastRow - 1, 1).getValues();

  const filas = [];
  for (let i = 0; i < columnaFam.length; i++) {
    if (objetivos.indexOf(nombreFamiliaPlano_(columnaFam[i][0])) !== -1) {
      filas.push(i + 2); // +2: los datos empiezan en la fila 2
    }
  }
  if (filas.length === 0) throw new Error('No se encontraron filas para esas familias');

  // Borrar de abajo hacia arriba para no desplazar los índices restantes
  filas.sort((a, b) => b - a).forEach(r => sheet.deleteRow(r));

  MailApp.sendEmail({
    to: EMAIL_NOTIFICACION,
    subject: '🗑️ Pedidos eliminados en lote · ' + familias.length + ' familia(s)',
    body:
      'Se eliminaron por completo los pedidos de:\n\n' +
      familias.map(f => '  • ' + f).join('\n') + '\n\n' +
      'Filas eliminadas en la hoja: ' + filas.length
  });

  return { familias: familias, filasEliminadas: filas.length };
}

// Devuelve el nombre de familia sin el prefijo ★ de la fila de total
function nombreFamiliaPlano_(valorCelda) {
  const s = String(valorCelda).trim();
  return s.charAt(0) === '★' ? s.replace(/^★\s*/, '').trim() : s;
}


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
    nombre: 'Josue Belmonte', telefono: '+58 414 1086197',
    email: 'info@smartlabd.com',
    fecha: new Date().toLocaleString('es-VE'), total: '23.55',
    libros: [{ ref: '20511', grade: 'Grado 1', desc: 'Estudios Sociales Alumno', qty: 2, usd: 11.71, subtotal: '23.42' }]
  });
  Logger.log('Emails enviados');
}

// Corre testPedido() primero para tener la "Familia Test" en la hoja.
function testSetEstado() {
  const r = setEstadoFamilias(JSON.stringify(['Familia Test']), 'Pagado');
  Logger.log(JSON.stringify(r));
}

function testDeleteFamilias() {
  const r = eliminarFamilias(JSON.stringify(['Familia Test']));
  Logger.log(JSON.stringify(r));
}
