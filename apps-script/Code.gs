/**
 * Súper de la casa: API sobre una planilla de Google.
 *
 * Instalación (una sola vez):
 *  1. Crear una planilla nueva en Google Sheets (por ejemplo "Súper de la casa").
 *  2. Extensiones > Apps Script. Borrar lo que haya y pegar este archivo entero.
 *  3. Configuración del proyecto (ícono de engranaje) > Propiedades de la secuencia de comandos >
 *     Agregar propiedad: CLAVE = una palabra clave que solo sepan Edi y Nati.
 *  4. Implementar > Nueva implementación > Tipo: Aplicación web.
 *     Ejecutar como: Yo. Quién tiene acceso: Cualquier usuario.
 *  5. Autorizar y copiar la URL que termina en /exec. Esa URL y la clave se cargan en la página.
 *
 * Para actualizar el código sin cambiar la URL: Implementar > Administrar implementaciones >
 *   lápiz (editar) > Versión: Nueva versión > Implementar.
 *
 * Los datos quedan en cuatro hojas que el script crea solo: compras, lista, config, facturas.
 * Las fotos de facturas van a la carpeta de Drive "Súper de la casa - Facturas":
 * primero en "Por cargar" y, una vez cargadas, en una subcarpeta por mes (AAAA-MM).
 * Sin la clave, el script no devuelve ni guarda nada.
 */

const HOJAS = {
  compras: ['id', 'fecha', 'lugar', 'total', 'categoria', 'nota', 'origen', 'items', 'creado'],
  lista: ['id', 'nombre', 'cantidad', 'comprado', 'creado'],
  config: ['clave', 'valor'],
  facturas: ['id', 'fecha', 'archivoId', 'nombre', 'estado', 'compraId', 'creado'],
};

function doGet(e) {
  return responder_(function () {
    verificarClave_(e.parameter.clave);
    if (e.parameter.accion === 'factura') return leerFactura_(e.parameter.id);
    return leerTodo_();
  });
}

function doPost(e) {
  return responder_(function () {
    const body = JSON.parse(e.postData.contents || '{}');
    verificarClave_(body.clave);
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    let nuevoId = null;
    try {
      switch (body.accion) {
        case 'agregar': nuevoId = agregar_(body.hoja, body.datos); break;
        case 'actualizar': actualizar_(body.hoja, body.id, body.datos); break;
        case 'borrar': borrar_(body.hoja, body.id); break;
        case 'presupuesto': guardarConfig_('presupuesto', Number(body.valor) || 0); break;
        case 'subirFactura': nuevoId = subirFactura_(body.imagen, body.mime); break;
        case 'marcarFactura': marcarFactura_(body.id, body.datos || {}); break;
        case 'borrarFactura': borrarFactura_(body.id); break;
        default: throw new Error('Acción desconocida');
      }
    } finally {
      lock.releaseLock();
    }
    const out = leerTodo_();
    out.nuevoId = nuevoId;
    return out;
  });
}

/* ---------- Facturas en Drive ---------- */

/** Ejecutar una vez desde el editor (botón Ejecutar) para dar permiso de Drive y crear la carpeta. */
function autorizar() {
  const c = carpeta_();
  subcarpeta_(c, 'Por cargar');
  Logger.log('Carpeta lista: ' + c.getUrl());
}

function carpeta_() {
  const p = PropertiesService.getScriptProperties();
  const id = p.getProperty('CARPETA');
  if (id) { try { const f = DriveApp.getFolderById(id); if (!f.isTrashed()) return f; } catch (e) {} }
  const nueva = DriveApp.createFolder('Súper de la casa - Facturas');
  p.setProperty('CARPETA', nueva.getId());
  return nueva;
}

function subcarpeta_(padre, nombre) {
  const it = padre.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : padre.createFolder(nombre);
}

function hoyTexto_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function subirFactura_(b64, mime) {
  if (!b64) throw new Error('Falta la imagen');
  const id = Utilities.getUuid();
  const hoy = hoyTexto_();
  const hora = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'HHmmss');
  const nombre = hoy + ' ' + hora + ' factura.jpg';
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime || 'image/jpeg', nombre);
  const file = subcarpeta_(carpeta_(), 'Por cargar').createFile(blob);
  hoja_('facturas').appendRow([id, "'" + hoy, file.getId(), nombre, 'pendiente', '', Date.now()]);
  return id;
}

function filaFactura_(id) {
  const fila = buscarFila_('facturas', id);
  if (fila < 0) throw new Error('No se encontró la factura');
  const v = hoja_('facturas').getRange(fila, 1, 1, HOJAS.facturas.length).getValues()[0];
  return { fila: fila, archivoId: String(v[2]), nombre: String(v[3]) };
}

function leerFactura_(id) {
  const f = filaFactura_(id);
  const blob = DriveApp.getFileById(f.archivoId).getBlob();
  return { id: id, nombre: f.nombre, mime: blob.getContentType(), imagen: Utilities.base64Encode(blob.getBytes()) };
}

function marcarFactura_(id, d) {
  const f = filaFactura_(id);
  const file = DriveApp.getFileById(f.archivoId);
  const fecha = String(d.fecha || hoyTexto_());
  const lugar = String(d.lugar || 'Compra').replace(/[\\/:*?"<>|]/g, '').slice(0, 40);
  const total = String(Math.round(Number(d.total) || 0)).replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  const nombre = fecha + ' ' + lugar + ' Gs ' + total + '.jpg';
  file.setName(nombre);
  file.moveTo(subcarpeta_(carpeta_(), fecha.slice(0, 7)));
  const sh = hoja_('facturas');
  sh.getRange(f.fila, 4, 1, 3).setValues([[nombre, 'cargada', String(d.compraId || '')]]);
}

function borrarFactura_(id) {
  const f = filaFactura_(id);
  try { DriveApp.getFileById(f.archivoId).setTrashed(true); } catch (e) {}
  hoja_('facturas').deleteRow(f.fila);
}

function responder_(fn) {
  let out;
  try {
    out = { ok: true, datos: fn() };
  } catch (err) {
    out = { ok: false, error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function verificarClave_(clave) {
  const real = PropertiesService.getScriptProperties().getProperty('CLAVE');
  if (!real) throw new Error('Falta configurar la propiedad CLAVE en el script');
  if (String(clave || '') !== real) throw new Error('Clave incorrecta');
}

function hoja_(nombre) {
  if (!HOJAS[nombre]) throw new Error('Hoja no válida');
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, HOJAS[nombre].length).setValues([HOJAS[nombre]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function filas_(nombre) {
  const sh = hoja_(nombre);
  const cols = HOJAS[nombre];
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, cols.length).getValues().map(function (r) {
    const o = {};
    cols.forEach(function (c, i) { o[c] = r[i]; });
    return o;
  });
}

function leerTodo_() {
  const compras = filas_('compras').map(function (c) {
    let items = [];
    try { items = c.items ? JSON.parse(c.items) : []; } catch (e) { items = []; }
    return {
      id: String(c.id), fecha: fechaTexto_(c.fecha), lugar: String(c.lugar || ''), total: Number(c.total) || 0,
      categoria: String(c.categoria || 'Otros'), nota: String(c.nota || ''), origen: String(c.origen || 'manual'),
      items: items, creado: Number(c.creado) || 0,
    };
  });
  const lista = filas_('lista').map(function (i) {
    return { id: String(i.id), nombre: String(i.nombre), cantidad: String(i.cantidad || ''), comprado: i.comprado === true || i.comprado === 'TRUE', creado: Number(i.creado) || 0 };
  });
  const cfg = {};
  filas_('config').forEach(function (r) { cfg[r.clave] = r.valor; });
  const facturas = filas_('facturas').map(function (f) {
    return { id: String(f.id), fecha: fechaTexto_(f.fecha), nombre: String(f.nombre), estado: String(f.estado), compraId: String(f.compraId || ''), creado: Number(f.creado) || 0 };
  });
  return { compras: compras, lista: lista, facturas: facturas, presupuesto: Number(cfg.presupuesto) || 1500000 };
}

function fechaTexto_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return String(v || '');
}

function aFila_(nombre, datos) {
  return HOJAS[nombre].map(function (c) {
    const v = datos[c];
    if (c === 'items') return JSON.stringify(Array.isArray(v) ? v : []);
    if (c === 'fecha') return "'" + String(v || ''); // texto, para que Sheets no la convierta
    return v === undefined || v === null ? '' : v;
  });
}

function agregar_(nombre, datos) {
  if (nombre === 'config' || nombre === 'facturas') throw new Error('Hoja no válida');
  datos = datos || {};
  datos.id = Utilities.getUuid();
  datos.creado = Date.now();
  hoja_(nombre).appendRow(aFila_(nombre, datos));
  return datos.id;
}

function buscarFila_(nombre, id) {
  const sh = hoja_(nombre);
  const n = sh.getLastRow() - 1;
  if (n < 1) return -1;
  const ids = sh.getRange(2, 1, n, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function actualizar_(nombre, id, datos) {
  const fila = buscarFila_(nombre, id);
  if (fila < 0) throw new Error('No se encontró el registro');
  const sh = hoja_(nombre);
  const cols = HOJAS[nombre];
  Object.keys(datos || {}).forEach(function (k) {
    const i = cols.indexOf(k);
    if (i > 0 && k !== 'creado') {
      let v = datos[k];
      if (k === 'items') v = JSON.stringify(v || []);
      if (k === 'fecha') v = "'" + v;
      sh.getRange(fila, i + 1).setValue(v);
    }
  });
}

function borrar_(nombre, id) {
  const fila = buscarFila_(nombre, id);
  if (fila > 0) hoja_(nombre).deleteRow(fila);
}

function guardarConfig_(clave, valor) {
  const sh = hoja_('config');
  const n = sh.getLastRow() - 1;
  if (n > 0) {
    const ks = sh.getRange(2, 1, n, 1).getValues();
    for (let i = 0; i < ks.length; i++) {
      if (ks[i][0] === clave) { sh.getRange(i + 2, 2).setValue(valor); return; }
    }
  }
  sh.appendRow([clave, valor]);
}
