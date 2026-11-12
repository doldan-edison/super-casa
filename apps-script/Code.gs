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
 * Los datos quedan en tres hojas que el script crea solo: compras, lista, config.
 * Sin la clave, el script no devuelve ni guarda nada.
 */

const HOJAS = {
  compras: ['id', 'fecha', 'lugar', 'total', 'categoria', 'nota', 'origen', 'items', 'creado'],
  lista: ['id', 'nombre', 'cantidad', 'comprado', 'creado'],
  config: ['clave', 'valor'],
};

function doGet(e) {
  return responder_(function () {
    verificarClave_(e.parameter.clave);
    return leerTodo_();
  });
}

function doPost(e) {
  return responder_(function () {
    const body = JSON.parse(e.postData.contents || '{}');
    verificarClave_(body.clave);
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      switch (body.accion) {
        case 'agregar': agregar_(body.hoja, body.datos); break;
        case 'actualizar': actualizar_(body.hoja, body.id, body.datos); break;
        case 'borrar': borrar_(body.hoja, body.id); break;
        case 'presupuesto': guardarConfig_('presupuesto', Number(body.valor) || 0); break;
        default: throw new Error('Acción desconocida');
      }
    } finally {
      lock.releaseLock();
    }
    return leerTodo_();
  });
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
  return { compras: compras, lista: lista, presupuesto: Number(cfg.presupuesto) || 1500000 };
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
  if (nombre === 'config') throw new Error('Hoja no válida');
  datos = datos || {};
  datos.id = Utilities.getUuid();
  datos.creado = Date.now();
  hoja_(nombre).appendRow(aFila_(nombre, datos));
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
