/**
 * Registro Inventario - Backend (Google Apps Script)
 *
 * Este script va "pegado" dentro del Google Sheet (Extensiones > Apps Script).
 * Hojas que usa:
 *   - "Registros":  una fila por línea de conteo (todas las vistas).
 *                  Columna C = Periodo de inventario (AAAA-MM, ej. 2026-12).
 *   - "Mantenedor": listas desplegables (columnas A a H, fila 1 = encabezado).
 *
 * Clave del mantenedor: NO va en el código. Se configura en
 *   Configuración del proyecto > Propiedades de la secuencia de comandos
 *   Propiedad: CLAVE_MANTENEDOR   Valor: <tu clave>
 */

const HOJA_REGISTROS = 'Registros';
const ZONA_HORARIA = 'America/Santiago';
const HOJA_MANTENEDOR = 'Mantenedor';

// Columnas del Mantenedor (1 = A)
const COLUMNAS_MANTENEDOR = {
  jefes:      { col: 1, nombre: 'Jefes de Área' },
  productos:  { col: 2, nombre: 'Productos' },
  clases:     { col: 3, nombre: 'Clases' },
  largos:     { col: 4, nombre: 'Largos' },
  diamRCC:    { col: 5, nombre: 'Diámetro RCC' },
  diamRC:     { col: 6, nombre: 'Diámetro RC-RCI' },
  diamRD:     { col: 7, nombre: 'Diámetro RD-RDI' },
  secado:     { col: 8, nombre: 'Nivel de Secado' }
};

const MAX_FILAS_MANTENEDOR = 500;
const TRABAJADORES_MAX = 10;
const PERIODO_RE = /^\d{4}-\d{2}$/;

// ---------- Entrada HTTP ----------

function doGet(e) {
  const accion = (e && e.parameter && e.parameter.action) || '';
  let salida;
  try {
    if (accion === 'lists') salida = obtenerListas();
    else if (accion === 'periodos') salida = obtenerPeriodos();
    else if (accion === 'registros') salida = obtenerRegistros(e.parameter.periodo);
    else if (accion === 'registrosVista') salida = obtenerRegistrosVista(e.parameter.periodo, e.parameter.vista);
    else salida = { ok: false, error: 'Acción desconocida' };
  } catch (err) {
    salida = { ok: false, error: String(err) };
  }
  return responder(salida);
}

function doPost(e) {
  let salida;
  try {
    const cuerpo = JSON.parse(e.postData.contents);
    if (cuerpo.action === 'registrar') salida = registrar(cuerpo.data);
    else if (cuerpo.action === 'actualizar') salida = actualizar(cuerpo.fila, cuerpo.data);
    else if (cuerpo.action === 'verificarClave') salida = { ok: claveValida(cuerpo.clave) };
    else if (cuerpo.action === 'guardarMantenedor') salida = guardarMantenedor(cuerpo.clave, cuerpo.listas);
    else salida = { ok: false, error: 'Acción desconocida' };
  } catch (err) {
    salida = { ok: false, error: String(err) };
  }
  return responder(salida);
}

function responder(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ---------- Mantenedor ----------

function obtenerListas() {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_MANTENEDOR);
  const salida = { ok: true };
  Object.keys(COLUMNAS_MANTENEDOR).forEach(function (clave) {
    const col = COLUMNAS_MANTENEDOR[clave].col;
    const valores = hoja.getRange(2, col, MAX_FILAS_MANTENEDOR, 1).getValues();
    salida[clave] = valores
      .map(function (f) { return f[0]; })
      .filter(function (v) { return v !== '' && v !== null; });
  });
  return salida;
}

function claveValida(clave) {
  const real = PropertiesService.getScriptProperties().getProperty('CLAVE_MANTENEDOR');
  return !!real && String(clave) === String(real);
}

function guardarMantenedor(clave, listas) {
  if (!claveValida(clave)) return { ok: false, error: 'Clave incorrecta' };
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_MANTENEDOR);
  Object.keys(COLUMNAS_MANTENEDOR).forEach(function (clave) {
    const col = COLUMNAS_MANTENEDOR[clave].col;
    const valores = (listas[clave] || []).map(function (v) {
      return [String(v).trim()];
    }).filter(function (f) { return f[0] !== ''; });
    hoja.getRange(2, col, MAX_FILAS_MANTENEDOR, 1).clearContent();
    if (valores.length > 0) {
      hoja.getRange(2, col, valores.length, 1).setValues(valores);
    }
  });
  return { ok: true };
}

// ---------- Registros ----------

function errorRegistro(d) {
  const faltantes = [];
  if (!d.vista) faltantes.push('vista');
  if (!d.periodo || !PERIODO_RE.test(String(d.periodo))) faltantes.push('periodo de inventario (AAAA-MM)');
  if (!d.jefe) faltantes.push('jefe de área');
  if (!d.producto) faltantes.push('producto');
  if (faltantes.length) return 'Falta: ' + faltantes.join(', ');

  const unidades = [d.produccion, d.recuperacion, d.rechazo];
  const hayUnidades = unidades.some(function (u) {
    return u !== '' && u !== null && u !== undefined && Number(u) > 0;
  });
  if (!hayUnidades) return 'Ingresa al menos una cantidad (producción, recuperación o rechazo)';
  return '';
}

// Valores de las columnas B a X (vista, periodo, jefe, trabajadores, datos de la línea).
function valoresFila(d) {
  const trabajadores = [];
  for (let i = 0; i < TRABAJADORES_MAX; i++) {
    trabajadores.push((d.trabajadores && d.trabajadores[i]) || '');
  }
  return [
    d.vista,
    String(d.periodo),
    d.jefe
  ].concat(trabajadores, [
    d.producto,
    d.clase || '',
    d.largo || '',
    d.diametro || '',
    numeroOVacio(d.produccion),
    numeroOVacio(d.recuperacion),
    numeroOVacio(d.rechazo),
    d.secado || '',
    d.sector || '',
    d.comentario || ''
  ]);
}

function registrar(d) {
  const error = errorRegistro(d);
  if (error) return { ok: false, error: error };
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_REGISTROS);
  hoja.appendRow([new Date()].concat(valoresFila(d)));
  return { ok: true };
}

// Reescribe una línea existente. Verifica que la fila siga siendo la misma vista y mes.
function actualizar(fila, d) {
  const n = Number(fila);
  const error = errorRegistro(d);
  if (error) return { ok: false, error: error };
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_REGISTROS);
  if (!n || n < 2 || n > hoja.getLastRow()) return { ok: false, error: 'Fila no válida' };
  const actual = hoja.getRange(n, 2, 1, 2).getValues()[0];
  if (String(actual[0]) !== String(d.vista) || periodoDeCelda(actual[1]) !== String(d.periodo)) {
    return { ok: false, error: 'La fila cambió. Recarga la lista antes de guardar.' };
  }
  hoja.getRange(n, 2, 1, 23).setValues([valoresFila(d)]);
  return { ok: true };
}

function numeroOVacio(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(v);
  return isNaN(n) ? '' : n;
}

function esFecha(v) {
  return v !== null && typeof v === 'object' && typeof v.getTime === 'function';
}

// Columna C puede venir como texto "2026-12" o como fecha (si Sheets la convirtió).
function periodoDeCelda(v) {
  if (esFecha(v)) return Utilities.formatDate(v, ZONA_HORARIA, 'yyyy-MM');
  return String(v || '').trim();
}

function obtenerPeriodos() {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_REGISTROS);
  const datos = hoja.getDataRange().getValues();
  const vistos = {};
  datos.slice(1).forEach(function (f) {
    const p = periodoDeCelda(f[2]);
    if (PERIODO_RE.test(p)) vistos[p] = true;
  });
  const periodos = Object.keys(vistos).sort().reverse();
  return { ok: true, periodos: periodos };
}

function obtenerRegistros(periodo) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_REGISTROS);
  const datos = hoja.getDataRange().getValues();
  const headers = datos[0];
  const filas = datos.slice(1).filter(function (f) {
    return periodoDeCelda(f[2]) === String(periodo || '');
  }).map(function (f) {
    const copia = f.slice();
    copia[2] = periodoDeCelda(f[2]);
    copia[0] = esFecha(f[0]) ? Utilities.formatDate(f[0], ZONA_HORARIA, 'yyyy-MM-dd HH:mm') : String(f[0]);
    return copia;
  });
  return { ok: true, headers: headers, rows: filas };
}

// Líneas de una vista y mes, con el número de fila de la hoja para poder editarlas.
function obtenerRegistrosVista(periodo, vista) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_REGISTROS);
  const datos = hoja.getDataRange().getValues();
  const filas = [];
  for (let i = 1; i < datos.length; i++) {
    const f = datos[i];
    if (String(f[1]) !== String(vista || '')) continue;
    if (periodoDeCelda(f[2]) !== String(periodo || '')) continue;
    filas.push({
      fila: i + 1,
      vista: String(f[1]),
      periodo: periodoDeCelda(f[2]),
      jefe: String(f[3] || ''),
      trabajadores: f.slice(4, 14).map(function (t) { return String(t || ''); }),
      producto: String(f[14] || ''),
      clase: String(f[15] || ''),
      largo: String(f[16] || ''),
      diametro: String(f[17] || ''),
      produccion: f[18] === '' ? '' : Number(f[18]),
      recuperacion: f[19] === '' ? '' : Number(f[19]),
      rechazo: f[20] === '' ? '' : Number(f[20]),
      secado: String(f[21] || ''),
      sector: String(f[22] || ''),
      comentario: String(f[23] || '')
    });
  }
  return { ok: true, rows: filas };
}
