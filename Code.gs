/**
 * Registro Inventario - Backend (Google Apps Script)
 *
 * Este script va "pegado" dentro del Google Sheet (Extensiones > Apps Script).
 * Hojas que usa:
 *   - "Registros":  una fila por línea de conteo (todas las vistas).
 *   - "Mantenedor": listas desplegables (columnas A a H, fila 1 = encabezado).
 *
 * Clave del mantenedor: NO va en el código. Se configura en
 *   Configuración del proyecto > Propiedades de la secuencia de comandos
 *   Propiedad: CLAVE_MANTENEDOR   Valor: <tu clave>
 */

const HOJA_REGISTROS = 'Registros';
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

// ---------- Entrada HTTP ----------

function doGet(e) {
  const accion = (e && e.parameter && e.parameter.action) || '';
  let salida;
  try {
    if (accion === 'lists') salida = obtenerListas();
    else if (accion === 'recientes') salida = obtenerRecientes(e.parameter.vista);
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

function registrar(d) {
  const faltantes = [];
  if (!d.vista) faltantes.push('vista');
  if (!d.fecha) faltantes.push('fecha');
  if (!d.jefe) faltantes.push('jefe de área');
  if (!d.producto) faltantes.push('producto');
  if (faltantes.length) {
    return { ok: false, error: 'Falta: ' + faltantes.join(', ') };
  }

  const unidades = [d.produccion, d.recuperacion, d.rechazo];
  const hayUnidades = unidades.some(function (u) {
    return u !== '' && u !== null && u !== undefined && Number(u) > 0;
  });
  if (!hayUnidades) {
    return { ok: false, error: 'Ingresa al menos una cantidad (producción, recuperación o rechazo)' };
  }

  const trabajadores = [];
  for (let i = 0; i < TRABAJADORES_MAX; i++) {
    trabajadores.push((d.trabajadores && d.trabajadores[i]) || '');
  }

  const fila = [
    new Date(),
    d.vista,
    d.fecha,
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

  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_REGISTROS);
  hoja.appendRow(fila);
  return { ok: true };
}

function numeroOVacio(v) {
  if (v === '' || v === null || v === undefined) return '';
  const n = Number(v);
  return isNaN(n) ? '' : n;
}

function obtenerRecientes(vista) {
  const hoja = SpreadsheetApp.getActive().getSheetByName(HOJA_REGISTROS);
  const datos = hoja.getDataRange().getValues();
  if (datos.length <= 1) return { ok: true, headers: datos[0] || [], rows: [] };
  const headers = datos[0];
  const filas = datos.slice(1).filter(function (f) {
    return !vista || String(f[1]) === String(vista);
  });
  return { ok: true, headers: headers, rows: filas.slice(-20).reverse() };
}
