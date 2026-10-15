// Mi Renta Variable: conector entre la app y esta planilla.
// Pega este código completo en Extensiones > Apps Script y no lo modifiques.
// La CLAVE impide que otra persona lea o escriba tus ventas: no la compartas.
var CLAVE = '__CLAVE__'; // La app genera tu clave: copia el código desde Ajustes, no desde este archivo.
var COLS = ['id', 'fecha', 'tipo', 'plan', 'uf', 'estado', 'nicho', 'medio_pago', 'beneficiarios', 'nombre', 'rut', 'telefono', 'correo', 'notas', 'actualizado', 'borrado', 'json'];

function doGet(e) {
  return responder_(function () {
    revisar_(e && e.parameter ? e.parameter.k : '');
    return todo_();
  });
}

function doPost(e) {
  return responder_(function () {
    var cuerpo = JSON.parse(e.postData.contents);
    revisar_(cuerpo.k);
    var lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      (cuerpo.ops || []).forEach(aplicar_);
      SpreadsheetApp.flush();
    } finally {
      lock.releaseLock();
    }
    return todo_();
  });
}

function responder_(fn) {
  var out;
  try {
    out = fn();
    out.ok = true;
  } catch (err) {
    out = { ok: false, error: String(err && err.message ? err.message : err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function revisar_(k) {
  if (!k || k !== CLAVE) throw new Error('Clave incorrecta');
}

function hoja_(nombre, encabezado, textoCols) {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    textoCols.forEach(function (c) { sh.getRange(c).setNumberFormat('@'); });
    sh.getRange(1, 1, 1, encabezado.length).setValues([encabezado]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function hojaVentas_() {
  var sh = hoja_('Ventas', COLS, ['A:D', 'F:H', 'J:O', 'Q:Q']);
  if (!sh.isColumnHiddenByUser(17)) {
    sh.getRange('E:E').setNumberFormat('0.00');
    sh.hideColumns(17);
  }
  return sh;
}
function hojaMeses_() { return hoja_('Meses', ['mes', 'actualizado', 'json'], ['A:C']); }
function hojaPerfil_() { return hoja_('Perfil', ['clave', 'valor'], ['A:B']); }

function esVerdadero_(x) { return x === true || String(x).toUpperCase() === 'TRUE'; }

function fila_(sh, valor) {
  var n = sh.getLastRow() - 1;
  if (n < 1) return 0;
  var vals = sh.getRange(2, 1, n, 1).getValues();
  for (var i = 0; i < n; i++) {
    if (String(vals[i][0]) === String(valor)) return i + 2;
  }
  return 0;
}

function todo_() {
  var sv = hojaVentas_(), sm = hojaMeses_(), sp = hojaPerfil_();
  var ventas = [];
  var vv = sv.getDataRange().getValues();
  for (var r = 1; r < vv.length; r++) {
    if (esVerdadero_(vv[r][15])) continue;
    try { if (vv[r][16]) ventas.push(JSON.parse(vv[r][16])); } catch (e) {}
  }
  var meses = {};
  var mv = sm.getDataRange().getValues();
  for (var i = 1; i < mv.length; i++) {
    try { if (mv[i][0]) meses[String(mv[i][0])] = JSON.parse(mv[i][2]); } catch (e) {}
  }
  var perfil = {};
  var pv = sp.getDataRange().getValues();
  for (var j = 1; j < pv.length; j++) {
    if (pv[j][0]) perfil[String(pv[j][0])] = pv[j][1];
  }
  return { ventas: ventas, meses: meses, perfil: perfil, planilla: SpreadsheetApp.getActive().getUrl(), hora: new Date().toISOString() };
}

function aplicar_(op) {
  if (!op) return;
  if (op.op === 'venta') guardarVenta_(op.v);
  else if (op.op === 'borrar') borrarVenta_(op.id, op.t);
  else if (op.op === 'mes') guardarMes_(op.mes, op.cfg, op.t);
  else if (op.op === 'perfil') guardarPerfil_(op.clave, op.valor);
}

function guardarVenta_(v) {
  if (!v || !v.id) return;
  var sh = hojaVentas_();
  var r = fila_(sh, v.id);
  if (r) {
    var prev = sh.getRange(r, 15, 1, 2).getValues()[0];
    if (esVerdadero_(prev[1])) return;
    if (prev[0] && v.actualizado && String(prev[0]) > String(v.actualizado)) return;
  }
  var c = v.cliente || {};
  var fila = [
    String(v.id), v.fecha || '', v.tipo || '', v.planNombre || '', Number(v.uf) || 0,
    v.tipo === 'isapre' ? (v.etapa || '') : (v.estado || ''), v.nicho ? 'sí' : 'no', v.medioPago || '',
    Number(v.beneficiarios) || 1, c.nombre || '', c.rut || '', c.telefono || '', c.correo || '', v.notas || '',
    v.actualizado || new Date().toISOString(), false, JSON.stringify(v)
  ];
  if (r) sh.getRange(r, 1, 1, fila.length).setValues([fila]);
  else sh.appendRow(fila);
}

function borrarVenta_(id, t) {
  if (!id) return;
  var sh = hojaVentas_();
  var r = fila_(sh, id);
  if (r) sh.getRange(r, 16).setValue(true);
  else sh.appendRow([String(id), '', '', '', 0, '', '', '', 0, '', '', '', '', '', t || '', true, '']);
}

function guardarMes_(mes, cfg, t) {
  if (!mes) return;
  var sh = hojaMeses_();
  var r = fila_(sh, mes);
  if (r) {
    var prev = sh.getRange(r, 2).getValue();
    if (prev && t && String(prev) > String(t)) return;
    sh.getRange(r, 1, 1, 3).setValues([[String(mes), t || '', JSON.stringify(cfg || {})]]);
  } else {
    sh.appendRow([String(mes), t || '', JSON.stringify(cfg || {})]);
  }
}

function guardarPerfil_(clave, valor) {
  if (!clave) return;
  var sh = hojaPerfil_();
  var r = fila_(sh, clave);
  if (r) sh.getRange(r, 2).setValue(String(valor));
  else sh.appendRow([String(clave), String(valor)]);
}
