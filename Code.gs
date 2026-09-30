/**********************************************************************
 * Boda Nicole & Rodrigo — Backend Google Sheets
 *
 * INSTALACIÓN (5 min, gratis):
 * 1. Abre tu Google Sheet y ve a Extensiones > Apps Script.
 * 2. Borra el contenido y pega este archivo completo.
 * 3. Ejecuta una vez la función `setup` (acepta los permisos).
 * 4. Implementar > Nueva implementación > Aplicación web:
 *    - Ejecutar como: Yo
 *    - Acceso: Cualquier usuario (anónimo)  ← necesario para el RSVP
 * 5. Copia la URL /exec y pégala en:
 *    - index.html → GOOGLE_APPS_SCRIPT_URL
 *    - admin.html → Configuración > Google Apps Script URL
 *
 * HOJAS QUE USA (se crean solas con `setup`):
 *  Invitados: id | nombre | telefono | pases | estado | mesaId
 *  Mesas:     id | nombre | capacidad | invitados (JSON de ids)
 *  RSVP:      timestamp | nombre | pases | asiste | mensaje | guestId | link
 **********************************************************************/

var SHARED_KEY = 'nicoleyrodrigo2026';
/* Sube este número cada vez que cambies este archivo para que el admin
   detecte despliegues viejos. */
var CODE_VERSION = 5;
var SHEET_INVITADOS = 'Invitados';
var SHEET_MESAS = 'Mesas';
var SHEET_RSVP = 'RSVP';

/* Cabeceras oficiales de cada hoja. */
var HEADERS = {
  guests: ['id', 'nombre', 'telefono', 'pases', 'estado', 'mesaId'],
  mesas: ['id', 'nombre', 'capacidad', 'invitados'],
  rsvp: ['timestamp', 'nombre', 'pases', 'asiste', 'mensaje', 'guestId', 'link']
};

/* Normaliza cabeceras: minúsculas, sin espacios ni tildes. */
function _normH(s) {
  return String(s == null ? '' : s).replace(/^\s+|\s+$/g, '').toLowerCase()
    .replace(/[áàäâ]/g, 'a').replace(/[éèëê]/g, 'e').replace(/[íìïî]/g, 'i')
    .replace(/[óòöô]/g, 'o').replace(/[úùüû]/g, 'u').replace(/ñ/g, 'n');
}

function _sheet(name) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var want = String(name).replace(/^\s+|\s+$/g, '').toLowerCase();
  var sheets = ss.getSheets();
  for (var i = 0; i < sheets.length; i++) {
    var n = String(sheets[i].getName()).replace(/^\s+|\s+$/g, '').toLowerCase();
    if (n === want) return sheets[i];
  }
  return ss.insertSheet(name);
}

function _json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function _keyOk(k) {
  return String(k || '') === SHARED_KEY;
}

function _norm(s) {
  return String(s == null ? '' : s).replace(/^\s+|\s+$/g, '').toLowerCase();
}

/** Hoja RSVP: nombre exacto, tolerante, con encabezado RSVP o cualquier hoja con datos. */
function _findRsvpSheet() {
  return _sheetByRole('rsvp');
}

/** Localiza hojas por rol aunque cambien mayúsculas, espacios o nombre:
    invitados = la que tenga telefono/estado; mesas = capacidad;
    rsvp = asiste+nombre, o cualquier otra hoja con datos. */
function _sheetByRole(role) {
  var want = role === 'guests' ? 'invitados' : (role === 'mesas' ? 'mesas' : 'rsvp');
  var keys = role === 'guests' ? ['telefono', 'estado']
           : (role === 'mesas' ? ['capacidad'] : ['asiste']);
  var sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
  var i, sh;
  for (i = 0; i < sheets.length; i++) {
    if (_norm(sheets[i].getName()) === want) return sheets[i];
  }
  for (i = 0; i < sheets.length; i++) {
    if (_headerHas(sheets[i], keys)) return sheets[i];
  }
  if (role === 'rsvp') {
    for (i = 0; i < sheets.length; i++) {
      if (_headerHas(sheets[i], ['nombre', 'asiste'])) return sheets[i];
    }
    var gN = _sheetByRole('guests').getName();
    var mN = _sheetByRole('mesas').getName();
    for (i = 0; i < sheets.length; i++) {
      sh = sheets[i];
      if (sh.getName() === gN || sh.getName() === mN) continue;
      if (_dataRows(sh, 7).length) return sh;
    }
  }
  return _sheet(role === 'guests' ? SHEET_INVITADOS : (role === 'mesas' ? SHEET_MESAS : SHEET_RSVP));
}

function _headerOf(sh) {
  if (sh.getLastRow() < 1) return [];
  return sh.getRange(1, 1, 1, Math.min(10, sh.getLastColumn())).getValues()[0].map(_normH);
}

function _headerHas(sh, words) {
  var head = _headerOf(sh);
  return words.every(function (w) { return head.indexOf(w) !== -1; });
}

/** Filas de datos: salta el encabezado si existe y quita filas vacías. */
function _dataRows(sh, ncols) {
  var last = sh.getLastRow();
  if (last < 1) return [];
  var vals = sh.getRange(1, 1, last, ncols).getValues();
  if (vals.length) {
    var h = vals[0].map(_norm).join('|');
    if (h.indexOf('nombre') !== -1 || h.indexOf('timestamp') !== -1 ||
        h.indexOf('capacidad') !== -1 || h.indexOf('invitados') !== -1 ||
        (h.indexOf('id') !== -1 && h.indexOf('telefono') !== -1)) vals.shift();
  }
  return vals.filter(function (r) { return r[0] !== '' && r[0] !== null; });
}

/** Papelera de borrados compartida: ids 'g:123' / 'm:456'.
    Lo borrado no resucita aunque otro teléfono lo re-suba. */
function _borradosList() {
  var sh = _sheet('Borrados');
  _ensureHeader(sh, ['id']);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 1).getValues().map(function (r) { return String(r[0]); });
}

function _borradosAdd(tags) {
  var sh = _sheet('Borrados');
  _ensureHeader(sh, ['id']);
  var have = {};
  _borradosList().forEach(function (t) { have[t] = 1; });
  tags.forEach(function (t) {
    t = String(t);
    if (t && !have[t]) { sh.appendRow([t]); have[t] = 1; }
  });
}

/** Dedupe por id (el último gana). NUNCA por nombre: los homónimos son válidos. */
function _dedupeById(rows) {
  var map = {};
  var order = [];
  rows.forEach(function (r) {
    var k = String(r.id);
    if (!map[k]) { map[k] = r; order.push(r); }
    else { order[order.indexOf(map[k])] = r; map[k] = r; }
  });
  return order;
}

/** Garantiza la fila 1 de cabecera. Si la fila 1 trae datos (hoja editada
    a mano sin cabecera), inserta la cabecera arriba en vez de asumirla:
    así la fila 1 nunca queda como dato inmortal. */
function _ensureHeader(sh, headers) {
  if (sh.getLastRow() < 1) { sh.appendRow(headers); return; }
  var a1 = _normH(sh.getRange(1, 1).getValue());
  if (a1 === 'id' || a1 === 'timestamp') return;
  sh.insertRowBefore(1);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]);
}

/** Mapea cada nombre lógico a su columna real según la fila 1.
    Tolera columnas extra o movidas; si falta, usa la posición oficial. */
function _colMap(sh, headers) {
  var n = Math.max(headers.length, sh.getLastColumn());
  var head = n > 0 ? sh.getRange(1, 1, 1, n).getValues()[0].map(_normH) : [];
  var map = {};
  headers.forEach(function (h, i) {
    var idx = head.indexOf(_normH(h));
    map[h] = idx !== -1 ? idx + 1 : i + 1;
  });
  return map;
}

/** Ids de la columna A (tope 1000) para diagnóstico. */
function _colAIds(sh) {
  var last = sh.getLastRow();
  if (last < 2) return [];
  var n = Math.min(last - 1, 1000);
  return sh.getRange(2, 1, n, 1).getValues().map(function (r) { return String(r[0]); });
}

function _trimId(v) { return String(v == null ? '' : v).replace(/^\s+|\s+$/g, ''); }

/** Borra TODAS las filas de datos con ese id en su columna (tolera espacios).
    Devuelve cuántas borró. */
function _deleteRowById(sh, headers, id) {
  _ensureHeader(sh, headers);
  var col = _colMap(sh, headers).id;
  var want = _trimId(id);
  if (!want) return 0;
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var vals = sh.getRange(2, col, last - 1, 1).getValues();
  var n = 0;
  for (var i = vals.length - 1; i >= 0; i--) {
    if (_trimId(vals[i][0]) === want) { sh.deleteRow(i + 2); n++; }
  }
  return n;
}

/** Borra filas de datos donde la columna col (1-based) == val. Devuelve conteo. */
function _deleteRowsWhere(sh, col, val) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var want = _trimId(val);
  if (!want) return 0;
  var vals = sh.getRange(2, col, last - 1, 1).getValues();
  var n = 0;
  for (var i = vals.length - 1; i >= 0; i--) {
    if (_trimId(vals[i][0]) === want) { sh.deleteRow(i + 2); n++; }
  }
  return n;
}

/** Limpia todas las filas de datos (fila 1 = cabecera) en todo el ancho. */
function _clearData(sh) {
  var last = sh.getLastRow();
  if (last < 2) return;
  sh.getRange(2, 1, last - 1, Math.min(Math.max(4, sh.getLastColumn()), 50)).clearContent();
}

/** Arma una fila en las columnas reales del mapa. */
function _mappedRow(map, headers, vals) {
  var width = 0;
  headers.forEach(function (h) { if (map[h] > width) width = map[h]; });
  var row = [];
  for (var i = 0; i < width; i++) row.push('');
  headers.forEach(function (h) { row[map[h] - 1] = vals[h] == null ? '' : vals[h]; });
  return row;
}

/** Fusiona mesas con id repetido: conserva la que tenga más invitados. */
function _dedupeMesas(mesas) {
  var map = {};
  var order = [];
  mesas.forEach(function (mm) {
    var k = String(mm.id);
    if (!map[k]) { map[k] = mm; order.push(mm); }
    else {
      var cur = map[k];
      var a = (cur.invitados || []).length, b = (mm.invitados || []).length;
      if (b > a) { order[order.indexOf(cur)] = mm; map[k] = mm; }
    }
  });
  return order;
}

/** Crea las 3 hojas con encabezados. Ejecutar una vez a mano. */
function setup() {
  var inv = _sheet(SHEET_INVITADOS);
  if (inv.getLastRow() === 0) inv.appendRow(['id', 'nombre', 'telefono', 'pases', 'estado', 'mesaId']);
  var mes = _sheet(SHEET_MESAS);
  if (mes.getLastRow() === 0) mes.appendRow(['id', 'nombre', 'capacidad', 'invitados']);
  var rsv = _sheet(SHEET_RSVP);
  if (rsv.getLastRow() === 0) rsv.appendRow(['timestamp', 'nombre', 'pases', 'asiste', 'mensaje', 'guestId', 'link']);
}

function _readGuests() {
  var sh = _sheetByRole('guests');
  _ensureHeader(sh, HEADERS.guests);
  var m = _colMap(sh, HEADERS.guests);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  return vals
    .filter(function (r) { return _trimId(r[m.id - 1]) !== ''; })
    .map(function (r) {
      var mesaRaw = r[m.mesaId - 1];
      return {
        id: Number(r[m.id - 1]),
        name: String(r[m.nombre - 1] || ''),
        phone: String(r[m.telefono - 1] || ''),
        seats: parseInt(r[m.pases - 1], 10) || 1,
        status: String(r[m.estado - 1] || 'pendiente'),
        mesaId: mesaRaw === '' || mesaRaw == null ? null : Number(mesaRaw)
      };
    });
}

function _readMesas() {
  var sh = _sheetByRole('mesas');
  _ensureHeader(sh, HEADERS.mesas);
  var m = _colMap(sh, HEADERS.mesas);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  return vals
    .filter(function (r) { return _trimId(r[m.id - 1]) !== ''; })
    .map(function (r) {
      var inv = [];
      try { inv = JSON.parse(String(r[m.invitados - 1] || '[]')) || []; } catch (e) { inv = []; }
      return {
        id: Number(r[m.id - 1]),
        nombre: String(r[m.nombre - 1] || ''),
        capacidad: parseInt(r[m.capacidad - 1], 10) || 6,
        invitados: inv
      };
    });
}

function _readRsvp() {
  var sh = _findRsvpSheet();
  _ensureHeader(sh, HEADERS.rsvp);
  var m = _colMap(sh, HEADERS.rsvp);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, sh.getLastColumn()).getValues();
  return vals
    .filter(function (r) { return _trimId(r[m.nombre - 1]) !== '' || _trimId(r[m.timestamp - 1]) !== ''; })
    .map(function (r) {
      return {
        timestamp: r[m.timestamp - 1] instanceof Date ? r[m.timestamp - 1].toISOString() : String(r[m.timestamp - 1] || ''),
        nombre: String(r[m.nombre - 1] || ''),
        pases: parseInt(r[m.pases - 1], 10) || 1,
        asiste: String(r[m.asiste - 1] || 'si'),
        mensaje: String(r[m.mensaje - 1] || ''),
        guestId: String(r[m.guestId - 1] || ''),
        link: String(r[m.link - 1] || '')
      };
    });
}

/* LECTURA: ?action=state | ?action=rsvpLog (ambas con &key=...).
   Toma el lock brevemente para no leer a mitad de una reescritura. */
function doGet(e) {
  var lock = LockService.getScriptLock();
  var got = false;
  try { lock.waitLock(8000); got = true; } catch (err) { got = false; }
  try {
    var p = (e && e.parameter) || {};
    if (p.action === 'state') {
      if (!_keyOk(p.key)) return _json({ ok: false, error: 'key', version: CODE_VERSION });
      return _json({ ok: true, guests: _readGuests(), mesas: _readMesas(), borrados: _borradosList(), version: CODE_VERSION });
    }
    if (p.action === 'rsvpLog') {
      if (!_keyOk(p.key)) return _json({ ok: false, error: 'key', version: CODE_VERSION });
      return _json({ ok: true, rsvp: _readRsvp(), version: CODE_VERSION });
    }
    return _json({ ok: true, service: 'boda-nicole-rodrigo', version: CODE_VERSION });
  } finally {
    if (got) lock.releaseLock();
  }
}

/* ESCRITURA: JSON {action, key, ...} o formulario clásico */
function doPost(e) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(15000); } catch (err) { return _json({ ok: false, error: 'lock' }); }
  try {
    var body = {};
    try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
    catch (err2) { body = {}; }
    var p = (e && e.parameter) || {};
    var action = body.action || p.action || '';

    /* --- RSVP desde la invitación (endpoint público, solo agrega) --- */
    if (action === 'rsvp' || body.nombre || p.name) {
      var choice = body.asiste || p.asiste || p.choice || '';
      var asiste = (choice === 'si' || choice === 'accepts') ? 'si' : 'no';
      var gid = String(body.guestId || p.guestId || p.id || '');
      _findRsvpSheet().appendRow([
        new Date(),
        body.nombre || p.nombre || p.name || '',
        parseInt(body.pases || p.pases || p.party, 10) || 0,
        asiste,
        body.mensaje || p.mensaje || p.message || '',
        gid,
        body.link || p.link || ''
      ]);
      /* Auto-confirmación: si el link traía id, marca al invitado */
      if (gid !== '') {
        var shG = _sheetByRole('guests');
        var lastG = shG.getLastRow();
        if (lastG > 1) {
          var ids = shG.getRange(2, 1, lastG - 1, 1).getValues();
          for (var k = 0; k < ids.length; k++) {
            if (String(ids[k][0]) === gid) {
              shG.getRange(k + 2, 5).setValue(asiste);
              break;
            }
          }
        }
      }
      return _json({ ok: true });
    }

    /* --- Todo lo demás exige la clave compartida --- */
    if (!_keyOk(body.key)) return _json({ ok: false, error: 'key' });

    /* --- Borrado por id (el admin lo llama al eliminar + sync de convergencia).
           Garantiza que el borrado pega aunque el sync completo falle o compita. --- */
    if (action === 'deleteGuest') {
      var gid = String(body.id != null ? body.id : (p.id || ''));
      if (!gid) return _json({ ok: false, error: 'id' });
      _borradosAdd(['g:' + gid]);
      var shG0 = _sheetByRole('guests');
      var gDel = _deleteRowById(shG0, HEADERS.guests, gid);
      var shM = _sheetByRole('mesas');
      _ensureHeader(shM, HEADERS.mesas);
      var mMap = _colMap(shM, HEADERS.mesas);
      var mLast = shM.getLastRow();
      if (mLast > 1) {
        var mInv = shM.getRange(2, mMap.invitados, mLast - 1, 1).getValues();
        for (var mi = 0; mi < mInv.length; mi++) {
          var arr = [];
          try { arr = JSON.parse(String(mInv[mi][0] || '[]')) || []; } catch (e9) { arr = []; }
          var f = arr.filter(function (x) { return _trimId(x) !== _trimId(gid); });
          if (f.length !== arr.length) shM.getRange(mi + 2, mMap.invitados).setValue(JSON.stringify(f));
        }
      }
      var shR0 = _findRsvpSheet();
      _ensureHeader(shR0, HEADERS.rsvp);
      var rDel = _deleteRowsWhere(shR0, _colMap(shR0, HEADERS.rsvp).guestId, gid);
      return _json({ ok: true, sheet: shG0.getName(), deletedGuests: gDel, deletedRsvp: rDel });
    }

    if (action === 'deleteMesa') {
      var mid = String(body.id != null ? body.id : (p.id || ''));
      if (!mid) return _json({ ok: false, error: 'id' });
      _borradosAdd(['m:' + mid]);
      var shM0 = _sheetByRole('mesas');
      var mDel = _deleteRowById(shM0, HEADERS.mesas, mid);
      var shG2 = _sheetByRole('guests');
      _ensureHeader(shG2, HEADERS.guests);
      var gMap = _colMap(shG2, HEADERS.guests);
      var gLast = shG2.getLastRow();
      var cleared = 0;
      if (gLast > 1) {
        var gMesa = shG2.getRange(2, gMap.mesaId, gLast - 1, 1).getValues();
        for (var gi = 0; gi < gMesa.length; gi++) {
          if (_trimId(gMesa[gi][0]) === _trimId(mid)) { shG2.getRange(gi + 2, gMap.mesaId).setValue(''); cleared++; }
        }
      }
      return _json({ ok: true, sheet: shM0.getName(), deletedMesas: mDel, guestsCleared: cleared });
    }

    /* --- Diagnóstico: qué archivo y pestañas ve el backend (para depurar) --- */
    if (action === 'debug') {
      var ss = SpreadsheetApp.getActiveSpreadsheet();
      var tabs = ss.getSheets().map(function (sh) {
        return { name: sh.getName(), rows: sh.getLastRow(), cols: sh.getLastColumn(), head: _headerOf(sh) };
      });
      var dg = _sheetByRole('guests');
      var dm = _sheetByRole('mesas');
      return _json({ ok: true, version: CODE_VERSION, file: ss.getName(),
        tabs: tabs, guestsTab: dg.getName(), guestsCols: _colMap(dg, HEADERS.guests),
        mesasTab: dm.getName(), mesasCols: _colMap(dm, HEADERS.mesas),
        guestIds: _colAIds(dg), mesaIds: _colAIds(dm), borrados: _borradosList() });
    }

    /* saveState con UNIÓN por id (multi-teléfono seguro):
       lo que subes se fusiona con lo que hay, nada se pisa;
       los ids en la papelera jamás vuelven. */
    if (action === 'saveState') {
      var incomingDel = body.deletedIds || [];
      _borradosAdd(incomingDel);
      var gone = {};
      _borradosList().forEach(function (t) { gone[t] = 1; });
      var guests = _dedupeById(body.guests || []);
      var mesas = _dedupeMesas(body.mesas || []);
      var shG = _sheetByRole('guests');
      _ensureHeader(shG, HEADERS.guests);
      var gm = _colMap(shG, HEADERS.guests);
      var gById = {};
      _readGuests().forEach(function (g) { gById['g:' + String(g.id)] = g; });
      guests.forEach(function (g) {
        gById['g:' + String(g.id)] = {
          id: g.id, name: g.name || '', phone: g.phone || '',
          seats: g.seats || 1, status: g.status || 'pendiente',
          mesaId: (g.mesaId == null || g.mesaId === '') ? '' : g.mesaId
        };
      });
      _clearData(shG);
      Object.keys(gById).forEach(function (k) {
        if (gone[k]) return;
        var g = gById[k];
        shG.appendRow(_mappedRow(gm, HEADERS.guests, {
          id: g.id, nombre: g.name, telefono: g.phone,
          pases: g.seats, estado: g.status, mesaId: g.mesaId
        }));
      });
      var shM = _sheetByRole('mesas');
      _ensureHeader(shM, HEADERS.mesas);
      var mm = _colMap(shM, HEADERS.mesas);
      var mById = {};
      _readMesas().forEach(function (m) { mById['m:' + String(m.id)] = m; });
      mesas.forEach(function (m) {
        mById['m:' + String(m.id)] = {
          id: m.id, nombre: m.nombre || '', capacidad: m.capacidad || 6,
          invitados: m.invitados || []
        };
      });
      _clearData(shM);
      Object.keys(mById).forEach(function (k) {
        if (gone[k]) return;
        var m = mById[k];
        shM.appendRow(_mappedRow(mm, HEADERS.mesas, {
          id: m.id, nombre: m.nombre, capacidad: m.capacidad,
          invitados: JSON.stringify(m.invitados || [])
        }));
      });
      return _json({ ok: true });
    }

    return _json({ ok: false, error: 'action' });
  } finally {
    lock.releaseLock();
  }
}
