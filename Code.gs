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
var CODE_VERSION = 3;
var SHEET_INVITADOS = 'Invitados';
var SHEET_MESAS = 'Mesas';
var SHEET_RSVP = 'RSVP';

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
  return sh.getRange(1, 1, 1, Math.min(7, sh.getLastColumn())).getValues()[0].map(_norm);
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

/** Fusiona duplicados por nombre (varios teléfonos crean el mismo
    invitado con distinto id): gana el estado más avanzado (si > no > pendiente). */
function _dedupeGuests(guests) {
  function rank(s) { return s === 'si' ? 2 : (s === 'no' ? 1 : 0); }
  var map = {};
  var order = [];
  guests.forEach(function (g) {
    var k = _norm(g.name);
    if (!k) { order.push(g); return; }
    if (!map[k]) { map[k] = g; order.push(g); }
    else {
      var cur = map[k];
      if (rank(g.status) > rank(cur.status)) {
        if (g.mesaId == null) g.mesaId = cur.mesaId;
        map[k] = g;
        order[order.indexOf(cur)] = g;
      }
    }
  });
  return order;
}

/** Borra las filas de datos cuyo valor en col A == id (fila 1 = cabecera). */
function _deleteRowByColA(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return false;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = ids.length - 1; i >= 0; i--) {
    if (String(ids[i][0]) === String(id)) sh.deleteRow(i + 2);
  }
  return true;
}

/** Borra filas de datos donde la columna col (1-based) == val. */
function _deleteRowsWhere(sh, col, val) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var vals = sh.getRange(2, col, last - 1, 1).getValues();
  var n = 0;
  for (var i = vals.length - 1; i >= 0; i--) {
    if (String(vals[i][0]) === String(val)) { sh.deleteRow(i + 2); n++; }
  }
  return n;
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
  var vals = _dataRows(_sheetByRole('guests'), 6);
  return vals
    .map(function (r) {
      return {
        id: Number(r[0]),
        name: String(r[1] || ''),
        phone: String(r[2] || ''),
        seats: parseInt(r[3], 10) || 1,
        status: String(r[4] || 'pendiente'),
        mesaId: r[5] === '' || r[5] === null ? null : Number(r[5])
      };
    });
}

function _readMesas() {
  var vals = _dataRows(_sheetByRole('mesas'), 4);
  return vals
    .map(function (r) {
      var inv = [];
      try { inv = JSON.parse(String(r[3] || '[]')) || []; } catch (e) { inv = []; }
      return {
        id: Number(r[0]),
        nombre: String(r[1] || ''),
        capacidad: parseInt(r[2], 10) || 6,
        invitados: inv
      };
    });
}

function _readRsvp() {
  var vals = _dataRows(_findRsvpSheet(), 7);
  return vals.map(function (r) {
    return {
      timestamp: r[0] instanceof Date ? r[0].toISOString() : String(r[0] || ''),
      nombre: String(r[1] || ''),
      pases: parseInt(r[2], 10) || 1,
      asiste: String(r[3] || 'si'),
      mensaje: String(r[4] || ''),
      guestId: String(r[5] || ''),
      link: String(r[6] || '')
    };
  });
}

/* LECTURA: ?action=state | ?action=rsvpLog (ambas con &key=...) */
function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'state') {
    if (!_keyOk(p.key)) return _json({ ok: false, error: 'key', version: CODE_VERSION });
    return _json({ ok: true, guests: _readGuests(), mesas: _readMesas(), version: CODE_VERSION });
  }
  if (p.action === 'rsvpLog') {
    if (!_keyOk(p.key)) return _json({ ok: false, error: 'key', version: CODE_VERSION });
    return _json({ ok: true, rsvp: _readRsvp(), version: CODE_VERSION });
  }
  return _json({ ok: true, service: 'boda-nicole-rodrigo', version: CODE_VERSION });
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
      _deleteRowByColA(_sheetByRole('guests'), gid);
      var shM = _sheetByRole('mesas');
      var mLast = shM.getLastRow();
      if (mLast > 1) {
        var mInv = shM.getRange(2, 4, mLast - 1, 1).getValues();
        for (var mi = 0; mi < mInv.length; mi++) {
          var arr = [];
          try { arr = JSON.parse(String(mInv[mi][0] || '[]')) || []; } catch (e9) { arr = []; }
          var f = arr.filter(function (x) { return String(x) !== gid; });
          if (f.length !== arr.length) shM.getRange(mi + 2, 4).setValue(JSON.stringify(f));
        }
      }
      _deleteRowsWhere(_findRsvpSheet(), 6, gid);
      return _json({ ok: true });
    }

    if (action === 'deleteMesa') {
      var mid = String(body.id != null ? body.id : (p.id || ''));
      if (!mid) return _json({ ok: false, error: 'id' });
      _deleteRowByColA(_sheetByRole('mesas'), mid);
      var shG2 = _sheetByRole('guests');
      var gLast = shG2.getLastRow();
      if (gLast > 1) {
        var gMesa = shG2.getRange(2, 6, gLast - 1, 1).getValues();
        for (var gi = 0; gi < gMesa.length; gi++) {
          if (String(gMesa[gi][0]) === mid) shG2.getRange(gi + 2, 6).setValue('');
        }
      }
      return _json({ ok: true });
    }

    if (action === 'saveState') {
      var guests = _dedupeGuests(body.guests || []);
      var mesas = body.mesas || [];
      var shG = _sheetByRole('guests');
      if (shG.getLastRow() > 1) shG.getRange(2, 1, shG.getLastRow() - 1, 6).clearContent();
      guests.forEach(function (g) {
        shG.appendRow([g.id, g.name || '', g.phone || '', g.seats || 1, g.status || 'pendiente', g.mesaId == null ? '' : g.mesaId]);
      });
      var shM = _sheetByRole('mesas');
      if (shM.getLastRow() > 1) shM.getRange(2, 1, shM.getLastRow() - 1, 4).clearContent();
      mesas.forEach(function (m) {
        shM.appendRow([m.id, m.nombre || '', m.capacidad || 6, JSON.stringify(m.invitados || [])]);
      });
      return _json({ ok: true });
    }

    return _json({ ok: false, error: 'action' });
  } finally {
    lock.releaseLock();
  }
}
