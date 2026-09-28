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
  var sh = _sheet(SHEET_INVITADOS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 6).getValues();
  return vals
    .filter(function (r) { return r[0] !== '' && r[0] !== null; })
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
  var sh = _sheet(SHEET_MESAS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 4).getValues();
  return vals
    .filter(function (r) { return r[0] !== '' && r[0] !== null; })
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
  var sh = _sheet(SHEET_RSVP);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var vals = sh.getRange(2, 1, last - 1, 7).getValues();
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
    if (!_keyOk(p.key)) return _json({ ok: false, error: 'key' });
    return _json({ ok: true, guests: _readGuests(), mesas: _readMesas() });
  }
  if (p.action === 'rsvpLog') {
    if (!_keyOk(p.key)) return _json({ ok: false, error: 'key' });
    return _json({ ok: true, rsvp: _readRsvp() });
  }
  return _json({ ok: true, service: 'boda-nicole-rodrigo' });
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
      _sheet(SHEET_RSVP).appendRow([
        new Date(),
        body.nombre || p.nombre || p.name || '',
        parseInt(body.pases || p.pases || p.party, 10) || 0,
        asiste,
        body.mensaje || p.mensaje || p.message || '',
        body.guestId || p.guestId || p.id || '',
        body.link || p.link || ''
      ]);
      return _json({ ok: true });
    }

    /* --- Todo lo demás exige la clave compartida --- */
    if (!_keyOk(body.key)) return _json({ ok: false, error: 'key' });

    if (action === 'saveState') {
      var guests = body.guests || [];
      var mesas = body.mesas || [];
      var shG = _sheet(SHEET_INVITADOS);
      if (shG.getLastRow() > 1) shG.getRange(2, 1, shG.getLastRow() - 1, 6).clearContent();
      guests.forEach(function (g) {
        shG.appendRow([g.id, g.name || '', g.phone || '', g.seats || 1, g.status || 'pendiente', g.mesaId == null ? '' : g.mesaId]);
      });
      var shM = _sheet(SHEET_MESAS);
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
