/**
 * ShopMan — Google Apps Script backend
 * เก็บรายการ shopping list ลง Google Sheet แล้วให้ทุกเครื่อง (คอม/มือถือ) sync กัน
 *
 * วิธี deploy:
 *   1. เปิด https://script.google.com → New project → วางโค้ดนี้
 *   2. Deploy → New deployment → Type: Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 *   3. ก๊อป Web app URL มาใส่ในแอป (หน้า Settings ของ ShopMan)
 *
 * Data model: 1 Sheet tab ชื่อ "items" — แถวแรกเป็น header
 *   id | text | done | addedAt | doneAt | updatedAt | deleted
 */

var SHEET_NAME = 'items';
var HEADERS = ['id', 'text', 'done', 'addedAt', 'doneAt', 'updatedAt', 'deleted'];

function _sheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sh;
}

function _readAll() {
  var sh = _sheet();
  var rng = sh.getDataRange().getValues();
  if (rng.length < 2) return [];
  var out = [];
  for (var i = 1; i < rng.length; i++) {
    var row = rng[i];
    if (!row[0]) continue;
    out.push({
      id: String(row[0]),
      text: String(row[1]),
      done: row[2] === true || row[2] === 'true' || row[2] === 'TRUE',
      addedAt: Number(row[3]) || null,
      doneAt: Number(row[4]) || null,
      updatedAt: Number(row[5]) || 0,
      deleted: row[6] === true || row[6] === 'true' || row[6] === 'TRUE'
    });
  }
  return out;
}

function _writeAll(items) {
  var sh = _sheet();
  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  if (!items.length) return;
  var rows = items.map(function (it) {
    return [it.id, it.text, !!it.done, it.addedAt || '', it.doneAt || '', it.updatedAt || 0, !!it.deleted];
  });
  sh.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
}

/** merge แบบ last-write-wins ด้วย updatedAt; deleted = tombstone */
function _merge(serverItems, clientItems) {
  var map = {};
  serverItems.forEach(function (it) { map[it.id] = it; });
  clientItems.forEach(function (c) {
    var s = map[c.id];
    if (!s || (c.updatedAt || 0) >= (s.updatedAt || 0)) map[c.id] = c;
  });
  return Object.keys(map).map(function (k) { return map[k]; });
}

function doGet() {
  return _json({ ok: true, items: _readAll().filter(function (x) { return !x.deleted; }) });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var client = body.items || [];
    var merged = _merge(_readAll(), client);
    _writeAll(merged);
    return _json({ ok: true, items: merged.filter(function (x) { return !x.deleted; }) });
  } catch (err) {
    return _json({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function _json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
