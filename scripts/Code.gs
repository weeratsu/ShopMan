/**
 * ShopMan — Google Apps Script backend
 * เก็บรายการ shopping list ลง Google Sheet แล้วให้ทุกเครื่อง (คอม/มือถือ) sync กัน
 *
 * วิธี deploy:
 *   1. สร้าง Google Sheet เปล่า 1 ไฟล์ → ก๊อป Spreadsheet ID จาก URL
 *      (ส่วนระหว่าง /d/ กับ /edit  เช่น docs.google.com/spreadsheets/d/<ID>/edit)
 *      แล้วเอามาใส่ใน SHEET_ID ด้านล่าง
 *   2. เปิด https://script.google.com → New project → วางโค้ดนี้
 *   2. Deploy → New deployment → Type: Web app
 *      - Execute as: Me
 *      - Who has access: Anyone
 *   3. ก๊อป Web app URL มาใส่ในแอป (หน้า Settings ของ ShopMan)
 *
 * Data model: 1 Sheet tab ชื่อ "items" — แถวแรกเป็น header
 *   id | text | done | addedAt | doneAt | updatedAt | deleted | order | store
 */

// ⬇️ ใส่ Spreadsheet ID ของ Sheet ที่จะใช้เก็บข้อมูล (ดูจาก URL ของ Sheet)
var SHEET_ID = '1SLnmo2QOyX7VfU7pd4kylkK_GQwvE5oM3JYObV9kjK0';
var SHEET_NAME = 'items';
var HEADERS = ['id', 'text', 'done', 'addedAt', 'doneAt', 'updatedAt', 'deleted', 'order', 'store'];
var STORES_SHEET = 'stores';   // tab แยกเก็บรายชื่อร้าน (คอลัมน์เดียว)

function _sheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  return sh;
}

function _storesSheet() {
  var ss = SpreadsheetApp.openById(SHEET_ID);
  var sh = ss.getSheetByName(STORES_SHEET);
  if (!sh) {
    sh = ss.insertSheet(STORES_SHEET);
    sh.getRange(1, 1).setValue('store');
  }
  return sh;
}

function _readStores() {
  var sh = _storesSheet();
  var rng = sh.getDataRange().getValues();
  var out = [];
  for (var i = 1; i < rng.length; i++) {
    if (rng[i][0]) out.push(String(rng[i][0]));
  }
  return out;
}

function _writeStores(stores) {
  var sh = _storesSheet();
  sh.clear();
  sh.getRange(1, 1).setValue('store');
  if (!stores || !stores.length) return;
  var rows = stores.map(function (s) { return [s]; });
  sh.getRange(2, 1, rows.length, 1).setValues(rows);
}

function _mergeStores(a, b) {
  var set = {}, out = [];
  (a || []).concat(b || []).forEach(function (s) { if (s && !set[s]) { set[s] = 1; out.push(s); } });
  return out;
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
      deleted: row[6] === true || row[6] === 'true' || row[6] === 'TRUE',
      order: Number(row[7]) || 0,
      store: row[8] != null ? String(row[8]) : ''
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
    return [it.id, it.text, !!it.done, it.addedAt || '', it.doneAt || '', it.updatedAt || 0, !!it.deleted, it.order || 0, it.store || ''];
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
  // ส่ง tombstone (deleted) กลับไปด้วย — client ต้องรู้ว่าถูกลบเพื่อ sync การลบข้ามเครื่อง
  return _json({ ok: true, items: _readAll(), stores: _readStores() });
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var body = JSON.parse(e.postData.contents || '{}');
    var client = body.items || [];
    var merged = _merge(_readAll(), client);
    _writeAll(merged);
    var stores = _readStores();
    if (body.stores) { stores = _mergeStores(stores, body.stores); _writeStores(stores); }
    // ส่ง tombstone กลับไปด้วย (ไม่ filter) — client จะได้ลบตาม
    return _json({ ok: true, items: merged, stores: stores });
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
