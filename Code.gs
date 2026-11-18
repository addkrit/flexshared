/**
 * Flex Shared (ของคุณเอง) — ตัวเก็บ Flex บน Google Sheet
 *
 * วิธีใช้: สร้าง Google Sheet ใหม่ > ส่วนขยาย > Apps Script > วางไฟล์นี้ทั้งไฟล์
 * แก้ LIFF_ID ด้านล่าง > ทำให้ใช้งานได้ (Deploy) > เว็บแอป
 *   - เรียกใช้ในฐานะ: ฉัน (Me)
 *   - ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน (Anyone)
 * แล้วเอา URL ที่ลงท้าย /exec ไปใส่ใน config.js
 *
 * สิ่งที่แก้จากของเดิม (botmarketing):
 *  - ใช้ Sheet ของคุณเอง (ของเดิมเจ้าของปิดสิทธิ์ → บันทึก/โหลด/ลบ ไม่ได้ทุกคน)
 *  - ยืนยันตัวตนด้วย LINE ID token จริง (ของเดิมใครรู้ id ก็ลบ/แก้ของคนอื่นได้)
 *  - Flex ยาวเกิน 50,000 ตัวอักษร (ขีดจำกัดต่อช่องของ Sheet) ก็บันทึกได้ — แบ่งเก็บหลายช่อง
 *  - กันข้อมูลถูก Sheet ตีความเป็นสูตร/ตัวเลข
 *  - ล็อกตอนเขียน กันบันทึกพร้อมกันแล้วข้อมูลทับกัน
 */

// ===== ตั้งค่า (แก้บรรทัดนี้บรรทัดเดียว) =====
const LIFF_ID = '2012675889-dcyNM9eN';   // LIFF ของช่อง Flex Shared (Provider: Digital)
// Google Sheet ที่เก็บข้อมูล (Flex Shared DB ของ goimage55) — ถ้าวางโค้ดนี้ใน Apps Script ของ Sheet โดยตรง ปล่อยว่าง '' ได้
const SHEET_ID = '19VKD_ktM14vSIKM7IQaP4kzwzORwhX3bH1MWkbxkVnM';

// ===== ค่าคงที่ (ไม่ต้องแก้) =====
const SHEET_NAME = 'flex';
const CHUNK = 40000;          // ตัวอักษรต่อช่อง (Sheet รับได้สูงสุด 50,000)
const FIRST_DATA_COL = 6;     // A id | B uid | C name | D created | E updated | F.. ข้อมูล JSON
const P = '~';                // prefix กัน Sheet แปลงข้อความเป็นสูตร/ตัวเลข
const MAX_FLEX_CHARS = 200000;

function doGet(e) { return handle_((e && e.parameter) || {}); }
function doPost(e) { return handle_((e && e.parameter) || {}); }

function handle_(p) {
  try {
    const q = p.q || (p.id ? 'one' : '');
    let out;
    if (q === 'ping') out = { result: 'success', message: 'ok' };
    else if (q === 'one') out = one_(p.id);
    else {
      const uid = verify_(p.token);
      if (q === 'get') out = list_(uid);
      else if (q === 'put') out = put_(uid, p.name, p.flex);
      else if (q === 'update') out = update_(uid, p.id, p.flex, p.name);
      else if (q === 'delete') out = del_(uid, p.id);
      else out = { result: 'error', message: 'คำสั่งไม่ถูกต้อง: ' + q };
    }
    return json_(out);
  } catch (err) {
    return json_({ result: 'error', message: String((err && err.message) || err) });
  }
}

// ---------- ยืนยันตัวตนด้วย LINE ID token ----------
function verify_(token) {
  if (!token) throw new Error('ไม่มี token — กรุณาล็อกอิน LINE ใหม่');
  const key = 'tok_' + Utilities.base64EncodeWebSafe(
    Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token)).slice(0, 40);
  const cache = CacheService.getScriptCache();
  const hit = cache.get(key);
  if (hit) return hit;
  const channelId = String(LIFF_ID).split('-')[0];
  const res = UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', {
    method: 'post',
    payload: { id_token: token, client_id: channelId },
    muteHttpExceptions: true,
  });
  const body = JSON.parse(res.getContentText() || '{}');
  if (res.getResponseCode() !== 200 || !body.sub) {
    throw new Error('ยืนยันตัวตนไม่ผ่าน (' + (body.error_description || body.error || res.getResponseCode()) +
      ') — ปิดแล้วเปิดหน้าใหม่ หรือเช็ก LIFF_ID ใน Code.gs');
  }
  const ttl = Math.max(60, Math.min(600, Math.floor(body.exp - Date.now() / 1000) - 30));
  cache.put(key, body.sub, ttl);
  return body.sub;
}

// ---------- Sheet ----------
function sheet_() {
  const ss = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, FIRST_DATA_COL).setValues([['id', 'uid', 'name', 'created', 'updated', 'flex (แบ่งหลายช่อง)']]);
    sh.setFrozenRows(1);
  }
  return sh;
}

function rows_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const width = Math.max(sh.getLastColumn(), FIRST_DATA_COL);
  return sh.getRange(2, 1, last - 1, width).getValues().map(function (r, i) {
    return { row: i + 2, values: r };
  });
}

function un_(v) { v = String(v == null ? '' : v); return v.charAt(0) === P ? v.slice(1) : v; }
function pre_(v) { return P + String(v); }

function flexOf_(values) {
  let s = '';
  for (let c = FIRST_DATA_COL - 1; c < values.length; c++) {
    if (values[c] === '' || values[c] == null) break;
    s += un_(values[c]);
  }
  return s;
}

function item_(r) {
  return { id: un_(r.values[0]), name: un_(r.values[2]), flex: flexOf_(r.values), updated: r.values[4] };
}

function writeFlex_(sh, row, flex) {
  const parts = [];
  for (let i = 0; i < flex.length; i += CHUNK) parts.push(pre_(flex.slice(i, i + CHUNK)));
  const width = sh.getLastColumn();
  if (width >= FIRST_DATA_COL) sh.getRange(row, FIRST_DATA_COL, 1, width - FIRST_DATA_COL + 1).clearContent();
  sh.getRange(row, FIRST_DATA_COL, 1, parts.length).setValues([parts]);
}

function cleanFlex_(flex, id) {
  if (!flex) throw new Error('ไม่มี JSON');
  let obj;
  try { obj = JSON.parse(flex); } catch (e) { throw new Error('JSON ไม่ถูกต้อง'); }
  let s = JSON.stringify(obj);
  if (id && s.indexOf('{{sharedFlex}}') > -1) s = s.split('{{sharedFlex}}').join(shareUrl_(id));
  if (s.length > MAX_FLEX_CHARS) throw new Error('Flex ใหญ่เกินไป');
  return s;
}

// ลิงก์แชร์ต่อ: เปิดหน้าเดียวกัน (index.html) ในโหมดแชร์
function shareUrl_(id) { return 'https://liff.line.me/' + LIFF_ID + '?flex=' + encodeURIComponent(id); }

function newId_() { return Utilities.getUuid().replace(/-/g, '').slice(0, 12); }

function find_(uid, id) {
  const r = rows_().filter(function (x) { return un_(x.values[0]) === String(id); })[0];
  if (!r) throw new Error('ไม่พบรายการนี้ (อาจถูกลบไปแล้ว)');
  if (un_(r.values[1]) !== uid) throw new Error('รายการนี้ไม่ใช่ของคุณ');
  return r;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

// ---------- คำสั่ง ----------
function list_(uid) {
  const data = rows_()
    .filter(function (r) { return un_(r.values[1]) === uid; })
    .map(item_);
  return { result: 'success', data: data };
}

function put_(uid, name, flex) {
  name = String(name || '').trim();
  if (!name) throw new Error('กรุณาใส่ชื่อ');
  if (name.length > 100) throw new Error('ชื่อยาวเกิน 100 ตัวอักษร');
  return withLock_(function () {
    const dup = rows_().some(function (r) {
      return un_(r.values[1]) === uid && un_(r.values[2]).toLowerCase() === name.toLowerCase();
    });
    if (dup) return { result: 'duplicate', message: 'ชื่อซ้ำ' };
    const id = newId_();
    const s = cleanFlex_(flex, id);
    const sh = sheet_();
    const row = sh.getLastRow() + 1;
    const now = new Date();
    sh.getRange(row, 1, 1, 5).setValues([[pre_(id), pre_(uid), pre_(name), now, now]]);
    writeFlex_(sh, row, s);
    return { result: 'success', data: { id: id, name: name, flex: s, updated: now } };
  });
}

function update_(uid, id, flex, name) {
  return withLock_(function () {
    const r = find_(uid, id);
    const sh = sheet_();
    const s = cleanFlex_(flex, id);
    let newName = un_(r.values[2]);
    if (name != null && String(name).trim() && String(name).trim() !== newName) {
      newName = String(name).trim();
      const dup = rows_().some(function (x) {
        return x.row !== r.row && un_(x.values[1]) === uid && un_(x.values[2]).toLowerCase() === newName.toLowerCase();
      });
      if (dup) return { result: 'duplicate', message: 'ชื่อซ้ำ' };
      sh.getRange(r.row, 3).setValue(pre_(newName));
    }
    const now = new Date();
    sh.getRange(r.row, 5).setValue(now);
    writeFlex_(sh, r.row, s);
    return { result: 'success', data: { id: id, name: newName, flex: s, updated: now } };
  });
}

function del_(uid, id) {
  return withLock_(function () {
    const r = find_(uid, id);
    sheet_().deleteRow(r.row);
    return { result: 'success' };
  });
}

// ใช้โดยลิงก์แชร์ (?flex=ID) — ใครมีลิงก์ก็เปิดส่งต่อได้ แต่แก้/ลบไม่ได้
function one_(id) {
  const r = rows_().filter(function (x) { return un_(x.values[0]) === String(id); })[0];
  if (!r) return { result: 'error', message: 'ไม่พบ Flex นี้ (อาจถูกลบไปแล้ว)' };
  return { result: 'success', data: { id: un_(r.values[0]), flex: flexOf_(r.values) } };
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

// กด Run ฟังก์ชันนี้ 1 ครั้งใน Apps Script เพื่ออนุญาตสิทธิ์ (Sheet + เชื่อมต่อ LINE) ก่อน Deploy
function setup() {
  sheet_();
  UrlFetchApp.fetch('https://api.line.me/oauth2/v2.1/verify', { method: 'post', payload: { id_token: 'x', client_id: '1' }, muteHttpExceptions: true });
  Logger.log('พร้อมใช้งาน: สร้างแท็บ "' + SHEET_NAME + '" แล้ว — ต่อไปกด ทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่ > เว็บแอป');
}
