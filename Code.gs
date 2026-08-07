/**
 * Code.gs — Backend API สำหรับแอปลงทะเบียนอาหารโรงครัว (มิตรลาว)
 * ------------------------------------------------------------
 * วิธีติดตั้ง (ทำครั้งเดียว):
 * 1) เปิด Google Sheet "ระบบลงทะเบียนอาหารโรงครัว - ฐานข้อมูล"
 * 2) เมนู Extensions > Apps Script
 * 3) ลบโค้ดเดิมในไฟล์ Code.gs (ถ้ามี) แล้ววางไฟล์นี้ทั้งหมดแทน
 * 4) กด Deploy > New deployment > เลือกประเภท "Web app"
 *      - Execute as: Me (บัญชีของคุณ)
 *      - Who has access: Anyone
 * 5) กด Deploy แล้วอนุญาตสิทธิ์ (Authorize access) ตามที่ Google ถาม
 * 6) คัดลอก "Web app URL" ที่ได้ (ลงท้ายด้วย /exec) ส่งให้ผู้พัฒนาแอป
 *    เพื่อนำไปใส่ในไฟล์แอป (ตัวแปร API_URL)
 *
 * ทุกครั้งที่แก้โค้ดนี้ ต้องกด Deploy > Manage deployments > แก้ไข (ไอคอนดินสอ)
 * > Version: New version > Deploy ใหม่ (URL เดิมยังใช้ได้)
 *
 * [อัปเดต ส.ค. 2026] เพิ่มระบบแคชฝั่งเซิร์ฟเวอร์ (CacheService)
 * - อ่าน state ซ้ำ ๆ จะเสิร์ฟจากแคชทันที ไม่ต้องอ่านชีททุกแท็บใหม่ทุกครั้ง
 * - ทุกการเขียน (ลงทะเบียน/บันทึกเมนู/ตั้งรอบใหม่/เช็คอิน) จะเลื่อนเลขรุ่นแคช
 *   ทำให้การอ่านครั้งถัดไปได้ข้อมูลใหม่เสมอ ไม่มีทางได้ข้อมูลค้าง
 */

const SS = SpreadsheetApp.getActiveSpreadsheet();
const SHEET_CONFIG = 'Config';
const SHEET_MENU = 'เมนูรายวัน';
const SHEET_REG = 'การลงทะเบียน';
const SHEET_HMENU = 'ประวัติเมนู';
const SHEET_HREG = 'ประวัติลงทะเบียน';
const SHEET_CI = 'เช็คอิน';

// ช่วงเวลาสแกน QR เช็คอิน (เวลาไทย) — ต้องตรงกับในแอป
const CI_WINDOWS = {
  breakfast: ['06:30', '08:00'],
  lunch: ['11:55', '13:10'],
  dinner: ['16:00', '18:30']
};

function getOrCreateSheet(name, headers) {
  var sh = SS.getSheetByName(name);
  if (!sh) {
    sh = SS.insertSheet(name);
    sh.appendRow(headers);
  }
  return sh;
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function sheet(name) {
  const sh = SS.getSheetByName(name);
  if (!sh) throw new Error('ไม่พบชีต: ' + name);
  return sh;
}

// Google Sheet อาจแปลงข้อความวันที่เป็น Date object อัตโนมัติ — แปลงกลับเป็น yyyy-MM-dd เสมอ
function isoD(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'Asia/Bangkok', 'yyyy-MM-dd');
  return String(v || '');
}

// ---------- SERVER CACHE (ทำให้ตอบไวขึ้นมาก) ----------
// หลักการ: เก็บ "เลขรุ่นข้อมูล" (sgen) ไว้ — ทุกครั้งที่มีการเขียน เลขนี้ +1
// ผล state ของแต่ละรุ่นถูกแคชแยกคีย์ตามเลขรุ่น การอ่านหลังการเขียนจึงพลาดแคชเก่าเสมอ (ไม่มีข้อมูลค้าง)
var CACHE_TTL = 21600; // 6 ชม. (ค่าสูงสุดของ Google)
function _cache() { return CacheService.getScriptCache(); }
function _getGen() {
  var c = _cache(); var g = c.get('sgen');
  if (g === null || g === undefined) {
    g = PropertiesService.getScriptProperties().getProperty('sgen') || '0';
    try { c.put('sgen', g, CACHE_TTL); } catch (e) {}
  }
  return g;
}
function _bumpGen() {
  try {
    var p = PropertiesService.getScriptProperties();
    var g = String((Number(p.getProperty('sgen')) || 0) + 1);
    p.setProperty('sgen', g);
    _cache().put('sgen', g, CACHE_TTL);
  } catch (e) {}
}
// state อาจใหญ่กว่า 100KB (ลิมิตต่อคีย์ของ CacheService) จึงหั่นเป็นท่อนละ 90,000 ตัวอักษร
function _cacheGetState(gen) {
  try {
    var c = _cache(); var n = Number(c.get('st_' + gen + '_n')) || 0;
    if (!n) return null;
    var keys = []; for (var i = 0; i < n; i++) keys.push('st_' + gen + '_' + i);
    var got = c.getAll(keys); var s = '';
    for (var j = 0; j < n; j++) { var part = got['st_' + gen + '_' + j]; if (part === undefined || part === null) return null; s += part; }
    return s;
  } catch (e) { return null; }
}
function _cachePutState(gen, s) {
  try {
    var c = _cache(); var CH = 90000; var n = Math.ceil(s.length / CH) || 1;
    var obj = {}; for (var i = 0; i < n; i++) obj['st_' + gen + '_' + i] = s.substr(i * CH, CH);
    c.putAll(obj, CACHE_TTL);
    c.put('st_' + gen + '_n', String(n), CACHE_TTL);
  } catch (e) { /* แคชไม่สำเร็จก็แค่เสิร์ฟช้าลง ข้อมูลยังถูกต้อง */ }
}

// ---------- READ STATE ----------
function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'state';
    if (action === 'state') {
      var gen = _getGen();
      var hit = _cacheGetState(gen);
      if (hit) return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
      var s = JSON.stringify(getState());
      _cachePutState(gen, s);
      return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.JSON);
    }
    return jsonOut({ error: 'unknown action: ' + action });
  } catch (err) {
    return jsonOut({ error: String(err) });
  }
}

function getState() {
  // Config
  const cfgSh = sheet(SHEET_CONFIG);
  const cfgVals = cfgSh.getDataRange().getValues(); // [ [key,value,desc], ... ]
  let rounds = 3, startDate = '', menuPub = 0, regGen = 0;
  for (let i = 1; i < cfgVals.length; i++) {
    const k = cfgVals[i][0];
    if (k === 'rounds') rounds = Number(cfgVals[i][1]) || 3;
    if (k === 'startDate') startDate = isoD(cfgVals[i][1]);
    if (k === 'menuPub') menuPub = Number(cfgVals[i][1]) || 0;
    if (k === 'regGen') regGen = Number(cfgVals[i][1]) || 0;
  }

  // เมนูรายวัน -> { dayIndex: { mealKey: [items...] } }
  const menuSh = sheet(SHEET_MENU);
  const menuVals = menuSh.getDataRange().getValues(); // header + rows: dayIndex,dateISO,mealKey,itemsJson
  const menu = {};
  for (let i = 1; i < menuVals.length; i++) {
    const row = menuVals[i];
    const dayIndex = row[0];
    if (dayIndex === '' || dayIndex === null || dayIndex === undefined) continue;
    const mealKey = row[2];
    let items = [];
    try { items = JSON.parse(row[3] || '[]'); } catch (e2) { items = []; }
    if (!menu[dayIndex]) menu[dayIndex] = {};
    menu[dayIndex][mealKey] = items;
  }

  // การลงทะเบียน -> { code: { dayIndex: { breakfast, lunch, dinner } } }
  const regSh = sheet(SHEET_REG);
  const regVals = regSh.getDataRange().getValues(); // header + rows: code,dayIndex,breakfast,lunch,dinner,updatedAt
  const registrations = {};
  for (let i = 1; i < regVals.length; i++) {
    const row = regVals[i];
    const code = row[0];
    if (code === '' || code === null || code === undefined) continue;
    const dayIndex = row[1];
    const rec = {};
    if (row[2] !== '' && row[2] !== null) rec.breakfast = !!row[2];
    if (row[3] !== '' && row[3] !== null) rec.lunch = !!row[3];
    if (row[4] !== '' && row[4] !== null) rec.dinner = !!row[4];
    if (!registrations[code]) registrations[code] = {};
    registrations[code][dayIndex] = rec;
  }

  // ประวัติรอบเก่า -> [ { startDate, rounds, menu, registrations } ]
  var history = [];
  var hmap = {};
  var hm = SS.getSheetByName(SHEET_HMENU);
  if (hm) {
    var hmVals = hm.getDataRange().getValues(); // roundId,dayIndex,dateISO,mealKey,itemsJson
    for (var i = 1; i < hmVals.length; i++) {
      var rid = hmVals[i][0];
      if (rid === '' || rid === null) continue;
      if (!hmap[rid]) { var p = String(rid).split('|'); hmap[rid] = { startDate: p[0], rounds: Number(p[1]) || 0, menu: {}, registrations: {} }; }
      var di = hmVals[i][1], mk = hmVals[i][3];
      var items = []; try { items = JSON.parse(hmVals[i][4] || '[]'); } catch (e3) { items = []; }
      if (!hmap[rid].menu[di]) hmap[rid].menu[di] = {};
      hmap[rid].menu[di][mk] = items;
    }
  }
  var hr = SS.getSheetByName(SHEET_HREG);
  if (hr) {
    var hrVals = hr.getDataRange().getValues(); // roundId,code,dayIndex,breakfast,lunch,dinner,updatedAt
    for (var j = 1; j < hrVals.length; j++) {
      var rid2 = hrVals[j][0];
      if (rid2 === '' || rid2 === null) continue;
      if (!hmap[rid2]) { var p2 = String(rid2).split('|'); hmap[rid2] = { startDate: p2[0], rounds: Number(p2[1]) || 0, menu: {}, registrations: {} }; }
      var code = hrVals[j][1], di2 = hrVals[j][2];
      var rec2 = {};
      if (hrVals[j][3] !== '' && hrVals[j][3] !== null) rec2.breakfast = !!hrVals[j][3];
      if (hrVals[j][4] !== '' && hrVals[j][4] !== null) rec2.lunch = !!hrVals[j][4];
      if (hrVals[j][5] !== '' && hrVals[j][5] !== null) rec2.dinner = !!hrVals[j][5];
      if (!hmap[rid2].registrations[code]) hmap[rid2].registrations[code] = {};
      hmap[rid2].registrations[code][di2] = rec2;
    }
  }
  Object.keys(hmap).sort().forEach(function (k) { history.push(hmap[k]); });

  // เช็คอิน (สแกน QR หน้าโรงอาหาร) -> { dateISO: { mealKey: count } }
  var checkins = {};
  var ciSh = SS.getSheetByName(SHEET_CI);
  if (ciSh) {
    var ciVals = ciSh.getDataRange().getValues(); // dateISO, mealKey, count
    for (var c = 1; c < ciVals.length; c++) {
      var cd = isoD(ciVals[c][0]);
      if (!cd) continue;
      if (!checkins[cd]) checkins[cd] = {};
      checkins[cd][ciVals[c][1]] = Number(ciVals[c][2]) || 0;
    }
  }

  return { rounds: rounds, startDate: startDate, menuPub: menuPub, regGen: regGen, menu: menu, registrations: registrations, history: history, checkins: checkins };
}

// ---------- WRITE ACTIONS ----------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    var out = null;
    if (action === 'saveConfig') out = saveConfig(data);
    else if (action === 'saveMenuDay') out = saveMenuDay(data);
    else if (action === 'saveRegistrationDay') out = saveRegistrationDay(data);
    else if (action === 'saveRegistration') out = saveRegistration(data);
    else if (action === 'newRound') out = newRound(data);
    else if (action === 'checkin') out = checkin(data);
    else if (action === 'clearRegistrations') out = clearRegistrations(data);
    if (out === null) return jsonOut({ ok: false, error: 'unknown action: ' + action });
    _bumpGen(); // ข้อมูลเปลี่ยน → เลื่อนรุ่นแคช การอ่านครั้งถัดไปจะได้ข้อมูลใหม่
    return jsonOut(out);
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

// ล้างเฉพาะข้อมูลการลงทะเบียนของรอบปัจจุบัน (เมนู/Config/ประวัติ ไม่ถูกแตะ)
function clearRegistrations(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet(SHEET_REG);
    const last = sh.getLastRow();
    if (last > 1) sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent();
    bumpRegGen();   // เปลี่ยน "รุ่นข้อมูล" เพื่อให้ทุกเครื่องรู้ว่าล้างแล้ว → ลงทะเบียนใหม่ได้
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// เพิ่มค่า regGen ใน Config ทีละ 1 (ใช้บอกทุกเครื่องว่าข้อมูลลงทะเบียนถูกรีเซ็ต)
function bumpRegGen() {
  const cfgSh = sheet(SHEET_CONFIG);
  const vals = cfgSh.getDataRange().getValues();
  for (let i = 1; i < vals.length; i++) {
    if (vals[i][0] === 'regGen') {
      cfgSh.getRange(i + 1, 2).setValue((Number(vals[i][1]) || 0) + 1);
      return;
    }
  }
  cfgSh.appendRow(['regGen', 1, 'รุ่นข้อมูลลงทะเบียน (เพิ่มขึ้นทุกครั้งที่ล้าง/เริ่มรอบใหม่)']);
}

function saveConfig(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cfgSh = sheet(SHEET_CONFIG);
    const vals = cfgSh.getDataRange().getValues();
    let foundPub = false;
    for (let i = 1; i < vals.length; i++) {
      if (vals[i][0] === 'rounds' && data.rounds !== undefined) cfgSh.getRange(i + 1, 2).setValue(data.rounds);
      if (vals[i][0] === 'startDate' && data.startDate !== undefined) cfgSh.getRange(i + 1, 2).setValue(data.startDate);
      if (vals[i][0] === 'menuPub') { foundPub = true; if (data.menuPub !== undefined) cfgSh.getRange(i + 1, 2).setValue(data.menuPub); }
    }
    if (!foundPub && data.menuPub !== undefined) cfgSh.appendRow(['menuPub', data.menuPub, 'สถานะยืนยันเมนูของรอบ (1=ยืนยันแล้ว)']);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// เริ่มรอบใหม่: ย้ายข้อมูลรอบเดิมไปเก็บใน "ประวัติเมนู/ประวัติลงทะเบียน" (เพื่อดูย้อนหลัง/สะสม)
// แล้วล้างตารางรอบปัจจุบัน และตั้งค่ารอบใหม่
function newRound(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // อ่านค่ารอบเดิมจาก Config ก่อน เพื่อใช้เป็น roundId ของประวัติ
    var cfgVals = sheet(SHEET_CONFIG).getDataRange().getValues();
    var oldRounds = 3, oldStart = '';
    for (var i = 1; i < cfgVals.length; i++) {
      if (cfgVals[i][0] === 'rounds') oldRounds = Number(cfgVals[i][1]) || 3;
      if (cfgVals[i][0] === 'startDate') oldStart = isoD(cfgVals[i][1]);
    }
    var roundId = oldStart + '|' + oldRounds;

    var menuSh = sheet(SHEET_MENU);
    var menuVals = menuSh.getDataRange().getValues();
    if (oldStart && menuVals.length > 1) {
      var hMenu = getOrCreateSheet(SHEET_HMENU, ['roundId', 'dayIndex', 'dateISO', 'mealKey', 'itemsJson']);
      for (var m = 1; m < menuVals.length; m++) {
        if (menuVals[m][0] === '' || menuVals[m][0] === null) continue;
        hMenu.appendRow([roundId, menuVals[m][0], menuVals[m][1], menuVals[m][2], menuVals[m][3]]);
      }
    }
    var regSh = sheet(SHEET_REG);
    var regVals = regSh.getDataRange().getValues();
    if (oldStart && regVals.length > 1) {
      var hReg = getOrCreateSheet(SHEET_HREG, ['roundId', 'code', 'dayIndex', 'breakfast', 'lunch', 'dinner', 'updatedAt']);
      for (var r = 1; r < regVals.length; r++) {
        if (regVals[r][0] === '' || regVals[r][0] === null) continue;
        hReg.appendRow([roundId, regVals[r][0], regVals[r][1], regVals[r][2], regVals[r][3], regVals[r][4], regVals[r][5]]);
      }
    }

    clearDataRows(menuSh);
    clearDataRows(regSh);
    saveConfig(data);
    bumpRegGen();   // รอบใหม่ = รุ่นข้อมูลใหม่ ทุกเครื่องลงทะเบียนใหม่ได้
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// เช็คอิน: นับจำนวนคนมาใช้สิทธิ์จริงต่อคาบ (ตรวจช่วงเวลาที่ฝั่งเซิร์ฟเวอร์ด้วย กันสแกนนอกเวลา)
// data: { date: 'yyyy-MM-dd', meal: 'breakfast'|'lunch'|'dinner' }
function checkin(data) {
  var meal = data.meal;
  if (!CI_WINDOWS[meal]) return { ok: false, error: 'invalid meal' };
  // ตรวจเวลาปัจจุบัน (เวลาไทย) ต้องอยู่ในช่วงของคาบนั้น
  var hm = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm');
  var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  var w = CI_WINDOWS[meal];
  if (hm < w[0] || hm > w[1]) return { ok: false, error: 'closed', window: w };
  if (String(data.date) !== today) return { ok: false, error: 'wrong date' };

  // ถ้าคาบนี้ของวันนี้ไม่มีเมนูอาหารในระบบ = ไม่รับเช็คอิน
  var menuVals = sheet(SHEET_MENU).getDataRange().getValues();
  var hasMenu = false;
  for (var mi = 1; mi < menuVals.length; mi++) {
    if (isoD(menuVals[mi][1]) === today && menuVals[mi][2] === meal) {
      var items = [];
      try { items = JSON.parse(menuVals[mi][3] || '[]'); } catch (e2) { items = []; }
      if (items.length) { hasMenu = true; break; }
    }
  }
  if (!hasMenu) return { ok: false, error: 'nomenu' };

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getOrCreateSheet(SHEET_CI, ['dateISO', 'mealKey', 'count', 'updatedAt']);
    var vals = sh.getDataRange().getValues();
    var now = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm:ss');
    for (var i = 1; i < vals.length; i++) {
      if (isoD(vals[i][0]) === today && vals[i][1] === meal) {
        var cnt = (Number(vals[i][2]) || 0) + 1;
        sh.getRange(i + 1, 3, 1, 2).setValues([[cnt, now]]);
        return { ok: true, count: cnt };
      }
    }
    sh.appendRow([today, meal, 1, now]);
    return { ok: true, count: 1 };
  } finally {
    lock.releaseLock();
  }
}

function clearDataRows(sh) {
  const lastRow = sh.getLastRow();
  const lastCol = sh.getLastColumn();
  if (lastRow > 1) sh.getRange(2, 1, lastRow - 1, lastCol).clearContent();
}

// data: { dayIndex, dateISO, mealKey, items: [{name,emoji,img}] }
function saveMenuDay(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet(SHEET_MENU);
    const vals = sh.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === String(data.dayIndex) && vals[i][2] === data.mealKey) {
        foundRow = i + 1;
        break;
      }
    }
    const itemsJson = JSON.stringify(data.items || []);
    if (foundRow > 0) {
      sh.getRange(foundRow, 1, 1, 4).setValues([[data.dayIndex, data.dateISO, data.mealKey, itemsJson]]);
    } else {
      sh.appendRow([data.dayIndex, data.dateISO, data.mealKey, itemsJson]);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// data: { code, dayIndex, breakfast, lunch, dinner }  (ค่าที่ไม่ส่งมาจะไม่ถูกแก้)
function saveRegistrationDay(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet(SHEET_REG);
    const vals = sh.getDataRange().getValues();
    let foundRow = -1;
    for (let i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === String(data.code) && String(vals[i][1]) === String(data.dayIndex)) {
        foundRow = i + 1;
        break;
      }
    }
    const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Vientiane', 'yyyy-MM-dd HH:mm:ss');
    const b = data.breakfast !== undefined ? data.breakfast : (foundRow > 0 ? vals[foundRow - 1][2] : '');
    const l = data.lunch !== undefined ? data.lunch : (foundRow > 0 ? vals[foundRow - 1][3] : '');
    const d = data.dinner !== undefined ? data.dinner : (foundRow > 0 ? vals[foundRow - 1][4] : '');
    if (foundRow > 0) {
      sh.getRange(foundRow, 1, 1, 6).setValues([[data.code, data.dayIndex, b, l, d, now]]);
    } else {
      sh.appendRow([data.code, data.dayIndex, b, l, d, now]);
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// บันทึกทั้งรอบครั้งเดียว (atomic): ลบแถวเดิมของรหัสนี้ทิ้ง แล้วเขียนใหม่ทุกวันในล็อกเดียว
// -> ข้อมูลของ 1 คน จะครบทุกวันเสมอ หรือไม่มีเลย (ไม่มีทางลงไม่ครบ)
// data: { code, days: { "0": {breakfast,lunch,dinner}, "1": {...}, ... } }
function saveRegistration(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = sheet(SHEET_REG);
    const vals = sh.getDataRange().getValues();
    // ลบแถวเดิมของรหัสนี้ (ไล่จากล่างขึ้นบน กัน index เลื่อน)
    for (var i = vals.length - 1; i >= 1; i--) {
      if (String(vals[i][0]) === String(data.code)) sh.deleteRow(i + 1);
    }
    var now = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Vientiane', 'yyyy-MM-dd HH:mm:ss');
    var days = data.days || {};
    var pick = function (v) { return v === undefined || v === null ? '' : v; };
    Object.keys(days).forEach(function (di) {
      var r = days[di] || {};
      sh.appendRow([data.code, Number(di), pick(r.breakfast), pick(r.lunch), pick(r.dinner), now]);
    });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}
