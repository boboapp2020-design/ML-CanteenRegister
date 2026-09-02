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
const SHEET_FB = 'ข้อเสนอแนะ';

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
var CACHE_TTL = 21600; // 6 ชม. — TTL ของ "เนื้อ state" (คีย์ตามเลขรุ่น จึงไม่ค้างถ้าเลขรุ่นเดิน)
var SGEN_TTL = 30;     // TTL สั้นของ "เลขรุ่น sgen" — กันแคชเลขรุ่นค้างยาว: ถ้าค้างจะหายเองใน 30 วิ
                       // (แหล่งจริงของ sgen คือ Script Properties ซึ่งเดินเสมอ; cache แค่ช่วยลดการอ่าน)
function _cache() { return CacheService.getScriptCache(); }
function _getGen() {
  var c = _cache(); var g = c.get('sgen');
  if (g === null || g === undefined) {
    g = PropertiesService.getScriptProperties().getProperty('sgen') || '0';
    try { c.put('sgen', g, SGEN_TTL); } catch (e) {}
  }
  return g;
}
function _bumpGen() {
  try {
    var p = PropertiesService.getScriptProperties();
    var g = String((Number(p.getProperty('sgen')) || 0) + 1);
    p.setProperty('sgen', g);
    _cache().put('sgen', g, SGEN_TTL);
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
  let rounds = 3, startDate = '', menuPub = 0, regGen = 0, regDeadline = '';
  for (let i = 1; i < cfgVals.length; i++) {
    const k = cfgVals[i][0];
    if (k === 'rounds') rounds = Number(cfgVals[i][1]) || 3;
    if (k === 'startDate') startDate = isoD(cfgVals[i][1]);
    if (k === 'menuPub') menuPub = Number(cfgVals[i][1]) || 0;
    if (k === 'regGen') regGen = Number(cfgVals[i][1]) || 0;
    if (k === 'regDeadline') { var _rd = cfgVals[i][1]; regDeadline = (_rd instanceof Date) ? _rd.toISOString() : String(_rd || ''); }
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

  // ข้อเสนอแนะ/ผลประเมิน -> [ { roundId, ts, unit, ratings, comment } ]  (ไม่ส่ง code/ชื่อ ออกไป เพื่อความเป็นส่วนตัว)
  var feedback = [];
  var fbSh = SS.getSheetByName(SHEET_FB);
  if (fbSh) {
    var fv = fbSh.getDataRange().getValues(); // roundId,ts,code,unit,taste,clean,portion,variety,comment
    for (var f = 1; f < fv.length; f++) {
      if ((fv[f][0] === '' || fv[f][0] === null) && (fv[f][1] === '' || fv[f][1] === null)) continue;
      var tt = fv[f][1];
      var tsStr = (tt instanceof Date) ? Utilities.formatDate(tt, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm') : String(tt || '');
      feedback.push({
        roundId: String(fv[f][0]), ts: tsStr, unit: String(fv[f][3] || ''),
        ratings: { taste: Number(fv[f][4]) || 0, clean: Number(fv[f][5]) || 0, portion: Number(fv[f][6]) || 0, variety: Number(fv[f][7]) || 0 },
        comment: String(fv[f][8] || '')
      });
    }
  }

  return { rounds: rounds, startDate: startDate, menuPub: menuPub, regGen: regGen, regDeadline: regDeadline, menu: menu, registrations: registrations, history: history, checkins: checkins, feedback: feedback };
}

// ---------- WRITE ACTIONS ----------
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    // ส่งแจ้งเตือน push (ไม่เปลี่ยนข้อมูล จึงไม่ต้องเลื่อนรุ่นแคช) — ตอบกลับทันที
    if (action === 'pushNotify') return jsonOut(osPush_(data.heading, data.content, data.url));
    if (action === 'getFeedback') return jsonOut(getFeedback(data));   // อ่านข้อเสนอแนะตรงจากชีต (ไม่ผ่านแคช)
    if (action === 'bustCache') return jsonOut(bustCache());           // ล้างแคช state ที่ค้าง
    var out = null;
    if (action === 'saveConfig') out = saveConfig(data);
    else if (action === 'saveMenuDay') out = saveMenuDay(data);
    else if (action === 'saveRegistrationDay') out = saveRegistrationDay(data);
    else if (action === 'saveRegistration') out = saveRegistration(data);
    else if (action === 'newRound') out = newRound(data);
    else if (action === 'checkin') out = checkin(data);
    else if (action === 'clearRegistrations') out = clearRegistrations(data);
    else if (action === 'restoreRound') out = restoreRound(data);
    else if (action === 'purgeHistory') out = purgeHistory(data);
    else if (action === 'saveFeedback') out = saveFeedback(data);
    if (out === null) return jsonOut({ ok: false, error: 'unknown action: ' + action });
    _bumpGen(); // ข้อมูลเปลี่ยน → เลื่อนรุ่นแคช การอ่านครั้งถัดไปจะได้ข้อมูลใหม่
    return jsonOut(out);
  } catch (err) {
    return jsonOut({ ok: false, error: String(err) });
  }
}

// กู้คืนรอบจากประวัติกลับเข้ารอบปัจจุบัน (เขียนกลับทีเดียวรวดเดียว เร็ว+ครบ)
// data: { roundId: "yyyy-MM-dd|N" }  — คัดลอกเมนู+ลงทะเบียนของ roundId นั้นทับรอบปัจจุบัน
function restoreRound(data) {
  var roundId = String(data.roundId || '');
  if (!roundId) return { ok: false, error: 'no roundId' };
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var parts = roundId.split('|');
    var startDate = parts[0];
    var rounds = Number(parts[1]) || 0;
    // ----- เมนู: อ่านจากประวัติเมนู เขียนทับเมนูปัจจุบัน -----
    var menuSh = sheet(SHEET_MENU);
    clearDataRows(menuSh);
    var mOut = [];
    var hm = SS.getSheetByName(SHEET_HMENU);
    if (hm) {
      var hv = hm.getDataRange().getValues(); // roundId,dayIndex,dateISO,mealKey,itemsJson
      for (var i = 1; i < hv.length; i++) {
        if (String(hv[i][0]) === roundId) mOut.push([hv[i][1], hv[i][2], hv[i][3], hv[i][4]]);
      }
    }
    if (mOut.length) menuSh.getRange(2, 1, mOut.length, 4).setValues(mOut);
    // ----- ลงทะเบียน: อ่านจากประวัติลงทะเบียน เขียนทับรอบปัจจุบัน -----
    var regSh = sheet(SHEET_REG);
    clearDataRows(regSh);
    var rOut = [];
    var hr = SS.getSheetByName(SHEET_HREG);
    if (hr) {
      var rv = hr.getDataRange().getValues(); // roundId,code,dayIndex,breakfast,lunch,dinner,updatedAt
      for (var j = 1; j < rv.length; j++) {
        if (String(rv[j][0]) === roundId) rOut.push([rv[j][1], rv[j][2], rv[j][3], rv[j][4], rv[j][5], rv[j][6]]);
      }
    }
    if (rOut.length) regSh.getRange(2, 1, rOut.length, 6).setValues(rOut);
    // ----- ตั้ง Config เป็นรอบนี้ + เผยแพร่ -----
    saveConfig({ rounds: rounds, startDate: startDate, menuPub: 1 });
    return { ok: true, menuRows: mOut.length, regRows: rOut.length, people: (function () { var s = {}; rOut.forEach(function (r) { s[r[0]] = 1; }); return Object.keys(s).length; })() };
  } finally {
    lock.releaseLock();
  }
}

// อ่านข้อเสนอแนะตรงจากชีต (ไม่ผ่านแคช getState) — คืนโดยไม่มี code/ชื่อ
function getFeedback(data) {
  var roundId = (data && data.roundId) ? String(data.roundId) : '';
  var arr = [];
  var fbSh = SS.getSheetByName(SHEET_FB);
  if (fbSh) {
    var fv = fbSh.getDataRange().getValues(); // roundId,ts,code,unit,taste,clean,portion,variety,comment
    for (var f = 1; f < fv.length; f++) {
      if ((fv[f][0] === '' || fv[f][0] === null) && (fv[f][1] === '' || fv[f][1] === null)) continue;
      if (roundId && String(fv[f][0]) !== roundId) continue;
      var tt = fv[f][1];
      var ts = (tt instanceof Date) ? Utilities.formatDate(tt, 'Asia/Bangkok', 'yyyy-MM-dd HH:mm') : String(tt || '');
      arr.push({ roundId: String(fv[f][0]), ts: ts, unit: String(fv[f][3] || ''),
        ratings: { taste: Number(fv[f][4]) || 0, clean: Number(fv[f][5]) || 0, portion: Number(fv[f][6]) || 0, variety: Number(fv[f][7]) || 0 },
        comment: String(fv[f][8] || '') });
    }
  }
  return { ok: true, feedback: arr };
}

// ล้างแคช state ที่ค้าง (แก้อาการ getState เสิร์ฟข้อมูลเก่าเพราะเลขรุ่น sgen หยุดเลื่อน)
function bustCache() {
  var out = { ok: true };
  try {
    var p = PropertiesService.getScriptProperties();
    var g = Number(p.getProperty('sgen')) || 0;
    var c = _cache();
    var keys = ['sgen'];
    for (var gg = 0; gg <= g + 2; gg++) { keys.push('st_' + gg + '_n'); for (var i = 0; i < 12; i++) keys.push('st_' + gg + '_' + i); }
    keys = keys.slice(0, 480);
    try { c.removeAll(keys); } catch (e) {}
    var ng = String(g + 1);
    try { p.setProperty('sgen', ng); } catch (e) {}
    try { c.put('sgen', ng, CACHE_TTL); } catch (e) {}
    out.gen = ng;
  } catch (e) { out.error = String(e); }
  return out;
}
// รันเองจาก Editor เพื่อล้างแคชทันที
function bustCacheNow() { return bustCache(); }

// บันทึกข้อเสนอแนะ/ผลประเมิน (เก็บ code ไว้กันซ้ำ 1 คน/รอบ แต่ getState จะไม่ส่ง code ออก)
// data: { roundId, code, unit, ratings:{taste,clean,portion,variety}, comment, ts }
function saveFeedback(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sh = getOrCreateSheet(SHEET_FB, ['roundId', 'ts', 'code', 'unit', 'taste', 'clean', 'portion', 'variety', 'comment']);
    var r = data.ratings || {};
    var row = [String(data.roundId || ''), String(data.ts || ''), String(data.code || ''), String(data.unit || ''),
      Number(r.taste) || '', Number(r.clean) || '', Number(r.portion) || '', Number(r.variety) || '', String(data.comment || '')];
    var vals = sh.getDataRange().getValues();
    var foundRow = -1;
    for (var i = 1; i < vals.length; i++) {
      if (String(vals[i][0]) === String(data.roundId) && String(vals[i][2]) === String(data.code)) { foundRow = i + 1; break; }
    }
    if (foundRow > 0) sh.getRange(foundRow, 1, 1, 9).setValues([row]);   // 1 คน/รอบ — แก้ทับของเดิม
    else sh.appendRow(row);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ล้างประวัติรอบเก่า (ประวัติเมนู + ประวัติลงทะเบียน) เพื่อให้ state เล็ก/เร็ว — รอบปัจจุบันไม่ถูกแตะ
// data: { keepDays }  0/ไม่ส่ง = ลบประวัติทั้งหมด · เช่น 30 = เก็บเฉพาะรอบที่เริ่มภายใน 30 วันล่าสุด
function purgeHistory(data) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var keepDays = Number(data && data.keepDays) || 0;
    var cutoff = keepDays > 0 ? (Date.now() - keepDays * 86400000) : null;
    function roundStartMs(rid) {
      var p = String(rid).split('|')[0];
      var d = new Date(p + 'T00:00:00+07:00'); if (isNaN(d.getTime())) d = new Date(p);
      return isNaN(d.getTime()) ? 0 : d.getTime();
    }
    function purgeSheet(name) {
      var sh = SS.getSheetByName(name); if (!sh) return { removed: 0, kept: 0 };
      var vals = sh.getDataRange().getValues(); if (vals.length < 2) return { removed: 0, kept: 0 };
      var head = vals[0], keep = [];
      for (var i = 1; i < vals.length; i++) {
        var rid = vals[i][0]; if (rid === '' || rid === null) continue;
        if (cutoff !== null && roundStartMs(rid) >= cutoff) keep.push(vals[i]);
      }
      sh.clearContents();
      sh.getRange(1, 1, 1, head.length).setValues([head]);
      if (keep.length) sh.getRange(2, 1, keep.length, keep[0].length).setValues(keep);
      return { removed: (vals.length - 1) - keep.length, kept: keep.length };
    }
    var m = purgeSheet(SHEET_HMENU), r = purgeSheet(SHEET_HREG);
    _cleanReminderProps_();
    return { ok: true, menuRowsRemoved: m.removed, regRowsRemoved: r.removed, menuRowsKept: m.kept, regRowsKept: r.kept };
  } finally {
    lock.releaseLock();
  }
}
// (ทางเลือก) ล้างประวัติอัตโนมัติทุกเดือน — รันฟังก์ชันนี้ครั้งเดียวจาก Editor เพื่อติดตั้ง
function setupMonthlyPurgeTrigger() {
  var t = ScriptApp.getProjectTriggers();
  for (var i = 0; i < t.length; i++) if (t[i].getHandlerFunction() === 'monthlyPurge') ScriptApp.deleteTrigger(t[i]);
  ScriptApp.newTrigger('monthlyPurge').timeBased().onMonthDay(1).atHour(3).create();
  return 'ตั้งล้างประวัติอัตโนมัติทุกวันที่ 1 (~03:00) แล้ว — เก็บรอบใน 30 วันล่าสุด ลบที่เก่ากว่า';
}
function monthlyPurge() { var r = purgeHistory({ keepDays: 30 }); _bumpGen(); return r; }

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
    let foundPub = false, foundDl = false;
    for (let i = 1; i < vals.length; i++) {
      if (vals[i][0] === 'rounds' && data.rounds !== undefined) cfgSh.getRange(i + 1, 2).setValue(data.rounds);
      if (vals[i][0] === 'startDate' && data.startDate !== undefined) cfgSh.getRange(i + 1, 2).setValue(data.startDate);
      if (vals[i][0] === 'menuPub') { foundPub = true; if (data.menuPub !== undefined) cfgSh.getRange(i + 1, 2).setValue(data.menuPub); }
      // regDeadline = เวลาปิดรับลงทะเบียนจริง (ISO) ที่แอปคำนวณให้ — ใช้ตอนส่งแจ้งเตือนอัตโนมัติ
      if (vals[i][0] === 'regDeadline') { foundDl = true; if (data.regDeadline !== undefined) cfgSh.getRange(i + 1, 2).setValue(data.regDeadline); }
    }
    if (!foundPub && data.menuPub !== undefined) cfgSh.appendRow(['menuPub', data.menuPub, 'สถานะยืนยันเมนูของรอบ (1=ยืนยันแล้ว)']);
    if (!foundDl && data.regDeadline !== undefined) cfgSh.appendRow(['regDeadline', data.regDeadline, 'เวลาปิดรับลงทะเบียนจริงของรอบ (ISO) สำหรับแจ้งเตือนอัตโนมัติ']);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// เริ่มรอบใหม่: ย้ายข้อมูลรอบเดิมไปเก็บใน "ประวัติเมนู/ประวัติลงทะเบียน" (เพื่อดูย้อนหลัง/สะสม)
// แล้วล้างตารางรอบปัจจุบัน และตั้งค่ารอบใหม่
function newRound(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    // อ่านค่ารอบเดิมจาก Config ก่อน เพื่อใช้เป็น roundId ของประวัติ
    var cfgVals = sheet(SHEET_CONFIG).getDataRange().getValues();
    var oldRounds = 3, oldStart = '';
    for (var i = 1; i < cfgVals.length; i++) {
      if (cfgVals[i][0] === 'rounds') oldRounds = Number(cfgVals[i][1]) || 3;
      if (cfgVals[i][0] === 'startDate') oldStart = isoD(cfgVals[i][1]);
    }
    var roundId = oldStart + '|' + oldRounds;

    // [แก้บั๊ก ส.ค. 2026] เดิม archive ด้วย appendRow ทีละแถว — พนักงานหลายร้อยคน
    // ทำให้ newRound ถือ LockService ค้างหลายสิบวินาที คำสั่ง saveMenuDay วันแรก
    // ที่แอดมินพิมพ์ทันทีหลังตั้งรอบเลยรอ lock ไม่ทัน (timeout) → เมนูวันแรกหาย
    // แก้เป็น: รวบทุกแถวเป็น array เดียวแล้วเขียนครั้งเดียวด้วย setValues() → เร็วขึ้นมาก
    var menuSh = sheet(SHEET_MENU);
    var menuVals = menuSh.getDataRange().getValues();
    if (oldStart && menuVals.length > 1) {
      var hMenu = getOrCreateSheet(SHEET_HMENU, ['roundId', 'dayIndex', 'dateISO', 'mealKey', 'itemsJson']);
      var mOut = [];
      for (var m = 1; m < menuVals.length; m++) {
        if (menuVals[m][0] === '' || menuVals[m][0] === null) continue;
        mOut.push([roundId, menuVals[m][0], menuVals[m][1], menuVals[m][2], menuVals[m][3]]);
      }
      if (mOut.length) hMenu.getRange(hMenu.getLastRow() + 1, 1, mOut.length, 5).setValues(mOut);
    }
    var regSh = sheet(SHEET_REG);
    var regVals = regSh.getDataRange().getValues();
    if (oldStart && regVals.length > 1) {
      var hReg = getOrCreateSheet(SHEET_HREG, ['roundId', 'code', 'dayIndex', 'breakfast', 'lunch', 'dinner', 'updatedAt']);
      var rOut = [];
      for (var r = 1; r < regVals.length; r++) {
        if (regVals[r][0] === '' || regVals[r][0] === null) continue;
        rOut.push([roundId, regVals[r][0], regVals[r][1], regVals[r][2], regVals[r][3], regVals[r][4], regVals[r][5]]);
      }
      if (rOut.length) hReg.getRange(hReg.getLastRow() + 1, 1, rOut.length, 7).setValues(rOut);
    }

    clearDataRows(menuSh);
    clearDataRows(regSh);
    saveConfig(data);
    bumpRegGen();   // รอบใหม่ = รุ่นข้อมูลใหม่ ทุกเครื่องลงทะเบียนใหม่ได้
    _cleanReminderProps_();   // ล้าง key กันเตือนซ้ำของรอบเก่า กัน Script Properties ล้น
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

// ---------- WEB PUSH (OneSignal) ----------
// ส่งแจ้งเตือนเด้งหน้าจอ (เหมือน LINE) ให้ทุกคนที่กด "เปิดการแจ้งเตือน" ไว้
// สำคัญ: REST API Key เป็นความลับ — เก็บใน Script Properties เท่านั้น ห้ามใส่ในไฟล์นี้ (ไฟล์นี้อยู่บน GitHub)
//   ตั้งค่า: Apps Script > Project Settings > Script Properties > เพิ่ม key ชื่อ ONESIGNAL_REST_KEY
var ONESIGNAL_APP_ID = 'e50edb91-4c00-44e6-ae2a-e0d3505d3720';
var APP_HOME_URL = 'https://ml-canteenregister.boboapp2020.workers.dev/';
function osKey_() { return PropertiesService.getScriptProperties().getProperty('ONESIGNAL_REST_KEY') || ''; }
// ส่งไปหาทุกคนที่สมัครไว้ — OneSignal แต่ละบัญชีตั้งชื่อ segment "ทุกคน" ไม่เหมือนกัน
// (รุ่นใหม่ = "Total Subscriptions", รุ่นเก่า = "Subscribed Users") จึงไล่ลองทีละชื่อจนเจออันที่มีคนรับจริง
var OS_SEGMENTS = ['Total Subscriptions', 'Subscribed Users', 'Active Subscriptions'];
function osPush_(heading, content, url) {
  var key = osKey_();
  if (!key) return { ok: false, error: 'ยังไม่ได้ตั้งค่า ONESIGNAL_REST_KEY ใน Script Properties' };
  if (!heading) heading = 'โรงครัวมิตรลาว';
  if (!content) content = '';
  var last = null;
  for (var i = 0; i < OS_SEGMENTS.length; i++) {
    var payload = {
      app_id: ONESIGNAL_APP_ID,
      headings: { en: heading },
      contents: { en: content },
      included_segments: [OS_SEGMENTS[i]],
      url: url || APP_HOME_URL
    };
    try {
      var res = UrlFetchApp.fetch('https://onesignal.com/api/v1/notifications', {
        method: 'post',
        contentType: 'application/json',
        headers: { Authorization: 'Basic ' + key },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true
      });
      var code = res.getResponseCode();
      var body = res.getContentText();
      var j = {};
      try { j = JSON.parse(body); } catch (e2) {}
      // OneSignal รุ่นใหม่ตอบ {"id":"...","external_id":null} โดยไม่มี recipients — ถ้าได้ id จริง = รับส่งแล้ว
      var hasId = j.id && String(j.id).length > 0;
      var hasErr = j.errors && ((j.errors.length) || Object.keys(j.errors).length);
      last = { ok: code === 200 && hasId && !hasErr, code: code, recipients: (j.recipients === undefined ? null : Number(j.recipients)), segment: OS_SEGMENTS[i], body: body };
      if (last.ok) return last;   // สำเร็จแล้ว หยุด (กันส่งซ้ำหลาย segment)
    } catch (err) {
      last = { ok: false, error: String(err), segment: OS_SEGMENTS[i] };
    }
  }
  return last;   // ไม่มี segment ไหนส่งถึงเลย — คืนผลล่าสุดไว้ดู error
}

// ทดสอบส่งเอง: รันฟังก์ชันนี้จาก Apps Script Editor เพื่อยิง push ทดสอบไปหาทุกคนที่สมัครไว้
// ผลจะถูก log ไว้ดูใน Execution log (บอกว่าใช้ segment ไหน ส่งถึงกี่คน)
function testPush() {
  var r = osPush_('🔔 ทดสอบแจ้งเตือน', 'ระบบแจ้งเตือนโรงครัวมิตรลาวพร้อมใช้งานแล้ว', APP_HOME_URL);
  Logger.log(JSON.stringify(r));
  return r;
}

// ---------- แจ้งเตือน "รีบลงทะเบียน" อัตโนมัติ จนถึงปิดรับ ----------
// ส่งซ้ำทุกวัน 2 เวลา (เช้า 07:00 และบ่าย 16:45) + เตือนด่วนก่อนปิด 1 ชม. แล้วหยุดเมื่อปิดรับ
// เวลาปิดรับ: ใช้ค่า regDeadline (ISO) จาก Config ถ้ามี (รองรับกรณีขยายเวลาพิเศษ)
//   ถ้าไม่มี ใช้กติกามาตรฐาน = 17:00 ของวันก่อนวันเริ่มรอบ
// ทำงานด้วย time trigger ทุก 5 นาที · เช้า/บ่ายวันละครั้ง · ชั่วโมงสุดท้ายเตือนทุก 10 นาที
var REMIND_MORNING = '07:00';     // "ก่อน 8 โมง"
var REMIND_AFTERNOON = '16:45';
function _cfgMap_() {
  var vals = sheet(SHEET_CONFIG).getDataRange().getValues();
  var m = {};
  for (var i = 1; i < vals.length; i++) m[vals[i][0]] = vals[i][1];
  return m;
}
function _effectiveDeadline_(cfg) {
  if (cfg.regDeadline) { var d = new Date(cfg.regDeadline); if (!isNaN(d.getTime())) return d; }
  var startDate = isoD(cfg.startDate); if (!startDate) return null;
  return new Date(new Date(startDate + 'T00:00:00+07:00').getTime() - 7 * 3600 * 1000);
}
function remindDeadline() {
  var cfg = _cfgMap_();
  var startDate = isoD(cfg.startDate);
  var rounds = Number(cfg.rounds) || 0;
  var menuPub = Number(cfg.menuPub) || 0;
  if (!startDate || !menuPub) return;                    // ไม่มีรอบ/ยังไม่เผยแพร่เมนู = ไม่เตือน
  remindMeals_(startDate, rounds);                       // A) เตือนก่อนเวลาทานอาหาร (วันที่มีเมนู)
  var deadline = _effectiveDeadline_(cfg); if (!deadline) return;  // B) เตือนรีบลงทะเบียน (จนถึงปิดรับ)
  var hoursUntil = (deadline.getTime() - Date.now()) / 3600000;
  if (hoursUntil <= 0) return;                           // ปิดรับแล้ว หยุดเตือน
  var roundId = startDate + '|' + rounds;
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  var hm = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm');
  var dlText = Utilities.formatDate(deadline, 'Asia/Bangkok', 'dd/MM HH:mm');
  function fire(slotKey, heading, content) {
    var k = 'rem_' + roundId + '_' + slotKey;
    if (props.getProperty(k) === '1') return;
    var r = osPush_(heading, content, APP_HOME_URL);
    if (r && r.ok) props.setProperty(k, '1');
  }
  // 1) ชั่วโมงสุดท้ายก่อนปิด — เตือนถี่ "ทุก 10 นาที" จนถึงเวลาปิด (นับถอยหลังเหลือกี่นาที)
  //    dedup ตาม "ช่วง 10 นาที" ของนาฬิกา เพื่อให้ส่งช่วงละครั้ง (trigger เดินทุก 5 นาที)
  if (hoursUntil <= 1) {
    var mm = Number(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'mm'));
    var bucket = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd_HH') + '_' + Math.floor(mm / 10);
    var minsLeft = Math.max(1, Math.round(hoursUntil * 60));
    fire('f_' + bucket, '⏰ ใกล้ปิดรับลงทะเบียน!', 'เหลืออีกประมาณ ' + minsLeft + ' นาที (ปิด ' + dlText + ' น.) รีบลงทะเบียนด่วน!');
    return;
  }
  // 2) รอบเช้า 07:00 (หน้าต่าง 07:00–07:59) วันละครั้ง
  if (hm >= REMIND_MORNING && hm < '08:00') { fire('m_' + today, '🍽️ อย่าลืมลงทะเบียนอาหาร', 'ยังเปิดรับลงทะเบียนอยู่ (ปิด ' + dlText + ' น.) รีบลงทะเบียนก่อนหมดเวลา'); return; }
  // 3) รอบบ่าย 16:45 (หน้าต่าง 16:45–16:59) วันละครั้ง
  if (hm >= REMIND_AFTERNOON && hm < '17:00') { fire('a_' + today, '🍽️ อย่าลืมลงทะเบียนอาหาร', 'ยังเปิดรับลงทะเบียนอยู่ (ปิด ' + dlText + ' น.) รีบลงทะเบียนก่อนหมดเวลา'); return; }
}
// ---------- แจ้งเตือน "ก่อนเวลาทานอาหาร" (ทุกวันที่มีเมนู ในช่วงรอบ) ----------
// เตือนแยกตามมื้อ · เฉพาะมื้อที่มีเมนูของวันนั้น · เตือน "ทุก 15 นาที" ตั้งแต่เวลาเริ่ม (s)
// จนถึงเวลาปิดของมื้อนั้น (e = หมดเวลาสแกน/รับอาหาร) — เช้า 08:00, เที่ยง 13:10, เย็น 18:30
var MEAL_REMIND = [
  { key: 'breakfast', s: '07:30', e: '08:00', label: 'อาหารเช้า' },
  { key: 'lunch',     s: '11:30', e: '13:10', label: 'อาหารกลางวัน' },
  { key: 'dinner',    s: '16:30', e: '18:30', label: 'อาหารเย็น' }
];
function _menuItemsFor_(menuVals, dayIndex, mealKey) {
  for (var i = 1; i < menuVals.length; i++) {
    if (String(menuVals[i][0]) === String(dayIndex) && menuVals[i][2] === mealKey) {
      try { return JSON.parse(menuVals[i][3] || '[]'); } catch (e) { return []; }
    }
  }
  return [];
}
function remindMeals_(startDate, rounds) {
  var today = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd');
  var hm = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH:mm');
  // วันนี้เป็นวันที่เท่าไรของรอบ (0 = วันแรก) — ต้องอยู่ในช่วง 0..rounds-1
  var startMs = new Date(startDate + 'T00:00:00+07:00').getTime();
  var todayMs = new Date(today + 'T00:00:00+07:00').getTime();
  var dayIndex = Math.round((todayMs - startMs) / 86400000);
  if (dayIndex < 0 || dayIndex >= rounds) return;        // วันนี้ไม่อยู่ในรอบ = ไม่เตือนมื้อ
  var menuVals = sheet(SHEET_MENU).getDataRange().getValues();
  var props = PropertiesService.getScriptProperties();
  for (var i = 0; i < MEAL_REMIND.length; i++) {
    var m = MEAL_REMIND[i];
    if (!(hm >= m.s && hm <= m.e)) continue;             // อยู่ในช่วง [เริ่ม, หมดเวลากิน] ของมื้อนี้
    // dedup ตามช่วง 15 นาที (เช่น 11:30, 11:45, 12:00 ...) — trigger เดินทุก 5 นาที ส่งช่วงละครั้ง
    var mm = Number(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'mm'));
    var bkt = Utilities.formatDate(new Date(), 'Asia/Bangkok', 'HH') + '_' + Math.floor(mm / 15);
    var k = 'meal_' + today + '_' + m.key + '_' + bkt;
    if (props.getProperty(k) === '1') continue;          // ช่วง 15 นาทีนี้ของมื้อนี้เตือนไปแล้ว
    var items = _menuItemsFor_(menuVals, dayIndex, m.key);
    if (!items.length) continue;                         // ไม่มีเมนูมื้อนี้ = ไม่เตือน
    var names = items.map(function (x) { return x.name; }).join(' / ');
    var r = osPush_('🍽️ ถึงเวลา' + m.label + 'แล้ว', 'เมนูวันนี้: ' + names + ' · เชิญรับประทานที่โรงอาหารได้เลย', APP_HOME_URL);
    if (r && r.ok) props.setProperty(k, '1');
  }
}
// รันฟังก์ชันนี้ "ครั้งเดียว" เพื่อติดตั้ง/อัปเดตตัวจับเวลา (หลังจากนั้นทำงานเองทุก 5 นาที)
function setupReminderTrigger() {
  var t = ScriptApp.getProjectTriggers();
  for (var i = 0; i < t.length; i++) if (t[i].getHandlerFunction() === 'remindDeadline') ScriptApp.deleteTrigger(t[i]);
  ScriptApp.newTrigger('remindDeadline').timeBased().everyMinutes(5).create();
  return 'ติดตั้งตัวจับเวลาแล้ว — เช็คทุก 5 นาที (เตือน 07:00, 16:45 และชั่วโมงสุดท้ายทุก 10 นาที)';
}
// ล้างเฉพาะ key "กันเตือนซ้ำ" (rem_/meal_/dl_sent_) — ไม่แตะ sgen, regGen, ONESIGNAL_REST_KEY
// เรียกตอนขึ้นรอบใหม่ เพื่อไม่ให้ Script Properties สะสมจนเต็ม (500KB)
function _cleanReminderProps_() {
  try {
    var props = PropertiesService.getScriptProperties();
    var all = props.getProperties();
    Object.keys(all).forEach(function (k) {
      if (k.indexOf('rem_') === 0 || k.indexOf('meal_') === 0 || k.indexOf('dl_sent_') === 0) {
        try { props.deleteProperty(k); } catch (e) {}
      }
    });
  } catch (e) {}
}
// รันเองได้จาก Editor เพื่อล้าง key เตือนเก่าทันที (เผื่อสะสมไว้เยอะก่อนมีฟังก์ชันนี้)
function cleanReminderPropsNow() { _cleanReminderProps_(); return 'ล้าง key กันเตือนซ้ำเก่าแล้ว'; }
