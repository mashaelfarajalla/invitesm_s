/**
 * سكربت استقبال ردود دعوة الزفاف
 * يُلصق في: Google Sheet ← Extensions ← Apps Script
 *
 * - تبويب "الحضور": كل تأكيد حضور بصف
 * - تبويب "المباركات": كل رسالة بصف، وعمود "إظهار" للتحكم بعرضها على الموقع
 */

var SHEET_RSVP = "الحضور";
var SHEET_MSGS = "المباركات";
var MAX_MESSAGES = 100;
var CACHE_KEY = "messages";
var CACHE_SECONDS = 600;

// استقبال النماذج من الموقع
function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var p = e.parameter;
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (p.type === "rsvp") {
      if (!p.name) return json_({ ok: false, error: "missing fields" });
      getSheet_(ss, SHEET_RSVP, ["التاريخ", "الاسم", "الجوال", "عدد المرافقين", "الحضور"])
        .appendRow([
          new Date(),
          clean_(p.name, 80),
          p.phone ? "'" + String(p.phone).trim().slice(0, 30) : "", // ' للحفاظ على الصفر في بداية الرقم
          Math.min(Math.max(Number(p.guests) || 0, 0), 20),
          p.attend === "no" ? "لن يحضر" : "سيحضر"
        ]);

    } else if (p.type === "blessing") {
      if (!p.name || !p.message) return json_({ ok: false, error: "missing fields" });
      var sheet = getSheet_(ss, SHEET_MSGS, ["التاريخ", "الاسم", "الرسالة", "إظهار"]);
      sheet.appendRow([new Date(), clean_(p.name, 60), clean_(p.message, 500), true]);
      sheet.getRange(sheet.getLastRow(), 4).insertCheckboxes().check();
      CacheService.getScriptCache().remove(CACHE_KEY);

    } else {
      return json_({ ok: false, error: "unknown type" });
    }

    return json_({ ok: true });
  } finally {
    lock.releaseLock();
  }
}

// إرجاع رسائل المباركة الظاهرة (الأحدث أولاً)، مع كاش لتسريع الاستجابة
function doGet() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(CACHE_KEY);
  if (hit) {
    return ContentService.createTextOutput(hit).setMimeType(ContentService.MimeType.JSON);
  }
  var body = JSON.stringify({ ok: true, messages: readMessages_() });
  try { cache.put(CACHE_KEY, body, CACHE_SECONDS); } catch (err) {} // الكاش محدود بـ 100KB
  return ContentService.createTextOutput(body).setMimeType(ContentService.MimeType.JSON);
}

// أي تعديل يدوي على الشيت (مثل إلغاء "إظهار") يمسح الكاش ليظهر التغيير فوراً
function onEdit() {
  CacheService.getScriptCache().remove(CACHE_KEY);
}

function readMessages_() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_MSGS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return rows
    .filter(function (r) { return r[3] === true && r[1] && r[2]; })
    .reverse()
    .slice(0, MAX_MESSAGES)
    .map(function (r) {
      return { name: String(r[1]).replace(/^'/, ""), message: String(r[2]).replace(/^'/, "") };
    });
}

function getSheet_(ss, name, headers) {
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold");
    sheet.setFrozenRows(1);
    sheet.setRightToLeft(true);
  }
  return sheet;
}

// تقصير النص ومنع تنفيذه كمعادلة داخل الشيت
function clean_(value, max) {
  var s = String(value || "").trim().slice(0, max);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
