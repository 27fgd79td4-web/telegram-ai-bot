/**
 * ============================================================================
 * GOOGLE APPS SCRIPT — TELEGRAM AI BOT ACTIVITY LOGGER
 * ============================================================================
 * Ushbu skript Telegram botdan kelgan barcha harakatlarni avtomatik ravishda
 * ushbu Google Sheet jadvaliga chiroyli va tartibli qilib yozib boradi.
 *
 * O'RNATISH (1 DAQIQA):
 * 1. Yangi Google Sheet oching (https://sheets.new).
 * 2. Yuqori menyudan: Kengaytmalar (Extensions) -> Apps Script ni bosing.
 * 3. Ochilgan oynadagi barcha kodni o'chirib, ushbu fayldagi kodni joylashtiring.
 * 4. "Deploy" (Joylashtirish) -> "New deployment" (Yangi joylashtirish) ni bosing.
 * 5. Turi: "Web app" (Veb-ilova) ni tanlang.
 *    - Description: Telegram Bot Logger
 *    - Execute as: Me (Mening nomimdan)
 *    - Who has access: Anyone (Har kim)
 * 6. "Deploy" ni bosing va berilgan "Web App URL" (https://script.google.com/macros/s/.../exec)
 *    havolasini nusxalab oling.
 * 7. Ushbu havolani bot sozlamalariga (GOOGLE_SHEETS_WEBHOOK_URL) qo'ying!
 * ============================================================================
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    setupHeaders(sheet);

    var rawData = e.postData.contents;
    var data = JSON.parse(rawData);

    // Agar ommaviy (batch) yozuvlar kelsa
    if (data.batch && Array.isArray(data.batch)) {
      data.batch.forEach(function(item) {
        appendLogRow(sheet, item);
      });
      return ContentService
        .createTextOutput(JSON.stringify({ status: "success", count: data.batch.length }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Bitta yozuv kelsa
    appendLogRow(sheet, data);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "online", message: "Telegram Bot Google Sheets Webhook faol!" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupHeaders(sheet) {
  if (sheet.getLastRow() === 0) {
    var headers = [
      "№",
      "Sana va Vaqt",
      "Telegram ID",
      "Username",
      "Foydalanuvchi Ismi",
      "Harakat Turi",
      "Foydalanuvchi So'rovi",
      "AI Javobi (Xulosa)",
      "Tokenlar",
      "Holat"
    ];

    sheet.appendRow(headers);

    // Sarlavha dizayni (Dark modern styling)
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setBackground("#1e293b");
    headerRange.setFontColor("#ffffff");
    headerRange.setFontWeight("bold");
    headerRange.setFontSize(11);
    headerRange.setHorizontalAlignment("center");
    headerRange.setVerticalAlignment("middle");
    sheet.setRowHeight(1, 38);
    sheet.setFrozenRows(1);

    // Ustun o'lchamlari
    sheet.setColumnWidth(1, 50);   // №
    sheet.setColumnWidth(2, 150);  // Sana/Vaqt
    sheet.setColumnWidth(3, 120);  // Telegram ID
    sheet.setColumnWidth(4, 130);  // Username
    sheet.setColumnWidth(5, 160);  // Ismi
    sheet.setColumnWidth(6, 140);  // Harakat
    sheet.setColumnWidth(7, 300);  // So'rov
    sheet.setColumnWidth(8, 320);  // Javob
    sheet.setColumnWidth(9, 90);   // Token
    sheet.setColumnWidth(10, 130); // Holat
  }
}

function appendLogRow(sheet, item) {
  var rowNum = Math.max(1, sheet.getLastRow()); // Sarlavha 1-qatorda bo'lsa, keyingisi 1 bo'ladi
  var nextId = rowNum;

  var timestamp = item.timestamp || Utilities.formatDate(new Date(), "Asia/Tashkent", "yyyy-MM-dd HH:mm:ss");
  var userId    = item.user_id || "-";
  var username  = item.username ? "@" + item.username.replace(/^@/, '') : "-";
  var fullName  = item.full_name || "-";
  var action    = item.action || "🤖 Chat";
  var query     = (item.query || "").toString().substring(0, 1000);
  var response  = (item.response || "").toString().substring(0, 1000);
  var tokens    = item.tokens || 0;
  var status    = item.status === "ERROR" ? "🔴 Xatolik" : "🟢 Muvaffaqiyatli";

  sheet.appendRow([
    nextId,
    timestamp,
    userId,
    username,
    fullName,
    action,
    query,
    response,
    tokens,
    status
  ]);

  var lastRow = sheet.getLastRow();
  var rowRange = sheet.getRange(lastRow, 1, 1, 10);
  rowRange.setVerticalAlignment("middle");
  rowRange.setFontFamily("Segoe UI");
  rowRange.setFontSize(10);

  // Status rangini sozlash
  var statusCell = sheet.getRange(lastRow, 10);
  if (item.status === "ERROR") {
    statusCell.setBackground("#fee2e2");
    statusCell.setFontColor("#991b1b");
    statusCell.setFontWeight("bold");
  } else {
    statusCell.setBackground("#dcfce7");
    statusCell.setFontColor("#166534");
    statusCell.setFontWeight("bold");
  }

  // Tokenlar va ID ni markazga joylash
  sheet.getRange(lastRow, 1).setHorizontalAlignment("center");
  sheet.getRange(lastRow, 2).setHorizontalAlignment("center");
  sheet.getRange(lastRow, 3).setHorizontalAlignment("center");
  sheet.getRange(lastRow, 9).setHorizontalAlignment("center");
  sheet.getRange(lastRow, 10).setHorizontalAlignment("center");
}
