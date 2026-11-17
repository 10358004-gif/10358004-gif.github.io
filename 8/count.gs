// Deployment:
// 1. Create a Google Apps Script project and paste this file into count.gs.
// 2. Run setupCounter() once and authorize access to the new spreadsheet.
// 3. Deploy as a web app, execute as you, and allow access to anyone.
// 4. Paste the deployed /exec URL into GAS_WEB_APP_URL in index.html.

const VISITS_SHEET_NAME = "Visits";
const SPREADSHEET_PROPERTY = "VISITOR_COUNTER_SPREADSHEET_ID";
const VISITS_HEADER = ["時間（UTC）", "頁面路徑", "來源頁面", "瀏覽器資訊"];

function setupCounter() {
  const properties = PropertiesService.getScriptProperties();
  let spreadsheetId = properties.getProperty(SPREADSHEET_PROPERTY);
  let spreadsheet;

  if (spreadsheetId) {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } else {
    spreadsheet = SpreadsheetApp.create("網頁訪客計數器紀錄");
    properties.setProperty(SPREADSHEET_PROPERTY, spreadsheet.getId());
  }

  let sheet = spreadsheet.getSheetByName(VISITS_SHEET_NAME);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(VISITS_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(VISITS_HEADER);
    sheet.setFrozenRows(1);
  }

  Logger.log("瀏覽紀錄試算表：" + spreadsheet.getUrl());
}

function doGet(e) {
  const callback = e && e.parameter ? e.parameter.callback : "";
  if (!isValidCallback_(callback)) {
    return ContentService.createTextOutput("Invalid callback")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  if (!e.parameter || e.parameter.action !== "visit") {
    return jsonp_(callback, { ok: false, error: "Unsupported action" });
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const sheet = getVisitsSheet_();
    sheet.appendRow([
      new Date().toISOString(),
      safeCellValue_(e.parameter.page, 500),
      safeCellValue_(e.parameter.referrer, 1000),
      safeCellValue_(e.parameter.userAgent, 500)
    ]);
    const count = Math.max(0, sheet.getLastRow() - 1);
    return jsonp_(callback, { ok: true, count: count });
  } catch (error) {
    console.error("Unable to record visitor: " + error);
    return jsonp_(callback, { ok: false, error: "Unable to record visit" });
  } finally {
    if (lock.hasLock()) {
      lock.releaseLock();
    }
  }
}

function getVisitsSheet_() {
  const spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty(SPREADSHEET_PROPERTY);
  if (!spreadsheetId) {
    throw new Error("Run setupCounter() and authorize the spreadsheet first.");
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(VISITS_SHEET_NAME);
  if (!sheet) {
    throw new Error("Visits sheet is missing; run setupCounter() again.");
  }
  return sheet;
}

function jsonp_(callback, value) {
  const output = ContentService.createTextOutput(
    callback + "(" + JSON.stringify(value) + ");"
  );
  return output.setMimeType(ContentService.MimeType.JAVASCRIPT);
}

function isValidCallback_(callback) {
  return typeof callback === "string" &&
    /^[A-Za-z_$][0-9A-Za-z_$]*$/.test(callback);
}

function safeCellValue_(value, maxLength) {
  const text = String(value || "").slice(0, maxLength);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}
