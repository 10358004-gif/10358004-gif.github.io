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
  const parameters = e && e.parameter ? e.parameter : {};
  if (parameters.action !== "visit") {
    return ContentService.createTextOutput(
      "Counter API is ready. Open the website to record a visit."
    ).setMimeType(ContentService.MimeType.TEXT);
  }

  const callback = parameters.callback || "";
  if (!isValidCallback_(callback)) {
    return ContentService.createTextOutput("Missing or invalid callback parameter.")
      .setMimeType(ContentService.MimeType.TEXT);
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const sheet = getVisitsSheet_();
    sheet.appendRow([
      new Date().toISOString(),
      safeCellValue_(parameters.page, 500),
      safeCellValue_(parameters.referrer, 1000),
      safeCellValue_(parameters.userAgent, 500)
    ]);
    const count = Math.max(0, sheet.getLastRow() - 1);
    return jsonp_(callback, { ok: true, count: count });
  } catch (error) {
    console.error("Unable to record visitor: " + error);
    const errorMessage = error && error.message ? error.message : "";
    const publicMessage = errorMessage === "COUNTER_NOT_INITIALIZED" ||
      errorMessage === "VISITS_SHEET_MISSING"
      ? "計數器尚未初始化。請在 Apps Script 執行 setupCounter() 並完成授權。"
      : "瀏覽紀錄寫入失敗。請到 Apps Script 的「執行項目」查看錯誤。";
    return jsonp_(callback, { ok: false, error: publicMessage });
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
    throw new Error("COUNTER_NOT_INITIALIZED");
  }

  const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  const sheet = spreadsheet.getSheetByName(VISITS_SHEET_NAME);
  if (!sheet) {
    throw new Error("VISITS_SHEET_MISSING");
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
