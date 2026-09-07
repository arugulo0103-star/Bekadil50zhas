const SHEET_NAME = "RSVP";
const ACCESS_TOKEN = "CHANGE_THIS_TOKEN_2026";

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(["Timestamp","Name","Attendance","Guests","Source"]);
  }
}

function doGet(e) {
  const p = e.parameter || {};
  if (p.action === "list") {
    if (p.token !== ACCESS_TOKEN) return json({ok:false,error:"Неверный токен"});
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sh) return json({ok:true,data:[]});
    const values = sh.getDataRange().getValues();
    const data = values.slice(1).filter(r=>r[1]).map(r=>({
      time: r[0] instanceof Date ? r[0].toISOString() : String(r[0]),
      name: String(r[1]),
      attendance: String(r[2]),
      guests: String(r[3] || 0),
      source: String(r[4] || "")
    }));
    return json({ok:true,data});
  }
  return HtmlService.createHtmlOutput("RSVP endpoint is active.");
}

function doPost(e) {
  const p = e.parameter || {};
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) ||
             SpreadsheetApp.getActiveSpreadsheet().insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) sh.appendRow(["Timestamp","Name","Attendance","Guests","Source"]);
  sh.appendRow([
    new Date(),
    String(p.name || ""),
    String(p.attendance || ""),
    String(p.guests || "0"),
    String(p.source || "")
  ]);
  return HtmlService.createHtmlOutput("<p style='font-family:Arial'>Спасибо! Ваш ответ принят.</p>");
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}