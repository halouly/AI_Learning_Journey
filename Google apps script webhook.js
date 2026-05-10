// HLearn AI Journey Builder v4 — Google Apps Script
// Setup: New Sheet → Extensions → Apps Script → paste → Deploy as Web app

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Responses");
    if (!sheet) {
      sheet = ss.insertSheet("Responses");
      sheet.appendRow([
        "Timestamp","Name","Email","Tracks","Level","Waitlist",
        "Overall Score","Overall Band",
        "D1","D2","D3","D4","D5","D6","D7","D8","D9","D10","D11",
        "SC Score","SC Band","PM Score","PM Band","P.Eng Score","P.Eng Band",
        "Strengths","Growth Areas",
        "Mastered","Recommended","Next Level",
        "Raw Answers"
      ]);
      sheet.getRange(1, 1, 1, 31).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      new Date(), data.name||"", data.email||"", data.tracks||"", data.level||"", data.waitlist||"",
      data.overallScore||"", data.overallBand||"",
      data.d1||"", data.d2||"", data.d3||"", data.d4||"", data.d5||"",
      data.d6||"", data.d7||"", data.d8||"", data.d9||"", data.d10||"", data.d11||"",
      data.scScore||"", data.scBand||"", data.pmScore||"", data.pmBand||"", data.peScore||"", data.peBand||"",
      data.strengths||"", data.growthAreas||"",
      data.mastered||"", data.recommended||"", data.nextLevel||"",
      data.rawAnswers||""
    ]);
    return ContentService.createTextOutput(JSON.stringify({status:"ok"})).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({status:"error",message:err.toString()})).setMimeType(ContentService.MimeType.JSON);
  }
}
function doGet(e) {
  return ContentService.createTextOutput(JSON.stringify({status:"ok",message:"Webhook is live"})).setMimeType(ContentService.MimeType.JSON);
}
