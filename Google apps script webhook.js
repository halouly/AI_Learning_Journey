// HLearn AI Journey Builder v3 — Google Apps Script Webhook
// Setup: New Sheet → Extensions → Apps Script → paste this → Deploy as Web app

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Responses");
    if (!sheet) {
      sheet = ss.insertSheet("Responses");
      sheet.appendRow([
        "Timestamp","Name","Email","Tracks","Waitlist",
        "Overall Score","Overall Band",
        "D1: Conceptual","D2: Capabilities","D3: Tool Proficiency","D4: Critical Eval","D5: Ethics & Risk",
        "D6: SC Use Cases","D7: SC Decisions",
        "D8: PM Practice","D9: PM Decisions",
        "D10: Engineering","D11: P.Eng Accountability",
        "SC Score","SC Band","PM Score","PM Band","P.Eng Score","P.Eng Band",
        "Strengths","Growth Areas",
        "Mastered Resources","Recommended Resources","Next Level Resources",
        "Raw Answers"
      ]);
      sheet.getRange(1, 1, 1, 30).setFontWeight("bold");
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      new Date(), data.name||"", data.email||"", data.tracks||"", data.waitlist||"",
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
