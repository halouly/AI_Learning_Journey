// ═══════════════════════════════════════════════════════════
// HLearn AI Journey Builder — Google Apps Script Webhook
// ═══════════════════════════════════════════════════════════
//
// SETUP INSTRUCTIONS:
//
// 1. Go to https://script.google.com → New Project
// 2. Delete the default code, paste this entire file
// 3. Click "Deploy" → "New deployment"
// 4. Type = "Web app"
// 5. Execute as = "Me"
// 6. Who has access = "Anyone" (required for cross-origin POST)
// 7. Click "Deploy" → copy the Web App URL
// 8. Paste that URL into your HTML app (WEBHOOK_URL variable)
//
// The script auto-creates the sheet with headers on first run.
// ═══════════════════════════════════════════════════════════

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    // Auto-create sheet if it doesn't exist
    var sheet = ss.getSheetByName("Responses");
    if (!sheet) {
      sheet = ss.insertSheet("Responses");
      sheet.appendRow([
        "Timestamp",
        "Name",
        "Email",
        "Track",
        "Overall Score",
        "Overall Band",
        "D1: Conceptual",
        "D2: Capabilities",
        "D3: Tool Proficiency",
        "D4: Critical Eval",
        "D5: Ethics & Risk",
        "D6: SC Use Cases",
        "D7: SC Decisions",
        "SC Score",
        "SC Band",
        "Strengths",
        "Growth Areas",
        "Raw Answers"
      ]);
      // Bold header row
      sheet.getRange(1, 1, 1, 18).setFontWeight("bold");
      // Freeze header
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      new Date(),
      data.name || "",
      data.email || "",
      data.track || "",
      data.overallScore || "",
      data.overallBand || "",
      data.d1 || "",
      data.d2 || "",
      data.d3 || "",
      data.d4 || "",
      data.d5 || "",
      data.d6 || "",
      data.d7 || "",
      data.scScore || "",
      data.scBand || "",
      data.strengths || "",
      data.growthAreas || "",
      data.rawAnswers || ""
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: "ok" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Required: handle CORS preflight
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "Webhook is live" }))
    .setMimeType(ContentService.MimeType.JSON);
}
