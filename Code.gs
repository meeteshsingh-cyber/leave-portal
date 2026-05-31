// ════════════════════════════════════════════════════════════
//  LeaveDesk — Google Apps Script Backend
//  Paste this entire file into Google Apps Script and deploy
//  as a Web App (see README.md for exact steps)
// ════════════════════════════════════════════════════════════

const SHEET_NAME = "Leaves";

// Column order in the sheet
const COLS = {
  ID:          1,
  NAME:        2,
  EMP_ID:      3,
  FROM:        4,
  TO:          5,
  REJOIN:      6,
  TYPE:        7,
  PURPOSE:     8,
  STATUS:      9,
  SUBMITTED:   10,
  ACTED_AT:    11,
};

// ── Utility ──────────────────────────────────────────────────
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    // Write header row
    sheet.appendRow([
      "ID", "Name", "Emp ID", "From", "To", "Rejoin",
      "Leave Type", "Purpose", "Status", "Submitted At", "Acted At"
    ]);
    sheet.setFrozenRows(1);
    // Style header
    const header = sheet.getRange(1, 1, 1, Object.keys(COLS).length);
    header.setBackground("#1A6B4A").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.setColumnWidths(1, Object.keys(COLS).length, 160);
  }
  return sheet;
}

function corsHeaders() {
  return ContentService.createTextOutput()
    .setMimeType(ContentService.MimeType.JSON);
}

function respond(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── GET — fetch all requests ──────────────────────────────────
function doGet(e) {
  try {
    const sheet = getSheet();
    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) return respond({ success: true, data: [] });

    const headers = rows[0];
    const data = rows.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });

    return respond({ success: true, data });
  } catch(err) {
    return respond({ success: false, error: err.message });
  }
}

// ── POST — handle actions ─────────────────────────────────────
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);
    const action = body.action;

    if (action === "submit") return submitLeave(body);
    if (action === "approve") return updateStatus(body.id, "approved");
    if (action === "reject")  return updateStatus(body.id, "rejected");
    if (action === "undo")    return updateStatus(body.id, "pending", true);

    return respond({ success: false, error: "Unknown action" });
  } catch(err) {
    return respond({ success: false, error: err.message });
  }
}

// ── Submit new leave request ──────────────────────────────────
function submitLeave(data) {
  const sheet = getSheet();
  const id = "REQ-" + Date.now().toString().slice(-6);
  sheet.appendRow([
    id,
    data.name,
    data.empId,
    data.from,
    data.to,
    data.rejoin || "",
    data.type,
    data.purpose,
    "pending",
    new Date().toISOString(),
    ""
  ]);
  return respond({ success: true, id });
}

// ── Update status (approve / reject / undo) ───────────────────
function updateStatus(id, status, isUndo) {
  const sheet = getSheet();
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][COLS.ID - 1] === id) {
      sheet.getRange(i + 1, COLS.STATUS).setValue(status);
      sheet.getRange(i + 1, COLS.ACTED_AT).setValue(isUndo ? "" : new Date().toISOString());

      // Colour-code the row
      const range = sheet.getRange(i + 1, 1, 1, Object.keys(COLS).length);
      if (status === "approved") range.setBackground("#E6F4EE");
      else if (status === "rejected") range.setBackground("#FDECEA");
      else range.setBackground("#FFFFFF");

      return respond({ success: true });
    }
  }
  return respond({ success: false, error: "Request not found" });
}
