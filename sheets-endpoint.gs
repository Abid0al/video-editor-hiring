// =====================================================================
//  VIDEO EDITOR APPLICATION → GOOGLE SHEETS ENDPOINT
//
//  SETUP (one-time, ~3 minutes):
//  1. Open Extensions → Apps Script in your spreadsheet.
//  2. Delete any existing code, paste this file, and save (Ctrl+S).
//  3. Select "setup" from the function dropdown and click Run.
//     This creates the sheet tab and writes all column headers.
//  4. Click Deploy → New deployment.
//     · Type: Web app
//     · Execute as: Me
//     · Who has access: Anyone
//  5. Click Deploy, then Authorize access (grant all requested permissions).
//  6. Copy the Web app URL shown after deployment.
//  7. Paste that URL into FORM_ENDPOINT in index.html.
// =====================================================================

const SPREADSHEET_ID = "1O0r4ym_kEZUysdj-YOMDInD0wqAI7hjpKejjl36mwfE";
const SHEET_NAME     = "Application";

// Maps form fields → exact column header in your sheet.
// partial:true does a startsWith match (used for the long rate column header).
// Columns not found in the sheet are appended automatically.
const FIELD_MAP = [
  { col: "Timestamp",                                                                                             value:   () => new Date() },
  { col: "Full Name",                                                                                             field:   "full_name" },
  { col: "Email Address",                                                                                         field:   "email" },
  { col: "How long have you been working as a video editor?",                                                     field:   "experience" },
  { col: "Please share your portfolio (Google Drive, YouTube, Vimeo, or website)",                               field:   "portfolio" },
  { col: "Which software are you proficient in? (Select all that apply)",                                        multi:   "software" },
  { col: "Can you confidently perform ALL of the following? (Required)",                                         multi:   "req" },
  { col: "Are you comfortable editing videos to match reference styles and following detailed feedback?",         field:   "feedback" },
  { col: "Are you willing to complete a 1-minute trial edit as part of the hiring process?",                     field:   "trial" },
  { col: "What is your expected fixed rate",                                                                      field:   "rate",    partial: true },
  { col: "Why do you want to join our team?",                                                                     field:   "why_join" },
  { col: "Anything else you'd like us to know?",                                                                  field:   "anything_else" },
  { col: "Country Code",                                                                                          field:   "country_code" },
  { col: "WhatsApp",                                                                                              field:   "whatsapp" },
];

// Run this once from the Apps Script editor to create the sheet and write headers.
function setup() {
  const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  let   sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Write every header from FIELD_MAP into row 1
  const headers = FIELD_MAP.map(f => f.col);
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Format header row
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight("bold")
             .setBackground("#1a1a2e")
             .setFontColor("#ffffff")
             .setWrap(false);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  Logger.log("Setup complete: " + headers.length + " columns written to \"" + SHEET_NAME + "\"");
}

function doPost(e) {
  try {
    Logger.log("doPost triggered");
    Logger.log("Parameters: " + JSON.stringify(e.parameter));

    const ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
    let   sheet = ss.getSheetByName(SHEET_NAME);
    Logger.log("Sheet found: " + (sheet ? "yes" : "no — will create"));

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.setFrozenRows(1);
    }

    // Read existing headers; first occurrence wins (handles duplicate columns)
    const lastCol    = Math.max(sheet.getLastColumn(), 1);
    const headerVals = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const colIndex   = {};
    headerVals.forEach((h, i) => {
      const key = String(h || "").trim();
      if (key && !colIndex[key]) colIndex[key] = i + 1;   // first occurrence wins
    });

    // Helper: find column by exact name, or startsWith for partial matches
    function findCol(name, partial) {
      if (colIndex[name]) return colIndex[name];
      if (partial) {
        for (const [header, ci] of Object.entries(colIndex)) {
          if (header.startsWith(name)) return ci;
        }
      }
      return null;
    }

    // Add any missing columns to the right
    FIELD_MAP.forEach(({ col, partial }) => {
      if (!findCol(col, partial)) {
        const newCol = sheet.getLastColumn() + 1;
        sheet.getRange(1, newCol).setValue(col)
             .setFontWeight("bold")
             .setBackground("#1a1a2e")
             .setFontColor("#ffffff");
        colIndex[col] = newCol;
      }
    });

    // Build the new row
    const p         = e.parameter;
    const pv        = e.parameters;
    const totalCols = sheet.getLastColumn();
    const rowData   = new Array(totalCols).fill("");

    FIELD_MAP.forEach(({ col, field, multi, value, partial }) => {
      const ci = (findCol(col, partial) || 0) - 1;
      if (ci < 0) return;
      if (value)      rowData[ci] = value();
      else if (multi) rowData[ci] = (pv[multi] || []).join(", ");
      else if (field) rowData[ci] = p[field] || "";
    });

    sheet.appendRow(rowData);
    Logger.log("Row appended successfully");

    const lastRow = sheet.getLastRow();
    if (lastRow % 20 === 0) sheet.autoResizeColumns(1, totalCols);

    return ContentService
      .createTextOutput(JSON.stringify({ result: "success", row: lastRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    Logger.log("ERROR: " + err.message + " | Stack: " + err.stack);
    return ContentService
      .createTextOutput(JSON.stringify({ result: "error", message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Visit the deployed URL in a browser to confirm the endpoint is live
function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: "ok", message: "Video editor application endpoint is live." }))
    .setMimeType(ContentService.MimeType.JSON);
}
