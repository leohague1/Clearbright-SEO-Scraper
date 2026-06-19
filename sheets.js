import { google } from "googleapis";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const SPREADSHEET_ID = "1t8X7bS98yRWQyZPchqwJhiwidML2Lrsp_8cJ5V0XUnw";
const SHEET_ID       = 0; // gid=0 = Sheet1
const SHEET_NAME     = "Sheet1";
const HEADERS    = [
  "Business Name", "Category", "Town", "Phone", "Email",
  "Address", "Website Status", "Rating", "Maps Link", "Website URL",
];

// Column widths in pixels
const COL_WIDTHS = [250, 160, 130, 130, 230, 300, 220, 80, 100, 300];

async function getSheets() {
  const auth = new google.auth.GoogleAuth({
    keyFile: path.join(__dirname, "credentials.json"),
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

// Applies header styling, freeze, filter, column widths, alternating rows,
// and conditional formatting for the Website Status column.
export async function formatSheet() {
  const sheets = await getSheets();
  const NUM_COLS = HEADERS.length;

  // ── Step 1: wipe all existing formatting, bandings and conditional rules ──
  const meta      = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
  const sheetMeta = meta.data.sheets.find((s) => s.properties.sheetId === SHEET_ID);

  const cleanupRequests = [];

  // Delete every banding (must be removed before adding new ones, and before clear)
  for (const band of sheetMeta?.bandedRanges || []) {
    cleanupRequests.push({ deleteBanding: { bandedRangeId: band.bandedRangeId } });
  }

  // Delete every conditional format rule (reverse order so indices stay valid)
  const numRules = sheetMeta?.conditionalFormats?.length || 0;
  for (let i = numRules - 1; i >= 0; i--) {
    cleanupRequests.push({ deleteConditionalFormatRule: { sheetId: SHEET_ID, index: i } });
  }

  // Reset ALL cell formatting across the whole sheet to default (empty = no colour, default text)
  cleanupRequests.push({
    repeatCell: {
      range: { sheetId: SHEET_ID },
      cell:  { userEnteredFormat: {} },
      fields: "userEnteredFormat",
    },
  });

  if (cleanupRequests.length) {
    await sheets.spreadsheets.batchUpdate({
      spreadsheetId: SPREADSHEET_ID,
      requestBody: { requests: cleanupRequests },
    });
  }

  // ── Step 2: apply fresh formatting ──
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: {
      requests: [
        // Header row — strictly columns A-J only (no column bounds = bleeds to K-Z)
        {
          repeatCell: {
            range: { sheetId: SHEET_ID, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 0, endColumnIndex: NUM_COLS },
            cell: {
              userEnteredFormat: {
                backgroundColor:     { red: 0.122, green: 0.306, blue: 0.475 },
                textFormat:          { bold: true, foregroundColor: { red: 1, green: 1, blue: 1 }, fontSize: 11 },
                horizontalAlignment: "CENTER",
                verticalAlignment:   "MIDDLE",
              },
            },
            fields: "userEnteredFormat(backgroundColor,textFormat,horizontalAlignment,verticalAlignment)",
          },
        },

        // Freeze header row
        {
          updateSheetProperties: {
            properties: { sheetId: SHEET_ID, gridProperties: { frozenRowCount: 1 } },
            fields: "gridProperties.frozenRowCount",
          },
        },

        // Auto-filter — columns A-J only
        {
          setBasicFilter: {
            filter: {
              range: { sheetId: SHEET_ID, startRowIndex: 0, startColumnIndex: 0, endColumnIndex: NUM_COLS },
            },
          },
        },

        // Column widths
        ...COL_WIDTHS.map((pixels, i) => ({
          updateDimensionProperties: {
            range: { sheetId: SHEET_ID, dimension: "COLUMNS", startIndex: i, endIndex: i + 1 },
            properties: { pixelSize: pixels },
            fields: "pixelSize",
          },
        })),

        // Conditional: "None" in Website Status (col 6) → red bold
        {
          addConditionalFormatRule: {
            rule: {
              ranges: [{ sheetId: SHEET_ID, startRowIndex: 1, startColumnIndex: 6, endColumnIndex: 7 }],
              booleanRule: {
                condition: { type: "TEXT_EQ", values: [{ userEnteredValue: "None" }] },
                format: { textFormat: { bold: true, foregroundColor: { red: 0.753, green: 0, blue: 0 } } },
              },
            },
            index: 0,
          },
        },

        // Conditional: "Poor" in Website Status → orange
        {
          addConditionalFormatRule: {
            rule: {
              ranges: [{ sheetId: SHEET_ID, startRowIndex: 1, startColumnIndex: 6, endColumnIndex: 7 }],
              booleanRule: {
                condition: { type: "TEXT_CONTAINS", values: [{ userEnteredValue: "Poor" }] },
                format: { textFormat: { foregroundColor: { red: 0.929, green: 0.490, blue: 0.192 } } },
              },
            },
            index: 1,
          },
        },
      ],
    },
  });
}

// Creates the header row if the sheet is empty, then applies formatting.
export async function initSheet() {
  const sheets = await getSheets();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A1:K1`,
  });
  if (!res.data.values?.length) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [HEADERS] },
    });
    await formatSheet();
  }
}

// Returns a Set of phone number strings already in the sheet.
export async function loadExistingPhones() {
  const sheets = await getSheets();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!D:D`,
  });
  const phones = new Set();
  (res.data.values || []).forEach((row, i) => {
    if (i === 0) return;
    const raw = row[0]?.toString().trim();
    if (!raw) return;
    const digits = raw.replace(/\D/g, "");
    // Normalize to 07XXXXXXXXX so dedup works regardless of how the number was stored
    if (digits.startsWith("447") && digits.length === 12) phones.add("0" + digits.slice(2));
    else if (digits.startsWith("07") && digits.length === 11) phones.add(digits);
    else if (digits.startsWith("7")  && digits.length === 10) phones.add("0" + digits);
    else phones.add(raw);
  });
  return phones;
}

// Formats a UK mobile number as 07XXX XXXXXX for easy dialling.
function formatPhone(phone) {
  if (!phone) return "";
  const digits = phone.toString().replace(/\D/g, "");
  if (digits.startsWith("07") && digits.length === 11) {
    return `${digits.slice(0, 5)} ${digits.slice(5)}`;
  }
  return phone;
}

// Appends an array of lead objects as new rows.
export async function appendLeads(leads) {
  if (!leads.length) return;
  const sheets = await getSheets();
  const rows = leads.map((l) => [
    l.businessName, l.category, l.town, formatPhone(l.phone), l.email,
    l.address, l.websiteStatus, l.rating,
    l.gbpUrl ? `=HYPERLINK("${l.gbpUrl}","Open")` : "",
    l.websiteUrl || "",
  ]);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:K`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: rows },
  });
}
