import { google } from "googleapis";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const SPREADSHEET_ID = "1o0YGJTzZwTEYlXN1J3ktQq-JqO2jkaPAL3rRHYNU7bc";
const SHEET_ID       = 0;
const SHEET_NAME     = "Sheet1";
const HEADERS        = [
  "Business Name", "Category", "Town", "Phone",
  "Address", "Website Status", "Rating", "Reviews", "Maps Link",
];

// Column widths in pixels
const COL_WIDTHS = [250, 160, 130, 130, 300, 220, 80, 80, 100];
const NUM_COLS   = HEADERS.length; // 9

async function getSheets() {
  const auth = new google.auth.GoogleAuth({
    keyFile: path.join(__dirname, "credentials.json"),
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

export async function formatSheet() {
  const sheets = await getSheets();

  const requests = [
    // Header row: dark blue background, white bold text, centred
    {
      repeatCell: {
        range: { sheetId: SHEET_ID, startRowIndex: 0, endRowIndex: 1 },
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

    // Auto-filter
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

    // Alternating row colours
    {
      addBanding: {
        bandedRange: {
          bandedRangeId: 1,
          range: { sheetId: SHEET_ID, startRowIndex: 1, startColumnIndex: 0, endColumnIndex: NUM_COLS },
          rowProperties: {
            firstBandColor:  { red: 1,     green: 1,     blue: 1     },
            secondBandColor: { red: 0.941, green: 0.957, blue: 0.980 },
          },
        },
      },
    },

    // Conditional: "None" in Website Status (col 5) → red bold
    {
      addConditionalFormatRule: {
        rule: {
          ranges: [{ sheetId: SHEET_ID, startRowIndex: 1, startColumnIndex: 5, endColumnIndex: 6 }],
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
          ranges: [{ sheetId: SHEET_ID, startRowIndex: 1, startColumnIndex: 5, endColumnIndex: 6 }],
          booleanRule: {
            condition: { type: "TEXT_CONTAINS", values: [{ userEnteredValue: "Poor" }] },
            format: { textFormat: { foregroundColor: { red: 0.929, green: 0.490, blue: 0.192 } } },
          },
        },
        index: 1,
      },
    },
  ];

  // Remove existing banding before re-applying to avoid duplicate errors
  try {
    const meta  = await sheets.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID });
    const sheet = meta.data.sheets.find((s) => s.properties.sheetId === SHEET_ID);
    for (const band of sheet?.bandedRanges || []) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId: SPREADSHEET_ID,
        requestBody: { requests: [{ deleteBanding: { bandedRangeId: band.bandedRangeId } }] },
      });
    }
  } catch { /* nothing to remove */ }

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SPREADSHEET_ID,
    requestBody: { requests },
  });
}

// Clears all data in the sheet (keeps the sheet itself).
export async function clearSheet() {
  const sheets = await getSheets();
  await sheets.spreadsheets.values.clear({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}`,
  });
}

// Creates header row (if empty) and applies formatting.
export async function initSheet() {
  const sheets = await getSheets();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A1:A1`,
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
    if (row[0]?.trim()) phones.add(row[0].trim());
  });
  return phones;
}

// Appends an array of lead objects as new rows.
export async function appendLeads(leads) {
  if (!leads.length) return;
  const sheets = await getSheets();
  const rows = leads.map((l) => [
    l.businessName, l.category, l.town, l.phone,
    l.address, l.websiteStatus, l.rating, l.reviews, l.gbpUrl,
  ]);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:I`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: rows },
  });
}
