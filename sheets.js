import { google } from "googleapis";
import path from "path";
import { fileURLToPath } from "url";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

const SPREADSHEET_ID = "1o0YGJTzZwTEYlXN1J3ktQq-JqO2jkaPAL3rRHYNU7bc";
const SHEET_NAME     = "Sheet1";
const HEADERS        = [
  "Business Name", "Category", "Town", "Phone", "Email",
  "Address", "Website Status", "Rating", "Reviews", "Maps Link",
];

async function getSheets() {
  const auth = new google.auth.GoogleAuth({
    keyFile: path.join(__dirname, "credentials.json"),
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

// Creates the header row if the sheet is empty.
export async function initSheet() {
  const sheets = await getSheets();
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A1:J1`,
  });
  if (!res.data.values?.length) {
    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `${SHEET_NAME}!A1`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [HEADERS] },
    });
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
    if (i === 0) return; // skip header
    if (row[0]?.trim()) phones.add(row[0].trim());
  });
  return phones;
}

// Appends an array of lead objects as new rows.
export async function appendLeads(leads) {
  if (!leads.length) return;
  const sheets = await getSheets();
  const rows = leads.map((l) => [
    l.businessName, l.category, l.town, l.phone, l.email,
    l.address, l.websiteStatus, l.rating, l.reviews, l.gbpUrl,
  ]);
  await sheets.spreadsheets.values.append({
    spreadsheetId: SPREADSHEET_ID,
    range: `${SHEET_NAME}!A:J`,
    valueInputOption: "USER_ENTERED",
    insertDataOption: "INSERT_ROWS",
    requestBody: { values: rows },
  });
}
