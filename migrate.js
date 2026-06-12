import ExcelJS from "exceljs";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import chalk from "chalk";
import { initSheet, loadExistingPhones, appendLeads } from "./sheets.js";

const __dirname        = path.dirname(fileURLToPath(import.meta.url));
const MASTER_XLSX_PATH = path.join(__dirname, "output", "leads_master.xlsx");

const HEADER_MAP = {
  "Business Name":  "businessName",
  "Category":       "category",
  "Town":           "town",
  "Phone":          "phone",
  "Email":          "email",
  "Address":        "address",
  "Website Status": "websiteStatus",
  "Rating":         "rating",
  "Reviews":        "reviews",
  "Maps Link":      "gbpUrl",
  "GBP URL":        "gbpUrl",
};

async function loadExcel() {
  if (!fs.existsSync(MASTER_XLSX_PATH)) {
    console.log(chalk.yellow("No Excel file found at output/leads_master.xlsx"));
    return [];
  }
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(MASTER_XLSX_PATH);
  const sheet = wb.getWorksheet("Leads");
  if (!sheet) return [];

  const headers = [];
  const results = [];
  sheet.eachRow((row, rowNum) => {
    if (rowNum === 1) {
      row.eachCell((cell, col) => {
        headers[col] = HEADER_MAP[cell.value?.toString().trim()] || cell.value;
      });
      return;
    }
    const obj = {};
    row.eachCell((cell, col) => {
      const val = cell.value?.hyperlink ? "" : (cell.value?.toString() ?? "");
      obj[headers[col]] = val;
    });
    if (obj.businessName) results.push(obj);
  });
  return results;
}

async function main() {
  console.log(chalk.cyan("Migrating leads from Excel → Google Sheets..."));

  await initSheet();
  const existingPhones = await loadExistingPhones();
  console.log(chalk.cyan(`${existingPhones.size} leads already in Google Sheets`));

  const leads = await loadExcel();
  if (!leads.length) {
    console.log(chalk.yellow("No leads found in Excel — nothing to migrate"));
    return;
  }
  console.log(chalk.cyan(`Found ${leads.length} leads in Excel`));

  const toMigrate = leads.filter((l) => l.phone && !existingPhones.has(l.phone));
  const alreadyThere = leads.length - toMigrate.length;
  console.log(chalk.cyan(`${toMigrate.length} to migrate, ${alreadyThere} already in Sheets`));

  if (!toMigrate.length) {
    console.log(chalk.green("All leads already in Google Sheets — nothing to do"));
    return;
  }

  const BATCH = 100;
  for (let i = 0; i < toMigrate.length; i += BATCH) {
    await appendLeads(toMigrate.slice(i, i + BATCH));
    console.log(chalk.green(`  Migrated ${Math.min(i + BATCH, toMigrate.length)} / ${toMigrate.length}`));
  }

  console.log(chalk.green(`\nDone! ${toMigrate.length} leads migrated to Google Sheets.`));
}

main().catch((err) => {
  console.error(chalk.red(`[FATAL] ${err.message}`));
  process.exit(1);
});
