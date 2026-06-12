import { chromium } from "playwright-extra";
import StealthPlugin from "puppeteer-extra-plugin-stealth";
import { createObjectCsvWriter } from "csv-writer";
import ExcelJS from "exceljs";
import chalk from "chalk";
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import { SEARCHES, MAX_RESULTS_PER_SEARCH, DELAYS, EXCLUDED_BUSINESS_TYPES } from "./config.js";

chromium.use(StealthPlugin());

// Domains that indicate a poor/placeholder web presence rather than a real website
const SOCIAL_DOMAINS    = ["facebook.com", "fb.com", "instagram.com", "twitter.com", "x.com", "tiktok.com", "linkedin.com", "linktree.com", "linktr.ee"];
const BUILDER_DOMAINS   = ["wix.com", "wixsite.com", "weebly.com", "squarespace.com", "jimdo.com", "yolasite.com", "webnode.com", "site123.com", "moonfruit.com", "godaddysites.com", "wordpress.com", "myfreesitemaker.com"];
const DIRECTORY_DOMAINS = ["yell.com", "checkatrade.com", "ratedpeople.com", "rated-people.com", "trustatrader.com", "bark.com", "mybuilder.com", "booksy.com", "treatwell.com", "fresha.com", "styleseat.com", "vagaro.com", "appointy.com", "setmore.com"];

function classifyWebsite(url) {
  if (!url) return "None";
  const lower = url.toLowerCase();
  if (SOCIAL_DOMAINS.some((d)    => lower.includes(d))) return "Poor - Social media page";
  if (BUILDER_DOMAINS.some((d)   => lower.includes(d))) return "Poor - Free website builder";
  if (DIRECTORY_DOMAINS.some((d) => lower.includes(d))) return "Poor - Directory/booking platform";
  return "Good";
}

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const CHECKPOINT_PATH  = path.join(__dirname, "checkpoint.json");
const SEEN_PATH        = path.join(__dirname, "seen_listings.json");
const OUTPUT_DIR       = path.join(__dirname, "output");
const MASTER_CSV_PATH  = path.join(OUTPUT_DIR, "leads_master.csv");
const MASTER_XLSX_PATH = path.join(OUTPUT_DIR, "leads_master.xlsx");

// Maps column header titles back to internal field names
const CSV_HEADER_MAP = {
  "Business Name":  "businessName",
  "Category":       "category",
  "Town":           "town",
  "Phone":          "phone",
  "Email":          "email",
  "Address":        "address",
  "Website Status": "websiteStatus",
  "Rating":         "rating",
  "Reviews":        "reviews",
  "GBP URL":        "gbpUrl",
  "Maps Link":      "gbpUrl",  // xlsx column header (CSV uses "GBP URL")
};

const CSV_HEADERS = [
  { id: "businessName",  title: "Business Name" },
  { id: "category",      title: "Category" },
  { id: "town",          title: "Town" },
  { id: "phone",         title: "Phone" },
  { id: "email",         title: "Email" },
  { id: "address",       title: "Address" },
  { id: "websiteStatus", title: "Website Status" },
  { id: "rating",        title: "Rating" },
  { id: "reviews",       title: "Reviews" },
  { id: "gbpUrl",        title: "GBP URL" },
];

function parseCSVLine(line) {
  const fields = [];
  let current = "";
  let inQuotes = false;
  for (let i = 0; i < line.length; i++) {
    const ch = line[i];
    if (ch === '"') {
      // Handle escaped double-quote ("") inside a quoted field
      if (inQuotes && line[i + 1] === '"') { current += '"'; i++; }
      else inQuotes = !inQuotes;
    } else if (ch === "," && !inQuotes) {
      fields.push(current);
      current = "";
    } else {
      current += ch;
    }
  }
  fields.push(current);
  return fields;
}

function loadMasterCsv() {
  if (!fs.existsSync(MASTER_CSV_PATH)) return [];
  const lines = fs.readFileSync(MASTER_CSV_PATH, "utf-8").split(/\r?\n/);
  if (lines.length < 2) return [];
  const headers = parseCSVLine(lines[0]);
  const results = [];
  for (let i = 1; i < lines.length; i++) {
    if (!lines[i].trim()) continue;
    const values = parseCSVLine(lines[i]);
    const obj = {};
    headers.forEach((h, idx) => {
      const field = CSV_HEADER_MAP[h.trim()] || h.trim();
      obj[field] = values[idx] ?? "";
    });
    results.push(obj);
  }
  console.log(chalk.cyan(`[MASTER] Loaded ${results.length} existing leads from leads_master.csv`));
  return results;
}

async function writeCsv(filePath, records) {
  const writer = createObjectCsvWriter({ path: filePath, header: CSV_HEADERS });
  await writer.writeRecords(records);
}

// Returns a Set of phone number strings already present in the master xlsx.
async function loadExistingPhones() {
  if (!fs.existsSync(MASTER_XLSX_PATH)) return new Set();
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.readFile(MASTER_XLSX_PATH);
  const sheet = wb.getWorksheet("Leads");
  if (!sheet) return new Set();

  let phoneCol = null;
  sheet.getRow(1).eachCell((cell, col) => {
    if (cell.value?.toString().trim() === "Phone") phoneCol = col;
  });
  if (phoneCol === null) return new Set();

  const phones = new Set();
  sheet.eachRow((row, rowNum) => {
    if (rowNum === 1) return;
    const val = row.getCell(phoneCol).value?.toString().trim();
    if (val) phones.add(val);
  });

  return phones;
}

// Column layout — single source of truth for both fresh creation and key re-establishment.
const SHEET_COLUMNS = [
  { header: "Business Name",  key: "businessName",  width: 35 },
  { header: "Category",       key: "category",      width: 22 },
  { header: "Town",           key: "town",          width: 18 },
  { header: "Phone",          key: "phone",         width: 18 },
  { header: "Email",          key: "email",         width: 32 },
  { header: "Address",        key: "address",       width: 42 },
  { header: "Website Status", key: "websiteStatus", width: 30 },
  { header: "Rating",         key: "rating",        width: 10 },
  { header: "Reviews",        key: "reviews",       width: 10 },
  { header: "Maps Link",      key: "gbpUrl",        width: 14 },
];

// Opens the master xlsx (creating it fresh if absent) and appends newRecords.
// Existing rows are never modified.
async function appendToMasterXlsx(newRecords) {
  if (newRecords.length === 0) return;

  const wb = new ExcelJS.Workbook();
  let sheet;
  let existingDataRows = 0;

  if (fs.existsSync(MASTER_XLSX_PATH)) {
    await wb.xlsx.readFile(MASTER_XLSX_PATH);
    sheet = wb.getWorksheet("Leads");
  }

  if (!sheet) {
    // Fresh file — set up headers and styling from scratch
    sheet = wb.addWorksheet("Leads");
    sheet.columns = SHEET_COLUMNS;

    const headerRow = sheet.getRow(1);
    headerRow.height = 22;
    headerRow.eachCell((cell) => {
      cell.font      = { bold: true, color: { argb: "FFFFFFFF" }, size: 11 };
      cell.fill      = { type: "pattern", pattern: "solid", fgColor: { argb: "FF1F4E79" } };
      cell.alignment = { vertical: "middle", horizontal: "center" };
      cell.border    = { bottom: { style: "thin", color: { argb: "FF2E75B6" } } };
    });
    sheet.views      = [{ state: "frozen", ySplit: 1, activeCell: "A2" }];
    sheet.autoFilter = { from: "A1", to: "J1" };
  } else {
    // Re-establish column keys from the header row so addRow works by key name
    sheet.getRow(1).eachCell((cell, colNum) => {
      const def = SHEET_COLUMNS.find((c) => c.header === cell.value?.toString().trim());
      if (def) sheet.getColumn(colNum).key = def.key;
    });
    existingDataRows = Math.max(0, sheet.rowCount - 1);
  }

  newRecords.forEach((record, index) => {
    const row = sheet.addRow({ ...record, gbpUrl: "" });
    row.height = 18;

    // Alternating background continues from the last existing row
    if ((existingDataRows + index) % 2 === 1) {
      row.eachCell({ includeEmpty: true }, (cell) => {
        cell.fill = { type: "pattern", pattern: "solid", fgColor: { argb: "FFF0F4FA" } };
      });
    }

    if (record.gbpUrl) {
      const linkCell = row.getCell("gbpUrl");
      linkCell.value = { text: "Open", hyperlink: record.gbpUrl };
      linkCell.font  = { color: { argb: "FF0563C1" }, underline: true };
    }

    const status = record.websiteStatus || "";
    const statusCell = row.getCell("websiteStatus");
    if (status === "None") {
      statusCell.font = { color: { argb: "FFC00000" }, bold: true };
    } else if (status.startsWith("Poor")) {
      statusCell.font = { color: { argb: "FFED7D31" } };
    }
  });

  await wb.xlsx.writeFile(MASTER_XLSX_PATH);
}

// Returns the normalised 07XXXXXXXXX string, or null if not a UK mobile.
function normalizeUkMobile(phone) {
  if (!phone) return null;
  const digits = phone.replace(/\D/g, "");
  if (digits.startsWith("447") && digits.length === 12) return "0" + digits.slice(2);
  if (digits.startsWith("07") && digits.length === 11) return digits;
  return null;
}

// Returns true when the business name or GBP category matches an excluded type.
function isExcludedType(businessName, category) {
  const haystack = `${businessName || ""} ${category || ""}`.toLowerCase();
  return EXCLUDED_BUSINESS_TYPES.some((term) => haystack.includes(term.toLowerCase()));
}


function randomDelay([min, max]) {
  return new Promise((r) => setTimeout(r, Math.floor(Math.random() * (max - min + 1)) + min));
}

function loadCheckpoint() {
  if (fs.existsSync(CHECKPOINT_PATH)) {
    return JSON.parse(fs.readFileSync(CHECKPOINT_PATH, "utf-8"));
  }
  return { completed: [], results_so_far: 0 };
}

function saveCheckpoint(checkpoint) {
  fs.writeFileSync(CHECKPOINT_PATH, JSON.stringify(checkpoint, null, 2));
}

function deleteCheckpoint() {
  if (fs.existsSync(CHECKPOINT_PATH)) fs.unlinkSync(CHECKPOINT_PATH);
}

function loadSeenListings() {
  if (fs.existsSync(SEEN_PATH)) {
    return new Set(JSON.parse(fs.readFileSync(SEEN_PATH, "utf-8")));
  }
  return new Set();
}

function saveSeenListings(seen) {
  fs.writeFileSync(SEEN_PATH, JSON.stringify([...seen]));
}

function buildOutputPath() {
  const now = new Date();
  const pad = (n) => String(n).padStart(2, "0");
  const stamp = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}_${pad(now.getHours())}-${pad(now.getMinutes())}`;
  return path.join(OUTPUT_DIR, `leads_${stamp}.csv`);
}

function deduplicateResults(results) {
  const seen = new Set();
  return results.filter((r) => {
    const key = `${r.businessName}|${r.phone}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function buildMapsUrl(category, town) {
  // Encode spaces as + per spec
  const query = `${category} in ${town} UK`.replace(/ /g, "+");
  return `https://www.google.com/maps/search/${query}?hl=en`;
}

async function dismissCookieBanner(page) {
  try {
    const candidates = [
      'button:has-text("Accept all")',
      'button:has-text("Reject all")',
      'button:has-text("Accept")',
      'button[aria-label*="Accept"]',
      'form:has(button[value="2"]) button', // Google consent form
    ];
    for (const sel of candidates) {
      const btn = page.locator(sel).first();
      if (await btn.isVisible({ timeout: 1500 }).catch(() => false)) {
        await btn.click();
        await randomDelay([800, 1500]);
        return;
      }
    }
  } catch { /* no banner */ }
}

// Collect place URLs from the results feed by scrolling incrementally
async function getListingUrls(page, maxResults) {
  const seenUrls = new Set();
  let noNewCount = 0;
  let lastCount = 0;

  while (seenUrls.size < maxResults) {
    const links = await page.locator('div[role="feed"] a[href*="/maps/place/"]').all();

    for (const link of links) {
      if (seenUrls.size >= maxResults) break;
      const href = await link.getAttribute("href").catch(() => null);
      if (!href) continue;
      // Strip query string to deduplicate across scroll positions
      const clean = href.split("?")[0];
      seenUrls.add(clean);
    }

    if (seenUrls.size === lastCount) {
      noNewCount++;
      if (noNewCount >= 3) break;
    } else {
      noNewCount = 0;
    }
    lastCount = seenUrls.size;

    // Scroll the feed panel
    const scrolled = await page.evaluate(() => {
      const feed = document.querySelector('div[role="feed"]');
      if (!feed) return false;
      feed.scrollTop += 600;
      return true;
    });

    if (!scrolled) break;
    await randomDelay([800, 1500]);
  }

  return Array.from(seenUrls).slice(0, maxResults);
}

async function extractListingDetails(page) {
  return page.evaluate(() => {
    // Business name
    const businessName = document.querySelector("h1")?.textContent?.trim() || null;

    // Phone — three fallback methods
    let phone = null;

    // 1. data-item-id starting with "phone:"
    const phoneEl = document.querySelector('[data-item-id^="phone:"]');
    if (phoneEl) {
      const raw = (phoneEl.getAttribute("aria-label") || phoneEl.textContent || "")
        .replace(/^phone:\s*/i, "").trim();
      if (raw.length > 3) phone = raw;
    }

    // 2. Any button/element with aria-label containing "Phone"
    if (!phone) {
      for (const el of document.querySelectorAll("[aria-label]")) {
        const lbl = el.getAttribute("aria-label") || "";
        if (/phone/i.test(lbl)) {
          const cleaned = lbl.replace(/^phone:\s*/i, "").trim();
          if (cleaned.length > 3) { phone = cleaned; break; }
        }
      }
    }

    // 3. tel: link
    if (!phone) {
      const telLink = document.querySelector('a[href^="tel:"]');
      if (telLink) phone = decodeURIComponent(telLink.href.replace("tel:", "")).trim();
    }

    // Extract the GBP "Website" button URL if present
    let websiteUrl = null;
    const authorityLink = document.querySelector('a[data-item-id="authority"]');
    if (authorityLink) {
      websiteUrl = authorityLink.href || null;
    }
    if (!websiteUrl) {
      // Fallback: aria-label explicitly says "website"
      for (const el of document.querySelectorAll("a[aria-label]")) {
        const label = (el.getAttribute("aria-label") || "").trim().toLowerCase();
        if (label === "website" || label === "open website" || label.startsWith("website:")) {
          websiteUrl = el.href || null;
          break;
        }
      }
    }

    // Address
    let address = null;
    const addrEl = document.querySelector('[data-item-id="address"]');
    if (addrEl) {
      address = (addrEl.getAttribute("aria-label") || addrEl.textContent || "")
        .replace(/^address:\s*/i, "").trim();
    }

    // GBP category (subtitle under name)
    let gbpCategory = null;
    const catCandidates = [
      document.querySelector('button[jsaction*="category"]'),
      document.querySelector(".DkEaL"),
      document.querySelector('[class*="fontBodyMedium"] button'),
    ];
    for (const el of catCandidates) {
      if (el?.textContent?.trim()) { gbpCategory = el.textContent.trim(); break; }
    }

    // Rating and reviews
    let rating = null;
    let reviews = null;
    const ratingEl = document.querySelector('[aria-label*="star"]');
    if (ratingEl) {
      const lbl = ratingEl.getAttribute("aria-label") || "";
      const rm = lbl.match(/([\d.]+)\s*star/i);
      if (rm) rating = rm[1];
      const vm = lbl.match(/([\d,]+)\s*review/i);
      if (vm) reviews = vm[1].replace(/,/g, "");
    }

    return { businessName, phone, websiteUrl, address, gbpCategory, rating, reviews };
  });
}

async function scrapeSearch(page, category, town, seenListings) {
  const results = [];
  const skipCounts = { hasWebsite: 0, noPhone: 0, notMobile: 0, excludedType: 0 };
  const mapsUrl = buildMapsUrl(category, town);

  console.log(chalk.green(`\n[START] ${category} in ${town}`));

  await page.goto(mapsUrl, { waitUntil: "domcontentloaded", timeout: 30000 });
  await randomDelay(DELAYS.betweenActions);
  await dismissCookieBanner(page);

  // Wait for the results feed to appear
  try {
    await page.waitForSelector('div[role="feed"]', { timeout: 15000 });
    await randomDelay([1000, 2000]);
  } catch {
    console.log(chalk.yellow(`  No results feed for: ${category} in ${town}`));
    return { results, skipCounts };
  }

  const allUrls = await getListingUrls(page, MAX_RESULTS_PER_SEARCH);
  const listingUrls = allUrls.filter((u) => !seenListings.has(u));
  const skippedSeen = allUrls.length - listingUrls.length;
  console.log(chalk.green(`  Found ${allUrls.length} listings — ${listingUrls.length} new, ${skippedSeen} already seen`))

  for (const placeUrl of listingUrls) {
    try {
      // Navigate directly to the place page
      await page.goto(`${placeUrl}?hl=en`, { waitUntil: "domcontentloaded", timeout: 20000 });
      await randomDelay(DELAYS.betweenActions);
      await page.waitForSelector("h1", { timeout: 10000 }).catch(() => {});
      await randomDelay([500, 1000]);

      const details = await extractListingDetails(page);
      const mapsPageUrl = page.url();

      if (!details.businessName) {
        console.log(chalk.yellow(`  [SKIP] Could not extract name`));
        seenListings.add(placeUrl);
        continue;
      }

      const websiteStatus = classifyWebsite(details.websiteUrl);

      if (websiteStatus === "Good") {
        console.log(chalk.yellow(`  [SKIP] ${details.businessName}: has website`));
        skipCounts.hasWebsite++;
        seenListings.add(placeUrl);
        continue;
      }

      if (!details.phone) {
        console.log(chalk.yellow(`  [SKIP] ${details.businessName}: no phone`));
        skipCounts.noPhone++;
        seenListings.add(placeUrl);
        continue;
      }

      const mobilePhone = normalizeUkMobile(details.phone);
      if (!mobilePhone) {
        console.log(chalk.yellow(`  [SKIP] ${details.businessName}: not a UK mobile (${details.phone})`));
        skipCounts.notMobile++;
        seenListings.add(placeUrl);
        continue;
      }

      const resolvedCategory = details.gbpCategory || category;
      if (isExcludedType(details.businessName, resolvedCategory)) {
        console.log(chalk.yellow(`  [SKIP] ${details.businessName}: excluded business type`));
        skipCounts.excludedType++;
        seenListings.add(placeUrl);
        continue;
      }

      const lead = {
        businessName: details.businessName,
        category: resolvedCategory,
        town,
        phone: mobilePhone,
        email: "",
        address: details.address || "",
        websiteStatus,
        rating: details.rating || "",
        reviews: details.reviews || "",
        gbpUrl: mapsPageUrl,
      };

      results.push(lead);
      seenListings.add(placeUrl);
      console.log(chalk.green(`  [LEAD] ${details.businessName} [${websiteStatus}]`));

    } catch (err) {
      // Don't add to seen on error — allow retry on next run
      console.log(chalk.yellow(`  [ERROR] ${err.message?.slice(0, 120)}`));
    }

    await randomDelay(DELAYS.betweenListings);
  }

  return { results, skipCounts };

}

async function main() {
  if (!fs.existsSync(OUTPUT_DIR)) fs.mkdirSync(OUTPUT_DIR, { recursive: true });

  // Load existing phone numbers upfront — used to skip duplicates after scraping
  const existingPhones = await loadExistingPhones();
  console.log(chalk.cyan(`[MASTER] ${existingPhones.size} leads already in leads_master.xlsx`));

  // Load seen listing URLs — these are skipped entirely on future runs
  const seenListings = loadSeenListings();
  console.log(chalk.cyan(`[SEEN]   ${seenListings.size} listings already processed (will be skipped)`));

  const checkpoint = loadCheckpoint();
  if (checkpoint.completed.length > 0) {
    console.log(chalk.cyan(`[RESUME] Checkpoint found — skipping ${checkpoint.completed.length} completed searches`));
  }

  const totalSkips = { hasWebsite: 0, noPhone: 0, notMobile: 0, excludedType: 0 };
  let totalAdded = 0;

  const browser = await chromium.launch({ headless: false });
  const context = await browser.newContext({
    userAgent: "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    locale: "en-GB",
    timezoneId: "Europe/London",
  });
  const page = await context.newPage();

  try {
    for (const { category, towns } of SEARCHES) {
      for (const town of towns) {
        const key = `${category}|${town}`;
        if (checkpoint.completed.includes(key)) {
          console.log(chalk.cyan(`[SKIP] Already completed: ${key}`));
          continue;
        }

        try {
          const { results, skipCounts } = await scrapeSearch(page, category, town, seenListings);

          totalSkips.hasWebsite   += skipCounts.hasWebsite;
          totalSkips.noPhone      += skipCounts.noPhone;
          totalSkips.notMobile    += skipCounts.notMobile;
          totalSkips.excludedType += skipCounts.excludedType;

          // Write leads to Excel immediately — nothing is lost if the run is stopped
          const unique = deduplicateResults(results).filter((r) => !existingPhones.has(r.phone));
          if (unique.length > 0) {
            await appendToMasterXlsx(unique);
            unique.forEach((r) => existingPhones.add(r.phone));
            totalAdded += unique.length;
          }

          checkpoint.completed.push(key);
          checkpoint.results_so_far = totalAdded;
          saveCheckpoint(checkpoint);
          saveSeenListings(seenListings);

          console.log(chalk.green(`  [DONE] ${key} — ${results.length} leads found, ${unique.length} added to master`));
        } catch (err) {
          console.log(chalk.red(`[FAIL] Search ${key}: ${err.message}`));
        }

        await randomDelay(DELAYS.betweenSearches);
      }
    }
  } finally {
    await browser.close();
  }

  const masterTotal = existingPhones.size;
  const totalSkipped = Object.values(totalSkips).reduce((a, b) => a + b, 0);
  console.log("\n" + "=".repeat(50));
  console.log(chalk.green(`New leads added this run:       ${totalAdded}`));
  console.log(chalk.green(`Total leads in master file:     ${masterTotal}`));
  console.log(chalk.yellow(`Total listings skipped:         ${totalSkipped}`));
  console.log(chalk.yellow(`  Has website:            ${totalSkips.hasWebsite}`));
  console.log(chalk.yellow(`  No phone:               ${totalSkips.noPhone}`));
  console.log(chalk.yellow(`  Not UK mobile:          ${totalSkips.notMobile}`));
  console.log(chalk.yellow(`  Excluded business type: ${totalSkips.excludedType}`));
  console.log(chalk.green(`Master file:                    ${MASTER_XLSX_PATH}`));
  console.log("=".repeat(50));

  deleteCheckpoint();
}

main().catch((err) => {
  console.error(chalk.red(`[FATAL] ${err.message}`));
  process.exit(1);
});
