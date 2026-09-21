
import { mkdir, writeFile } from "node:fs/promises";
import * as XLSX from "xlsx";

const PERFORMANCE_PAGE = "https://www.dol.gov/agencies/eta/foreign-labor/performance";
const OUTPUT_PATH = new URL("../public/perm-dashboard.json", import.meta.url);
const MAX_WORKBOOK_BYTES = 200 * 1024 * 1024;
const DOL_HEADERS = { "user-agent": "PERM-Dashboard-Importer/1.0 (public data dashboard)", "accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8", "referer": "https://www.dol.gov/" };

const landing = await fetch(PERFORMANCE_PAGE, { headers: DOL_HEADERS });
if (!landing.ok) throw new Error("DOL performance index failed (" + landing.status + ").");
const workbookUrl = newestDisclosureUrl(await landing.text());
if (!workbookUrl) throw new Error("No PERM disclosure workbook was found on the DOL performance page.");
const workbookResponse = await fetch(workbookUrl, { headers: DOL_HEADERS });
if (!workbookResponse.ok) throw new Error("DOL workbook download failed (" + workbookResponse.status + ").");
const size = Number(workbookResponse.headers.get("content-length") || 0);
if (size && size > MAX_WORKBOOK_BYTES) throw new Error("The DOL workbook exceeded the importer's configured size limit.");
const workbook = XLSX.read(Buffer.from(await workbookResponse.arrayBuffer()), { type: "buffer", cellDates: true });
const records = workbook.SheetNames.flatMap((sheetName) => readRecords(workbook.Sheets[sheetName]));
if (!records.length) throw new Error("The DOL workbook did not have decision rows with a usable DECISION_DATE column.");
const certified = records.filter((record) => /CERTIFIED|CERTIFICATION/.test(record.status));
const payload = { generatedAt: new Date().toISOString(), sourceUrl: workbookUrl, sourceRelease: workbookUrl.split("/").pop(), chartSeries: { Certified: buildChartSeries(certified), Processed: buildChartSeries(records) } };
await mkdir(new URL("../public/", import.meta.url), { recursive: true });
await writeFile(OUTPUT_PATH, JSON.stringify(payload) + String.fromCharCode(10));

function newestDisclosureUrl(html) {
  const urls = [...String(html).matchAll(/href=["']([^"']+[.]xlsx(?:[?][^"']*)?)["']/gi)].map((match) => new URL(match[1], PERFORMANCE_PAGE).href).filter((url) => /PERM/i.test(url) && !/record[ _-]?layout/i.test(url));
  return [...new Set(urls)].sort((left, right) => fiscalYear(right) - fiscalYear(left))[0];
}
function fiscalYear(url) { return Number(String(url).match(/FY(20[0-9]{2})/i)?.[1] || 0); }
function normalizeHeader(value) { return String(value || "").toUpperCase().replace(/[^A-Z0-9]/g, "_"); }
function firstColumn(columns, names) { return names.map((name) => columns[name]).find((value) => value !== undefined); }
function cellValue(sheet, rowIndex, columnIndex) { return columnIndex === undefined ? "" : sheet[XLSX.utils.encode_cell({ r: rowIndex, c: columnIndex })]?.v; }
function readRecords(sheet) {
  if (!sheet) return [];
  const range = XLSX.utils.decode_range(sheet["!ref"] || "A1:A1");
  const columns = {};
  for (let columnIndex = range.s.c; columnIndex <= range.e.c; columnIndex += 1) {
    const header = normalizeHeader(sheet[XLSX.utils.encode_cell({ r: range.s.r, c: columnIndex })]?.v);
    if (header) columns[header] = columnIndex;
  }
  const decisionColumn = firstColumn(columns, ["DECISION_DATE", "DATE_OF_DECISION", "CASE_DECISION_DATE", "DETERMINATION_DATE", "CASE_DETERMINATION_DATE", "FINAL_DECISION_DATE"]);
  const statusColumn = firstColumn(columns, ["CASE_STATUS", "STATUS", "FINAL_DECISION"]);
  const employerColumn = firstColumn(columns, ["EMPLOYER_NAME", "EMPLOYER_BUSINESS_NAME", "EMP_BUSINESS_NAME"]);
  if (decisionColumn === undefined) return [];
  const records = [];
  for (let rowIndex = range.s.r + 1; rowIndex <= range.e.r; rowIndex += 1) {
    const decisionDate = toDate(cellValue(sheet, rowIndex, decisionColumn));
    if (decisionDate) records.push({ decisionDate, status: String(cellValue(sheet, rowIndex, statusColumn) || "").toUpperCase(), employer: String(cellValue(sheet, rowIndex, employerColumn) || "") });
  }
  return records;
}
function toDate(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number") { const parsed = XLSX.SSF.parse_date_code(value); return parsed ? new Date(parsed.y, parsed.m - 1, parsed.d) : null; }
  const date = new Date(value); return Number.isNaN(date.getTime()) ? null : date;
}
function buildChartSeries(records) {
  const latest = records.reduce((result, record) => record.decisionDate > result ? record.decisionDate : result, new Date(0));
  const dailyDates = Array.from({ length: 30 }, (_, index) => addDays(latest, index - 29));
  const monthDates = Array.from({ length: 12 }, (_, index) => new Date(latest.getFullYear(), latest.getMonth() - 11 + index, 1));
  const dailyCounts = countBy(records, (record) => dayKey(record.decisionDate));
  const weeklyCounts = countBy(records, (record) => weekKey(record.decisionDate));
  const monthlyCounts = countBy(records, (record) => monthKey(record.decisionDate));
  const letters = countBy(records, (record) => employerLetter(record.employer));
  const latestMonth = monthKey(latest);
  return { daily: dailyDates.map((date) => dailyCounts[dayKey(date)] || 0), dailyLabels: dailyDates.map(formatShortDate), weekly: [1, 2, 3, 4, 5, 6, 0].map((day) => records.filter((record) => record.decisionDate.getDay() === day).length), weeklyLabels: ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"], weeklyVolume: Array.from({ length: 4 }, (_, index) => weeklyCounts[weekKey(addDays(latest, -21 + index * 7))] || 0), weeklyVolumeLabels: Array.from({ length: 4 }, (_, index) => formatShortDate(addDays(latest, -21 + index * 7))), monthlyVolume: monthDates.map((date) => monthlyCounts[monthKey(date)] || 0), monthlyVolumeLabels: monthDates.map((date) => date.toLocaleDateString("en-US", { month: "short", year: "numeric" })), activity: "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("").map((letter) => letters[letter] || 0), activeLetters: "ABCDEFGHIJKLMNOPQRSTUVWXYZ".split("").map((letter) => records.filter((record) => monthKey(record.decisionDate) === latestMonth && employerLetter(record.employer) === letter).length), detailedAsOf: latest.toISOString() };
}
function countBy(items, keyForItem) { return items.reduce((counts, item) => { const key = keyForItem(item); counts[key] = (counts[key] || 0) + 1; return counts; }, {}); }
function addDays(date, amount) { const result = new Date(date); result.setDate(result.getDate() + amount); return result; }
function dayKey(date) { return date.toISOString().slice(0, 10); }
function weekKey(date) { return dayKey(addDays(date, -((date.getDay() + 6) % 7))); }
function monthKey(date) { return date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0"); }
function formatShortDate(date) { return date.toLocaleDateString("en-US", { month: "short", day: "numeric" }); }
function employerLetter(value) { const letter = String(value || "").trim().charAt(0).toUpperCase(); return /[A-Z]/.test(letter) ? letter : "X"; }
