
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
console.log("Workbook diagnostics", JSON.stringify(workbook.SheetNames.map((sheetName) => { const sheet = workbook.Sheets[sheetName]; return { sheetName, ref: sheet?.["!ref"], firstRow: sheet ? Array.from({ length: 12 }, (_, columnIndex) => sheet[XLSX.utils.encode_cell({ r: 0, c: columnIndex })]?.v ?? null) : null }; })));
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
