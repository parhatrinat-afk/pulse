/**
 * Pulse 0.3.0 — Weekly Mapping Attention Foundation.
 *
 * Materializes repository-validated, derived Mapping evidence only. Mapping
 * Rules, Effective Mapping, facts, weekly cache, Performance, and Reports are
 * never changed. Run in bounded Prepare -> Write -> Finalize actions.
 */
function main(
  workbook: ExcelScript.Workbook,
  action: string,
  payloadJson: string
): string {
  const requested = text(action);
  const payload = parsePayload(payloadJson);
  if (requested === "Prepare") return prepare(workbook, payload);
  if (requested === "Write") return writeChunk(workbook, payload);
  if (requested === "Finalize") return finalize(workbook, payload);
  throw new Error(`PUL-030MA-200: Unsupported action ${requested || "(blank)"}.`);
}

type CellValue = string | number | boolean;
type Payload = {
  projectionFingerprint?: string;
  activeCacheVersion?: string;
  activeCacheFingerprint?: string;
  section?: string;
  offset?: number;
  totalRowCount?: number;
  startRow?: number;
  startColumn?: number;
  values?: CellValue[][];
};
type Section = {
  name: string;
  tableName: string;
  startRow: number;
  startColumn: number;
  headers: string[];
  rowCount: number;
};
type Metric = { factCount: number; salesNok: number; quantity: number };

const FINAL_SHEET = "_Weekly_Mapping_Attention";
const STAGING_SHEET = "_Weekly_Mapping_Attn_Stage";
const EXPECTED_PROJECTION = "WMA-637d7a94536ac4ed";
const EXPECTED_CACHE_VERSION = "WCV-1b0b195c210da456";
const EXPECTED_CACHE_FINGERPRINT = "WCC-26c195956ebc2823";
const EXPECTED_MAPPING_CONTENT = "MCF-759cc92c4304a913";
const EXPECTED_CATALOG_CONTENT = "ICC-5644a77c18a97437";
const EXPECTED_IDENTITY_PREFLIGHT = "IDP-4ae1a62974cca3af";
const EXPECTED_PERIODS = 85;
const EXPECTED_PRODUCTS = 1237;
const EXPECTED_EXISTING_PRODUCTS = 1041;
const EXPECTED_WEEKLY_PRODUCTS = 196;
const EXPECTED_FACTS = 248572;
const EXPECTED_SALES = 490365140.75;
const EXPECTED_QUANTITY = 2499642.44;

const CONTROL_HEADERS = [
  "ProjectionVersion", "ValidationStatus", "HealthStatus", "ProjectionFingerprint",
  "ActiveCacheVersion", "ActiveCacheFingerprint", "SourceCorpusFingerprint",
  "MappingContentFingerprint", "CatalogContentFingerprint", "IdentityPreflightFingerprint",
  "ThroughPeriodLabel", "PeriodRowCount", "ProductRowCount", "ExistingProductCount",
  "WeeklyAddedProductCount", "SourceFactCount", "SourceSalesNOK", "SourceQuantity"
];
const PRODUCT_HEADERS = [
  "ProductID", "SourceSystemID", "Item", "Main Category", "Subcategory", "Sales Account",
  "SourceClassificationID", "ProductKey", "Identity Origin", "ReportingGroupID",
  "Reporting Group", "Resolution", "Mapping State", "WinningRuleID", "Historical Facts",
  "Historical Sales NOK", "Historical Quantity", "Hierarchy Attention", "Hierarchy Alternatives"
];

function prepare(workbook: ExcelScript.Workbook, payload: Payload): string {
  validatePayload(payload);
  const authority = validateActiveAuthority(workbook);
  const existing = workbook.getWorksheet(FINAL_SHEET);
  if (existing) {
    const control = workbook.getTable("tblWeeklyMappingAttentionControl");
    const rows = workbook.getTable("tblWeeklyMappingAttention");
    if (control && rows) {
      const values = control.getRangeBetweenHeaderAndTotal().getValues();
      if (values.length === 1 && text(values[0][3]) === EXPECTED_PROJECTION &&
          text(values[0][4]) === EXPECTED_CACHE_VERSION && rows.getRowCount() === EXPECTED_PRODUCTS) {
        writeMappingHealth(workbook, existing, false);
        return JSON.stringify({ status: "Already Materialized", authority, products: rows.getRowCount() });
      }
    }
    throw new Error("PUL-030MA-201: Existing Mapping attention surface differs from accepted evidence.");
  }
  const prior = workbook.getWorksheet(STAGING_SHEET);
  if (prior) prior.delete();
  const staging = workbook.addWorksheet(STAGING_SHEET);
  staging.setVisibility(ExcelScript.SheetVisibility.hidden);
  const all = sections();
  for (let index = 0; index < all.length; index += 1) {
    const section = all[index];
    checkedRange(staging, section.startRow, section.startColumn, 1, section.headers.length,
      `${section.name} header`).setValues([section.headers]);
  }
  return JSON.stringify({ status: "Prepared", sheet: STAGING_SHEET, authority });
}

function writeChunk(workbook: ExcelScript.Workbook, payload: Payload): string {
  validatePayload(payload);
  const staging = workbook.getWorksheet(STAGING_SHEET);
  if (!staging) throw new Error("PUL-030MA-202: Staging sheet is missing. Run Prepare first.");
  if (workbook.getWorksheet(FINAL_SHEET)) {
    throw new Error("PUL-030MA-203: Final attention surface already exists.");
  }
  const section = requiredSection(payload.section);
  const values = payload.values || [];
  const offset = integer(payload.offset, "offset");
  const startRow = integer(payload.startRow, "startRow");
  const startColumn = integer(payload.startColumn, "startColumn");
  const total = integer(payload.totalRowCount, "totalRowCount");
  if (total !== section.rowCount || startRow !== section.startRow + 1 + offset ||
      startColumn !== section.startColumn || offset + values.length > section.rowCount) {
    throw new Error(`PUL-030MA-204: ${section.name} chunk geometry is invalid.`);
  }
  for (let index = 0; index < values.length; index += 1) {
    if (values[index].length !== section.headers.length) {
      throw new Error(`PUL-030MA-205: ${section.name} chunk has an invalid column count.`);
    }
  }
  if (values.length) {
    checkedRange(staging, startRow, startColumn, values.length, section.headers.length,
      `${section.name} values`).setValues(values);
  }
  return JSON.stringify({ status: "Written", section: section.name, offset, rows: values.length });
}

function finalize(workbook: ExcelScript.Workbook, payload: Payload): string {
  validatePayload(payload);
  if (workbook.getWorksheet(FINAL_SHEET)) {
    throw new Error("PUL-030MA-206: Final attention surface already exists. Run Prepare for idempotent verification.");
  }
  const staging = workbook.getWorksheet(STAGING_SHEET);
  if (!staging) throw new Error("PUL-030MA-207: Staging sheet is missing.");
  const authority = validateActiveAuthority(workbook);
  const all = sections();
  for (let index = 0; index < all.length; index += 1) assertHeaders(staging, all[index]);
  const control = checkedRange(staging, 2, 1, 1, CONTROL_HEADERS.length,
    "control values").getValues()[0];
  const products = checkedRange(staging, 6, 1, EXPECTED_PRODUCTS, PRODUCT_HEADERS.length,
    "product values").getValues();
  const validation = validateProjection(control, products);
  for (let index = 0; index < all.length; index += 1) {
    const section = all[index];
    const range = checkedRange(staging, section.startRow, section.startColumn,
      section.rowCount + 1, section.headers.length, `${section.name} table`);
    const table = staging.addTable(range, true);
    table.setName(section.tableName);
    table.setPredefinedTableStyle("TableStyleMedium2");
  }
  formatEngineeringSurface(staging);
  staging.setName(FINAL_SHEET);
  staging.setVisibility(ExcelScript.SheetVisibility.hidden);
  writeMappingHealth(workbook, staging, false);
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  assertInteractionQa(workbook, "tblPerformanceInteractionQA");
  assertInteractionQa(workbook, "tblWeeklyPerformanceQA");
  return JSON.stringify({
    status: "PASS", sheet: FINAL_SHEET, visibility: "Hidden", authority, validation
  });
}

function validateProjection(control: CellValue[], rows: CellValue[][]): {
  productRows: number; states: { [key: string]: { products: number; facts: number; salesNok: number } };
  source: Metric;
} {
  if (text(control[0]) !== "0.3.0-weekly-mapping-attention-v1" ||
      text(control[1]) !== "PASS" || text(control[2]) !== "Attention required" ||
      text(control[3]) !== EXPECTED_PROJECTION || text(control[4]) !== EXPECTED_CACHE_VERSION ||
      text(control[5]) !== EXPECTED_CACHE_FINGERPRINT ||
      text(control[7]) !== EXPECTED_MAPPING_CONTENT || text(control[8]) !== EXPECTED_CATALOG_CONTENT ||
      text(control[9]) !== EXPECTED_IDENTITY_PREFLIGHT || text(control[10]) !== "2026 W33" ||
      numberValue(control[11]) !== EXPECTED_PERIODS || numberValue(control[12]) !== EXPECTED_PRODUCTS ||
      numberValue(control[13]) !== EXPECTED_EXISTING_PRODUCTS ||
      numberValue(control[14]) !== EXPECTED_WEEKLY_PRODUCTS || numberValue(control[15]) !== EXPECTED_FACTS ||
      Math.abs(numberValue(control[16]) - EXPECTED_SALES) > 0.005 ||
      Math.abs(numberValue(control[17]) - EXPECTED_QUANTITY) > 0.000005) {
    throw new Error("PUL-030MA-208: Control row differs from accepted 85-week evidence.");
  }
  const productIds: { [key: string]: boolean } = {};
  const productKeys: { [key: string]: boolean } = {};
  const states: { [key: string]: { products: number; facts: number; salesNok: number } } = {};
  const source = emptyMetric();
  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];
    const productId = text(row[0]);
    const productKey = text(row[7]);
    if (!productId || productIds[productId]) throw new Error(`PUL-030MA-209: Duplicate/blank ProductID ${productId}.`);
    if (productKey && productKeys[productKey]) throw new Error(`PUL-030MA-210: Duplicate ProductKey ${productKey}.`);
    productIds[productId] = true;
    if (productKey) productKeys[productKey] = true;
    const stateName = text(row[12]);
    if (!states[stateName]) states[stateName] = { products: 0, facts: 0, salesNok: 0 };
    states[stateName].products += 1;
    states[stateName].facts += numberValue(row[14]);
    states[stateName].salesNok += numberValue(row[15]);
    source.factCount += numberValue(row[14]);
    source.salesNok += numberValue(row[15]);
    source.quantity += numberValue(row[16]);
  }
  if (source.factCount !== EXPECTED_FACTS || Math.abs(source.salesNok - EXPECTED_SALES) > 0.005 ||
      Math.abs(source.quantity - EXPECTED_QUANTITY) > 0.000005 ||
      stateProducts(states, "Mapped") !== 929 || stateProducts(states, "Unmapped") !== 302 ||
      stateProducts(states, "Identity Pending") !== 6 || stateProducts(states, "Conflict") !== 0 ||
      stateProducts(states, "Inactive Target") !== 0 || stateFacts(states, "Identity Pending") !== 126 ||
      Math.abs(stateSales(states, "Identity Pending") - 120048) > 0.005) {
    throw new Error("PUL-030MA-211: Product projection does not reconcile to accepted coverage.");
  }
  return { productRows: rows.length, states, source };
}

function validateActiveAuthority(workbook: ExcelScript.Workbook): {
  cacheVersion: string; cacheFingerprint: string; activeCount: number;
} {
  const table = requiredTable(workbook, "tblWeeklyCacheVersions");
  const headers = headerMap(table);
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  let activeCount = 0;
  let version = "";
  let fingerprint = "";
  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];
    if (text(row[headers.CacheStatus]) !== "Active" || text(row[headers.ActivationState]) !== "Active") continue;
    activeCount += 1;
    version = text(row[headers.CacheVersion]);
    fingerprint = text(row[headers.CacheFingerprint]);
    if (text(row[headers.ValidationStatus]) !== "PASS" ||
        text(row[headers.MappingContentFingerprint]) !== EXPECTED_MAPPING_CONTENT ||
        text(row[headers.CatalogContentFingerprint]) !== EXPECTED_CATALOG_CONTENT ||
        text(row[headers.IdentityPreflightFingerprint]) !== EXPECTED_IDENTITY_PREFLIGHT) {
      throw new Error("PUL-030MA-212: Active weekly authority is stale or unvalidated.");
    }
  }
  if (activeCount !== 1 || version !== EXPECTED_CACHE_VERSION || fingerprint !== EXPECTED_CACHE_FINGERPRINT) {
    throw new Error("PUL-030MA-213: Expected exactly one accepted Active / Active weekly cache.");
  }
  return { cacheVersion: version, cacheFingerprint: fingerprint, activeCount };
}

function writeMappingHealth(
  workbook: ExcelScript.Workbook,
  attentionSheet: ExcelScript.Worksheet,
  refreshRequired: boolean
): void {
  const mapping = workbook.getWorksheet("Mapping");
  if (!mapping) throw new Error("PUL-030MA-214: Mapping sheet is missing.");
  const productTable = requiredTable(workbook, "tblWeeklyMappingAttention");
  const controlTable = requiredTable(workbook, "tblWeeklyMappingAttentionControl");
  if (productTable.getWorksheet().getName() !== attentionSheet.getName()) {
    throw new Error("PUL-030MA-215: Attention table is outside its engineering sheet.");
  }
  const rows = productTable.getRangeBetweenHeaderAndTotal().getValues();
  const headers = headerMap(productTable);
  const stateNames = ["Mapped", "Unmapped", "Identity Pending", "Conflict", "Inactive Target"];
  const counts: { [key: string]: { products: number; facts: number; salesNok: number } } = {};
  for (let stateIndex = 0; stateIndex < stateNames.length; stateIndex += 1) {
    counts[stateNames[stateIndex]] = { products: 0, facts: 0, salesNok: 0 };
  }
  for (let index = 0; index < rows.length; index += 1) {
    const state = text(rows[index][headers["Mapping State"]]);
    if (!counts[state]) continue;
    counts[state].products += 1;
    counts[state].facts += numberValue(rows[index][headers["Historical Facts"]]);
    counts[state].salesNok += numberValue(rows[index][headers["Historical Sales NOK"]]);
  }
  const control = controlTable.getRangeBetweenHeaderAndTotal().getValues()[0];
  const stateHeaders = ["Metric", "Mapped", "Unmapped", "Identity Pending", "Conflict", "Inactive Target"];
  const productValues: CellValue[] = ["Products"];
  const factValues: CellValue[] = ["Historical facts"];
  const salesValues: CellValue[] = ["Historical Sales NOK"];
  for (let index = 0; index < stateNames.length; index += 1) {
    const state = counts[stateNames[index]];
    productValues.push(state.products);
    factValues.push(state.facts);
    salesValues.push(state.salesNok);
  }
  mapping.getRange("A4:N9").unmerge();
  mapping.getRange("A4:N9").clear(ExcelScript.ClearApplyTo.formats);
  mapping.getRange("A4:N4").merge();
  mapping.getRange("A4").setValue(`Weekly mapping health — ${text(control[2])}`);
  mapping.getRange("A5:F5").setValues([stateHeaders]);
  mapping.getRange("A6:F6").setValues([productValues]);
  mapping.getRange("A7:F7").setValues([factValues]);
  mapping.getRange("A8:F8").setValues([salesValues]);
  mapping.getRange("A9:N9").merge();
  const through = text(control[10]);
  mapping.getRange("A9").setValue(refreshRequired
    ? "Performance refresh required"
    : `Performance classifications are up to date through ${through}`);
  const navy = "#17365D";
  const blue = "#D9EAF7";
  const amber = "#FFF4CE";
  mapping.getRange("A4:N4").getFormat().getFill().setColor(navy);
  mapping.getRange("A4:N4").getFormat().getFont().setColor("#FFFFFF");
  mapping.getRange("A4:N4").getFormat().getFont().setBold(true);
  mapping.getRange("A5:F5").getFormat().getFill().setColor(blue);
  mapping.getRange("A5:F5").getFormat().getFont().setBold(true);
  mapping.getRange("A6:A8").getFormat().getFont().setBold(true);
  mapping.getRange("B6:F7").setNumberFormat("#,##0");
  mapping.getRange("B8:F8").setNumberFormat("#,##0.00");
  mapping.getRange("A9:N9").getFormat().getFill().setColor(refreshRequired ? amber : "#E2F0D9");
  mapping.getRange("A9:N9").getFormat().getFont().setBold(true);
  mapping.getRange("A4:N9").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  mapping.getRange("A4:N9").getFormat().setWrapText(true);
  mapping.getRange("4:5").getFormat().setRowHeight(26);
  mapping.getRange("6:8").getFormat().setRowHeight(22);
  mapping.getRange("9:9").getFormat().setRowHeight(28);
}

function assertInteractionQa(workbook: ExcelScript.Workbook, tableName: string): void {
  const table = requiredTable(workbook, tableName);
  const headers = headerMap(table);
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  const resultColumn = headers.Status === undefined ? headers.Result : headers.Status;
  if (resultColumn === undefined) {
    throw new Error(`PUL-030MA-224: ${tableName} has no Status/Result column.`);
  }
  let pass = 0;
  for (let index = 0; index < rows.length; index += 1) {
    if (text(rows[index][resultColumn]) === "PASS") pass += 1;
  }
  if (rows.length !== 16 || pass !== 16) {
    throw new Error(`PUL-030MA-216: ${tableName} is ${pass}/${rows.length}, expected 16/16 PASS.`);
  }
}

function sections(): Section[] {
  return [
    { name: "control", tableName: "tblWeeklyMappingAttentionControl", startRow: 1,
      startColumn: 1, headers: CONTROL_HEADERS, rowCount: 1 },
    { name: "products", tableName: "tblWeeklyMappingAttention", startRow: 5,
      startColumn: 1, headers: PRODUCT_HEADERS, rowCount: EXPECTED_PRODUCTS },
  ];
}

function requiredSection(name: string | undefined): Section {
  const all = sections();
  for (let index = 0; index < all.length; index += 1) if (all[index].name === text(name)) return all[index];
  throw new Error(`PUL-030MA-217: Unknown section ${text(name)}.`);
}

function formatEngineeringSurface(sheet: ExcelScript.Worksheet): void {
  sheet.getFreezePanes().freezeRows(5);
  sheet.getRange("A1:S1242").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRange("O6:O1242").setNumberFormat("#,##0");
  sheet.getRange("P6:Q1242").setNumberFormat("#,##0.00");
  sheet.getRange("A:S").getFormat().setColumnWidth(110);
  sheet.getRange("C:C").getFormat().setColumnWidth(220);
  sheet.getRange("D:F").getFormat().setColumnWidth(170);
  sheet.getRange("H:H").getFormat().setColumnWidth(280);
  sheet.getRange("R:S").getFormat().setColumnWidth(320);
}

function validatePayload(payload: Payload): void {
  if (text(payload.projectionFingerprint) !== EXPECTED_PROJECTION ||
      text(payload.activeCacheVersion) !== EXPECTED_CACHE_VERSION ||
      text(payload.activeCacheFingerprint) !== EXPECTED_CACHE_FINGERPRINT) {
    throw new Error("PUL-030MA-218: Payload does not identify the accepted weekly attention evidence.");
  }
}

function assertHeaders(sheet: ExcelScript.Worksheet, section: Section): void {
  const values = checkedRange(sheet, section.startRow, section.startColumn, 1,
    section.headers.length, `${section.name} headers`).getValues()[0];
  for (let index = 0; index < section.headers.length; index += 1) {
    if (text(values[index]) !== section.headers[index]) {
      throw new Error(`PUL-030MA-219: ${section.name} header ${index + 1} differs.`);
    }
  }
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-030MA-220: Required table ${name} is missing.`);
  return table;
}

function headerMap(table: ExcelScript.Table): { [key: string]: number } {
  const values = table.getHeaderRowRange().getValues()[0];
  const result: { [key: string]: number } = {};
  for (let index = 0; index < values.length; index += 1) result[text(values[index])] = index;
  return result;
}

function checkedRange(
  sheet: ExcelScript.Worksheet, startRow: number, startColumn: number,
  rowCount: number, columnCount: number, label: string
): ExcelScript.Range {
  if (startRow < 1 || startColumn < 1 || rowCount < 1 || columnCount < 1) {
    throw new Error(`PUL-030MA-221: Invalid ${label} range geometry.`);
  }
  return sheet.getRangeByIndexes(startRow - 1, startColumn - 1, rowCount, columnCount);
}

function parsePayload(value: string): Payload {
  try {
    const parsed = JSON.parse(value || "{}");
    return parsed as Payload;
  } catch (error) {
    throw new Error(`PUL-030MA-222: payloadJson is invalid JSON. ${text(error)}`);
  }
}

function integer(value: number | undefined, label: string): number {
  const result = Number(value);
  if (!Number.isInteger(result) || result < 0) throw new Error(`PUL-030MA-223: ${label} is invalid.`);
  return result;
}

function emptyMetric(): Metric { return { factCount: 0, salesNok: 0, quantity: 0 }; }
function stateProducts(states: { [key: string]: { products: number } }, name: string): number {
  return states[name] ? states[name].products : 0;
}
function stateFacts(states: { [key: string]: { facts: number } }, name: string): number {
  return states[name] ? states[name].facts : 0;
}
function stateSales(states: { [key: string]: { salesNok: number } }, name: string): number {
  return states[name] ? states[name].salesNok : 0;
}
function text(value: unknown): string { return value === null || value === undefined ? "" : String(value); }
function numberValue(value: unknown): number { const result = Number(value); return Number.isFinite(result) ? result : 0; }
