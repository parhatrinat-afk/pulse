/**
 * Pulse Build 0.3.0 — activate the accepted compact weekly cache.
 *
 * The script repeats the full semantic-content and materialized-cache
 * preflight, then changes only CacheStatus and ActivationState on the accepted
 * version row. It does not change analytical cache rows, Performance, Reports,
 * imports, facts, selectors, mappings, or workbook visibility.
 */
function main(workbook: ExcelScript.Workbook): string {
  const cacheSheet = workbook.getWorksheet(CACHE_SHEET);
  if (!cacheSheet || cacheSheet.getVisibility() !== ExcelScript.SheetVisibility.hidden) {
    throw new Error("PUL-030A-100: The hidden weekly-cache engineering surface is missing.");
  }
  const versionTable = requiredTable(workbook, "tblWeeklyCacheVersions");
  const periodTable = requiredTable(workbook, "tblWeeklyPeriodManifest");
  const scopeTable = requiredTable(workbook, "tblWeeklyScopeCache");
  const rpgTable = requiredTable(workbook, "tblWeeklyRPGCache");
  assertHeaders(versionTable, VERSION_HEADERS);
  assertHeaders(periodTable, PERIOD_HEADERS);
  assertHeaders(scopeTable, SCOPE_HEADERS);
  assertHeaders(rpgTable, RPG_HEADERS);

  const versionRows = tableRows(versionTable);
  const periodRows = tableRows(periodTable);
  const scopeRows = tableRows(scopeTable);
  const rpgRows = tableRows(rpgTable);
  const candidate = validateMaterializedCache(versionRows, periodRows, scopeRows, rpgRows);
  const live = validateLiveState(workbook);
  assertFreshContent(candidate, live);
  const protectedBefore = protectedSnapshot(workbook);

  const versionHeaders = headerMap(versionTable);
  let targetIndex = -1;
  let activeCount = 0;
  for (let index = 0; index < versionRows.length; index += 1) {
    const row = versionRows[index];
    const cacheVersion = text(row[versionHeaders.CacheVersion]);
    const cacheStatus = text(row[versionHeaders.CacheStatus]);
    const activationState = text(row[versionHeaders.ActivationState]);
    if (cacheStatus === "Active" || activationState === "Active") activeCount += 1;
    if (cacheVersion === EXPECTED_CACHE_VERSION) targetIndex = index;
  }
  if (targetIndex < 0) throw new Error("PUL-030A-101: Accepted cache version is missing.");
  const target = versionRows[targetIndex];
  const status = text(target[versionHeaders.CacheStatus]);
  const activation = text(target[versionHeaders.ActivationState]);
  if (status === "Active" && activation === "Active") {
    if (activeCount !== 1) throw new Error(`PUL-030A-102: Weekly cache has ${activeCount} Active authority rows.`);
    return JSON.stringify({
      status: "Already Active",
      cacheVersion: EXPECTED_CACHE_VERSION,
      cacheFingerprint: EXPECTED_CACHE_FINGERPRINT,
      freshness: "Available",
      live,
      candidate,
      protected: protectedBefore
    });
  }
  if (activeCount !== 0) throw new Error(`PUL-030A-103: Another weekly cache authority is Active (${activeCount}).`);
  if (status !== "Candidate" || activation !== "Not Active") {
    throw new Error(`PUL-030A-104: Accepted cache is ${status} / ${activation}; expected Candidate / Not Active.`);
  }
  if (versionHeaders.ActivationState !== versionHeaders.CacheStatus + 1) {
    throw new Error("PUL-030A-105: Cache authority columns are not contiguous.");
  }

  const authority = versionTable.getRangeBetweenHeaderAndTotal()
    .getCell(targetIndex, versionHeaders.CacheStatus).getResizedRange(0, 1);
  authority.setValues([["Active", "Active"]]);

  const afterRows = tableRows(versionTable);
  const afterTarget = afterRows[targetIndex];
  if (text(afterTarget[versionHeaders.CacheStatus]) !== "Active" ||
      text(afterTarget[versionHeaders.ActivationState]) !== "Active") {
    throw new Error("PUL-030A-106: Weekly cache authority update did not persist.");
  }
  const afterCandidate = validateMaterializedCache(afterRows, periodRows, scopeRows, rpgRows);
  const liveAfter = validateLiveState(workbook);
  assertFreshContent(afterCandidate, liveAfter);
  const protectedAfter = protectedSnapshot(workbook);
  if (JSON.stringify(protectedBefore) !== JSON.stringify(protectedAfter)) {
    throw new Error("PUL-030A-107: Protected Performance, Reports, metric-result, or import state changed.");
  }
  return JSON.stringify({
    status: "PASS",
    cacheVersion: EXPECTED_CACHE_VERSION,
    cacheFingerprint: EXPECTED_CACHE_FINGERPRINT,
    cacheStatus: "Active",
    activationState: "Active",
    freshness: "Available",
    live: liveAfter,
    candidate: afterCandidate,
    protected: protectedAfter
  });
}

type CellValue = string | number | boolean;
type Metric = { factCount: number; salesNok: number; quantity: number };
type LiveState = {
  mappingContentFingerprint: string;
  catalogContentFingerprint: string;
  performanceRestaurantScopeFingerprint: string;
  qaPassCount: number;
};

const CACHE_SHEET = "_Weekly_Cache";
const EXPECTED_CACHE_VERSION = "WCV-1a34ad1f46763d9b";
const EXPECTED_CACHE_FINGERPRINT = "WCC-508dd608166cdb6e";
const EXPECTED_MAPPING_CONTENT = "MCF-759cc92c4304a913";
const EXPECTED_CATALOG_CONTENT = "ICC-5644a77c18a97437";
const EXPECTED_IDENTITY_PREFLIGHT = "IDP-062c182f23905ae8";
const EXPECTED_RESTAURANT_SCOPE = "RSC-08df626f217dd94b";
const EXPECTED_FACTS = 245632;
const EXPECTED_SALES = 484728367.25;
const EXPECTED_QUANTITY = 2469988.09;
const EXPECTED_PENDING_FACTS = 120;
const EXPECTED_PENDING_SALES = 114876;
const EXPECTED_PENDING_QUANTITY = 951;

const VERSION_HEADERS = [
  "CacheVersion", "CacheSchemaVersion", "CacheStatus", "ActivationState",
  "ValidationStatus", "SourceSystemID", "ParserVersion", "IdentityContractVersion",
  "SourceCorpusFingerprint", "IdentityPreflightFingerprint", "CatalogFingerprint",
  "CatalogContentFingerprint", "MappingContentFingerprint", "Phase2AMappingFingerprint",
  "MappingAsOfDate", "ActiveReportingGroupFingerprint",
  "PerformanceRestaurantScopeFingerprint", "PeriodRowCount", "ScopeCacheRowCount",
  "DenseRPGCacheRowCount", "NonzeroRPGCacheRowCount", "CacheFingerprint"
];
const PERIOD_HEADERS = [
  "WeeklyPeriodManifestRowID", "CacheVersion", "SourcePeriodKey", "PeriodStart", "PeriodEnd",
  "ISOYear", "ISOWeek", "SourceFileID", "SourceSemanticFingerprint",
  "SourceBinaryFingerprint", "ScopeID", "ScopeFingerprint", "SourceFactCount",
  "SourceSalesNOK", "SourceQuantity", "SourceRestaurantCount"
];
const SCOPE_HEADERS = [
  "WeeklyScopeCacheRowID", "CacheVersion", "SourcePeriodKey", "RestaurantID",
  "PerformanceEligible", "SourceFactCount", "SourceSalesNOK", "SourceQuantity",
  "MappedFactCount", "MappedSalesNOK", "MappedQuantity", "UnmappedFactCount",
  "UnmappedSalesNOK", "UnmappedQuantity", "IdentityPendingFactCount",
  "IdentityPendingSalesNOK", "IdentityPendingQuantity", "ConflictFactCount",
  "ConflictSalesNOK", "ConflictQuantity", "InactiveTargetFactCount",
  "InactiveTargetSalesNOK", "InactiveTargetQuantity"
];
const RPG_HEADERS = [
  "WeeklyRPGCacheRowID", "CacheVersion", "SourcePeriodKey", "RestaurantID",
  "ReportingGroupID", "MappedFactCount", "MappedSalesNOK", "MappedQuantity"
];

function validateMaterializedCache(
  versionRows: CellValue[][],
  periodRows: CellValue[][],
  scopeRows: CellValue[][],
  rpgRows: CellValue[][]
): { cacheVersion: string; cacheFingerprint: string; completeRows: number; source: Metric; identityPending: Metric } {
  if (versionRows.length !== 1 || periodRows.length !== 84 || scopeRows.length !== 1421 || rpgRows.length !== 12789) {
    throw new Error(`PUL-030A-108: Cache row counts differ: ${versionRows.length}/${periodRows.length}/${scopeRows.length}/${rpgRows.length}.`);
  }
  const version = versionRows[0];
  const status = text(version[2]);
  const activation = text(version[3]);
  const validAuthority = (status === "Candidate" && activation === "Not Active") ||
    (status === "Active" && activation === "Active");
  if (text(version[0]) !== EXPECTED_CACHE_VERSION || !validAuthority || text(version[4]) !== "PASS" ||
      text(version[9]) !== EXPECTED_IDENTITY_PREFLIGHT ||
      text(version[11]) !== EXPECTED_CATALOG_CONTENT || text(version[12]) !== EXPECTED_MAPPING_CONTENT ||
      text(version[16]) !== EXPECTED_RESTAURANT_SCOPE || number(version[17]) !== 84 ||
      number(version[18]) !== 1421 || number(version[19]) !== 12789 ||
      text(version[21]) !== EXPECTED_CACHE_FINGERPRINT) {
    throw new Error("PUL-030A-109: Version manifest differs from the accepted cache contract.");
  }
  const source = emptyMetric();
  const mapped = emptyMetric();
  const identityPending = emptyMetric();
  const conflict = emptyMetric();
  const inactive = emptyMetric();
  const scopeGrains: { [key: string]: boolean } = {};
  for (const row of scopeRows) {
    if (text(row[1]) !== EXPECTED_CACHE_VERSION) throw new Error("PUL-030A-110: Scope row has another cache version.");
    const grain = `${text(row[2])}|${text(row[3])}`;
    if (!text(row[0]) || scopeGrains[grain]) throw new Error(`PUL-030A-111: Duplicate/blank scope grain ${grain}.`);
    scopeGrains[grain] = true;
    const rowSource = metric(row, 5);
    const coverage = emptyMetric();
    addMetric(source, rowSource);
    const rowMapped = metric(row, 8); addMetric(mapped, rowMapped); addMetric(coverage, rowMapped);
    addMetric(coverage, metric(row, 11));
    const pending = metric(row, 14); addMetric(identityPending, pending); addMetric(coverage, pending);
    const rowConflict = metric(row, 17); addMetric(conflict, rowConflict); addMetric(coverage, rowConflict);
    const rowInactive = metric(row, 20); addMetric(inactive, rowInactive); addMetric(coverage, rowInactive);
    assertMetricEqual(rowSource, coverage, `scope coverage ${grain}`);
    if ((text(row[3]) === "RST-0017" || text(row[3]) === "RST-0018") && text(row[4]) !== "No") {
      throw new Error(`PUL-030A-112: ${text(row[3])} must remain Performance-ineligible.`);
    }
  }
  const periodGrains: { [key: string]: boolean } = {};
  for (const row of periodRows) {
    if (text(row[1]) !== EXPECTED_CACHE_VERSION) throw new Error("PUL-030A-113: Period row has another cache version.");
    const key = text(row[2]);
    if (!text(row[0]) || periodGrains[key]) throw new Error(`PUL-030A-114: Duplicate/blank period ${key}.`);
    periodGrains[key] = true;
  }
  const rpgGrains: { [key: string]: boolean } = {};
  const mappedRpg = emptyMetric();
  for (const row of rpgRows) {
    if (text(row[1]) !== EXPECTED_CACHE_VERSION) throw new Error("PUL-030A-115: RPG row has another cache version.");
    const grain = `${text(row[2])}|${text(row[3])}|${text(row[4])}`;
    if (!text(row[0]) || rpgGrains[grain]) throw new Error(`PUL-030A-116: Duplicate/blank RPG grain ${grain}.`);
    rpgGrains[grain] = true;
    addMetric(mappedRpg, metric(row, 5));
  }
  assertMetricEqual(source, { factCount: EXPECTED_FACTS, salesNok: EXPECTED_SALES, quantity: EXPECTED_QUANTITY }, "corpus source");
  assertMetricEqual(identityPending, {
    factCount: EXPECTED_PENDING_FACTS, salesNok: EXPECTED_PENDING_SALES,
    quantity: EXPECTED_PENDING_QUANTITY
  }, "Identity Pending");
  assertMetricEqual(conflict, emptyMetric(), "Conflict");
  assertMetricEqual(inactive, emptyMetric(), "Inactive Target");
  assertMetricEqual(mappedRpg, mapped, "RPG numerators to Mapped scope");
  const cacheFingerprint = fingerprintCache(version, periodRows, scopeRows, rpgRows);
  if (cacheFingerprint !== EXPECTED_CACHE_FINGERPRINT) {
    throw new Error(`PUL-030A-117: CacheFingerprint ${cacheFingerprint} differs from accepted ${EXPECTED_CACHE_FINGERPRINT}.`);
  }
  return {
    cacheVersion: EXPECTED_CACHE_VERSION,
    cacheFingerprint,
    completeRows: versionRows.length + periodRows.length + scopeRows.length + rpgRows.length,
    source: finalMetric(source),
    identityPending: finalMetric(identityPending)
  };
}

function validateLiveState(workbook: ExcelScript.Workbook): LiveState {
  const groupsTable = requiredTable(workbook, "tblReportingGroups");
  const rulesTable = requiredTable(workbook, "tblMappingRules");
  const productsTable = requiredTable(workbook, "tblProducts");
  const classificationsTable = requiredTable(workbook, "tblSourceClassifications");
  const effectiveTable = requiredTable(workbook, "tblEffectiveMapping");
  const restaurantsTable = requiredTable(workbook, "tblRestaurants");
  const groups = tableRows(groupsTable);
  const rules = tableRows(rulesTable);
  const products = tableRows(productsTable);
  const classifications = tableRows(classificationsTable);
  const effective = tableRows(effectiveTable);
  const restaurants = tableRows(restaurantsTable);
  const mappingContentFingerprint = computeMappingContentFingerprint(
    groupsTable, groups, rulesTable, rules, productsTable, products,
    classificationsTable, classifications, effectiveTable, effective
  );
  return {
    mappingContentFingerprint,
    catalogContentFingerprint: computeCatalogContentFingerprint(
      mappingContentFingerprint, groupsTable, groups, rulesTable, rules,
      productsTable, products, classificationsTable, classifications,
      restaurantsTable, restaurants
    ),
    performanceRestaurantScopeFingerprint: restaurantScopeFingerprint(restaurantsTable, restaurants),
    qaPassCount: validateInteractionQa(workbook)
  };
}

function assertFreshContent(
  candidate: { cacheVersion: string; cacheFingerprint: string },
  live: LiveState
): void {
  if (candidate.cacheVersion !== EXPECTED_CACHE_VERSION ||
      candidate.cacheFingerprint !== EXPECTED_CACHE_FINGERPRINT ||
      live.mappingContentFingerprint !== EXPECTED_MAPPING_CONTENT ||
      live.catalogContentFingerprint !== EXPECTED_CATALOG_CONTENT ||
      live.performanceRestaurantScopeFingerprint !== EXPECTED_RESTAURANT_SCOPE ||
      live.qaPassCount !== 16) {
    throw new Error(`PUL-030A-118: Weekly cache freshness preflight failed. ${JSON.stringify(live)}`);
  }
}

function computeMappingContentFingerprint(
  groupsTable: ExcelScript.Table, groups: CellValue[][],
  rulesTable: ExcelScript.Table, rules: CellValue[][],
  productsTable: ExcelScript.Table, products: CellValue[][],
  classificationsTable: ExcelScript.Table, classifications: CellValue[][],
  effectiveTable: ExcelScript.Table, effective: CellValue[][]
): string {
  const gh = headerMap(groupsTable); const rh = headerMap(rulesTable);
  const ph = headerMap(productsTable); const ch = headerMap(classificationsTable);
  const eh = headerMap(effectiveTable);
  const classificationById: { [key: string]: { sourceSystemId: string; main: string } } = {};
  for (const row of classifications) classificationById[text(row[ch.SourceClassificationID])] = {
    sourceSystemId: text(row[ch.SourceSystemID]), main: text(row[ch.SourceMainCategory])
  };
  const records: string[] = [record("V", ["PULSE-MAPPING-CONTENT-V1"])];
  for (const row of groups) records.push(record("G", [
    row[gh.ReportingGroupID], row[gh.ReportingGroupName], row[gh.Active], row[gh.SortOrder]
  ]));
  for (const row of rules) {
    if (!text(row[rh.MappingRuleID])) continue;
    records.push(record("R", [
      row[rh.MappingRuleID], row[rh.SourceSystemID], row[rh.ScopeType], row[rh.NodeID],
      rh.RuleAction === undefined ? "Map" : text(row[rh.RuleAction]) || "Map",
      row[rh.TargetReportingGroupID], boundary(row[rh.EffectiveFrom]), boundary(row[rh.EffectiveTo]),
      row[rh.Status]
    ]));
  }
  for (const row of products) {
    const productId = text(row[ph.ProductID]);
    if (!productId) continue;
    const classificationId = text(row[ph.SourceClassificationID]);
    const classification = classificationById[classificationId];
    if (!classification) throw new Error(`PUL-030A-119: Product ${productId} has missing classification ${classificationId}.`);
    records.push(record("P", [
      productId, row[ph.SourceSystemID],
      `${classification.sourceSystemId} || Main || ${classification.main}`, classificationId
    ]));
  }
  const effectiveProducts: { [key: string]: boolean } = {};
  for (const row of effective) {
    const productId = text(row[eh.ProductID]);
    if (!productId || effectiveProducts[productId]) throw new Error(`PUL-030A-120: Effective Mapping repeats/omits ProductID ${productId}.`);
    effectiveProducts[productId] = true;
    records.push(record("E", [
      productId, row[eh.EffectiveReportingGroupID], row[eh.ResolutionSource],
      row[eh.ResolutionState], row[eh.ResolutionStatus], normalizeDelimited(row[eh.WinningRuleID])
    ]));
  }
  if (Object.keys(effectiveProducts).length !== products.filter(row => text(row[ph.ProductID])).length) {
    throw new Error("PUL-030A-121: Effective Mapping is not one row per current ProductID.");
  }
  records.sort();
  return hashStrings(records, "MCF-");
}

function computeCatalogContentFingerprint(
  mappingContentFingerprint: string,
  groupsTable: ExcelScript.Table, groups: CellValue[][],
  rulesTable: ExcelScript.Table, rules: CellValue[][],
  productsTable: ExcelScript.Table, products: CellValue[][],
  classificationsTable: ExcelScript.Table, classifications: CellValue[][],
  restaurantsTable: ExcelScript.Table, restaurants: CellValue[][]
): string {
  const gh = headerMap(groupsTable); const rh = headerMap(rulesTable);
  const ph = headerMap(productsTable); const ch = headerMap(classificationsTable);
  const th = headerMap(restaurantsTable);
  const records: string[] = [record("CATALOG_CONTENT", [mappingContentFingerprint])];
  for (const row of restaurants) records.push(record("RESTAURANT", [
    row[th.RestaurantID], row[th.SourceSystemID], row[th.SourceRestaurantName],
    row[th.Status], row[th.ReportingEnabled]
  ]));
  for (const row of products) {
    const sourceSystemId = text(row[ph.SourceSystemID]);
    records.push(record("PRODUCT", [
      row[ph.ProductID],
      `${sourceSystemId} || ${text(row[ph.SourceProductName])} || ${text(row[ph.SalesAccount])}`,
      row[ph.SourceClassificationID], row[ph.ProductStatus]
    ]));
  }
  for (const row of classifications) {
    const sourceSystemId = text(row[ch.SourceSystemID]);
    records.push(record("CLASSIFICATION", [
      row[ch.SourceClassificationID],
      `${sourceSystemId} || ${text(row[ch.SourceMainCategory])} || ${text(row[ch.SourceSubCategory])}`,
      row[ch.Status]
    ]));
  }
  for (const row of groups) records.push(record("REPORTING_GROUP", [
    row[gh.ReportingGroupID], row[gh.ReportingGroupName], row[gh.Active], row[gh.SortOrder]
  ]));
  for (const row of rules) records.push(record("MAPPING_RULE", [
    row[rh.MappingRuleID], row[rh.SourceSystemID], row[rh.ScopeType], row[rh.NodeID],
    row[rh.TargetReportingGroupID], boundary(row[rh.EffectiveFrom]), boundary(row[rh.EffectiveTo]),
    row[rh.Status], rh.RuleAction === undefined ? "Map" : text(row[rh.RuleAction]) || "Map"
  ]));
  records.sort();
  return hashStrings(records, "ICC-");
}

function restaurantScopeFingerprint(table: ExcelScript.Table, rows: CellValue[][]): string {
  const h = headerMap(table);
  const ids: string[] = [];
  for (const row of rows) {
    if (text(row[h.Status]) === "Active" && text(row[h.ReportingEnabled]) === "Yes") {
      ids.push(text(row[h.RestaurantID]));
    }
  }
  ids.sort();
  const serialized = ids.map(id => `${id.length}:${id}`).join("|");
  return `RSC-${hashText(`ENABLED-RESTAURANTS|${serialized}`)}`;
}

function fingerprintCache(
  version: CellValue[], periodRows: CellValue[][],
  scopeRows: CellValue[][], rpgRows: CellValue[][]
): string {
  const records: string[] = [record("CACHE_VERSION", [text(version[0])])];
  for (const row of periodRows) records.push(record("PERIOD", [
    row[2], canonicalIsoDate(row[3]), canonicalIsoDate(row[4]), row[7], row[8], row[12],
    sales(row[13]), quantity(row[14])
  ]));
  for (const row of scopeRows) records.push(record("SCOPE", [
    row[2], row[3], row[4],
    row[5], sales(row[6]), quantity(row[7]), row[8], sales(row[9]), quantity(row[10]),
    row[11], sales(row[12]), quantity(row[13]), row[14], sales(row[15]), quantity(row[16]),
    row[17], sales(row[18]), quantity(row[19]), row[20], sales(row[21]), quantity(row[22])
  ]));
  for (const row of rpgRows) records.push(record("RPG", [
    row[2], row[3], row[4], row[5], sales(row[6]), quantity(row[7])
  ]));
  records.sort();
  return hashStrings(records, "WCC-");
}

function protectedSnapshot(workbook: ExcelScript.Workbook): { [key: string]: string } {
  const performance = workbook.getWorksheet("Performance");
  const reports = workbook.getWorksheet("Reports");
  if (!performance || !reports) throw new Error("PUL-030A-122: Protected user sheets are missing.");
  return {
    performance: rangeFingerprint(performance.getRange("A1:T47")),
    reports: rangeFingerprint(reports.getRange("A1:H20")),
    metricResults: rangeFingerprint(requiredTable(workbook, "tblMetricRPGResults").getRange()),
    imports: rangeFingerprint(requiredTable(workbook, "tblImports").getRange()),
    phase2C: String(validateInteractionQa(workbook))
  };
}

function validateInteractionQa(workbook: ExcelScript.Workbook): number {
  const table = requiredTable(workbook, "tblPerformanceInteractionQA");
  const h = headerMap(table);
  const rows = tableRows(table);
  const seen: { [key: string]: boolean } = {};
  for (const row of rows) {
    const id = text(row[h.CheckID]);
    if (/^QA-0302C-(0[1-9]|1[0-6])$/.test(id) && text(row[h.Result]) === "PASS") seen[id] = true;
  }
  if (Object.keys(seen).length !== 16) {
    throw new Error(`PUL-030A-123: Phase 2C Interaction QA is ${Object.keys(seen).length}/16 PASS.`);
  }
  return 16;
}

function rangeFingerprint(range: ExcelScript.Range): string {
  const values = range.getValues();
  const formulas = range.getFormulas();
  return hashText(JSON.stringify([values, formulas]));
}

function assertHeaders(table: ExcelScript.Table, expected: string[]): void {
  const actual = table.getHeaderRowRange().getTexts()[0];
  if (actual.join("|") !== expected.join("|")) {
    throw new Error(`PUL-030A-124: ${table.getName()} headers differ from the accepted schema.`);
  }
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-030A-125: Required table ${name} is missing.`);
  return table;
}

function tableRows(table: ExcelScript.Table): CellValue[][] {
  return table.getRangeBetweenHeaderAndTotal().getValues();
}

function headerMap(table: ExcelScript.Table): { [key: string]: number } {
  const values = table.getHeaderRowRange().getTexts()[0];
  const result: { [key: string]: number } = {};
  for (let index = 0; index < values.length; index += 1) result[text(values[index])] = index;
  return result;
}

function metric(row: CellValue[], start: number): Metric {
  return { factCount: number(row[start]), salesNok: number(row[start + 1]), quantity: number(row[start + 2]) };
}

function emptyMetric(): Metric { return { factCount: 0, salesNok: 0, quantity: 0 }; }

function addMetric(target: Metric, value: Metric): void {
  target.factCount += value.factCount;
  target.salesNok += value.salesNok;
  target.quantity += value.quantity;
}

function finalMetric(value: Metric): Metric {
  return { factCount: number(value.factCount), salesNok: round(value.salesNok, 2), quantity: round(value.quantity, 6) };
}

function assertMetricEqual(actual: Metric, expected: Metric, label: string): void {
  const left = finalMetric(actual); const right = finalMetric(expected);
  if (left.factCount !== right.factCount || !close(left.salesNok, right.salesNok) ||
      !close(left.quantity, right.quantity)) {
    throw new Error(`PUL-030A-126: ${label} differs. Actual ${JSON.stringify(left)} expected ${JSON.stringify(right)}.`);
  }
}

function canonicalIsoDate(value: unknown): string {
  const raw = text(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const serial = number(value);
  if (!serial) return raw;
  const date = new Date(Math.round((serial - 25569) * 86400000));
  return date.toISOString().slice(0, 10);
}

function record(kind: string, values: unknown[]): string {
  return `${kind}|${values.map(value => {
    const normalized = value === null || value === undefined ? "" : String(value).trim();
    return `${normalized.length}:${normalized}`;
  }).join("|")}`;
}

function hashStrings(values: string[], prefix: string): string {
  let left = 0; let right = 0;
  for (const item of values) {
    const value = `${item}\n`;
    for (let index = 0; index < value.length; index += 1) {
      const code = value.charCodeAt(index);
      left = (left * 131 + code) % 2147483647;
      right = (right * 137 + code) % 2147483629;
    }
  }
  return `${prefix}${left.toString(16).padStart(8, "0")}${right.toString(16).padStart(8, "0")}`;
}

function hashText(value: string): string {
  let left = 0; let right = 0; const input = `${value}\n`;
  for (let index = 0; index < input.length; index += 1) {
    const code = input.charCodeAt(index);
    left = (left * 131 + code) % 2147483647;
    right = (right * 137 + code) % 2147483629;
  }
  return left.toString(16).padStart(8, "0") + right.toString(16).padStart(8, "0");
}

function boundary(value: CellValue): number { const parsed = Number(value); return value === "" || !Number.isFinite(parsed) ? 0 : parsed; }
function sales(value: CellValue): string { return number(value).toFixed(2); }
function quantity(value: CellValue): string { return number(value).toFixed(6); }
function round(value: number, decimals: number): number { const factor = Math.pow(10, decimals); return Math.round((value + Number.EPSILON) * factor) / factor; }
function number(value: unknown): number { const parsed = Number(value); return Number.isFinite(parsed) ? parsed : 0; }
function text(value: unknown): string { return String(value === null || value === undefined ? "" : value).trim(); }
function close(left: number, right: number): boolean { return Math.abs(left - right) <= 0.000001; }
function normalizeDelimited(value: unknown): string { return text(value).split(",").map(item => item.trim()).filter(item => item.length > 0).sort().join(", "); }
