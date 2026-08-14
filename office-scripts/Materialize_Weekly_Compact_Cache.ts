/**
 * Pulse Build 0.3.0 — Candidate Compact Weekly Cache materialization.
 *
 * This script writes one validated, inactive candidate to a hidden engineering
 * sheet. It never activates a cache, changes Performance/Reports, adds period
 * selectors, changes legacy imports, or reads raw weekly source workbooks.
 *
 * Run it in bounded steps with repository-generated payloads:
 *   Prepare -> Write (one or more chunks) -> Finalize.
 */
function main(
  workbook: ExcelScript.Workbook,
  action: string,
  payloadJson: string
): string {
  const requestedAction = text(action);
  const payload = parsePayload(payloadJson);
  if (requestedAction === "Prepare") return prepareMaterialization(workbook, payload);
  if (requestedAction === "Write") return writeMaterializationChunk(workbook, payload);
  if (requestedAction === "Finalize") return finalizeMaterialization(workbook, payload);
  throw new Error(`PUL-030M-100: Unsupported materialization action ${requestedAction || "(blank)"}.`);
}

type CellValue = string | number | boolean;
type Payload = {
  cacheVersion?: string;
  cacheFingerprint?: string;
  mappingContentFingerprint?: string;
  catalogContentFingerprint?: string;
  identityPreflightFingerprint?: string;
  phase2AMappingFingerprint?: string;
  mappingAsOfDate?: string | number;
  performanceRestaurantScopeFingerprint?: string;
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
type LiveState = {
  mappingContentFingerprint: string;
  catalogContentFingerprint: string;
  phase2AMappingFingerprint: string;
  mappingAsOfDate: number;
  performanceRestaurantScopeFingerprint: string;
  qaPassCount: number;
};

const FINAL_SHEET = "_Weekly_Cache";
const STAGING_SHEET = "_Weekly_Cache_Staging";
const EXPECTED_CACHE_VERSION = "WCV-1a34ad1f46763d9b";
const EXPECTED_CACHE_FINGERPRINT = "WCC-508dd608166cdb6e";
const EXPECTED_MAPPING_CONTENT = "MCF-759cc92c4304a913";
const EXPECTED_CATALOG_CONTENT = "ICC-5644a77c18a97437";
const EXPECTED_IDENTITY_PREFLIGHT = "IDP-062c182f23905ae8";
const EXPECTED_PHASE2A_MAPPING = "MAP-34202a7a1a922bd0";
const EXPECTED_MAPPING_AS_OF_DATE = "2026-08-12";
const EXPECTED_MAPPING_AS_OF_SERIAL = 46246;
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

function prepareMaterialization(workbook: ExcelScript.Workbook, payload: Payload): string {
  validatePayloadContract(payload);
  const live = validateLiveState(workbook);
  assertAcceptedLiveState(live, payload);
  const existing = workbook.getWorksheet(FINAL_SHEET);
  if (existing) {
    const result = validateFinalCandidate(existing);
    if (result.cacheVersion === EXPECTED_CACHE_VERSION &&
        result.cacheFingerprint === EXPECTED_CACHE_FINGERPRINT &&
        result.cacheStatus === "Candidate" && result.activationState === "Not Active") {
      return JSON.stringify({ status: "Already Materialized", live, candidate: result });
    }
    throw new Error("PUL-030M-101: Existing weekly cache surface differs from the accepted inactive candidate.");
  }
  const priorStaging = workbook.getWorksheet(STAGING_SHEET);
  if (priorStaging) priorStaging.delete();
  const staging = workbook.addWorksheet(STAGING_SHEET);
  staging.setVisibility(ExcelScript.SheetVisibility.hidden);
  for (const section of sections()) {
    checkedRange(staging, section.startRow, section.startColumn, 1, section.headers.length,
      `${section.name} header`).setValues([section.headers]);
  }
  return JSON.stringify({
    status: "Prepared",
    sheet: STAGING_SHEET,
    cacheVersion: EXPECTED_CACHE_VERSION,
    cacheFingerprint: EXPECTED_CACHE_FINGERPRINT,
    live
  });
}

function writeMaterializationChunk(workbook: ExcelScript.Workbook, payload: Payload): string {
  validatePayloadContract(payload);
  const staging = workbook.getWorksheet(STAGING_SHEET);
  if (!staging) throw new Error("PUL-030M-102: Weekly cache staging sheet is missing. Run Prepare first.");
  if (workbook.getWorksheet(FINAL_SHEET)) {
    throw new Error("PUL-030M-103: Final cache surface already exists; staging writes are not allowed.");
  }
  const section = requiredSection(payload.section);
  const values = payload.values || [];
  const offset = integer(payload.offset, "offset");
  const startRow = integer(payload.startRow, "startRow");
  const startColumn = integer(payload.startColumn, "startColumn");
  const totalRows = integer(payload.totalRowCount, "totalRowCount");
  if (totalRows !== section.rowCount || startRow !== section.startRow + 1 + offset ||
      startColumn !== section.startColumn || offset + values.length > section.rowCount) {
    throw new Error(`PUL-030M-104: ${section.name} chunk geometry differs from the accepted layout.`);
  }
  for (const row of values) {
    if (row.length !== section.headers.length) {
      throw new Error(`PUL-030M-105: ${section.name} chunk has an invalid column count.`);
    }
  }
  if (values.length) {
    checkedRange(staging, startRow, startColumn, values.length, section.headers.length,
      `${section.name} values`).setValues(values);
  }
  return JSON.stringify({
    status: "Written",
    section: section.name,
    offset,
    rows: values.length,
    totalRows: section.rowCount
  });
}

function finalizeMaterialization(workbook: ExcelScript.Workbook, payload: Payload): string {
  validatePayloadContract(payload);
  if (workbook.getWorksheet(FINAL_SHEET)) {
    throw new Error("PUL-030M-106: Final cache surface already exists. Run Prepare to verify idempotency.");
  }
  const staging = workbook.getWorksheet(STAGING_SHEET);
  if (!staging) throw new Error("PUL-030M-107: Weekly cache staging sheet is missing.");
  const live = validateLiveState(workbook);
  assertAcceptedLiveState(live, payload);
  const allSections = sections();
  const versionSection = allSections[0];
  const periodSection = allSections[1];
  const scopeSection = allSections[2];
  const rpgSection = allSections[3];
  assertHeaders(staging, versionSection);
  assertHeaders(staging, periodSection);
  assertHeaders(staging, scopeSection);
  assertHeaders(staging, rpgSection);
  const sectionValues: { [key: string]: CellValue[][] } = {
    version: checkedRange(staging, 2, 1, 1, VERSION_HEADERS.length,
      "version validation values").getValues(),
    period: checkedRange(staging, 6, 1, 84, PERIOD_HEADERS.length,
      "period validation values").getValues(),
    scope: checkedRange(staging, 93, 1, 1421, SCOPE_HEADERS.length,
      "scope validation values").getValues(),
    rpg: checkedRange(staging, 1517, 1, 12789, RPG_HEADERS.length,
      "RPG validation values").getValues()
  };
  const validation = validateStagedCandidate(sectionValues);
  if (validation.cacheFingerprint !== EXPECTED_CACHE_FINGERPRINT) {
    throw new Error(`PUL-030M-108: Staged CacheFingerprint ${validation.cacheFingerprint} differs from accepted ${EXPECTED_CACHE_FINGERPRINT}.`);
  }
  for (const section of sections()) {
    const range = checkedRange(staging, section.startRow, section.startColumn,
      section.rowCount + 1, section.headers.length, `${section.name} table`);
    const table = staging.addTable(range, true);
    table.setName(section.tableName);
    table.setPredefinedTableStyle("TableStyleMedium2");
  }
  formatEngineeringSurface(staging);
  staging.setName(FINAL_SHEET);
  staging.setVisibility(ExcelScript.SheetVisibility.hidden);
  const qaAfter = validateInteractionQa(workbook);
  if (qaAfter !== 16) throw new Error("PUL-030M-109: Phase 2C Interaction QA changed during materialization.");
  return JSON.stringify({
    status: "PASS",
    sheet: FINAL_SHEET,
    visibility: "Hidden",
    cacheStatus: "Candidate",
    activationState: "Not Active",
    live,
    validation
  });
}

function validateStagedCandidate(values: { [key: string]: CellValue[][] }): {
  cacheVersion: string; cacheFingerprint: string; completeRows: number;
  source: Metric; identityPending: Metric; conflict: Metric; inactiveTarget: Metric;
  rangeFixtures: { [key: string]: { [key: string]: number } };
} {
  const version = values.version[0];
  if (text(version[0]) !== EXPECTED_CACHE_VERSION || text(version[2]) !== "Candidate" ||
      text(version[3]) !== "Not Active" || text(version[4]) !== "PASS" ||
      text(version[9]) !== EXPECTED_IDENTITY_PREFLIGHT ||
      text(version[11]) !== EXPECTED_CATALOG_CONTENT ||
      text(version[12]) !== EXPECTED_MAPPING_CONTENT ||
      text(version[13]) !== EXPECTED_PHASE2A_MAPPING || !acceptedMappingDate(version[14]) ||
      text(version[16]) !== EXPECTED_RESTAURANT_SCOPE ||
      number(version[17]) !== 84 || number(version[18]) !== 1421 ||
      number(version[19]) !== 12789 || text(version[21]) !== EXPECTED_CACHE_FINGERPRINT) {
    throw new Error("PUL-030M-110: Version manifest differs from the accepted inactive candidate.");
  }
  const scopeRows = values.scope;
  const rpgRows = values.rpg;
  const periodRows = values.period;
  const source = emptyMetric();
  const states = {
    mapped: emptyMetric(), unmapped: emptyMetric(), identityPending: emptyMetric(),
    conflict: emptyMetric(), inactiveTarget: emptyMetric()
  };
  const scopeGrains: { [key: string]: boolean } = {};
  const periodSource: { [key: string]: Metric } = {};
  const periodCoverage: { [key: string]: Metric } = {};
  for (const row of scopeRows) {
    const grain = `${text(row[1])}|${text(row[2])}|${text(row[3])}`;
    if (!text(row[0]) || scopeGrains[grain]) throw new Error(`PUL-030M-111: Duplicate/blank scope grain ${grain}.`);
    scopeGrains[grain] = true;
    const rowSource = metric(row, 5);
    addMetric(source, rowSource);
    const coverage = emptyMetric();
    const mapped = metric(row, 8); addMetric(states.mapped, mapped); addMetric(coverage, mapped);
    const unmapped = metric(row, 11); addMetric(states.unmapped, unmapped); addMetric(coverage, unmapped);
    const pending = metric(row, 14); addMetric(states.identityPending, pending); addMetric(coverage, pending);
    const conflict = metric(row, 17); addMetric(states.conflict, conflict); addMetric(coverage, conflict);
    const inactive = metric(row, 20); addMetric(states.inactiveTarget, inactive); addMetric(coverage, inactive);
    assertMetricEqual(rowSource, coverage, `scope coverage ${grain}`);
    const periodKey = text(row[2]);
    if (!periodSource[periodKey]) periodSource[periodKey] = emptyMetric();
    if (!periodCoverage[periodKey]) periodCoverage[periodKey] = emptyMetric();
    addMetric(periodSource[periodKey], rowSource);
    addMetric(periodCoverage[periodKey], coverage);
    if ((text(row[3]) === "RST-0017" || text(row[3]) === "RST-0018") && text(row[4]) !== "No") {
      throw new Error(`PUL-030M-112: ${text(row[3])} must remain Performance-ineligible.`);
    }
  }
  assertMetricEqual(source, { factCount: EXPECTED_FACTS, salesNok: EXPECTED_SALES, quantity: EXPECTED_QUANTITY }, "corpus source");
  assertMetricEqual(states.identityPending, {
    factCount: EXPECTED_PENDING_FACTS, salesNok: EXPECTED_PENDING_SALES,
    quantity: EXPECTED_PENDING_QUANTITY
  }, "Identity Pending");
  assertMetricEqual(states.conflict, emptyMetric(), "Conflict");
  assertMetricEqual(states.inactiveTarget, emptyMetric(), "Inactive Target");
  const periodIds: { [key: string]: boolean } = {};
  for (const row of periodRows) {
    const key = text(row[2]);
    if (!text(row[0]) || periodIds[key]) throw new Error(`PUL-030M-113: Duplicate/blank period ${key}.`);
    periodIds[key] = true;
    const manifestMetric = metric(row, 12);
    assertMetricEqual(periodSource[key] || emptyMetric(), manifestMetric, `period source ${key}`);
    assertMetricEqual(periodCoverage[key] || emptyMetric(), manifestMetric, `period coverage ${key}`);
  }
  const rpgGrains: { [key: string]: boolean } = {};
  const mappedRpg = emptyMetric();
  for (const row of rpgRows) {
    const grain = `${text(row[1])}|${text(row[2])}|${text(row[3])}|${text(row[4])}`;
    if (!text(row[0]) || rpgGrains[grain]) throw new Error(`PUL-030M-114: Duplicate/blank RPG grain ${grain}.`);
    rpgGrains[grain] = true;
    addMetric(mappedRpg, metric(row, 5));
  }
  assertMetricEqual(mappedRpg, states.mapped, "RPG numerators to Mapped scope");
  const fingerprint = fingerprintCache(version, periodRows, scopeRows, rpgRows);
  const rangeFixtures = validateRangeFixtures(periodRows, scopeRows, rpgRows);
  return {
    cacheVersion: text(version[0]), cacheFingerprint: fingerprint,
    completeRows: 1 + periodRows.length + scopeRows.length + rpgRows.length,
    source: finalMetric(source), identityPending: finalMetric(states.identityPending),
    conflict: finalMetric(states.conflict), inactiveTarget: finalMetric(states.inactiveTarget),
    rangeFixtures
  };
}

function validateRangeFixtures(
  periodRows: CellValue[][],
  scopeRows: CellValue[][],
  rpgRows: CellValue[][]
): { [key: string]: { [key: string]: number } } {
  const fixtures = [
    { key: "W31", cy: 2026, cs: 31, ce: 31, py: 2025, ps: 31, pe: 31,
      share: 0.011882782054146212, compare: 0.006700865293813728,
      sales: 80263, impact: 35001.58321842167 },
    { key: "W01-W32", cy: 2026, cs: 1, ce: 32, py: 2025, ps: 1, pe: 32,
      share: 0.013073558139563134, compare: 0.007336150732445015,
      sales: 2383679, impact: 1046091.4622320954 },
    { key: "W20-W30", cy: 2026, cs: 20, ce: 30, py: 2025, ps: 20, pe: 30,
      share: 0.012494359243136773, compare: 0.006825152231488558,
      sales: 835122, impact: 378929.35570764536 }
  ];
  const periodIndex: { [key: string]: { year: number; week: number } } = {};
  for (const row of periodRows) periodIndex[text(row[2])] = { year: number(row[5]), week: number(row[6]) };
  const eligibleGrains: { [key: string]: boolean } = {};
  for (const row of scopeRows) {
    eligibleGrains[`${text(row[2])}|${text(row[3])}`] = text(row[4]) === "Yes";
  }
  const result: { [key: string]: { [key: string]: number } } = {};
  for (const fixture of fixtures) {
    const current = aggregateRange(periodIndex, eligibleGrains, scopeRows, rpgRows,
      fixture.cy, fixture.cs, fixture.ce, "RPG-0001");
    const comparison = aggregateRange(periodIndex, eligibleGrains, scopeRows, rpgRows,
      fixture.py, fixture.ps, fixture.pe, "RPG-0001");
    const currentShare = current.denominator === 0 ? 0 : current.numerator / current.denominator;
    const comparisonShare = comparison.denominator === 0 ? 0 : comparison.numerator / comparison.denominator;
    const impact = current.numerator - comparisonShare * current.denominator;
    assertClose(currentShare, fixture.share, `${fixture.key} current share`, 0.000000000001);
    assertClose(comparisonShare, fixture.compare, `${fixture.key} comparison share`, 0.000000000001);
    assertClose(current.numerator, fixture.sales, `${fixture.key} current Sales NOK`);
    assertClose(impact, fixture.impact, `${fixture.key} NOK Impact`);
    result[fixture.key] = {
      currentShare, comparisonShare, ppChange: (currentShare - comparisonShare) * 100,
      currentSalesNok: current.numerator, nokImpact: impact
    };
  }
  return result;
}

function aggregateRange(
  periodIndex: { [key: string]: { year: number; week: number } },
  eligibleGrains: { [key: string]: boolean },
  scopeRows: CellValue[][],
  rpgRows: CellValue[][],
  year: number,
  weekStart: number,
  weekEnd: number,
  reportingGroupId: string
): { denominator: number; numerator: number } {
  let denominator = 0;
  let numerator = 0;
  for (const row of scopeRows) {
    const period = periodIndex[text(row[2])];
    if (period && period.year === year && period.week >= weekStart && period.week <= weekEnd &&
        text(row[4]) === "Yes") denominator += number(row[6]);
  }
  for (const row of rpgRows) {
    const period = periodIndex[text(row[2])];
    if (period && period.year === year && period.week >= weekStart && period.week <= weekEnd &&
        text(row[4]) === reportingGroupId && eligibleGrains[`${text(row[2])}|${text(row[3])}`]) {
      numerator += number(row[6]);
    }
  }
  return { denominator: round(denominator, 2), numerator: round(numerator, 2) };
}

function fingerprintCache(
  version: CellValue[],
  periodRows: CellValue[][],
  scopeRows: CellValue[][],
  rpgRows: CellValue[][]
): string {
  const records: string[] = [record("CACHE_VERSION", [text(version[0])])];
  for (const row of periodRows) records.push(record("PERIOD", [
    row[2], canonicalIsoDate(row[3]), canonicalIsoDate(row[4]), row[7], row[8], row[12],
    sales(row[13]), quantity(row[14])
  ]));
  for (const row of scopeRows) records.push(record("SCOPE", [
    row[2], row[3], row[4],
    row[5], sales(row[6]), quantity(row[7]),
    row[8], sales(row[9]), quantity(row[10]),
    row[11], sales(row[12]), quantity(row[13]),
    row[14], sales(row[15]), quantity(row[16]),
    row[17], sales(row[18]), quantity(row[19]),
    row[20], sales(row[21]), quantity(row[22])
  ]));
  for (const row of rpgRows) records.push(record("RPG", [
    row[2], row[3], row[4], row[5], sales(row[6]), quantity(row[7])
  ]));
  records.sort();
  return hashStrings(records, "WCC-");
}

function validateLiveState(workbook: ExcelScript.Workbook): LiveState {
  const groupsTable = requiredTable(workbook, "tblReportingGroups");
  const rulesTable = requiredTable(workbook, "tblMappingRules");
  const productsTable = requiredTable(workbook, "tblProducts");
  const classificationsTable = requiredTable(workbook, "tblSourceClassifications");
  const effectiveTable = requiredTable(workbook, "tblEffectiveMapping");
  const restaurantsTable = requiredTable(workbook, "tblRestaurants");
  const bridgeTable = requiredTable(workbook, "tblMetricRPGFacts");
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
  const catalogContentFingerprint = computeCatalogContentFingerprint(
    mappingContentFingerprint, groupsTable, groups, rulesTable, rules,
    productsTable, products, classificationsTable, classifications,
    restaurantsTable, restaurants
  );
  const bridgeHeaders = headerMap(bridgeTable);
  const bridgeFingerprints = bridgeTable.getColumnByName("MappingFingerprint")
    .getRangeBetweenHeaderAndTotal().getValues();
  const bridgeDates = bridgeTable.getColumnByName("MappingAsOfDate")
    .getRangeBetweenHeaderAndTotal().getValues();
  let phase2AMappingFingerprint = "";
  let mappingAsOfDate = 0;
  for (let index = 0; index < bridgeFingerprints.length; index += 1) {
    const currentFingerprint = text(bridgeFingerprints[index][0]);
    const currentDate = number(bridgeDates[index][0]);
    if (!phase2AMappingFingerprint) phase2AMappingFingerprint = currentFingerprint;
    if (!mappingAsOfDate) mappingAsOfDate = currentDate;
    if (currentFingerprint !== phase2AMappingFingerprint || currentDate !== mappingAsOfDate) {
      throw new Error("PUL-030M-115: Phase 2A bridge contains mixed mapping audit state.");
    }
  }
  if (bridgeHeaders.MappingFingerprint === undefined || bridgeHeaders.MappingAsOfDate === undefined) {
    throw new Error("PUL-030M-116: Phase 2A bridge audit columns are missing.");
  }
  return {
    mappingContentFingerprint,
    catalogContentFingerprint,
    phase2AMappingFingerprint,
    mappingAsOfDate,
    performanceRestaurantScopeFingerprint: restaurantScopeFingerprint(restaurantsTable, restaurants),
    qaPassCount: validateInteractionQa(workbook)
  };
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
    if (!classification) throw new Error(`PUL-030M-117: Product ${productId} has missing classification ${classificationId}.`);
    const sourceSystemId = text(row[ph.SourceSystemID]);
    records.push(record("P", [
      productId, sourceSystemId, `${classification.sourceSystemId} || Main || ${classification.main}`,
      classificationId
    ]));
  }
  const effectiveProducts: { [key: string]: boolean } = {};
  for (const row of effective) {
    const productId = text(row[eh.ProductID]);
    if (!productId || effectiveProducts[productId]) throw new Error(`PUL-030M-118: Effective Mapping repeats/omits ProductID ${productId}.`);
    effectiveProducts[productId] = true;
    records.push(record("E", [
      productId, row[eh.EffectiveReportingGroupID], row[eh.ResolutionSource],
      row[eh.ResolutionState], row[eh.ResolutionStatus], normalizeDelimited(row[eh.WinningRuleID])
    ]));
  }
  if (Object.keys(effectiveProducts).length !== products.filter(row => text(row[ph.ProductID])).length) {
    throw new Error("PUL-030M-119: Effective Mapping is not one row per current ProductID.");
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
    const productKey = `${sourceSystemId} || ${text(row[ph.SourceProductName])} || ${text(row[ph.SalesAccount])}`;
    records.push(record("PRODUCT", [
      row[ph.ProductID], productKey, row[ph.SourceClassificationID], row[ph.ProductStatus]
    ]));
  }
  for (const row of classifications) {
    const sourceSystemId = text(row[ch.SourceSystemID]);
    const classificationKey = `${sourceSystemId} || ${text(row[ch.SourceMainCategory])} || ${text(row[ch.SourceSubCategory])}`;
    records.push(record("CLASSIFICATION", [
      row[ch.SourceClassificationID], classificationKey, row[ch.Status]
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

function validateInteractionQa(workbook: ExcelScript.Workbook): number {
  const table = requiredTable(workbook, "tblPerformanceInteractionQA");
  const h = headerMap(table);
  const rows = tableRows(table);
  let passCount = 0;
  const seen: { [key: string]: boolean } = {};
  for (const row of rows) {
    const id = text(row[h.CheckID]);
    if (/^QA-0302C-(0[1-9]|1[0-6])$/.test(id) && text(row[h.Result]) === "PASS") {
      seen[id] = true;
      passCount += 1;
    }
  }
  if (Object.keys(seen).length !== 16 || passCount !== 16) {
    throw new Error(`PUL-030M-120: Phase 2C Interaction QA is ${Object.keys(seen).length}/16 PASS.`);
  }
  return 16;
}

function assertAcceptedLiveState(live: LiveState, payload: Payload): void {
  if (live.mappingContentFingerprint !== EXPECTED_MAPPING_CONTENT ||
      live.catalogContentFingerprint !== EXPECTED_CATALOG_CONTENT ||
      live.phase2AMappingFingerprint !== EXPECTED_PHASE2A_MAPPING ||
      live.mappingAsOfDate !== EXPECTED_MAPPING_AS_OF_SERIAL ||
      live.performanceRestaurantScopeFingerprint !== EXPECTED_RESTAURANT_SCOPE ||
      live.qaPassCount !== 16 ||
      text(payload.mappingContentFingerprint) !== EXPECTED_MAPPING_CONTENT ||
      text(payload.catalogContentFingerprint) !== EXPECTED_CATALOG_CONTENT ||
      text(payload.identityPreflightFingerprint) !== EXPECTED_IDENTITY_PREFLIGHT ||
      text(payload.phase2AMappingFingerprint) !== EXPECTED_PHASE2A_MAPPING ||
      !acceptedMappingDate(payload.mappingAsOfDate) ||
      text(payload.performanceRestaurantScopeFingerprint) !== EXPECTED_RESTAURANT_SCOPE) {
    throw new Error(`PUL-030M-121: Live weekly-cache freshness preflight failed. ${JSON.stringify(live)}`);
  }
}

function validatePayloadContract(payload: Payload): void {
  if (text(payload.cacheVersion) !== EXPECTED_CACHE_VERSION ||
      text(payload.cacheFingerprint) !== EXPECTED_CACHE_FINGERPRINT) {
    throw new Error("PUL-030M-122: Payload does not identify the accepted weekly candidate.");
  }
}

function acceptedMappingDate(value: unknown): boolean {
  return text(value) === EXPECTED_MAPPING_AS_OF_DATE || number(value) === EXPECTED_MAPPING_AS_OF_SERIAL;
}

function canonicalIsoDate(value: unknown): string {
  const raw = text(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const serial = number(value);
  if (!serial) return raw;
  const date = new Date(Math.round((serial - 25569) * 86400000));
  return date.toISOString().slice(0, 10);
}

function validateFinalCandidate(sheet: ExcelScript.Worksheet): {
  cacheVersion: string; cacheFingerprint: string; cacheStatus: string; activationState: string;
} {
  const table = sheet.getTable("tblWeeklyCacheVersions");
  if (!table || table.getRowCount() !== 1) throw new Error("PUL-030M-123: Existing version manifest is invalid.");
  const h = headerMap(table);
  const row = table.getRangeBetweenHeaderAndTotal().getValues()[0];
  return {
    cacheVersion: text(row[h.CacheVersion]),
    cacheFingerprint: text(row[h.CacheFingerprint]),
    cacheStatus: text(row[h.CacheStatus]),
    activationState: text(row[h.ActivationState])
  };
}

function formatEngineeringSurface(sheet: ExcelScript.Worksheet): void {
  sheet.getFreezePanes().freezeRows(1);
  const used = sheet.getUsedRange();
  used.getFormat().getFont().setName("Aptos");
  used.getFormat().getFont().setSize(9);
  for (const section of sections()) {
    const header = checkedRange(sheet, section.startRow, section.startColumn, 1,
      section.headers.length, `${section.name} format header`);
    header.getFormat().getFill().setColor("#17233A");
    header.getFormat().getFont().setColor("#FFFFFF");
    header.getFormat().getFont().setBold(true);
    header.getFormat().setWrapText(true);
  }
  sheet.getRange("D6:E89").setNumberFormat("yyyy-mm-dd");
  sheet.getRange("N6:O89").setNumberFormat("#,##0.00");
  sheet.getRange("G93:W1513").setNumberFormat("#,##0.00");
  sheet.getRange("G1517:H14305").setNumberFormat("#,##0.00");
  sheet.getRange("A:W").getFormat().setColumnWidth(110);
}

function sections(): Section[] {
  return [
    { name: "version", tableName: "tblWeeklyCacheVersions", startRow: 1, startColumn: 1,
      headers: VERSION_HEADERS, rowCount: 1 },
    { name: "period", tableName: "tblWeeklyPeriodManifest", startRow: 5, startColumn: 1,
      headers: PERIOD_HEADERS, rowCount: 84 },
    { name: "scope", tableName: "tblWeeklyScopeCache", startRow: 92, startColumn: 1,
      headers: SCOPE_HEADERS, rowCount: 1421 },
    { name: "rpg", tableName: "tblWeeklyRPGCache", startRow: 1516, startColumn: 1,
      headers: RPG_HEADERS, rowCount: 12789 }
  ];
}

function requiredSection(name: string): Section {
  const candidate = text(name);
  for (const section of sections()) if (section.name === candidate) return section;
  throw new Error(`PUL-030M-124: Unknown materialization section ${candidate}.`);
}

function assertHeaders(sheet: ExcelScript.Worksheet, section: Section): void {
  const actual = checkedRange(sheet, section.startRow, section.startColumn, 1,
    section.headers.length, `${section.name} headers`).getTexts()[0];
  if (actual.join("|") !== section.headers.join("|")) {
    throw new Error(`PUL-030M-125: ${section.name} headers differ from the materialization contract.`);
  }
}

function checkedRange(
  sheet: ExcelScript.Worksheet,
  startRow: number,
  startColumn: number,
  rowCount: number,
  columnCount: number,
  label: string
): ExcelScript.Range {
  if (startRow < 1 || startColumn < 1 || rowCount < 1 || columnCount < 1 ||
      startRow + rowCount - 1 > 1048576 || startColumn + columnCount - 1 > 16384) {
    throw new Error(`PUL-030M-126: Invalid ${label} geometry ${startRow},${startColumn},${rowCount},${columnCount}.`);
  }
  return sheet.getRangeByIndexes(startRow - 1, startColumn - 1, rowCount, columnCount);
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-030M-127: Required table ${name} is missing.`);
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

function parsePayload(value: string): Payload {
  try {
    return JSON.parse(value || "{}") as Payload;
  } catch (error) {
    throw new Error(`PUL-030M-128: payloadJson is not valid JSON. ${String(error)}`);
  }
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
    throw new Error(`PUL-030M-129: ${label} differs. Actual ${JSON.stringify(left)} expected ${JSON.stringify(right)}.`);
  }
}

function assertClose(actual: number, expected: number, label: string, tolerance = 0.000001): void {
  if (Math.abs(actual - expected) > tolerance) {
    throw new Error(`PUL-030M-130: ${label} ${actual} differs from ${expected}.`);
  }
}

function close(left: number, right: number): boolean { return Math.abs(left - right) <= 0.000001; }
function boundary(value: CellValue): number { const parsed = Number(value); return value === "" || !Number.isFinite(parsed) ? 0 : parsed; }
function sales(value: CellValue): string { return number(value).toFixed(2); }
function quantity(value: CellValue): string { return number(value).toFixed(6); }
function round(value: number, decimals: number): number { const factor = Math.pow(10, decimals); return Math.round((value + Number.EPSILON) * factor) / factor; }
function number(value: unknown): number { const parsed = Number(value); return Number.isFinite(parsed) ? parsed : 0; }
function integer(value: unknown, label: string): number { const parsed = Number(value); if (!Number.isInteger(parsed) || parsed < 0) throw new Error(`PUL-030M-131: ${label} must be a non-negative integer.`); return parsed; }
function text(value: unknown): string { return String(value === null || value === undefined ? "" : value).trim(); }
function normalizeDelimited(value: unknown): string { return text(value).split(",").map(item => item.trim()).filter(item => item.length > 0).sort().join(", "); }

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
