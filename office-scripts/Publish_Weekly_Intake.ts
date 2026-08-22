/**
 * Pulse Build 0.3.0 — automated New-week publication runtime.
 *
 * Power Automate calls Build Candidate with the complete accepted parser
 * result, then calls Activate Candidate with the returned WCV/WCC. Business
 * calculations stay inside Pulse. The Candidate is fully materialized and
 * validated before activation; the prior Active cache is copied to a hidden
 * rollback surface before the final two-field authority switch.
 */
function main(
  workbook: ExcelScript.Workbook,
  operation: string = "Build Candidate",
  payloadJson: string = "{}"
): WeeklyPublicationResult {
  const payload = parseJson(payloadJson);
  if (operation === "Build Candidate") return buildCandidate(workbook, payload);
  if (operation === "Activate Candidate") return activateCandidate(workbook, payload);
  throw new Error(`PUL-030WPR-100: Unsupported publication operation ${operation}.`);
}

type CellValue = string | number | boolean;
type DataRow = { [key: string]: CellValue };
type Metric = { factCount: number; salesNok: number; quantity: number };
type ParsedRow = {
  sourceRowId: string; sourceFileId: string; sourcePeriodKey: string;
  restaurant: string; mainCategory: string; subCategory: string;
  salesAccount: string; item: string; quantity: number; salesNok: number;
};
type ParsedManifest = {
  sourceFileId: string; sourceLocator: string; sourceBinaryFingerprint: string;
  semanticFingerprint: string; sourcePeriodKey: string; periodStart: string;
  periodEnd: string; isoYear: number; isoWeek: number; parserVersion: string;
  schemaVersion: string; scopeId: string; scopeFingerprint: string;
  sourceRowCount: number; totalSalesNok: number; totalQuantity: number;
  restaurantCount: number; contentReconciliationStatus: string;
};
type ParsedReport = {
  parserVersion: string; schemaVersion: string;
  scopeContract: { sourceSystemId: string };
  manifest: ParsedManifest; rows: ParsedRow[];
};
type PublicationPayload = {
  parsedReport?: ParsedReport; candidateVersion?: string;
  candidateFingerprint?: string; sourcePeriodKey?: string;
};
type Assignment = {
  sourceRowId: string; restaurantId: string; productId: string;
  sourceClassificationId: string; identityState: string;
  identityPendingReason: string; mappingStatus: string;
  effectiveReportingGroupId: string;
};
type IdentityPlan = {
  assignments: Assignment[]; restaurants: DataRow[];
  restaurantCandidates: DataRow[]; classificationCandidates: DataRow[];
  productCandidates: DataRow[]; identityEvidenceFingerprint: string;
};
type IdentityRegistry = {
  control: DataRow; restaurants: DataRow[];
  classifications: DataRow[]; products: DataRow[];
};
type ActiveSnapshot = {
  versionTable: ExcelScript.Table; periodTable: ExcelScript.Table;
  scopeTable: ExcelScript.Table; rpgTable: ExcelScript.Table;
  versionRows: DataRow[]; activeVersion: DataRow;
  periodRows: DataRow[]; scopeRows: DataRow[]; rpgRows: DataRow[];
  identityRegistry: IdentityRegistry;
};
interface WeeklyPublicationResult {
  status: string; period: string; message: string; cacheChanged: boolean;
  priorCacheVersion: string; resultingCacheVersion: string;
  cacheFingerprint: string; archiveReady: boolean;
  periodRowCount: number; scopeRowCount: number; rpgRowCount: number;
  identityPendingFactCount: number; identityPendingSalesNok: number;
  identityPendingQuantity: number;
}

const PARSER_VERSION = "0.3.0-weekly-parser-v1";
const SCHEMA_VERSION = "sales-per-item-v1";
const CACHE_SCHEMA_VERSION = "0.3.0-weekly-compact-cache-v1";
const PUBLISHER_VERSION = "0.3.0-weekly-intake-publisher-v1";
const IDENTITY_EVIDENCE_VERSION = "0.3.0-weekly-intake-identity-evidence-v1";
const IDENTITY_REGISTRY_VERSION = "0.3.0-weekly-identity-registry-v1";
const CANDIDATE_SHEET = "_Weekly_Cache_Candidate";
const ROLLBACK_SHEET = "_Weekly_Cache_Rollback";
const CACHE_SHEET = "_Weekly_Cache";
const IDENTITY_SHEET = "_Weekly_Identity";
const CANDIDATE_TABLES = [
  "tblWeeklyIntakeCandidateVersions", "tblWeeklyIntakeCandidatePeriodManifest",
  "tblWeeklyIntakeCandidateScopeCache", "tblWeeklyIntakeCandidateRPGCache"
];
const ROLLBACK_TABLES = [
  "tblWeeklyRollbackVersions", "tblWeeklyRollbackPeriodManifest",
  "tblWeeklyRollbackScopeCache", "tblWeeklyRollbackRPGCache"
];
const IDENTITY_TABLES = [
  "tblWeeklyIdentityRegistryControl", "tblWeeklyIdentityRestaurants",
  "tblWeeklyIdentityClassifications", "tblWeeklyIdentityProducts"
];
const CANDIDATE_IDENTITY_TABLES = [
  "tblWeeklyIntakeCandidateIdentityControl", "tblWeeklyIntakeCandidateIdentityRestaurants",
  "tblWeeklyIntakeCandidateIdentityClassifications", "tblWeeklyIntakeCandidateIdentityProducts"
];
const ROLLBACK_IDENTITY_TABLES = [
  "tblWeeklyRollbackIdentityControl", "tblWeeklyRollbackIdentityRestaurants",
  "tblWeeklyRollbackIdentityClassifications", "tblWeeklyRollbackIdentityProducts"
];
const STATES = ["Mapped", "Unmapped", "Identity Pending", "Conflict", "Inactive Target"];
const STATE_PREFIXES = ["mapped", "unmapped", "identityPending", "conflict", "inactiveTarget"];

const VERSION_COLUMNS = [
  ["CacheVersion", "cacheVersion"], ["CacheSchemaVersion", "cacheSchemaVersion"],
  ["CacheStatus", "cacheStatus"], ["ActivationState", "activationState"],
  ["ValidationStatus", "validationStatus"], ["SourceSystemID", "sourceSystemId"],
  ["ParserVersion", "parserVersion"], ["IdentityContractVersion", "identityContractVersion"],
  ["SourceCorpusFingerprint", "sourceCorpusFingerprint"],
  ["IdentityPreflightFingerprint", "identityPreflightFingerprint"],
  ["CatalogFingerprint", "catalogFingerprint"],
  ["CatalogContentFingerprint", "catalogContentFingerprint"],
  ["MappingContentFingerprint", "mappingContentFingerprint"],
  ["Phase2AMappingFingerprint", "mappingFingerprint"],
  ["MappingAsOfDate", "mappingAsOfDate"],
  ["ActiveReportingGroupFingerprint", "activeReportingGroupFingerprint"],
  ["PerformanceRestaurantScopeFingerprint", "performanceRestaurantScopeFingerprint"],
  ["PeriodRowCount", "periodRowCount"], ["ScopeCacheRowCount", "scopeCacheRowCount"],
  ["DenseRPGCacheRowCount", "denseRpgCacheRowCount"],
  ["NonzeroRPGCacheRowCount", "nonzeroRpgCacheRowCount"],
  ["CacheFingerprint", "cacheFingerprint"]
];
const PERIOD_COLUMNS = [
  ["WeeklyPeriodManifestRowID", "weeklyPeriodManifestRowId"], ["CacheVersion", "cacheVersion"],
  ["SourcePeriodKey", "sourcePeriodKey"], ["PeriodStart", "periodStart"],
  ["PeriodEnd", "periodEnd"], ["ISOYear", "isoYear"], ["ISOWeek", "isoWeek"],
  ["SourceFileID", "sourceFileId"], ["SourceSemanticFingerprint", "sourceSemanticFingerprint"],
  ["SourceBinaryFingerprint", "sourceBinaryFingerprint"], ["ScopeID", "scopeId"],
  ["ScopeFingerprint", "scopeFingerprint"], ["SourceFactCount", "sourceFactCount"],
  ["SourceSalesNOK", "sourceSalesNok"], ["SourceQuantity", "sourceQuantity"],
  ["SourceRestaurantCount", "sourceRestaurantCount"]
];
const SCOPE_COLUMNS = [
  ["WeeklyScopeCacheRowID", "weeklyScopeCacheRowId"], ["CacheVersion", "cacheVersion"],
  ["SourcePeriodKey", "sourcePeriodKey"], ["RestaurantID", "restaurantId"],
  ["PerformanceEligible", "performanceEligible"], ["SourceFactCount", "sourceFactCount"],
  ["SourceSalesNOK", "sourceSalesNok"], ["SourceQuantity", "sourceQuantity"],
  ["MappedFactCount", "mappedFactCount"], ["MappedSalesNOK", "mappedSalesNok"],
  ["MappedQuantity", "mappedQuantity"], ["UnmappedFactCount", "unmappedFactCount"],
  ["UnmappedSalesNOK", "unmappedSalesNok"], ["UnmappedQuantity", "unmappedQuantity"],
  ["IdentityPendingFactCount", "identityPendingFactCount"],
  ["IdentityPendingSalesNOK", "identityPendingSalesNok"],
  ["IdentityPendingQuantity", "identityPendingQuantity"],
  ["ConflictFactCount", "conflictFactCount"], ["ConflictSalesNOK", "conflictSalesNok"],
  ["ConflictQuantity", "conflictQuantity"],
  ["InactiveTargetFactCount", "inactiveTargetFactCount"],
  ["InactiveTargetSalesNOK", "inactiveTargetSalesNok"],
  ["InactiveTargetQuantity", "inactiveTargetQuantity"]
];
const RPG_COLUMNS = [
  ["WeeklyRPGCacheRowID", "weeklyRpgCacheRowId"], ["CacheVersion", "cacheVersion"],
  ["SourcePeriodKey", "sourcePeriodKey"], ["RestaurantID", "restaurantId"],
  ["ReportingGroupID", "reportingGroupId"], ["MappedFactCount", "mappedFactCount"],
  ["MappedSalesNOK", "mappedSalesNok"], ["MappedQuantity", "mappedQuantity"]
];
const PRODUCT_COLUMNS = [
  ["ProductID", "productId"], ["SourceSystemID", "sourceSystemId"],
  ["SourceProductName", "sourceProductName"], ["SalesAccount", "salesAccount"],
  ["SourceClassificationID", "sourceClassificationId"], ["ProductKey", "productKey"],
  ["ProductStatus", "productStatus"]
];
const CLASS_COLUMNS = [
  ["SourceClassificationID", "sourceClassificationId"], ["SourceSystemID", "sourceSystemId"],
  ["SourceMainCategory", "sourceMainCategory"], ["SourceSubCategory", "sourceSubCategory"],
  ["SourceClassificationKey", "sourceClassificationKey"], ["Status", "status"]
];
const RESTAURANT_COLUMNS = [
  ["RestaurantID", "restaurantId"], ["SourceSystemID", "sourceSystemId"],
  ["SourceRestaurantName", "sourceRestaurantName"], ["DisplayName", "displayName"],
  ["Status", "status"], ["ReportingEnabled", "reportingEnabled"]
];
const GROUP_COLUMNS = [
  ["ReportingGroupID", "reportingGroupId"], ["ReportingGroupName", "reportingGroupName"],
  ["Active", "active"], ["SortOrder", "sortOrder"]
];
const RULE_COLUMNS = [
  ["MappingRuleID", "mappingRuleId"], ["SourceSystemID", "sourceSystemId"],
  ["ScopeType", "scopeType"], ["NodeID", "nodeId"],
  ["TargetReportingGroupID", "targetReportingGroupId"], ["EffectiveFrom", "effectiveFrom"],
  ["EffectiveTo", "effectiveTo"], ["Status", "status"], ["RuleAction", "ruleAction"]
];
const EFFECTIVE_COLUMNS = [
  ["ProductID", "productId"], ["EffectiveReportingGroupID", "effectiveReportingGroupId"],
  ["ResolutionStatus", "resolutionStatus"], ["WinningRuleID", "winningRuleId"]
];
const IDENTITY_CONTROL_COLUMNS = [
  ["IdentityRegistryVersion", "identityRegistryVersion"],
  ["IdentityPreflightFingerprint", "identityPreflightFingerprint"],
  ["IdentityRegistryFingerprint", "identityRegistryFingerprint"],
  ["RestaurantCount", "restaurantCount"],
  ["ClassificationCount", "classificationCount"],
  ["ProductCount", "productCount"]
];
const IDENTITY_PRODUCT_COLUMNS = [
  ["ProductID", "productId"], ["SourceSystemID", "sourceSystemId"],
  ["SourceProductName", "sourceProductName"], ["SalesAccount", "salesAccount"],
  ["SourceClassificationID", "sourceClassificationId"], ["ProductKey", "productKey"],
  ["ProductStatus", "productStatus"], ["HierarchyStatus", "hierarchyStatus"],
  ["ObservedHierarchyPaths", "observedHierarchyPaths"]
];

function buildCandidate(workbook: ExcelScript.Workbook, payload: PublicationPayload): WeeklyPublicationResult {
  const report = payload.parsedReport ? payload.parsedReport : payload as unknown as ParsedReport;
  validateParsedReport(report);
  const active = activeSnapshot(workbook);
  const existingPeriod = rowsForPeriod(active.periodRows, report.manifest.sourcePeriodKey);
  if (existingPeriod.length) {
    if (existingPeriod.length === 1 && text(existingPeriod[0].sourceSemanticFingerprint) === report.manifest.semanticFingerprint) {
      return publicationResult("Duplicate", report.manifest.sourcePeriodKey,
        "Same period and semantic fingerprint; no cache change.", false,
        text(active.activeVersion.cacheVersion), "", text(active.activeVersion.cacheFingerprint), true, active, zeroMetric());
    }
    throw new Error(`PUL-030WPR-101: Period ${report.manifest.sourcePeriodKey} already exists or is ambiguous.`);
  }
  const candidate = buildCandidateRows(workbook, active, report);
  materializeCandidate(workbook, candidate);
  const newLedger = intakeLedgerEntry(report, "New",
    `Validated Candidate ${text(candidate.version.cacheVersion)} prepared; not active.`,
    text(active.activeVersion.cacheVersion), text(candidate.version.cacheVersion), "");
  upsertLedger(workbook, newLedger);
  const pending = stateTotals(candidate.scopeRows, "identityPending");
  return publicationResult("Candidate Ready", report.manifest.sourcePeriodKey,
    `Candidate ${text(candidate.version.cacheVersion)} is PASS / Not Active.`, false,
    text(active.activeVersion.cacheVersion), text(candidate.version.cacheVersion),
    text(candidate.version.cacheFingerprint), false, candidate, pending);
}

function activateCandidate(workbook: ExcelScript.Workbook, payload: PublicationPayload): WeeklyPublicationResult {
  const expectedVersion = text(payload.candidateVersion);
  const expectedFingerprint = text(payload.candidateFingerprint);
  const active = activeSnapshot(workbook);
  if (expectedVersion && text(active.activeVersion.cacheVersion) === expectedVersion) {
    if (expectedFingerprint && text(active.activeVersion.cacheFingerprint) !== expectedFingerprint) {
      throw new Error(`PUL-030WPR-103: Active fingerprint ${text(active.activeVersion.cacheFingerprint)} differs from ${expectedFingerprint}.`);
    }
    const period = text(payload.sourcePeriodKey) || newestPeriod(active.periodRows);
    const periodRows = rowsForPeriod(active.periodRows, period);
    if (periodRows.length !== 1) throw new Error(`PUL-030WPR-145: Active cache does not contain exactly one ${period}.`);
    const rollback = active.versionRows.filter(row => text(row.cacheStatus) === "Rollback" &&
      text(row.activationState) === "Not Active");
    const priorVersion = rollback.length === 1 ? text(rollback[0].cacheVersion) : "";
    const published = publicationLedgerEntry(workbook, periodRows[0], active.activeVersion, priorVersion);
    upsertLedger(workbook, published);
    return publicationResult("Published", period,
      `Published ${period} in ${expectedVersion}; activation was already complete.`, true,
      priorVersion, expectedVersion, text(active.activeVersion.cacheFingerprint), true,
      active, stateTotals(active.scopeRows, "identityPending"));
  }
  const candidate = readCandidate(workbook);
  if (expectedVersion && expectedVersion !== text(candidate.version.cacheVersion)) {
    throw new Error(`PUL-030WPR-102: Candidate version ${text(candidate.version.cacheVersion)} differs from ${expectedVersion}.`);
  }
  if (expectedFingerprint && expectedFingerprint !== text(candidate.version.cacheFingerprint)) {
    throw new Error(`PUL-030WPR-103: Candidate fingerprint ${text(candidate.version.cacheFingerprint)} differs from ${expectedFingerprint}.`);
  }
  validateCandidateForActivation(workbook, active, candidate, text(payload.sourcePeriodKey));
  const priorVersion = text(active.activeVersion.cacheVersion);
  const resultVersion = text(candidate.version.cacheVersion);
  const priorIdentity = text(active.activeVersion.identityPreflightFingerprint);
  backupActiveCache(workbook, active);
  let canonicalMutationStarted = false;
  try {
    const versionRows: DataRow[] = [];
    const rollbackVersion = cloneRow(active.activeVersion);
    rollbackVersion.cacheStatus = "Rollback";
    rollbackVersion.activationState = "Not Active";
    const candidateVersion = cloneRow(candidate.version);
    candidateVersion.cacheStatus = "Candidate";
    candidateVersion.activationState = "Not Active";
    versionRows.push(rollbackVersion); versionRows.push(candidateVersion);
    canonicalMutationStarted = true;
    writeIdentityRegistry(workbook, candidate.identityRegistry);
    writeCanonicalCache(active, versionRows, candidate.periodRows, candidate.scopeRows, candidate.rpgRows);
    requiredSheet(workbook, "_Metric_Calc").getRange("AL30")
      .setValue(text(candidate.version.identityPreflightFingerprint));
    const statusRange = active.versionTable.getRangeBetweenHeaderAndTotal().getCell(0, 2)
      .getResizedRange(1, 1);
    statusRange.setValues([["Rollback", "Not Active"], ["Active", "Active"]]);
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);
    validatePublishedPostconditions(workbook, resultVersion, text(candidate.version.cacheFingerprint),
      text(payload.sourcePeriodKey));
    const period = text(payload.sourcePeriodKey) || newestPeriod(candidate.periodRows);
    const source = candidate.periodRows.filter(row => text(row.sourcePeriodKey) === period)[0];
    const published = publicationLedgerEntry(workbook, source, candidate.version, priorVersion);
    upsertLedger(workbook, published);
    const candidateSheet = workbook.getWorksheet(CANDIDATE_SHEET);
    if (candidateSheet) {
      try { candidateSheet.delete(); } catch (ignored) { /* A published Candidate is harmless and cleaned on the next build. */ }
    }
    const pending = stateTotals(candidate.scopeRows, "identityPending");
    return publicationResult("Published", period,
      `Published ${period} in ${resultVersion}; prior ${priorVersion} retained for rollback.`, true,
      priorVersion, resultVersion, text(candidate.version.cacheFingerprint), true, candidate, pending);
  } catch (error) {
    if (canonicalMutationStarted) {
      writeIdentityRegistry(workbook, active.identityRegistry);
      restoreCanonicalCache(active);
      requiredSheet(workbook, "_Metric_Calc").getRange("AL30").setValue(priorIdentity);
      workbook.getApplication().calculate(ExcelScript.CalculationType.full);
    }
    throw new Error(`PUL-030WPR-104: Publication failed; prior Active cache restored. ${String(error)}`);
  }
}

function buildCandidateRows(workbook: ExcelScript.Workbook, active: ActiveSnapshot, report: ParsedReport): {
  version: DataRow; periodRows: DataRow[]; scopeRows: DataRow[]; rpgRows: DataRow[];
  activeReportingGroups: DataRow[]; identityRegistry: IdentityRegistry;
} {
  const products = readObjects(requiredTable(workbook, "tblProducts"), PRODUCT_COLUMNS);
  const classifications = readObjects(requiredTable(workbook, "tblSourceClassifications"), CLASS_COLUMNS);
  const restaurants = readObjects(requiredTable(workbook, "tblRestaurants"), RESTAURANT_COLUMNS);
  const groups = readObjects(requiredTable(workbook, "tblReportingGroups"), GROUP_COLUMNS);
  const rules = readObjects(requiredTable(workbook, "tblMappingRules"), RULE_COLUMNS);
  const effective = readObjects(requiredTable(workbook, "tblEffectiveMapping"), EFFECTIVE_COLUMNS);
  validateIdentityRegistryAgainstCatalogs(active.identityRegistry, products,
    classifications, restaurants);
  const sourceSystemId = text(active.activeVersion.sourceSystemId);
  if (report.scopeContract.sourceSystemId !== sourceSystemId) {
    throw new Error(`PUL-030WPR-105: Parsed source system differs from ${sourceSystemId}.`);
  }
  const identityProducts = products.concat(active.identityRegistry.products);
  const identityClassifications = classifications.concat(active.identityRegistry.classifications);
  const identityRestaurants = restaurants.concat(active.identityRegistry.restaurants);
  const identity = buildIdentityPlan(report, sourceSystemId, identityProducts,
    identityClassifications, identityRestaurants, groups, rules, effective,
    number(active.activeVersion.mappingAsOfDate));
  const activeGroups = validatedActiveReportingGroups(groups);
  const incoming = buildIncomingWeek(report, identity, activeGroups);
  const periodBase: DataRow[] = [];
  for (const row of active.periodRows) periodBase.push(cloneRow(row));
  periodBase.push(incoming.periodRow);
  periodBase.sort(comparePeriod);
  const corpusFingerprint = fingerprintCorpus(periodBase);
  const evidence: DataRow[] = [
    { evidenceKey: `BASE:${text(active.activeVersion.cacheVersion)}`,
      fingerprint: text(active.activeVersion.identityPreflightFingerprint) },
    { evidenceKey: report.manifest.sourcePeriodKey, fingerprint: identity.identityEvidenceFingerprint }
  ];
  evidence.sort((left, right) => compareText(left.evidenceKey, right.evidenceKey));
  const identityRecords = [
    record("CONTRACT", [PUBLISHER_VERSION]), record("CORPUS", [corpusFingerprint])
  ];
  for (const row of evidence) identityRecords.push(record("IDENTITY_EVIDENCE", [row.evidenceKey, row.fingerprint]));
  const identityFingerprint = hashStrings(identityRecords, "IDP-");
  const groupIds = activeGroups.map(row => text(row.reportingGroupId)).sort(compareTextValue);
  const versionRecords = [
    record("CONTRACT", [PUBLISHER_VERSION]), record("SCHEMA", [CACHE_SCHEMA_VERSION]),
    record("CORPUS", [corpusFingerprint]), record("IDENTITY", [identityFingerprint]),
    record("CATALOG", [active.activeVersion.catalogContentFingerprint]),
    record("MAPPING_CONTENT", [active.activeVersion.mappingContentFingerprint]),
    record("RESTAURANT_SCOPE", [active.activeVersion.performanceRestaurantScopeFingerprint])
  ];
  for (const id of groupIds) versionRecords.push(record("RPG", [id]));
  const cacheVersion = hashStrings(versionRecords, "WCV-");
  if (cacheVersion === text(active.activeVersion.cacheVersion)) {
    throw new Error("PUL-030WPR-107: Candidate CacheVersion equals the Active version.");
  }
  const periodRows: DataRow[] = [];
  for (const row of periodBase) periodRows.push(rekeyPeriod(row, cacheVersion));
  const scopeRows: DataRow[] = [];
  for (const row of active.scopeRows) scopeRows.push(rekeyScope(row, cacheVersion));
  for (const row of incoming.scopeRows) scopeRows.push(rekeyScope(row, cacheVersion));
  scopeRows.sort(compareScope);
  const rpgRows: DataRow[] = [];
  for (const row of active.rpgRows) rpgRows.push(rekeyRpg(row, cacheVersion));
  for (const row of incoming.rpgRows) rpgRows.push(rekeyRpg(row, cacheVersion));
  rpgRows.sort(compareRpg);
  validateCompleteRows(periodRows, scopeRows, rpgRows, activeGroups);
  const cacheFingerprint = fingerprintCache(cacheVersion, periodRows, scopeRows, rpgRows);
  const version = cloneRow(active.activeVersion);
  version.cacheVersion = cacheVersion;
  version.cacheStatus = "Candidate";
  version.activationState = "Not Active";
  version.validationStatus = "PASS";
  version.sourceCorpusFingerprint = corpusFingerprint;
  version.identityPreflightFingerprint = identityFingerprint;
  version.periodRowCount = periodRows.length;
  version.scopeCacheRowCount = scopeRows.length;
  version.denseRpgCacheRowCount = rpgRows.length;
  version.nonzeroRpgCacheRowCount = rpgRows.filter(isNonzeroRpg).length;
  version.cacheFingerprint = cacheFingerprint;
  const identityRegistry = mergeIdentityRegistry(active.identityRegistry, identity, identityFingerprint);
  return { version, periodRows, scopeRows, rpgRows, activeReportingGroups: activeGroups,
    identityRegistry };
}

function buildIdentityPlan(
  report: ParsedReport, sourceSystemId: string, products: DataRow[], classifications: DataRow[],
  restaurants: DataRow[], groups: DataRow[], rules: DataRow[], effective: DataRow[], asOf: number
): IdentityPlan {
  const restaurantIndex = indexRows(restaurants, row => restaurantKey(sourceSystemId, text(row.sourceRestaurantName)));
  const classIndex = indexRows(classifications, row => text(row.sourceClassificationKey));
  const productIndex = indexRows(products, row => text(row.productKey));
  const effectiveIndex = indexRows(effective, row => text(row.productId));
  const observedRestaurants: { [key: string]: string } = {};
  const observedClasses: { [key: string]: DataRow } = {};
  const observedProducts: { [key: string]: DataRow } = {};
  const productPaths: { [key: string]: string[] } = {};
  for (const row of report.rows) {
    const rKey = restaurantKey(sourceSystemId, row.restaurant);
    const cKey = classificationKey(sourceSystemId, row.mainCategory, row.subCategory);
    const pKey = productKey(sourceSystemId, row.item, row.salesAccount);
    observedRestaurants[keyToken(rKey)] = row.restaurant;
    observedClasses[keyToken(cKey)] = {
      sourceSystemId, sourceMainCategory: row.mainCategory, sourceSubCategory: row.subCategory,
      sourceClassificationKey: cKey
    };
    observedProducts[keyToken(pKey)] = {
      sourceSystemId, sourceProductName: row.item, salesAccount: row.salesAccount, productKey: pKey
    };
    const token = keyToken(pKey);
    const paths = productPaths[token] || [];
    if (paths.indexOf(cKey) < 0) paths.push(cKey);
    productPaths[token] = paths;
  }
  const restaurantNewKeys = unmatchedKeys(observedRestaurants, restaurantIndex);
  const classNewKeys = unmatchedObjectKeys(observedClasses, classIndex, "sourceClassificationKey");
  const productNewKeys = unmatchedObjectKeys(observedProducts, productIndex, "productKey");
  const restaurantIds = allocateIds(restaurants, "restaurantId", "RST", 4, restaurantNewKeys);
  const classIds = allocateIds(classifications, "sourceClassificationId", "SCL", 5, classNewKeys);
  const productIds = allocateIds(products, "productId", "PRD", 6, productNewKeys);
  const restaurantCandidates: DataRow[] = [];
  for (const key of restaurantNewKeys) restaurantCandidates.push({
    restaurantId: restaurantIds[keyToken(key)], sourceSystemId,
    sourceRestaurantName: observedRestaurants[keyToken(key)], displayName: observedRestaurants[keyToken(key)],
    status: "Active", reportingEnabled: "No"
  });
  const classificationCandidates: DataRow[] = [];
  for (const key of classNewKeys) {
    const value = observedClasses[keyToken(key)];
    const row = cloneRow(value); row.sourceClassificationId = classIds[keyToken(key)]; row.status = "Active";
    classificationCandidates.push(row);
  }
  const allClasses: DataRow[] = [];
  for (const row of classifications) allClasses.push(row);
  for (const row of classificationCandidates) allClasses.push(row);
  const allClassIndex = indexRows(allClasses, row => text(row.sourceClassificationKey));
  const productCandidates: DataRow[] = [];
  for (const key of productNewKeys) {
    const value = observedProducts[keyToken(key)];
    const paths = productPaths[keyToken(key)].slice().sort(compareTextValue);
    const classMatch = paths.length === 1 ? singleRow(allClassIndex, paths[0]) : null;
    productCandidates.push({
      productId: productIds[keyToken(key)], sourceSystemId,
      sourceProductName: value.sourceProductName, salesAccount: value.salesAccount,
      sourceClassificationId: classMatch ? classMatch.sourceClassificationId : "",
      productKey: key, productStatus: "Active",
      hierarchyStatus: classMatch ? "Proposed exact hierarchy" : "Identity Pending",
      observedHierarchyPaths: paths.join("\u001f")
    });
  }
  const allRestaurants: DataRow[] = [];
  for (const row of restaurants) allRestaurants.push(row);
  for (const row of restaurantCandidates) allRestaurants.push(row);
  const allProducts: DataRow[] = [];
  for (const row of products) allProducts.push(row);
  for (const row of productCandidates) allProducts.push(row);
  const allRestaurantIndex = indexRows(allRestaurants, row => restaurantKey(sourceSystemId, text(row.sourceRestaurantName)));
  const allProductIndex = indexRows(allProducts, row => text(row.productKey));
  const allClassByKey = indexRows(allClasses, row => text(row.sourceClassificationKey));
  const classById = indexRows(allClasses, row => text(row.sourceClassificationId));
  const candidateProductIds: { [key: string]: boolean } = {};
  for (const row of productCandidates) candidateProductIds[keyToken(text(row.productId))] = true;
  const mappingByProduct: { [key: string]: DataRow } = {};
  for (const product of allProducts) {
    const productId = text(product.productId);
    const resolved = singleRow(effectiveIndex, productId);
    if (resolved) {
      mappingByProduct[keyToken(productId)] = resolved ? {
        resolutionStatus: resolved.resolutionStatus,
        effectiveReportingGroupId: resolved.effectiveReportingGroupId,
        winningRuleId: resolved.winningRuleId,
        identityPendingReason: ""
      } : {};
    } else if (candidateProductIds[keyToken(productId)] || text(product.hierarchyStatus)) {
      const classification = singleRow(classById, text(product.sourceClassificationId));
      mappingByProduct[keyToken(productId)] = classification ?
        resolveNewProduct(product, classification, rules, groups, asOf) :
        { resolutionStatus: "Identity Pending", effectiveReportingGroupId: "", winningRuleId: "",
          identityPendingReason: "Product hierarchy identity is pending." };
    } else {
      mappingByProduct[keyToken(productId)] = { resolutionStatus: "Identity Pending",
        effectiveReportingGroupId: "", winningRuleId: "",
        identityPendingReason: "Product Effective Mapping is missing or ambiguous." };
    }
  }
  const assignments: Assignment[] = [];
  for (const row of report.rows) {
    const r = singleRow(allRestaurantIndex, restaurantKey(sourceSystemId, row.restaurant));
    const p = singleRow(allProductIndex, productKey(sourceSystemId, row.item, row.salesAccount));
    const c = singleRow(allClassByKey, classificationKey(sourceSystemId, row.mainCategory, row.subCategory));
    const mapping = p ? mappingByProduct[keyToken(text(p.productId))] : null;
    const reasons: string[] = [];
    if (!r) reasons.push("Restaurant identity collision");
    if (!p) reasons.push("Product identity collision");
    if (!c) reasons.push("Source classification identity collision");
    if (mapping && text(mapping.identityPendingReason)) reasons.push(text(mapping.identityPendingReason));
    const pending = reasons.length > 0;
    assignments.push({
      sourceRowId: row.sourceRowId, restaurantId: r ? text(r.restaurantId) : "",
      productId: p ? text(p.productId) : "", sourceClassificationId: c ? text(c.sourceClassificationId) : "",
      identityState: pending ? "Identity Pending" : "Stable",
      identityPendingReason: reasons.join("; "),
      mappingStatus: pending ? "Identity Pending" : text(mapping ? mapping.resolutionStatus : "Identity Pending"),
      effectiveReportingGroupId: !pending && mapping && text(mapping.resolutionStatus) === "Mapped"
        ? text(mapping.effectiveReportingGroupId) : ""
    });
  }
  const evidenceRecords = [record("CONTRACT", [IDENTITY_EVIDENCE_VERSION])];
  for (const row of assignments) evidenceRecords.push(record("ASSIGNMENT", [
    row.sourceRowId, row.restaurantId, row.productId, row.sourceClassificationId,
    row.identityState, row.identityPendingReason, row.mappingStatus, row.effectiveReportingGroupId
  ]));
  for (const row of restaurantCandidates) evidenceRecords.push(record("RESTAURANT_CANDIDATE", [
    row.restaurantId, row.sourceSystemId, row.sourceRestaurantName, row.status, row.reportingEnabled
  ]));
  for (const row of classificationCandidates) evidenceRecords.push(record("CLASSIFICATION_CANDIDATE", [
    row.sourceClassificationId, row.sourceSystemId, row.sourceClassificationKey, row.status
  ]));
  for (const row of productCandidates) {
    const values: CellValue[] = [row.productId, row.sourceSystemId, row.productKey,
      row.sourceClassificationId, row.productStatus, row.hierarchyStatus];
    const paths = text(row.observedHierarchyPaths).split("\u001f").filter(value => value !== "");
    for (const path of paths) values.push(path);
    evidenceRecords.push(record("PRODUCT_CANDIDATE", values));
  }
  evidenceRecords.sort(compareTextValue);
  return {
    assignments, restaurants: allRestaurants, restaurantCandidates,
    classificationCandidates, productCandidates,
    identityEvidenceFingerprint: hashStrings(evidenceRecords, "WIE-")
  };
}

function resolveNewProduct(product: DataRow, classification: DataRow, rules: DataRow[],
  groups: DataRow[], asOf: number): DataRow {
  const groupIndex = indexRows(groups, row => text(row.reportingGroupId));
  const levels = [
    ["Product", text(product.productId)],
    ["SourceSubCategory", text(classification.sourceClassificationId)],
    ["SourceMainCategory", `${text(product.sourceSystemId)} || Main || ${text(classification.sourceMainCategory)}`]
  ];
  for (const level of levels) {
    const matches: DataRow[] = [];
    for (const rule of rules) {
      if (text(rule.status) !== "Active" || text(rule.sourceSystemId) !== text(product.sourceSystemId) ||
          text(rule.scopeType) !== level[0] || text(rule.nodeId) !== level[1]) continue;
      const start = nullableNumber(rule.effectiveFrom, Number.NEGATIVE_INFINITY);
      const end = nullableNumber(rule.effectiveTo, Number.POSITIVE_INFINITY);
      if (start <= asOf && end >= asOf) matches.push(rule);
    }
    if (matches.length > 1) return { resolutionStatus: "Conflict", effectiveReportingGroupId: "",
      winningRuleId: matches.map(row => text(row.mappingRuleId)).join(", "), identityPendingReason: "" };
    if (matches.length === 1) {
      const rule = matches[0]; const action = text(rule.ruleAction) || "Map";
      if (action === "Exclude") return { resolutionStatus: "Unmapped", effectiveReportingGroupId: "",
        winningRuleId: rule.mappingRuleId, identityPendingReason: "" };
      const group = singleRow(groupIndex, text(rule.targetReportingGroupId));
      return { resolutionStatus: group && text(group.active) === "Yes" ? "Mapped" : "Inactive Target",
        effectiveReportingGroupId: rule.targetReportingGroupId,
        winningRuleId: rule.mappingRuleId, identityPendingReason: "" };
    }
  }
  return { resolutionStatus: "Unmapped", effectiveReportingGroupId: "", winningRuleId: "",
    identityPendingReason: "" };
}

function buildIncomingWeek(report: ParsedReport, identity: IdentityPlan, groups: DataRow[]): {
  periodRow: DataRow; scopeRows: DataRow[]; rpgRows: DataRow[];
} {
  const assignmentIndex: { [key: string]: Assignment } = {};
  for (const row of identity.assignments) assignmentIndex[keyToken(row.sourceRowId)] = row;
  const restaurantIndex = indexRows(identity.restaurants, row => text(row.restaurantId));
  const groupIndex = indexRows(groups, row => text(row.reportingGroupId));
  const scopes: { [key: string]: DataRow } = {};
  const rpgs: { [key: string]: Metric } = {};
  for (const row of report.rows) {
    const assignment = assignmentIndex[keyToken(row.sourceRowId)];
    if (!assignment || !assignment.restaurantId) throw new Error(`PUL-030WPR-108: Row ${row.sourceRowId} lacks stable identity.`);
    const restaurant = singleRow(restaurantIndex, assignment.restaurantId);
    if (!restaurant) throw new Error(`PUL-030WPR-109: Restaurant ${assignment.restaurantId} is unavailable.`);
    const key = keyToken(`${row.sourcePeriodKey}|${assignment.restaurantId}`);
    let scope = scopes[key];
    if (!scope) {
      scope = { sourcePeriodKey: row.sourcePeriodKey, restaurantId: assignment.restaurantId,
        performanceEligible: text(restaurant.status) === "Active" && text(restaurant.reportingEnabled) === "Yes" ? "Yes" : "No" };
      setMetric(scope, "source", zeroMetric());
      for (const prefix of STATE_PREFIXES) setMetric(scope, prefix, zeroMetric());
      scopes[key] = scope;
    }
    addMetricToRow(scope, "source", row);
    const stateIndex = STATES.indexOf(assignment.mappingStatus);
    if (stateIndex < 0) throw new Error(`PUL-030WPR-110: Unsupported mapping state ${assignment.mappingStatus}.`);
    addMetricToRow(scope, STATE_PREFIXES[stateIndex], row);
    if (assignment.mappingStatus === "Mapped") {
      if (!singleRow(groupIndex, assignment.effectiveReportingGroupId)) {
        throw new Error(`PUL-030WPR-111: Mapped row targets unavailable ${assignment.effectiveReportingGroupId}.`);
      }
      const rpgKey = keyToken(`${row.sourcePeriodKey}|${assignment.restaurantId}|${assignment.effectiveReportingGroupId}`);
      const metric = rpgs[rpgKey] || zeroMetric(); addMetric(metric, row); rpgs[rpgKey] = metric;
    }
  }
  const scopeRows: DataRow[] = [];
  for (const key in scopes) scopeRows.push(finalizeMetricRow(scopes[key]));
  scopeRows.sort(compareScope);
  const rpgRows: DataRow[] = [];
  for (const scope of scopeRows) for (const group of groups) {
    const key = keyToken(`${text(scope.sourcePeriodKey)}|${text(scope.restaurantId)}|${text(group.reportingGroupId)}`);
    const metric = finalizeMetric(rpgs[key] || zeroMetric());
    rpgRows.push({ sourcePeriodKey: scope.sourcePeriodKey, restaurantId: scope.restaurantId,
      reportingGroupId: group.reportingGroupId, mappedFactCount: metric.factCount,
      mappedSalesNok: metric.salesNok, mappedQuantity: metric.quantity });
  }
  return {
    periodRow: {
      sourcePeriodKey: report.manifest.sourcePeriodKey, periodStart: excelDateSerial(report.manifest.periodStart),
      periodEnd: excelDateSerial(report.manifest.periodEnd), isoYear: report.manifest.isoYear, isoWeek: report.manifest.isoWeek,
      sourceFileId: report.manifest.sourceFileId,
      sourceSemanticFingerprint: report.manifest.semanticFingerprint,
      sourceBinaryFingerprint: report.manifest.sourceBinaryFingerprint,
      scopeId: report.manifest.scopeId, scopeFingerprint: report.manifest.scopeFingerprint,
      sourceFactCount: report.manifest.sourceRowCount, sourceSalesNok: report.manifest.totalSalesNok,
      sourceQuantity: report.manifest.totalQuantity, sourceRestaurantCount: report.manifest.restaurantCount
    }, scopeRows, rpgRows
  };
}

function activeSnapshot(workbook: ExcelScript.Workbook): ActiveSnapshot {
  const versionTable = requiredTable(workbook, "tblWeeklyCacheVersions");
  const periodTable = requiredTable(workbook, "tblWeeklyPeriodManifest");
  const scopeTable = requiredTable(workbook, "tblWeeklyScopeCache");
  const rpgTable = requiredTable(workbook, "tblWeeklyRPGCache");
  const versionRows = readObjects(versionTable, VERSION_COLUMNS);
  const activeRows = versionRows.filter(row => text(row.cacheStatus) === "Active" && text(row.activationState) === "Active");
  if (activeRows.length !== 1) throw new Error(`PUL-030WPR-112: Expected one Active / Active cache; found ${activeRows.length}.`);
  const activeVersion = activeRows[0];
  const periodRows = readObjects(periodTable, PERIOD_COLUMNS).filter(row => text(row.cacheVersion) === text(activeVersion.cacheVersion));
  const scopeRows = readObjects(scopeTable, SCOPE_COLUMNS).filter(row => text(row.cacheVersion) === text(activeVersion.cacheVersion));
  const rpgRows = readObjects(rpgTable, RPG_COLUMNS).filter(row => text(row.cacheVersion) === text(activeVersion.cacheVersion));
  if (periodRows.length !== number(activeVersion.periodRowCount) ||
      scopeRows.length !== number(activeVersion.scopeCacheRowCount) ||
      rpgRows.length !== number(activeVersion.denseRpgCacheRowCount)) {
    throw new Error("PUL-030WPR-113: Active cache row counts differ from manifest.");
  }
  const calc = requiredSheet(workbook, "_Metric_Calc");
  if (calc.getRange("AL16").getText() !== "Available" || calc.getRange("AL24").getText() !== text(activeVersion.cacheVersion) ||
      calc.getRange("AL27").getText() !== text(activeVersion.mappingContentFingerprint) ||
      calc.getRange("AL28").getText() !== text(activeVersion.catalogContentFingerprint) ||
      calc.getRange("AL29").getText() !== text(activeVersion.performanceRestaurantScopeFingerprint) ||
      calc.getRange("AL30").getText() !== text(activeVersion.identityPreflightFingerprint) ||
      qaPassCount(workbook, "tblWeeklyPerformanceQA") !== 16 ||
      qaPassCount(workbook, "tblPerformanceInteractionQA") !== 16) {
    throw new Error("PUL-030WPR-114: Active cache freshness or QA is unavailable.");
  }
  validateCompleteRows(periodRows, scopeRows, rpgRows,
    readObjects(requiredTable(workbook, "tblReportingGroups"), GROUP_COLUMNS).filter(row => text(row.active) === "Yes"));
  if (fingerprintCache(text(activeVersion.cacheVersion), periodRows, scopeRows, rpgRows) !== text(activeVersion.cacheFingerprint)) {
    throw new Error("PUL-030WPR-115: Active cache fingerprint differs from materialized rows.");
  }
  const identityRegistry = readIdentityRegistry(workbook, IDENTITY_SHEET, IDENTITY_TABLES,
    text(activeVersion.identityPreflightFingerprint));
  return { versionTable, periodTable, scopeTable, rpgTable, versionRows,
    activeVersion, periodRows, scopeRows, rpgRows, identityRegistry };
}

function materializeCandidate(workbook: ExcelScript.Workbook, candidate: {
  version: DataRow; periodRows: DataRow[]; scopeRows: DataRow[]; rpgRows: DataRow[];
  identityRegistry: IdentityRegistry;
}): void {
  const existing = workbook.getWorksheet(CANDIDATE_SHEET);
  if (existing) {
    const versionTable = workbook.getTable(CANDIDATE_TABLES[0]);
    if (versionTable) {
      const rows = readObjects(versionTable, VERSION_COLUMNS);
      const activeRows = readObjects(requiredTable(workbook, "tblWeeklyCacheVersions"), VERSION_COLUMNS)
        .filter(row => text(row.cacheStatus) === "Active" && text(row.activationState) === "Active");
      if (rows.length === 1 && activeRows.length === 1 &&
          text(rows[0].cacheVersion) === text(activeRows[0].cacheVersion)) {
        existing.delete();
        materializeCandidate(workbook, candidate);
        return;
      }
      if (rows.length === 1 && text(rows[0].cacheVersion) === text(candidate.version.cacheVersion) &&
          text(rows[0].cacheFingerprint) === text(candidate.version.cacheFingerprint)) {
        readIdentityRegistry(workbook, CANDIDATE_SHEET, CANDIDATE_IDENTITY_TABLES,
          text(candidate.version.identityPreflightFingerprint));
        return;
      }
    }
    throw new Error("PUL-030WPR-116: A different or partial Candidate surface already exists.");
  }
  const sheet = workbook.addWorksheet(CANDIDATE_SHEET);
  try {
    sheet.setVisibility(ExcelScript.SheetVisibility.hidden);
    const starts = layout(candidate.periodRows.length, candidate.scopeRows.length);
    createSection(sheet, starts[0], VERSION_COLUMNS, [candidate.version], CANDIDATE_TABLES[0]);
    createSection(sheet, starts[1], PERIOD_COLUMNS, candidate.periodRows, CANDIDATE_TABLES[1]);
    createSection(sheet, starts[2], SCOPE_COLUMNS, candidate.scopeRows, CANDIDATE_TABLES[2]);
    createSection(sheet, starts[3], RPG_COLUMNS, candidate.rpgRows, CANDIDATE_TABLES[3]);
    createIdentitySections(sheet, candidate.identityRegistry, CANDIDATE_IDENTITY_TABLES, 26);
    sheet.getRange("Y1:Z7").setValues([
      ["Weekly intake Candidate", "Value"], ["CandidateVersion", candidate.version.cacheVersion],
      ["CandidateFingerprint", candidate.version.cacheFingerprint], ["Status", "Candidate"],
      ["Authority", "Not Active"], ["Validation", "PASS"], ["CompleteRows", 1 + candidate.periodRows.length + candidate.scopeRows.length + candidate.rpgRows.length]
    ]);
  } catch (error) {
    try { sheet.delete(); } catch (ignored) { /* next rerun fails clearly if cleanup is unavailable */ }
    throw error;
  }
}

function readCandidate(workbook: ExcelScript.Workbook): {
  version: DataRow; periodRows: DataRow[]; scopeRows: DataRow[]; rpgRows: DataRow[];
  identityRegistry: IdentityRegistry;
} {
  const sheet = requiredSheet(workbook, CANDIDATE_SHEET);
  if (sheet.getVisibility() !== ExcelScript.SheetVisibility.hidden) throw new Error("PUL-030WPR-117: Candidate sheet must remain hidden.");
  const versions = readObjects(requiredTable(workbook, CANDIDATE_TABLES[0]), VERSION_COLUMNS);
  if (versions.length !== 1) throw new Error("PUL-030WPR-118: Candidate version row is missing or duplicated.");
  return { version: versions[0],
    periodRows: readObjects(requiredTable(workbook, CANDIDATE_TABLES[1]), PERIOD_COLUMNS),
    scopeRows: readObjects(requiredTable(workbook, CANDIDATE_TABLES[2]), SCOPE_COLUMNS),
    rpgRows: readObjects(requiredTable(workbook, CANDIDATE_TABLES[3]), RPG_COLUMNS),
    identityRegistry: readIdentityRegistry(workbook, CANDIDATE_SHEET,
      CANDIDATE_IDENTITY_TABLES, text(versions[0].identityPreflightFingerprint)) };
}

function validateCandidateForActivation(workbook: ExcelScript.Workbook, active: ActiveSnapshot,
  candidate: { version: DataRow; periodRows: DataRow[]; scopeRows: DataRow[]; rpgRows: DataRow[];
    identityRegistry: IdentityRegistry },
  requestedPeriod: string): void {
  if (text(candidate.version.cacheStatus) !== "Candidate" || text(candidate.version.activationState) !== "Not Active" ||
      text(candidate.version.validationStatus) !== "PASS") throw new Error("PUL-030WPR-119: Candidate is not PASS / Not Active.");
  if (text(candidate.version.mappingContentFingerprint) !== text(active.activeVersion.mappingContentFingerprint) ||
      text(candidate.version.catalogContentFingerprint) !== text(active.activeVersion.catalogContentFingerprint) ||
      text(candidate.version.performanceRestaurantScopeFingerprint) !== text(active.activeVersion.performanceRestaurantScopeFingerprint)) {
    throw new Error("PUL-030WPR-120: Candidate freshness differs from the current Active cache.");
  }
  if (candidate.periodRows.length !== number(candidate.version.periodRowCount) ||
      candidate.scopeRows.length !== number(candidate.version.scopeCacheRowCount) ||
      candidate.rpgRows.length !== number(candidate.version.denseRpgCacheRowCount)) {
    throw new Error("PUL-030WPR-121: Candidate counts differ from manifest.");
  }
  validateCompleteRows(candidate.periodRows, candidate.scopeRows, candidate.rpgRows,
    readObjects(requiredTable(workbook, "tblReportingGroups"), GROUP_COLUMNS).filter(row => text(row.active) === "Yes"));
  const fingerprint = fingerprintCache(text(candidate.version.cacheVersion), candidate.periodRows,
    candidate.scopeRows, candidate.rpgRows);
  if (fingerprint !== text(candidate.version.cacheFingerprint)) throw new Error("PUL-030WPR-122: Candidate WCC differs.");
  if (candidate.periodRows.length !== active.periodRows.length + 1) throw new Error("PUL-030WPR-123: Candidate must add exactly one period.");
  validateIdentityRegistryCarryForward(active.identityRegistry, candidate.identityRegistry);
  for (const period of active.periodRows) {
    const matches = candidate.periodRows.filter(row => text(row.sourcePeriodKey) === text(period.sourcePeriodKey) &&
      text(row.sourceSemanticFingerprint) === text(period.sourceSemanticFingerprint));
    if (matches.length !== 1) throw new Error(`PUL-030WPR-124: Candidate does not preserve ${text(period.sourcePeriodKey)}.`);
  }
  if (requestedPeriod && candidate.periodRows.filter(row => text(row.sourcePeriodKey) === requestedPeriod).length !== 1) {
    throw new Error(`PUL-030WPR-125: Candidate does not contain requested ${requestedPeriod}.`);
  }
}

function backupActiveCache(workbook: ExcelScript.Workbook, active: ActiveSnapshot): void {
  const existing = workbook.getWorksheet(ROLLBACK_SHEET);
  if (existing) existing.delete();
  const sheet = workbook.addWorksheet(ROLLBACK_SHEET);
  sheet.setVisibility(ExcelScript.SheetVisibility.hidden);
  const rollbackVersion = cloneRow(active.activeVersion);
  rollbackVersion.cacheStatus = "Rollback"; rollbackVersion.activationState = "Not Active";
  const starts = layout(active.periodRows.length, active.scopeRows.length);
  createSection(sheet, starts[0], VERSION_COLUMNS, [rollbackVersion], ROLLBACK_TABLES[0]);
  createSection(sheet, starts[1], PERIOD_COLUMNS, active.periodRows, ROLLBACK_TABLES[1]);
  createSection(sheet, starts[2], SCOPE_COLUMNS, active.scopeRows, ROLLBACK_TABLES[2]);
  createSection(sheet, starts[3], RPG_COLUMNS, active.rpgRows, ROLLBACK_TABLES[3]);
  createIdentitySections(sheet, active.identityRegistry, ROLLBACK_IDENTITY_TABLES, 26);
}

function writeCanonicalCache(active: ActiveSnapshot, versions: DataRow[], periods: DataRow[],
  scopes: DataRow[], rpgs: DataRow[]): void {
  ensureTableSpace(active.periodTable, active.scopeTable, periods.length);
  ensureTableSpace(active.scopeTable, active.rpgTable, scopes.length);
  resizeAndWrite(active.rpgTable, RPG_COLUMNS, rpgs);
  resizeAndWrite(active.scopeTable, SCOPE_COLUMNS, scopes);
  resizeAndWrite(active.periodTable, PERIOD_COLUMNS, periods);
  resizeAndWrite(active.versionTable, VERSION_COLUMNS, versions);
}

function restoreCanonicalCache(active: ActiveSnapshot): void {
  const version = cloneRow(active.activeVersion); version.cacheStatus = "Active"; version.activationState = "Active";
  resizeAndWrite(active.rpgTable, RPG_COLUMNS, active.rpgRows);
  resizeAndWrite(active.scopeTable, SCOPE_COLUMNS, active.scopeRows);
  resizeAndWrite(active.periodTable, PERIOD_COLUMNS, active.periodRows);
  resizeAndWrite(active.versionTable, VERSION_COLUMNS, [version]);
}

function validatePublishedPostconditions(workbook: ExcelScript.Workbook, version: string,
  fingerprint: string, period: string): void {
  const active = activeSnapshot(workbook);
  if (text(active.activeVersion.cacheVersion) !== version || text(active.activeVersion.cacheFingerprint) !== fingerprint) {
    throw new Error("PUL-030WPR-126: Published authority does not match Candidate.");
  }
  if (period && active.periodRows.filter(row => text(row.sourcePeriodKey) === period).length !== 1) {
    throw new Error(`PUL-030WPR-127: Published period ${period} is unavailable.`);
  }
  if (!workbook.getWorksheet(ROLLBACK_SHEET) || qaPassCount(workbook, "tblWeeklyPerformanceQA") !== 16 ||
      qaPassCount(workbook, "tblPerformanceInteractionQA") !== 16) {
    throw new Error("PUL-030WPR-128: Rollback or QA postcondition failed.");
  }
}

function validateParsedReport(report: ParsedReport): void {
  if (!report || report.parserVersion !== PARSER_VERSION || report.schemaVersion !== SCHEMA_VERSION ||
      !report.manifest || report.manifest.parserVersion !== PARSER_VERSION ||
      report.manifest.schemaVersion !== SCHEMA_VERSION ||
      report.manifest.contentReconciliationStatus !== "PASS" || !Array.isArray(report.rows)) {
    throw new Error("PUL-030WPR-129: Parsed report contract is invalid.");
  }
  if (report.rows.length !== number(report.manifest.sourceRowCount)) throw new Error("PUL-030WPR-130: Parser row count differs.");
  const total = zeroMetric();
  const seen: { [key: string]: boolean } = {};
  for (const row of report.rows) {
    if (row.sourcePeriodKey !== report.manifest.sourcePeriodKey || row.sourceFileId !== report.manifest.sourceFileId ||
        seen[keyToken(row.sourceRowId)]) throw new Error("PUL-030WPR-131: Parser lineage is invalid or duplicated.");
    seen[keyToken(row.sourceRowId)] = true; addMetric(total, row);
  }
  const value = finalizeMetric(total);
  if (value.factCount !== report.manifest.sourceRowCount ||
      Math.abs(value.salesNok - report.manifest.totalSalesNok) > 0.005 ||
      Math.abs(value.quantity - report.manifest.totalQuantity) > 0.0000005) {
    throw new Error("PUL-030WPR-132: Parser totals do not reconcile.");
  }
}

function validateCompleteRows(periods: DataRow[], scopes: DataRow[], rpgs: DataRow[], groups: DataRow[]): void {
  const activeGroups = validatedActiveReportingGroups(groups);
  if (rpgs.length !== scopes.length * activeGroups.length) {
    throw new Error(`PUL-030WPR-133: Dense RPG grain has ${rpgs.length} rows; expected ${scopes.length * activeGroups.length}.`);
  }
  const periodSeen: { [key: string]: boolean } = {};
  const scopeSeen: { [key: string]: boolean } = {};
  const rpgSeen: { [key: string]: boolean } = {};
  const activeGroupIds: { [key: string]: boolean } = {};
  for (const group of activeGroups) activeGroupIds[keyToken(text(group.reportingGroupId))] = true;
  const mappedScope = zeroMetric(); const mappedRpg = zeroMetric();
  for (const scope of scopes) {
    const key = keyToken(`${text(scope.cacheVersion)}|${text(scope.sourcePeriodKey)}|${text(scope.restaurantId)}`);
    if (scopeSeen[key]) throw new Error("PUL-030WPR-134: Duplicate scope grain."); scopeSeen[key] = true;
    const source = metricFromRow(scope, "source"); const coverage = zeroMetric();
    for (const prefix of STATE_PREFIXES) addMetric(coverage, metricFromRow(scope, prefix));
    if (!equalMetric(source, coverage)) throw new Error(`PUL-030WPR-135: ${text(scope.sourcePeriodKey)} scope does not reconcile.`);
    addMetric(mappedScope, metricFromRow(scope, "mapped"));
  }
  for (const rpg of rpgs) {
    const key = keyToken(`${text(rpg.cacheVersion)}|${text(rpg.sourcePeriodKey)}|${text(rpg.restaurantId)}|${text(rpg.reportingGroupId)}`);
    if (rpgSeen[key]) throw new Error("PUL-030WPR-136: Duplicate RPG grain."); rpgSeen[key] = true;
    if (!activeGroupIds[keyToken(text(rpg.reportingGroupId))]) {
      throw new Error(`PUL-030WPR-136: RPG grain references inactive or unknown ReportingGroupID ${text(rpg.reportingGroupId)}.`);
    }
    addMetric(mappedRpg, { factCount: number(rpg.mappedFactCount), salesNok: number(rpg.mappedSalesNok), quantity: number(rpg.mappedQuantity) });
  }
  for (const scope of scopes) {
    for (const group of activeGroups) {
      const key = keyToken(`${text(scope.cacheVersion)}|${text(scope.sourcePeriodKey)}|${text(scope.restaurantId)}|${text(group.reportingGroupId)}`);
      if (!rpgSeen[key]) {
        throw new Error(`PUL-030WPR-136: ${text(scope.sourcePeriodKey)}/${text(scope.restaurantId)} is missing dense ReportingGroupID ${text(group.reportingGroupId)}.`);
      }
    }
  }
  if (!equalMetric(mappedScope, mappedRpg)) throw new Error("PUL-030WPR-137: Mapped scope and RPG rows differ.");
  for (const period of periods) {
    const key = keyToken(text(period.sourcePeriodKey)); if (periodSeen[key]) throw new Error("PUL-030WPR-138: Duplicate period.");
    periodSeen[key] = true; const total = zeroMetric();
    for (const scope of scopes) if (text(scope.sourcePeriodKey) === text(period.sourcePeriodKey)) addMetric(total, metricFromRow(scope, "source"));
    if (!equalMetric(total, { factCount: number(period.sourceFactCount), salesNok: number(period.sourceSalesNok), quantity: number(period.sourceQuantity) })) {
      throw new Error(`PUL-030WPR-139: ${text(period.sourcePeriodKey)} does not reconcile.`);
    }
  }
}

function validatedActiveReportingGroups(groups: DataRow[]): DataRow[] {
  const catalogIds: { [key: string]: boolean } = {};
  const activeSortOrders: { [key: string]: boolean } = {};
  const active: DataRow[] = [];
  for (const row of groups) {
    const reportingGroupId = text(row.reportingGroupId);
    if (!reportingGroupId) throw new Error("PUL-030WPR-106: Reporting Group catalog contains a blank ReportingGroupID.");
    const idKey = keyToken(reportingGroupId);
    if (catalogIds[idKey]) throw new Error(`PUL-030WPR-106: Reporting Group catalog repeats ${reportingGroupId}.`);
    catalogIds[idKey] = true;
    if (text(row.active) !== "Yes") continue;
    const sortOrder = Number(row.sortOrder);
    if (!Number.isFinite(sortOrder)) {
      throw new Error(`PUL-030WPR-106: Active Reporting Group ${reportingGroupId} has an invalid SortOrder.`);
    }
    const sortKey = keyToken(String(sortOrder));
    if (activeSortOrders[sortKey]) {
      throw new Error(`PUL-030WPR-106: Active Reporting Groups repeat SortOrder ${sortOrder}.`);
    }
    activeSortOrders[sortKey] = true;
    active.push(row);
  }
  if (!active.length) throw new Error("PUL-030WPR-106: At least one active Reporting Group is required.");
  active.sort((left, right) => number(left.sortOrder) - number(right.sortOrder) ||
    compareText(left.reportingGroupId, right.reportingGroupId));
  return active;
}

function fingerprintCorpus(rows: DataRow[]): string {
  const records: string[] = [];
  for (const row of rows) records.push(record("FILE", [row.sourceFileId, row.sourceSemanticFingerprint,
    row.sourcePeriodKey, row.sourceFactCount, row.sourceSalesNok, row.sourceQuantity, row.scopeFingerprint]));
  return hashStrings(records, "WSC-");
}

function fingerprintCache(version: string, periods: DataRow[], scopes: DataRow[], rpgs: DataRow[]): string {
  const records = [record("CACHE_VERSION", [version])];
  for (const row of periods) records.push(record("PERIOD", [row.sourcePeriodKey, canonicalDate(row.periodStart),
    canonicalDate(row.periodEnd), row.sourceFileId, row.sourceSemanticFingerprint, row.sourceFactCount,
    sales(row.sourceSalesNok), quantity(row.sourceQuantity)]));
  for (const row of scopes) {
    const values: CellValue[] = [row.sourcePeriodKey, row.restaurantId, row.performanceEligible];
    appendMetricFields(values, metricFromRow(row, "source"));
    for (const prefix of STATE_PREFIXES) appendMetricFields(values, metricFromRow(row, prefix));
    records.push(record("SCOPE", values));
  }
  for (const row of rpgs) records.push(record("RPG", [row.sourcePeriodKey, row.restaurantId,
    row.reportingGroupId, row.mappedFactCount, sales(row.mappedSalesNok), quantity(row.mappedQuantity)]));
  records.sort(compareTextValue);
  return hashStrings(records, "WCC-");
}

function rekeyPeriod(row: DataRow, version: string): DataRow {
  const result = cloneRow(row); result.cacheVersion = version;
  result.weeklyPeriodManifestRowId = stableId("WPER-", [version, row.sourcePeriodKey]); return result;
}
function rekeyScope(row: DataRow, version: string): DataRow {
  const result = cloneRow(row); result.cacheVersion = version;
  result.weeklyScopeCacheRowId = stableId("WSCP-", [version, row.sourcePeriodKey, row.restaurantId]); return result;
}
function rekeyRpg(row: DataRow, version: string): DataRow {
  const result = cloneRow(row); result.cacheVersion = version;
  result.weeklyRpgCacheRowId = stableId("WRPG-", [version, row.sourcePeriodKey, row.restaurantId, row.reportingGroupId]); return result;
}

function createSection(sheet: ExcelScript.Worksheet, startRow: number, columns: string[][],
  rows: DataRow[], tableName: string, startColumn: number = 0): void {
  const headers = columns.map(value => value[0]);
  sheet.getRangeByIndexes(startRow - 1, startColumn, 1, columns.length).setValues([headers]);
  const values = rowsToValues(rows, columns);
  const chunk = 500;
  for (let offset = 0; offset < values.length; offset += chunk) {
    const part = values.slice(offset, Math.min(values.length, offset + chunk));
    sheet.getRangeByIndexes(startRow + offset, startColumn, part.length, columns.length).setValues(part);
  }
  const table = sheet.addTable(sheet.getRangeByIndexes(startRow - 1, startColumn,
    rows.length + 1, columns.length), true);
  table.setName(tableName); table.setPredefinedTableStyle("TableStyleMedium2");
}

function createIdentitySections(sheet: ExcelScript.Worksheet, registry: IdentityRegistry,
  tableNames: string[], startColumn: number): void {
  const control = 1;
  const restaurants = control + 4;
  const classifications = restaurants + registry.restaurants.length + 3;
  const products = classifications + registry.classifications.length + 3;
  createSection(sheet, control, IDENTITY_CONTROL_COLUMNS, [registry.control], tableNames[0], startColumn);
  createSection(sheet, restaurants, RESTAURANT_COLUMNS, registry.restaurants, tableNames[1], startColumn);
  createSection(sheet, classifications, CLASS_COLUMNS, registry.classifications, tableNames[2], startColumn);
  createSection(sheet, products, IDENTITY_PRODUCT_COLUMNS, registry.products, tableNames[3], startColumn);
}

function readIdentityRegistry(workbook: ExcelScript.Workbook, sheetName: string,
  tableNames: string[], expectedIdentityPreflight: string): IdentityRegistry {
  const sheet = requiredSheet(workbook, sheetName);
  if (sheet.getVisibility() !== ExcelScript.SheetVisibility.hidden) {
    throw new Error(`PUL-030WPR-146: ${sheetName} must remain hidden.`);
  }
  const controls = readObjects(requiredTable(workbook, tableNames[0]), IDENTITY_CONTROL_COLUMNS);
  if (controls.length !== 1) throw new Error(`PUL-030WPR-147: ${tableNames[0]} must contain one row.`);
  const registry: IdentityRegistry = {
    control: controls[0],
    restaurants: readObjects(requiredTable(workbook, tableNames[1]), RESTAURANT_COLUMNS),
    classifications: readObjects(requiredTable(workbook, tableNames[2]), CLASS_COLUMNS),
    products: readObjects(requiredTable(workbook, tableNames[3]), IDENTITY_PRODUCT_COLUMNS)
  };
  const control = registry.control;
  if (text(control.identityRegistryVersion) !== IDENTITY_REGISTRY_VERSION ||
      text(control.identityPreflightFingerprint) !== expectedIdentityPreflight ||
      number(control.restaurantCount) !== registry.restaurants.length ||
      number(control.classificationCount) !== registry.classifications.length ||
      number(control.productCount) !== registry.products.length) {
    throw new Error(`PUL-030WPR-148: ${sheetName} identity registry control is stale.`);
  }
  const fingerprint = fingerprintIdentityRegistry(registry);
  if (fingerprint !== text(control.identityRegistryFingerprint)) {
    throw new Error(`PUL-030WPR-149: ${sheetName} identity registry fingerprint differs.`);
  }
  validateUniqueIdentityRows(registry);
  return registry;
}

function writeIdentityRegistry(workbook: ExcelScript.Workbook, registry: IdentityRegistry): void {
  const existing = workbook.getWorksheet(IDENTITY_SHEET);
  if (existing) existing.delete();
  const sheet = workbook.addWorksheet(IDENTITY_SHEET);
  sheet.setVisibility(ExcelScript.SheetVisibility.hidden);
  createIdentitySections(sheet, registry, IDENTITY_TABLES, 0);
}

function mergeIdentityRegistry(active: IdentityRegistry, identity: IdentityPlan,
  identityPreflightFingerprint: string): IdentityRegistry {
  const restaurants = mergeIdentityRows(active.restaurants, identity.restaurantCandidates,
    "restaurantId");
  const classifications = mergeIdentityRows(active.classifications,
    identity.classificationCandidates, "sourceClassificationId");
  const products = mergeIdentityRows(active.products, identity.productCandidates, "productId");
  const registry: IdentityRegistry = {
    control: {}, restaurants, classifications, products
  };
  registry.control = {
    identityRegistryVersion: IDENTITY_REGISTRY_VERSION,
    identityPreflightFingerprint,
    identityRegistryFingerprint: fingerprintIdentityRegistry(registry),
    restaurantCount: restaurants.length,
    classificationCount: classifications.length,
    productCount: products.length
  };
  return registry;
}

function mergeIdentityRows(existing: DataRow[], candidates: DataRow[], idField: string): DataRow[] {
  const result: DataRow[] = [];
  const index: { [key: string]: DataRow } = {};
  for (const row of existing) {
    const value = cloneRow(row); result.push(value); index[keyToken(text(value[idField]))] = value;
  }
  for (const candidate of candidates) {
    const token = keyToken(text(candidate[idField])); const prior = index[token];
    if (prior && canonicalRow(prior) !== canonicalRow(candidate)) {
      throw new Error(`PUL-030WPR-150: Accepted identity ${text(candidate[idField])} changed.`);
    }
    if (!prior) {
      const value = cloneRow(candidate); result.push(value); index[token] = value;
    }
  }
  return result.sort((left, right) => compareText(left[idField], right[idField]));
}

function validateIdentityRegistryCarryForward(active: IdentityRegistry,
  candidate: IdentityRegistry): void {
  validateIdentityRowsCarryForward(active.restaurants, candidate.restaurants, "restaurantId");
  validateIdentityRowsCarryForward(active.classifications, candidate.classifications,
    "sourceClassificationId");
  validateIdentityRowsCarryForward(active.products, candidate.products, "productId");
}

function validateIdentityRowsCarryForward(active: DataRow[], candidate: DataRow[], idField: string): void {
  const index: { [key: string]: DataRow } = {};
  for (const row of candidate) index[keyToken(text(row[idField]))] = row;
  for (const row of active) {
    const value = index[keyToken(text(row[idField]))];
    if (!value || canonicalRow(value) !== canonicalRow(row)) {
      throw new Error(`PUL-030WPR-151: Candidate does not preserve ${text(row[idField])}.`);
    }
  }
}

function validateIdentityRegistryAgainstCatalogs(registry: IdentityRegistry, products: DataRow[],
  classifications: DataRow[], restaurants: DataRow[]): void {
  validateNoIdentityCollision(registry.products, products, "productId", "productKey");
  validateNoIdentityCollision(registry.classifications, classifications,
    "sourceClassificationId", "sourceClassificationKey");
  validateNoIdentityCollision(registry.restaurants, restaurants, "restaurantId",
    "sourceRestaurantName");
}

function validateNoIdentityCollision(registry: DataRow[], catalog: DataRow[],
  idField: string, keyField: string): void {
  const ids: { [key: string]: boolean } = {}; const keys: { [key: string]: boolean } = {};
  for (const row of catalog) {
    ids[keyToken(text(row[idField]))] = true; keys[keyToken(text(row[keyField]))] = true;
  }
  for (const row of registry) {
    if (ids[keyToken(text(row[idField]))] || keys[keyToken(text(row[keyField]))]) {
      throw new Error(`PUL-030WPR-152: Identity registry collides with catalog ${text(row[idField])}.`);
    }
  }
}

function validateUniqueIdentityRows(registry: IdentityRegistry): void {
  validateUniqueIdentityType(registry.restaurants, "restaurantId", "sourceRestaurantName");
  validateUniqueIdentityType(registry.classifications, "sourceClassificationId",
    "sourceClassificationKey");
  validateUniqueIdentityType(registry.products, "productId", "productKey");
}

function validateUniqueIdentityType(rows: DataRow[], idField: string, keyField: string): void {
  const ids: { [key: string]: boolean } = {}; const keys: { [key: string]: boolean } = {};
  for (const row of rows) {
    const id = keyToken(text(row[idField])); const key = keyToken(text(row[keyField]));
    if (ids[id] || keys[key]) throw new Error(`PUL-030WPR-153: Duplicate identity ${text(row[idField])}.`);
    ids[id] = true; keys[key] = true;
  }
}

function fingerprintIdentityRegistry(registry: IdentityRegistry): string {
  const records = [record("CONTRACT", [IDENTITY_REGISTRY_VERSION])];
  for (const row of registry.restaurants) records.push(record("RESTAURANT", [
    row.restaurantId, row.sourceSystemId, row.sourceRestaurantName, row.displayName,
    row.status, row.reportingEnabled
  ]));
  for (const row of registry.classifications) records.push(record("CLASSIFICATION", [
    row.sourceClassificationId, row.sourceSystemId, row.sourceMainCategory,
    row.sourceSubCategory, row.sourceClassificationKey, row.status
  ]));
  for (const row of registry.products) {
    const values: CellValue[] = [row.productId, row.sourceSystemId, row.sourceProductName,
      row.salesAccount, row.sourceClassificationId, row.productKey, row.productStatus,
      row.hierarchyStatus];
    const paths = text(row.observedHierarchyPaths).split("\u001f")
      .filter(value => value !== "").sort(compareTextValue);
    for (const path of paths) values.push(path);
    records.push(record("PRODUCT", values));
  }
  records.sort(compareTextValue);
  return hashStrings(records, "WIR-");
}

function canonicalRow(row: DataRow): string {
  const keys: string[] = [];
  for (const key in row) keys.push(key);
  keys.sort(compareTextValue);
  const values: CellValue[] = [];
  for (const key of keys) values.push(key, row[key]);
  return record("ROW", values);
}

function resizeAndWrite(table: ExcelScript.Table, columns: string[][], rows: DataRow[]): void {
  const range = table.getRange(); const sheet = range.getWorksheet();
  const target = sheet.getRangeByIndexes(range.getRowIndex(), range.getColumnIndex(), rows.length + 1, columns.length);
  table.resize(target);
  const values = rowsToValues(rows, columns); const chunk = 500;
  const body = table.getRangeBetweenHeaderAndTotal();
  for (let offset = 0; offset < values.length; offset += chunk) {
    const part = values.slice(offset, Math.min(values.length, offset + chunk));
    body.getCell(offset, 0).getResizedRange(part.length - 1, columns.length - 1).setValues(part);
  }
}

function ensureTableSpace(current: ExcelScript.Table, next: ExcelScript.Table, requiredRows: number): void {
  const currentRange = current.getRange(); const nextRange = next.getRange();
  const requiredEnd = currentRange.getRowIndex() + requiredRows + 1;
  const gap = nextRange.getRowIndex() - requiredEnd;
  if (gap >= 0) return;
  const insertCount = -gap;
  const first = nextRange.getRowIndex() + 1;
  nextRange.getWorksheet().getRange(`${first}:${first + insertCount - 1}`)
    .insert(ExcelScript.InsertShiftDirection.down);
}

function rowsToValues(rows: DataRow[], columns: string[][]): CellValue[][] {
  const values: CellValue[][] = [];
  for (const row of rows) {
    const value: CellValue[] = [];
    for (const column of columns) value.push(row[column[1]] === undefined ? "" : row[column[1]]);
    values.push(value);
  }
  return values;
}

function readObjects(table: ExcelScript.Table, columns: string[][]): DataRow[] {
  const headers = table.getHeaderRowRange().getTexts()[0]; const values = table.getRangeBetweenHeaderAndTotal().getValues();
  const positions: number[] = [];
  for (const column of columns) {
    const index = headers.indexOf(column[0]);
    if (index < 0) throw new Error(`PUL-030WPR-140: ${table.getName()} lacks ${column[0]}.`);
    positions.push(index);
  }
  const result: DataRow[] = [];
  for (const value of values) {
    if (!text(value[0])) continue; const row: DataRow = {};
    for (let index = 0; index < columns.length; index += 1) row[columns[index][1]] = value[positions[index]];
    result.push(row);
  }
  return result;
}

function indexRows(rows: DataRow[], keySelector: (row: DataRow) => string): { [key: string]: DataRow[] } {
  const result: { [key: string]: DataRow[] } = {};
  for (const row of rows) {
    const key = keyToken(keySelector(row)); const group = result[key] || []; group.push(row); result[key] = group;
  }
  return result;
}
function singleRow(index: { [key: string]: DataRow[] }, key: string): DataRow | null {
  const rows = index[keyToken(key)] || []; return rows.length === 1 ? rows[0] : null;
}
function unmatchedKeys(observed: { [key: string]: string }, existing: { [key: string]: DataRow[] }): string[] {
  const result: string[] = [];
  for (const token in observed) if (!existing[token]) result.push(token.substring(2));
  return result.sort(compareTextValue);
}
function unmatchedObjectKeys(observed: { [key: string]: DataRow }, existing: { [key: string]: DataRow[] }, field: string): string[] {
  const result: string[] = [];
  for (const token in observed) if (!existing[token]) result.push(text(observed[token][field]));
  return result.sort(compareTextValue);
}
function allocateIds(rows: DataRow[], field: string, prefix: string, digits: number, keys: string[]): { [key: string]: string } {
  let maximum = 0; const result: { [key: string]: string } = {};
  for (const row of rows) { const match = new RegExp(`^${prefix}-(\\d+)$`).exec(text(row[field])); if (match) maximum = Math.max(maximum, Number(match[1])); }
  for (const key of keys.slice().sort(compareTextValue)) { maximum += 1; result[keyToken(key)] = `${prefix}-${String(maximum).padStart(digits, "0")}`; }
  return result;
}

function setMetric(row: DataRow, prefix: string, metric: Metric): void {
  row[`${prefix}FactCount`] = metric.factCount; row[`${prefix}SalesNok`] = metric.salesNok; row[`${prefix}Quantity`] = metric.quantity;
}
function addMetricToRow(row: DataRow, prefix: string, source: ParsedRow): void {
  row[`${prefix}FactCount`] = number(row[`${prefix}FactCount`]) + 1;
  row[`${prefix}SalesNok`] = number(row[`${prefix}SalesNok`]) + source.salesNok;
  row[`${prefix}Quantity`] = number(row[`${prefix}Quantity`]) + source.quantity;
}
function finalizeMetricRow(row: DataRow): DataRow {
  const result = cloneRow(row);
  for (const prefix of ["source", "mapped", "unmapped", "identityPending", "conflict", "inactiveTarget"]) {
    result[`${prefix}FactCount`] = number(result[`${prefix}FactCount`]);
    result[`${prefix}SalesNok`] = round(number(result[`${prefix}SalesNok`]), 2);
    result[`${prefix}Quantity`] = round(number(result[`${prefix}Quantity`]), 6);
  }
  return result;
}
function metricFromRow(row: DataRow, prefix: string): Metric {
  return { factCount: number(row[`${prefix}FactCount`]), salesNok: number(row[`${prefix}SalesNok`]), quantity: number(row[`${prefix}Quantity`]) };
}
function stateTotals(rows: DataRow[], prefix: string): Metric {
  const value = zeroMetric(); for (const row of rows) addMetric(value, metricFromRow(row, prefix)); return finalizeMetric(value);
}
function zeroMetric(): Metric { return { factCount: 0, salesNok: 0, quantity: 0 }; }
function addMetric(metric: Metric, row: { factCount?: number; salesNok?: number; quantity?: number }): void {
  metric.factCount += number(row.factCount === undefined ? 1 : row.factCount);
  metric.salesNok += number(row.salesNok); metric.quantity += number(row.quantity);
}
function finalizeMetric(metric: Metric): Metric { return { factCount: number(metric.factCount), salesNok: round(metric.salesNok, 2), quantity: round(metric.quantity, 6) }; }
function equalMetric(left: Metric, right: Metric): boolean { const a = finalizeMetric(left); const b = finalizeMetric(right); return a.factCount === b.factCount && Math.abs(a.salesNok - b.salesNok) <= 0.005 && Math.abs(a.quantity - b.quantity) <= 0.0000005; }
function appendMetricFields(values: CellValue[], metric: Metric): void { const value = finalizeMetric(metric); values.push(value.factCount, sales(value.salesNok), quantity(value.quantity)); }

function intakeLedgerEntry(report: ParsedReport, status: string, message: string,
  prior: string, resulting: string, superseded: string): DataRow {
  return { intakeEventId: stableId("WINT-", [report.manifest.sourcePeriodKey, report.manifest.semanticFingerprint, status, resulting]),
    sourceLocator: report.manifest.sourceLocator, sourceFileId: report.manifest.sourceFileId,
    sourcePeriodKey: report.manifest.sourcePeriodKey, sourceSemanticFingerprint: report.manifest.semanticFingerprint,
    identityPreflightFingerprint: "", intakeStatus: status, statusMessage: message,
    sourceRowCount: report.manifest.sourceRowCount, sourceSalesNok: round(report.manifest.totalSalesNok, 2),
    processedAt: new Date().toISOString(), priorCacheVersion: prior,
    resultingCacheVersion: resulting, supersededCacheVersion: superseded };
}
function publicationLedgerEntry(workbook: ExcelScript.Workbook, period: DataRow,
  version: DataRow, prior: string): DataRow {
  const key = text(period.sourcePeriodKey); const semantic = text(period.sourceSemanticFingerprint);
  return { intakeEventId: stableId("WINT-", [key, semantic, "Published", version.cacheVersion]),
    sourceLocator: ledgerSourceLocator(workbook, key, semantic),
    sourceFileId: period.sourceFileId, sourcePeriodKey: key,
    sourceSemanticFingerprint: semantic, identityPreflightFingerprint: version.identityPreflightFingerprint,
    intakeStatus: "Published", statusMessage: `Published ${key} in ${text(version.cacheVersion)}; prior ${prior} retained for rollback.`,
    sourceRowCount: period.sourceFactCount, sourceSalesNok: period.sourceSalesNok,
    processedAt: new Date().toISOString(), priorCacheVersion: prior,
    resultingCacheVersion: version.cacheVersion, supersededCacheVersion: prior };
}

function ledgerSourceLocator(workbook: ExcelScript.Workbook, period: string, semantic: string): string {
  const table = requiredTable(workbook, "tblWeeklyIntakeLog");
  const headers = table.getHeaderRowRange().getTexts()[0];
  const values = table.getRangeBetweenHeaderAndTotal().getTexts();
  const locator = headers.indexOf("SourceLocator");
  const periodIndex = headers.indexOf("SourcePeriodKey");
  const semanticIndex = headers.indexOf("SourceSemanticFingerprint");
  for (let index = values.length - 1; index >= 0; index -= 1) {
    const row = values[index];
    if (row[periodIndex] === period && row[semanticIndex] === semantic) return row[locator];
  }
  return "";
}

function upsertLedger(workbook: ExcelScript.Workbook, entry: DataRow): void {
  const table = requiredTable(workbook, "tblWeeklyIntakeLog");
  const headers = ["IntakeEventID", "SourceLocator", "SourceFileID", "SourcePeriodKey",
    "SourceSemanticFingerprint", "IdentityPreflightFingerprint", "IntakeStatus", "StatusMessage",
    "SourceRowCount", "SourceSalesNOK", "ProcessedAt", "PriorCacheVersion",
    "ResultingCacheVersion", "SupersededCacheVersion"];
  const columns = ["intakeEventId", "sourceLocator", "sourceFileId", "sourcePeriodKey",
    "sourceSemanticFingerprint", "identityPreflightFingerprint", "intakeStatus", "statusMessage",
    "sourceRowCount", "sourceSalesNok", "processedAt", "priorCacheVersion",
    "resultingCacheVersion", "supersededCacheVersion"];
  if (table.getHeaderRowRange().getTexts()[0].join("|") !== headers.join("|")) throw new Error("PUL-030WPR-141: Intake ledger schema differs.");
  const rows = table.getRangeBetweenHeaderAndTotal().getValues(); const values: CellValue[] = [];
  for (const column of columns) values.push(entry[column] === undefined ? "" : entry[column]);
  for (const row of rows) if (text(row[0]) === text(entry.intakeEventId)) return;
  table.addRow(-1, values);
}

function publicationResult(status: string, period: string, message: string, changed: boolean,
  prior: string, resulting: string, fingerprint: string, archive: boolean,
  value: { periodRows: DataRow[]; scopeRows: DataRow[]; rpgRows: DataRow[] }, pending: Metric): WeeklyPublicationResult {
  return { status, period, message, cacheChanged: changed, priorCacheVersion: prior,
    resultingCacheVersion: resulting, cacheFingerprint: fingerprint, archiveReady: archive,
    periodRowCount: value.periodRows.length, scopeRowCount: value.scopeRows.length,
    rpgRowCount: value.rpgRows.length, identityPendingFactCount: pending.factCount,
    identityPendingSalesNok: pending.salesNok, identityPendingQuantity: pending.quantity };
}

function layout(periodCount: number, scopeCount: number): number[] {
  const version = 1; const period = 5; const scope = period + periodCount + 3;
  const rpg = scope + scopeCount + 3; return [version, period, scope, rpg];
}
function rowsForPeriod(rows: DataRow[], key: string): DataRow[] { return rows.filter(row => text(row.sourcePeriodKey) === key); }
function newestPeriod(rows: DataRow[]): string { return rows.slice().sort(comparePeriod).map(row => text(row.sourcePeriodKey)).pop() || ""; }
function comparePeriod(left: DataRow, right: DataRow): number { return compareTextValue(canonicalDate(left.periodStart), canonicalDate(right.periodStart)) || compareText(left.sourcePeriodKey, right.sourcePeriodKey); }
function compareScope(left: DataRow, right: DataRow): number { return compareText(left.sourcePeriodKey, right.sourcePeriodKey) || compareText(left.restaurantId, right.restaurantId); }
function compareRpg(left: DataRow, right: DataRow): number { return compareText(left.sourcePeriodKey, right.sourcePeriodKey) || compareText(left.restaurantId, right.restaurantId) || compareText(left.reportingGroupId, right.reportingGroupId); }
function isNonzeroRpg(row: DataRow): boolean { return number(row.mappedFactCount) !== 0 || number(row.mappedSalesNok) !== 0 || number(row.mappedQuantity) !== 0; }
function cloneRow(row: DataRow): DataRow { const result: DataRow = {}; for (const key in row) result[key] = row[key]; return result; }
function keyToken(value: string): string { return `K|${value}`; }
function restaurantKey(source: string, name: string): string { return `${source} || ${name}`; }
function productKey(source: string, item: string, account: string): string { return `${source} || ${item} || ${account}`; }
function classificationKey(source: string, main: string, sub: string): string { return `${source} || ${main} || ${sub}`; }
function stableId(prefix: string, values: CellValue[]): string { return hashStrings([record("ID", values)], prefix); }
function record(kind: string, values: CellValue[]): string { const parts: string[] = []; for (const value of values) { const normalized = value === null || value === undefined ? "" : String(value); parts.push(`${normalized.length}:${normalized}`); } return `${kind}|${parts.join("|")}`; }
function hashStrings(values: string[], prefix: string): string { let left = 0; let right = 0; for (const item of values) { const value = `${item}\n`; for (let index = 0; index < value.length; index += 1) { const code = value.charCodeAt(index); left = (left * 131 + code) % 2147483647; right = (right * 137 + code) % 2147483629; } } return `${prefix}${left.toString(16).padStart(8, "0")}${right.toString(16).padStart(8, "0")}`; }
function qaPassCount(workbook: ExcelScript.Workbook, name: string): number { const table = requiredTable(workbook, name); const headers = table.getHeaderRowRange().getTexts()[0]; const index = headers.indexOf("Result"); let count = 0; const rows = table.getRangeBetweenHeaderAndTotal().getTexts(); for (const row of rows) if (row[index] === "PASS") count += 1; return count; }
function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table { const table = workbook.getTable(name); if (!table) throw new Error(`PUL-030WPR-142: Required table ${name} is missing.`); return table; }
function requiredSheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet { const sheet = workbook.getWorksheet(name); if (!sheet) throw new Error(`PUL-030WPR-143: Required sheet ${name} is missing.`); return sheet; }
function parseJson(value: string): PublicationPayload { try { return JSON.parse(value || "{}") as PublicationPayload; } catch (error) { throw new Error(`PUL-030WPR-144: payloadJson is invalid. ${String(error)}`); } }
function canonicalDate(value: CellValue): string { const raw = text(value); if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw; const serial = number(value); if (!serial) return raw; return new Date(Math.round((serial - 25569) * 86400000)).toISOString().slice(0, 10); }
function excelDateSerial(value: string): number { return Math.round(Date.parse(`${value}T00:00:00Z`) / 86400000) + 25569; }
function nullableNumber(value: CellValue, fallback: number): number { return value === "" || value === null || value === undefined ? fallback : number(value); }
function sales(value: CellValue): string { return number(value).toFixed(2); }
function quantity(value: CellValue): string { return number(value).toFixed(6); }
function round(value: number, decimals: number): number { const factor = Math.pow(10, decimals); return Math.round((value + Number.EPSILON) * factor) / factor; }
function number(value: unknown): number { const parsed = Number(value); return Number.isFinite(parsed) ? parsed : 0; }
function text(value: unknown): string { return String(value === null || value === undefined ? "" : value).trim(); }
function compareText(left: CellValue, right: CellValue): number { return compareTextValue(text(left), text(right)); }
function compareTextValue(left: string, right: string): number { return left < right ? -1 : left > right ? 1 : 0; }
