/**
 * Table-shaped materialization contract for the derived weekly Mapping
 * attention projection. This module does not open or mutate a workbook.
 */

export const WEEKLY_MAPPING_ATTENTION_SHEET = "_Weekly_Mapping_Attention";
export const WEEKLY_MAPPING_ATTENTION_STAGING_SHEET = "_Weekly_Mapping_Attn_Stage";

export const WEEKLY_MAPPING_ATTENTION_TABLES = Object.freeze({
  control: "tblWeeklyMappingAttentionControl",
  products: "tblWeeklyMappingAttention",
});

export const WEEKLY_MAPPING_ATTENTION_COLUMNS = Object.freeze({
  control: Object.freeze([
    ["ProjectionVersion", "projectionVersion"],
    ["ValidationStatus", "validationStatus"],
    ["HealthStatus", "healthStatus"],
    ["ProjectionFingerprint", "projectionFingerprint"],
    ["ActiveCacheVersion", "activeCacheVersion"],
    ["ActiveCacheFingerprint", "activeCacheFingerprint"],
    ["SourceCorpusFingerprint", "sourceCorpusFingerprint"],
    ["MappingContentFingerprint", "mappingContentFingerprint"],
    ["CatalogContentFingerprint", "catalogContentFingerprint"],
    ["IdentityPreflightFingerprint", "identityPreflightFingerprint"],
    ["ThroughPeriodLabel", "throughPeriodLabel"],
    ["PeriodRowCount", "periodRowCount"],
    ["ProductRowCount", "productRowCount"],
    ["ExistingProductCount", "existingProductCount"],
    ["WeeklyAddedProductCount", "weeklyAddedProductCount"],
    ["SourceFactCount", "sourceFactCount"],
    ["SourceSalesNOK", "sourceSalesNok"],
    ["SourceQuantity", "sourceQuantity"],
  ]),
  products: Object.freeze([
    ["ProductID", "productId"],
    ["SourceSystemID", "sourceSystemId"],
    ["Item", "item"],
    ["Main Category", "sourceMainCategory"],
    ["Subcategory", "sourceSubCategory"],
    ["Sales Account", "salesAccount"],
    ["SourceClassificationID", "sourceClassificationId"],
    ["ProductKey", "productKey"],
    ["Identity Origin", "identityOrigin"],
    ["ReportingGroupID", "effectiveReportingGroupId"],
    ["Reporting Group", "effectiveReportingGroupName"],
    ["Resolution", "resolutionType"],
    ["Mapping State", "mappingStatus"],
    ["WinningRuleID", "winningRuleId"],
    ["Historical Facts", "historicalFactCount"],
    ["Historical Sales NOK", "historicalSalesNok"],
    ["Historical Quantity", "historicalQuantity"],
    ["Hierarchy Attention", "hierarchyAttention"],
    ["Hierarchy Alternatives", "hierarchyAlternatives"],
  ]),
});

export function buildWeeklyMappingAttentionMaterializationPlan(cache) {
  const projection = cache?.mappingAttentionProjection ?? cache?.identityPreflight?.mappingAttentionProjection;
  const manifest = cache?.versionManifest;
  if (!projection || projection.validationStatus !== "PASS") {
    throw new Error("PUL-030MA-101: Mapping attention projection must pass before materialization.");
  }
  if (!manifest || manifest.cacheStatus !== "Active" || manifest.activationState !== "Active" ||
      manifest.validationStatus !== "PASS") {
    throw new Error("PUL-030MA-102: Mapping attention requires one validated Active / Active cache.");
  }
  const cacheSource = cache?.validation?.corpusReconciliation?.source;
  if (!cacheSource || projection.sourceTotals.factCount !== cacheSource.factCount ||
      Math.abs(projection.sourceTotals.salesNok - cacheSource.salesNok) > 0.005 ||
      Math.abs(projection.sourceTotals.quantity - cacheSource.quantity) > 0.000005) {
    throw new Error("PUL-030MA-103: Projection totals differ from the Active weekly cache.");
  }
  const latestPeriod = [...cache.periodManifest].sort((left, right) =>
    Number(left.isoYear) - Number(right.isoYear) || Number(left.isoWeek) - Number(right.isoWeek)).at(-1);
  if (!latestPeriod) throw new Error("PUL-030MA-104: Active cache has no periods.");
  const throughPeriodLabel = `${latestPeriod.isoYear} W${String(latestPeriod.isoWeek).padStart(2, "0")}`;
  const control = {
    projectionVersion: projection.contractVersion,
    validationStatus: projection.validationStatus,
    healthStatus: projection.healthStatus,
    projectionFingerprint: projection.projectionFingerprint,
    activeCacheVersion: manifest.cacheVersion,
    activeCacheFingerprint: manifest.cacheFingerprint,
    sourceCorpusFingerprint: manifest.sourceCorpusFingerprint,
    mappingContentFingerprint: manifest.mappingContentFingerprint,
    catalogContentFingerprint: manifest.catalogContentFingerprint,
    identityPreflightFingerprint: manifest.identityPreflightFingerprint,
    throughPeriodLabel,
    periodRowCount: cache.periodManifest.length,
    productRowCount: projection.rows.length,
    existingProductCount: projection.existingProductCount,
    weeklyAddedProductCount: projection.weeklyAddedProductCount,
    sourceFactCount: projection.sourceTotals.factCount,
    sourceSalesNok: projection.sourceTotals.salesNok,
    sourceQuantity: projection.sourceTotals.quantity,
  };
  const sources = { control: [control], products: projection.rows };
  const layout = {
    control: { startRow: 1, startColumn: 1 },
    products: { startRow: 5, startColumn: 1 },
  };
  const sections = {};
  for (const name of Object.keys(WEEKLY_MAPPING_ATTENTION_COLUMNS)) {
    const columns = WEEKLY_MAPPING_ATTENTION_COLUMNS[name];
    const values = sources[name].map(row => columns.map(([, key]) => row[key] ?? ""));
    sections[name] = {
      name,
      tableName: WEEKLY_MAPPING_ATTENTION_TABLES[name],
      startRow: layout[name].startRow,
      startColumn: layout[name].startColumn,
      headers: columns.map(([header]) => header),
      values,
      rowCount: values.length,
      columnCount: columns.length,
    };
  }
  return {
    sheetName: WEEKLY_MAPPING_ATTENTION_SHEET,
    stagingSheetName: WEEKLY_MAPPING_ATTENTION_STAGING_SHEET,
    projectionFingerprint: projection.projectionFingerprint,
    activeCacheVersion: manifest.cacheVersion,
    activeCacheFingerprint: manifest.cacheFingerprint,
    throughPeriodLabel,
    stateCoverage: projection.stateCoverage,
    sections,
  };
}

export function weeklyMappingAttentionChunk(plan, sectionName, offset = 0, limit = 400) {
  const section = plan.sections[sectionName];
  if (!section) throw new Error(`PUL-030MA-105: Unknown attention section ${sectionName}.`);
  if (!Number.isInteger(offset) || offset < 0 || offset > section.rowCount) {
    throw new Error(`PUL-030MA-106: Invalid ${sectionName} row offset ${offset}.`);
  }
  if (!Number.isInteger(limit) || limit < 1) {
    throw new Error(`PUL-030MA-107: Invalid ${sectionName} chunk size ${limit}.`);
  }
  const values = section.values.slice(offset, offset + limit);
  return {
    projectionFingerprint: plan.projectionFingerprint,
    activeCacheVersion: plan.activeCacheVersion,
    activeCacheFingerprint: plan.activeCacheFingerprint,
    section: sectionName,
    offset,
    totalRowCount: section.rowCount,
    startRow: section.startRow + 1 + offset,
    startColumn: section.startColumn,
    values,
  };
}
