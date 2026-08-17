/**
 * Pulse Build 0.3.0 — Weekly intake runtime adapter.
 *
 * Power Automate calls Process with the accepted weekly parser metadata. This
 * script performs only bounded authority/manifest classification and ledger
 * writes. Identity resolution, mapping and Candidate construction remain in
 * the repository publisher. A repository-approved New Candidate may be
 * materialized with Prepare New -> Write New -> Finalize New; it stays hidden
 * and non-authoritative until a separately approved activation.
 */
function main(
  workbook: ExcelScript.Workbook,
  operation: string = "Process",
  payloadJson: string = "{}"
): WeeklyIntakeResult {
  const payload = parsePayload(payloadJson);
  const requested = text(operation) || text(payload.operation) || "Process";
  if (requested === "Process") return processIntake(workbook, payload);
  if (requested === "Prepare New") return prepareNewCandidate(workbook, payload);
  if (requested === "Write New") return writeNewCandidateChunk(workbook, payload);
  if (requested === "Finalize New") return finalizeNewCandidate(workbook, payload);
  throw new Error(`PUL-030WIA-100: Unsupported adapter operation ${requested}.`);
}

type CellValue = string | number | boolean;
type ParsedManifest = {
  sourceFileId?: string;
  sourceLocator?: string;
  semanticFingerprint?: string;
  sourcePeriodKey?: string;
  parserVersion?: string;
  schemaVersion?: string;
  sourceRowCount?: number;
  totalSalesNok?: number;
  contentReconciliationStatus?: string;
};
type ParsedReportMetadata = {
  parserVersion?: string;
  schemaVersion?: string;
  manifest?: ParsedManifest;
};
type LedgerEntry = {
  intakeEventId?: string;
  sourceLocator?: string;
  sourceFileId?: string;
  sourcePeriodKey?: string;
  sourceSemanticFingerprint?: string;
  identityPreflightFingerprint?: string;
  intakeStatus?: string;
  statusMessage?: string;
  sourceRowCount?: number;
  sourceSalesNok?: number;
  processedAt?: string;
  priorCacheVersion?: string;
  resultingCacheVersion?: string;
  supersededCacheVersion?: string;
};
type IntakePayload = {
  adapterVersion?: string;
  publisherVersion?: string;
  operation?: string;
  parseError?: string;
  processedAt?: string;
  expectedOutcome?: string;
  expectedLedgerEntry?: LedgerEntry;
  parsedReport?: ParsedReportMetadata;
  candidateVersion?: string;
  candidateFingerprint?: string;
  priorCacheVersion?: string;
  layoutFingerprint?: string;
  counts?: { period?: number; scope?: number; rpg?: number };
  section?: string;
  offset?: number;
  totalRowCount?: number;
  startRow?: number;
  startColumn?: number;
  values?: CellValue[][];
  chunkFingerprint?: string;
  versionValues?: CellValue[];
  ledgerEntry?: LedgerEntry;
};
type Authority = {
  available: boolean;
  cacheVersion: string;
  cacheFingerprint: string;
  identityPreflightFingerprint: string;
  message: string;
};
type CandidateLayout = {
  version: CandidateSection;
  period: CandidateSection;
  scope: CandidateSection;
  rpg: CandidateSection;
};
type CandidateSection = {
  name: string;
  tableName: string;
  startRow: number;
  startColumn: number;
  rowCount: number;
  columnCount: number;
  headers: string[];
};
interface WeeklyIntakeResult {
  status: string;
  period: string;
  message: string;
  cacheChanged: boolean;
  activeCacheVersion: string;
  resultingCacheVersion: string;
  ledgerEventId: string;
  ledgerAction: string;
  archiveReady: boolean;
  operation: string;
}

const ADAPTER_VERSION = "0.3.0-weekly-intake-adapter-v1";
const PUBLISHER_VERSION = "0.3.0-weekly-intake-publisher-v1";
const PARSER_VERSION = "0.3.0-weekly-parser-v1";
const SCHEMA_VERSION = "sales-per-item-v1";
const CACHE_SHEET = "_Weekly_Cache";
const CANDIDATE_SHEET = "_Weekly_Cache_Candidate";
const LEDGER_TABLE = "tblWeeklyIntakeLog";
const LEDGER_RANGE = "Y1:AL2";
const LEDGER_HEADERS = [
  "IntakeEventID", "SourceLocator", "SourceFileID", "SourcePeriodKey",
  "SourceSemanticFingerprint", "IdentityPreflightFingerprint", "IntakeStatus",
  "StatusMessage", "SourceRowCount", "SourceSalesNOK", "ProcessedAt",
  "PriorCacheVersion", "ResultingCacheVersion", "SupersededCacheVersion"
];
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

function processIntake(workbook: ExcelScript.Workbook, payload: IntakePayload): WeeklyIntakeResult {
  validateAdapterPayload(payload);
  const parsed = payload.parsedReport;
  const manifest = parsed && parsed.manifest ? parsed.manifest : {};
  const metadataErrors = parserMetadataErrors(parsed);
  let status = "";
  let message = "";
  let authority = unavailableAuthority("");
  if (text(payload.parseError)) {
    status = "Rejected";
    message = text(payload.parseError);
  } else if (metadataErrors.length) {
    status = "Rejected";
    message = metadataErrors.join(" ");
  } else {
    authority = resolveActiveAuthority(workbook);
    if (!authority.available) {
      status = "Cache Stale";
      message = authority.message;
    } else {
      const existing = activePeriodRows(workbook, authority.cacheVersion, text(manifest.sourcePeriodKey));
      if (existing.length > 1) {
        status = "Cache Stale";
        message = `Active cache repeats period ${text(manifest.sourcePeriodKey)}.`;
      } else if (!existing.length) {
        status = "New";
        message = "Period is not present in the Active cache; an accepted publisher Candidate is required.";
      } else if (text(existing[0][8]) === text(manifest.semanticFingerprint)) {
        status = "Duplicate";
        message = "Same period and semantic fingerprint; no cache change.";
      } else {
        status = "Correction Review";
        message = `Period ${text(manifest.sourcePeriodKey)} already exists with different content; explicit supersession is required.`;
      }
    }
  }
  if (text(payload.expectedOutcome) && text(payload.expectedOutcome) !== status) {
    throw new Error(`PUL-030WIA-101: Runtime outcome ${status} differs from accepted publisher ${text(payload.expectedOutcome)}.`);
  }
  if (status === "New") {
    return result(status, text(manifest.sourcePeriodKey), message, false,
      authority.cacheVersion, "", "", "Not written", false, "Process");
  }
  const entry = ledgerEntryForOutcome(
    status, message, parsed, text(payload.parseError), authority.cacheVersion,
    authority.identityPreflightFingerprint, "", text(payload.processedAt),
    payload.expectedLedgerEntry
  );
  const ledgerAction = upsertLedger(workbook, entry);
  return result(status, text(entry.sourcePeriodKey), message, false,
    authority.cacheVersion, "", text(entry.intakeEventId), ledgerAction,
    status === "Duplicate", "Process");
}

function prepareNewCandidate(workbook: ExcelScript.Workbook, payload: IntakePayload): WeeklyIntakeResult {
  validateCandidateEnvelope(payload);
  const authority = requireFreshAuthority(workbook, payload);
  const layout = candidateLayout(payload);
  const existing = workbook.getWorksheet(CANDIDATE_SHEET);
  if (existing) {
    validateCandidateControl(existing, payload);
    return result("New", "", "Candidate surface is already prepared.", false,
      authority.cacheVersion, text(payload.candidateVersion), "", "Not written", false, "Prepare New");
  }
  const sheet = workbook.addWorksheet(CANDIDATE_SHEET);
  sheet.setVisibility(ExcelScript.SheetVisibility.hidden);
  writeCandidateHeaders(sheet, layout);
  writeCandidateControl(sheet, payload, "Prepared");
  sheet.getUsedRange().getFormat().getFont().setName("Aptos");
  sheet.getUsedRange().getFormat().getFont().setSize(9);
  return result("New", "", "Candidate surface prepared; no Active cache change.", false,
    authority.cacheVersion, text(payload.candidateVersion), "", "Not written", false, "Prepare New");
}

function writeNewCandidateChunk(workbook: ExcelScript.Workbook, payload: IntakePayload): WeeklyIntakeResult {
  validateCandidateEnvelope(payload);
  const authority = requireFreshAuthority(workbook, payload);
  const sheet = requiredCandidateSheet(workbook);
  validateCandidateControl(sheet, payload);
  const layout = candidateLayout(payload);
  const section = candidateSection(layout, text(payload.section));
  if (section.name === "version") throw new Error("PUL-030WIA-102: Version manifest is written only by Finalize New.");
  const values = payload.values || [];
  const offset = integer(payload.offset, "offset");
  const totalRows = integer(payload.totalRowCount, "totalRowCount");
  const startRow = integer(payload.startRow, "startRow");
  const startColumn = integer(payload.startColumn, "startColumn");
  if (totalRows !== section.rowCount || startRow !== section.startRow + 1 + offset ||
      startColumn !== section.startColumn || offset + values.length > section.rowCount ||
      text(payload.chunkFingerprint) !== chunkFingerprint(payload)) {
    throw new Error(`PUL-030WIA-103: ${section.name} Candidate chunk contract is invalid.`);
  }
  validateChunkRows(section, values, text(payload.candidateVersion));
  if (values.length) {
    const target = checkedRange(sheet, startRow, startColumn, values.length,
      section.columnCount, `${section.name} Candidate chunk`);
    const prior = target.getValues();
    let blankRows = 0;
    let equalRows = 0;
    for (let index = 0; index < prior.length; index += 1) {
      if (blankRow(prior[index])) blankRows += 1;
      else if (JSON.stringify(prior[index]) === JSON.stringify(values[index])) equalRows += 1;
      else throw new Error(`PUL-030WIA-104: ${section.name} Candidate chunk would overwrite different data.`);
    }
    if (blankRows && equalRows) throw new Error("PUL-030WIA-105: Candidate chunk is partially written.");
    if (blankRows === prior.length) target.setValues(values);
  }
  return result("New", "", `Candidate ${section.name} chunk ${offset}+${values.length} accepted.`,
    true, authority.cacheVersion, text(payload.candidateVersion), "", "Not written", false, "Write New");
}

function finalizeNewCandidate(workbook: ExcelScript.Workbook, payload: IntakePayload): WeeklyIntakeResult {
  validateCandidateEnvelope(payload);
  const authority = requireFreshAuthority(workbook, payload);
  const sheet = requiredCandidateSheet(workbook);
  validateCandidateControl(sheet, payload);
  const layout = candidateLayout(payload);
  const versionValues = payload.versionValues || [];
  if (versionValues.length !== VERSION_HEADERS.length ||
      text(versionValues[0]) !== text(payload.candidateVersion) ||
      text(versionValues[2]) !== "Candidate" || text(versionValues[3]) !== "Not Active" ||
      text(versionValues[4]) !== "PASS" ||
      number(versionValues[17]) !== layout.period.rowCount ||
      number(versionValues[18]) !== layout.scope.rowCount ||
      number(versionValues[19]) !== layout.rpg.rowCount ||
      text(versionValues[21]) !== text(payload.candidateFingerprint)) {
    throw new Error("PUL-030WIA-106: Candidate version manifest is invalid.");
  }
  const existingVersionTable = workbook.getTable("tblWeeklyIntakeCandidateVersions");
  if (existingVersionTable) {
    const existingVersion = tableRows(existingVersionTable)[0];
    if (!existingVersion || JSON.stringify(existingVersion) !== JSON.stringify(versionValues)) {
      throw new Error("PUL-030WIA-107: Finalized Candidate version differs from the request.");
    }
    const existingEntry = requiredLedgerEntry(payload.ledgerEntry);
    const ledgerAction = upsertLedger(workbook, existingEntry);
    return result("New", text(existingEntry.sourcePeriodKey), "Candidate is already finalized and remains Not Active.",
      false, authority.cacheVersion, text(payload.candidateVersion), text(existingEntry.intakeEventId),
      ledgerAction, false, "Finalize New");
  }
  assertCandidateHeaders(sheet, layout);
  const periodRows = candidateValues(sheet, layout.period);
  const scopeRows = candidateValues(sheet, layout.scope);
  const rpgRows = candidateValues(sheet, layout.rpg);
  validateCompleteCandidateRows(text(payload.candidateVersion), periodRows, scopeRows, rpgRows);
  const fingerprint = fingerprintCache(versionValues, periodRows, scopeRows, rpgRows);
  if (fingerprint !== text(payload.candidateFingerprint)) {
    throw new Error(`PUL-030WIA-108: Candidate fingerprint ${fingerprint} differs from ${text(payload.candidateFingerprint)}.`);
  }
  checkedRange(sheet, 2, 1, 1, VERSION_HEADERS.length, "Candidate version manifest")
    .setValues([versionValues]);
  createCandidateTables(sheet, layout);
  writeCandidateControl(sheet, payload, "Finalized / Not Active");
  sheet.setVisibility(ExcelScript.SheetVisibility.hidden);
  const entry = requiredLedgerEntry(payload.ledgerEntry);
  if (text(entry.resultingCacheVersion) !== text(payload.candidateVersion) ||
      text(entry.intakeStatus) !== "New") {
    throw new Error("PUL-030WIA-109: New Candidate ledger entry is invalid.");
  }
  const ledgerAction = upsertLedger(workbook, entry);
  const activeAfter = resolveActiveAuthority(workbook);
  if (!activeAfter.available || activeAfter.cacheVersion !== authority.cacheVersion ||
      activeAfter.cacheFingerprint !== authority.cacheFingerprint) {
    throw new Error("PUL-030WIA-110: Active cache changed while finalizing the inactive Candidate.");
  }
  return result("New", text(entry.sourcePeriodKey),
    `Candidate ${text(payload.candidateVersion)} finalized as Not Active.`, true,
    authority.cacheVersion, text(payload.candidateVersion), text(entry.intakeEventId),
    ledgerAction, false, "Finalize New");
}

function resolveActiveAuthority(workbook: ExcelScript.Workbook): Authority {
  const versions = requiredTable(workbook, "tblWeeklyCacheVersions");
  const vh = headerMap(versions);
  const rows = tableRows(versions);
  const active: CellValue[][] = [];
  for (const row of rows) {
    if (text(row[vh.CacheStatus]) === "Active" && text(row[vh.ActivationState]) === "Active") active.push(row);
  }
  if (active.length !== 1) return unavailableAuthority(`Weekly cache requires exactly one Active / Active version; found ${active.length}.`);
  const row = active[0];
  const version = text(row[vh.CacheVersion]);
  const cacheFingerprint = text(row[vh.CacheFingerprint]);
  const identityPreflightFingerprint = text(row[vh.IdentityPreflightFingerprint]);
  if (text(row[vh.ValidationStatus]) !== "PASS" || !text(row[vh.CacheFingerprint]) ||
      !text(row[vh.MappingContentFingerprint]) || !text(row[vh.CatalogContentFingerprint]) ||
      !text(row[vh.IdentityPreflightFingerprint]) || !text(row[vh.PerformanceRestaurantScopeFingerprint])) {
    return unavailableAuthority(`Active cache ${version} is not validated/fingerprinted.`,
      version, cacheFingerprint, identityPreflightFingerprint);
  }
  const periodTable = requiredTable(workbook, "tblWeeklyPeriodManifest");
  const periodRows = tableRows(periodTable); const ph = headerMap(periodTable);
  const scopeRows = tableRows(requiredTable(workbook, "tblWeeklyScopeCache"));
  const rpgRows = tableRows(requiredTable(workbook, "tblWeeklyRPGCache"));
  let periods = 0; let scopes = 0; let rpgs = 0;
  const seenPeriods: { [key: string]: boolean } = {};
  for (const period of periodRows) {
    if (text(period[ph.CacheVersion]) !== version) continue;
    periods += 1;
    const key = `${number(period[ph.ISOYear])}|${number(period[ph.ISOWeek])}`;
    if (seenPeriods[key]) return unavailableAuthority(`Active cache repeats ISO period ${key}.`,
      version, cacheFingerprint, identityPreflightFingerprint);
    seenPeriods[key] = true;
  }
  for (const scope of scopeRows) if (text(scope[1]) === version) scopes += 1;
  for (const rpg of rpgRows) if (text(rpg[1]) === version) rpgs += 1;
  if (periods !== number(row[vh.PeriodRowCount]) || scopes !== number(row[vh.ScopeCacheRowCount]) ||
      rpgs !== number(row[vh.DenseRPGCacheRowCount])) {
    return unavailableAuthority(`Active cache ${version} row counts differ from its manifest.`,
      version, cacheFingerprint, identityPreflightFingerprint);
  }
  const calc = requiredSheet(workbook, "_Metric_Calc");
  if (text(calc.getRange("AL16").getValue()) !== "Available" ||
      text(calc.getRange("AL24").getValue()) !== version ||
      text(calc.getRange("AL27").getValue()) !== text(row[vh.MappingContentFingerprint]) ||
      text(calc.getRange("AL28").getValue()) !== text(row[vh.CatalogContentFingerprint]) ||
      text(calc.getRange("AL29").getValue()) !== text(row[vh.PerformanceRestaurantScopeFingerprint]) ||
      text(calc.getRange("AL30").getValue()) !== text(row[vh.IdentityPreflightFingerprint])) {
    return unavailableAuthority(`Active cache ${version} freshness guard is stale / unavailable.`,
      version, cacheFingerprint, identityPreflightFingerprint);
  }
  if (qaPassCount(workbook, "tblWeeklyPerformanceQA") !== 16 ||
      qaPassCount(workbook, "tblPerformanceInteractionQA") !== 16) {
    return unavailableAuthority("Weekly Performance or Phase 2C QA is not 16/16 PASS.",
      version, cacheFingerprint, identityPreflightFingerprint);
  }
  return {
    available: true,
    cacheVersion: version,
    cacheFingerprint,
    identityPreflightFingerprint,
    message: "Available"
  };
}

function requireFreshAuthority(workbook: ExcelScript.Workbook, payload: IntakePayload): Authority {
  const authority = resolveActiveAuthority(workbook);
  if (!authority.available) throw new Error(`PUL-030WIA-111: ${authority.message}`);
  if (text(payload.priorCacheVersion) !== authority.cacheVersion) {
    throw new Error(`PUL-030WIA-112: Candidate prior cache ${text(payload.priorCacheVersion)} differs from Active ${authority.cacheVersion}.`);
  }
  return authority;
}

function activePeriodRows(workbook: ExcelScript.Workbook, cacheVersion: string, periodKey: string): CellValue[][] {
  const table = requiredTable(workbook, "tblWeeklyPeriodManifest");
  const h = headerMap(table); const matches: CellValue[][] = [];
  const rows = tableRows(table);
  for (const row of rows) {
    if (text(row[h.CacheVersion]) === cacheVersion && text(row[h.SourcePeriodKey]) === periodKey) matches.push(row);
  }
  return matches;
}

function ensureLedger(workbook: ExcelScript.Workbook): ExcelScript.Table {
  const sheet = requiredSheet(workbook, CACHE_SHEET);
  if (sheet.getVisibility() !== ExcelScript.SheetVisibility.hidden) {
    throw new Error("PUL-030WIA-113: Weekly cache engineering surface must remain hidden.");
  }
  const existing = workbook.getTable(LEDGER_TABLE);
  if (existing) {
    assertTableHeaders(existing, LEDGER_HEADERS, LEDGER_TABLE);
    return existing;
  }
  const target = sheet.getRange(LEDGER_RANGE);
  if (!blankMatrix(target.getValues())) {
    throw new Error(`PUL-030WIA-114: ${LEDGER_RANGE} is not empty for the intake ledger.`);
  }
  target.setValues([LEDGER_HEADERS, blankValues(LEDGER_HEADERS.length)]);
  const table = sheet.addTable(target, true);
  table.setName(LEDGER_TABLE);
  table.setPredefinedTableStyle("TableStyleMedium2");
  sheet.getRange("Y:AL").getFormat().setColumnWidth(120);
  sheet.getRange("Y1:AL1").getFormat().setWrapText(true);
  table.getColumnByName("SourceSalesNOK").getRange().setNumberFormat("#,##0.00");
  return table;
}

function upsertLedger(workbook: ExcelScript.Workbook, entry: LedgerEntry): string {
  const table = ensureLedger(workbook);
  const values = ledgerValues(entry);
  const rows = tableRows(table);
  let blankIndex = -1;
  for (let index = 0; index < rows.length; index += 1) {
    if (!text(rows[index][0]) && blankIndex < 0) blankIndex = index;
    if (text(rows[index][0]) !== text(entry.intakeEventId)) continue;
    const invariantIndexes = [0, 2, 3, 4, 5, 6, 8, 9, 11, 12, 13];
    for (const column of invariantIndexes) {
      if (text(rows[index][column]) !== text(values[column])) {
        throw new Error(`PUL-030WIA-115: Existing ledger event ${text(entry.intakeEventId)} differs at ${LEDGER_HEADERS[column]}.`);
      }
    }
    return "Existing";
  }
  if (blankIndex >= 0) table.getRangeBetweenHeaderAndTotal().getCell(blankIndex, 0)
    .getResizedRange(0, LEDGER_HEADERS.length - 1).setValues([values]);
  else table.addRow(-1, values);
  return "Added";
}

function ledgerEntryForOutcome(
  status: string,
  message: string,
  parsed: ParsedReportMetadata | undefined,
  parseError: string,
  activeCacheVersion: string,
  identityPreflightFingerprint: string,
  resultingCacheVersion: string,
  processedAt: string,
  expected: LedgerEntry | undefined
): LedgerEntry {
  const manifest = parsed && parsed.manifest ? parsed.manifest : {};
  const locator = text(manifest.sourceLocator);
  const period = text(manifest.sourcePeriodKey);
  const semantic = text(manifest.semanticFingerprint);
  const eventKey: CellValue[] = semantic ? [period, semantic, status, resultingCacheVersion] :
    [locator, status, message || parseError];
  const entry: LedgerEntry = {
    intakeEventId: hashStrings([record("INTAKE", eventKey)], "WINT-"),
    sourceLocator: locator,
    sourceFileId: text(manifest.sourceFileId),
    sourcePeriodKey: period,
    sourceSemanticFingerprint: semantic,
    identityPreflightFingerprint: status === "Rejected" ? "" : identityPreflightFingerprint,
    intakeStatus: status,
    statusMessage: message,
    sourceRowCount: number(manifest.sourceRowCount),
    sourceSalesNok: round(number(manifest.totalSalesNok), 2),
    processedAt: processedAt || new Date().toISOString(),
    priorCacheVersion: activeCacheVersion,
    resultingCacheVersion,
    supersededCacheVersion: ""
  };
  if (expected && text(expected.intakeEventId)) {
    const expectedValues = ledgerValues(expected); const actualValues = ledgerValues(entry);
    const compareIndexes = [0, 2, 3, 4, 6, 8, 9, 11, 12, 13];
    for (const index of compareIndexes) {
      if (text(expectedValues[index]) !== text(actualValues[index])) {
        throw new Error(`PUL-030WIA-116: Accepted ledger entry differs at ${LEDGER_HEADERS[index]}.`);
      }
    }
    entry.identityPreflightFingerprint = text(expected.identityPreflightFingerprint);
  }
  return entry;
}

function requiredLedgerEntry(value: LedgerEntry | undefined): LedgerEntry {
  if (!value || !text(value.intakeEventId) || !text(value.intakeStatus)) {
    throw new Error("PUL-030WIA-117: Finalize New requires the accepted publisher ledger entry.");
  }
  return value;
}

function ledgerValues(entry: LedgerEntry): CellValue[] {
  return [
    text(entry.intakeEventId), text(entry.sourceLocator), text(entry.sourceFileId),
    text(entry.sourcePeriodKey), text(entry.sourceSemanticFingerprint),
    text(entry.identityPreflightFingerprint), text(entry.intakeStatus),
    text(entry.statusMessage), number(entry.sourceRowCount), round(number(entry.sourceSalesNok), 2),
    text(entry.processedAt), text(entry.priorCacheVersion), text(entry.resultingCacheVersion),
    text(entry.supersededCacheVersion)
  ];
}

function validateAdapterPayload(payload: IntakePayload): void {
  if (text(payload.adapterVersion) !== ADAPTER_VERSION ||
      text(payload.publisherVersion) !== PUBLISHER_VERSION) {
    throw new Error("PUL-030WIA-118: Adapter/publisher payload version is unsupported.");
  }
}

function validateCandidateEnvelope(payload: IntakePayload): void {
  validateAdapterPayload(payload);
  if (!/^WCV-[0-9a-f]{16}$/.test(text(payload.candidateVersion)) ||
      !/^WCC-[0-9a-f]{16}$/.test(text(payload.candidateFingerprint)) ||
      !/^WLAY-[0-9a-f]{16}$/.test(text(payload.layoutFingerprint))) {
    throw new Error("PUL-030WIA-119: Candidate identifiers are invalid.");
  }
}

function parserMetadataErrors(parsed: ParsedReportMetadata | undefined): string[] {
  if (!parsed || !parsed.manifest) return ["Parsed weekly report metadata is missing."];
  const errors: string[] = []; const manifest = parsed.manifest;
  if (text(parsed.parserVersion) !== PARSER_VERSION || text(manifest.parserVersion) !== PARSER_VERSION) {
    errors.push(`ParserVersion must be ${PARSER_VERSION}.`);
  }
  if (text(parsed.schemaVersion) !== SCHEMA_VERSION || text(manifest.schemaVersion) !== SCHEMA_VERSION) {
    errors.push(`SchemaVersion must be ${SCHEMA_VERSION}.`);
  }
  if (text(manifest.contentReconciliationStatus) !== "PASS") errors.push("Source report reconciliation is not PASS.");
  if (!text(manifest.sourceFileId) || !text(manifest.semanticFingerprint) || !text(manifest.sourcePeriodKey)) {
    errors.push("SourceFileID, semantic fingerprint and period are required.");
  }
  if (!Number.isInteger(number(manifest.sourceRowCount)) || number(manifest.sourceRowCount) < 1 ||
      !Number.isFinite(Number(manifest.totalSalesNok))) {
    errors.push("Source row count or Sales NOK is invalid.");
  }
  return errors;
}

function candidateLayout(payload: IntakePayload): CandidateLayout {
  const counts = payload.counts || {};
  const periodCount = positiveInteger(counts.period, "period count");
  const scopeCount = positiveInteger(counts.scope, "scope count");
  const rpgCount = positiveInteger(counts.rpg, "RPG count");
  const version = section("version", "tblWeeklyIntakeCandidateVersions", 1, 1, VERSION_HEADERS);
  const period = section("period", "tblWeeklyIntakeCandidatePeriodManifest", 5, periodCount, PERIOD_HEADERS);
  const scopeStart = period.startRow + period.rowCount + 3;
  const scope = section("scope", "tblWeeklyIntakeCandidateScopeCache", scopeStart, scopeCount, SCOPE_HEADERS);
  const rpgStart = scope.startRow + scope.rowCount + 3;
  const rpg = section("rpg", "tblWeeklyIntakeCandidateRPGCache", rpgStart, rpgCount, RPG_HEADERS);
  const layout = { version, period, scope, rpg };
  if (candidateLayoutFingerprint(layout) !== text(payload.layoutFingerprint)) {
    throw new Error("PUL-030WIA-120: Candidate layout fingerprint is invalid.");
  }
  return layout;
}

function section(name: string, tableName: string, startRow: number, rowCount: number,
  headers: string[]): CandidateSection {
  return { name, tableName, startRow, startColumn: 1, rowCount, columnCount: headers.length, headers };
}

function candidateLayoutFingerprint(layout: CandidateLayout): string {
  const values = [layout.version, layout.period, layout.scope, layout.rpg];
  const records: string[] = [];
  for (const value of values) {
    const fields: unknown[] = [
      value.name, value.tableName, value.startRow, value.startColumn,
      value.rowCount, value.columnCount
    ];
    for (const header of value.headers) fields.push(header);
    records.push(record("SECTION", fields));
  }
  return hashStrings(records, "WLAY-");
}

function chunkFingerprint(payload: IntakePayload): string {
  return hashStrings([record("CHUNK", [
    payload.candidateVersion, payload.section, payload.offset, payload.totalRowCount,
    payload.startRow, payload.startColumn, JSON.stringify(payload.values || [])
  ])], "WCHK-");
}

function writeCandidateHeaders(sheet: ExcelScript.Worksheet, layout: CandidateLayout): void {
  writeCandidateHeader(sheet, layout.version);
  writeCandidateHeader(sheet, layout.period);
  writeCandidateHeader(sheet, layout.scope);
  writeCandidateHeader(sheet, layout.rpg);
}

function assertCandidateHeaders(sheet: ExcelScript.Worksheet, layout: CandidateLayout): void {
  assertCandidateHeader(sheet, layout.version);
  assertCandidateHeader(sheet, layout.period);
  assertCandidateHeader(sheet, layout.scope);
  assertCandidateHeader(sheet, layout.rpg);
}

function writeCandidateHeader(sheet: ExcelScript.Worksheet, value: CandidateSection): void {
  checkedRange(sheet, value.startRow, value.startColumn, 1, value.columnCount,
    `${value.name} Candidate header`).setValues([value.headers]);
}

function assertCandidateHeader(sheet: ExcelScript.Worksheet, value: CandidateSection): void {
  const actual = checkedRange(sheet, value.startRow, value.startColumn, 1,
    value.columnCount, `${value.name} Candidate header`).getTexts()[0];
  if (actual.join("|") !== value.headers.join("|")) {
    throw new Error(`PUL-030WIA-121: ${value.name} Candidate headers differ.`);
  }
}

function writeCandidateControl(sheet: ExcelScript.Worksheet, payload: IntakePayload, state: string): void {
  sheet.getRange("Y1:Z6").setValues([
    ["Weekly intake Candidate", "Value"],
    ["CandidateVersion", text(payload.candidateVersion)],
    ["CandidateFingerprint", text(payload.candidateFingerprint)],
    ["PriorCacheVersion", text(payload.priorCacheVersion)],
    ["LayoutFingerprint", text(payload.layoutFingerprint)],
    ["State", state]
  ]);
}

function validateCandidateControl(sheet: ExcelScript.Worksheet, payload: IntakePayload): void {
  if (sheet.getVisibility() !== ExcelScript.SheetVisibility.hidden) {
    throw new Error("PUL-030WIA-122: Candidate engineering surface must remain hidden.");
  }
  const values = sheet.getRange("Y1:Z6").getTexts();
  if (values[1][1] !== text(payload.candidateVersion) ||
      values[2][1] !== text(payload.candidateFingerprint) ||
      values[3][1] !== text(payload.priorCacheVersion) ||
      values[4][1] !== text(payload.layoutFingerprint)) {
    throw new Error("PUL-030WIA-123: Existing Candidate control differs from the request.");
  }
}

function validateChunkRows(sectionValue: CandidateSection, rows: CellValue[][], cacheVersion: string): void {
  for (const row of rows) {
    if (row.length !== sectionValue.columnCount || text(row[1]) !== cacheVersion) {
      throw new Error(`PUL-030WIA-124: ${sectionValue.name} Candidate row schema/version is invalid.`);
    }
    let expectedId = "";
    if (sectionValue.name === "period") expectedId = stableId("WPER-", [cacheVersion, row[2]]);
    else if (sectionValue.name === "scope") expectedId = stableId("WSCP-", [cacheVersion, row[2], row[3]]);
    else expectedId = stableId("WRPG-", [cacheVersion, row[2], row[3], row[4]]);
    if (text(row[0]) !== expectedId) throw new Error(`PUL-030WIA-125: ${sectionValue.name} Candidate row ID is invalid.`);
  }
}

function validateCompleteCandidateRows(
  cacheVersion: string,
  periodRows: CellValue[][],
  scopeRows: CellValue[][],
  rpgRows: CellValue[][]
): void {
  const periodIds: { [key: string]: boolean } = {};
  for (const row of periodRows) {
    validateChunkRows(section("period", "", 1, 1, PERIOD_HEADERS), [row], cacheVersion);
    const grain = text(row[2]);
    if (periodIds[grain]) throw new Error(`PUL-030WIA-126: Candidate repeats period ${grain}.`);
    periodIds[grain] = true;
  }
  const scopeIds: { [key: string]: boolean } = {};
  for (const row of scopeRows) {
    validateChunkRows(section("scope", "", 1, 1, SCOPE_HEADERS), [row], cacheVersion);
    const grain = `${text(row[2])}|${text(row[3])}`;
    if (scopeIds[grain]) throw new Error(`PUL-030WIA-127: Candidate repeats scope ${grain}.`);
    scopeIds[grain] = true;
  }
  const rpgIds: { [key: string]: boolean } = {};
  for (const row of rpgRows) {
    validateChunkRows(section("rpg", "", 1, 1, RPG_HEADERS), [row], cacheVersion);
    const grain = `${text(row[2])}|${text(row[3])}|${text(row[4])}`;
    if (rpgIds[grain]) throw new Error(`PUL-030WIA-128: Candidate repeats RPG grain ${grain}.`);
    rpgIds[grain] = true;
  }
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

function createCandidateTables(sheet: ExcelScript.Worksheet, layout: CandidateLayout): void {
  createCandidateTable(sheet, layout.version);
  createCandidateTable(sheet, layout.period);
  createCandidateTable(sheet, layout.scope);
  createCandidateTable(sheet, layout.rpg);
}

function createCandidateTable(sheet: ExcelScript.Worksheet, value: CandidateSection): void {
  const range = checkedRange(sheet, value.startRow, value.startColumn,
    value.rowCount + 1, value.columnCount, `${value.name} Candidate table`);
  const table = sheet.addTable(range, true);
  table.setName(value.tableName);
  table.setPredefinedTableStyle("TableStyleMedium2");
}

function candidateValues(sheet: ExcelScript.Worksheet, sectionValue: CandidateSection): CellValue[][] {
  const values = checkedRange(sheet, sectionValue.startRow + 1, sectionValue.startColumn,
    sectionValue.rowCount, sectionValue.columnCount, `${sectionValue.name} Candidate values`).getValues();
  for (const row of values) if (blankRow(row)) {
    throw new Error(`PUL-030WIA-129: ${sectionValue.name} Candidate contains an unwritten row.`);
  }
  return values;
}

function candidateSection(layout: CandidateLayout, name: string): CandidateSection {
  if (name === "version") return layout.version;
  if (name === "period") return layout.period;
  if (name === "scope") return layout.scope;
  if (name === "rpg") return layout.rpg;
  throw new Error(`PUL-030WIA-130: Unsupported Candidate section ${name}.`);
}

function qaPassCount(workbook: ExcelScript.Workbook, tableName: string): number {
  const table = requiredTable(workbook, tableName); const h = headerMap(table);
  const rows = tableRows(table);
  let count = 0;
  for (const row of rows) if (text(row[h.Result]) === "PASS") count += 1;
  return count;
}

function assertTableHeaders(table: ExcelScript.Table, expected: string[], label: string): void {
  const actual = table.getHeaderRowRange().getTexts()[0];
  if (actual.join("|") !== expected.join("|")) throw new Error(`PUL-030WIA-131: ${label} headers differ.`);
}

function result(
  status: string, period: string, message: string, cacheChanged: boolean,
  activeCacheVersion: string, resultingCacheVersion: string, ledgerEventId: string,
  ledgerAction: string, archiveReady: boolean, operation: string
): WeeklyIntakeResult {
  return {
    status, period, message, cacheChanged, activeCacheVersion,
    resultingCacheVersion, ledgerEventId, ledgerAction, archiveReady, operation
  };
}

function unavailableAuthority(
  message: string,
  cacheVersion: string = "",
  cacheFingerprint: string = "",
  identityPreflightFingerprint: string = ""
): Authority {
  return { available: false, cacheVersion, cacheFingerprint, identityPreflightFingerprint, message };
}

function requiredCandidateSheet(workbook: ExcelScript.Workbook): ExcelScript.Worksheet {
  const sheet = workbook.getWorksheet(CANDIDATE_SHEET);
  if (!sheet) throw new Error("PUL-030WIA-132: Candidate surface is missing. Run Prepare New first.");
  return sheet;
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-030WIA-133: Required table ${name} is missing.`);
  return table;
}

function requiredSheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  const sheet = workbook.getWorksheet(name);
  if (!sheet) throw new Error(`PUL-030WIA-134: Required sheet ${name} is missing.`);
  return sheet;
}

function tableRows(table: ExcelScript.Table): CellValue[][] {
  return table.getRangeBetweenHeaderAndTotal().getValues();
}

function headerMap(table: ExcelScript.Table): { [key: string]: number } {
  const headers = table.getHeaderRowRange().getTexts()[0]; const result: { [key: string]: number } = {};
  for (let index = 0; index < headers.length; index += 1) result[text(headers[index])] = index;
  return result;
}

function parsePayload(value: string): IntakePayload {
  try { return JSON.parse(value || "{}") as IntakePayload; }
  catch (error) { throw new Error(`PUL-030WIA-135: payloadJson is invalid JSON. ${String(error)}`); }
}

function checkedRange(
  sheet: ExcelScript.Worksheet, startRow: number, startColumn: number,
  rowCount: number, columnCount: number, label: string
): ExcelScript.Range {
  if (startRow < 1 || startColumn < 1 || rowCount < 1 || columnCount < 1 ||
      startRow + rowCount - 1 > 1048576 || startColumn + columnCount - 1 > 16384) {
    throw new Error(`PUL-030WIA-136: Invalid ${label} geometry.`);
  }
  return sheet.getRangeByIndexes(startRow - 1, startColumn - 1, rowCount, columnCount);
}

function blankMatrix(values: CellValue[][]): boolean {
  for (const row of values) if (!blankRow(row)) return false;
  return true;
}

function blankRow(row: CellValue[]): boolean {
  for (const value of row) if (text(value)) return false;
  return true;
}

function blankValues(count: number): string[] {
  const values: string[] = [];
  for (let index = 0; index < count; index += 1) values.push("");
  return values;
}

function stableId(prefix: string, values: unknown[]): string {
  return hashStrings([record("ID", values)], prefix);
}

function canonicalIsoDate(value: unknown): string {
  const raw = text(value);
  if (/^\d{4}-\d{2}-\d{2}$/.test(raw)) return raw;
  const serial = number(value); if (!serial) return raw;
  const date = new Date(Math.round((serial - 25569) * 86400000));
  return date.toISOString().slice(0, 10);
}

function record(kind: string, values: unknown[]): string {
  return `${kind}|${values.map(value => {
    const normalized = value === null || value === undefined ? "" : String(value);
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

function positiveInteger(value: unknown, label: string): number {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 1) throw new Error(`PUL-030WIA-137: ${label} must be positive.`);
  return parsed;
}

function integer(value: unknown, label: string): number {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 0) throw new Error(`PUL-030WIA-138: ${label} must be non-negative.`);
  return parsed;
}

function sales(value: unknown): string { return number(value).toFixed(2); }
function quantity(value: unknown): string { return number(value).toFixed(6); }
function round(value: number, decimals: number): number {
  const factor = Math.pow(10, decimals);
  return Math.round((value + Number.EPSILON) * factor) / factor;
}
function number(value: unknown): number { const parsed = Number(value); return Number.isFinite(parsed) ? parsed : 0; }
function text(value: unknown): string { return String(value === null || value === undefined ? "" : value).trim(); }
