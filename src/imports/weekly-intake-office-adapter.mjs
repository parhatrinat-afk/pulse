import {
  WEEKLY_CACHE_COLUMNS,
} from "./weekly-cache-materialization.mjs";
import {
  validateActiveWeeklyCacheFreshness,
} from "./weekly-cache-activation.mjs";
import {
  WEEKLY_INTAKE_LOG_COLUMNS,
  WEEKLY_INTAKE_OUTCOMES,
  WEEKLY_INTAKE_PUBLISHER_VERSION,
} from "./weekly-intake-publisher.mjs";
import {
  WEEKLY_SALES_PARSER_VERSION,
  WEEKLY_SALES_SCHEMA_VERSION,
} from "./weekly-sales-parser.mjs";

export const WEEKLY_INTAKE_ADAPTER_VERSION = "0.3.0-weekly-intake-adapter-v1";
export const WEEKLY_INTAKE_LEDGER_TABLE = "tblWeeklyIntakeLog";
export const WEEKLY_INTAKE_LEDGER_SHEET = "_Weekly_Cache";
export const WEEKLY_INTAKE_LEDGER_ADDRESS = "Y1:AL2";
export const WEEKLY_INTAKE_CANDIDATE_SHEET = "_Weekly_Cache_Candidate";

export const WEEKLY_INTAKE_CANDIDATE_TABLES = Object.freeze({
  version: "tblWeeklyIntakeCandidateVersions",
  period: "tblWeeklyIntakeCandidatePeriodManifest",
  scope: "tblWeeklyIntakeCandidateScopeCache",
  rpg: "tblWeeklyIntakeCandidateRPGCache",
});

const LAYOUT_GAP_ROWS = 2;

/**
 * Mirror the bounded runtime classification performed by the Office Script.
 * It deliberately stops before identity resolution, mapping, or cache math.
 */
export function classifyWeeklyIntakeMetadata({
  parsedReport,
  parseError = "",
  versionManifests,
  periodManifest,
  currentFreshness,
}) {
  if (String(parseError)) {
    return classification("Rejected", String(parseError));
  }
  const validation = validateParserMetadata(parsedReport);
  if (validation.length) return classification("Rejected", validation.join(" "));
  const freshness = validateActiveWeeklyCacheFreshness({
    versionManifests,
    current: currentFreshness,
  });
  if (freshness.status !== "Available") {
    return classification("Cache Stale", freshness.errors.join(" "), freshness.activeVersion ?? "");
  }
  const manifest = parsedReport.manifest;
  const existing = (periodManifest ?? []).filter(row =>
    row.cacheVersion === freshness.activeVersion &&
    row.sourcePeriodKey === manifest.sourcePeriodKey);
  if (existing.length > 1) {
    return classification(
      "Cache Stale",
      `Active cache repeats period ${manifest.sourcePeriodKey}.`,
      freshness.activeVersion,
    );
  }
  if (!existing.length) {
    return classification(
      "New",
      "Period is not present in the Active cache; an accepted publisher Candidate is required.",
      freshness.activeVersion,
    );
  }
  if (existing[0].sourceSemanticFingerprint === manifest.semanticFingerprint) {
    return classification(
      "Duplicate",
      "Same period and semantic fingerprint; no cache change.",
      freshness.activeVersion,
    );
  }
  return classification(
    "Correction Review",
    `Period ${manifest.sourcePeriodKey} already exists with different content; explicit supersession is required.`,
    freshness.activeVersion,
  );
}

/** Build the small Process request that Power Automate passes to Excel. */
export function buildWeeklyIntakeProcessPayload({
  parsedReport,
  parseError = "",
  processedAt = "",
  expectedOutcome = "",
  expectedLedgerEntry = null,
}) {
  if (expectedOutcome && !WEEKLY_INTAKE_OUTCOMES.includes(expectedOutcome)) {
    fail("PUL-030WIA-001", `Unsupported expected outcome ${expectedOutcome}.`);
  }
  return {
    adapterVersion: WEEKLY_INTAKE_ADAPTER_VERSION,
    publisherVersion: WEEKLY_INTAKE_PUBLISHER_VERSION,
    operation: "Process",
    parseError: String(parseError),
    processedAt: String(processedAt),
    expectedOutcome: String(expectedOutcome),
    expectedLedgerEntry: expectedLedgerEntry ? { ...expectedLedgerEntry } : null,
    parsedReport: compactParsedReport(parsedReport),
  };
}

/**
 * Convert the already-validated repository publisher result into the bounded
 * workbook materialization plan. No business classification is performed here.
 */
export function buildWeeklyIntakeCandidatePlan(publicationResult) {
  if (publicationResult?.outcome !== "New" || !publicationResult.candidatePrepared ||
      !publicationResult.candidate) {
    fail("PUL-030WIA-002", "Only an accepted New publisher result can create a Candidate plan.");
  }
  const candidate = publicationResult.candidate;
  const manifest = candidate.versionManifest;
  if (manifest.cacheStatus !== "Candidate" || manifest.activationState !== "Not Active" ||
      manifest.validationStatus !== "PASS" || candidate.validation?.status !== "PASS") {
    fail("PUL-030WIA-003", "Publisher Candidate must be Candidate / Not Active / PASS.");
  }
  const sources = {
    version: [manifest],
    period: candidate.periodManifest,
    scope: candidate.scopeCacheRows,
    rpg: candidate.weeklyRpgCacheRows,
  };
  const expectedCounts = {
    period: Number(manifest.periodRowCount),
    scope: Number(manifest.scopeCacheRowCount),
    rpg: Number(manifest.denseRpgCacheRowCount),
  };
  for (const name of ["period", "scope", "rpg"]) {
    if (!Array.isArray(sources[name]) || sources[name].length !== expectedCounts[name]) {
      fail("PUL-030WIA-004", `${name} Candidate rows differ from the version manifest.`);
    }
  }
  const layout = dynamicCandidateLayout(expectedCounts);
  const sections = {};
  for (const name of ["version", "period", "scope", "rpg"]) {
    const columns = WEEKLY_CACHE_COLUMNS[name];
    const values = sources[name].map(row => columns.map(([, key]) => row[key] ?? ""));
    sections[name] = {
      name,
      tableName: WEEKLY_INTAKE_CANDIDATE_TABLES[name],
      startRow: layout[name].startRow,
      startColumn: 1,
      headers: columns.map(([header]) => header),
      values,
      rowCount: values.length,
      columnCount: columns.length,
    };
  }
  const layoutFingerprint = candidateLayoutFingerprint(sections);
  return {
    adapterVersion: WEEKLY_INTAKE_ADAPTER_VERSION,
    publisherVersion: WEEKLY_INTAKE_PUBLISHER_VERSION,
    sheetName: WEEKLY_INTAKE_CANDIDATE_SHEET,
    candidateVersion: manifest.cacheVersion,
    candidateFingerprint: manifest.cacheFingerprint,
    priorCacheVersion: publicationResult.activeCacheVersion,
    layoutFingerprint,
    sections,
    ledgerEntry: { ...publicationResult.ledgerEntry },
    completeCandidateRows: 1 + expectedCounts.period + expectedCounts.scope + expectedCounts.rpg,
    status: "New",
    cacheStatus: "Candidate",
    activationState: "Not Active",
  };
}

export function weeklyIntakeCandidatePreparePayload(plan) {
  return candidateEnvelope(plan, "Prepare New", {
    counts: sectionCounts(plan),
  });
}

export function weeklyIntakeCandidateChunk(plan, sectionName, offset = 0, limit = 500) {
  const section = plan?.sections?.[sectionName];
  if (!section || sectionName === "version") {
    fail("PUL-030WIA-005", `Unsupported Candidate chunk section ${sectionName}.`);
  }
  if (!Number.isInteger(offset) || offset < 0 || offset > section.rowCount ||
      !Number.isInteger(limit) || limit < 1) {
    fail("PUL-030WIA-006", `Invalid Candidate chunk ${sectionName}/${offset}/${limit}.`);
  }
  const values = section.values.slice(offset, offset + limit);
  const payload = candidateEnvelope(plan, "Write New", {
    section: sectionName,
    offset,
    totalRowCount: section.rowCount,
    startRow: section.startRow + 1 + offset,
    startColumn: section.startColumn,
    values,
  });
  return {
    ...payload,
    chunkFingerprint: chunkFingerprint(payload),
  };
}

export function weeklyIntakeCandidateFinalizePayload(plan) {
  return candidateEnvelope(plan, "Finalize New", {
    counts: sectionCounts(plan),
    versionValues: plan.sections.version.values[0],
    ledgerEntry: { ...plan.ledgerEntry },
  });
}

export function candidateLayoutFingerprint(sections) {
  const records = ["version", "period", "scope", "rpg"].map(name => {
    const section = sections[name];
    return record("SECTION", [
      name,
      section.tableName,
      section.startRow,
      section.startColumn,
      section.rowCount,
      section.columnCount,
      ...section.headers,
    ]);
  });
  return hashStrings(records, "WLAY-");
}

export function chunkFingerprint(payload) {
  return hashStrings([record("CHUNK", [
    payload.candidateVersion,
    payload.section,
    payload.offset,
    payload.totalRowCount,
    payload.startRow,
    payload.startColumn,
    JSON.stringify(payload.values),
  ])], "WCHK-");
}

export function weeklyIntakeLedgerHeaders() {
  return WEEKLY_INTAKE_LOG_COLUMNS.map(([header]) => header);
}

function candidateEnvelope(plan, operation, values) {
  return {
    adapterVersion: plan.adapterVersion,
    publisherVersion: plan.publisherVersion,
    operation,
    candidateVersion: plan.candidateVersion,
    candidateFingerprint: plan.candidateFingerprint,
    priorCacheVersion: plan.priorCacheVersion,
    layoutFingerprint: plan.layoutFingerprint,
    ...values,
  };
}

function dynamicCandidateLayout(counts) {
  const version = { startRow: 1 };
  const period = { startRow: 5 };
  const scope = { startRow: period.startRow + counts.period + 1 + LAYOUT_GAP_ROWS };
  const rpg = { startRow: scope.startRow + counts.scope + 1 + LAYOUT_GAP_ROWS };
  return { version, period, scope, rpg };
}

function sectionCounts(plan) {
  return {
    period: plan.sections.period.rowCount,
    scope: plan.sections.scope.rowCount,
    rpg: plan.sections.rpg.rowCount,
  };
}

function compactParsedReport(report) {
  if (!report) return null;
  return {
    parserVersion: report.parserVersion,
    schemaVersion: report.schemaVersion,
    manifest: report.manifest ? { ...report.manifest } : null,
  };
}

function validateParserMetadata(report) {
  const errors = [];
  const manifest = report?.manifest;
  if (!manifest) return ["Parsed weekly report metadata is missing."];
  if (report.parserVersion !== WEEKLY_SALES_PARSER_VERSION ||
      manifest.parserVersion !== WEEKLY_SALES_PARSER_VERSION) {
    errors.push(`ParserVersion must be ${WEEKLY_SALES_PARSER_VERSION}.`);
  }
  if (report.schemaVersion !== WEEKLY_SALES_SCHEMA_VERSION ||
      manifest.schemaVersion !== WEEKLY_SALES_SCHEMA_VERSION) {
    errors.push(`SchemaVersion must be ${WEEKLY_SALES_SCHEMA_VERSION}.`);
  }
  if (manifest.contentReconciliationStatus !== "PASS") {
    errors.push("Source report reconciliation is not PASS.");
  }
  for (const field of ["sourceFileId", "semanticFingerprint", "sourcePeriodKey"]) {
    if (!String(manifest[field] ?? "").trim()) errors.push(`${field} is required.`);
  }
  if (!Number.isInteger(Number(manifest.sourceRowCount)) || Number(manifest.sourceRowCount) < 1 ||
      !Number.isFinite(Number(manifest.totalSalesNok))) {
    errors.push("Source row count or Sales NOK is invalid.");
  }
  return errors;
}

function classification(outcome, message, activeCacheVersion = "") {
  return {
    outcome,
    message,
    activeCacheVersion,
    cacheChanged: false,
    resultingCacheVersion: "",
  };
}

function record(kind, values) {
  return `${kind}|${values.map(value => {
    const normalized = value === null || value === undefined ? "" : String(value);
    return `${normalized.length}:${normalized}`;
  }).join("|")}`;
}

function hashStrings(values, prefix) {
  let left = 0;
  let right = 0;
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

function fail(code, message) {
  throw new Error(`${code}: ${message}`);
}
