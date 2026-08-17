import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import {
  WEEKLY_INTAKE_ADAPTER_VERSION,
  WEEKLY_INTAKE_CANDIDATE_SHEET,
  WEEKLY_INTAKE_CANDIDATE_TABLES,
  WEEKLY_INTAKE_LEDGER_ADDRESS,
  WEEKLY_INTAKE_LEDGER_SHEET,
  WEEKLY_INTAKE_LEDGER_TABLE,
  buildWeeklyIntakeCandidatePlan,
  buildWeeklyIntakeProcessPayload,
  candidateLayoutFingerprint,
  classifyWeeklyIntakeMetadata,
  weeklyIntakeCandidateChunk,
  weeklyIntakeCandidateFinalizePayload,
  weeklyIntakeCandidatePreparePayload,
  weeklyIntakeLedgerHeaders,
} from "../src/imports/weekly-intake-office-adapter.mjs";

const scriptUrl = new URL("../office-scripts/Process_Weekly_Intake.ts", import.meta.url);
const script = await readFile(scriptUrl, "utf8");

test("runtime metadata classification supports exactly the five accepted outcomes", () => {
  const parsed = report();
  const active = activeVersion();
  const current = freshness(active);
  const period = periodRow(parsed);

  assert.equal(classifyWeeklyIntakeMetadata({
    parsedReport: parsed,
    versionManifests: [active],
    periodManifest: [],
    currentFreshness: current,
  }).outcome, "New");
  assert.equal(classifyWeeklyIntakeMetadata({
    parsedReport: parsed,
    versionManifests: [active],
    periodManifest: [period],
    currentFreshness: current,
  }).outcome, "Duplicate");
  assert.equal(classifyWeeklyIntakeMetadata({
    parsedReport: changedReport(),
    versionManifests: [active],
    periodManifest: [period],
    currentFreshness: current,
  }).outcome, "Correction Review");
  assert.equal(classifyWeeklyIntakeMetadata({
    parsedReport: parsed,
    parseError: "PUL-030I-008: source header is invalid",
    versionManifests: [active],
    periodManifest: [period],
    currentFreshness: current,
  }).outcome, "Rejected");
  assert.equal(classifyWeeklyIntakeMetadata({
    parsedReport: parsed,
    versionManifests: [active],
    periodManifest: [period],
    currentFreshness: { ...current, mappingContentFingerprint: "MCF-changed" },
  }).outcome, "Cache Stale");
});

test("zero, multiple, failed or duplicate-grain active authorities never guess", () => {
  const parsed = report();
  const active = activeVersion();
  const current = freshness(active);
  assert.equal(classifyWeeklyIntakeMetadata({
    parsedReport: parsed, versionManifests: [], periodManifest: [], currentFreshness: current,
  }).outcome, "Cache Stale");
  assert.equal(classifyWeeklyIntakeMetadata({
    parsedReport: parsed,
    versionManifests: [active, { ...active, cacheVersion: "WCV-1111111111111111" }],
    periodManifest: [], currentFreshness: current,
  }).outcome, "Cache Stale");
  assert.equal(classifyWeeklyIntakeMetadata({
    parsedReport: parsed,
    versionManifests: [{ ...active, validationStatus: "FAIL" }],
    periodManifest: [], currentFreshness: current,
  }).outcome, "Cache Stale");
  const duplicatePeriod = periodRow(parsed);
  assert.equal(classifyWeeklyIntakeMetadata({
    parsedReport: parsed, versionManifests: [active],
    periodManifest: [duplicatePeriod, { ...duplicatePeriod }], currentFreshness: current,
  }).outcome, "Cache Stale");
});

test("Process payload contains parser metadata only and preserves publisher expectation", () => {
  const expectedLedgerEntry = ledgerEntry();
  const payload = buildWeeklyIntakeProcessPayload({
    parsedReport: { ...report(), rows: [["must not cross the adapter boundary"]] },
    expectedOutcome: "Duplicate",
    expectedLedgerEntry,
    processedAt: "2026-08-17T12:00:00.000Z",
  });
  assert.equal(payload.adapterVersion, WEEKLY_INTAKE_ADAPTER_VERSION);
  assert.equal(payload.operation, "Process");
  assert.equal(payload.expectedOutcome, "Duplicate");
  assert.equal(payload.parsedReport.manifest.sourcePeriodKey, "WEEK-2026-W01");
  assert.equal("rows" in payload.parsedReport, false);
  assert.deepEqual(payload.expectedLedgerEntry, expectedLedgerEntry);
  assert.throws(() => buildWeeklyIntakeProcessPayload({
    parsedReport: report(), expectedOutcome: "Review Later",
  }), /PUL-030WIA-001/);
});

test("ledger uses the approved hidden engineering surface and exact minimal schema", () => {
  assert.equal(WEEKLY_INTAKE_LEDGER_TABLE, "tblWeeklyIntakeLog");
  assert.equal(WEEKLY_INTAKE_LEDGER_SHEET, "_Weekly_Cache");
  assert.equal(WEEKLY_INTAKE_LEDGER_ADDRESS, "Y1:AL2");
  assert.deepEqual(weeklyIntakeLedgerHeaders(), [
    "IntakeEventID", "SourceLocator", "SourceFileID", "SourcePeriodKey",
    "SourceSemanticFingerprint", "IdentityPreflightFingerprint", "IntakeStatus",
    "StatusMessage", "SourceRowCount", "SourceSalesNOK", "ProcessedAt",
    "PriorCacheVersion", "ResultingCacheVersion", "SupersededCacheVersion",
  ]);
  assert.match(script, /const LEDGER_RANGE = "Y1:AL2"/);
  assert.match(script, /requiredSheet\(workbook, CACHE_SHEET\)/);
  assert.match(script, /SheetVisibility\.hidden/);
  assert.doesNotMatch(script, /addWorksheet\([^)]*IntakeLog|addWorksheet\([^)]*Imports/);
});

test("accepted publisher Candidate becomes deterministic bounded workbook payloads", () => {
  const plan = buildWeeklyIntakeCandidatePlan(publisherResult());
  assert.equal(plan.sheetName, WEEKLY_INTAKE_CANDIDATE_SHEET);
  assert.deepEqual(WEEKLY_INTAKE_CANDIDATE_TABLES, {
    version: "tblWeeklyIntakeCandidateVersions",
    period: "tblWeeklyIntakeCandidatePeriodManifest",
    scope: "tblWeeklyIntakeCandidateScopeCache",
    rpg: "tblWeeklyIntakeCandidateRPGCache",
  });
  assert.equal(plan.status, "New");
  assert.equal(plan.cacheStatus, "Candidate");
  assert.equal(plan.activationState, "Not Active");
  assert.equal(plan.completeCandidateRows, 23);
  assert.equal(plan.sections.version.startRow, 1);
  assert.equal(plan.sections.period.startRow, 5);
  assert.equal(plan.sections.scope.startRow, 10);
  assert.equal(plan.sections.rpg.startRow, 15);
  assert.equal(candidateLayoutFingerprint(plan.sections), plan.layoutFingerprint);

  const prepare = weeklyIntakeCandidatePreparePayload(plan);
  const first = weeklyIntakeCandidateChunk(plan, "rpg", 0, 5);
  const second = weeklyIntakeCandidateChunk(plan, "rpg", 5, 5);
  const finalize = weeklyIntakeCandidateFinalizePayload(plan);
  assert.equal(prepare.operation, "Prepare New");
  assert.equal(first.operation, "Write New");
  assert.equal(first.values.length, 5);
  assert.equal(second.startRow, plan.sections.rpg.startRow + 6);
  assert.match(first.chunkFingerprint, /^WCHK-[0-9a-f]{16}$/);
  assert.deepEqual(first, weeklyIntakeCandidateChunk(plan, "rpg", 0, 5));
  assert.equal(finalize.operation, "Finalize New");
  assert.equal(finalize.versionValues[2], "Candidate");
  assert.equal(finalize.versionValues[3], "Not Active");
  assert.equal(finalize.versionValues[4], "PASS");
});

test("Office Script keeps New materialization inactive and authority switching out of scope", () => {
  assert.match(script, /Prepare New -> Write New -> Finalize New/);
  assert.match(script, /Finalized \/ Not Active/);
  assert.match(script, /Candidate .* finalized as Not Active/);
  assert.match(script, /const activeAfter = resolveActiveAuthority\(workbook\)/);
  assert.doesNotMatch(script, /setValues\(\[\["Active", "Active"\]\]\)/);
  assert.doesNotMatch(script, /setValue\("Active"\)|setValues\([^\n]*"Active"/);
  assert.doesNotMatch(script, /tblSalesFacts|tblMetricRPGResults|Mapping Rules|tblMappingRules/);
  assert.doesNotMatch(script, /getWorksheet\("Performance"\)|getWorksheet\("Reports"\)|getWorksheet\("Imports"\)/);
});

test("Duplicate is the only archive-ready no-op and ledger upsert is idempotent", () => {
  assert.match(script, /status === "Duplicate", "Process"/);
  assert.match(script, /return "Existing"/);
  assert.match(script, /return "Added"/);
  assert.match(script, /hashStrings\(\[record\("INTAKE", eventKey\)\], "WINT-"\)/);
  assert.doesNotMatch(script, /table\.addRow\([^\n]*Duplicate/);
  const first = buildWeeklyIntakeProcessPayload({
    parsedReport: report(), expectedOutcome: "Duplicate", expectedLedgerEntry: ledgerEntry(),
  });
  const second = buildWeeklyIntakeProcessPayload({
    parsedReport: report(), expectedOutcome: "Duplicate", expectedLedgerEntry: ledgerEntry(),
  });
  assert.deepEqual(first, second);
});

test("typed return is small and Power Automate branches only on adapter status", () => {
  assert.match(script, /interface WeeklyIntakeResult/);
  for (const field of [
    "status", "period", "message", "cacheChanged", "resultingCacheVersion",
    "ledgerEventId", "archiveReady",
  ]) assert.match(script, new RegExp(`${field}:`));
  assert.match(script, /status: string/);
  assert.match(script, /cacheChanged: boolean/);
  assert.match(script, /archiveReady: boolean/);
});

test("Office Script stays batched and compatible with the Excel for web compiler target", () => {
  assert.match(script, /^function main\(/m);
  assert.doesNotMatch(script, /^(export|async|public|private|protected)\s+function\s+main/m);
  assert.doesNotMatch(script, /new\s+Map\s*</);
  assert.doesNotMatch(script, /new\s+Set\s*</);
  assert.doesNotMatch(script, /\.(entries|keys|values)\s*\(\)/);
  assert.doesNotMatch(script, /Array\.from\s*\(/);
  assert.doesNotMatch(script, /\.\.\./);
  const loopBodies = Array.from(script.matchAll(/for\s*\([^)]*\)\s*\{([\s\S]*?)\n  \}/g), match => match[1]);
  for (const body of loopBodies) {
    assert.doesNotMatch(body, /workbook\.get|sheet\.get|table\.get|\.getValues\(|\.getTexts\(/);
  }
  assertBalanced(script);
});

function report() {
  return {
    parserVersion: "0.3.0-weekly-parser-v1",
    schemaVersion: "sales-per-item-v1",
    manifest: {
      parserVersion: "0.3.0-weekly-parser-v1",
      schemaVersion: "sales-per-item-v1",
      sourceFileId: "SRCFILE-aaaaaaaaaaaaaaaa",
      sourceLocator: "OneDrive/Pulse/Incoming reports/banana.xlsx",
      semanticFingerprint: "WSF-aaaaaaaaaaaaaaaa",
      sourcePeriodKey: "WEEK-2026-W01",
      sourceRowCount: 10,
      totalSalesNok: 1000.25,
      contentReconciliationStatus: "PASS",
    },
  };
}

function changedReport() {
  const parsed = report();
  parsed.manifest.semanticFingerprint = "WSF-bbbbbbbbbbbbbbbb";
  parsed.manifest.sourceFileId = "SRCFILE-bbbbbbbbbbbbbbbb";
  parsed.manifest.totalSalesNok = 1001.25;
  return parsed;
}

function activeVersion() {
  return {
    cacheVersion: "WCV-1a34ad1f46763d9b",
    cacheFingerprint: "WCC-508dd608166cdb6e",
    cacheStatus: "Active",
    activationState: "Active",
    validationStatus: "PASS",
    mappingContentFingerprint: "MCF-759cc92c4304a913",
    catalogContentFingerprint: "ICC-5644a77c18a97437",
    identityPreflightFingerprint: "IDP-062c182f23905ae8",
    performanceRestaurantScopeFingerprint: "RSC-08df626f217dd94b",
  };
}

function freshness(active) {
  return {
    mappingContentFingerprint: active.mappingContentFingerprint,
    catalogContentFingerprint: active.catalogContentFingerprint,
    identityPreflightFingerprint: active.identityPreflightFingerprint,
    performanceRestaurantScopeFingerprint: active.performanceRestaurantScopeFingerprint,
  };
}

function periodRow(parsed) {
  return {
    cacheVersion: activeVersion().cacheVersion,
    sourcePeriodKey: parsed.manifest.sourcePeriodKey,
    sourceSemanticFingerprint: parsed.manifest.semanticFingerprint,
  };
}

function ledgerEntry() {
  return {
    intakeEventId: "WINT-1111111111111111",
    sourceLocator: report().manifest.sourceLocator,
    sourceFileId: report().manifest.sourceFileId,
    sourcePeriodKey: report().manifest.sourcePeriodKey,
    sourceSemanticFingerprint: report().manifest.semanticFingerprint,
    identityPreflightFingerprint: activeVersion().identityPreflightFingerprint,
    intakeStatus: "Duplicate",
    statusMessage: "Same period and semantic fingerprint; no cache change.",
    sourceRowCount: 10,
    sourceSalesNok: 1000.25,
    processedAt: "2026-08-17T12:00:00.000Z",
    priorCacheVersion: activeVersion().cacheVersion,
    resultingCacheVersion: "",
    supersededCacheVersion: "",
  };
}

function publisherResult() {
  const version = {
    cacheVersion: "WCV-2222222222222222",
    cacheSchemaVersion: "weekly-rpg-cache-v1",
    cacheStatus: "Candidate",
    activationState: "Not Active",
    validationStatus: "PASS",
    sourceSystemId: "POS-KATRIA",
    parserVersion: "0.3.0-weekly-parser-v1",
    identityContractVersion: "0.3.0-weekly-identity-v1",
    sourceCorpusFingerprint: "WSC-2222222222222222",
    identityPreflightFingerprint: "IDP-2222222222222222",
    catalogFingerprint: "ICF-2222222222222222",
    catalogContentFingerprint: "ICC-5644a77c18a97437",
    mappingContentFingerprint: "MCF-759cc92c4304a913",
    mappingFingerprint: "MAP-34202a7a1a922bd0",
    mappingAsOfDate: "2026-08-17",
    activeReportingGroupFingerprint: "RPG-2222222222222222",
    performanceRestaurantScopeFingerprint: "RSC-08df626f217dd94b",
    periodRowCount: 2,
    scopeCacheRowCount: 2,
    denseRpgCacheRowCount: 18,
    nonzeroRpgCacheRowCount: 2,
    cacheFingerprint: "WCC-2222222222222222",
  };
  return {
    outcome: "New",
    candidatePrepared: true,
    activeCacheVersion: activeVersion().cacheVersion,
    ledgerEntry: { ...ledgerEntry(), intakeStatus: "New", resultingCacheVersion: version.cacheVersion },
    candidate: {
      versionManifest: version,
      validation: { status: "PASS" },
      periodManifest: [objectRow("period", version.cacheVersion, 1), objectRow("period", version.cacheVersion, 2)],
      scopeCacheRows: [objectRow("scope", version.cacheVersion, 1), objectRow("scope", version.cacheVersion, 2)],
      weeklyRpgCacheRows: Array.from({ length: 18 }, (_, index) =>
        objectRow("rpg", version.cacheVersion, index + 1)),
    },
  };
}

function objectRow(kind, version, suffix) {
  if (kind === "period") return {
    weeklyPeriodManifestRowId: `WPER-${suffix}`, cacheVersion: version,
    sourcePeriodKey: `WEEK-2026-W0${suffix}`, periodStart: "2026-01-01", periodEnd: "2026-01-07",
    isoYear: 2026, isoWeek: suffix, sourceFileId: `SRCFILE-${suffix}`,
    sourceSemanticFingerprint: `WSF-${suffix}`, sourceBinaryFingerprint: "NOT-PROVIDED",
    scopeId: "SCOPE", scopeFingerprint: "SFP", sourceFactCount: 1,
    sourceSalesNok: 10, sourceQuantity: 1, sourceRestaurantCount: 1,
  };
  if (kind === "scope") return {
    weeklyScopeCacheRowId: `WSCP-${suffix}`, cacheVersion: version,
    sourcePeriodKey: `WEEK-2026-W0${suffix}`, restaurantId: `RST-${suffix}`,
    performanceEligible: "Yes", sourceFactCount: 1, sourceSalesNok: 10, sourceQuantity: 1,
    mappedFactCount: 1, mappedSalesNok: 10, mappedQuantity: 1,
    unmappedFactCount: 0, unmappedSalesNok: 0, unmappedQuantity: 0,
    identityPendingFactCount: 0, identityPendingSalesNok: 0, identityPendingQuantity: 0,
    conflictFactCount: 0, conflictSalesNok: 0, conflictQuantity: 0,
    inactiveTargetFactCount: 0, inactiveTargetSalesNok: 0, inactiveTargetQuantity: 0,
  };
  const rpg = ((suffix - 1) % 9) + 1;
  const period = suffix <= 9 ? 1 : 2;
  return {
    weeklyRpgCacheRowId: `WRPG-${suffix}`, cacheVersion: version,
    sourcePeriodKey: `WEEK-2026-W0${period}`, restaurantId: `RST-${period}`,
    reportingGroupId: `RPG-${String(rpg).padStart(4, "0")}`,
    mappedFactCount: rpg === 1 ? 1 : 0, mappedSalesNok: rpg === 1 ? 10 : 0,
    mappedQuantity: rpg === 1 ? 1 : 0,
  };
}

function assertBalanced(source) {
  let braces = 0; let parentheses = 0; let brackets = 0; let quote = "";
  let escaped = false; let lineComment = false; let blockComment = false;
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index]; const next = source[index + 1];
    if (lineComment) { if (char === "\n") lineComment = false; continue; }
    if (blockComment) { if (char === "*" && next === "/") { blockComment = false; index += 1; } continue; }
    if (quote) {
      if (escaped) escaped = false;
      else if (char === "\\") escaped = true;
      else if (char === quote) quote = "";
      continue;
    }
    if (char === "/" && next === "/") { lineComment = true; index += 1; continue; }
    if (char === "/" && next === "*") { blockComment = true; index += 1; continue; }
    if (char === '"' || char === "'" || char === "`") { quote = char; continue; }
    if (char === "{") braces += 1; else if (char === "}") braces -= 1;
    if (char === "(") parentheses += 1; else if (char === ")") parentheses -= 1;
    if (char === "[") brackets += 1; else if (char === "]") brackets -= 1;
    assert.ok(braces >= 0 && parentheses >= 0 && brackets >= 0, "source closes before it opens");
  }
  assert.deepEqual({ braces, parentheses, brackets, quote },
    { braces: 0, parentheses: 0, brackets: 0, quote: "" });
}
