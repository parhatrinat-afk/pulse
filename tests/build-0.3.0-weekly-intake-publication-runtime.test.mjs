import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import {
  applyWeeklyIntakePublishedActivation,
  planWeeklyIntakePublishedActivation,
} from "../src/imports/weekly-intake-publication-runtime.mjs";
import {
  fingerprintWeeklyCacheRows,
} from "../src/imports/weekly-compact-cache.mjs";
import {
  fingerprintWeeklyIntakeIdentityEvidence,
} from "../src/imports/weekly-intake-publisher.mjs";

const scriptUrl = new URL("../office-scripts/Publish_Weekly_Intake.ts", import.meta.url);
const script = await readFile(scriptUrl, "utf8");
const registryScriptUrl = new URL("../office-scripts/Install_Weekly_Identity_Registry.ts",
  import.meta.url);
const registryScript = await readFile(registryScriptUrl, "utf8");
const auditUrl = new URL("../src/imports/audit-weekly-intake-publication.mjs", import.meta.url);
const audit = await readFile(auditUrl, "utf8");
const w33EvidenceUrl = new URL("./expected-build-0.3.0-weekly-w33-publication.json",
  import.meta.url);

test("accepted New Candidate plans one Published authority and one rollback version", () => {
  const fixture = publicationFixture();
  const plan = planWeeklyIntakePublishedActivation(fixture);
  const final = applyWeeklyIntakePublishedActivation(fixture.versionManifests, plan);

  assert.equal(plan.action, "Publish");
  assert.equal(plan.priorCacheVersion, "WCV-prior");
  assert.equal(plan.resultingCacheVersion, "WCV-candidate");
  assert.equal(plan.result.status, "Published");
  assert.equal(plan.result.cacheChanged, true);
  assert.equal(plan.result.archiveReady, true);
  assert.deepEqual(plan.counts, { period: 1, scope: 1, rpg: 9 });
  assert.equal(final.filter(row => row.cacheStatus === "Active" &&
    row.activationState === "Active").length, 1);
  assert.equal(final.filter(row => row.cacheStatus === "Rollback" &&
    row.activationState === "Not Active").length, 1);
});

test("publication planning is deterministic and does not mutate prior Active or Candidate", () => {
  const fixture = publicationFixture();
  const before = structuredClone(fixture);
  const first = planWeeklyIntakePublishedActivation(fixture);
  const second = planWeeklyIntakePublishedActivation(fixture);

  assert.deepEqual(first, second);
  assert.deepEqual(fixture, before);
  assert.equal(first.ledgerEntry.intakeStatus, "Published");
  assert.equal(first.ledgerEntry.sourceLocator,
    "OneDrive/Pulse/Incoming reports/W33.xlsx");
  assert.match(first.ledgerEntry.intakeEventId, /^WINT-[0-9a-f]{16}$/);
});

test("stale, corrupt, active, or count-mismatched Candidates cannot publish", () => {
  const stale = publicationFixture();
  stale.currentFreshness.mappingContentFingerprint = "MCF-changed";
  assert.throws(() => planWeeklyIntakePublishedActivation(stale), /PUL-030WPR-002/);

  const corrupt = publicationFixture();
  corrupt.publicationResult.candidate.versionManifest.cacheFingerprint = "WCC-corrupt";
  assert.throws(() => planWeeklyIntakePublishedActivation(corrupt), /PUL-030WPR-007/);

  const active = publicationFixture();
  active.publicationResult.candidate.versionManifest.cacheStatus = "Active";
  assert.throws(() => planWeeklyIntakePublishedActivation(active), /PUL-030WPR-004/);

  const count = publicationFixture();
  count.publicationResult.candidate.versionManifest.scopeCacheRowCount = 2;
  assert.throws(() => planWeeklyIntakePublishedActivation(count), /PUL-030WPR-006/);
});

test("compact weekly identity evidence is order-independent and assignment-sensitive", () => {
  const preflight = identityPreflight();
  const reversed = {
    ...preflight,
    rowAssignments: [...preflight.rowAssignments].reverse(),
    newIdentityCandidates: {
      ...preflight.newIdentityCandidates,
      products: [...preflight.newIdentityCandidates.products].reverse(),
    },
  };
  assert.equal(
    fingerprintWeeklyIntakeIdentityEvidence(preflight),
    fingerprintWeeklyIntakeIdentityEvidence(reversed),
  );
  const changed = structuredClone(preflight);
  changed.rowAssignments[0].productId = "PRD-999999";
  assert.notEqual(
    fingerprintWeeklyIntakeIdentityEvidence(preflight),
    fingerprintWeeklyIntakeIdentityEvidence(changed),
  );
});

test("Office Script exposes two bounded calls and keeps authority switch final", () => {
  assert.match(script, /^function main\(/m);
  assert.match(script, /operation === "Build Candidate"/);
  assert.match(script, /operation === "Activate Candidate"/);
  assert.match(script, /materializeCandidate\(workbook, candidate\)/);
  assert.match(script, /Candidate .*PASS \/ Not Active/);
  assert.match(script, /backupActiveCache\(workbook, active\)/);
  assert.match(script, /writeCanonicalCache\(active, versionRows/);
  assert.match(script, /statusRange\.setValues\(\[\["Rollback", "Not Active"\], \["Active", "Active"\]\]\)/);
  assert.ok(
    script.indexOf("writeCanonicalCache(active, versionRows") <
      script.indexOf('statusRange.setValues([["Rollback", "Not Active"], ["Active", "Active"]])'),
  );
  assert.match(script, /if \(canonicalMutationStarted\) \{\s*writeIdentityRegistry\(workbook, active\.identityRegistry\);\s*restoreCanonicalCache\(active\)/);
  assert.match(script, /writeIdentityRegistry\(workbook, candidate\.identityRegistry\)/);
  assert.match(script, /writeIdentityRegistry\(workbook, active\.identityRegistry\)/);
  assert.match(script, /tblWeeklyIdentityRegistryControl/);
  assert.match(script, /validateIdentityRegistryCarryForward/);
  assert.match(script, /activation was already complete/);
  assert.match(script, /sourceLocator: ledgerSourceLocator\(workbook, key, semantic\)/);
  assert.match(script, /row\[periodIndex\] === period && row\[semanticIndex\] === semantic/);
});

test("accepted 84-week identity registry installer is narrow and idempotent", () => {
  assert.match(registryScript, /^function main\(/m);
  assert.match(registryScript, /EXPECTED_ACTIVE_VERSION = "WCV-1a34ad1f46763d9b"/);
  assert.match(registryScript, /EXPECTED_IDP = "IDP-062c182f23905ae8"/);
  assert.match(registryScript, /EXPECTED_WIR = "WIR-776953cb0144af11"/);
  assert.match(registryScript, /RESTAURANTS\.length/);
  assert.match(registryScript, /CLASSIFICATIONS\.length/);
  assert.match(registryScript, /PRODUCTS\.length/);
  assert.match(registryScript, /Already installed/);
  assert.doesNotMatch(registryScript,
    /tblSalesFacts|tblMetricRPGResults|tblMappingRules|getWorksheet\("Performance"\)|getWorksheet\("Reports"\)|getWorksheet\("Imports"\)/);
});

test("Office Script is isolated from facts, Performance, Reports, Imports, and mapping writes", () => {
  assert.doesNotMatch(script, /tblSalesFacts|tblMetricRPGResults/);
  assert.doesNotMatch(script, /getWorksheet\("Performance"\)|getWorksheet\("Reports"\)|getWorksheet\("Imports"\)/);
  assert.doesNotMatch(script, /tblMappingRules[^\n]*(addRow|setValues|deleteRowsAt)/);
  assert.doesNotMatch(script, /tblProducts[^\n]*(addRow|setValues|deleteRowsAt)/);
  assert.doesNotMatch(script, /tblRestaurants[^\n]*(addRow|setValues|deleteRowsAt)/);
});

test("Office Script uses bounded writes and stays compatible with Excel for web", () => {
  for (const source of [script, registryScript]) {
    assert.doesNotMatch(source, /^(export|async|public|private|protected)\s+function\s+main/m);
    assert.doesNotMatch(source, /new\s+Map\s*</);
    assert.doesNotMatch(source, /new\s+Set\s*</);
    assert.doesNotMatch(source, /\.(entries|keys|values)\s*\(\)/);
    assert.doesNotMatch(source, /Array\.from\s*\(/);
    assert.doesNotMatch(source, /\.\.\./);
    assertBalanced(source);
  }
  assert.match(script, /const chunk = 500/);
  assert.match(script, /getResizedRange\(part\.length - 1, columns\.length - 1\)\.setValues\(part\)/);
  assert.match(registryScript, /const chunk = 500/);
});

test("publication audit requires exact inputs and remains read-only", () => {
  assert.match(audit, /exactPath\(process\.argv\[2\]/);
  assert.match(audit, /flag\("--incoming"\)/);
  assert.match(audit, /flag\("--catalog"\)/);
  assert.match(audit, /businessRowsEqual\(candidate, full\)/);
  assert.doesNotMatch(audit, /writeFile|appendFile|rename\(|unlink|rm\(|ExcelScript|Pulse_Current\.xlsx/);
});

test("frozen W33 publication evidence records the accepted 85-week build", async () => {
  const evidence = JSON.parse(await readFile(w33EvidenceUrl, "utf8"));
  assert.deepEqual(evidence.period, {
    sourcePeriodKey: "PERIOD-2026-08-10-2026-08-16",
    periodStart: "2026-08-10", periodEnd: "2026-08-16",
    isoYear: 2026, isoWeek: 33, sourceRowCount: 2940,
    salesNok: 5636773.5, quantity: 29654.35,
    semanticFingerprint: "WSF-641061337dfbfd59",
  });
  assert.equal(evidence.candidate.cacheVersion, "WCV-1b0b195c210da456");
  assert.equal(evidence.candidate.cacheFingerprint, "WCC-26c195956ebc2823");
  assert.equal(evidence.candidate.identityPreflightFingerprint, "IDP-4ae1a62974cca3af");
  assert.equal(evidence.candidate.identityRegistryFingerprint, "WIR-00fc39ff746cb4d1");
  assert.equal(evidence.candidate.periodRows, 85);
  assert.equal(evidence.candidate.scopeRows, 1438);
  assert.equal(evidence.candidate.rpgRows, 12942);
  assert.equal(evidence.candidate.completeRows, 14466);
  assert.deepEqual(evidence.candidate.source, {
    factCount: 248572, salesNok: 490365140.75, quantity: 2499642.44,
  });
  assert.equal(evidence.w33.newProducts.length, 3);
  assert.deepEqual(evidence.independentFullRebuild, {
    status: "PASS", businessRowsEqual: true, sourceTotalsEqual: true,
  });
});

function publicationFixture() {
  const currentFreshness = {
    mappingContentFingerprint: "MCF-current",
    catalogContentFingerprint: "ICC-current",
    identityPreflightFingerprint: "IDP-current",
    performanceRestaurantScopeFingerprint: "RSC-current",
  };
  const active = {
    cacheVersion: "WCV-prior",
    cacheFingerprint: "WCC-prior",
    cacheStatus: "Active",
    activationState: "Active",
    validationStatus: "PASS",
    ...currentFreshness,
  };
  const periodManifest = [{
    weeklyPeriodManifestRowId: "WPER-1", cacheVersion: "WCV-candidate",
    sourcePeriodKey: "WEEK-2026-W33", periodStart: "2026-08-10", periodEnd: "2026-08-16",
    sourceFileId: "SRCFILE-1", sourceSemanticFingerprint: "WSF-1",
    sourceFactCount: 1, sourceSalesNok: 100, sourceQuantity: 1,
  }];
  const scopeCacheRows = [{
    weeklyScopeCacheRowId: "WSCP-1", cacheVersion: "WCV-candidate",
    sourcePeriodKey: "WEEK-2026-W33", restaurantId: "RST-0001", performanceEligible: "Yes",
    sourceFactCount: 1, sourceSalesNok: 100, sourceQuantity: 1,
    mappedFactCount: 1, mappedSalesNok: 100, mappedQuantity: 1,
    unmappedFactCount: 0, unmappedSalesNok: 0, unmappedQuantity: 0,
    identityPendingFactCount: 0, identityPendingSalesNok: 0, identityPendingQuantity: 0,
    conflictFactCount: 0, conflictSalesNok: 0, conflictQuantity: 0,
    inactiveTargetFactCount: 0, inactiveTargetSalesNok: 0, inactiveTargetQuantity: 0,
  }];
  const weeklyRpgCacheRows = Array.from({ length: 9 }, (_, index) => ({
    weeklyRpgCacheRowId: `WRPG-${index + 1}`, cacheVersion: "WCV-candidate",
    sourcePeriodKey: "WEEK-2026-W33", restaurantId: "RST-0001",
    reportingGroupId: `RPG-${String(index + 1).padStart(4, "0")}`,
    mappedFactCount: index === 0 ? 1 : 0,
    mappedSalesNok: index === 0 ? 100 : 0,
    mappedQuantity: index === 0 ? 1 : 0,
  }));
  const cacheFingerprint = fingerprintWeeklyCacheRows({
    cacheVersion: "WCV-candidate", periodManifest, scopeCacheRows, weeklyRpgCacheRows,
  });
  const candidate = {
    versionManifest: {
      cacheVersion: "WCV-candidate", cacheFingerprint,
      cacheStatus: "Candidate", activationState: "Not Active", validationStatus: "PASS",
      periodRowCount: 1, scopeCacheRowCount: 1, denseRpgCacheRowCount: 9,
      ...currentFreshness,
    },
    periodManifest, scopeCacheRows, weeklyRpgCacheRows,
    validation: { status: "PASS" },
  };
  return {
    publicationResult: {
      outcome: "New", candidatePrepared: true, activeCacheVersion: "WCV-prior",
      candidate,
      ledgerEntry: {
        sourcePeriodKey: "WEEK-2026-W33", sourceSemanticFingerprint: "WSF-1",
        sourceLocator: "OneDrive/Pulse/Incoming reports/W33.xlsx", processedAt: "",
      },
    },
    versionManifests: [active], currentFreshness, processedAt: "2026-08-17T12:00:00Z",
  };
}

function identityPreflight() {
  return {
    reconciliation: { status: "PASS" },
    rowAssignments: [{
      sourceRowId: "ROW-1", restaurantId: "RST-0001", productId: "PRD-001042",
      sourceClassificationId: "SCL-00146", identityState: "Stable",
      identityPendingReason: "", mappingStatus: "Mapped", effectiveReportingGroupId: "RPG-0001",
    }, {
      sourceRowId: "ROW-2", restaurantId: "RST-0001", productId: "PRD-000001",
      sourceClassificationId: "SCL-00001", identityState: "Stable",
      identityPendingReason: "", mappingStatus: "Unmapped", effectiveReportingGroupId: "",
    }],
    newIdentityCandidates: {
      restaurants: [], classifications: [], products: [{
        productId: "PRD-001042", sourceSystemId: "SRC-TEST-SALES",
        productKey: "SRC-TEST-SALES || New || Sales", sourceClassificationId: "SCL-00146",
        productStatus: "Active", hierarchyStatus: "Proposed exact hierarchy",
        observedHierarchyPaths: ["SRC-TEST-SALES || Main || Sub"],
      }],
    },
  };
}

function assertBalanced(source) {
  const pairs = [["{", "}"], ["(", ")"], ["[", "]"]];
  for (const [open, close] of pairs) {
    let depth = 0;
    for (const character of source) {
      if (character === open) depth += 1;
      if (character === close) depth -= 1;
      assert.ok(depth >= 0, `${close} closed before ${open}`);
    }
    assert.equal(depth, 0, `${open}${close} are not balanced`);
  }
}
