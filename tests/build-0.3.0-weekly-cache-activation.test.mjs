import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import {
  applyWeeklyCacheActivationPlan,
  planWeeklyCacheActivation,
  validateActiveWeeklyCacheFreshness,
} from "../src/imports/weekly-cache-activation.mjs";

const scriptPath = new URL("../office-scripts/Activate_Weekly_Compact_Cache.ts", import.meta.url);
const script = await readFile(scriptPath, "utf8");

test("exact validated candidate produces the two-field Active authority transition", () => {
  const before = candidate();
  const plan = planWeeklyCacheActivation(contract([before]));
  const after = applyWeeklyCacheActivationPlan([before], plan)[0];

  assert.deepEqual(plan, {
    action: "Activate",
    cacheVersion: EXPECTED.cacheVersion,
    updates: { cacheStatus: "Active", activationState: "Active" },
  });
  const changed = Object.keys(after).filter(key => after[key] !== before[key]);
  assert.deepEqual(changed, ["cacheStatus", "activationState"]);
  assert.equal(after.cacheStatus, "Active");
  assert.equal(after.activationState, "Active");
});

test("exact Active rerun is idempotent", () => {
  const active = candidate({ cacheStatus: "Active", activationState: "Active" });
  const plan = planWeeklyCacheActivation(contract([active]));
  assert.deepEqual(plan, {
    action: "Already Active",
    cacheVersion: EXPECTED.cacheVersion,
    updates: {},
  });
  assert.deepEqual(applyWeeklyCacheActivationPlan([active], plan), [active]);
});

test("activation rejects another Active version or split authority state", () => {
  assert.throws(() => planWeeklyCacheActivation(contract([
    candidate(),
    candidate({ cacheVersion: "WCV-other", cacheStatus: "Active", activationState: "Active" }),
  ])), /PUL-030A-002/);
  assert.throws(() => planWeeklyCacheActivation(contract([
    candidate({ cacheStatus: "Candidate", activationState: "Active" }),
  ])), /PUL-030A-005/);
});

test("activation rejects stale fingerprints, row counts, reconciliation or failed QA", () => {
  assert.throws(() => planWeeklyCacheActivation(contract([candidate()], {
    current: { mappingContentFingerprint: "MCF-changed" },
  })), /PUL-030A-004.*MappingContentFingerprint/);
  assert.throws(() => planWeeklyCacheActivation(contract([candidate()], {
    materialized: { denseRpgCacheRowCount: 12788 },
  })), /PUL-030A-010/);
  assert.throws(() => planWeeklyCacheActivation(contract([candidate()], {
    materialized: { reconciliationStatus: "FAIL" },
  })), /PUL-030A-011/);
  assert.throws(() => planWeeklyCacheActivation(contract([
    candidate({ validationStatus: "FAIL" }),
  ])), /PUL-030A-009/);
});

test("one exact Active version is available to future weekly consumers", () => {
  const active = candidate({ cacheStatus: "Active", activationState: "Active" });
  assert.deepEqual(validateActiveWeeklyCacheFreshness({
    versionManifests: [active],
    current: current(),
  }), {
    status: "Available",
    activeVersion: EXPECTED.cacheVersion,
    errors: [],
  });
});

test("date-only mapping audit changes do not stale the weekly cache", () => {
  const active = candidate({
    cacheStatus: "Active",
    activationState: "Active",
    mappingAsOfDate: "2026-08-11",
    mappingFingerprint: "MAP-date-sensitive-11",
  });
  const live = {
    ...current(),
    mappingAsOfDate: "2026-08-12",
    mappingFingerprint: "MAP-date-sensitive-12",
  };
  assert.equal(validateActiveWeeklyCacheFreshness({
    versionManifests: [active], current: live,
  }).status, "Available");
});

test("mapping, catalog, identity or ReportingEnabled scope changes make the Active cache stale", () => {
  const active = candidate({ cacheStatus: "Active", activationState: "Active" });
  const cases = [
    ["mappingContentFingerprint", "MCF-changed", "MappingContentFingerprint"],
    ["catalogContentFingerprint", "ICC-changed", "CatalogContentFingerprint"],
    ["identityPreflightFingerprint", "IDP-changed", "IdentityPreflightFingerprint"],
    ["performanceRestaurantScopeFingerprint", "RSC-changed", "PerformanceRestaurantScopeFingerprint"],
  ];
  for (const [field, changed, label] of cases) {
    const result = validateActiveWeeklyCacheFreshness({
      versionManifests: [active],
      current: { ...current(), [field]: changed },
    });
    assert.equal(result.status, "Stale");
    assert.match(result.errors.join(" "), new RegExp(label));
  }
});

test("missing, multiple or failed Active authority is unavailable/stale", () => {
  assert.equal(validateActiveWeeklyCacheFreshness({
    versionManifests: [candidate()], current: current(),
  }).status, "Unavailable");
  assert.equal(validateActiveWeeklyCacheFreshness({
    versionManifests: [
      candidate({ cacheStatus: "Active", activationState: "Active" }),
      candidate({ cacheVersion: "WCV-other", cacheStatus: "Active", activationState: "Active" }),
    ],
    current: current(),
  }).status, "Unavailable");
  assert.equal(validateActiveWeeklyCacheFreshness({
    versionManifests: [candidate({
      cacheStatus: "Active", activationState: "Active", validationStatus: "FAIL",
    })],
    current: current(),
  }).status, "Stale");
});

test("Office Script activates only the exact accepted candidate and ignores mapping date", () => {
  assert.match(script, /EXPECTED_CACHE_VERSION = "WCV-1a34ad1f46763d9b"/);
  assert.match(script, /EXPECTED_CACHE_FINGERPRINT = "WCC-508dd608166cdb6e"/);
  assert.match(script, /EXPECTED_MAPPING_CONTENT = "MCF-759cc92c4304a913"/);
  assert.match(script, /EXPECTED_CATALOG_CONTENT = "ICC-5644a77c18a97437"/);
  assert.match(script, /EXPECTED_IDENTITY_PREFLIGHT = "IDP-062c182f23905ae8"/);
  assert.match(script, /EXPECTED_RESTAURANT_SCOPE = "RSC-08df626f217dd94b"/);
  assert.doesNotMatch(script, /EXPECTED_MAPPING_AS_OF|version\[14\]/);
  assert.match(script, /authority\.setValues\(\[\["Active", "Active"\]\]\)/);
  assert.equal((script.match(/\.setValues\s*\(/g) || []).length, 1);
});

test("Office Script preflight proves exact materialized rows, reconciliation and Phase 2C protection", () => {
  assert.match(script, /periodRows\.length !== 84/);
  assert.match(script, /scopeRows\.length !== 1421/);
  assert.match(script, /rpgRows\.length !== 12789/);
  assert.match(script, /EXPECTED_FACTS = 245632/);
  assert.match(script, /EXPECTED_SALES = 484728367\.25/);
  assert.match(script, /EXPECTED_QUANTITY = 2469988\.09/);
  assert.match(script, /EXPECTED_PENDING_FACTS = 120/);
  assert.match(script, /assertMetricEqual\(mappedRpg, mapped/);
  assert.match(script, /fingerprintCache\(version, periodRows, scopeRows, rpgRows\)/);
  assert.match(script, /protectedSnapshot\(workbook\)/);
  assert.match(script, /tblPerformanceInteractionQA/);
  assert.match(script, /Object\.keys\(seen\)\.length !== 16/);
});

test("activation has no Performance, Reports, import, selector or analytical-row mutation path", () => {
  assert.doesNotMatch(script, /addWorksheet|delete\(\)|setVisibility|setPosition|setFormula|setFormulas/);
  assert.doesNotMatch(script, /Period Selector|From week|To week|supersed/i);
  const setCall = script.match(/[^\n]*\.setValues\([^\n]*/g) || [];
  assert.deepEqual(setCall.map(line => line.trim()), ["authority.setValues([[\"Active\", \"Active\"]]);"]);
  assert.match(script, /performance: rangeFingerprint\(performance\.getRange\("A1:T47"\)\)/);
  assert.match(script, /reports: rangeFingerprint\(reports\.getRange\("A1:H20"\)\)/);
  assert.match(script, /metricResults: rangeFingerprint\(requiredTable\(workbook, "tblMetricRPGResults"\)\.getRange\(\)\)/);
  assert.match(script, /imports: rangeFingerprint\(requiredTable\(workbook, "tblImports"\)\.getRange\(\)\)/);
});

test("activation Office Script remains compiler/performance compatible", () => {
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

const EXPECTED = Object.freeze({
  cacheVersion: "WCV-1a34ad1f46763d9b",
  cacheFingerprint: "WCC-508dd608166cdb6e",
  mappingContentFingerprint: "MCF-759cc92c4304a913",
  catalogContentFingerprint: "ICC-5644a77c18a97437",
  identityPreflightFingerprint: "IDP-062c182f23905ae8",
  performanceRestaurantScopeFingerprint: "RSC-08df626f217dd94b",
  periodRowCount: 84,
  scopeCacheRowCount: 1421,
  denseRpgCacheRowCount: 12789,
});

function candidate(overrides = {}) {
  return {
    ...EXPECTED,
    cacheStatus: "Candidate",
    activationState: "Not Active",
    validationStatus: "PASS",
    mappingAsOfDate: "2026-08-12",
    mappingFingerprint: "MAP-34202a7a1a922bd0",
    ...overrides,
  };
}

function current(overrides = {}) {
  return {
    mappingContentFingerprint: EXPECTED.mappingContentFingerprint,
    catalogContentFingerprint: EXPECTED.catalogContentFingerprint,
    identityPreflightFingerprint: EXPECTED.identityPreflightFingerprint,
    performanceRestaurantScopeFingerprint: EXPECTED.performanceRestaurantScopeFingerprint,
    ...overrides,
  };
}

function contract(versionManifests, overrides = {}) {
  return {
    versionManifests,
    expected: EXPECTED,
    materialized: {
      cacheFingerprint: EXPECTED.cacheFingerprint,
      reconciliationStatus: "PASS",
      periodRowCount: EXPECTED.periodRowCount,
      scopeCacheRowCount: EXPECTED.scopeCacheRowCount,
      denseRpgCacheRowCount: EXPECTED.denseRpgCacheRowCount,
      ...overrides.materialized,
    },
    current: current(overrides.current),
  };
}

function assertBalanced(source) {
  let braces = 0; let parentheses = 0; let brackets = 0; let quote = "";
  let escaped = false; let lineComment = false; let blockComment = false;
  for (let index = 0; index < source.length; index += 1) {
    const char = source[index]; const next = source[index + 1];
    if (lineComment) { if (char === "\n") lineComment = false; continue; }
    if (blockComment) { if (char === "*" && next === "/") { blockComment = false; index += 1; } continue; }
    if (quote) { if (escaped) escaped = false; else if (char === "\\") escaped = true; else if (char === quote) quote = ""; continue; }
    if (char === "/" && next === "/") { lineComment = true; index += 1; continue; }
    if (char === "/" && next === "*") { blockComment = true; index += 1; continue; }
    if (char === '"' || char === "'" || char === "`") quote = char;
    else if (char === "{") braces += 1; else if (char === "}") braces -= 1;
    else if (char === "(") parentheses += 1; else if (char === ")") parentheses -= 1;
    else if (char === "[") brackets += 1; else if (char === "]") brackets -= 1;
    assert.ok(braces >= 0 && parentheses >= 0 && brackets >= 0, `Unbalanced delimiter at ${index}`);
  }
  assert.equal(quote, ""); assert.equal(blockComment, false);
  assert.equal(braces, 0); assert.equal(parentheses, 0); assert.equal(brackets, 0);
}
