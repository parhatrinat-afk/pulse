import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import {
  WEEKLY_CACHE_COLUMNS,
  WEEKLY_CACHE_LAYOUT,
  WEEKLY_CACHE_SHEET,
  WEEKLY_CACHE_STAGING_SHEET,
  WEEKLY_CACHE_TABLES,
  buildWeeklyCacheMaterializationPlan,
  decideWeeklyCacheMaterialization,
  materializationChunk,
  materializedSnapshotFromPlan,
} from "../src/imports/weekly-cache-materialization.mjs";

const officeScriptPath = new URL("../office-scripts/Materialize_Weekly_Compact_Cache.ts", import.meta.url);
const auditPath = new URL("../src/imports/audit-weekly-cache-materialization.mjs", import.meta.url);
const expectedPath = new URL("./expected-build-0.3.0-weekly-compact-cache.json", import.meta.url);
const script = await readFile(officeScriptPath, "utf8");
const audit = await readFile(auditPath, "utf8");

test("materialization plan uses four bounded tables on one hidden engineering surface", () => {
  const plan = buildWeeklyCacheMaterializationPlan(candidate());
  const snapshot = materializedSnapshotFromPlan(plan);

  assert.equal(plan.sheetName, WEEKLY_CACHE_SHEET);
  assert.equal(plan.stagingSheetName, WEEKLY_CACHE_STAGING_SHEET);
  assert.deepEqual(WEEKLY_CACHE_TABLES, {
    version: "tblWeeklyCacheVersions",
    period: "tblWeeklyPeriodManifest",
    scope: "tblWeeklyScopeCache",
    rpg: "tblWeeklyRPGCache",
  });
  assert.deepEqual(Object.fromEntries(Object.entries(plan.sections).map(([name, section]) => [name, {
    address: section.address,
    rows: section.rowCount,
    columns: section.columnCount,
  }])), {
    version: { address: "A1:V2", rows: 1, columns: 22 },
    period: { address: "A5:P7", rows: 2, columns: 16 },
    scope: { address: "A92:W93", rows: 1, columns: 23 },
    rpg: { address: "A1516:H1524", rows: 8, columns: 8 },
  });
  assert.equal(snapshot.status, "Candidate");
  assert.equal(snapshot.activationState, "Not Active");
  assert.equal(snapshot.completeCandidateRows, 12);
});

test("materialization chunks preserve exact row order and bounded geometry", () => {
  const plan = buildWeeklyCacheMaterializationPlan(candidate());
  const first = materializationChunk(plan, "rpg", 0, 3);
  const second = materializationChunk(plan, "rpg", 3, 3);

  assert.equal(first.startRow, WEEKLY_CACHE_LAYOUT.rpg.startRow + 1);
  assert.equal(first.startColumn, 1);
  assert.deepEqual(first.headers, WEEKLY_CACHE_COLUMNS.rpg.map(([header]) => header));
  assert.deepEqual(second.headers, []);
  assert.equal(second.startRow, WEEKLY_CACHE_LAYOUT.rpg.startRow + 4);
  assert.deepEqual([...first.values, ...second.values], plan.sections.rpg.values.slice(0, 6));
  assert.throws(() => materializationChunk(plan, "rpg", 9, 1), /PUL-030M-006/);
});

test("only validated Candidate / Not Active caches can be materialized", () => {
  const active = candidate();
  active.versionManifest.activationState = "Active";
  assert.throws(() => buildWeeklyCacheMaterializationPlan(active), /PUL-030M-001/);

  const failed = candidate();
  failed.validation.status = "FAIL";
  assert.throws(() => buildWeeklyCacheMaterializationPlan(failed), /PUL-030M-002/);
});

test("idempotency reuses an exact inactive candidate and rejects active or differing versions", () => {
  assert.equal(decideWeeklyCacheMaterialization([], "WCV-test", "WCC-test"), "Prepare");
  assert.equal(decideWeeklyCacheMaterialization([{
    cacheVersion: "WCV-test", cacheFingerprint: "WCC-test",
    cacheStatus: "Candidate", activationState: "Not Active",
  }], "WCV-test", "WCC-test"), "Already Materialized");
  assert.throws(() => decideWeeklyCacheMaterialization([{
    cacheVersion: "WCV-test", cacheFingerprint: "WCC-test",
    cacheStatus: "Active", activationState: "Active",
  }], "WCV-test", "WCC-test"), /PUL-030M-008/);
  assert.throws(() => decideWeeklyCacheMaterialization([{
    cacheVersion: "WCV-test", cacheFingerprint: "WCC-other",
    cacheStatus: "Candidate", activationState: "Not Active",
  }], "WCV-test", "WCC-test"), /PUL-030M-009/);
});

test("frozen 84-week evidence matches the exact Excel materialization contract", async () => {
  const expected = JSON.parse(await readFile(expectedPath, "utf8"));
  assert.equal(expected.cache_version, "WCV-1a34ad1f46763d9b");
  assert.equal(expected.cache_fingerprint, "WCC-508dd608166cdb6e");
  assert.equal(expected.mapping_content_fingerprint, "MCF-759cc92c4304a913");
  assert.equal(expected.identity_preflight_fingerprint, "IDP-062c182f23905ae8");
  assert.equal(expected.counts.periodRows, 84);
  assert.equal(expected.counts.cacheVersionRows, 1);
  assert.equal(expected.counts.scopeCacheRows, 1421);
  assert.equal(expected.counts.denseRpgCacheRows, 12789);
  assert.equal(expected.counts.completeCandidateRows, 14295);
  assert.deepEqual(expected.identity_pending, {
    fact_count: 120, sales_nok: 114876, quantity: 951,
  });
});

test("MappingAsOfDate accepts the exact ISO audit date or its Excel serial", () => {
  assert.match(script, /acceptedMappingDate\(version\[14\]\)/);
  assert.match(script, /text\(value\) === EXPECTED_MAPPING_AS_OF_DATE \|\| number\(value\) === EXPECTED_MAPPING_AS_OF_SERIAL/);
});

test("Excel-coerced period dates canonicalize back to ISO for the cache fingerprint", () => {
  assert.match(script, /canonicalIsoDate\(row\[3\]\)/);
  assert.match(script, /canonicalIsoDate\(row\[4\]\)/);
  assert.match(script, /serial - 25569/);
});

test("Office Script fails stale preflight before creating a staging surface", () => {
  const prepare = functionBody("prepareMaterialization");
  assert.ok(prepare.indexOf("validateLiveState(workbook)") < prepare.indexOf("workbook.addWorksheet"));
  assert.ok(prepare.indexOf("assertAcceptedLiveState(live, payload)") < prepare.indexOf("workbook.addWorksheet"));
  assert.match(script, /EXPECTED_MAPPING_CONTENT = "MCF-759cc92c4304a913"/);
  assert.match(script, /EXPECTED_CATALOG_CONTENT = "ICC-5644a77c18a97437"/);
  assert.match(script, /EXPECTED_PHASE2A_MAPPING = "MAP-34202a7a1a922bd0"/);
  assert.match(script, /EXPECTED_MAPPING_AS_OF_DATE = "2026-08-12"/);
  assert.match(script, /EXPECTED_MAPPING_AS_OF_SERIAL = 46246/);
});

test("Office Script is staging-atomic, inactive and idempotent", () => {
  assert.match(script, /"Candidate"/);
  assert.match(script, /"Not Active"/);
  assert.match(script, /"Already Materialized"/);
  assert.match(script, /priorStaging\.delete\(\)/);
  assert.match(script, /setVisibility\(ExcelScript\.SheetVisibility\.hidden\)/);
  assert.match(script, /staging\.setName\(FINAL_SHEET\)/);
  assert.doesNotMatch(script, /ActivationState[^\n]*"Active"|cacheStatus[^\n]*"Active"/);
});

test("materializer reconciles states and validates additive range fixtures from staged rows", () => {
  assert.match(script, /assertMetricEqual\(states\.identityPending/);
  assert.match(script, /assertMetricEqual\(states\.conflict, emptyMetric\(\), "Conflict"\)/);
  assert.match(script, /assertMetricEqual\(states\.inactiveTarget, emptyMetric\(\), "Inactive Target"\)/);
  assert.match(script, /assertMetricEqual\(mappedRpg, states\.mapped/);
  assert.match(script, /W31/);
  assert.match(script, /W01-W32/);
  assert.match(script, /W20-W30/);
  assert.match(script, /current\.numerator - comparisonShare \* current\.denominator/);
  assert.doesNotMatch(script, /average|AVERAGE/i);
});

test("materializer cannot cut over Performance, Reports, imports, facts or selectors", () => {
  assert.doesNotMatch(script, /tblSalesFacts|tblImports|tblMetricRPGResults|_Metric_Calc/);
  assert.doesNotMatch(script, /getWorksheet\("Performance"\)|getWorksheet\("Reports"\)/);
  assert.doesNotMatch(script, /period selector|Period Selector|supersed/i);
  assert.doesNotMatch(script, /setFormula|setFormulas|setPosition|setSelected/);
  assert.match(script, /tblPerformanceInteractionQA/);
});

test("payload generator is read-only and uses one exact caller-supplied corpus path", () => {
  assert.match(audit, /Provide one exact fixture corpus path/);
  assert.match(audit, /readOnly: true/);
  assert.doesNotMatch(audit, /\/Users\/|process\.env\.(HOME|CODEX_HOME)|homedir\(/);
  assert.doesNotMatch(audit, /writeFile|copyFile|rename\(|unlink\(|rm\(/);
  assert.doesNotMatch(audit, /Pulse_Current\.xlsx|tblMetricRPGResults|Power Automate/);
});

test("Office Scripts compiler/performance compatibility guards remain clean", () => {
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

function candidate() {
  const versionManifest = Object.fromEntries(WEEKLY_CACHE_COLUMNS.version.map(([, key]) => [key, `${key}-value`]));
  Object.assign(versionManifest, {
    cacheVersion: "WCV-test", cacheStatus: "Candidate", activationState: "Not Active",
    validationStatus: "PASS", periodRowCount: 2, scopeCacheRowCount: 1,
    denseRpgCacheRowCount: 8, cacheFingerprint: "WCC-test",
  });
  return {
    versionManifest,
    validation: { status: "PASS" },
    periodManifest: [row(WEEKLY_CACHE_COLUMNS.period, 1), row(WEEKLY_CACHE_COLUMNS.period, 2)],
    scopeCacheRows: [row(WEEKLY_CACHE_COLUMNS.scope, 1)],
    weeklyRpgCacheRows: Array.from({ length: 8 }, (_, index) => row(WEEKLY_CACHE_COLUMNS.rpg, index + 1)),
  };
}

function row(columns, suffix) {
  return Object.fromEntries(columns.map(([, key]) => [key, `${key}-${suffix}`]));
}

function functionBody(name) {
  const start = script.indexOf(`function ${name}(`);
  assert.notEqual(start, -1, `${name} must exist`);
  const open = script.indexOf("{", start);
  let depth = 0;
  for (let index = open; index < script.length; index += 1) {
    if (script[index] === "{") depth += 1;
    else if (script[index] === "}") {
      depth -= 1;
      if (depth === 0) return script.slice(open + 1, index);
    }
  }
  throw new Error(`Unbalanced function ${name}`);
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
