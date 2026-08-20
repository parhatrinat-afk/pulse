import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  buildWeeklyMappingAttentionProjection,
} from "../src/mapping/weekly-mapping-attention.mjs";
import {
  buildWeeklyMappingAttentionMaterializationPlan,
  weeklyMappingAttentionChunk,
} from "../src/mapping/weekly-mapping-attention-materialization.mjs";

const expectedPath = new URL("./expected-build-0.3.0-weekly-mapping-attention.json", import.meta.url);
const officeScriptPath = new URL("../office-scripts/Materialize_Weekly_Mapping_Attention.ts", import.meta.url);
const phase1Path = new URL("../office-scripts/Build_0_3_0_Phase1.ts", import.meta.url);

test("weekly Mapping attention is one deterministic row per stable ProductID", () => {
  const projection = fixtureProjection();
  assert.equal(projection.validationStatus, "PASS");
  assert.equal(projection.totalProductCount, 5);
  assert.equal(projection.existingProductCount, 4);
  assert.equal(projection.weeklyAddedProductCount, 1);
  assert.deepEqual(projection.rows.map(row => row.productId), [
    "PRD-000001", "PRD-000002", "PRD-000003", "PRD-000004", "PRD-000005",
  ]);
  assert.deepEqual(projection.rows.map(row => row.resolutionType), [
    "Inherited Main", "Inherited Subcategory", "Explicit Product", "Explicit exclusion",
    "Identity Pending",
  ]);
  assert.deepEqual(projection.sourceTotals, { factCount: 5, salesNok: 1500, quantity: 15 });
  assert.deepEqual(projection.projectionTotals, projection.sourceTotals);
  assert.deepEqual(projection.duplicateProductIds, []);
  assert.deepEqual(projection.duplicateProductKeys, []);
});

test("Identity Pending stays distinct from Unmapped with hierarchy alternatives and impact", () => {
  const projection = fixtureProjection();
  const pending = projection.rows.find(row => row.productId === "PRD-000005");
  assert.equal(pending.mappingStatus, "Identity Pending");
  assert.equal(pending.effectiveReportingGroupId, "");
  assert.equal(pending.historicalFactCount, 1);
  assert.equal(pending.historicalSalesNok, 500);
  assert.match(pending.hierarchyAttention, /Identity Pending/);
  assert.match(pending.hierarchyAlternatives, /Main A/);
  assert.equal(projection.stateCoverage.Unmapped.productCount, 1);
  assert.equal(projection.stateCoverage["Identity Pending"].productCount, 1);
});

test("projection consumes accepted resolver output and never creates mapping semantics", () => {
  const first = fixtureProjection();
  const second = fixtureProjection({ reverseSourceRows: true });
  assert.equal(first.projectionFingerprint, second.projectionFingerprint);
  assert.deepEqual(first.rows, second.rows);
  assert.equal(first.rows.find(row => row.productId === "PRD-000003").winningRuleId, "MAP-000003");
  assert.equal(first.rows.find(row => row.productId === "PRD-000004").mappingStatus, "Unmapped");
});

test("materialization plan is hidden-table shaped, bounded, and tied to Active authority", () => {
  const projection = fixtureProjection();
  const cache = fixtureCache(projection);
  const plan = buildWeeklyMappingAttentionMaterializationPlan(cache);
  assert.equal(plan.sheetName, "_Weekly_Mapping_Attention");
  assert.equal(plan.stagingSheetName, "_Weekly_Mapping_Attn_Stage");
  assert.ok(plan.stagingSheetName.length <= 31);
  assert.equal(plan.sections.control.tableName, "tblWeeklyMappingAttentionControl");
  assert.equal(plan.sections.products.tableName, "tblWeeklyMappingAttention");
  assert.equal(plan.sections.products.rowCount, 5);
  assert.equal(plan.sections.products.headers[0], "ProductID");
  assert.equal(plan.sections.products.headers[1], "SourceSystemID");
  assert.equal(plan.throughPeriodLabel, "2026 W33");
  const chunk = weeklyMappingAttentionChunk(plan, "products", 2, 2);
  assert.equal(chunk.offset, 2);
  assert.equal(chunk.startRow, 8);
  assert.equal(chunk.values.length, 2);
  assert.throws(() => buildWeeklyMappingAttentionMaterializationPlan({
    ...cache, versionManifest: { ...cache.versionManifest, cacheStatus: "Candidate" },
  }), /Active \/ Active/);
});

test("frozen 85-week evidence records exact universe, coverage, and review identities", async () => {
  const expected = JSON.parse(await readFile(expectedPath, "utf8"));
  assert.equal(expected.productUniverse.total, 1237);
  assert.equal(expected.productUniverse.existingCatalog, 1041);
  assert.equal(expected.productUniverse.weeklyAdded, 196);
  assert.deepEqual(expected.stateCoverage.Mapped,
    { productCount: 929, factCount: 229190, salesNok: 479649885.1, quantity: 2367590.46 });
  assert.deepEqual(expected.stateCoverage.Unmapped,
    { productCount: 302, factCount: 19256, salesNok: 10595207.65, quantity: 131052.98 });
  assert.deepEqual(expected.stateCoverage["Identity Pending"],
    { productCount: 6, factCount: 126, salesNok: 120048, quantity: 999 });
  assert.equal(expected.identityPending.length, 6);
  assert.equal(expected.hierarchyAttentionProductIds.length, 9);
  assert.ok(expected.hierarchyAttentionProductIds.includes("PRD-000689"));
});

test("Office Script remains compatibility-safe and Phase 1 restores health without redesign", async () => {
  const [materializer, phase1] = await Promise.all([
    readFile(officeScriptPath, "utf8"), readFile(phase1Path, "utf8"),
  ]);
  const combined = `${materializer}\n${phase1}`;
  assert.doesNotMatch(combined, /for\s*\([^)]*\bof\s+\w+\.(entries|keys|values)\s*\(/);
  assert.doesNotMatch(combined, /\.entries\s*\(|\.keys\s*\(|\.values\s*\(/);
  assert.doesNotMatch(materializer, /Performance.*setFormula|Reports.*setFormula|tblMappingRules.*setValues/s);
  assert.match(materializer, /_Weekly_Mapping_Attention/);
  assert.match(materializer, /tblWeeklyMappingAttention/);
  assert.match(materializer, /headers\.Result/);
  assert.match(materializer, /A4:N9/);
  assert.match(phase1, /restoreWeeklyMappingHealthBlock/);
  assert.match(phase1, /Performance refresh required/);
  assert.doesNotMatch(combined, /Power Automate|Month|Quarter|Phase 3/);
});

function fixtureProjection({ reverseSourceRows = false } = {}) {
  const classifications = [
    classification("SCL-00001", "Main A", "Sub A"),
    classification("SCL-00002", "Main A", "Sub B"),
  ];
  const products = [
    product("PRD-000001", "SCL-00001"), product("PRD-000002", "SCL-00002"),
    product("PRD-000003", "SCL-00001"), product("PRD-000004", "SCL-00001"),
    { ...product("PRD-000005", ""), hierarchyStatus: "Identity Pending",
      observedHierarchyPaths: ["SRC || Main B || Sub B", "SRC || Main A || Sub A"] },
  ];
  const mappingByProduct = new Map([
    ["PRD-000001", resolution("Mapped", "SourceMainCategory", "Inherited", "RPG-0001", "MAP-000001")],
    ["PRD-000002", resolution("Mapped", "SourceSubCategory", "Inherited", "RPG-0002", "MAP-000002")],
    ["PRD-000003", resolution("Mapped", "Product", "Explicit", "RPG-0002", "MAP-000003")],
    ["PRD-000004", resolution("Unmapped", "Product", "Explicit exclusion", "", "MAP-000004")],
    ["PRD-000005", resolution("Identity Pending", "Identity Pending", "Identity Pending", "", "")],
  ]);
  const sourceRows = products.map((row, index) => ({
    sourceRowId: `ROW-${index + 1}`, factCount: 1, salesNok: (index + 1) * 100,
    quantity: index + 1,
  }));
  const rows = reverseSourceRows ? [...sourceRows].reverse() : sourceRows;
  const rowAssignments = products.map((row, index) => ({
    sourceRowId: `ROW-${index + 1}`, productId: row.productId,
    identityState: row.productId === "PRD-000005" ? "Identity Pending" : "Stable",
    mappingStatus: mappingByProduct.get(row.productId).resolutionStatus,
  }));
  return buildWeeklyMappingAttentionProjection({
    products, classifications,
    reportingGroups: [
      { reportingGroupId: "RPG-0001", reportingGroupName: "Group A" },
      { reportingGroupId: "RPG-0002", reportingGroupName: "Group B" },
    ],
    mappingByProduct, sourceRows: rows, rowAssignments,
    existingCatalogProductIds: products.slice(0, 4).map(row => row.productId), hierarchyReview: [],
  });
}

function fixtureCache(projection) {
  return {
    versionManifest: {
      cacheVersion: "WCV-test", cacheFingerprint: "WCC-test", cacheStatus: "Active",
      activationState: "Active", validationStatus: "PASS", sourceCorpusFingerprint: "WSC-test",
      mappingContentFingerprint: "MCF-test", catalogContentFingerprint: "ICC-test",
      identityPreflightFingerprint: "IDP-test",
    },
    periodManifest: [{ isoYear: 2026, isoWeek: 32 }, { isoYear: 2026, isoWeek: 33 }],
    validation: { corpusReconciliation: { source: projection.sourceTotals } },
    mappingAttentionProjection: projection,
  };
}

function classification(sourceClassificationId, sourceMainCategory, sourceSubCategory) {
  return { sourceClassificationId, sourceSystemId: "SRC", sourceMainCategory, sourceSubCategory };
}

function product(productId, sourceClassificationId) {
  return {
    productId, sourceSystemId: "SRC", sourceProductName: `Item ${productId}`,
    salesAccount: "3000", sourceClassificationId, productKey: `SRC || ${productId} || 3000`,
  };
}

function resolution(resolutionStatus, resolutionSource, resolutionState, groupId, ruleId) {
  return {
    resolutionStatus, resolutionSource, resolutionState,
    effectiveReportingGroupId: groupId, ruleId,
  };
}
