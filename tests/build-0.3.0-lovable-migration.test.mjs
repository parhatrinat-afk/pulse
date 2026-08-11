import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";
import {
  LOVABLE_MIGRATION_DECISIONS,
  LOVABLE_PRODUCT_EXCLUSIONS,
  planLovableMigration,
  semanticKey,
} from "../src/mapping/lovable-migration.mjs";

const expected = JSON.parse(fs.readFileSync(
  new URL("./expected-build-0.3.0-lovable-migration.json", import.meta.url),
  "utf8",
));

const existingAddOnsRule = {
  mappingRuleId: "MAP-000001",
  sourceSystemId: "SRC-TEST-SALES",
  scopeType: "SourceMainCategory",
  nodeId: "SRC-TEST-SALES || Main || Add-ons",
  targetReportingGroupId: "RPG-0001",
  effectiveFrom: 1,
  effectiveTo: null,
  status: "Active",
};

test("frozen migration contract represents exactly the approved 129 decisions", () => {
  const counts = LOVABLE_MIGRATION_DECISIONS.reduce((result, rule) => {
    const key = `${rule.scopeType}:${rule.ruleAction}`;
    result[key] = (result[key] ?? 0) + 1;
    return result;
  }, {});
  assert.equal(LOVABLE_MIGRATION_DECISIONS.length, expected.logical_decisions.total);
  assert.deepEqual(counts, {
    "SourceMainCategory:Map": expected.logical_decisions.source_main_category_maps,
    "SourceSubCategory:Map": expected.logical_decisions.source_subcategory_maps,
    "Product:Map": expected.logical_decisions.product_maps,
    "Product:Exclude": expected.logical_decisions.product_exclusions,
  });
  assert.equal(new Set(LOVABLE_MIGRATION_DECISIONS.map(semanticKey)).size, 129);
  assert.equal(new Set(LOVABLE_MIGRATION_DECISIONS.map(rule =>
    [rule.sourceSystemId, rule.scopeType, rule.nodeId].join("|"),
  )).size, 129);
});

test("Product exclusions are exactly the eight approved stable ProductIDs", () => {
  assert.deepEqual([...LOVABLE_PRODUCT_EXCLUSIONS].sort(), [
    "PRD-000220", "PRD-000221", "PRD-000259", "PRD-000260",
    "PRD-000546", "PRD-000566", "PRD-000567", "PRD-000942",
  ]);
  const exclusions = LOVABLE_MIGRATION_DECISIONS.filter(rule => rule.ruleAction === "Exclude");
  assert.ok(exclusions.every(rule =>
    rule.scopeType === "Product" && rule.targetReportingGroupId === "" && rule.nodeId.startsWith("PRD-"),
  ));
});

test("idempotent planner reuses MAP-000001 and never duplicates migrated rules", () => {
  const first = planLovableMigration({ existingRules: [existingAddOnsRule], asOf: 50000 });
  assert.deepEqual(first.reused.map(row => row.mappingRuleId), ["MAP-000001"]);
  assert.equal(first.additions.length, 128);
  assert.equal(new Set(first.additions.map(rule => rule.mappingRuleId)).size, 128);

  const afterFirstRun = [existingAddOnsRule, ...first.additions];
  const second = planLovableMigration({ existingRules: afterFirstRun, asOf: 50000 });
  assert.equal(second.reused.length, 129);
  assert.equal(second.additions.length, 0);
  assert.equal(new Set(second.reused.map(row => row.mappingRuleId)).size, 129);
});

test("migration planner refuses an active same-node semantic collision", () => {
  assert.throws(() => planLovableMigration({
    existingRules: [{ ...existingAddOnsRule, targetReportingGroupId: "RPG-0002" }],
    asOf: 50000,
  }), /conflicts with approved decision/);
});

test("Office Script carries the same stable-ID contract without fuzzy runtime matching", () => {
  const script = fs.readFileSync(
    new URL("../office-scripts/Migrate_Lovable_Mapping.ts", import.meta.url),
    "utf8",
  );
  const compact = script.replace(/,\s+/g, ",");
  for (const rule of LOVABLE_MIGRATION_DECISIONS.filter(value => value.ruleAction === "Map")) {
    assert.ok(
      compact.includes(JSON.stringify([rule.nodeId, rule.targetReportingGroupId])),
      `Office Script is missing ${semanticKey(rule)}`,
    );
  }
  for (const productId of LOVABLE_PRODUCT_EXCLUSIONS) {
    assert.match(script, new RegExp(`"${productId}"`));
  }
  assert.match(script, /decisions\.length !== 129/);
  assert.match(script, /mainCount !== 70/);
  assert.match(script, /subCount !== 2/);
  assert.match(script, /productMapCount !== 49/);
  assert.match(script, /exclusionCount !== 8/);
  assert.match(script, /Existing Add-ons rule MAP-000001 was not recognized for reuse/);
  assert.doesNotMatch(script, /fuzzy|levenshtein|similarity|toLowerCase\(\)/i);
});

test("migration Office Script avoids unsupported Map and Set iterator constructs", () => {
  const script = fs.readFileSync(
    new URL("../office-scripts/Migrate_Lovable_Mapping.ts", import.meta.url),
    "utf8",
  );
  assert.doesNotMatch(script, /Array\.from\s*\(/);
  assert.doesNotMatch(script, /\.(?:entries|keys|values)\s*\(/);
  assert.doesNotMatch(script, /for\s*\([^)]*\bof\s+(?:displayByNode|active|seen)\b/);
  assert.doesNotMatch(script, /\.\.\.(?:displayByNode|active|seen)\b/);
});

test("accepted checkpoint totals reconcile exactly across states and Reporting Groups", () => {
  const states = Object.values(expected.states);
  assert.equal(sum(states, "fact_count"), expected.source.fact_count);
  assert.ok(almostEqual(sum(states, "sales_nok"), expected.source.sales_nok));
  assert.ok(almostEqual(sum(states, "quantity"), expected.source.quantity));

  const groups = Object.values(expected.reporting_groups);
  assert.equal(sum(groups, "fact_count"), expected.states.Mapped.fact_count);
  assert.ok(almostEqual(sum(groups, "sales_nok"), expected.states.Mapped.sales_nok));
  assert.ok(almostEqual(sum(groups, "quantity"), expected.states.Mapped.quantity));
  assert.equal(expected.states.Conflict.fact_count, 0);
  assert.equal(expected.states["Inactive Target"].fact_count, 0);
  assert.equal(percent(expected.states.Mapped.fact_count, expected.source.fact_count), expected.coverage_percent.facts);
  assert.equal(percent(expected.states.Mapped.sales_nok, expected.source.sales_nok), expected.coverage_percent.sales_nok);
  assert.equal(percent(expected.states.Mapped.quantity, expected.source.quantity), expected.coverage_percent.quantity);
});

test("accepted live Excel evidence is internally reconciled and identifies the mapping state", () => {
  const evidence = expected.live_excel_qa;
  assert.equal(evidence.status, "PASS");
  assert.match(evidence.mapping_fingerprint, /^MAP-[0-9a-f]{16}$/);
  assert.equal(evidence.highest_mapping_rule_id, "MAP-000133");
  assert.equal(evidence.all_reconciliation_scopes, "PASS");
  assert.equal(evidence.source_facts_unchanged, "PASS");

  const scopes = Object.values(evidence.scopes);
  assert.ok(scopes.every(scope => scope.mapped_facts + scope.unmapped_facts === scope.total_facts));
  assert.equal(sum(scopes, "mapped_facts"), expected.states.Mapped.fact_count);
  assert.equal(sum(scopes, "unmapped_facts"), expected.states.Unmapped.fact_count);
  assert.equal(sum(scopes, "total_facts"), expected.source.fact_count);
});

function sum(rows, field) {
  return rows.reduce((total, row) => total + row[field], 0);
}

function almostEqual(left, right) {
  return Math.abs(left - right) <= 0.000001;
}

function percent(part, total) {
  return Math.round((part / total) * 10000) / 100;
}
