import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";
import { findRuleConflicts, resolveProduct } from "../src/mapping/hierarchical-resolver.mjs";

const groups = [
  { reportingGroupId: "RPG-0001", active: "Yes" },
  { reportingGroupId: "RPG-0002", active: "Yes" },
  { reportingGroupId: "RPG-0003", active: "No" },
];
const product = {
  productId: "PRD-1",
  sourceSystemId: "SRC-1",
  mainNodeId: "SRC-1 || Drinks",
  subNodeId: "SCL-1",
};
const rule = (id, scopeType, nodeId, target, extra = {}) => ({
  mappingRuleId: id,
  sourceSystemId: "SRC-1",
  scopeType,
  nodeId,
  targetReportingGroupId: target,
  effectiveFrom: 1,
  effectiveTo: null,
  status: "Active",
  ...extra,
});

test("nearest mapped ancestor is inherited", () => {
  const result = resolveProduct({
    product, groups, asOf: 10,
    rules: [rule("MAP-1", "SourceMainCategory", product.mainNodeId, "RPG-0001")],
  });
  assert.equal(result.effectiveReportingGroupId, "RPG-0001");
  assert.equal(result.resolutionState, "Inherited");
});

test("most-specific explicit mapping wins", () => {
  const rules = [
    rule("MAP-1", "SourceMainCategory", product.mainNodeId, "RPG-0001"),
    rule("MAP-2", "SourceSubCategory", product.subNodeId, "RPG-0002"),
    rule("MAP-3", "Product", product.productId, "RPG-0001"),
  ];
  assert.equal(resolveProduct({ product, groups, rules, asOf: 10 }).ruleId, "MAP-3");
});

test("subcategory overrides main category when no product rule exists", () => {
  const rules = [
    rule("MAP-1", "SourceMainCategory", product.mainNodeId, "RPG-0001"),
    rule("MAP-2", "SourceSubCategory", product.subNodeId, "RPG-0002"),
  ];
  const result = resolveProduct({ product, groups, rules, asOf: 10 });
  assert.equal(result.ruleId, "MAP-2");
  assert.equal(result.resolutionState, "Inherited");
});

test("product exception survives parent remapping", () => {
  const productOverride = rule("MAP-3", "Product", product.productId, "RPG-0002");
  const before = resolveProduct({
    product, groups, asOf: 10,
    rules: [rule("MAP-1", "SourceMainCategory", product.mainNodeId, "RPG-0001"), productOverride],
  });
  const after = resolveProduct({
    product, groups, asOf: 10,
    rules: [rule("MAP-4", "SourceMainCategory", product.mainNodeId, "RPG-0002"), productOverride],
  });
  assert.equal(before.ruleId, "MAP-3");
  assert.equal(after.ruleId, "MAP-3");
});

test("inactive or cleared override restores ancestor", () => {
  const rules = [
    rule("MAP-1", "SourceMainCategory", product.mainNodeId, "RPG-0001"),
    rule("MAP-2", "Product", product.productId, "RPG-0002", { status: "Inactive" }),
  ];
  assert.equal(resolveProduct({ product, groups, rules, asOf: 10 }).ruleId, "MAP-1");
});

test("explicit Product exclusion beats inherited parent mapping", () => {
  const rules = [
    rule("MAP-1", "SourceMainCategory", product.mainNodeId, "RPG-0001"),
    rule("MAP-2", "Product", product.productId, "", { ruleAction: "Exclude" }),
  ];
  const result = resolveProduct({ product, groups, rules, asOf: 10 });
  assert.deepEqual(result, {
    effectiveReportingGroupId: "",
    resolutionSource: "Product",
    resolutionState: "Explicit exclusion",
    resolutionStatus: "Unmapped",
    ruleId: "MAP-2",
  });
});

test("deactivating Product exclusion restores inherited mapping", () => {
  const rules = [
    rule("MAP-1", "SourceMainCategory", product.mainNodeId, "RPG-0001"),
    rule("MAP-2", "Product", product.productId, "", { ruleAction: "Exclude", status: "Inactive" }),
  ];
  const result = resolveProduct({ product, groups, rules, asOf: 10 });
  assert.equal(result.ruleId, "MAP-1");
  assert.equal(result.resolutionStatus, "Mapped");
});

test("overlapping Product Map and Exclude rules surface a conflict", () => {
  const rules = [
    rule("MAP-1", "Product", product.productId, "RPG-0001"),
    rule("MAP-2", "Product", product.productId, "", { ruleAction: "Exclude" }),
  ];
  assert.deepEqual(findRuleConflicts(rules), [["MAP-1", "MAP-2"]]);
  assert.equal(resolveProduct({ product, groups, rules, asOf: 10 }).resolutionStatus, "Conflict");
});

test("unmapped remains visible", () => {
  assert.equal(resolveProduct({ product, groups, rules: [], asOf: 10 }).resolutionStatus, "Unmapped");
});

test("same-node overlapping rules are conflicts", () => {
  const rules = [
    rule("MAP-1", "Product", product.productId, "RPG-0001", { effectiveFrom: 1, effectiveTo: 20 }),
    rule("MAP-2", "Product", product.productId, "RPG-0002", { effectiveFrom: 10, effectiveTo: 30 }),
  ];
  assert.deepEqual(findRuleConflicts(rules), [["MAP-1", "MAP-2"]]);
  assert.equal(resolveProduct({ product, groups, rules, asOf: 15 }).resolutionStatus, "Conflict");
});

test("non-overlapping effective intervals do not conflict", () => {
  const rules = [
    rule("MAP-1", "Product", product.productId, "RPG-0001", { effectiveFrom: 1, effectiveTo: 9 }),
    rule("MAP-2", "Product", product.productId, "RPG-0002", { effectiveFrom: 10, effectiveTo: 20 }),
  ];
  assert.deepEqual(findRuleConflicts(rules), []);
  assert.equal(resolveProduct({ product, groups, rules, asOf: 10 }).ruleId, "MAP-2");
});

test("inactive target is surfaced", () => {
  const rules = [rule("MAP-1", "Product", product.productId, "RPG-0003")];
  assert.equal(resolveProduct({ product, groups, rules, asOf: 10 }).resolutionStatus, "Inactive Target");
});

test("classification cannot change fact totals", () => {
  const facts = [
    { productId: "PRD-1", sales: 125, quantity: 2 },
    { productId: "PRD-2", sales: 75, quantity: 1 },
  ];
  const total = facts.reduce((a, fact) => ({ sales: a.sales + fact.sales, quantity: a.quantity + fact.quantity }), { sales: 0, quantity: 0 });
  const mapped = { sales: facts[0].sales, quantity: facts[0].quantity };
  const unmapped = { sales: facts[1].sales, quantity: facts[1].quantity };
  assert.equal(mapped.sales + unmapped.sales, total.sales);
  assert.equal(mapped.quantity + unmapped.quantity, total.quantity);
});

test("Office Script list validation uses resilient literal value sources", () => {
  const script = fs.readFileSync(
    new URL("../office-scripts/Build_0_3_0_Phase1.ts", import.meta.url),
    "utf8",
  );
  const validationBody = script.match(
    /function wireMappingValidation[\s\S]*?\n}\n\nfunction applyListValidation/,
  )?.[0] ?? "";

  assert.ok(validationBody, "wireMappingValidation function must exist");
  assert.equal((validationBody.match(/applyListValidation\(/g) ?? []).length, 7);
  assert.equal((validationBody.match(/,activeGroupSource,/g) ?? []).length, 2);
  assert.doesNotMatch(validationBody, /setRule\(\{list:\{[^}]*source:(?:actionSource|scopeSource|activeGroupSource|statusSource|ruleActionSource)/);
  assert.match(script, /sourceRange\.getValues\(\)\.forEach/);
  assert.match(script, /const source=items\.join\(","\)/);
  assert.match(script, /validation\.setRule\(\{list:\{inCellDropDown:true,source:source\}\}\)/);
  assert.match(script, /catch\(error\)\{failures\.push/);
  assert.match(validationBody, /PUL-0301-013/);
  assert.match(validationBody, /sheet\.getRange\("E8"\)\.setValue\(message\)/);
  assert.match(script, /Dropdown validation ready \(7\/7\)/);
  assert.match(script, /"RuleAction"/);
  assert.match(script, /resolutionState: "Explicit exclusion", resolutionStatus: "Unmapped"/);
});

test("Office Script avoids unsupported Map and Set iterator constructs", () => {
  const script = fs.readFileSync(
    new URL("../office-scripts/Build_0_3_0_Phase1.ts", import.meta.url),
    "utf8",
  );
  const collectionNames = "groupById|seen|classificationById|context|resultByProduct|main|subs|mappedIds|map";

  assert.doesNotMatch(script, /Array\.from\s*\(/);
  assert.doesNotMatch(script, /\.(?:entries|keys|values)\s*\(/);
  assert.doesNotMatch(script, new RegExp(`for\\s*\\([^)]*\\bof\\s+(?:${collectionNames})\\b`));
  assert.doesNotMatch(script, new RegExp(`\\.\\.\\.(?:${collectionNames})\\b`));
  assert.match(script, /context\.forEach\s*\(/);
  assert.match(script, /main\.forEach\s*\(/);
  assert.match(script, /subs\.forEach\s*\(/);
});
