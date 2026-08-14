import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  IDENTITY_PENDING,
  TEST_DEPARTMENT_SOURCE_NAMES,
  buildWeeklyIdentityPreflight,
} from "../src/imports/weekly-identity-preflight.mjs";
import {
  WEEKLY_SALES_HEADERS,
  parseWeeklySalesMatrix,
} from "../src/imports/weekly-sales-parser.mjs";

const expectedPath = new URL("./expected-build-0.3.0-weekly-identity.json", import.meta.url);
const preflightPath = new URL("../src/imports/weekly-identity-preflight.mjs", import.meta.url);
const auditPath = new URL("../src/imports/audit-weekly-identity-preflight.mjs", import.meta.url);

test("known exact keys reuse stable IDs and approved product override precedence", () => {
  const catalog = fixtureCatalog();
  catalog.mappingRules.push(rule("MAP-000003", "Product", "PRD-000001", "RPG-0002"));
  const report = parsedReport([
    row("Known Restaurant", "Main A", "Sub A", "Known Product", 2, 200),
  ]);

  const result = buildWeeklyIdentityPreflight({ parsedReports: [report], catalogs: catalog });

  assert.equal(result.knownIdentitiesReused.restaurants.distinctIdentityCount, 1);
  assert.equal(result.knownIdentitiesReused.products.distinctIdentityCount, 1);
  assert.equal(result.knownIdentitiesReused.classifications.distinctIdentityCount, 1);
  assert.equal(result.newIdentityCandidates.products.length, 0);
  assert.deepEqual(result.rowAssignments[0], {
    sourceRowId: report.rows[0].sourceRowId,
    sourcePeriodKey: "PERIOD-2026-01-05-2026-01-11",
    restaurantId: "RST-0001",
    productId: "PRD-000001",
    sourceClassificationId: "SCL-00001",
    identityState: "Stable",
    identityPendingReason: "",
    mappingStatus: "Mapped",
    effectiveReportingGroupId: "RPG-0002",
  });
});

test("new exact identities receive deterministic candidates without implying a mapping", () => {
  const report = parsedReport([
    row("New Restaurant", "Main A", "Sub A", "A New Mapped", 1, 100),
    row("New Restaurant", "New Main", "New Sub", "B New Unmapped", 3, 450),
  ]);

  const result = buildWeeklyIdentityPreflight({ parsedReports: [report], catalogs: fixtureCatalog() });
  const [mapped, unmapped] = result.newIdentityCandidates.products;

  assert.equal(result.newIdentityCandidates.restaurants[0].restaurantId, "RST-0002");
  assert.equal(mapped.productId, "PRD-000002");
  assert.equal(mapped.sourceClassificationId, "SCL-00001");
  assert.equal(unmapped.productId, "PRD-000003");
  assert.equal(unmapped.sourceClassificationId, "SCL-00003");
  assert.equal(result.newIdentityCandidates.classifications[0].sourceClassificationId, "SCL-00003");
  assert.equal(result.rowAssignments[0].mappingStatus, "Mapped");
  assert.equal(result.rowAssignments[0].effectiveReportingGroupId, "RPG-0001");
  assert.equal(result.rowAssignments[1].mappingStatus, "Unmapped");
  assert.equal(result.rowAssignments[1].effectiveReportingGroupId, "");
  assert.equal(result.reconciliation.status, "PASS");
  assert.deepEqual(result.reconciliation.duplicateProposedStableIds, []);
  assert.deepEqual(result.reconciliation.duplicateProposedStableKeys, []);
});

test("the two Test Department spellings remain distinct and Performance-disabled", () => {
  const report = parsedReport(TEST_DEPARTMENT_SOURCE_NAMES.map((name, index) =>
    row(name, "Main A", "Sub A", "Known Product", index + 1, (index + 1) * 100)));

  const result = buildWeeklyIdentityPreflight({ parsedReports: [report], catalogs: fixtureCatalog() });
  const candidates = result.newIdentityCandidates.restaurants;

  assert.deepEqual(candidates.map(value => value.sourceRestaurantName), TEST_DEPARTMENT_SOURCE_NAMES);
  assert.deepEqual(candidates.map(value => value.restaurantId), ["RST-0002", "RST-0003"]);
  assert.ok(candidates.every(value => value.reportingEnabled === "No"));
  assert.ok(candidates.every(value => value.testDepartment));
});

test("new multi-hierarchy ProductKey is stable but hierarchy remains explicitly pending", () => {
  const report = parsedReport([
    row("Known Restaurant", "Main A", "Sub A", "New Multi Path", 1, 100),
    row("Known Restaurant", "Main B", "Sub B", "New Multi Path", 2, 250),
  ]);

  const result = buildWeeklyIdentityPreflight({ parsedReports: [report], catalogs: fixtureCatalog() });
  const candidate = result.newIdentityCandidates.products[0];

  assert.equal(candidate.productId, "PRD-000002");
  assert.equal(candidate.hierarchyStatus, IDENTITY_PENDING);
  assert.equal(candidate.sourceClassificationId, "");
  assert.deepEqual(candidate.observedHierarchyPaths, [
    "SRC-TEST-SALES || Main A || Sub A",
    "SRC-TEST-SALES || Main B || Sub B",
  ]);
  assert.ok(result.rowAssignments.every(value => value.productId === "PRD-000002"));
  assert.ok(result.rowAssignments.every(value => value.identityState === IDENTITY_PENDING));
  assert.ok(result.rowAssignments.every(value => value.mappingStatus === IDENTITY_PENDING));
  assert.deepEqual(result.identityStateCoverage[IDENTITY_PENDING], {
    factCount: 2,
    salesNok: 350,
    quantity: 3,
  });
});

test("colliding exact catalog keys are never guessed and carry source impact", () => {
  const catalog = fixtureCatalog();
  catalog.products.push({ ...catalog.products[0], productId: "PRD-000002" });
  const report = parsedReport([
    row("Known Restaurant", "Main A", "Sub A", "Known Product", 4, 800),
  ]);

  const result = buildWeeklyIdentityPreflight({ parsedReports: [report], catalogs: catalog });
  const pending = result.identityPendingItems.find(value => value.entityType === "Product");

  assert.deepEqual(pending.candidateStableIds, ["PRD-000001", "PRD-000002"]);
  assert.deepEqual(pending.impact, { factCount: 1, salesNok: 800, quantity: 4 });
  assert.equal(result.rowAssignments[0].productId, "");
  assert.equal(result.rowAssignments[0].identityState, IDENTITY_PENDING);
  assert.equal(result.rowAssignments[0].mappingStatus, IDENTITY_PENDING);
});

test("a current Product with missing hierarchy authority is visible as Identity Pending", () => {
  const catalog = fixtureCatalog();
  catalog.classifications = catalog.classifications.filter(value => value.sourceClassificationId !== "SCL-00001");
  const report = parsedReport([
    row("Known Restaurant", "Main A", "Sub A", "Known Product", 4, 800),
  ]);

  const result = buildWeeklyIdentityPreflight({ parsedReports: [report], catalogs: catalog });
  const pending = result.identityPendingItems.find(value => value.stableId === "PRD-000001");

  assert.match(pending.reason, /missing or ambiguous/);
  assert.deepEqual(pending.impact, { factCount: 1, salesNok: 800, quantity: 4 });
  assert.equal(result.rowAssignments[0].productId, "PRD-000001");
  assert.equal(result.rowAssignments[0].mappingStatus, IDENTITY_PENDING);
});

test("current Product hierarchy remains authority and row order cannot change inheritance", () => {
  const rows = [
    row("Known Restaurant", "Main A", "Sub A", "Known Product", 1, 100),
    row("Known Restaurant", "Main B", "Sub B", "Known Product", 2, 300),
  ];
  const first = buildWeeklyIdentityPreflight({
    parsedReports: [parsedReport(rows)],
    catalogs: fixtureCatalog(),
  });
  const reversed = buildWeeklyIdentityPreflight({
    parsedReports: [parsedReport([...rows].reverse())],
    catalogs: fixtureCatalog(),
  });

  assert.ok(first.rowAssignments.every(value => value.productId === "PRD-000001"));
  assert.ok(first.rowAssignments.every(value => value.effectiveReportingGroupId === "RPG-0001"));
  assert.ok(reversed.rowAssignments.every(value => value.effectiveReportingGroupId === "RPG-0001"));
  assert.equal(first.hierarchyReview.length, 1);
  assert.equal(first.hierarchyReview[0].currentSourceClassificationId, "SCL-00001");
  assert.equal(first.hierarchyReview[0].alternatePaths[0].alternateOutcome.effectiveReportingGroupId, "RPG-0002");
  assert.equal(first.fingerprints.preflightFingerprint, reversed.fingerprints.preflightFingerprint);
});

test("candidate allocation and identity are independent of locator and input row order", () => {
  const rows = [
    row("New Restaurant", "New Main", "New Sub", "Zulu Product", 1, 100),
    row("New Restaurant", "New Main", "New Sub", "Alpha Product", 2, 200),
  ];
  const normal = parsedReport(rows, "source/2026.xlsx");
  const banana = parsedReport([...rows].reverse(), "banana.xlsx");
  const first = buildWeeklyIdentityPreflight({ parsedReports: [normal], catalogs: fixtureCatalog() });
  const second = buildWeeklyIdentityPreflight({ parsedReports: [banana], catalogs: fixtureCatalog() });

  assert.deepEqual(first.newIdentityCandidates, second.newIdentityCandidates);
  assert.deepEqual(first.fingerprints, second.fingerprints);
  assert.notEqual(normal.manifest.sourceLocator, banana.manifest.sourceLocator);
});

test("an accepted exact candidate catalog is idempotently reused on rerun", () => {
  const report = parsedReport([
    row("New Restaurant", "New Main", "New Sub", "New Product", 1, 100),
  ]);
  const catalog = fixtureCatalog();
  const first = buildWeeklyIdentityPreflight({ parsedReports: [report], catalogs: catalog });
  const accepted = {
    ...catalog,
    restaurants: [...catalog.restaurants, ...first.newIdentityCandidates.restaurants],
    products: [...catalog.products, ...first.newIdentityCandidates.products],
    classifications: [...catalog.classifications, ...first.newIdentityCandidates.classifications],
  };
  const second = buildWeeklyIdentityPreflight({ parsedReports: [report], catalogs: accepted });

  assert.equal(second.newIdentityCandidates.restaurants.length, 0);
  assert.equal(second.newIdentityCandidates.products.length, 0);
  assert.equal(second.newIdentityCandidates.classifications.length, 0);
  assert.equal(second.rowAssignments[0].restaurantId, first.newIdentityCandidates.restaurants[0].restaurantId);
  assert.equal(second.rowAssignments[0].productId, first.newIdentityCandidates.products[0].productId);
});

test("identity and mapping states account for every row and reconcile all source measures", () => {
  const report = parsedReport([
    row("Known Restaurant", "Main A", "Sub A", "Known Product", 1.25, 100.1),
    row("New Restaurant", "New Main", "New Sub", "New Product", 2.5, 200.2),
    row("New Restaurant", "Main A", "Sub A", "Another Product", 3.75, 300.3),
  ]);

  const result = buildWeeklyIdentityPreflight({ parsedReports: [report], catalogs: fixtureCatalog() });

  assert.equal(result.rowAssignments.length, 3);
  assert.deepEqual(result.sourceTotals, { factCount: 3, salesNok: 600.6, quantity: 7.5 });
  assert.equal(result.reconciliation.assignedFactCount, 3);
  assert.equal(result.reconciliation.assignedSalesNok, 600.6);
  assert.equal(result.reconciliation.assignedQuantity, 7.5);
  assert.equal(result.reconciliation.status, "PASS");
});

test("frozen 84-week identity evidence records the exact candidate contract", async () => {
  const expected = JSON.parse(await readFile(expectedPath, "utf8"));

  assert.equal(expected.report_count, 84);
  assert.deepEqual(expected.source, {
    fact_count: 245632,
    sales_nok: 484728367.25,
    quantity: 2469988.09,
  });
  assert.equal(expected.new_candidates.products, 193);
  assert.equal(expected.new_candidates.classifications, 19);
  assert.equal(expected.identity_pending_count, 6);
  assert.deepEqual(expected.hierarchy_review_product_ids, [
    "PRD-000689",
    "PRD-000027",
    "PRD-000104",
    "PRD-000233",
    "PRD-000296",
    "PRD-000365",
    "PRD-000449",
    "PRD-000596",
    "PRD-000870",
  ]);
  assert.equal(expected.mapping_states.Mapped.fact_count +
    expected.mapping_states.Unmapped.fact_count +
    expected.mapping_states.Conflict.fact_count +
    expected.mapping_states["Inactive Target"].fact_count +
    expected.mapping_states[IDENTITY_PENDING].fact_count, expected.source.fact_count);
  assert.equal(expected.reconciliation_status, "PASS");
  assert.deepEqual(expected.duplicate_proposed_stable_ids, []);
  assert.deepEqual(expected.duplicate_proposed_stable_keys, []);
});

test("preflight remains a read-only caller-supplied audit with no cache or workbook mutation", async () => {
  const [preflight, audit] = await Promise.all([
    readFile(preflightPath, "utf8"),
    readFile(auditPath, "utf8"),
  ]);

  assert.match(audit, /Provide one exact fixture corpus path/);
  assert.match(audit, /readOnly: true/);
  assert.doesNotMatch(audit, /\/Users\/|process\.env\.(HOME|CODEX_HOME)|homedir\(/);
  assert.doesNotMatch(`${preflight}\n${audit}`, /writeFile|rename\(|unlink\(|rm\(|copyFile/);
  assert.doesNotMatch(`${preflight}\n${audit}`, /Performance|Power Automate|compact cache|supersede/i);
});

function fixtureCatalog() {
  return {
    sourceSystemId: "SRC-TEST-SALES",
    catalogAsOfDate: "2026-08-12",
    catalogAsOfExcelSerial: 46246,
    mappingFingerprint: "MAP-fixture",
    restaurants: [{
      restaurantId: "RST-0001",
      sourceSystemId: "SRC-TEST-SALES",
      sourceRestaurantName: "Known Restaurant",
      displayName: "Known Restaurant",
      status: "Active",
      reportingEnabled: "Yes",
    }],
    classifications: [
      classification("SCL-00001", "Main A", "Sub A"),
      classification("SCL-00002", "Main B", "Sub B"),
    ],
    products: [{
      productId: "PRD-000001",
      sourceSystemId: "SRC-TEST-SALES",
      sourceProductName: "Known Product",
      salesAccount: "3000 - Sales 25%",
      sourceClassificationId: "SCL-00001",
      productKey: "SRC-TEST-SALES || Known Product || 3000 - Sales 25%",
      productStatus: "Active",
    }],
    reportingGroups: [
      { reportingGroupId: "RPG-0001", reportingGroupName: "Group A", active: "Yes", sortOrder: 10 },
      { reportingGroupId: "RPG-0002", reportingGroupName: "Group B", active: "Yes", sortOrder: 20 },
    ],
    mappingRules: [
      rule("MAP-000001", "SourceMainCategory", "SRC-TEST-SALES || Main || Main A", "RPG-0001"),
      rule("MAP-000002", "SourceMainCategory", "SRC-TEST-SALES || Main || Main B", "RPG-0002"),
    ],
  };
}

function classification(id, main, sub) {
  return {
    sourceClassificationId: id,
    sourceSystemId: "SRC-TEST-SALES",
    sourceMainCategory: main,
    sourceSubCategory: sub,
    sourceClassificationKey: `SRC-TEST-SALES || ${main} || ${sub}`,
    status: "Active",
  };
}

function rule(id, scopeType, nodeId, reportingGroupId) {
  return {
    mappingRuleId: id,
    sourceSystemId: "SRC-TEST-SALES",
    scopeType,
    nodeId,
    targetReportingGroupId: reportingGroupId,
    effectiveFrom: 45000,
    effectiveTo: null,
    status: "Active",
    ruleAction: "Map",
  };
}

function parsedReport(rows, sourceLocator = "fixture.xlsx") {
  return parseWeeklySalesMatrix([
    ["Period: 2026-01-05 - 2026-01-11"],
    [...WEEKLY_SALES_HEADERS],
    ...rows,
  ], { sourceLocator });
}

function row(restaurant, main, sub, item, quantity, amount) {
  return [restaurant, main, sub, "3000 - Sales 25%", item, quantity, amount];
}
