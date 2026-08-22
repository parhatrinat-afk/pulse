import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";
import {
  buildReportingGroupMembershipView,
} from "../src/mapping/weekly-mapping-attention.mjs";

const scriptPath = new URL("../office-scripts/Build_0_3_0_Phase1.ts", import.meta.url);
const expectedPath = new URL("./expected-build-0.3.0-reporting-group-membership.json", import.meta.url);

test("reverse membership uses accepted mapped Product rows and active stable group IDs", () => {
  const view = fixtureView();
  assert.deepEqual(view.members.map(row => row.productId), ["PRD-000001", "PRD-000002"]);
  assert.deepEqual(view.members.map(row => row.mappingState), ["Inherited", "Custom"]);
  assert.deepEqual(view.overview.map(row => [row.reportingGroupId, row.productCount]), [
    ["RPG-0001", 1], ["RPG-0002", 1],
  ]);
  assert.deepEqual(view.totals, { productCount: 2, factCount: 5, salesNok: 575 });
});

test("Unmapped, exclusions, Identity Pending, Conflict and Inactive Target never appear as members", () => {
  const view = fixtureView();
  assert.equal(view.members.length, 2);
  assert.ok(view.members.every(row => row.reportingGroupName !== "Unmapped"));
  assert.ok(!view.members.some(row => ["PRD-000003", "PRD-000004", "PRD-000005", "PRD-000006"]
    .includes(row.productId)));
});

test("duplicate mapped ProductIDs and inactive targets are rejected rather than double-counted", () => {
  const rows = fixtureRows();
  assert.throws(() => buildReportingGroupMembershipView({
    rows: [...rows, { ...rows[0] }], reportingGroups: fixtureGroups(),
  }), /Mapped ProductID PRD-000001 is duplicated/);
  assert.throws(() => buildReportingGroupMembershipView({
    rows: [{ ...rows[0], effectiveReportingGroupId: "RPG-9999" }],
    reportingGroups: fixtureGroups(),
  }), /inactive or unknown RPG-9999/);
});

test("duplicate active business names fail before stable IDs can be conflated", () => {
  const groups = fixtureGroups().map(group => group.reportingGroupId === "RPG-0002"
    ? { ...group, reportingGroupName: "Add-ons" } : group);
  assert.throws(() => buildReportingGroupMembershipView({
    rows: fixtureRows(), reportingGroups: groups,
  }), /active ReportingGroupName Add-ons is duplicated/);
});

test("bounded membership capacity fails clearly before presentation truncates", () => {
  assert.throws(() => buildReportingGroupMembershipView({
    rows: fixtureRows().filter(row => row.mappingStatus === "Mapped"),
    reportingGroups: fixtureGroups(), memberCapacity: 0,
  }), /bounded capacity is 0/);
});

test("Office Script adds a read-only name-first Reporting Group mode in the shared workspace", async () => {
  const script = await readFile(scriptPath, "utf8");
  assert.match(script, /\["Source Category", "Reporting Group"\]/);
  assert.match(script, /browseBy: "Source Category"/);
  assert.match(script, /Reporting Group membership is read-only/);
  assert.match(script, /tblMappingReportingGroupOverview/);
  assert.match(script, /applyFinalMappingPresentation/);
  assert.match(script, /memberTable\.setName\("tblMappingMemberWorkspace"\)/);
  assert.match(script, /if \(table\) table\.delete\(\)/);
  assert.match(script, /tblMappingReportingGroupCatalog/);
  assert.match(script, /const reportingGroupMemberCapacity = 400/);
  assert.match(script, /product\.resolution\.resolutionStatus !== "Mapped"/);
  assert.match(script, /product\.resolution\.resolutionSource === "Product" \? "Custom" : "Inherited"/);
  assert.match(script, /_Mapping_Audit'!\$BU\$2:\$BU\$\$\{reportingGroupCatalogEndRow\}/);
  assert.match(script, /Mapping workspace dropdowns ready \(12\/12\)/);
  assert.match(script, /PUL-0301-033: Active Reporting Group name is duplicated/);
  const finalHeaders = script.match(/const memberHeaders = \["Select", "Level", "Item", "Main Category"[\s\S]*?\];/)?.[0] ?? "";
  assert.ok(finalHeaders, "shared business-first member headers must exist");
  assert.match(finalHeaders, /"Item".*"Main Category".*"Subcategory".*"Sales Account".*"Mapping state".*"Facts".*"Sales NOK"/s);
});

test("Reporting Group facade preserves the accepted Source Category mutation workspace", async () => {
  const script = await readFile(scriptPath, "utf8");
  assert.match(script, /const memberCapacity = 150/);
  assert.match(script, /tblMappingCategoryOverview/);
  assert.match(script, /tblMappingMemberWorkspace/);
  assert.match(script, /Assign Reporting Group/);
  assert.match(script, /Leave Unmapped/);
  assert.match(script, /Remove custom mapping/);
  assert.match(script, /Read-only membership inspection\. Return to Source Category to change mappings/);
  assert.doesNotMatch(script, /function resolveReportingGroupMembership/);
});

test("frozen 85-week Reporting Group totals reconcile to mapped attention", async () => {
  const expected = JSON.parse(await readFile(expectedPath, "utf8"));
  assert.equal(expected.reportingGroups.length, 9);
  assert.deepEqual(expected.totals, {
    products: 929, facts: 229190, salesNok: 479649885.1,
  });
  const totals = expected.reportingGroups.reduce((result, row) => ({
    products: result.products + row.products,
    facts: result.facts + row.facts,
    salesNok: result.salesNok + row.salesNok,
  }), { products: 0, facts: 0, salesNok: 0 });
  totals.salesNok = Math.round(totals.salesNok * 100) / 100;
  assert.deepEqual(totals, expected.totals);
  assert.equal(Math.max(...expected.reportingGroups.map(row => row.products)), 352);
  assert.equal(expected.nonAlcohol.products, 60);
  assert.equal(expected.mains.products, 352);
});

function fixtureView() {
  return buildReportingGroupMembershipView({
    rows: fixtureRows(), reportingGroups: fixtureGroups(), memberCapacity: 10,
  });
}

function fixtureGroups() {
  return [
    { reportingGroupId: "RPG-0001", reportingGroupName: "Add-ons", active: "Yes", sortOrder: 10 },
    { reportingGroupId: "RPG-0002", reportingGroupName: "Non-Alcohol", active: "Yes", sortOrder: 20 },
    { reportingGroupId: "RPG-0003", reportingGroupName: "Inactive fixture", active: "No", sortOrder: 30 },
  ];
}

function fixtureRows() {
  const base = {
    sourceMainCategory: "Main", sourceSubCategory: "Sub", salesAccount: "3000",
    historicalQuantity: 1, hierarchyAttention: "", hierarchyAlternatives: "",
  };
  return [
    { ...base, productId: "PRD-000001", item: "Inherited item", effectiveReportingGroupId: "RPG-0001",
      effectiveReportingGroupName: "Add-ons", resolutionType: "Inherited Main", mappingStatus: "Mapped",
      historicalFactCount: 2, historicalSalesNok: 200 },
    { ...base, productId: "PRD-000002", item: "Custom item", effectiveReportingGroupId: "RPG-0002",
      effectiveReportingGroupName: "Non-Alcohol", resolutionType: "Explicit Product", mappingStatus: "Mapped",
      historicalFactCount: 3, historicalSalesNok: 375 },
    { ...base, productId: "PRD-000003", item: "Unmapped", effectiveReportingGroupId: "",
      resolutionType: "Unmapped", mappingStatus: "Unmapped", historicalFactCount: 1, historicalSalesNok: 10 },
    { ...base, productId: "PRD-000004", item: "Excluded", effectiveReportingGroupId: "",
      resolutionType: "Explicit exclusion", mappingStatus: "Unmapped", historicalFactCount: 1, historicalSalesNok: 10 },
    { ...base, productId: "PRD-000005", item: "Pending", effectiveReportingGroupId: "",
      resolutionType: "Identity Pending", mappingStatus: "Identity Pending", historicalFactCount: 1, historicalSalesNok: 10 },
    { ...base, productId: "PRD-000006", item: "Conflict", effectiveReportingGroupId: "",
      resolutionType: "Conflict", mappingStatus: "Conflict", historicalFactCount: 1, historicalSalesNok: 10 },
  ];
}
