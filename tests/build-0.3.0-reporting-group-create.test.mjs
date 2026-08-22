import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  buildReportingGroupBusinessOverview,
  planReportingGroupCreation,
  validateReportingGroupAuthority,
} from "../src/reporting/reporting-group-administration.mjs";
import { activeReportingGroups } from "../src/imports/weekly-compact-cache.mjs";
import {
  buildWeeklyPerformanceLayout,
  planWeeklyPerformanceRpgSelection,
} from "../src/reporting/weekly-performance.mjs";

const officeScriptUrl = new URL("../office-scripts/Create_Reporting_Group.ts", import.meta.url);
let officeScript = "";
try { officeScript = await readFile(officeScriptUrl, "utf8"); } catch { /* added later in this slice */ }

test("creation allocates RPG-0010 and SortOrder 100 without user-authored IDs", () => {
  const plan = planReportingGroupCreation({
    reportingGroups: groups(9),
    reportingGroupName: "Kids Menu",
    description: "Meals intended for children.",
    notes: "Created by an administrator.",
  });
  assert.deepEqual(plan.reportingGroup, {
    reportingGroupId: "RPG-0010",
    reportingGroupName: "Kids Menu",
    domainId: "DOM-SALES",
    active: "Yes",
    sortOrder: 100,
    description: "Meals intended for children.",
    notes: "Created by an administrator.",
  });
  assert.equal(plan.mappingRulesCreated, 0);
  assert.equal(plan.performanceRefreshRequired, true);
});

test("ID allocation includes inactive historical rows and never reuses IDs", () => {
  const authority = validateReportingGroupAuthority([
    ...groups(9),
    { ...group(10), active: "No" },
    { ...group(12), active: "No", sortOrder: 120 },
  ]);
  assert.equal(authority.nextReportingGroupId, "RPG-0013");
  assert.equal(authority.nextSortOrder, 130);
});

test("blank and duplicate active business names are rejected before mutation", () => {
  assert.throws(() => planReportingGroupCreation({
    reportingGroups: groups(9), reportingGroupName: "   ",
  }), /name is required/i);
  assert.throws(() => planReportingGroupCreation({
    reportingGroups: groups(9), reportingGroupName: " add-ONS ",
  }), /already exists/i);
  assert.throws(() => planReportingGroupCreation({
    reportingGroups: groups(9).map(row => row.reportingGroupId === "RPG-0002"
      ? { ...row, reportingGroupName: "Kids Menu" }
      : row),
    reportingGroupName: "Kids   Menu",
  }), /already exists/i);
});

test("malformed authority, duplicate IDs, and duplicate SortOrder are rejected", () => {
  assert.throws(() => validateReportingGroupAuthority([
    ...groups(9), { ...group(10), reportingGroupId: "RPG-0009" },
  ]), /repeats ReportingGroupID/);
  assert.throws(() => validateReportingGroupAuthority([
    ...groups(9), { ...group(10), sortOrder: 90 },
  ]), /repeats SortOrder/);
  assert.throws(() => validateReportingGroupAuthority([
    ...groups(9), { ...group(10), reportingGroupId: "KIDS" },
  ]), /invalid stable ID/);
});

test("repeated execution fails clearly and cannot create a duplicate row", () => {
  const first = planReportingGroupCreation({
    reportingGroups: groups(9), reportingGroupName: "Kids Menu",
  }).reportingGroup;
  assert.throws(() => planReportingGroupCreation({
    reportingGroups: groups(9).concat(first), reportingGroupName: "Kids Menu",
  }), /already exists/i);
});

test("business overview exposes no IDs and gives the new group zero impact", () => {
  const created = planReportingGroupCreation({
    reportingGroups: groups(9), reportingGroupName: "Kids Menu",
  }).reportingGroup;
  const overview = buildReportingGroupBusinessOverview(groups(9).concat(created), [
    { reportingGroupId: "RPG-0001", products: 10, salesNok: 1250 },
  ]);
  assert.deepEqual(overview[0], {
    reportingGroupId: "RPG-0001", reportingGroupName: "Add-ons",
    status: "Active", products: 10, salesNok: 1250,
  });
  assert.deepEqual(overview[9], {
    reportingGroupId: "RPG-0010", reportingGroupName: "Kids Menu",
    status: "Active", products: 0, salesNok: 0,
  });
});

test("tenth group is dense-cache eligible but contributes zero until mapped", () => {
  const catalog = groups(10);
  const active = activeReportingGroups(catalog);
  const scopeRows = 1438;
  assert.equal(active.length, 10);
  assert.equal(scopeRows * active.length, 14380);
  assert.equal(active[9].reportingGroupId, "RPG-0010");
});

test("existing Performance selections survive and the created group defaults No", () => {
  const catalog = groups(10);
  const selection = planWeeklyPerformanceRpgSelection({
    reportingGroups: catalog,
    priorCatalogExists: true,
    priorRows: groups(9).map((row, index) => ({
      id: row.reportingGroupId,
      include: index === 4 ? "No" : "Yes",
    })),
  });
  assert.equal(selection.length, 10);
  assert.equal(selection[4].include, "No");
  assert.deepEqual(selection[9], { id: "RPG-0010", name: "Group 10", include: "No" });
  assert.equal(buildWeeklyPerformanceLayout({
    reportingGroups: catalog, restaurantCapacity: 16,
  }).groupCapacity, 10);
});

test("creation changes the active-group contract and therefore requires cache refresh", () => {
  const prior = activeReportingGroups(groups(9)).map(row => row.reportingGroupId);
  const next = activeReportingGroups(groups(10)).map(row => row.reportingGroupId);
  assert.notDeepEqual(next, prior);
  assert.deepEqual(next.slice(0, 9), prior);
  assert.equal(next[9], "RPG-0010");
});

test("Office Script is a bounded create entry point with no rename/deactivate/delete path", () => {
  assert.ok(officeScript.length > 0, "Create_Reporting_Group.ts must exist");
  assert.match(officeScript, /^function main\(workbook: ExcelScript\.Workbook\): string/m);
  assert.match(officeScript, /tblReportingGroups/);
  assert.match(officeScript, /RPG-/);
  assert.match(officeScript, /Performance refresh required/);
  assert.match(officeScript, /tblPerformanceRPGSelection/);
  assert.match(officeScript, /tblMappingReportingGroupCatalog/);
  assert.match(officeScript, /rows\[index\]\[h\.ReportingGroupID\]/);
  assert.match(officeScript, /priorIncludeById/);
  assert.match(officeScript, /priorIncludeById\[group\.id\] !== undefined[\s\S]*\? priorIncludeById\[group\.id\][\s\S]*: "No"/);
  assert.doesNotMatch(officeScript, /function\s+(rename|deactivate|delete)ReportingGroup/i);
  assert.doesNotMatch(officeScript, /tblMappingRules.*addRow|mappingRules.*addRow/i);
});

test("Office Script validates completely before its single authority append", () => {
  const append = officeScript.indexOf("groupsTable.addRow");
  assert.ok(append > 0);
  for (const marker of [
    "validateAuthority", "planCreation", "captureProtectedState", "validateDownstreamSurfaces",
  ]) {
    assert.ok(officeScript.indexOf(marker) >= 0 && officeScript.indexOf(marker) < append,
      `${marker} must run before the authority append`);
  }
  assert.match(officeScript, /deleteRowsAt\(priorGroupCount, 1\)/);
});

test("Office Script remains compatible with the Office Scripts TypeScript target", () => {
  assert.doesNotMatch(officeScript, /^(export|async|public|private|protected)\s+function\s+main/m);
  assert.doesNotMatch(officeScript, /new\s+Map\s*</);
  assert.doesNotMatch(officeScript, /new\s+Set\s*</);
  assert.doesNotMatch(officeScript, /\.(entries|keys|values)\s*\(\)/);
  assert.doesNotMatch(officeScript, /Array\.from\s*\(/);
  assert.doesNotMatch(officeScript, /\.\.\./);
  const loopBodies = Array.from(officeScript.matchAll(/for\s*\([^)]*\)\s*\{([\s\S]*?)\n  \}/g), match => match[1]);
  for (const body of loopBodies) {
    assert.doesNotMatch(body, /workbook\.get|sheet\.get|table\.get|\.getValues\(|\.getTexts\(/);
  }
  assertBalanced(officeScript);
});

function groups(count) {
  return Array.from({ length: count }, (_, index) => group(index + 1));
}

function group(number) {
  return {
    reportingGroupId: `RPG-${String(number).padStart(4, "0")}`,
    reportingGroupName: number === 1 ? "Add-ons" : `Group ${number}`,
    domainId: "DOM-SALES",
    active: "Yes",
    sortOrder: number * 10,
    description: "",
    notes: "",
  };
}

function assertBalanced(source) {
  const stack = [];
  const pairs = { ")": "(", "]": "[", "}": "{" };
  let quote = "";
  let escaped = false;
  for (const char of source) {
    if (quote) {
      if (escaped) escaped = false;
      else if (char === "\\") escaped = true;
      else if (char === quote) quote = "";
      continue;
    }
    if (char === "\"" || char === "'" || char === "`") { quote = char; continue; }
    if (char === "(" || char === "[" || char === "{") stack.push(char);
    else if (pairs[char]) assert.equal(stack.pop(), pairs[char]);
  }
  assert.equal(quote, "");
  assert.deepEqual(stack, []);
}
