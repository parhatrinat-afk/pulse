import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const scriptPath = new URL("../office-scripts/Build_0_3_0_Phase1.ts", import.meta.url);
const docsPath = new URL("../docs/BUILD_0_3_0_MAPPING_UX_WORKSPACE.md", import.meta.url);

async function source() {
  return readFile(scriptPath, "utf8");
}

test("final Mapping presentation consolidates both browse modes into one member workspace", async () => {
  const script = await source();
  assert.match(script, /applyFinalMappingPresentation\(workbook, memberCapacity, reportingGroupMemberCapacity\)/);
  assert.match(script, /function applyFinalMappingPresentation/);
  assert.match(script, /A\$\{memberHeaderRow\}:V\$\{memberBodyEndRow\}/);
  assert.match(script, /memberTable\.setName\("tblMappingMemberWorkspace"\)/);
  assert.match(script, /"tblMappingReportingGroupMembers"/);
  assert.match(script, /if \(table\) table\.delete\(\)/);
  assert.equal((script.match(/function applyFinalMappingPresentation/g) ?? []).length, 1);
});

test("source editing remains bounded to 150 rows while Reporting Group inspection supports 400", async () => {
  const script = await source();
  assert.match(script, /const memberCapacity = 150/);
  assert.match(script, /const reportingGroupMemberCapacity = 400/);
  assert.match(script, /rowIndex < sourceMemberCapacity/);
  assert.match(script, /applyLiteralValidation\(sheet\.getRange\("A23:A172"\), \["Yes", "No"\]/);
  assert.match(script, /PUL-0301-028/);
  assert.match(script, /PUL-0301-032/);
});

test("shared formulas switch between accepted hidden catalogs without another resolver", async () => {
  const script = await source();
  const body = script.match(/function applyFinalMappingPresentation[\s\S]*?\n}\n\nfunction writeMappingLists/)?.[0] ?? "";
  assert.ok(body);
  assert.match(body, /IF\(\$B\$10="Reporting Group"/);
  assert.match(body, /'_Mapping_Audit'!\$J\$2:\$J\$/);
  assert.match(body, /'_Mapping_Audit'!\$BU\$2:\$BU\$/);
  assert.match(body, /requiredTable\(workbook, "tblMappingMemberCatalog"\)/);
  assert.match(body, /requiredTable\(workbook, "tblMappingReportingGroupCatalog"\)/);
  assert.doesNotMatch(body, /resolveProduct|new Map|fuzzy/i);
});

test("selected context and active member list stay near the top", async () => {
  const script = await source();
  assert.match(script, /const memberSectionRow = 21/);
  assert.match(script, /const memberHeaderRow = 22/);
  assert.match(script, /getRange\("A19:K19"\)\.merge/);
  assert.match(script, /current category mapping:/);
  assert.match(script, /" inherited · "/);
  assert.match(script, /" custom"/);
  assert.match(script, /Members in /);
  assert.match(script, /Products in /);
  assert.match(script, /FIXED\(\$AY\$3,0,FALSE\)/);
  assert.match(script, /FIXED\(\$AX\$4,2,FALSE\)/);
  assert.match(script, /getRange\("E13"\)\.setFormula\('\=COUNTIF\(tblMappingMemberWorkspace\[Select\]/);
  assert.match(script, /getRange\("H13"\)\.setFormula\('\=SUMIFS\(tblMappingMemberWorkspace\[Facts\]/);
  assert.match(script, /getRange\("J13"\)\.setFormula\('\=SUMIFS\(tblMappingMemberWorkspace\[Sales NOK\]/);
});

test("category and Reporting Group overviews remain secondary and retain authoritative names", async () => {
  const script = await source();
  assert.match(script, /const categorySectionRow = 425/);
  assert.match(script, /Source Category overview — navigation and impact/);
  assert.match(script, /Reporting Group overview — current effective Product membership/);
  assert.match(script, /finalCategoryTable\.setName\("tblMappingCategoryOverview"\)/);
  assert.match(script, /finalReportingGroupTable\.setName\("tblMappingReportingGroupOverview"\)/);
  assert.match(script, /setValues\(categoryValues\)/);
  assert.match(script, /setValues\(reportingGroupOverviewValues\)/);
});

test("compact health cards preserve exact Product, Fact, Sales and freshness evidence", async () => {
  const script = await source();
  const body = script.match(/function restoreWeeklyMappingHealthBlock[\s\S]*?\n}\n\nfunction applyListValidation/)?.[0] ?? "";
  assert.ok(body);
  assert.match(body, /const states = \["Mapped", "Unmapped", "Identity Pending", "Conflict", "Inactive Target"\]/);
  assert.match(body, /#,##0 "products"/);
  assert.match(body, /#,##0 "facts"/);
  assert.match(body, /"NOK " #,##0\.00/);
  assert.match(body, /Performance classifications are up to date through/);
  assert.match(body, /Mapping health — Attention required/);
  assert.match(body, /Mapping health — Structural review required/);
  assert.doesNotMatch(body, /fingerprint|CacheVersion|WCV-|WCC-/i);
});

test("final Mapping corrections make deselection explicit and remove visible validation debug", async () => {
  const script = await source();
  const validationBody = script.match(/function wireMappingValidation[\s\S]*?\n}\n\nfunction restoreWeeklyMappingHealthBlock/)?.[0] ?? "";
  assert.ok(validationBody);
  assert.match(script, /Use Select = Yes or No/);
  assert.match(validationBody, /\["Yes", "No"\]/);
  assert.match(validationBody, /statusBand\.clear\(ExcelScript\.ClearApplyTo\.contents\)/);
  assert.doesNotMatch(validationBody, /getRange\("A18"\)\.setValue\(message\)/);
  assert.match(script, /setWidths\(sheet, \[68, 82, 112, 80, 82, 92, 96, 72, 90, 96, 86\]\)/);
  assert.match(script, /C\$\{memberBodyStartRow\}:H\$\{memberBodyEndRow\}.*setWrapText\(true\)/);
  assert.match(script, /C\$\{memberBodyStartRow\}:K\$\{memberBodyEndRow\}.*autofitRows\(\)/);
});

test("upper Mapping workspace uses compact geometry and intentional alignment", async () => {
  const script = await source();
  assert.match(script, /const centeredControlRanges = \["B10:C10", "E10:F10", "B12:C12", "E12:F12", "B13:C13"/);
  assert.match(script, /"B15:C15", "E15:F15", "G10:K10"/);
  assert.match(script, /const leftAlignedControlRanges = \["A10", "D10", "A11:K11", "A12", "D12"/);
  assert.match(script, /getRange\("10:10"\).*setRowHeight\(28\)/);
  assert.match(script, /getRange\("11:13"\).*setRowHeight\(26\)/);
  assert.match(script, /getRange\("14:14"\).*setRowHeight\(34\)/);
  assert.match(script, /getRange\("15:17"\).*setRowHeight\(26\)/);
  assert.match(script, /getRange\("4:5"\).*setRowHeight\(24\)/);
  assert.match(script, /getRange\("6:8"\).*setRowHeight\(20\)/);
  assert.match(script, /getRange\("9:9"\).*setRowHeight\(26\)/);
});

test("Mapping tables align identities as text and compact measures as values", async () => {
  const script = await source();
  assert.match(script, /const centeredMemberColumns = \["Select", "Level", "Reporting Group", "Mapping state", "Facts", "Sales NOK"\]/);
  assert.match(script, /const leftAlignedMemberColumns = \["Item", "Main Category", "Subcategory", "Sales Account", "Attention"\]/);
  assert.match(script, /const centeredCategoryColumns = \["Current Reporting Group", "Mapping state", "Subcategories", "Products"/);
  assert.match(script, /finalCategoryTable\.getColumn\("Main Category"\)/);
  assert.match(script, /finalCategoryTable\.getColumn\("Attention"\)/);
  assert.match(script, /finalReportingGroupTable\.getColumn\("Reporting Group"\)/);
  assert.match(script, /const centeredReportingGroupColumns = \["Products", "Facts", "Sales NOK"\]/);
  assert.match(script, /cardStarts\[stateIndex\].*cardEnds\[stateIndex\].*HorizontalAlignment\.center/s);
});

test("visible merged action controls use selection-aware same-sheet validation sources", async () => {
  const script = await source();
  const validationBody = script.match(/function wireMappingValidation[\s\S]*?\n}\n\nfunction restoreWeeklyMappingHealthBlock/)?.[0] ?? "";
  assert.ok(validationBody);
  assert.match(validationBody, /applyRangeValidation\(sheet\.getRange\("B15:C15"\), sheet\.getRange\("AM2:AM4"\)/);
  assert.match(validationBody, /applyRangeValidation\(sheet\.getRange\("E15:F15"\), sheet\.getRange\(`AN2:AN/);
  assert.doesNotMatch(validationBody, /applyLiteralValidation\(sheet\.getRange\("B15"\)/);
  assert.doesNotMatch(validationBody, /applyLiteralValidation\(sheet\.getRange\("E15"\)/);
  assert.match(script, /"Mapping action state", "Eligible action choices", "Eligible Reporting Group choices"/);
  assert.match(script, /IF\(\$AL\$8,"Assign Reporting Group",""\)/);
  assert.match(script, /IF\(AND\(\$AL\$8,\$B\$13="Selected members",\$AL\$5=\$AL\$2\),"Leave Unmapped",""\)/);
  assert.match(script, /IF\(AND\(\$AL\$8,IF\(\$B\$13="Selected members",\$AL\$6=\$AL\$2,\$AL\$7>0\)\),"Remove custom mapping",""\)/);
  assert.match(script, /IF\(AND\(\$AL\$8,\$B\$15="Assign Reporting Group"\),\$AS\$/);
  assert.match(script, /const activeGroupCount = Math\.max\(0, reportingGroupOverviewValues\.length - 1\)/);
});

test("action eligibility excludes review-only and structurally unsafe batches", async () => {
  const script = await source();
  assert.match(script, /IdentityState\],"Identity Pending"/);
  assert.match(script, /Mapping state\],"Conflict"/);
  assert.match(script, /Mapping state\],"Inactive Target"/);
  assert.match(script, /ScopeType\],"SourceMainCategory"/);
  assert.match(script, /ScopeType\],"SourceSubCategory"/);
  assert.match(script, /ParentSubNodeID/);
  assert.match(script, /LEN\(tblMappingMemberWorkspace\[CurrentExplicitRuleID\]\)>0/);
  assert.match(script, /LEN\(tblMappingMemberCatalog\[CurrentExplicitRuleID\]\)>0/);
  assert.match(script, /LEN\(tblMappingMemberWorkspace\[NodeID\]\)>0/);
});

test("browse, category and view controls lock while stable-ID selections exist", async () => {
  const script = await source();
  assert.match(script, /IF\(COUNTIF\(tblMappingMemberWorkspace\[Select\],"Yes"\)>0,\$B\$10,\$AW\$2:\$AW\$3\)/);
  assert.match(script, /IF\(COUNTIF\(tblMappingMemberWorkspace\[Select\],"Yes"\)>0,\$B\$12,SORT/);
  assert.match(script, /IF\(COUNTIF\(tblMappingMemberWorkspace\[Select\],"Yes"\)>0,\$E\$12/);
  assert.match(script, /applyRangeValidation\(sheet\.getRange\("B10"\), sheet\.getRange\("AV2:AV3"\)/);
});

test("visible workspace is business-first while IDs and lineage remain hidden", async () => {
  const script = await source();
  const headerBlock = script.match(/const memberHeaders = \["Select", "Level", "Item", "Main Category"[\s\S]*?\];/)?.[0] ?? "";
  assert.ok(headerBlock);
  const visibleHeaderText = headerBlock.split('"Historical Quantity"')[0];
  assert.doesNotMatch(visibleHeaderText, /ProductID|NodeID|RuleID|ReportingGroupID/);
  assert.match(headerBlock, /SourceSystemID.*NodeID.*ProductID.*CurrentExplicitRuleID/s);
  assert.match(script, /sheet\.getRange\("L:AY"\)\.setColumnHidden\(true\)/);
  assert.match(script, /auditSheet\.setVisibility\(ExcelScript\.SheetVisibility\.hidden\)/);
});

test("documentation records the one-workspace contract", async () => {
  const docs = await readFile(docsPath, "utf8");
  assert.match(docs, /one shared member workspace near the top/);
  assert.match(docs, /Source Category mode\s+uses its accepted 150-row selectable bound/);
  assert.match(docs, /Reporting Group mode uses the\s+same physical area with a 400-row read-only bound/);
  assert.match(docs, /secondary 93-category and nine-group overview tables/);
});
