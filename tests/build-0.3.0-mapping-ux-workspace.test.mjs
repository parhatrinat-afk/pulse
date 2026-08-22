import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import test from "node:test";

const scriptPath = new URL("../office-scripts/Build_0_3_0_Phase1.ts", import.meta.url);

async function source() {
  return readFile(scriptPath, "utf8");
}

test("Mapping workspace remains a facade over the accepted authorities", async () => {
  const script = await source();
  assert.match(script, /requiredTable\(workbook, "tblWeeklyMappingAttention"\)/);
  assert.match(script, /resolveProduct\(node, rules, groupById, asOf\)/);
  assert.match(script, /tblMappingRules/);
  assert.match(script, /tblReportingGroups/);
  assert.match(script, /resetOutputSheet\(workbook, "_Mapping_Audit"/);
  assert.match(script, /auditSheet\.setVisibility\(ExcelScript\.SheetVisibility\.hidden\)/);
  for (const table of ["tblMappingMainNodes", "tblMappingSubcategoryNodes", "tblMappingProducts",
    "tblMappingCategoryCatalog", "tblMappingMemberCatalog"]) {
    assert.match(script, new RegExp(`setName\\("${table}"\\)`));
  }
  assert.doesNotMatch(script, /Technical hierarchy audit — authoritative tables retained below/);
  assert.doesNotMatch(script, /resetOutputSheet\(workbook, "Mapping"[^;]*tblMappingRules/);
});

test("category overview is navigation and never an always-on bulk selector", async () => {
  const script = await source();
  assert.match(script, /Category overview — navigation and impact/);
  assert.match(script, /"Main Category", "Current Reporting Group", "Mapping state", "Subcategories", "Products"/);
  assert.match(script, /"Historical Facts", "Historical Sales NOK", "Attention"/);
  assert.match(script, /categoryTable\.setName\("tblMappingCategoryOverview"\)/);
  const categoryHeaders = script.match(/const categoryHeaders = \[[\s\S]*?\];/)?.[0] ?? "";
  assert.doesNotMatch(categoryHeaders, /Select|Include/);
});

test("Show category and view drive a bounded formula-backed member working set", async () => {
  const script = await source();
  assert.match(script, /setValue\("Show category"\)/);
  assert.match(script, /const views = \["All", "Unmapped", "Custom", "Identity Pending", "Excluded"\]/);
  assert.match(script, /const memberCapacity = 150/);
  assert.match(script, /getRange\("AP1"\)\.setValue\("Catalog row index"\)/);
  assert.match(script, /AGGREGATE\(15,6/);
  assert.match(script, /COUNT\(\$AP\$2:\$AP\$151\)/);
  assert.match(script, /INDEX\('_Mapping_Audit'!\$\$\{sourceColumn\}\$2:\$\$\{sourceColumn\}\$\$\{catalogEndRow\},\$AP\$\{rowIndex \+ 2\}\)/);
  assert.match(script, /setFormulas\(formulas\)/);
  assert.doesNotMatch(script, /FILTER\(tblMappingMemberCatalog/);
  assert.match(script, /PUL-0301-028/);
  assert.doesNotMatch(script, /"Main Category"\], \["Subcategory"\], \["Product"\]/);
});

test("regeneration clears stale merged cells and shown sales count products once", async () => {
  const script = await source();
  assert.match(script, /if\(used\)\{used\.unmerge\(\);used\.clear\(ExcelScript\.ClearApplyTo\.all\);\}/);
  assert.match(script, /SUMIFS\(\$I\$\$\{memberBodyStartRow\}:\$I\$\$\{memberBodyEndRow\},\$B\$\$\{memberBodyStartRow\}:\$B\$\$\{memberBodyEndRow\},"Product"\)/);
});

test("only Select is editable while stable IDs stay backstage", async () => {
  const script = await source();
  assert.match(script, /const memberHeaders = \["Select", "Level", "Item", "Subcategory", "Sales Account", "Reporting Group", "Mapping state"/);
  assert.match(script, /"SourceSystemID", "ScopeType", "NodeID", "ProductID"/);
  assert.match(script, /sheet\.getRange\("K:U"\)\.setColumnHidden\(true\)/);
  assert.match(script, /sheet\.getRange\("K:AU"\)\.setColumnHidden\(true\)/);
  assert.match(script, /applyLiteralValidation\(sheet\.getRange\("A23:A172"\), \["Yes", "No"\]/);
});

test("member selections lock category and view to prevent stable-ID drift", async () => {
  const script = await source();
  assert.match(script, /IF\(COUNTIF\(tblMappingMemberWorkspace\[Select\],"Yes"\)>0,\$B\$12,SORT\(tblMappingCategoryCatalog\[CategoryChoice\]\)\)/);
  assert.match(script, /IF\(COUNTIF\(tblMappingMemberWorkspace\[Select\],"Yes"\)>0,\$E\$12,\$AT\$2:\$AT\$6\)/);
  assert.match(script, /applyRangeValidation\(sheet\.getRange\("B12"\), sheet\.getRange\("AQ2:AQ200"\)/);
  assert.match(script, /applyRangeValidation\(sheet\.getRange\("E12"\), sheet\.getRange\("AR2:AR10"\)/);
});

test("bulk actions are name-first and preserve established rule semantics", async () => {
  const script = await source();
  for (const action of ["Assign Reporting Group", "Leave Unmapped", "Remove custom mapping"]) {
    assert.match(script, new RegExp(action));
  }
  assert.match(script, /targetMatches = groups\.filter\(group => group\.name === state\.targetGroupName/);
  assert.match(script, /ruleAction = intent === "Leave Unmapped" \? "Exclude" : "Map"/);
  assert.match(script, /normal hierarchy inheritance is restored/);
});

test("whole-category mapping is a separate intentional scope", async () => {
  const script = await source();
  assert.match(script, /\["Selected members", "Entire shown category"\]/);
  assert.match(script, /Clear member selections before choosing Entire shown category/);
  assert.match(script, /scopeType: "SourceMainCategory"/);
  assert.match(script, /Use Entire shown category only for an intentional category-wide rule/);
  assert.match(script, /Leave Unmapped is supported only for selected Products/);
});

test("complete batch safety validation precedes all writes", async () => {
  const script = await source();
  const body = script.match(/function applyPendingAction[\s\S]*?\n}\n\nfunction readMappingWorkspaceState/)?.[0] ?? "";
  assert.ok(body);
  const writeIndex = body.indexOf("table.addRows(-1, newRows)");
  assert.ok(writeIndex > 0);
  for (const guard of ["activeIndexes.length > 1", "Identity Pending", "mappingState === \"Conflict\"",
    "selectedSubcategories[target.parentSubNodeId]", "Remove custom mapping"]) {
    assert.ok(body.indexOf(guard) >= 0 && body.indexOf(guard) < writeIndex, `${guard} must precede mutation`);
  }
  assert.match(body, /table\.deleteRowsAt\(table\.getRowCount\(\) - appendedCount, appendedCount\)/);
  assert.match(body, /setValues\(originalStatuses\)/);
});

test("ancestor plus descendant selection is rejected atomically", async () => {
  const script = await source();
  assert.match(script, /target\.scopeType === "SourceSubCategory"/);
  assert.match(script, /target\.scopeType === "Product" && target\.parentSubNodeId/);
  assert.match(script, /Do not select both .* and its Subcategory ancestor in one batch/);
});

test("Identity Pending remains review-only with weekly hierarchy evidence", async () => {
  const script = await source();
  assert.match(script, /is Identity Pending and remains review-only/);
  assert.match(script, /Identity review required/);
  assert.match(script, /Hierarchy Alternatives/);
  assert.match(script, /hierarchyAlternatives/);
});

test("mapping mutations mark historical Performance classifications stale", async () => {
  const script = await source();
  assert.match(script, /const mappingActionRequested = workspaceState\.intent !== ""/);
  assert.match(script, /restoreWeeklyMappingHealthBlock\(workbook, mappingActionRequested\)/);
  assert.match(script, /Performance refresh required/);
});

test("workspace keeps restrained Pulse styling and technical audit access", async () => {
  const script = await source();
  assert.match(script, /#17365D/);
  assert.match(script, /#4F8CFF/);
  assert.match(script, /#FFF4CE/);
  assert.match(script, /setShowGridlines\(false\)/);
  assert.match(script, /freezeRows\(2\)/);
  assert.match(script, /setWidths\(sheet,\[105,130,175,115,145,125,105,105,105,170/);
  assert.match(script, /setNumberFormat\("#,##0"\)/);
  assert.match(script, /setNumberFormat\("#,##0\.00"\)/);
  assert.match(script, /H21:H\$\{categoryEndRow\}.*setWrapText\(true\)/s);
  assert.match(script, /J\$\{memberBodyStartRow\}:J\$\{memberBodyEndRow\}.*setWrapText\(true\)/s);
});

test("backstage hierarchy audit preserves live formula lineage", async () => {
  const script = await source();
  assert.match(script, /tblSalesFacts\[SalesAmount\].*tblSalesFacts\[ProductID\].*\$BH\$\{formulaRow\}/);
  assert.match(script, /tblSalesFacts\[Quantity\].*tblSalesFacts\[ProductID\].*\$BH\$\{formulaRow\}/);
  assert.match(script, /tblMappingProducts\[Sales NOK \(Active\)\].*\$AH\$\{formulaRow\}/);
  assert.match(script, /tblMappingProducts\[Quantity \(Active\)\].*\$AH\$\{formulaRow\}/);
  assert.match(script, /tblMappingProducts\[Sales NOK \(Active\)\].*\$AU\$\{formulaRow\}.*\$AV\$\{formulaRow\}/);
  assert.match(script, /tblMappingProducts\[Quantity \(Active\)\].*\$AU\$\{formulaRow\}.*\$AV\$\{formulaRow\}/);
  assert.match(script, /setFormulas\(productImpactFormulas\)/);
  assert.match(script, /setFormulas\(mainImpactFormulas\)/);
  assert.match(script, /setFormulas\(subImpactFormulas\)/);
});
