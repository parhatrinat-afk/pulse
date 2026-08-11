/**
 * Pulse Build 0.3.0 Phase 1 — Reporting Groups + hierarchical mapping foundation.
 *
 * Run against excel/Pulse_Build_0_2_0_QA.xlsx.xlsx.
 * This script is additive: it does not edit raw imports, staging, tblSalesFacts,
 * Performance, Reports, or the legacy 0.2.0 category metric path.
 *
 * Rerun behavior:
 * - Reporting Group edits and Mapping Rules are preserved.
 * - Effective Mapping, Mapping browse tables, and Mapping QA are refreshed.
 * - If Mapping!B5 is "Apply", the action form is validated and appended as a
 *   new explicit rule before outputs are refreshed.
 */
function main(workbook: ExcelScript.Workbook): string {
  const sourceClassifications = requiredTable(workbook, "tblSourceClassifications");
  const products = requiredTable(workbook, "tblProducts");
  const facts = requiredTable(workbook, "tblSalesFacts");
  const reportingCategories = requiredSheet(workbook, "Reporting Categories");
  const environment = requiredTable(workbook, "tblEnvironment");
  const buildLog = requiredTable(workbook, "tblBuildLog");

  const factSnapshot = snapshotFacts(facts);
  const groups = ensureReportingGroups(workbook);
  const mappingRules = ensureMappingRules(workbook);
  markLegacyCategories(reportingCategories);

  const mappingSheet = workbook.getWorksheet("Mapping");
  const actionMessage = mappingSheet
    ? applyPendingAction(mappingSheet, mappingRules, groups, sourceClassifications, products)
    : "No pending mapping action.";

  const hierarchy = buildHierarchy(sourceClassifications, products);
  const rules = readRules(mappingRules);
  const groupById = new Map<string, ReportingGroup>();
  for (const group of groups) groupById.set(group.id, group);
  const conflicts = findConflicts(rules);
  const salesContext = aggregateSalesContext(facts);
  const resolved = hierarchy.products.map(product =>
    resolveProduct(product, rules, groupById, excelToday())
  );

  writeEffectiveMapping(workbook, resolved);
  writeMappingSurface(workbook, hierarchy, resolved, salesContext, groups, actionMessage);
  writeMappingLists(workbook, groups);
  wireMappingValidation(workbook);
  writeMappingQA(workbook, groups, rules, conflicts, resolved, factSnapshot, salesContext);

  const factAfter = snapshotFacts(facts);
  if (!sameFactSnapshot(factSnapshot, factAfter)) {
    throw new Error("PUL-0301-001: Fact snapshot changed during mapping migration.");
  }

  updateEnvironment(environment, "BuildID", "0.3.0-Phase1", "Mapping foundation only; Performance remains on legacy categories.");
  updateEnvironment(environment, "BuildVersion", "0.3.0-Phase1", "Reporting Groups and hierarchy-aware mapping resolver.");
  appendBuildLog(buildLog, [
    nextId(buildLog, "LOG", 6), excelNow(), "0.3.0-Phase1",
    "Refresh hierarchical mapping foundation", "Success", "Mapping + Reporting Groups",
    `${resolved.length} products resolved; ${conflicts.length} overlapping rule conflict(s); facts unchanged.`
  ]);
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  return `Pulse 0.3.0 Phase 1 refreshed. ${groups.length} Reporting Groups, ${rules.length} rule(s), ${resolved.length} product(s). ${actionMessage}`;
}

type ReportingGroup = { id: string; name: string; active: string; sortOrder: number };
type MappingRule = {
  id: string; sourceSystemId: string; scopeType: string; nodeId: string;
  nodeDisplay: string; hierarchyLevel: number; targetGroupId: string;
  effectiveFrom: number; effectiveTo: number; status: string; createdAt: number; notes: string;
};
type ProductNode = {
  productId: string; sourceSystemId: string; productName: string; salesAccount: string;
  sourceClassificationId: string; mainCategory: string; subCategory: string;
  mainNodeId: string; subNodeId: string;
};
type Resolution = ProductNode & {
  mainRuleIds: string; mainTargetIds: string; subRuleIds: string; subTargetIds: string;
  productRuleIds: string; productTargetIds: string; effectiveGroupId: string;
  effectiveGroupName: string; resolutionSource: string; resolutionState: string;
  resolutionStatus: string; winningRuleId: string;
};
type Conflict = { leftRuleId: string; rightRuleId: string; scopeType: string; nodeId: string };
type FactSnapshot = { rows: number; quantity: number; sales: number; activeQuantity: number; activeSales: number; firstId: string; lastId: string };

function ensureReportingGroups(workbook: ExcelScript.Workbook): ReportingGroup[] {
  const sheet = workbook.getWorksheet("Reporting Groups") ?? workbook.addWorksheet("Reporting Groups");
  let table = workbook.getTable("tblReportingGroups");
  if (!table) {
    sheet.getRange("A1:G20").clear(ExcelScript.ClearApplyTo.all);
    writeTitle(sheet, "Reporting Groups", "Business-owned semantic configuration. IDs are stable; names and status are editable.", "G");
    const headers = [["ReportingGroupID", "ReportingGroupName", "DomainID", "Active", "SortOrder", "Description", "Notes"]];
    sheet.getRange("A4:G4").setValues(headers);
    const seeds: (string | number)[][] = [
      ["RPG-0001", "Add-ons", "DOM-SALES", "Yes", 10, "Optional extras and additions.", "Seed configuration; human editable."],
      ["RPG-0002", "Non-Alcohol", "DOM-SALES", "Yes", 20, "Non-alcoholic beverages.", "Seed configuration; human editable."],
      ["RPG-0003", "Spirits/Cocktails", "DOM-SALES", "Yes", 30, "Spirits, cocktails, and alcoholic mixed drinks.", "Seed configuration; human editable."],
      ["RPG-0004", "Coffee & Tea", "DOM-SALES", "Yes", 40, "Coffee, tea, and related hot drinks.", "Seed configuration; human editable."],
      ["RPG-0005", "Beer & Cider", "DOM-SALES", "Yes", 50, "Beer and cider.", "Seed configuration; human editable."],
      ["RPG-0006", "Desserts", "DOM-SALES", "Yes", 60, "Desserts and sweet finishes.", "Seed configuration; human editable."],
      ["RPG-0007", "Wine & Sake", "DOM-SALES", "Yes", 70, "Wine, sparkling wine, and sake.", "Seed configuration; human editable."],
      ["RPG-0008", "Starters", "DOM-SALES", "Yes", 80, "Starters and small opening dishes.", "Seed configuration; human editable."],
      ["RPG-0009", "Mains", "DOM-SALES", "Yes", 90, "Main dishes and meal centrepieces.", "Seed configuration; human editable."]
    ];
    sheet.getRange(`A5:G${seeds.length + 4}`).setValues(seeds);
    table = sheet.addTable(`A4:G${seeds.length + 4}`, true);
    table.setName("tblReportingGroups");
    table.setPredefinedTableStyle("TableStyleMedium2");
    sheet.getRange("D5:D200").getDataValidation().setRule({ list: { inCellDropDown: true, source: "Yes,No" } });
    sheet.getFreezePanes().freezeRows(4);
    setWidths(sheet, [110, 180, 100, 70, 75, 260, 250]);
  }
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  const groups: ReportingGroup[] = [];
  const seen = new Set<string>();
  for (const row of rows) {
    const id = text(row[0]);
    if (!id) continue;
    if (seen.has(id)) throw new Error(`PUL-0301-002: Duplicate ReportingGroupID ${id}.`);
    seen.add(id);
    groups.push({ id, name: text(row[1]), active: text(row[3]), sortOrder: numberValue(row[4]) });
  }
  return groups;
}

function ensureMappingRules(workbook: ExcelScript.Workbook): ExcelScript.Table {
  const sheet = workbook.getWorksheet("Mapping Rules") ?? workbook.addWorksheet("Mapping Rules");
  let table = workbook.getTable("tblMappingRules");
  if (!table) {
    sheet.getRange("A1:L20").clear(ExcelScript.ClearApplyTo.all);
    writeTitle(sheet, "Mapping Rules", "Authoritative explicit rules. Use the Mapping browse/action surface; descendants inherit computed state.", "L");
    sheet.getRange("A4:L4").setValues([[
      "MappingRuleID", "SourceSystemID", "ScopeType", "NodeID", "NodeDisplay", "HierarchyLevel",
      "TargetReportingGroupID", "EffectiveFrom", "EffectiveTo", "Status", "CreatedAt", "Notes"
    ]]);
    table = sheet.addTable("A4:L4", true);
    table.setName("tblMappingRules");
    table.setPredefinedTableStyle("TableStyleMedium2");
    table.getRangeBetweenHeaderAndTotal().clear(ExcelScript.ClearApplyTo.contents);
    sheet.getRange("H5:I1000").setNumberFormatLocal("dd.mm.yyyy");
    sheet.getRange("K5:K1000").setNumberFormatLocal("dd.mm.yyyy hh:mm");
    sheet.getFreezePanes().freezeRows(4);
    setWidths(sheet, [105, 110, 130, 210, 210, 85, 145, 100, 100, 85, 120, 240]);
  }
  return table;
}

function markLegacyCategories(sheet: ExcelScript.Worksheet): void {
  sheet.getRange("A2").setValue("Legacy/source-default classification retained for Build 0.2.0 compatibility. Reporting Groups are the new business semantic layer.");
  sheet.getRange("A2:G2").getFormat().getFill().setColor("#FFF4CE");
  sheet.getRange("A2:G2").getFormat().setWrapText(true);
  sheet.getRange("2:2").getFormat().setRowHeight(34);
}

function applyPendingAction(
  sheet: ExcelScript.Worksheet, table: ExcelScript.Table, groups: ReportingGroup[],
  classifications: ExcelScript.Table, products: ExcelScript.Table
): string {
  const action = text(sheet.getRange("B5").getValue());
  if (!action) return "No pending mapping action.";
  if (action === "Deactivate") {
    const ruleId = text(sheet.getRange("B8").getValue());
    const h = headerMap(table);
    const body = table.getRangeBetweenHeaderAndTotal();
    const rows = body.getValues();
    for (let i = 0; i < rows.length; i++) {
      if (text(rows[i][h.MappingRuleID]) !== ruleId) continue;
      body.getCell(i, h.Status).setValue("Inactive");
      sheet.getRange("B5:B12").clear(ExcelScript.ClearApplyTo.contents);
      return `Deactivated explicit rule ${ruleId}; nearest valid ancestor now resolves descendants.`;
    }
    throw new Error(`PUL-0301-011: MappingRuleID not found for deactivation: ${ruleId}.`);
  }
  if (action !== "Apply") throw new Error(`PUL-0301-012: Unsupported Mapping action ${action}.`);
  const sourceSystemId = text(sheet.getRange("B6").getValue());
  const scopeType = text(sheet.getRange("B7").getValue());
  const nodeId = text(sheet.getRange("B8").getValue());
  const targetGroupId = text(sheet.getRange("B9").getValue());
  const effectiveFrom = numberValue(sheet.getRange("B10").getValue()) || excelToday();
  const effectiveTo = numberValue(sheet.getRange("B11").getValue());
  const notes = text(sheet.getRange("B12").getValue());
  if (!sourceSystemId || !nodeId || !targetGroupId) throw new Error("PUL-0301-003: Mapping action is incomplete.");
  const levelByScope: { [key: string]: number } = { SourceMainCategory: 1, SourceSubCategory: 2, Product: 3 };
  if (!levelByScope[scopeType]) throw new Error(`PUL-0301-004: Unsupported current scope ${scopeType}.`);
  if (!groups.some(group => group.id === targetGroupId && group.active === "Yes")) {
    throw new Error(`PUL-0301-005: Mapping actions may target only an active Reporting Group: ${targetGroupId}.`);
  }
  if (effectiveTo && effectiveTo < effectiveFrom) throw new Error("PUL-0301-006: EffectiveTo precedes EffectiveFrom.");
  const nodeDisplay = resolveNodeDisplay(sourceSystemId, scopeType, nodeId, classifications, products);
  if (!nodeDisplay) throw new Error(`PUL-0301-007: NodeID does not exist for ${scopeType}: ${nodeId}.`);
  table.addRow(-1, [
    nextId(table, "MAP", 6), sourceSystemId, scopeType, nodeId, nodeDisplay, levelByScope[scopeType],
    targetGroupId, effectiveFrom, effectiveTo || "", "Active", excelNow(), notes
  ]);
  sheet.getRange("B5:B12").clear(ExcelScript.ClearApplyTo.contents);
  return `Applied explicit ${scopeType} rule for ${nodeDisplay}.`;
}

function resolveNodeDisplay(sourceSystemId: string, scopeType: string, nodeId: string, classifications: ExcelScript.Table, products: ExcelScript.Table): string {
  if (scopeType === "Product") {
    const rows = products.getRangeBetweenHeaderAndTotal().getValues();
    const h = headerMap(products);
    for (const row of rows) if (text(row[h.ProductID]) === nodeId && text(row[h.SourceSystemID]) === sourceSystemId) return text(row[h.SourceProductName]);
  } else {
    const rows = classifications.getRangeBetweenHeaderAndTotal().getValues();
    const h = headerMap(classifications);
    for (const row of rows) {
      if (text(row[h.SourceSystemID]) !== sourceSystemId) continue;
      if (scopeType === "SourceSubCategory" && text(row[h.SourceClassificationID]) === nodeId) return text(row[h.SourceSubCategory]);
      if (scopeType === "SourceMainCategory") {
        const id = mainNodeId(text(row[h.SourceSystemID]), text(row[h.SourceMainCategory]));
        if (id === nodeId) return text(row[h.SourceMainCategory]) || "(blank main category)";
      }
    }
  }
  return "";
}

function buildHierarchy(classifications: ExcelScript.Table, products: ExcelScript.Table): { products: ProductNode[] } {
  const cRows = classifications.getRangeBetweenHeaderAndTotal().getValues();
  const ch = headerMap(classifications);
  const classificationById = new Map<string, { sourceSystemId: string; main: string; sub: string }>();
  for (const row of cRows) {
    classificationById.set(text(row[ch.SourceClassificationID]), {
      sourceSystemId: text(row[ch.SourceSystemID]), main: text(row[ch.SourceMainCategory]), sub: text(row[ch.SourceSubCategory])
    });
  }
  const pRows = products.getRangeBetweenHeaderAndTotal().getValues();
  const ph = headerMap(products);
  const result: ProductNode[] = [];
  for (const row of pRows) {
    const productId = text(row[ph.ProductID]);
    if (!productId) continue;
    const classificationId = text(row[ph.SourceClassificationID]);
    const classification = classificationById.get(classificationId);
    if (!classification) throw new Error(`PUL-0301-008: Product ${productId} references missing classification ${classificationId}.`);
    result.push({
      productId, sourceSystemId: text(row[ph.SourceSystemID]), productName: text(row[ph.SourceProductName]),
      salesAccount: text(row[ph.SalesAccount]), sourceClassificationId: classificationId,
      mainCategory: classification.main, subCategory: classification.sub,
      mainNodeId: mainNodeId(classification.sourceSystemId, classification.main), subNodeId: classificationId
    });
  }
  return { products: result };
}

function readRules(table: ExcelScript.Table): MappingRule[] {
  const h = headerMap(table);
  return table.getRangeBetweenHeaderAndTotal().getValues().filter(row => text(row[h.MappingRuleID])).map(row => ({
    id: text(row[h.MappingRuleID]), sourceSystemId: text(row[h.SourceSystemID]), scopeType: text(row[h.ScopeType]),
    nodeId: text(row[h.NodeID]), nodeDisplay: text(row[h.NodeDisplay]), hierarchyLevel: numberValue(row[h.HierarchyLevel]),
    targetGroupId: text(row[h.TargetReportingGroupID]), effectiveFrom: numberValue(row[h.EffectiveFrom]),
    effectiveTo: numberValue(row[h.EffectiveTo]), status: text(row[h.Status]), createdAt: numberValue(row[h.CreatedAt]), notes: text(row[h.Notes])
  }));
}

function findConflicts(rules: MappingRule[]): Conflict[] {
  const conflicts: Conflict[] = [];
  for (let i = 0; i < rules.length; i++) {
    const a = rules[i];
    if (a.status !== "Active") continue;
    for (let j = i + 1; j < rules.length; j++) {
      const b = rules[j];
      if (b.status !== "Active" || a.sourceSystemId !== b.sourceSystemId || a.scopeType !== b.scopeType || a.nodeId !== b.nodeId) continue;
      const aStart = a.effectiveFrom || -1e15; const bStart = b.effectiveFrom || -1e15;
      const aEnd = a.effectiveTo || 1e15; const bEnd = b.effectiveTo || 1e15;
      if (aStart <= bEnd && bStart <= aEnd) conflicts.push({ leftRuleId: a.id, rightRuleId: b.id, scopeType: a.scopeType, nodeId: a.nodeId });
    }
  }
  return conflicts;
}

function resolveProduct(
  product: ProductNode, rules: MappingRule[], groups: Map<string, ReportingGroup>, asOf: number
): Resolution {
  const scopes = [
    { type: "SourceMainCategory", nodeId: product.mainNodeId },
    { type: "SourceSubCategory", nodeId: product.subNodeId },
    { type: "Product", nodeId: product.productId }
  ];
  const applicable: { [key: string]: MappingRule[] } = {};
  for (const scope of scopes) {
    applicable[scope.type] = rules.filter(rule => rule.status === "Active" && rule.sourceSystemId === product.sourceSystemId &&
      rule.scopeType === scope.type && rule.nodeId === scope.nodeId && (!rule.effectiveFrom || rule.effectiveFrom <= asOf) &&
      (!rule.effectiveTo || rule.effectiveTo >= asOf));
  }
  const base = {
    ...product,
    mainRuleIds: ids(applicable.SourceMainCategory), mainTargetIds: targets(applicable.SourceMainCategory),
    subRuleIds: ids(applicable.SourceSubCategory), subTargetIds: targets(applicable.SourceSubCategory),
    productRuleIds: ids(applicable.Product), productTargetIds: targets(applicable.Product)
  };
  for (const scopeType of ["Product", "SourceSubCategory", "SourceMainCategory"]) {
    const candidates = applicable[scopeType];
    if (candidates.length > 1) {
      return { ...base, effectiveGroupId: "", effectiveGroupName: "", resolutionSource: scopeType,
        resolutionState: "Explicit conflict", resolutionStatus: "Conflict", winningRuleId: ids(candidates) };
    }
    if (candidates.length === 1) {
      const rule = candidates[0]; const group = groups.get(rule.targetGroupId);
      return { ...base, effectiveGroupId: rule.targetGroupId, effectiveGroupName: group?.name ?? "",
        resolutionSource: scopeType, resolutionState: scopeType === "Product" ? "Explicit" : "Inherited",
        resolutionStatus: !group || group.active !== "Yes" ? "Inactive Target" : "Mapped", winningRuleId: rule.id };
    }
  }
  return { ...base, effectiveGroupId: "", effectiveGroupName: "", resolutionSource: "Unmapped",
    resolutionState: "Unmapped", resolutionStatus: "Unmapped", winningRuleId: "" };
}

function writeEffectiveMapping(workbook: ExcelScript.Workbook, rows: Resolution[]): void {
  const sheet = resetOutputSheet(workbook, "Effective Mapping", "tblEffectiveMapping");
  writeTitle(sheet, "Effective Mapping", "Read-only product-level audit. Most-specific explicit rule wins; otherwise nearest mapped ancestor is inherited.", "V");
  const headers = ["ProductID", "SourceSystemID", "SourceProductName", "SalesAccount", "SourceMainCategory", "SourceSubCategory",
    "SourceClassificationID", "MainCategoryRuleID", "MainCategoryTargetID", "SubcategoryRuleID", "SubcategoryTargetID",
    "ProductRuleID", "ProductTargetID", "EffectiveReportingGroupID", "EffectiveReportingGroupName", "ResolutionSource",
    "ResolutionState", "ResolutionStatus", "WinningRuleID", "MainNodeID", "SubcategoryNodeID", "AsOfDate"];
  sheet.getRange("A4:V4").setValues([headers]);
  const values = rows.map(row => [row.productId, row.sourceSystemId, row.productName, row.salesAccount, row.mainCategory, row.subCategory,
    row.sourceClassificationId, row.mainRuleIds, row.mainTargetIds, row.subRuleIds, row.subTargetIds, row.productRuleIds,
    row.productTargetIds, row.effectiveGroupId, row.effectiveGroupName, row.resolutionSource, row.resolutionState, row.resolutionStatus,
    row.winningRuleId, row.mainNodeId, row.subNodeId, excelToday()]);
  if (values.length) sheet.getRangeByIndexes(4, 0, values.length, headers.length).setValues(values);
  const table = sheet.addTable(`A4:V${Math.max(4, values.length + 4)}`, true); table.setName("tblEffectiveMapping"); table.setPredefinedTableStyle("TableStyleMedium2");
  sheet.getRange(`V5:V${Math.max(5, values.length + 4)}`).setNumberFormatLocal("dd.mm.yyyy");
  sheet.getFreezePanes().freezeRows(4); setWidths(sheet, [105, 105, 210, 170, 160, 180, 115, 120, 120, 120, 120, 120, 120, 135, 170, 130, 100, 105, 110, 220, 115, 90]);
}

function writeMappingSurface(
  workbook: ExcelScript.Workbook, hierarchy: { products: ProductNode[] }, resolved: Resolution[],
  context: Map<string, { sales: number; quantity: number }>, groups: ReportingGroup[], message: string
): void {
  const sheet = resetOutputSheet(workbook, "Mapping", "tblMappingMainNodes", "tblMappingSubcategoryNodes", "tblMappingProducts");
  writeTitle(sheet, "Mapping", "Browse source branch → inspect descendants and impact → map broad node when safe → inherit → override exceptions.", "N");
  sheet.getRange("A4:C4").setValues([["Mapping action", "Value", "How to use"]]); styleHeader(sheet.getRange("A4:C4"));
  sheet.getRange("A5:C12").setValues([
    ["Action", "", "Apply a node mapping, or Deactivate using its RuleID in the next NodeID field."], ["SourceSystemID", "", "Copy from the filtered browse row; not needed for Deactivate."],
    ["ScopeType", "", "Current hierarchy: SourceMainCategory / SourceSubCategory / Product."],
    ["NodeID / RuleID", "", "Copy NodeID for Apply, or Winning RuleID for Deactivate."], ["TargetReportingGroupID", "", "Active Reporting Groups only."],
    ["EffectiveFrom", excelToday(), "Blank defaults to today."], ["EffectiveTo", "", "Blank is open-ended."], ["Notes", "", "Human-owned context / audit note."]
  ]);
  sheet.getRange("A5:A12").getFormat().getFill().setColor("#EEF1F5"); sheet.getRange("A5:A12").getFormat().getFont().setBold(true);
  sheet.getRange("B5:B12").getFormat().getFill().setColor("#FFFFFF"); sheet.getRange("B5:B12").getFormat().getFont().setBold(true);
  sheet.getRange("B10:B11").setNumberFormatLocal("dd.mm.yyyy");
  sheet.getRange("E4:N4").setValues([["Workflow status", "", "", "", "", "", "", "", "", ""]]); styleHeader(sheet.getRange("E4:N4"));
  sheet.getRange("E5:N7").setValues([
    [message, "", "", "", "", "", "", "", "", ""],
    ["Filter a hierarchy table below. Broad rows show descendant counts and all Active Finalized sales/quantity breadth before assignment.", "", "", "", "", "", "", "", "", ""],
    ["Explicit lower-level rules remain separate rows in Mapping Rules and are never overwritten by parent remapping.", "", "", "", "", "", "", "", "", ""]
  ]); sheet.getRange("E5:N7").getFormat().setWrapText(true);

  const resultByProduct = new Map<string, Resolution>(); for (const row of resolved) resultByProduct.set(row.productId, row);
  const main = new Map<string, { source: string; name: string; products: Set<string>; subs: Set<string>; ruleIds: Set<string>; targetIds: Set<string>; inherited: number; sales: number; qty: number; exceptions: number }>();
  const subs = new Map<string, { source: string; main: string; name: string; products: Set<string>; ruleIds: Set<string>; targetIds: Set<string>; inherited: number; sales: number; qty: number; exceptions: number }>();
  for (const product of hierarchy.products) {
    const amount = context.get(product.productId) ?? { sales: 0, quantity: 0 }; const resolution = resultByProduct.get(product.productId)!;
    if (!main.has(product.mainNodeId)) main.set(product.mainNodeId, { source: product.sourceSystemId, name: product.mainCategory, products: new Set(), subs: new Set(), ruleIds: new Set(), targetIds: new Set(), inherited: 0, sales: 0, qty: 0, exceptions: 0 });
    const m = main.get(product.mainNodeId)!; m.products.add(product.productId); m.subs.add(product.subNodeId); addDelimited(m.ruleIds, resolution.mainRuleIds); addDelimited(m.targetIds, resolution.mainTargetIds); m.sales += amount.sales; m.qty += amount.quantity; if (resolution.resolutionSource === "SourceMainCategory" && resolution.resolutionStatus !== "Conflict") m.inherited++; if (resolution.resolutionSource === "SourceSubCategory" || resolution.resolutionSource === "Product") m.exceptions++;
    if (!subs.has(product.subNodeId)) subs.set(product.subNodeId, { source: product.sourceSystemId, main: product.mainCategory, name: product.subCategory, products: new Set(), ruleIds: new Set(), targetIds: new Set(), inherited: 0, sales: 0, qty: 0, exceptions: 0 });
    const s = subs.get(product.subNodeId)!; s.products.add(product.productId); addDelimited(s.ruleIds, resolution.subRuleIds); addDelimited(s.targetIds, resolution.subTargetIds); s.sales += amount.sales; s.qty += amount.quantity; if (resolution.resolutionSource === "SourceSubCategory" && resolution.resolutionStatus !== "Conflict") s.inherited++; if (resolution.resolutionSource === "Product") s.exceptions++;
  }
  let row = 15;
  sheet.getRange(`A${row}:L${row}`).setValues([["SourceSystemID", "ScopeType", "NodeID", "SourceMainCategory", "Explicit Rule(s)", "Target Group ID(s)", "Subcategories", "Products", "Inheriting products", "Sales NOK (Active)", "Quantity (Active)", "Descendant exceptions"]]);
  const mainValues: (string|number)[][] = [];
  main.forEach((v,id) => mainValues.push([v.source,"SourceMainCategory",id,v.name,joinSet(v.ruleIds),joinSet(v.targetIds),v.subs.size,v.products.size,v.inherited,v.sales,v.qty,v.exceptions]));
  mainValues.sort((a,b)=>text(a[3]).localeCompare(text(b[3])));
  if (mainValues.length) sheet.getRangeByIndexes(row,0,mainValues.length,12).setValues(mainValues); const mt=sheet.addTable(`A${row}:L${Math.max(row,row+mainValues.length)}`,true); mt.setName("tblMappingMainNodes"); mt.setPredefinedTableStyle("TableStyleMedium2");
  if (mainValues.length) { sheet.getRange(`G${row + 1}:I${row + mainValues.length}`).setNumberFormat("#,##0"); sheet.getRange(`J${row + 1}:K${row + mainValues.length}`).setNumberFormat("#,##0.00"); sheet.getRange(`L${row + 1}:L${row + mainValues.length}`).setNumberFormat("#,##0"); }
  row += Math.max(2, mainValues.length + 3);
  sheet.getRange(`A${row}:M${row}`).setValues([["SourceSystemID", "ScopeType", "NodeID", "SourceMainCategory", "SourceSubCategory", "Explicit Rule(s)", "Target Group ID(s)", "Products", "Inheriting products", "Sales NOK (Active)", "Quantity (Active)", "Product exceptions", "Browse cue"]]);
  const subValues: (string|number)[][] = [];
  subs.forEach((v,id) => subValues.push([v.source,"SourceSubCategory",id,v.main,v.name,joinSet(v.ruleIds),joinSet(v.targetIds),v.products.size,v.inherited,v.sales,v.qty,v.exceptions,"Filter products below by NodeID"]));
  subValues.sort((a,b)=>text(a[3]).localeCompare(text(b[3]))||text(a[4]).localeCompare(text(b[4])));
  if(subValues.length) sheet.getRangeByIndexes(row,0,subValues.length,13).setValues(subValues); const st=sheet.addTable(`A${row}:M${Math.max(row,row+subValues.length)}`,true); st.setName("tblMappingSubcategoryNodes"); st.setPredefinedTableStyle("TableStyleMedium2");
  if (subValues.length) { sheet.getRange(`H${row + 1}:I${row + subValues.length}`).setNumberFormat("#,##0"); sheet.getRange(`J${row + 1}:K${row + subValues.length}`).setNumberFormat("#,##0.00"); sheet.getRange(`L${row + 1}:L${row + subValues.length}`).setNumberFormat("#,##0"); }
  row += Math.max(2, subValues.length + 3);
  sheet.getRange(`A${row}:N${row}`).setValues([["SourceSystemID","ScopeType","NodeID","SourceMainCategory","SourceSubCategory","Product","SalesAccount","Sales NOK (Active)","Quantity (Active)","Effective Reporting Group","Resolution Source","State","Status","Winning Rule"]]);
  const productValues = hierarchy.products.map(product => { const amount=context.get(product.productId)??{sales:0,quantity:0}; const r=resultByProduct.get(product.productId)!; return [product.sourceSystemId,"Product",product.productId,product.mainCategory,product.subCategory,product.productName,product.salesAccount,amount.sales,amount.quantity,r.effectiveGroupName,r.resolutionSource,r.resolutionState,r.resolutionStatus,r.winningRuleId]; });
  if(productValues.length) sheet.getRangeByIndexes(row,0,productValues.length,14).setValues(productValues); const pt=sheet.addTable(`A${row}:N${Math.max(row,row+productValues.length)}`,true); pt.setName("tblMappingProducts"); pt.setPredefinedTableStyle("TableStyleMedium2");
  if (productValues.length) sheet.getRange(`H${row + 1}:I${row + productValues.length}`).setNumberFormat("#,##0.00");
  sheet.getFreezePanes().freezeRows(4); setWidths(sheet,[110,135,220,170,180,150,145,105,105,110,105,120,180,110]);
}

function writeMappingLists(workbook: ExcelScript.Workbook, groups: ReportingGroup[]): void {
  const sheet = workbook.getWorksheet("_Mapping_Lists") ?? workbook.addWorksheet("_Mapping_Lists");
  sheet.getRange("A1:D1000").clear(ExcelScript.ClearApplyTo.all);
  sheet.getRange("A1:D1").setValues([["ActiveReportingGroupID", "ScopeType", "Action", "Status"]]); styleHeader(sheet.getRange("A1:D1"));
  const active = groups.filter(g=>g.active==="Yes").sort((a,b)=>a.sortOrder-b.sortOrder||a.name.localeCompare(b.name));
  if(active.length) sheet.getRange(`A2:A${active.length+1}`).setValues(active.map(g=>[g.id]));
  sheet.getRange("B2:B4").setValues([["SourceMainCategory"],["SourceSubCategory"],["Product"]]);
  sheet.getRange("C2:C3").setValues([["Apply"],["Deactivate"]]); sheet.getRange("D2:D3").setValues([["Active"],["Inactive"]]);
}

function wireMappingValidation(workbook: ExcelScript.Workbook): void {
  const sheet = requiredSheet(workbook,"Mapping");
  const listsSheet = requiredSheet(workbook,"_Mapping_Lists");
  const groupCount=countPopulated(listsSheet.getRange("A2:A1000").getValues());
  const actionSource = listsSheet.getRange("C2:C3");
  const scopeSource = listsSheet.getRange("B2:B4");
  const activeGroupSource = listsSheet.getRange(`A2:A${Math.max(2,groupCount+1)}`);
  const statusSource = listsSheet.getRange("D2:D3");
  sheet.getRange("B5").getDataValidation().setRule({list:{inCellDropDown:true,source:actionSource}});
  sheet.getRange("B7").getDataValidation().setRule({list:{inCellDropDown:true,source:scopeSource}});
  sheet.getRange("B9").getDataValidation().setRule({list:{inCellDropDown:true,source:activeGroupSource}});
  const rulesSheet = requiredSheet(workbook,"Mapping Rules");
  rulesSheet.getRange("C5:C1000").getDataValidation().setRule({list:{inCellDropDown:true,source:scopeSource}});
  rulesSheet.getRange("G5:G1000").getDataValidation().setRule({list:{inCellDropDown:true,source:activeGroupSource}});
  rulesSheet.getRange("J5:J1000").getDataValidation().setRule({list:{inCellDropDown:true,source:statusSource}});
}

function writeMappingQA(
  workbook: ExcelScript.Workbook, groups: ReportingGroup[], rules: MappingRule[], conflicts: Conflict[],
  resolved: Resolution[], before: FactSnapshot, context: Map<string,{sales:number;quantity:number}>
): void {
  const sheet=resetOutputSheet(workbook,"Mapping QA","tblMappingQA","tblMappingConflicts"); writeTitle(sheet,"Mapping QA","Deterministic Phase 1 checks. Mapping redistributes classification only; raw/fact totals remain unchanged.","H");
  const duplicateGroupIds=groups.length-new Set(groups.map(g=>g.id)).size; const mappedIds=new Set(resolved.filter(r=>r.resolutionStatus==="Mapped"||r.resolutionStatus==="Inactive Target").map(r=>r.productId));
  let mappedSales=0,mappedQty=0,unmappedSales=0,unmappedQty=0;
  context.forEach((v,productId) => { if(mappedIds.has(productId)){mappedSales+=v.sales;mappedQty+=v.quantity;}else{unmappedSales+=v.sales;unmappedQty+=v.quantity;} });
  const checks:(string|number)[][]=[
    ["QA-0301-01","Reporting Group IDs unique",duplicateGroupIds===0?"PASS":"FAIL",duplicateGroupIds,"Stable IDs must be unique."],
    ["QA-0301-02","Seed registry present",groups.length>=9?"PASS":"FAIL",groups.length,"At least the nine Phase 1 seeds must exist."],
    ["QA-0301-03","Overlapping active conflicts surfaced",conflicts.length===0?"PASS":"WARN",conflicts.length,"Conflicts are listed below and never row-order resolved."],
    ["QA-0301-04","Unmapped visible","PASS",resolved.filter(r=>r.resolutionStatus==="Unmapped").length,"Unmapped is an explicit status."],
    ["QA-0301-05","Inactive targets visible","PASS",resolved.filter(r=>r.resolutionStatus==="Inactive Target").length,"Inactive mappings remain stored and are surfaced."],
    ["QA-0301-06","Mapped + unmapped sales reconcile",almostEqual(mappedSales+unmappedSales,before.activeSales)?"PASS":"FAIL",mappedSales+unmappedSales-before.activeSales,"Must equal Active Finalized tblSalesFacts SalesAmount total."],
    ["QA-0301-07","Mapped + unmapped quantity reconcile",almostEqual(mappedQty+unmappedQty,before.activeQuantity)?"PASS":"FAIL",mappedQty+unmappedQty-before.activeQuantity,"Must equal Active Finalized tblSalesFacts Quantity total."],
    ["QA-0301-08","Fact row count retained","PASS",before.rows,"Script snapshots and rechecks facts before completion."],
    ["QA-0301-09","Performance migration deferred","PASS",0,"Phase 1 leaves Performance/Reports and _Metric_Calc category logic unchanged."]
  ];
  sheet.getRange("A4:E4").setValues([["CheckID","Check","Result","Observed","Explanation"]]); sheet.getRangeByIndexes(4,0,checks.length,5).setValues(checks); const qt=sheet.addTable(`A4:E${checks.length+4}`,true);qt.setName("tblMappingQA");qt.setPredefinedTableStyle("TableStyleMedium2");
  const start=checks.length+8; sheet.getRange(`A${start}:E${start}`).setValues([["LeftRuleID","RightRuleID","ScopeType","NodeID","Status"]]);
  const conflictRows=conflicts.map(c=>[c.leftRuleId,c.rightRuleId,c.scopeType,c.nodeId,"Conflict"]); if(conflictRows.length)sheet.getRangeByIndexes(start,0,conflictRows.length,5).setValues(conflictRows);
  const ct=sheet.addTable(`A${start}:E${Math.max(start,start+conflictRows.length)}`,true);ct.setName("tblMappingConflicts");ct.setPredefinedTableStyle("TableStyleMedium3");
  sheet.getFreezePanes().freezeRows(4);setWidths(sheet,[100,250,80,120,310,90,90,90]);
}

function aggregateSalesContext(facts: ExcelScript.Table): Map<string,{sales:number;quantity:number}> {
  const h=headerMap(facts); const map=new Map<string,{sales:number;quantity:number}>();
  for(const row of facts.getRangeBetweenHeaderAndTotal().getValues()){
    if(text(row[h.PublicationState])!=="Active Finalized")continue; const id=text(row[h.ProductID]); if(!id)continue;
    const current=map.get(id)??{sales:0,quantity:0};current.sales+=numberValue(row[h.SalesAmount]);current.quantity+=numberValue(row[h.Quantity]);map.set(id,current);
  } return map;
}

function snapshotFacts(table: ExcelScript.Table): FactSnapshot {
  const h=headerMap(table);const rows=table.getRangeBetweenHeaderAndTotal().getValues();let quantity=0,sales=0,activeQuantity=0,activeSales=0,firstId="",lastId="";
  for(const row of rows){const id=text(row[h.SalesFactID]);if(!firstId&&id)firstId=id;if(id)lastId=id;const q=numberValue(row[h.Quantity]);const s=numberValue(row[h.SalesAmount]);quantity+=q;sales+=s;if(text(row[h.PublicationState])==="Active Finalized"){activeQuantity+=q;activeSales+=s;}}
  return{rows:rows.length,quantity,sales,activeQuantity,activeSales,firstId,lastId};
}
function sameFactSnapshot(a:FactSnapshot,b:FactSnapshot):boolean{return a.rows===b.rows&&almostEqual(a.quantity,b.quantity)&&almostEqual(a.sales,b.sales)&&almostEqual(a.activeQuantity,b.activeQuantity)&&almostEqual(a.activeSales,b.activeSales)&&a.firstId===b.firstId&&a.lastId===b.lastId;}
function almostEqual(a:number,b:number):boolean{return Math.abs(a-b)<=Math.max(0.000001,Math.abs(b)*1e-12);}
function ids(rules:MappingRule[]):string{return rules.map(r=>r.id).join(", ");}
function targets(rules:MappingRule[]):string{return rules.map(r=>r.targetGroupId).join(", ");}
function addDelimited(target:Set<string>,value:string):void{for(const item of value.split(",")){const normalized=item.trim();if(normalized)target.add(normalized);}}
function joinSet(values:Set<string>):string{const items:string[]=[];values.forEach(value=>items.push(value));return items.join(", ");}
function mainNodeId(sourceSystemId:string,mainCategory:string):string{return `${sourceSystemId} || Main || ${mainCategory}`;}

function resetOutputSheet(workbook:ExcelScript.Workbook,name:string,...tableNames:string[]):ExcelScript.Worksheet{
  for(const tableName of tableNames){const table=workbook.getTable(tableName);if(table)table.delete();}
  const sheet=workbook.getWorksheet(name)??workbook.addWorksheet(name);const used=sheet.getUsedRange();if(used)used.clear(ExcelScript.ClearApplyTo.all);return sheet;
}
function requiredSheet(workbook:ExcelScript.Workbook,name:string):ExcelScript.Worksheet{const sheet=workbook.getWorksheet(name);if(!sheet)throw new Error(`PUL-0301-009: Required sheet missing: ${name}`);return sheet;}
function requiredTable(workbook:ExcelScript.Workbook,name:string):ExcelScript.Table{const table=workbook.getTable(name);if(!table)throw new Error(`PUL-0301-010: Required table missing: ${name}`);return table;}
function headerMap(table:ExcelScript.Table):{[key:string]:number}{const headers=table.getHeaderRowRange().getValues()[0];const map:{[key:string]:number}={};for(let i=0;i<headers.length;i++)map[text(headers[i])]=i;return map;}
function writeTitle(sheet:ExcelScript.Worksheet,title:string,subtitle:string,endColumn:string):void{sheet.getRange(`A1:${endColumn}1`).getFormat().getFill().setColor("#172033");sheet.getRange(`A1:${endColumn}1`).getFormat().getFont().setColor("#FFFFFF");sheet.getRange(`A1:${endColumn}1`).getFormat().getFont().setBold(true);sheet.getRange(`A1:${endColumn}1`).getFormat().getFont().setSize(18);sheet.getRange("A1").setValue(title);sheet.getRange(`A2:${endColumn}2`).getFormat().getFill().setColor("#EAF2FF");sheet.getRange(`A2:${endColumn}2`).getFormat().setWrapText(true);sheet.getRange("A2").setValue(subtitle);sheet.getRange("1:1").getFormat().setRowHeight(32);sheet.getRange("2:2").getFormat().setRowHeight(30);}
function styleHeader(range:ExcelScript.Range):void{range.getFormat().getFill().setColor("#4F8CFF");range.getFormat().getFont().setColor("#FFFFFF");range.getFormat().getFont().setBold(true);range.getFormat().setWrapText(true);}
function setWidths(sheet:ExcelScript.Worksheet,widths:number[]):void{for(let i=0;i<widths.length;i++)sheet.getRangeByIndexes(0,i,1,1).getEntireColumn().getFormat().setColumnWidth(widths[i]);}
function countPopulated(values:(string|number|boolean)[][]):number{let count=0;for(const row of values)if(text(row[0]))count++;return count;}
function updateEnvironment(table:ExcelScript.Table,key:string,value:string,note:string):void{const rows=table.getRangeBetweenHeaderAndTotal().getValues();for(let i=0;i<rows.length;i++){if(text(rows[i][0])===key){table.getRangeBetweenHeaderAndTotal().getCell(i,1).setValue(value);table.getRangeBetweenHeaderAndTotal().getCell(i,2).setValue(note);return;}}table.addRow(-1,[key,value,note]);}
function appendBuildLog(table:ExcelScript.Table,row:(string|number|boolean)[]):void{table.addRow(-1,row);}
function nextId(table:ExcelScript.Table,prefix:string,digits:number):string{let max=0;for(const row of table.getRangeBetweenHeaderAndTotal().getValues()){const match=text(row[0]).match(new RegExp(`^${prefix}-(\\d+)$`));if(match)max=Math.max(max,Number(match[1]));}return `${prefix}-${String(max+1).padStart(digits,"0")}`;}
function excelNow():number{return Date.now()/86400000+25569;}
function excelToday():number{return Math.floor(excelNow());}
function numberValue(value:unknown):number{const n=Number(value);return Number.isFinite(n)?n:0;}
function text(value:unknown):string{return String(value??"").trim();}
