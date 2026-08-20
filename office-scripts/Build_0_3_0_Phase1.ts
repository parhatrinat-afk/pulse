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
 * - Mapping exposes a name-first category overview and a bounded member
 *   workspace over the stable-ID hierarchy. Requested bulk changes are fully
 *   validated before authoritative rules change.
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
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  const workspaceState = mappingSheet ? readMappingWorkspaceState(workbook, mappingSheet) : emptyWorkspaceState();
  const mappingActionRequested = workspaceState.intent !== "";
  const actionMessage = mappingSheet
    ? applyPendingAction(workbook, mappingSheet, mappingRules, groups, sourceClassifications, products, workspaceState)
    : "No pending mapping action.";
  if (mappingActionRequested) {
    workspaceState.intent = "";
    workspaceState.targetGroupName = "";
    workspaceState.notes = "";
    workspaceState.selectedTargets = [];
  }

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
  writeMappingSurface(workbook, hierarchy, resolved, salesContext, groups, rules, actionMessage, workspaceState);
  writeMappingLists(workbook, groups);
  const validationMessage = wireMappingValidation(workbook);
  restoreWeeklyMappingHealthBlock(workbook, mappingActionRequested);
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
  return `Pulse 0.3.0 Phase 1 refreshed. ${groups.length} Reporting Groups, ${rules.length} rule(s), ${resolved.length} product(s). ${actionMessage} ${validationMessage}`;
}

type ReportingGroup = { id: string; name: string; active: string; sortOrder: number };
type MappingRule = {
  id: string; sourceSystemId: string; scopeType: string; nodeId: string;
  nodeDisplay: string; hierarchyLevel: number; targetGroupId: string; ruleAction: string;
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
type MappingSelectionTarget = {
  sourceSystemId: string; scopeType: string; nodeId: string; productId: string;
  parentSubNodeId: string; identityState: string; currentExplicitRuleId: string;
  currentRuleAction: string; currentTargetId: string; selectedName: string;
  mappingState: string;
};
type MappingWorkspaceState = {
  view: string; categoryChoice: string; applyTo: string;
  intent: string; targetGroupName: string; notes: string;
  selectedTargets: MappingSelectionTarget[];
};

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
    sheet.getRange("A1:M20").clear(ExcelScript.ClearApplyTo.all);
    writeTitle(sheet, "Mapping Rules", "Authoritative explicit rules. Product exclusions resolve to Unmapped; all other descendants inherit computed state.", "M");
    sheet.getRange("A4:M4").setValues([[
      "MappingRuleID", "SourceSystemID", "ScopeType", "NodeID", "NodeDisplay", "HierarchyLevel",
      "TargetReportingGroupID", "EffectiveFrom", "EffectiveTo", "Status", "CreatedAt", "Notes", "RuleAction"
    ]]);
    table = sheet.addTable("A4:M4", true);
    table.setName("tblMappingRules");
    table.setPredefinedTableStyle("TableStyleMedium2");
    table.getRangeBetweenHeaderAndTotal().clear(ExcelScript.ClearApplyTo.contents);
    sheet.getRange("H5:I1000").setNumberFormatLocal("dd.mm.yyyy");
    sheet.getRange("K5:K1000").setNumberFormatLocal("dd.mm.yyyy hh:mm");
    sheet.getFreezePanes().freezeRows(4);
    setWidths(sheet, [105, 110, 130, 210, 210, 85, 145, 100, 100, 85, 120, 240, 90]);
  } else if (headerMap(table).RuleAction === undefined) {
    table.addColumn(-1, undefined, "RuleAction");
  }
  const ruleHeaders = headerMap(table);
  const ruleValues = table.getRangeBetweenHeaderAndTotal().getValues();
  const actionRange = table.getColumn("RuleAction").getRangeBetweenHeaderAndTotal();
  const actionValues = actionRange.getValues();
  for (let i = 0; i < actionValues.length; i++) {
    if (text(ruleValues[i][ruleHeaders.MappingRuleID]) && !text(actionValues[i][0])) actionValues[i][0] = "Map";
  }
  if (actionValues.length) actionRange.setValues(actionValues);
  return table;
}

function markLegacyCategories(sheet: ExcelScript.Worksheet): void {
  sheet.getRange("A2").setValue("Legacy/source-default classification retained for Build 0.2.0 compatibility. Reporting Groups are the new business semantic layer.");
  sheet.getRange("A2:G2").getFormat().getFill().setColor("#FFF4CE");
  sheet.getRange("A2:G2").getFormat().setWrapText(true);
  sheet.getRange("2:2").getFormat().setRowHeight(34);
}

function applyPendingAction(
  workbook: ExcelScript.Workbook, sheet: ExcelScript.Worksheet, table: ExcelScript.Table,
  groups: ReportingGroup[], classifications: ExcelScript.Table, products: ExcelScript.Table,
  state: MappingWorkspaceState
): string {
  const intent = state.intent;
  if (!intent) return "No pending mapping action.";
  const allowed = ["Assign Reporting Group", "Leave Unmapped", "Remove custom mapping"];
  if (allowed.indexOf(intent) < 0) throw new Error(`PUL-0301-012: Unsupported Mapping action ${intent}.`);
  const selected = state.selectedTargets;
  if (!selected.length) throw new Error("PUL-0301-003: Select at least one safe category member before applying a mapping action.");
  const levelByScope: { [key: string]: number } = { SourceMainCategory: 1, SourceSubCategory: 2, Product: 3 };
  const targetMatches = groups.filter(group => group.name === state.targetGroupName && group.active === "Yes");
  const needsTarget = intent === "Assign Reporting Group";
  if (needsTarget && targetMatches.length !== 1) {
    throw new Error(`PUL-0301-005: Select one active Reporting Group by business name; found ${targetMatches.length}.`);
  }
  const h = headerMap(table);
  const body = table.getRangeBetweenHeaderAndTotal();
  const rows = body.getValues();
  const activeByNode: { [key: string]: number[] } = {};
  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];
    if (text(row[h.Status]) !== "Active" ||
        (numberValue(row[h.EffectiveFrom]) && numberValue(row[h.EffectiveFrom]) > excelToday()) ||
        (numberValue(row[h.EffectiveTo]) && numberValue(row[h.EffectiveTo]) < excelToday())) continue;
    const key = mappingNodeKey(text(row[h.SourceSystemID]), text(row[h.ScopeType]), text(row[h.NodeID]));
    if (!activeByNode[key]) activeByNode[key] = [];
    activeByNode[key].push(index);
  }

  const selectedKeys: { [key: string]: boolean } = {};
  const selectedSubcategories: { [key: string]: boolean } = {};
  for (let index = 0; index < selected.length; index += 1) {
    const target = selected[index];
    if (!target.sourceSystemId || !levelByScope[target.scopeType] || !target.nodeId || !target.selectedName) {
      throw new Error("PUL-0301-004: The selected batch contains an invalid stable hierarchy target.");
    }
    const key = mappingNodeKey(target.sourceSystemId, target.scopeType, target.nodeId);
    if (selectedKeys[key]) throw new Error(`PUL-0301-022: Duplicate selected target ${target.selectedName}. No rule was changed.`);
    selectedKeys[key] = true;
    if (target.scopeType === "SourceSubCategory") selectedSubcategories[target.nodeId] = true;
    if (target.identityState === "Identity Pending" || target.mappingState === "Identity Pending") {
      throw new Error(`PUL-0301-017: ${target.selectedName} is Identity Pending and remains review-only. No rule was changed.`);
    }
    if (target.mappingState === "Conflict" || target.mappingState === "Inactive Target") {
      throw new Error(`PUL-0301-023: ${target.selectedName} is ${target.mappingState}; resolve that state before bulk mapping.`);
    }
    const activeIndexes = activeByNode[key] || [];
    if (activeIndexes.length > 1) {
      throw new Error(`PUL-0301-019: ${activeIndexes.length} active explicit rules overlap at ${target.selectedName}; resolve the conflict first.`);
    }
  }
  for (let index = 0; index < selected.length; index += 1) {
    const target = selected[index];
    if (target.scopeType === "Product" && target.parentSubNodeId && selectedSubcategories[target.parentSubNodeId]) {
      throw new Error(`PUL-0301-024: Do not select both ${target.selectedName} and its Subcategory ancestor in one batch.`);
    }
  }
  if (intent === "Leave Unmapped" && selected.some(target => target.scopeType !== "Product")) {
    throw new Error("PUL-0301-018: Leave Unmapped is supported only for selected Products.");
  }

  const targetGroupId = needsTarget ? targetMatches[0].id : "";
  const newRows: (string | number | boolean)[][] = [];
  const deactivate: number[] = [];
  let nextNumber = nextNumericId(table, "MAP") + 1;
  for (let index = 0; index < selected.length; index += 1) {
    const target = selected[index];
    const key = mappingNodeKey(target.sourceSystemId, target.scopeType, target.nodeId);
    const activeIndexes = activeByNode[key] || [];
    const currentIndex = activeIndexes.length === 1 ? activeIndexes[0] : -1;
    const currentAction = currentIndex >= 0 ? text(rows[currentIndex][h.RuleAction]) || "Map" : "";
    const currentTarget = currentIndex >= 0 ? text(rows[currentIndex][h.TargetReportingGroupID]) : "";
    if (intent === "Remove custom mapping") {
      if (currentIndex < 0) {
        throw new Error(`PUL-0301-021: ${target.selectedName} has no custom rule to remove; its current result is inherited or Unmapped.`);
      }
      deactivate.push(currentIndex);
      continue;
    }
    const ruleAction = intent === "Leave Unmapped" ? "Exclude" : "Map";
    if (intent === "Assign Reporting Group" && currentIndex < 0 && target.mappingState === "Inherited" &&
        target.currentTargetId === targetGroupId) continue;
    if (currentIndex >= 0 && currentAction === ruleAction &&
        (ruleAction === "Exclude" || currentTarget === targetGroupId)) continue;
    if (currentIndex >= 0) deactivate.push(currentIndex);
    const newRuleId = `MAP-${String(nextNumber).padStart(6, "0")}`;
    nextNumber += 1;
    newRows.push([
      newRuleId, target.sourceSystemId, target.scopeType, target.nodeId, target.selectedName,
      levelByScope[target.scopeType], ruleAction === "Exclude" ? "" : targetGroupId, excelToday(), "",
      "Active", excelNow(), state.notes, ruleAction
    ]);
  }

  if (!newRows.length && !deactivate.length) {
    clearWorkspaceAction(sheet);
    return `No change: all ${selected.length} selected item(s) already have the requested explicit state.`;
  }
  const originalStatuses = body.getColumn(h.Status).getValues();
  let appendedCount = 0;
  try {
    if (newRows.length) {
      table.addRows(-1, newRows);
      appendedCount = newRows.length;
    }
    const updatedStatuses = originalStatuses.map(row => [row[0]]);
    for (let index = 0; index < deactivate.length; index += 1) updatedStatuses[deactivate[index]][0] = "Inactive";
    if (updatedStatuses.length) table.getColumn("Status").getRangeBetweenHeaderAndTotal()
      .getResizedRange(-(appendedCount), 0).setValues(updatedStatuses);
  } catch (error) {
    if (appendedCount) table.deleteRowsAt(table.getRowCount() - appendedCount, appendedCount);
    if (originalStatuses.length) table.getColumn("Status").getRangeBetweenHeaderAndTotal().setValues(originalStatuses);
    throw error;
  }
  clearWorkspaceAction(sheet);
  if (intent === "Leave Unmapped") return `Saved explicit Unmapped decisions for ${selected.length} selected Product(s).`;
  if (intent === "Remove custom mapping") return `Removed ${deactivate.length} custom rule(s); normal hierarchy inheritance is restored.`;
  return `Assigned ${selected.length} selected item(s) to ${state.targetGroupName}.`;
}

function readMappingWorkspaceState(workbook: ExcelScript.Workbook, sheet: ExcelScript.Worksheet): MappingWorkspaceState {
  const empty = emptyWorkspaceState();
  if (text(sheet.getRange("A15").getValue()) !== "Action") return empty;
  const memberTable = workbook.getTable("tblMappingMemberWorkspace");
  const categoryTable = workbook.getTable("tblMappingCategoryCatalog");
  const categoryChoice = text(sheet.getRange("B12").getValue());
  const applyTo = text(sheet.getRange("B13").getValue()) || "Selected members";
  const selectedTargets: MappingSelectionTarget[] = [];
  if (memberTable) {
    const h = headerMap(memberTable);
    const rows = memberTable.getRangeBetweenHeaderAndTotal().getValues();
    for (let index = 0; index < rows.length; index += 1) {
      const row = rows[index];
      if (text(row[h.Select]) !== "Yes") continue;
      selectedTargets.push({
        sourceSystemId: text(row[h.SourceSystemID]), scopeType: text(row[h.ScopeType]), nodeId: text(row[h.NodeID]),
        productId: text(row[h.ProductID]), parentSubNodeId: text(row[h.ParentSubNodeID]),
        identityState: text(row[h.IdentityState]), currentExplicitRuleId: text(row[h.CurrentExplicitRuleID]),
        currentRuleAction: text(row[h.CurrentRuleAction]), currentTargetId: text(row[h.CurrentTargetID]),
        selectedName: text(row[h.Item]), mappingState: text(row[h["Mapping state"]])
      });
    }
  }
  if (applyTo === "Entire shown category") {
    if (selectedTargets.length) {
      throw new Error("PUL-0301-025: Clear member selections before choosing Entire shown category.");
    }
    if (!categoryTable) throw new Error("PUL-0301-026: Category overview is unavailable.");
    const h = headerMap(categoryTable);
    const rows = categoryTable.getRangeBetweenHeaderAndTotal().getValues();
    for (let index = 0; index < rows.length; index += 1) {
      const row = rows[index];
      if (text(row[h.CategoryChoice]) !== categoryChoice) continue;
      const nodeId = text(row[h.MainNodeID]);
      if (!nodeId) throw new Error("PUL-0301-027: The shown category has no approved stable Main Category node.");
      selectedTargets.push({
        sourceSystemId: text(row[h.SourceSystemID]), scopeType: "SourceMainCategory", nodeId,
        productId: "", parentSubNodeId: "", identityState: "", currentExplicitRuleId: text(row[h.CurrentExplicitRuleID]),
        currentRuleAction: text(row[h.CurrentRuleAction]), currentTargetId: text(row[h.CurrentTargetID]),
        selectedName: text(row[h["Main Category"]]), mappingState: text(row[h["Mapping state"]])
      });
      break;
    }
  }
  return { view: text(sheet.getRange("E12").getValue()) || "All", categoryChoice, applyTo,
    intent: text(sheet.getRange("B15").getValue()), targetGroupName: text(sheet.getRange("E15").getValue()),
    notes: text(sheet.getRange("H15").getValue()), selectedTargets };
}

function emptyWorkspaceState(): MappingWorkspaceState {
  return { view: "All", categoryChoice: "", applyTo: "Selected members", intent: "",
    targetGroupName: "", notes: "", selectedTargets: [] };
}

function clearWorkspaceAction(sheet: ExcelScript.Worksheet): void {
  sheet.getRange("B15").clear(ExcelScript.ClearApplyTo.contents);
  sheet.getRange("E15").clear(ExcelScript.ClearApplyTo.contents);
  sheet.getRange("H15:J16").clear(ExcelScript.ClearApplyTo.contents);
  const table = sheet.getTables().find(candidate => candidate.getName() === "tblMappingMemberWorkspace");
  if (table) table.getColumn("Select").getRangeBetweenHeaderAndTotal().clear(ExcelScript.ClearApplyTo.contents);
}

function mappingNodeKey(sourceSystemId: string, scopeType: string, nodeId: string): string {
  return `${sourceSystemId}\u001f${scopeType}\u001f${nodeId}`;
}

function nextNumericId(table: ExcelScript.Table, prefix: string): number {
  const values = table.getColumn("MappingRuleID").getRangeBetweenHeaderAndTotal().getValues();
  let maximum = 0;
  for (let index = 0; index < values.length; index += 1) {
    const value = text(values[index][0]);
    if (value.indexOf(`${prefix}-`) !== 0) continue;
    const parsed = Number(value.substring(prefix.length + 1));
    if (Number.isFinite(parsed) && parsed > maximum) maximum = parsed;
  }
  return maximum;
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
  const rules = table.getRangeBetweenHeaderAndTotal().getValues().filter(row => text(row[h.MappingRuleID])).map(row => ({
    id: text(row[h.MappingRuleID]), sourceSystemId: text(row[h.SourceSystemID]), scopeType: text(row[h.ScopeType]),
    nodeId: text(row[h.NodeID]), nodeDisplay: text(row[h.NodeDisplay]), hierarchyLevel: numberValue(row[h.HierarchyLevel]),
    targetGroupId: text(row[h.TargetReportingGroupID]), effectiveFrom: numberValue(row[h.EffectiveFrom]),
    effectiveTo: numberValue(row[h.EffectiveTo]), status: text(row[h.Status]), createdAt: numberValue(row[h.CreatedAt]),
    notes: text(row[h.Notes]), ruleAction: text(row[h.RuleAction]) || "Map"
  }));
  for (const rule of rules) validateRuleAction(rule);
  return rules;
}

function validateRuleAction(rule: MappingRule): void {
  if (rule.ruleAction !== "Map" && rule.ruleAction !== "Exclude") {
    throw new Error(`PUL-0301-014: Rule ${rule.id} has unsupported RuleAction ${rule.ruleAction}.`);
  }
  if (rule.ruleAction === "Exclude" && (rule.scopeType !== "Product" || !!rule.targetGroupId)) {
    throw new Error(`PUL-0301-015: Rule ${rule.id} must be a Product exclusion with a blank Reporting Group target.`);
  }
  if (rule.ruleAction === "Map" && !rule.targetGroupId) {
    throw new Error(`PUL-0301-016: Map rule ${rule.id} requires a Reporting Group target.`);
  }
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
      const rule = candidates[0];
      if (rule.ruleAction === "Exclude") {
        return { ...base, effectiveGroupId: "", effectiveGroupName: "", resolutionSource: "Product",
          resolutionState: "Explicit exclusion", resolutionStatus: "Unmapped", winningRuleId: rule.id };
      }
      const group = groups.get(rule.targetGroupId);
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

type MappingWorkspaceProduct = {
  productId: string; sourceSystemId: string; item: string; mainCategory: string; subcategory: string;
  salesAccount: string; sourceClassificationId: string; identityState: string; facts: number; sales: number;
  quantity: number; hierarchyAttention: string; hierarchyAlternatives: string; resolution: Resolution;
};
type MappingWorkspaceMember = {
  level: string; item: string; subcategory: string; salesAccount: string; reportingGroup: string;
  mappingState: string; facts: number; sales: number; quantity: number; attention: string;
  sourceSystemId: string; scopeType: string; nodeId: string; productId: string; parentSubNodeId: string;
  identityState: string; currentExplicitRuleId: string; currentRuleAction: string; currentTargetId: string;
  categoryChoice: string;
};
type MappingWorkspaceCategory = {
  mainCategory: string; currentGroup: string; mappingState: string; subcategories: number; products: number;
  facts: number; sales: number; attention: string; sourceSystemId: string; mainNodeId: string;
  categoryChoice: string; currentExplicitRuleId: string; currentRuleAction: string; currentTargetId: string;
};

function buildMappingWorkspaceData(
  attention: ExcelScript.Table, groups: ReportingGroup[], rules: MappingRule[]
): { categories: MappingWorkspaceCategory[]; members: MappingWorkspaceMember[] } {
  const h = headerMap(attention);
  const rows = attention.getRangeBetweenHeaderAndTotal().getValues();
  const asOf = excelToday();
  const groupById = new Map<string, ReportingGroup>();
  for (let index = 0; index < groups.length; index += 1) groupById.set(groups[index].id, groups[index]);
  const activeByNode: { [key: string]: MappingRule[] } = {};
  for (let index = 0; index < rules.length; index += 1) {
    const rule = rules[index];
    if (rule.status !== "Active" || (rule.effectiveFrom && rule.effectiveFrom > asOf) ||
        (rule.effectiveTo && rule.effectiveTo < asOf)) continue;
    const key = mappingNodeKey(rule.sourceSystemId, rule.scopeType, rule.nodeId);
    if (!activeByNode[key]) activeByNode[key] = [];
    activeByNode[key].push(rule);
  }

  const products: MappingWorkspaceProduct[] = [];
  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];
    const productId = text(row[h.ProductID]);
    if (!productId) continue;
    const sourceSystemId = text(row[h.SourceSystemID]);
    const item = text(row[h.Item]);
    const mainCategory = text(row[h["Main Category"]]);
    const subcategory = text(row[h.Subcategory]);
    const sourceClassificationId = text(row[h.SourceClassificationID]);
    const node: ProductNode = {
      productId, sourceSystemId, productName: item, salesAccount: text(row[h["Sales Account"]]),
      sourceClassificationId, mainCategory, subCategory: subcategory,
      mainNodeId: mainNodeId(sourceSystemId, mainCategory), subNodeId: sourceClassificationId
    };
    const identityPending = text(row[h["Mapping State"]]) === "Identity Pending";
    const resolution = identityPending ? {
      ...node, mainRuleIds: "", mainTargetIds: "", subRuleIds: "", subTargetIds: "", productRuleIds: "",
      productTargetIds: "", effectiveGroupId: "", effectiveGroupName: "", resolutionSource: "Identity Pending",
      resolutionState: "Identity Pending", resolutionStatus: "Identity Pending", winningRuleId: ""
    } : resolveProduct(node, rules, groupById, asOf);
    products.push({
      productId, sourceSystemId, item, mainCategory, subcategory, salesAccount: node.salesAccount,
      sourceClassificationId, identityState: identityPending ? "Identity Pending" : "",
      facts: numberValue(row[h["Historical Facts"]]), sales: numberValue(row[h["Historical Sales NOK"]]),
      quantity: numberValue(row[h["Historical Quantity"]]), hierarchyAttention: text(row[h["Hierarchy Attention"]]),
      hierarchyAlternatives: text(row[h["Hierarchy Alternatives"]]), resolution
    });
  }

  type StateCounts = { custom: number; unmapped: number; pending: number; excluded: number; conflict: number; inactive: number };
  type CategoryAccumulator = {
    sourceSystemId: string; mainCategory: string; displayMain: string; mainNodeId: string; categoryChoice: string;
    productCount: number; subIds: Set<string>; facts: number; sales: number; quantity: number; states: StateCounts;
  };
  type SubcategoryAccumulator = {
    sourceSystemId: string; mainCategory: string; subcategory: string; nodeId: string; categoryKey: string;
    productCount: number; facts: number; sales: number; quantity: number; states: StateCounts;
  };
  const categories = new Map<string, CategoryAccumulator>();
  const subcategories = new Map<string, SubcategoryAccumulator>();
  const displayCounts: { [key: string]: number } = {};
  const emptyCounts = (): StateCounts => ({ custom: 0, unmapped: 0, pending: 0, excluded: 0, conflict: 0, inactive: 0 });
  const presentationState = (resolution: Resolution): string => {
    if (resolution.resolutionStatus === "Identity Pending") return "Identity Pending";
    if (resolution.resolutionStatus === "Conflict") return "Conflict";
    if (resolution.resolutionStatus === "Inactive Target") return "Inactive Target";
    if (resolution.resolutionStatus === "Unmapped" && resolution.resolutionState === "Explicit exclusion") return "Excluded";
    if (resolution.resolutionStatus === "Unmapped") return "Unmapped";
    if (resolution.resolutionSource === "Product") return "Custom";
    return "Inherited";
  };
  const incrementState = (counts: StateCounts, state: string): void => {
    if (state === "Custom") counts.custom += 1;
    else if (state === "Unmapped") counts.unmapped += 1;
    else if (state === "Identity Pending") counts.pending += 1;
    else if (state === "Excluded") counts.excluded += 1;
    else if (state === "Conflict") counts.conflict += 1;
    else if (state === "Inactive Target") counts.inactive += 1;
  };
  for (let index = 0; index < products.length; index += 1) {
    const product = products[index];
    const displayMain = product.mainCategory || "(No approved Main Category)";
    const categoryKey = `${product.sourceSystemId}\u001f${product.mainCategory}`;
    let category = categories.get(categoryKey);
    if (!category) {
      category = { sourceSystemId: product.sourceSystemId, mainCategory: product.mainCategory, displayMain,
        mainNodeId: product.mainCategory ? mainNodeId(product.sourceSystemId, product.mainCategory) : "",
        categoryChoice: displayMain, productCount: 0, subIds: new Set<string>(), facts: 0, sales: 0, quantity: 0,
        states: emptyCounts() };
      categories.set(categoryKey, category);
      displayCounts[displayMain] = (displayCounts[displayMain] || 0) + 1;
    }
    category.productCount += 1; category.facts += product.facts; category.sales += product.sales; category.quantity += product.quantity;
    if (product.sourceClassificationId) category.subIds.add(product.sourceClassificationId);
    incrementState(category.states, presentationState(product.resolution));
    if (product.sourceClassificationId) {
      const subKey = mappingNodeKey(product.sourceSystemId, "SourceSubCategory", product.sourceClassificationId);
      let sub = subcategories.get(subKey);
      if (!sub) {
        sub = { sourceSystemId: product.sourceSystemId, mainCategory: product.mainCategory,
          subcategory: product.subcategory || "(No approved Subcategory)", nodeId: product.sourceClassificationId,
          categoryKey, productCount: 0, facts: 0, sales: 0, quantity: 0, states: emptyCounts() };
        subcategories.set(subKey, sub);
      }
      sub.productCount += 1; sub.facts += product.facts; sub.sales += product.sales; sub.quantity += product.quantity;
      incrementState(sub.states, presentationState(product.resolution));
    }
  }
  categories.forEach(category => {
    category.categoryChoice = displayCounts[category.displayMain] > 1
      ? `${category.displayMain} · ${category.sourceSystemId}` : category.displayMain;
  });

  const categoryChoiceByKey: { [key: string]: string } = {};
  categories.forEach((category, key) => { categoryChoiceByKey[key] = category.categoryChoice; });
  const nodePresentation = (sourceSystemId: string, scopeType: string, nodeId: string, parentNodeId: string): {
    group: string; state: string; ruleId: string; ruleAction: string; targetId: string;
  } => {
    const own = nodeId ? (activeByNode[mappingNodeKey(sourceSystemId, scopeType, nodeId)] || []) : [];
    const parent = parentNodeId ? (activeByNode[mappingNodeKey(sourceSystemId, "SourceMainCategory", parentNodeId)] || []) : [];
    if (own.length > 1) return { group: "Conflict", state: "Conflict", ruleId: ids(own), ruleAction: "", targetId: "" };
    if (own.length === 1) {
      const rule = own[0];
      if (rule.ruleAction === "Exclude") return { group: "Unmapped", state: "Excluded", ruleId: rule.id, ruleAction: rule.ruleAction, targetId: "" };
      const group = groupById.get(rule.targetGroupId);
      if (!group || group.active !== "Yes") return { group: "Inactive Target", state: "Inactive Target", ruleId: rule.id, ruleAction: rule.ruleAction, targetId: rule.targetGroupId };
      return { group: group.name, state: "Custom", ruleId: rule.id, ruleAction: rule.ruleAction, targetId: rule.targetGroupId };
    }
    if (parent.length > 1) return { group: "Conflict", state: "Conflict", ruleId: "", ruleAction: "", targetId: "" };
    if (parent.length === 1) {
      const rule = parent[0];
      const group = groupById.get(rule.targetGroupId);
      if (!group || group.active !== "Yes") return { group: "Inactive Target", state: "Inactive Target", ruleId: "", ruleAction: "", targetId: "" };
      return { group: group.name, state: "Inherited", ruleId: "", ruleAction: "", targetId: rule.targetGroupId };
    }
    return { group: "Unmapped", state: "Unmapped", ruleId: "", ruleAction: "", targetId: "" };
  };
  const attentionSummary = (counts: StateCounts): string => {
    const parts: string[] = [];
    if (counts.unmapped) parts.push(`${counts.unmapped} Unmapped`);
    if (counts.pending) parts.push(`${counts.pending} Identity Pending`);
    if (counts.excluded) parts.push(`${counts.excluded} Excluded`);
    if (counts.custom) parts.push(`${counts.custom} Custom`);
    if (counts.conflict) parts.push(`${counts.conflict} Conflict`);
    if (counts.inactive) parts.push(`${counts.inactive} Inactive Target`);
    return parts.length ? parts.join(" · ") : "No current attention";
  };

  const categoryRows: MappingWorkspaceCategory[] = [];
  categories.forEach(category => {
    const current = nodePresentation(category.sourceSystemId, "SourceMainCategory", category.mainNodeId, "");
    categoryRows.push({
      mainCategory: category.categoryChoice, currentGroup: current.state === "Unmapped" ? "—" : current.group,
      mappingState: current.state === "Unmapped" ? "No category rule" : current.state,
      subcategories: category.subIds.size, products: category.productCount, facts: category.facts, sales: category.sales,
      attention: attentionSummary(category.states), sourceSystemId: category.sourceSystemId, mainNodeId: category.mainNodeId,
      categoryChoice: category.categoryChoice, currentExplicitRuleId: current.ruleId,
      currentRuleAction: current.ruleAction, currentTargetId: current.targetId
    });
  });
  categoryRows.sort((left, right) => left.categoryChoice.localeCompare(right.categoryChoice));

  const memberRows: MappingWorkspaceMember[] = [];
  subcategories.forEach(sub => {
    const parentNodeId = sub.mainCategory ? mainNodeId(sub.sourceSystemId, sub.mainCategory) : "";
    const current = nodePresentation(sub.sourceSystemId, "SourceSubCategory", sub.nodeId, parentNodeId);
    memberRows.push({
      level: "Subcategory", item: sub.subcategory, subcategory: "—", salesAccount: "", reportingGroup: current.group,
      mappingState: current.state, facts: sub.facts, sales: sub.sales, quantity: sub.quantity,
      attention: attentionSummary(sub.states), sourceSystemId: sub.sourceSystemId, scopeType: "SourceSubCategory",
      nodeId: sub.nodeId, productId: "", parentSubNodeId: "", identityState: "",
      currentExplicitRuleId: current.ruleId, currentRuleAction: current.ruleAction, currentTargetId: current.targetId,
      categoryChoice: categoryChoiceByKey[sub.categoryKey]
    });
  });
  for (let index = 0; index < products.length; index += 1) {
    const product = products[index];
    const state = presentationState(product.resolution);
    const own = activeByNode[mappingNodeKey(product.sourceSystemId, "Product", product.productId)] || [];
    const ruleId = own.length ? ids(own) : "";
    const ruleAction = own.length === 1 ? own[0].ruleAction : "";
    const currentTargetId = own.length === 1 ? own[0].targetGroupId : "";
    let productAttention = product.hierarchyAttention;
    if (state === "Identity Pending") productAttention = product.hierarchyAlternatives
      ? `Identity review required · ${product.hierarchyAlternatives}` : "Identity review required";
    else if (!productAttention && state === "Unmapped") productAttention = "Needs a Reporting Group";
    else if (!productAttention && state === "Excluded") productAttention = "Intentional explicit exclusion";
    const categoryKey = `${product.sourceSystemId}\u001f${product.mainCategory}`;
    memberRows.push({
      level: "Product", item: product.item, subcategory: product.subcategory || "(No approved Subcategory)",
      salesAccount: product.salesAccount, reportingGroup: product.resolution.effectiveGroupName ||
        (state === "Excluded" || state === "Unmapped" ? "Unmapped" : state), mappingState: state,
      facts: product.facts, sales: product.sales, quantity: product.quantity, attention: productAttention,
      sourceSystemId: product.sourceSystemId, scopeType: "Product", nodeId: product.productId, productId: product.productId,
      parentSubNodeId: product.sourceClassificationId, identityState: product.identityState,
      currentExplicitRuleId: ruleId, currentRuleAction: ruleAction,
      currentTargetId: product.resolution.effectiveGroupId || currentTargetId,
      categoryChoice: categoryChoiceByKey[categoryKey]
    });
  }
  memberRows.sort((left, right) => left.categoryChoice.localeCompare(right.categoryChoice) ||
    (left.level === right.level ? 0 : left.level === "Subcategory" ? -1 : 1) ||
    left.subcategory.localeCompare(right.subcategory) || left.item.localeCompare(right.item) ||
    left.nodeId.localeCompare(right.nodeId));
  return { categories: categoryRows, members: memberRows };
}

function writeMappingSurface(
  workbook: ExcelScript.Workbook, hierarchy: { products: ProductNode[] }, resolved: Resolution[],
  context: Map<string, { sales: number; quantity: number }>, groups: ReportingGroup[], rules: MappingRule[],
  message: string, priorState: MappingWorkspaceState
): void {
  const attention = requiredTable(workbook, "tblWeeklyMappingAttention");
  const workspace = buildMappingWorkspaceData(attention, groups, rules);
  const memberCapacity = 150;
  let maximumMembers = 0;
  for (let categoryIndex = 0; categoryIndex < workspace.categories.length; categoryIndex += 1) {
    const choice = workspace.categories[categoryIndex].categoryChoice;
    const count = workspace.members.filter(member => member.categoryChoice === choice).length;
    if (count > maximumMembers) maximumMembers = count;
  }
  if (maximumMembers > memberCapacity) {
    throw new Error(`PUL-0301-028: The largest category has ${maximumMembers} members; the safe workspace capacity is ${memberCapacity}.`);
  }
  const sheet = resetOutputSheet(workbook, "Mapping", "tblMappingMainNodes", "tblMappingSubcategoryNodes",
    "tblMappingProducts", "tblMappingUXControl", "tblMappingCategoryOverview", "tblMappingCategoryCatalog",
    "tblMappingMemberWorkspace", "tblMappingMemberCatalog");
  const auditSheet = resetOutputSheet(workbook, "_Mapping_Audit", "tblMappingMainNodes", "tblMappingSubcategoryNodes",
    "tblMappingProducts", "tblMappingCategoryCatalog", "tblMappingMemberCatalog");
  auditSheet.setVisibility(ExcelScript.SheetVisibility.hidden);
  sheet.getRange("A1:N18").unmerge();
  writeTitle(sheet, "Mapping", "Browse source categories, inspect their members, and apply one validated mapping decision to one or many stable hierarchy items.", "N");

  sheet.getRange("A4:N4").setValues([["Weekly mapping health", "", "", "", "", "", "", "", "", "", "", "", "", ""]]);
  styleNavyHeader(sheet.getRange("A4:N4"));
  sheet.getRange("A5:F8").getFormat().getFill().setColor("#FFFFFF");
  sheet.getRange("A9:N9").getFormat().getFill().setColor("#E2F0D9");

  const categoryChoices = workspace.categories.map(category => category.categoryChoice);
  const defaultCategory = categoryChoices.indexOf(priorState.categoryChoice) >= 0 ? priorState.categoryChoice
    : categoryChoices.indexOf("Soft Drinks") >= 0 ? "Soft Drinks" : categoryChoices[0] || "";
  const views = ["All", "Unmapped", "Custom", "Identity Pending", "Excluded"];
  const defaultView = views.indexOf(priorState.view) >= 0 ? priorState.view : "All";
  const applyTo = priorState.applyTo === "Entire shown category" ? priorState.applyTo : "Selected members";

  sheet.getRange("A11:N11").setValues([["Category and bulk mapping", "", "", "", "", "", "", "", "", "", "", "", "", ""]]);
  styleSectionHeader(sheet.getRange("A11:N11"));
  sheet.getRange("A12").setValue("Show category"); sheet.getRange("B12").setValue(defaultCategory);
  sheet.getRange("D12").setValue("View"); sheet.getRange("E12").setValue(defaultView);
  sheet.getRange("G12").setValue("Active members"); sheet.getRange("H12").setFormula('=COUNT($AP$2:$AP$151)');
  sheet.getRange("I12").setValue("Shown Sales NOK"); sheet.getRange("J12").setValue(0);
  sheet.getRange("A13").setValue("Apply to"); sheet.getRange("B13").setValue(applyTo);
  sheet.getRange("D13").setValue("Selected items"); sheet.getRange("E13").setFormula('=COUNTIF(tblMappingMemberWorkspace[Select],"Yes")');
  sheet.getRange("G13").setValue("Selected facts"); sheet.getRange("H13").setFormula('=SUMIFS(tblMappingMemberWorkspace[Facts],tblMappingMemberWorkspace[Select],"Yes")');
  sheet.getRange("I13").setValue("Selected Sales NOK"); sheet.getRange("J13").setFormula('=SUMIFS(tblMappingMemberWorkspace[Sales NOK],tblMappingMemberWorkspace[Select],"Yes")');
  sheet.getRange("A12:A13").getFormat().getFont().setBold(true);
  sheet.getRange("D12:D13").getFormat().getFont().setBold(true);
  sheet.getRange("G12:G13").getFormat().getFont().setBold(true);
  sheet.getRange("I12:I13").getFormat().getFont().setBold(true);
  styleInput(sheet.getRange("B12")); styleInput(sheet.getRange("E12")); styleInput(sheet.getRange("B13"));
  sheet.getRange("A14:N14").merge();
  sheet.getRange("A14").setValue("Select one or more members below, choose an action, and apply once. Category and View lock while selected. Use Entire shown category only for an intentional category-wide rule.");
  sheet.getRange("A14:N14").getFormat().getFill().setColor("#EAF2FF"); sheet.getRange("A14:N14").getFormat().setWrapText(true);
  sheet.getRange("A15").setValue("Action"); sheet.getRange("B15").setValue(priorState.intent);
  sheet.getRange("D15").setValue("Assign to"); sheet.getRange("E15").setValue(priorState.targetGroupName);
  sheet.getRange("G15").setValue("Notes"); sheet.getRange("H15:J16").merge(); sheet.getRange("H15").setValue(priorState.notes);
  sheet.getRange("A15").getFormat().getFont().setBold(true);
  sheet.getRange("D15").getFormat().getFont().setBold(true);
  sheet.getRange("G15").getFormat().getFont().setBold(true);
  styleInput(sheet.getRange("B15")); styleInput(sheet.getRange("E15")); styleInput(sheet.getRange("H15:J16"));
  sheet.getRange("A17:J17").merge(); sheet.getRange("A17").setValue(message);
  sheet.getRange("A17:J17").getFormat().getFill().setColor("#F7F9FC"); sheet.getRange("A17:J17").getFormat().setWrapText(true);
  sheet.getRange("A18:J18").merge(); sheet.getRange("A18").setValue("Validation wiring pending.");

  sheet.getRange("A19:N19").setValues([["Category overview — navigation and impact", "", "", "", "", "", "", "", "", "", "", "", "", ""]]);
  styleSectionHeader(sheet.getRange("A19:N19"));
  const categoryHeaders = ["Main Category", "Current Reporting Group", "Mapping state", "Subcategories", "Products",
    "Historical Facts", "Historical Sales NOK", "Attention"];
  sheet.getRange("A20:H20").setValues([categoryHeaders]);
  const categoryValues = workspace.categories.map(category => [category.mainCategory, category.currentGroup, category.mappingState,
    category.subcategories, category.products, category.facts, category.sales, category.attention]);
  if (categoryValues.length) sheet.getRangeByIndexes(20, 0, categoryValues.length, categoryHeaders.length).setValues(categoryValues);
  const categoryEndRow = Math.max(20, 20 + categoryValues.length);
  const categoryTable = sheet.addTable(`A20:H${categoryEndRow}`, true); categoryTable.setName("tblMappingCategoryOverview");
  categoryTable.setPredefinedTableStyle("TableStyleMedium2");
  if (categoryValues.length) {
    sheet.getRange(`D21:F${categoryEndRow}`).setNumberFormat("#,##0");
    sheet.getRange(`G21:G${categoryEndRow}`).setNumberFormat("#,##0.00");
  }
  const categoryCatalogHeaders = ["Main Category", "Mapping state", "SourceSystemID", "MainNodeID", "CategoryChoice",
    "CurrentExplicitRuleID", "CurrentRuleAction", "CurrentTargetID"];
  auditSheet.getRange("A1:H1").setValues([categoryCatalogHeaders]);
  const categoryCatalogValues = workspace.categories.map(category => [category.mainCategory, category.mappingState,
    category.sourceSystemId, category.mainNodeId, category.categoryChoice, category.currentExplicitRuleId,
    category.currentRuleAction, category.currentTargetId]);
  if (categoryCatalogValues.length) auditSheet.getRangeByIndexes(1, 0, categoryCatalogValues.length, categoryCatalogHeaders.length).setValues(categoryCatalogValues);
  const categoryCatalogEndRow = Math.max(1, categoryCatalogValues.length + 1);
  const categoryCatalogTable = auditSheet.addTable(`A1:H${categoryCatalogEndRow}`, true); categoryCatalogTable.setName("tblMappingCategoryCatalog");
  categoryCatalogTable.setPredefinedTableStyle("TableStyleLight1");

  const memberSectionRow = categoryEndRow + 3;
  const memberHeaderRow = memberSectionRow + 1;
  const memberBodyStartRow = memberHeaderRow + 1;
  const memberBodyEndRow = memberHeaderRow + memberCapacity;
  sheet.getRange(`A${memberSectionRow}:N${memberSectionRow}`).setValues([["Members in shown category", "", "", "", "", "", "", "", "", "", "", "", "", ""]]);
  styleSectionHeader(sheet.getRange(`A${memberSectionRow}:N${memberSectionRow}`));
  const memberHeaders = ["Select", "Level", "Item", "Subcategory", "Sales Account", "Reporting Group", "Mapping state",
    "Facts", "Sales NOK", "Attention", "Historical Quantity", "SourceSystemID", "ScopeType", "NodeID", "ProductID",
    "ParentSubNodeID", "IdentityState", "CurrentExplicitRuleID", "CurrentRuleAction", "CurrentTargetID", "CategoryChoice"];
  sheet.getRange(`A${memberHeaderRow}:U${memberHeaderRow}`).setValues([memberHeaders]);
  sheet.getRange(`A${memberBodyStartRow}:U${memberBodyEndRow}`).clear(ExcelScript.ClearApplyTo.contents);
  const memberTable = sheet.addTable(`A${memberHeaderRow}:U${memberBodyEndRow}`, true); memberTable.setName("tblMappingMemberWorkspace");
  memberTable.setPredefinedTableStyle("TableStyleLight9"); memberTable.setShowBandedRows(false);
  sheet.getRange("J12").setFormula(`=IFERROR(SUMIFS($I$${memberBodyStartRow}:$I$${memberBodyEndRow},$B$${memberBodyStartRow}:$B$${memberBodyEndRow},"Product"),0)`);

  const catalogHeaders = memberHeaders.slice(1);
  auditSheet.getRange("J1:AC1").setValues([catalogHeaders]);
  const catalogValues = workspace.members.map(member => [member.level, member.item, member.subcategory, member.salesAccount,
    member.reportingGroup, member.mappingState, member.facts, member.sales, member.attention, member.quantity,
    member.sourceSystemId, member.scopeType, member.nodeId, member.productId, member.parentSubNodeId, member.identityState,
    member.currentExplicitRuleId, member.currentRuleAction, member.currentTargetId, member.categoryChoice]);
  if (catalogValues.length) auditSheet.getRangeByIndexes(1, 9, catalogValues.length, catalogHeaders.length).setValues(catalogValues);
  const catalogEndRow = Math.max(1, catalogValues.length + 1);
  const catalogTable = auditSheet.addTable(`J1:AC${catalogEndRow}`, true); catalogTable.setName("tblMappingMemberCatalog");
  catalogTable.setPredefinedTableStyle("TableStyleLight1");

  sheet.getRange("AP1").setValue("Catalog row index");
  const indexFormulas: string[][] = [];
  for (let rowIndex = 0; rowIndex < memberCapacity; rowIndex += 1) {
    indexFormulas.push([`=IFERROR(AGGREGATE(15,6,(ROW('_Mapping_Audit'!$J$2:$J$${catalogEndRow})-ROW('_Mapping_Audit'!$J$2)+1)/(('_Mapping_Audit'!$AC$2:$AC$${catalogEndRow}=$B$12)*IF($E$12="All",1,'_Mapping_Audit'!$O$2:$O$${catalogEndRow}=$E$12)),ROWS($AP$2:AP${rowIndex + 2})),"")`]);
  }
  sheet.getRange(`AP2:AP${memberCapacity + 1}`).setFormulas(indexFormulas);
  sheet.getRange("AQ1:AT1").setValues([["Category choices", "View choices", "Reporting Group choices", "View source"]]);
  sheet.getRange("AQ2").setFormula('=IF(COUNTIF(tblMappingMemberWorkspace[Select],"Yes")>0,$B$12,SORT(tblMappingCategoryCatalog[CategoryChoice]))');
  sheet.getRange("AR2").setFormula('=IF(COUNTIF(tblMappingMemberWorkspace[Select],"Yes")>0,$E$12,$AT$2:$AT$6)');
  sheet.getRange("AT2:AT6").setValues(views.map(view => [view]));
  const activeGroups = groups.filter(group => group.active === "Yes").sort((left, right) => left.sortOrder - right.sortOrder || left.name.localeCompare(right.name));
  if (activeGroups.length) sheet.getRange(`AS2:AS${activeGroups.length + 1}`).setValues(activeGroups.map(group => [group.name]));

  const formulas: string[][] = [];
  const catalogColumns = ["J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA", "AB", "AC"];
  const numericCatalogColumns: { [key: string]: boolean } = { "P": true, "Q": true, "S": true };
  for (let rowIndex = 0; rowIndex < memberCapacity; rowIndex += 1) {
    const formulaRow: string[] = [];
    for (let columnIndex = 0; columnIndex < catalogHeaders.length; columnIndex += 1) {
      const sourceColumn = catalogColumns[columnIndex];
      const indexedValue = `INDEX('_Mapping_Audit'!$${sourceColumn}$2:$${sourceColumn}$${catalogEndRow},$AP${rowIndex + 2})`;
      formulaRow.push(numericCatalogColumns[sourceColumn]
        ? `=IF($AP${rowIndex + 2}="","",IFERROR(${indexedValue},""))`
        : `=IF($AP${rowIndex + 2}="","",IFERROR(${indexedValue}&"",""))`);
    }
    formulas.push(formulaRow);
  }
  sheet.getRange(`B${memberBodyStartRow}:U${memberBodyEndRow}`).setFormulas(formulas);
  const selectedKeys: { [key: string]: boolean } = {};
  for (let index = 0; index < priorState.selectedTargets.length; index += 1) {
    const target = priorState.selectedTargets[index];
    selectedKeys[mappingNodeKey(target.sourceSystemId, target.scopeType, target.nodeId)] = true;
  }
  const shownMembers = workspace.members.filter(member => member.categoryChoice === defaultCategory &&
    (defaultView === "All" || member.mappingState === defaultView));
  const includeValues: string[][] = [];
  for (let rowIndex = 0; rowIndex < memberCapacity; rowIndex += 1) {
    const member = shownMembers[rowIndex];
    includeValues.push([member && selectedKeys[mappingNodeKey(member.sourceSystemId, member.scopeType, member.nodeId)] ? "Yes" : ""]);
  }
  sheet.getRange(`A${memberBodyStartRow}:A${memberBodyEndRow}`).setValues(includeValues);
  sheet.getRange(`H${memberBodyStartRow}:H${memberBodyEndRow}`).setNumberFormat("#,##0");
  sheet.getRange(`I${memberBodyStartRow}:I${memberBodyEndRow}`).setNumberFormat("#,##0.00");
  sheet.getRange(`K${memberBodyStartRow}:K${memberBodyEndRow}`).setNumberFormat("#,##0.000000");
  const selectFormat = sheet.getRange(`A${memberBodyStartRow}:A${memberBodyEndRow}`)
    .addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
  selectFormat.getCustom().getRule().setFormula(`=$B${memberBodyStartRow}<>""`);
  selectFormat.getCustom().getFormat().getFill().setColor("#FFF4CE");

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  sheet.getRange("K:U").setColumnHidden(true);
  sheet.getRange("K:AT").setColumnHidden(true);

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
  const mainAuditRow = 1;
  auditSheet.getRange(`AE${mainAuditRow}:AP${mainAuditRow}`).setValues([["SourceSystemID", "ScopeType", "NodeID", "SourceMainCategory", "Explicit Rule(s)", "Target Group ID(s)", "Subcategories", "Products", "Inheriting products", "Sales NOK (Active)", "Quantity (Active)", "Descendant exceptions"]]);
  const mainValues: (string|number)[][] = [];
  main.forEach((v,id) => mainValues.push([v.source,"SourceMainCategory",id,v.name,joinSet(v.ruleIds),joinSet(v.targetIds),v.subs.size,v.products.size,v.inherited,v.sales,v.qty,v.exceptions]));
  mainValues.sort((a,b)=>text(a[3]).localeCompare(text(b[3])));
  if (mainValues.length) auditSheet.getRangeByIndexes(mainAuditRow,30,mainValues.length,12).setValues(mainValues); const mt=auditSheet.addTable(`AE${mainAuditRow}:AP${Math.max(mainAuditRow,mainAuditRow+mainValues.length)}`,true); mt.setName("tblMappingMainNodes"); mt.setPredefinedTableStyle("TableStyleMedium2");
  if (mainValues.length) { auditSheet.getRange(`AK${mainAuditRow + 1}:AM${mainAuditRow + mainValues.length}`).setNumberFormat("#,##0"); auditSheet.getRange(`AN${mainAuditRow + 1}:AO${mainAuditRow + mainValues.length}`).setNumberFormat("#,##0.00"); auditSheet.getRange(`AP${mainAuditRow + 1}:AP${mainAuditRow + mainValues.length}`).setNumberFormat("#,##0"); }
  const subAuditRow = 1;
  auditSheet.getRange(`AR${subAuditRow}:BD${subAuditRow}`).setValues([["SourceSystemID", "ScopeType", "NodeID", "SourceMainCategory", "SourceSubCategory", "Explicit Rule(s)", "Target Group ID(s)", "Products", "Inheriting products", "Sales NOK (Active)", "Quantity (Active)", "Product exceptions", "Browse cue"]]);
  const subValues: (string|number)[][] = [];
  subs.forEach((v,id) => subValues.push([v.source,"SourceSubCategory",id,v.main,v.name,joinSet(v.ruleIds),joinSet(v.targetIds),v.products.size,v.inherited,v.sales,v.qty,v.exceptions,"Filter products below by NodeID"]));
  subValues.sort((a,b)=>text(a[3]).localeCompare(text(b[3]))||text(a[4]).localeCompare(text(b[4])));
  if(subValues.length) auditSheet.getRangeByIndexes(subAuditRow,43,subValues.length,13).setValues(subValues); const st=auditSheet.addTable(`AR${subAuditRow}:BD${Math.max(subAuditRow,subAuditRow+subValues.length)}`,true); st.setName("tblMappingSubcategoryNodes"); st.setPredefinedTableStyle("TableStyleMedium2");
  if (subValues.length) { auditSheet.getRange(`AY${subAuditRow + 1}:AZ${subAuditRow + subValues.length}`).setNumberFormat("#,##0"); auditSheet.getRange(`BA${subAuditRow + 1}:BB${subAuditRow + subValues.length}`).setNumberFormat("#,##0.00"); auditSheet.getRange(`BC${subAuditRow + 1}:BC${subAuditRow + subValues.length}`).setNumberFormat("#,##0"); }
  const productAuditRow = 1;
  auditSheet.getRange(`BF${productAuditRow}:BS${productAuditRow}`).setValues([["SourceSystemID","ScopeType","NodeID","SourceMainCategory","SourceSubCategory","Product","SalesAccount","Sales NOK (Active)","Quantity (Active)","Effective Reporting Group","Resolution Source","State","Status","Winning Rule"]]);
  const productValues = hierarchy.products.map(product => { const amount=context.get(product.productId)??{sales:0,quantity:0}; const r=resultByProduct.get(product.productId)!; return [product.sourceSystemId,"Product",product.productId,product.mainCategory,product.subCategory,product.productName,product.salesAccount,amount.sales,amount.quantity,r.effectiveGroupName,r.resolutionSource,r.resolutionState,r.resolutionStatus,r.winningRuleId]; });
  if(productValues.length) auditSheet.getRangeByIndexes(productAuditRow,57,productValues.length,14).setValues(productValues); const pt=auditSheet.addTable(`BF${productAuditRow}:BS${Math.max(productAuditRow,productAuditRow+productValues.length)}`,true); pt.setName("tblMappingProducts"); pt.setPredefinedTableStyle("TableStyleMedium2");
  if (productValues.length) {
    const productImpactFormulas: string[][] = [];
    for (let index = 0; index < productValues.length; index += 1) {
      const formulaRow = productAuditRow + index + 1;
      productImpactFormulas.push([
        `=SUMIFS(tblSalesFacts[SalesAmount],tblSalesFacts[ProductID],$BH${formulaRow},tblSalesFacts[PublicationState],"Active Finalized")`,
        `=SUMIFS(tblSalesFacts[Quantity],tblSalesFacts[ProductID],$BH${formulaRow},tblSalesFacts[PublicationState],"Active Finalized")`
      ]);
    }
    auditSheet.getRange(`BM${productAuditRow + 1}:BN${productAuditRow + productValues.length}`).setFormulas(productImpactFormulas);
    auditSheet.getRange(`BM${productAuditRow + 1}:BN${productAuditRow + productValues.length}`).setNumberFormat("#,##0.00");
  }
  if (mainValues.length) {
    const mainImpactFormulas: string[][] = [];
    for (let index = 0; index < mainValues.length; index += 1) {
      const formulaRow = mainAuditRow + index + 1;
      mainImpactFormulas.push([
        `=SUMIFS(tblMappingProducts[Sales NOK (Active)],tblMappingProducts[SourceMainCategory],$AH${formulaRow})`,
        `=SUMIFS(tblMappingProducts[Quantity (Active)],tblMappingProducts[SourceMainCategory],$AH${formulaRow})`
      ]);
    }
    auditSheet.getRange(`AN${mainAuditRow + 1}:AO${mainAuditRow + mainValues.length}`).setFormulas(mainImpactFormulas);
  }
  if (subValues.length) {
    const subImpactFormulas: string[][] = [];
    for (let index = 0; index < subValues.length; index += 1) {
      const formulaRow = subAuditRow + index + 1;
      subImpactFormulas.push([
        `=SUMIFS(tblMappingProducts[Sales NOK (Active)],tblMappingProducts[SourceMainCategory],$AU${formulaRow},tblMappingProducts[SourceSubCategory],$AV${formulaRow})`,
        `=SUMIFS(tblMappingProducts[Quantity (Active)],tblMappingProducts[SourceMainCategory],$AU${formulaRow},tblMappingProducts[SourceSubCategory],$AV${formulaRow})`
      ]);
    }
    auditSheet.getRange(`BA${subAuditRow + 1}:BB${subAuditRow + subValues.length}`).setFormulas(subImpactFormulas);
  }
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);
  sheet.getFreezePanes().freezeRows(2);
  setWidths(sheet,[105,130,175,115,145,125,105,105,105,170,70,70,70,70]);
  sheet.getRange("1:1").getFormat().setRowHeight(32); sheet.getRange("2:2").getFormat().setRowHeight(32);
  sheet.getRange("4:4").getFormat().setRowHeight(26); sheet.getRange("5:9").getFormat().setRowHeight(22);
  sheet.getRange("11:11").getFormat().setRowHeight(26); sheet.getRange("12:13").getFormat().setRowHeight(28);
  sheet.getRange("14:14").getFormat().setRowHeight(30); sheet.getRange("15:16").getFormat().setRowHeight(28);
  sheet.getRange("17:18").getFormat().setRowHeight(26); sheet.getRange("19:20").getFormat().setRowHeight(28);
  sheet.getRange(`21:${categoryEndRow}`).getFormat().setRowHeight(30);
  sheet.getRange(`${memberSectionRow}:${memberHeaderRow}`).getFormat().setRowHeight(28);
  sheet.getRange(`${memberBodyStartRow}:${memberBodyEndRow}`).getFormat().setRowHeight(24);
  sheet.getRange(`A11:J${memberBodyEndRow}`).getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRange(`A11:J${memberBodyEndRow}`).getFormat().setWrapText(false);
  sheet.getRange("A14:N14").getFormat().setWrapText(true);
  sheet.getRange("A17:J18").getFormat().setWrapText(true);
  sheet.getRange(`A20:H20`).getFormat().setWrapText(true);
  sheet.getRange(`A21:A${categoryEndRow}`).getFormat().setWrapText(true);
  sheet.getRange(`H21:H${categoryEndRow}`).getFormat().setWrapText(true);
  sheet.getRange(`A${memberHeaderRow}:J${memberHeaderRow}`).getFormat().setWrapText(true);
  sheet.getRange(`J${memberBodyStartRow}:J${memberBodyEndRow}`).getFormat().setWrapText(true);
  sheet.getRange("H12:H13").setNumberFormat("#,##0");
  sheet.getRange("J12:J13").setNumberFormat("#,##0.00");
  sheet.setShowGridlines(false);
}

function writeMappingLists(workbook: ExcelScript.Workbook, groups: ReportingGroup[]): void {
  const sheet = workbook.getWorksheet("_Mapping_Lists") ?? workbook.addWorksheet("_Mapping_Lists");
  sheet.getRange("A1:E1000").clear(ExcelScript.ClearApplyTo.all);
  sheet.getRange("A1:E1").setValues([["ActiveReportingGroupID", "ScopeType", "Action", "Status", "RuleAction"]]); styleHeader(sheet.getRange("A1:E1"));
  const active = groups.filter(g=>g.active==="Yes").sort((a,b)=>a.sortOrder-b.sortOrder||a.name.localeCompare(b.name));
  if(active.length) sheet.getRange(`A2:A${active.length+1}`).setValues(active.map(g=>[g.id]));
  sheet.getRange("B2:B4").setValues([["SourceMainCategory"],["SourceSubCategory"],["Product"]]);
  sheet.getRange("C2:C3").setValues([["Apply"],["Deactivate"]]); sheet.getRange("D2:D3").setValues([["Active"],["Inactive"]]);
  sheet.getRange("E2:E3").setValues([["Map"],["Exclude"]]);
}

function wireMappingValidation(workbook: ExcelScript.Workbook): string {
  const sheet = requiredSheet(workbook,"Mapping");
  const listsSheet = requiredSheet(workbook,"_Mapping_Lists");
  const groupCount=countPopulated(listsSheet.getRange("A2:A1000").getValues());
  const rulesSheet = requiredSheet(workbook,"Mapping Rules");
  const scopeSource = listsSheet.getRange("B2:B4");
  const activeGroupSource = listsSheet.getRange(`A2:A${Math.max(2,groupCount+1)}`);
  const statusSource = listsSheet.getRange("D2:D3");
  const ruleActionSource = listsSheet.getRange("E2:E3");
  const failures: string[] = [];
  applyRangeValidation(sheet.getRange("B12"), sheet.getRange("AQ2:AQ200"), "Mapping!B12 Show category", failures);
  applyRangeValidation(sheet.getRange("E12"), sheet.getRange("AR2:AR10"), "Mapping!E12 member view", failures);
  applyLiteralValidation(sheet.getRange("B13"), ["Selected members", "Entire shown category"], "Mapping!B13 apply scope", failures);
  applyLiteralValidation(sheet.getRange("B15"), ["Assign Reporting Group", "Leave Unmapped", "Remove custom mapping"], "Mapping!B15 action", failures);
  const groupNames = requiredTable(workbook, "tblReportingGroups").getRangeBetweenHeaderAndTotal().getValues()
    .filter(row => text(row[3]) === "Yes").map(row => text(row[1])).filter(value => value !== "");
  if (groupNames.some(value => value.indexOf(",") >= 0)) {
    applyRangeValidation(sheet.getRange("E15"), sheet.getRange(`AS2:AS${Math.max(2, groupNames.length + 1)}`), "Mapping!E15 Reporting Group", failures);
  } else {
    applyLiteralValidation(sheet.getRange("E15"), groupNames, "Mapping!E15 Reporting Group", failures);
  }
  const memberTable = requiredTable(workbook, "tblMappingMemberWorkspace");
  applyLiteralValidation(memberTable.getColumn("Select").getRangeBetweenHeaderAndTotal(), ["Yes"], "Mapping member selections", failures);
  applyListValidation(rulesSheet.getRange("C5:C1000"),scopeSource,"Mapping Rules!C5:C1000 ScopeType",failures);
  applyListValidation(rulesSheet.getRange("G5:G1000"),activeGroupSource,"Mapping Rules!G5:G1000 TargetReportingGroupID",failures);
  applyListValidation(rulesSheet.getRange("J5:J1000"),statusSource,"Mapping Rules!J5:J1000 Status",failures);
  applyListValidation(rulesSheet.getRange("M5:M1000"),ruleActionSource,"Mapping Rules!M5:M1000 RuleAction",failures);
  const message = failures.length
    ? `PUL-0301-013: ${failures.length} dropdown validation(s) unavailable; mapping refresh completed. ${failures.join(" | ")}`
    : "Mapping workspace dropdowns ready (10/10).";
  sheet.getRange("A18").setValue(message);
  const statusBand=sheet.getRange("A18:J18");
  statusBand.getFormat().getFill().setColor(failures.length?"#FCE8E6":"#E2F0D9");
  statusBand.getFormat().getFont().setBold(true);
  statusBand.getFormat().setWrapText(true);
  sheet.getRange("18:18").getFormat().setRowHeight(failures.length?72:28);
  failures.forEach(failure=>console.log(`PUL-0301-013 ${failure}`));
  return message;
}

function restoreWeeklyMappingHealthBlock(workbook: ExcelScript.Workbook, refreshRequired: boolean): void {
  const mapping = workbook.getWorksheet("Mapping");
  const products = workbook.getTable("tblWeeklyMappingAttention");
  const control = workbook.getTable("tblWeeklyMappingAttentionControl");
  if (!mapping || !products || !control) return;
  const rows = products.getRangeBetweenHeaderAndTotal().getValues();
  const headers = headerMap(products);
  const states = ["Mapped", "Unmapped", "Identity Pending", "Conflict", "Inactive Target"];
  const counts: { [key: string]: { products: number; facts: number; sales: number } } = {};
  for (let stateIndex = 0; stateIndex < states.length; stateIndex += 1) {
    counts[states[stateIndex]] = { products: 0, facts: 0, sales: 0 };
  }
  for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
    const state = text(rows[rowIndex][headers["Mapping State"]]);
    if (!counts[state]) continue;
    counts[state].products += 1;
    counts[state].facts += numberValue(rows[rowIndex][headers["Historical Facts"]]);
    counts[state].sales += numberValue(rows[rowIndex][headers["Historical Sales NOK"]]);
  }
  const controlValues = control.getRangeBetweenHeaderAndTotal().getValues();
  if (controlValues.length !== 1) return;
  const controlHeaders = headerMap(control);
  const health = text(controlValues[0][controlHeaders.HealthStatus]);
  const through = text(controlValues[0][controlHeaders.ThroughPeriodLabel]);
  const productValues: (string | number)[] = ["Products"];
  const factValues: (string | number)[] = ["Historical facts"];
  const salesValues: (string | number)[] = ["Historical Sales NOK"];
  for (let stateIndex = 0; stateIndex < states.length; stateIndex += 1) {
    productValues.push(counts[states[stateIndex]].products);
    factValues.push(counts[states[stateIndex]].facts);
    salesValues.push(counts[states[stateIndex]].sales);
  }
  mapping.getRange("A4:N9").unmerge();
  mapping.getRange("A4:N9").clear(ExcelScript.ClearApplyTo.formats);
  mapping.getRange("A4:N4").merge();
  mapping.getRange("A4").setValue(`Weekly mapping health — ${health}`);
  mapping.getRange("A5:F5").setValues([[
    "Metric", "Mapped", "Unmapped", "Identity Pending", "Conflict", "Inactive Target"
  ]]);
  mapping.getRange("A6:F6").setValues([productValues]);
  mapping.getRange("A7:F7").setValues([factValues]);
  mapping.getRange("A8:F8").setValues([salesValues]);
  mapping.getRange("A9:N9").merge();
  mapping.getRange("A9").setValue(refreshRequired
    ? "Performance refresh required"
    : `Performance classifications are up to date through ${through}`);
  mapping.getRange("A4:N4").getFormat().getFill().setColor("#17365D");
  mapping.getRange("A4:N4").getFormat().getFont().setColor("#FFFFFF");
  mapping.getRange("A4:N4").getFormat().getFont().setBold(true);
  mapping.getRange("A5:F5").getFormat().getFill().setColor("#D9EAF7");
  mapping.getRange("A5:F5").getFormat().getFont().setBold(true);
  mapping.getRange("A6:A8").getFormat().getFont().setBold(true);
  mapping.getRange("B6:F7").setNumberFormat("#,##0");
  mapping.getRange("B8:F8").setNumberFormat("#,##0.00");
  mapping.getRange("A9:N9").getFormat().getFill().setColor(refreshRequired ? "#FFF4CE" : "#E2F0D9");
  mapping.getRange("A9:N9").getFormat().getFont().setBold(true);
  mapping.getRange("A4:N9").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  mapping.getRange("A4:N9").getFormat().setWrapText(true);
  mapping.getRange("4:5").getFormat().setRowHeight(26);
  mapping.getRange("6:8").getFormat().setRowHeight(22);
  mapping.getRange("9:9").getFormat().setRowHeight(28);
}

function applyListValidation(target:ExcelScript.Range,sourceRange:ExcelScript.Range,label:string,failures:string[]):void{
  try{
    const validation=target.getDataValidation();validation.clear();
    const items:string[]=[];
    sourceRange.getValues().forEach(row=>{const item=text(row[0]);if(item)items.push(item);});
    if(!items.length)throw new Error("source list is empty");
    if(items.some(item=>item.indexOf(",")>=0))throw new Error("source item contains a comma");
    const source=items.join(",");
    validation.setRule({list:{inCellDropDown:true,source:source}});
  }catch(error){failures.push(`${label}: ${text(error)}`);}
}

function applyLiteralValidation(target:ExcelScript.Range,items:string[],label:string,failures:string[]):void{
  try{
    const normalized=items.map(item=>text(item)).filter(item=>item!=="");
    if(!normalized.length)throw new Error("source list is empty");
    if(normalized.some(item=>item.indexOf(",")>=0))throw new Error("source item contains a comma");
    const validation=target.getDataValidation();validation.clear();
    validation.setRule({list:{inCellDropDown:true,source:normalized.join(",")}});
  }catch(error){failures.push(`${label}: ${text(error)}`);}
}

function applyRangeValidation(target:ExcelScript.Range,sourceRange:ExcelScript.Range,label:string,failures:string[]):void{
  try{
    const validation=target.getDataValidation();validation.clear();
    validation.setRule({list:{inCellDropDown:true,source:sourceRange}});
  }catch(error){failures.push(`${label}: ${text(error)}`);}
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
  const sheet=workbook.getWorksheet(name)??workbook.addWorksheet(name);const used=sheet.getUsedRange();
  if(used){used.unmerge();used.clear(ExcelScript.ClearApplyTo.all);}return sheet;
}
function requiredSheet(workbook:ExcelScript.Workbook,name:string):ExcelScript.Worksheet{const sheet=workbook.getWorksheet(name);if(!sheet)throw new Error(`PUL-0301-009: Required sheet missing: ${name}`);return sheet;}
function requiredTable(workbook:ExcelScript.Workbook,name:string):ExcelScript.Table{const table=workbook.getTable(name);if(!table)throw new Error(`PUL-0301-010: Required table missing: ${name}`);return table;}
function headerMap(table:ExcelScript.Table):{[key:string]:number}{const headers=table.getHeaderRowRange().getValues()[0];const map:{[key:string]:number}={};for(let i=0;i<headers.length;i++)map[text(headers[i])]=i;return map;}
function writeTitle(sheet:ExcelScript.Worksheet,title:string,subtitle:string,endColumn:string):void{sheet.getRange(`A1:${endColumn}1`).getFormat().getFill().setColor("#172033");sheet.getRange(`A1:${endColumn}1`).getFormat().getFont().setColor("#FFFFFF");sheet.getRange(`A1:${endColumn}1`).getFormat().getFont().setBold(true);sheet.getRange(`A1:${endColumn}1`).getFormat().getFont().setSize(18);sheet.getRange("A1").setValue(title);sheet.getRange(`A2:${endColumn}2`).getFormat().getFill().setColor("#EAF2FF");sheet.getRange(`A2:${endColumn}2`).getFormat().setWrapText(true);sheet.getRange("A2").setValue(subtitle);sheet.getRange("1:1").getFormat().setRowHeight(32);sheet.getRange("2:2").getFormat().setRowHeight(30);}
function styleHeader(range:ExcelScript.Range):void{range.getFormat().getFill().setColor("#4F8CFF");range.getFormat().getFont().setColor("#FFFFFF");range.getFormat().getFont().setBold(true);range.getFormat().setWrapText(true);}
function styleNavyHeader(range:ExcelScript.Range):void{range.getFormat().getFill().setColor("#17365D");range.getFormat().getFont().setColor("#FFFFFF");range.getFormat().getFont().setBold(true);range.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);}
function styleSectionHeader(range:ExcelScript.Range):void{range.getFormat().getFill().setColor("#4F8CFF");range.getFormat().getFont().setColor("#FFFFFF");range.getFormat().getFont().setBold(true);range.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);}
function styleTableHeader(range:ExcelScript.Range):void{range.getFormat().getFill().setColor("#D9EAF7");range.getFormat().getFont().setBold(true);range.getFormat().setWrapText(true);range.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);}
function styleInput(range:ExcelScript.Range):void{range.getFormat().getFill().setColor("#FFF4CE");range.getFormat().getFont().setBold(true);range.getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);}
function setWidths(sheet:ExcelScript.Worksheet,widths:number[]):void{for(let i=0;i<widths.length;i++)sheet.getRangeByIndexes(0,i,1,1).getEntireColumn().getFormat().setColumnWidth(widths[i]);}
function countPopulated(values:(string|number|boolean)[][]):number{let count=0;for(const row of values)if(text(row[0]))count++;return count;}
function updateEnvironment(table:ExcelScript.Table,key:string,value:string,note:string):void{const rows=table.getRangeBetweenHeaderAndTotal().getValues();for(let i=0;i<rows.length;i++){if(text(rows[i][0])===key){table.getRangeBetweenHeaderAndTotal().getCell(i,1).setValue(value);table.getRangeBetweenHeaderAndTotal().getCell(i,2).setValue(note);return;}}table.addRow(-1,[key,value,note]);}
function appendBuildLog(table:ExcelScript.Table,row:(string|number|boolean)[]):void{table.addRow(-1,row);}
function nextId(table:ExcelScript.Table,prefix:string,digits:number):string{let max=0;for(const row of table.getRangeBetweenHeaderAndTotal().getValues()){const match=text(row[0]).match(new RegExp(`^${prefix}-(\\d+)$`));if(match)max=Math.max(max,Number(match[1]));}return `${prefix}-${String(max+1).padStart(digits,"0")}`;}
function excelNow():number{return Date.now()/86400000+25569;}
function excelToday():number{return Math.floor(excelNow());}
function numberValue(value:unknown):number{const n=Number(value);return Number.isFinite(n)?n:0;}
function text(value:unknown):string{return String(value??"").trim();}
