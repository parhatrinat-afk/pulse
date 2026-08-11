/**
 * Pulse Build 0.3.0 — frozen Lovable business-definition migration.
 *
 * Run after the merged Phase 2A checkpoint, then run Build_0_3_0_Phase1.ts
 * followed by Build_0_3_0_Phase2A.ts. This script only extends Mapping Rules.
 * It never chooses a target by name: every decision below uses an approved,
 * stable Pulse NodeID/ProductID and ReportingGroupID.
 */
function main(workbook: ExcelScript.Workbook): string {
  const rulesTable = requiredTable(workbook, "tblMappingRules");
  const groupsTable = requiredTable(workbook, "tblReportingGroups");
  const classificationsTable = requiredTable(workbook, "tblSourceClassifications");
  const productsTable = requiredTable(workbook, "tblProducts");
  const factsTable = requiredTable(workbook, "tblSalesFacts");
  const buildLog = requiredTable(workbook, "tblBuildLog");
  const factsBefore = snapshotFacts(factsTable);

  ensureRuleActionColumn(rulesTable);
  const decisions = migrationDecisions();
  validateDecisionContract(decisions);
  const displayByNode = readApprovedNodeDisplays(classificationsTable, productsTable);
  validateTargets(groupsTable, decisions);
  for (const decision of decisions) {
    if (!displayByNode.has(decision.nodeId)) {
      throw new Error(`PUL-030M-001: Approved stable node is absent from Pulse: ${decision.nodeId}. No rules were added.`);
    }
  }

  const asOf = excelToday();
  const plan = planMigration(rulesTable, decisions, asOf);
  if (!plan.reusedRuleIds.some(id => id === "MAP-000001")) {
    throw new Error("PUL-030M-002: Existing Add-ons rule MAP-000001 was not recognized for reuse. No rules were added.");
  }

  if (plan.additions.length) {
    const h = headerMap(rulesTable);
    const columnCount = rulesTable.getHeaderRowRange().getColumnCount();
    const createdAt = excelNow();
    const rows: (string | number | boolean)[][] = [];
    for (const addition of plan.additions) {
      const row: (string | number | boolean)[] = [];
      for (let i = 0; i < columnCount; i++) row.push("");
      row[h.MappingRuleID] = addition.mappingRuleId;
      row[h.SourceSystemID] = addition.sourceSystemId;
      row[h.ScopeType] = addition.scopeType;
      row[h.NodeID] = addition.nodeId;
      row[h.NodeDisplay] = displayByNode.get(addition.nodeId) || addition.nodeId;
      row[h.HierarchyLevel] = hierarchyLevel(addition.scopeType);
      row[h.TargetReportingGroupID] = addition.targetGroupId;
      row[h.EffectiveFrom] = asOf;
      row[h.EffectiveTo] = "";
      row[h.Status] = "Active";
      row[h.CreatedAt] = createdAt;
      row[h.Notes] = "Approved Lovable business-definition migration; stable Pulse ID decision.";
      row[h.RuleAction] = addition.ruleAction;
      rows.push(row);
    }
    rulesTable.addRows(-1, rows);
  }

  const factsAfter = snapshotFacts(factsTable);
  if (!sameFactSnapshot(factsBefore, factsAfter)) {
    throw new Error("PUL-030M-003: Source facts changed during mapping-rule migration.");
  }
  appendBuildLog(buildLog, [
    nextId(buildLog, "LOG", 6), excelNow(), "0.3.0-Lovable-Mapping",
    "Migrate approved Lovable reporting definitions", "Success", "Mapping Rules only",
    `${decisions.length} logical decisions; ${plan.reusedRuleIds.length} active rule(s) reused; ${plan.additions.length} rule(s) added. Run Phase 1 and Phase 2A next.`
  ]);
  return `Lovable mapping migration ready: ${decisions.length} decisions, ${plan.reusedRuleIds.length} reused, ${plan.additions.length} added. Run Build_0_3_0_Phase1.ts, then Build_0_3_0_Phase2A.ts.`;
}

type MigrationDecision = {
  sourceSystemId: string; scopeType: string; nodeId: string; ruleAction: string; targetGroupId: string;
};
type PlannedAddition = MigrationDecision & { mappingRuleId: string };
type FactSnapshot = {
  rows: number; sales: number; quantity: number; firstId: string; lastId: string;
};

function migrationDecisions(): MigrationDecision[] {
  const sourceSystemId = "SRC-TEST-SALES";
  const main: string[][] = [
    ["SRC-TEST-SALES || Main || *Bjørvika Special*", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Add-ons", "RPG-0001"],
    ["SRC-TEST-SALES || Main || Alcoholic Drinks", "RPG-0003"],
    ["SRC-TEST-SALES || Main || Bao Buns", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Bao Buns*", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Bao Buns.", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Beer*", "RPG-0005"],
    ["SRC-TEST-SALES || Main || COCKTAILS*", "RPG-0003"],
    ["SRC-TEST-SALES || Main || Campaign", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Campaign*", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Champagne", "RPG-0007"],
    ["SRC-TEST-SALES || Main || Cider", "RPG-0005"],
    ["SRC-TEST-SALES || Main || Cider*", "RPG-0005"],
    ["SRC-TEST-SALES || Main || Classic Maki", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Classic Sushi*", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Cocktails", "RPG-0003"],
    ["SRC-TEST-SALES || Main || Coctails", "RPG-0003"],
    ["SRC-TEST-SALES || Main || Coffee & Tea", "RPG-0004"],
    ["SRC-TEST-SALES || Main || Coffee & Tea*", "RPG-0004"],
    ["SRC-TEST-SALES || Main || Coffee.", "RPG-0004"],
    ["SRC-TEST-SALES || Main || Combinations", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Combinations*", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Deluxe Combination", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Deluxe Maki", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Deluxe Maki*", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Dessert", "RPG-0006"],
    ["SRC-TEST-SALES || Main || Dirty Maki", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Dirty Maki*", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Drink Mix", "RPG-0003"],
    ["SRC-TEST-SALES || Main || Drinks", "RPG-0002"],
    ["SRC-TEST-SALES || Main || Extra", "RPG-0001"],
    ["SRC-TEST-SALES || Main || Extras", "RPG-0001"],
    ["SRC-TEST-SALES || Main || Fast Snacks", "RPG-0001"],
    ["SRC-TEST-SALES || Main || Finish Line*", "RPG-0006"],
    ["SRC-TEST-SALES || Main || Hot Food", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Kampanje", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Lunch Offer", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Lunch Specials*", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Main Courses*", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Mains", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Mocktail", "RPG-0002"],
    ["SRC-TEST-SALES || Main || Mocktails", "RPG-0002"],
    ["SRC-TEST-SALES || Main || Nigiri", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Poké Bowls", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Poké Bowls.", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Quick Drinks", "RPG-0005"],
    ["SRC-TEST-SALES || Main || Sashimi", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Sauces", "RPG-0001"],
    ["SRC-TEST-SALES || Main || Sauser", "RPG-0001"],
    ["SRC-TEST-SALES || Main || Snacks", "RPG-0001"],
    ["SRC-TEST-SALES || Main || Soft Drinks", "RPG-0002"],
    ["SRC-TEST-SALES || Main || Softdrinks", "RPG-0002"],
    ["SRC-TEST-SALES || Main || Softdrinks*", "RPG-0002"],
    ["SRC-TEST-SALES || Main || Solheimsviken Special", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Special Wine", "RPG-0007"],
    ["SRC-TEST-SALES || Main || Spirits", "RPG-0003"],
    ["SRC-TEST-SALES || Main || Spirits*", "RPG-0003"],
    ["SRC-TEST-SALES || Main || Starters", "RPG-0008"],
    ["SRC-TEST-SALES || Main || Sushi", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Sushi + Hot Food", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Sushi Combinations", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Sushi For Many", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Sushi for One", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Sushi*", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Tilbud", "RPG-0009"],
    ["SRC-TEST-SALES || Main || Upsell", "RPG-0001"],
    ["SRC-TEST-SALES || Main || White Wine*", "RPG-0007"],
    ["SRC-TEST-SALES || Main || Wine", "RPG-0007"],
    ["SRC-TEST-SALES || Main || Wine & Bubbles*", "RPG-0007"],
    ["SRC-TEST-SALES || Main || Wine & Sake", "RPG-0007"]
  ];
  const subcategory: string[][] = [
    ["SCL-00017", "RPG-0005"], ["SCL-00057", "RPG-0003"]
  ];
  const product: string[][] = [
    ["PRD-000035", "RPG-0005"], ["PRD-000133", "RPG-0004"],
    ["PRD-000222", "RPG-0007"], ["PRD-000223", "RPG-0007"],
    ["PRD-000238", "RPG-0009"], ["PRD-000239", "RPG-0009"],
    ["PRD-000286", "RPG-0006"], ["PRD-000287", "RPG-0009"],
    ["PRD-000288", "RPG-0009"], ["PRD-000289", "RPG-0009"],
    ["PRD-000290", "RPG-0008"], ["PRD-000319", "RPG-0009"],
    ["PRD-000328", "RPG-0008"], ["PRD-000329", "RPG-0008"],
    ["PRD-000337", "RPG-0003"], ["PRD-000338", "RPG-0003"],
    ["PRD-000369", "RPG-0009"], ["PRD-000370", "RPG-0009"],
    ["PRD-000374", "RPG-0002"], ["PRD-000458", "RPG-0005"],
    ["PRD-000462", "RPG-0005"], ["PRD-000479", "RPG-0009"],
    ["PRD-000494", "RPG-0005"], ["PRD-000507", "RPG-0005"],
    ["PRD-000536", "RPG-0005"], ["PRD-000537", "RPG-0005"],
    ["PRD-000556", "RPG-0009"], ["PRD-000557", "RPG-0009"],
    ["PRD-000626", "RPG-0009"], ["PRD-000655", "RPG-0007"],
    ["PRD-000672", "RPG-0007"], ["PRD-000673", "RPG-0007"],
    ["PRD-000680", "RPG-0002"], ["PRD-000682", "RPG-0002"],
    ["PRD-000690", "RPG-0001"], ["PRD-000710", "RPG-0007"],
    ["PRD-000711", "RPG-0007"], ["PRD-000789", "RPG-0009"],
    ["PRD-000827", "RPG-0002"], ["PRD-000871", "RPG-0009"],
    ["PRD-000879", "RPG-0001"], ["PRD-000920", "RPG-0007"],
    ["PRD-000925", "RPG-0007"], ["PRD-000932", "RPG-0009"],
    ["PRD-000953", "RPG-0003"], ["PRD-000956", "RPG-0003"],
    ["PRD-001008", "RPG-0007"], ["PRD-001012", "RPG-0002"],
    ["PRD-001013", "RPG-0002"]
  ];
  const exclusions = [
    "PRD-000220", "PRD-000221", "PRD-000259", "PRD-000260",
    "PRD-000546", "PRD-000566", "PRD-000567", "PRD-000942"
  ];
  const result: MigrationDecision[] = [];
  for (const row of main) result.push(decision(sourceSystemId, "SourceMainCategory", row[0], "Map", row[1]));
  for (const row of subcategory) result.push(decision(sourceSystemId, "SourceSubCategory", row[0], "Map", row[1]));
  for (const row of product) result.push(decision(sourceSystemId, "Product", row[0], "Map", row[1]));
  for (const nodeId of exclusions) result.push(decision(sourceSystemId, "Product", nodeId, "Exclude", ""));
  return result;
}

function decision(sourceSystemId: string, scopeType: string, nodeId: string, ruleAction: string, targetGroupId: string): MigrationDecision {
  return { sourceSystemId, scopeType, nodeId, ruleAction, targetGroupId };
}

function validateDecisionContract(decisions: MigrationDecision[]): void {
  let mainCount = 0, subCount = 0, productMapCount = 0, exclusionCount = 0;
  const seen = new Set<string>();
  const seenNodes = new Set<string>();
  for (const value of decisions) {
    const key = semanticKey(value);
    if (seen.has(key)) throw new Error(`PUL-030M-004: Duplicate frozen migration decision ${key}.`);
    seen.add(key);
    const nodeKey = [value.sourceSystemId, value.scopeType, value.nodeId].join("|");
    if (seenNodes.has(nodeKey)) throw new Error(`PUL-030M-014: More than one frozen decision targets ${nodeKey}.`);
    seenNodes.add(nodeKey);
    if (value.ruleAction === "Exclude") {
      if (value.scopeType !== "Product" || !!value.targetGroupId) {
        throw new Error(`PUL-030M-005: Invalid Product exclusion ${value.nodeId}.`);
      }
      exclusionCount++;
    } else if (value.ruleAction === "Map" && !!value.targetGroupId) {
      if (value.scopeType === "SourceMainCategory") mainCount++;
      else if (value.scopeType === "SourceSubCategory") subCount++;
      else if (value.scopeType === "Product") productMapCount++;
      else throw new Error(`PUL-030M-006: Unsupported scope ${value.scopeType}.`);
    } else {
      throw new Error(`PUL-030M-007: Invalid Map decision ${value.nodeId}.`);
    }
  }
  if (decisions.length !== 129 || mainCount !== 70 || subCount !== 2 || productMapCount !== 49 || exclusionCount !== 8) {
    throw new Error(`PUL-030M-008: Frozen migration contract count mismatch: total ${decisions.length}, main ${mainCount}, sub ${subCount}, product ${productMapCount}, exclusions ${exclusionCount}.`);
  }
}

function ensureRuleActionColumn(table: ExcelScript.Table): void {
  if (headerMap(table).RuleAction === undefined) table.addColumn(-1, undefined, "RuleAction");
  const h = headerMap(table);
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  const range = table.getColumn("RuleAction").getRangeBetweenHeaderAndTotal();
  const values = range.getValues();
  for (let i = 0; i < values.length; i++) {
    if (text(rows[i][h.MappingRuleID]) && !text(values[i][0])) values[i][0] = "Map";
  }
  if (values.length) range.setValues(values);
}

function readApprovedNodeDisplays(
  classifications: ExcelScript.Table, products: ExcelScript.Table
): Map<string, string> {
  const result = new Map<string, string>();
  const ch = headerMap(classifications);
  const classificationValues = classifications.getRangeBetweenHeaderAndTotal().getValues();
  for (const row of classificationValues) {
    const sourceSystemId = text(row[ch.SourceSystemID]);
    if (sourceSystemId !== "SRC-TEST-SALES") continue;
    const main = text(row[ch.SourceMainCategory]);
    result.set(`${sourceSystemId} || Main || ${main}`, main || "(blank main category)");
    result.set(text(row[ch.SourceClassificationID]), text(row[ch.SourceSubCategory]) || "(blank subcategory)");
  }
  const ph = headerMap(products);
  const productValues = products.getRangeBetweenHeaderAndTotal().getValues();
  for (const row of productValues) {
    if (text(row[ph.SourceSystemID]) !== "SRC-TEST-SALES") continue;
    result.set(text(row[ph.ProductID]), text(row[ph.SourceProductName]));
  }
  return result;
}

function validateTargets(groups: ExcelScript.Table, decisions: MigrationDecision[]): void {
  const h = headerMap(groups);
  const active = new Set<string>();
  const values = groups.getRangeBetweenHeaderAndTotal().getValues();
  for (const row of values) if (text(row[h.Active]) === "Yes") active.add(text(row[h.ReportingGroupID]));
  for (const value of decisions) {
    if (value.ruleAction === "Map" && !active.has(value.targetGroupId)) {
      throw new Error(`PUL-030M-009: Approved target is missing or inactive: ${value.targetGroupId}. No rules were added.`);
    }
  }
}

function planMigration(
  table: ExcelScript.Table, decisions: MigrationDecision[], asOf: number
): { reusedRuleIds: string[]; additions: PlannedAddition[] } {
  const h = headerMap(table);
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  const applicable: { id: string; sourceSystemId: string; scopeType: string; nodeId: string; ruleAction: string; targetGroupId: string }[] = [];
  let maxRuleNumber = 0;
  for (const row of rows) {
    const id = text(row[h.MappingRuleID]);
    if (!id) continue;
    const match = id.match(/^MAP-(\d+)$/);
    if (match) maxRuleNumber = Math.max(maxRuleNumber, Number(match[1]));
    const effectiveFrom = numberValue(row[h.EffectiveFrom]);
    const effectiveTo = numberValue(row[h.EffectiveTo]);
    if (text(row[h.Status]) !== "Active" || (effectiveFrom && effectiveFrom > asOf) || (effectiveTo && effectiveTo < asOf)) continue;
    applicable.push({
      id,
      sourceSystemId: text(row[h.SourceSystemID]),
      scopeType: text(row[h.ScopeType]),
      nodeId: text(row[h.NodeID]),
      ruleAction: text(row[h.RuleAction]) || "Map",
      targetGroupId: text(row[h.TargetReportingGroupID])
    });
  }
  const reusedRuleIds: string[] = [];
  const additions: PlannedAddition[] = [];
  for (const requested of decisions) {
    const sameNode = applicable.filter(rule => rule.sourceSystemId === requested.sourceSystemId &&
      rule.scopeType === requested.scopeType && rule.nodeId === requested.nodeId);
    const identical = sameNode.filter(rule => semanticKey(rule) === semanticKey(requested));
    if (identical.length > 1) {
      throw new Error(`PUL-030M-010: Multiple active identical rules already exist for ${requested.nodeId}. No rules were added.`);
    }
    if (identical.length === 1) {
      reusedRuleIds.push(identical[0].id);
      continue;
    }
    if (sameNode.length) {
      throw new Error(`PUL-030M-011: Active rule ${sameNode.map(rule => rule.id).join(", ")} conflicts with approved decision ${semanticKey(requested)}. No rules were added.`);
    }
    maxRuleNumber++;
    additions.push({ ...requested, mappingRuleId: `MAP-${String(maxRuleNumber).padStart(6, "0")}` });
  }
  return { reusedRuleIds, additions };
}

function semanticKey(value: { sourceSystemId: string; scopeType: string; nodeId: string; ruleAction: string; targetGroupId: string }): string {
  return [value.sourceSystemId, value.scopeType, value.nodeId, value.ruleAction,
    value.ruleAction === "Exclude" ? "" : value.targetGroupId].join("|");
}

function hierarchyLevel(scopeType: string): number {
  if (scopeType === "SourceMainCategory") return 1;
  if (scopeType === "SourceSubCategory") return 2;
  if (scopeType === "Product") return 3;
  throw new Error(`PUL-030M-012: Unsupported hierarchy scope ${scopeType}.`);
}

function snapshotFacts(table: ExcelScript.Table): FactSnapshot {
  const h = headerMap(table);
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  let sales = 0, quantity = 0, firstId = "", lastId = "";
  for (const row of rows) {
    const id = text(row[h.SalesFactID]);
    if (!firstId && id) firstId = id;
    if (id) lastId = id;
    sales += numberValue(row[h.SalesAmount]);
    quantity += numberValue(row[h.Quantity]);
  }
  return { rows: rows.length, sales, quantity, firstId, lastId };
}

function sameFactSnapshot(left: FactSnapshot, right: FactSnapshot): boolean {
  return left.rows === right.rows && almostEqual(left.sales, right.sales) &&
    almostEqual(left.quantity, right.quantity) && left.firstId === right.firstId && left.lastId === right.lastId;
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-030M-013: Required table missing: ${name}.`);
  return table;
}
function headerMap(table: ExcelScript.Table): { [key: string]: number } {
  const headers = table.getHeaderRowRange().getValues()[0];
  const map: { [key: string]: number } = {};
  for (let i = 0; i < headers.length; i++) map[text(headers[i])] = i;
  return map;
}
function appendBuildLog(table: ExcelScript.Table, row: (string | number | boolean)[]): void { table.addRow(-1, row); }
function nextId(table: ExcelScript.Table, prefix: string, digits: number): string {
  let max = 0;
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  for (const row of values) {
    const match = text(row[0]).match(new RegExp(`^${prefix}-(\\d+)$`));
    if (match) max = Math.max(max, Number(match[1]));
  }
  return `${prefix}-${String(max + 1).padStart(digits, "0")}`;
}
function almostEqual(left: number, right: number): boolean { return Math.abs(left - right) <= Math.max(0.000001, Math.abs(right) * 1e-12); }
function excelNow(): number { return Date.now() / 86400000 + 25569; }
function excelToday(): number { return Math.floor(excelNow()); }
function numberValue(value: unknown): number { const result = Number(value); return Number.isFinite(result) ? result : 0; }
function text(value: unknown): string { return String(value ?? "").trim(); }
