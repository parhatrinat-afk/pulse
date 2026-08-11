/**
 * Frozen, human-approved Lovable -> Pulse mapping migration contract.
 *
 * Every decision below targets a stable Pulse NodeID/ProductID. Source labels
 * are deliberately absent: names are display metadata and never decide a
 * migration target at runtime.
 */

export const LOVABLE_MIGRATION_SOURCE_SYSTEM_ID = "SRC-TEST-SALES";
export const LOVABLE_MIGRATION_NOTE =
  "Approved Lovable business-definition migration; stable Pulse ID decision.";

const MAIN_RULES = [
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
  ["SRC-TEST-SALES || Main || Wine & Sake", "RPG-0007"],
];

const SUBCATEGORY_RULES = [
  ["SCL-00017", "RPG-0005"],
  ["SCL-00057", "RPG-0003"],
];

const PRODUCT_RULES = [
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
  ["PRD-001013", "RPG-0002"],
];

export const LOVABLE_PRODUCT_EXCLUSIONS = Object.freeze([
  "PRD-000220", "PRD-000221", "PRD-000259", "PRD-000260",
  "PRD-000546", "PRD-000566", "PRD-000567", "PRD-000942",
]);

export const LOVABLE_MIGRATION_DECISIONS = Object.freeze([
  ...MAIN_RULES.map(([nodeId, targetReportingGroupId]) => decision(
    "SourceMainCategory", nodeId, "Map", targetReportingGroupId,
  )),
  ...SUBCATEGORY_RULES.map(([nodeId, targetReportingGroupId]) => decision(
    "SourceSubCategory", nodeId, "Map", targetReportingGroupId,
  )),
  ...PRODUCT_RULES.map(([nodeId, targetReportingGroupId]) => decision(
    "Product", nodeId, "Map", targetReportingGroupId,
  )),
  ...LOVABLE_PRODUCT_EXCLUSIONS.map(nodeId => decision("Product", nodeId, "Exclude", "")),
]);

export function planLovableMigration({ existingRules, asOf, decisions = LOVABLE_MIGRATION_DECISIONS }) {
  const requestedNodes = new Set();
  for (const requested of decisions) {
    const nodeKey = [requested.sourceSystemId, requested.scopeType, requested.nodeId].join("|");
    if (requestedNodes.has(nodeKey)) {
      throw new Error(`More than one migration decision targets ${nodeKey}.`);
    }
    requestedNodes.add(nodeKey);
  }
  const applicable = existingRules.filter(rule =>
    normalizeStatus(rule.status) === "Active" &&
    (number(rule.effectiveFrom) === 0 || number(rule.effectiveFrom) <= asOf) &&
    (number(rule.effectiveTo) === 0 || number(rule.effectiveTo) >= asOf),
  );
  const reused = [];
  const additions = [];
  let nextNumber = maxRuleNumber(existingRules) + 1;

  for (const requested of decisions) {
    validateDecision(requested);
    const sameNode = applicable.filter(rule =>
      String(rule.sourceSystemId).trim() === requested.sourceSystemId &&
      String(rule.scopeType).trim() === requested.scopeType &&
      String(rule.nodeId).trim() === requested.nodeId,
    );
    const identical = sameNode.filter(rule => semanticKey(rule) === semanticKey(requested));
    if (identical.length > 1) {
      throw new Error(`Multiple active semantically identical rules already exist for ${requested.nodeId}.`);
    }
    if (identical.length === 1) {
      reused.push({ decision: requested, mappingRuleId: ruleId(identical[0]) });
      continue;
    }
    if (sameNode.length > 0) {
      throw new Error(
        `Active rule ${sameNode.map(ruleId).join(", ")} conflicts with approved decision ${semanticKey(requested)}.`,
      );
    }
    additions.push({
      ...requested,
      mappingRuleId: `MAP-${String(nextNumber).padStart(6, "0")}`,
      effectiveFrom: asOf,
      effectiveTo: 0,
      status: "Active",
      notes: LOVABLE_MIGRATION_NOTE,
    });
    nextNumber += 1;
  }

  return { reused, additions };
}

export function semanticKey(rule) {
  const action = normalizeRuleAction(rule);
  return [
    String(rule.sourceSystemId ?? "").trim(),
    String(rule.scopeType ?? "").trim(),
    String(rule.nodeId ?? "").trim(),
    action,
    action === "Exclude" ? "" : String(rule.targetReportingGroupId ?? rule.targetGroupId ?? "").trim(),
  ].join("|");
}

export function normalizeRuleAction(rule) {
  return String(rule.ruleAction ?? rule.action ?? "").trim() || "Map";
}

function decision(scopeType, nodeId, ruleAction, targetReportingGroupId) {
  return Object.freeze({
    sourceSystemId: LOVABLE_MIGRATION_SOURCE_SYSTEM_ID,
    scopeType,
    nodeId,
    ruleAction,
    targetReportingGroupId,
  });
}

function validateDecision(value) {
  if (value.ruleAction === "Exclude" && (value.scopeType !== "Product" || value.targetReportingGroupId)) {
    throw new Error(`Explicit exclusion must be Product-scoped with no Reporting Group target: ${value.nodeId}.`);
  }
  if (value.ruleAction === "Map" && !value.targetReportingGroupId) {
    throw new Error(`Map decision requires a Reporting Group target: ${value.nodeId}.`);
  }
}

function maxRuleNumber(rules) {
  return rules.reduce((maximum, rule) => {
    const match = ruleId(rule).match(/^MAP-(\d+)$/);
    return match ? Math.max(maximum, Number(match[1])) : maximum;
  }, 0);
}

function ruleId(rule) {
  return String(rule.mappingRuleId ?? rule.id ?? "").trim();
}

function normalizeStatus(value) {
  return String(value ?? "").trim();
}

function number(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}
