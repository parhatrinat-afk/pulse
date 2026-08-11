/**
 * Pure deterministic mapping resolver used by Phase 1 fixtures.
 * Workbook automation implements the same generic scope/node contract.
 */

export const CURRENT_HIERARCHY = [
  { scopeType: "SourceMainCategory", level: 1 },
  { scopeType: "SourceSubCategory", level: 2 },
  { scopeType: "Product", level: 3 },
];

export function intervalsOverlap(a, b) {
  const aStart = a.effectiveFrom ?? Number.NEGATIVE_INFINITY;
  const bStart = b.effectiveFrom ?? Number.NEGATIVE_INFINITY;
  const aEnd = a.effectiveTo ?? Number.POSITIVE_INFINITY;
  const bEnd = b.effectiveTo ?? Number.POSITIVE_INFINITY;
  return aStart <= bEnd && bStart <= aEnd;
}

export function findRuleConflicts(rules) {
  const conflicts = [];
  for (let i = 0; i < rules.length; i += 1) {
    const left = rules[i];
    if (left.status !== "Active") continue;
    for (let j = i + 1; j < rules.length; j += 1) {
      const right = rules[j];
      if (right.status !== "Active") continue;
      if (left.sourceSystemId !== right.sourceSystemId) continue;
      if (left.scopeType !== right.scopeType || left.nodeId !== right.nodeId) continue;
      if (intervalsOverlap(left, right)) {
        conflicts.push([left.mappingRuleId, right.mappingRuleId]);
      }
    }
  }
  return conflicts;
}

export function resolveProduct({ product, rules, groups, asOf, hierarchy = CURRENT_HIERARCHY }) {
  const groupById = new Map(groups.map(group => [group.reportingGroupId, group]));
  const nodes = {
    SourceMainCategory: product.mainNodeId,
    SourceSubCategory: product.subNodeId,
    Product: product.productId,
  };
  const applicableByScope = new Map();

  for (const level of hierarchy) {
    const candidates = rules.filter(rule =>
      rule.status === "Active" &&
      rule.sourceSystemId === product.sourceSystemId &&
      rule.scopeType === level.scopeType &&
      rule.nodeId === nodes[level.scopeType] &&
      (rule.effectiveFrom == null || rule.effectiveFrom <= asOf) &&
      (rule.effectiveTo == null || rule.effectiveTo >= asOf)
    );
    applicableByScope.set(level.scopeType, candidates);
  }

  for (const level of [...hierarchy].sort((a, b) => b.level - a.level)) {
    const candidates = applicableByScope.get(level.scopeType) ?? [];
    if (candidates.length > 1) {
      return {
        effectiveReportingGroupId: "",
        resolutionSource: level.scopeType,
        resolutionState: "Explicit conflict",
        resolutionStatus: "Conflict",
        ruleId: candidates.map(rule => rule.mappingRuleId).join(", "),
      };
    }
    if (candidates.length === 1) {
      const rule = candidates[0];
      const group = groupById.get(rule.targetReportingGroupId);
      if (!group || group.active !== "Yes") {
        return {
          effectiveReportingGroupId: rule.targetReportingGroupId,
          resolutionSource: level.scopeType,
          resolutionState: level.scopeType === "Product" ? "Explicit" : "Inherited",
          resolutionStatus: "Inactive Target",
          ruleId: rule.mappingRuleId,
        };
      }
      return {
        effectiveReportingGroupId: rule.targetReportingGroupId,
        resolutionSource: level.scopeType,
        resolutionState: level.scopeType === "Product" ? "Explicit" : "Inherited",
        resolutionStatus: "Mapped",
        ruleId: rule.mappingRuleId,
      };
    }
  }

  return {
    effectiveReportingGroupId: "",
    resolutionSource: "Unmapped",
    resolutionState: "Unmapped",
    resolutionStatus: "Unmapped",
    ruleId: "",
  };
}
