export const WEEKLY_MAPPING_ATTENTION_VERSION =
  "0.3.0-weekly-mapping-attention-v1";

export const WEEKLY_MAPPING_ATTENTION_STATES = Object.freeze([
  "Mapped",
  "Unmapped",
  "Identity Pending",
  "Conflict",
  "Inactive Target",
]);

/**
 * Project the already-resolved weekly Product universe into one compact
 * administrative row per stable ProductID. Mapping authority remains the
 * existing Product hierarchy plus the accepted hierarchical resolver supplied
 * by the identity preflight; this function only aggregates evidence.
 */
export function buildWeeklyMappingAttentionProjection({
  products,
  classifications,
  reportingGroups,
  mappingByProduct,
  sourceRows,
  rowAssignments,
  existingCatalogProductIds,
  hierarchyReview = [],
}) {
  requireArray(products, "products");
  requireArray(classifications, "classifications");
  requireArray(reportingGroups, "reportingGroups");
  requireArray(sourceRows, "sourceRows");
  requireArray(rowAssignments, "rowAssignments");
  if (!(mappingByProduct instanceof Map)) {
    fail("PUL-030MA-001", "mappingByProduct must be the accepted resolver output Map.");
  }

  const existingIds = new Set(existingCatalogProductIds ?? []);
  const classificationById = uniqueIndex(
    classifications,
    row => row.sourceClassificationId,
    "SourceClassificationID",
  );
  const groupById = uniqueIndex(
    reportingGroups,
    row => row.reportingGroupId,
    "ReportingGroupID",
  );
  const productById = uniqueIndex(products, row => row.productId, "ProductID");
  const reviewByProductId = uniqueIndex(
    hierarchyReview,
    row => row.productId,
    "hierarchy-review ProductID",
  );
  const sourceRowById = uniqueIndex(sourceRows, row => row.sourceRowId, "SourceRowID");
  const impactByProductId = new Map();
  const assignmentStatusByProductId = new Map();
  const unassignedIdentityPendingTotals = emptyMetric();

  for (const assignment of rowAssignments) {
    const sourceRow = sourceRowById.get(assignment.sourceRowId);
    if (!sourceRow) {
      fail("PUL-030MA-002", `Assignment ${assignment.sourceRowId} has no source row.`);
    }
    const productId = String(assignment.productId ?? "");
    if (!productId) {
      if (String(assignment.identityState ?? "") !== "Identity Pending" ||
          String(assignment.mappingStatus ?? "") !== "Identity Pending") {
        fail("PUL-030MA-012", `Assignment ${assignment.sourceRowId} is unassigned without Identity Pending state.`);
      }
      addMetric(unassignedIdentityPendingTotals, sourceRow);
      continue;
    }
    if (!productById.has(productId)) {
      fail("PUL-030MA-003", `Assignment ProductID ${productId} is outside the current Product universe.`);
    }
    const current = impactByProductId.get(productId) ?? emptyMetric();
    addMetric(current, sourceRow);
    impactByProductId.set(productId, current);
    const statuses = assignmentStatusByProductId.get(productId) ?? new Set();
    statuses.add(requiredText(assignment.mappingStatus, "assignment mapping status"));
    assignmentStatusByProductId.set(productId, statuses);
  }
  if (sourceRowById.size !== rowAssignments.length) {
    fail(
      "PUL-030MA-004",
      `Source rows ${sourceRowById.size} and assignments ${rowAssignments.length} are not one-to-one.`,
    );
  }

  const rows = [];
  for (const product of products) {
    const productId = requiredText(product.productId, "ProductID");
    const resolution = mappingByProduct.get(productId);
    if (!resolution) {
      fail("PUL-030MA-005", `Product ${productId} has no accepted resolver output.`);
    }
    const mappingStatus = requiredText(resolution.resolutionStatus, "ResolutionStatus");
    if (!WEEKLY_MAPPING_ATTENTION_STATES.includes(mappingStatus)) {
      fail("PUL-030MA-006", `Product ${productId} has unsupported state ${mappingStatus}.`);
    }
    const assignmentStatuses = assignmentStatusByProductId.get(productId) ?? new Set();
    if (assignmentStatuses.size > 1 ||
        (assignmentStatuses.size === 1 && !assignmentStatuses.has(mappingStatus))) {
      fail(
        "PUL-030MA-007",
        `Product ${productId} row states ${[...assignmentStatuses].sort(compareText).join(", ")} differ from current ${mappingStatus}.`,
      );
    }
    const classificationId = String(product.sourceClassificationId ?? "");
    const classification = classificationId ? classificationById.get(classificationId) : undefined;
    const review = reviewByProductId.get(productId);
    const hierarchy = hierarchyAttention(product, review);
    const metric = finalizeMetric(impactByProductId.get(productId) ?? emptyMetric());
    const groupId = String(resolution.effectiveReportingGroupId ?? "");
    rows.push({
      productId,
      sourceSystemId: String(product.sourceSystemId ?? ""),
      item: String(product.sourceProductName ?? ""),
      sourceMainCategory: String(classification?.sourceMainCategory ?? ""),
      sourceSubCategory: String(classification?.sourceSubCategory ?? ""),
      salesAccount: String(product.salesAccount ?? ""),
      sourceClassificationId: classificationId,
      productKey: String(product.productKey ?? ""),
      identityOrigin: existingIds.has(productId) ? "Existing Product catalog" : "Weekly identity registry",
      effectiveReportingGroupId: groupId,
      effectiveReportingGroupName: String(groupById.get(groupId)?.reportingGroupName ?? ""),
      resolutionType: resolutionType(resolution),
      mappingStatus,
      winningRuleId: String(resolution.ruleId ?? ""),
      historicalFactCount: metric.factCount,
      historicalSalesNok: metric.salesNok,
      historicalQuantity: metric.quantity,
      hierarchyAttention: hierarchy.attention,
      hierarchyAlternatives: hierarchy.alternatives,
    });
  }
  rows.sort((left, right) => compareText(left.productId, right.productId));

  const duplicateProductIds = duplicates(rows.map(row => row.productId));
  const duplicateProductKeys = duplicates(rows.map(row => row.productKey).filter(Boolean));
  const sourceTotals = finalizeMetric(sumMetrics(sourceRows));
  const projectionTotals = finalizeMetric(sumMetrics(rows.map(row => ({
    factCount: row.historicalFactCount,
    salesNok: row.historicalSalesNok,
    quantity: row.historicalQuantity,
  }))));
  const unassignedTotals = finalizeMetric(unassignedIdentityPendingTotals);
  const accountedTotals = finalizeMetric(sumMetrics([projectionTotals, unassignedTotals]));
  const stateCoverage = Object.fromEntries(WEEKLY_MAPPING_ATTENTION_STATES.map(state => [
    state,
    { productCount: 0, ...emptyMetric() },
  ]));
  const resolutionTypeCounts = {};
  for (const row of rows) {
    const state = stateCoverage[row.mappingStatus];
    state.productCount += 1;
    state.factCount += row.historicalFactCount;
    state.salesNok += row.historicalSalesNok;
    state.quantity += row.historicalQuantity;
    resolutionTypeCounts[row.resolutionType] = (resolutionTypeCounts[row.resolutionType] ?? 0) + 1;
  }
  stateCoverage["Identity Pending"].factCount += unassignedTotals.factCount;
  stateCoverage["Identity Pending"].salesNok += unassignedTotals.salesNok;
  stateCoverage["Identity Pending"].quantity += unassignedTotals.quantity;
  for (const value of Object.values(stateCoverage)) {
    value.salesNok = round(value.salesNok, 2);
    value.quantity = round(value.quantity, 6);
  }

  const errors = [];
  if (duplicateProductIds.length) errors.push(`Duplicate ProductIDs: ${duplicateProductIds.join(", ")}.`);
  if (duplicateProductKeys.length) errors.push(`Duplicate ProductKeys: ${duplicateProductKeys.join(", ")}.`);
  reconcileMetric(sourceTotals, accountedTotals, errors);
  const existingProductCount = rows.filter(row => row.identityOrigin === "Existing Product catalog").length;
  const weeklyAddedProductCount = rows.length - existingProductCount;
  if (existingProductCount !== existingIds.size) {
    errors.push(`Existing Product count ${existingProductCount} differs from expected ${existingIds.size}.`);
  }
  const projectionFingerprint = fingerprintRows(rows);
  const attentionProducts = stateCoverage.Unmapped.productCount +
    stateCoverage["Identity Pending"].productCount;
  const criticalProducts = stateCoverage.Conflict.productCount +
    stateCoverage["Inactive Target"].productCount;

  return {
    contractVersion: WEEKLY_MAPPING_ATTENTION_VERSION,
    validationStatus: errors.length ? "FAIL" : "PASS",
    healthStatus: criticalProducts ? "Action required" : attentionProducts
      ? "Attention required"
      : "Mapping healthy",
    projectionFingerprint,
    totalProductCount: rows.length,
    existingProductCount,
    weeklyAddedProductCount,
    sourceTotals,
    projectionTotals,
    unassignedIdentityPendingTotals: unassignedTotals,
    accountedTotals,
    stateCoverage,
    resolutionTypeCounts,
    duplicateProductIds,
    duplicateProductKeys,
    errors,
    rows,
  };
}

/**
 * Build the read-only reverse membership view used by Mapping administration.
 * The weekly attention rows already contain the accepted resolver outcome; this
 * function only filters and aggregates that authority by active ReportingGroupID.
 */
export function buildReportingGroupMembershipView({
  rows,
  reportingGroups,
  memberCapacity = 400,
}) {
  requireArray(rows, "rows");
  requireArray(reportingGroups, "reportingGroups");
  const activeGroups = reportingGroups
    .filter(group => String(group.active ?? "Yes") === "Yes")
    .map(group => ({
      reportingGroupId: requiredText(group.reportingGroupId, "ReportingGroupID"),
      reportingGroupName: requiredText(group.reportingGroupName, "ReportingGroupName"),
      sortOrder: Number(group.sortOrder ?? 0),
    }))
    .sort((left, right) => left.sortOrder - right.sortOrder ||
      compareText(left.reportingGroupName, right.reportingGroupName));
  const groupById = uniqueIndex(
    activeGroups,
    group => group.reportingGroupId,
    "active ReportingGroupID",
  );
  uniqueIndex(
    activeGroups,
    group => group.reportingGroupName,
    "active ReportingGroupName",
  );
  const overviewById = new Map(activeGroups.map(group => [group.reportingGroupId, {
    ...group,
    productCount: 0,
    factCount: 0,
    salesNok: 0,
  }]));
  const productIds = new Set();
  const members = [];
  for (const row of rows) {
    if (String(row.mappingStatus ?? "") !== "Mapped") continue;
    const productId = requiredText(row.productId, "mapped ProductID");
    if (productIds.has(productId)) {
      fail("PUL-030RG-001", `Mapped ProductID ${productId} is duplicated.`);
    }
    productIds.add(productId);
    const reportingGroupId = requiredText(
      row.effectiveReportingGroupId,
      `mapped Product ${productId} ReportingGroupID`,
    );
    const group = groupById.get(reportingGroupId);
    if (!group) {
      fail("PUL-030RG-002", `Mapped Product ${productId} targets inactive or unknown ${reportingGroupId}.`);
    }
    const resolutionType = requiredText(row.resolutionType, `mapped Product ${productId} resolution`);
    if (!["Explicit Product", "Inherited Subcategory", "Inherited Main"].includes(resolutionType)) {
      fail("PUL-030RG-003", `Mapped Product ${productId} has unsupported resolution ${resolutionType}.`);
    }
    const member = {
      productId,
      product: String(row.item ?? ""),
      mainCategory: String(row.sourceMainCategory ?? ""),
      subcategory: String(row.sourceSubCategory ?? ""),
      salesAccount: String(row.salesAccount ?? ""),
      reportingGroupId,
      reportingGroupName: group.reportingGroupName,
      mappingState: resolutionType === "Explicit Product" ? "Custom" : "Inherited",
      factCount: Number(row.historicalFactCount ?? 0),
      salesNok: round(Number(row.historicalSalesNok ?? 0), 2),
    };
    members.push(member);
    const overview = overviewById.get(reportingGroupId);
    overview.productCount += 1;
    overview.factCount += member.factCount;
    overview.salesNok += member.salesNok;
  }
  members.sort((left, right) => {
    const groupCompare = activeGroups.findIndex(group => group.reportingGroupId === left.reportingGroupId) -
      activeGroups.findIndex(group => group.reportingGroupId === right.reportingGroupId);
    return groupCompare || compareText(left.product, right.product) ||
      compareText(left.salesAccount, right.salesAccount) || compareText(left.productId, right.productId);
  });
  const overview = activeGroups.map(group => {
    const value = overviewById.get(group.reportingGroupId);
    return { ...value, salesNok: round(value.salesNok, 2) };
  });
  const maximumGroupProducts = overview.reduce((maximum, row) =>
    Math.max(maximum, row.productCount), 0);
  if (maximumGroupProducts > memberCapacity) {
    fail(
      "PUL-030RG-004",
      `Largest Reporting Group has ${maximumGroupProducts} Products; bounded capacity is ${memberCapacity}.`,
    );
  }
  const totals = overview.reduce((result, row) => ({
    productCount: result.productCount + row.productCount,
    factCount: result.factCount + row.factCount,
    salesNok: result.salesNok + row.salesNok,
  }), { productCount: 0, factCount: 0, salesNok: 0 });
  totals.salesNok = round(totals.salesNok, 2);
  return { activeGroups, overview, members, totals, maximumGroupProducts };
}

function hierarchyAttention(product, review) {
  const observed = [...(product.observedHierarchyPaths ?? [])].sort(compareText);
  if (String(product.hierarchyStatus ?? "") === "Identity Pending") {
    return {
      attention: `Identity Pending — ${observed.length} observed hierarchy paths`,
      alternatives: observed.join(" | "),
    };
  }
  if (!review) return { attention: "", alternatives: "" };
  const alternatives = (review.alternatePaths ?? []).map(path => {
    const group = String(path.alternateOutcome?.effectiveReportingGroupId ?? "");
    const status = String(path.alternateOutcome?.resolutionStatus ?? "");
    const outcome = group || status || "Unmapped";
    return `${path.sourceMainCategory} › ${path.sourceSubCategory} → ${outcome}`;
  }).sort(compareText);
  return {
    attention: review.reason || "Alternate observed hierarchy",
    alternatives: alternatives.join(" | "),
  };
}

function resolutionType(resolution) {
  const status = String(resolution.resolutionStatus ?? "");
  const source = String(resolution.resolutionSource ?? "");
  const state = String(resolution.resolutionState ?? "");
  if (status === "Identity Pending") return "Identity Pending";
  if (status === "Conflict") return "Conflict";
  if (status === "Inactive Target") return "Inactive Target";
  if (status === "Unmapped" && state === "Explicit exclusion") return "Explicit exclusion";
  if (status === "Unmapped") return "Unmapped";
  if (source === "Product") return "Explicit Product";
  if (source === "SourceSubCategory") return "Inherited Subcategory";
  if (source === "SourceMainCategory") return "Inherited Main";
  return source || status;
}

function fingerprintRows(rows) {
  const records = [record("CONTRACT", [WEEKLY_MAPPING_ATTENTION_VERSION])];
  for (const row of rows) records.push(record("PRODUCT", [
    row.productId,
    row.sourceSystemId,
    row.item,
    row.sourceMainCategory,
    row.sourceSubCategory,
    row.salesAccount,
    row.sourceClassificationId,
    row.productKey,
    row.identityOrigin,
    row.effectiveReportingGroupId,
    row.effectiveReportingGroupName,
    row.resolutionType,
    row.mappingStatus,
    row.winningRuleId,
    row.historicalFactCount,
    row.historicalSalesNok.toFixed(2),
    row.historicalQuantity.toFixed(6),
    row.hierarchyAttention,
    row.hierarchyAlternatives,
  ]));
  return hashStrings(records.sort(compareText), "WMA-");
}

function uniqueIndex(rows, selector, label) {
  const result = new Map();
  for (const row of rows) {
    const key = String(selector(row) ?? "");
    if (!key) continue;
    if (result.has(key)) fail("PUL-030MA-009", `${label} ${key} is duplicated.`);
    result.set(key, row);
  }
  return result;
}

function reconcileMetric(source, projection, errors) {
  if (source.factCount !== projection.factCount) {
    errors.push(`Fact count ${projection.factCount} differs from source ${source.factCount}.`);
  }
  if (Math.abs(source.salesNok - projection.salesNok) > 0.005) {
    errors.push(`Sales NOK ${projection.salesNok} differs from source ${source.salesNok}.`);
  }
  if (Math.abs(source.quantity - projection.quantity) > 0.000005) {
    errors.push(`Quantity ${projection.quantity} differs from source ${source.quantity}.`);
  }
}

function sumMetrics(rows) {
  const result = emptyMetric();
  for (const row of rows) {
    result.factCount += Number(row.factCount ?? 1);
    result.salesNok += Number(row.salesNok ?? 0);
    result.quantity += Number(row.quantity ?? 0);
  }
  return result;
}

function addMetric(metric, row) {
  metric.factCount += 1;
  metric.salesNok += Number(row.salesNok ?? 0);
  metric.quantity += Number(row.quantity ?? 0);
}

function finalizeMetric(metric) {
  return {
    factCount: Number(metric.factCount ?? 0),
    salesNok: round(Number(metric.salesNok ?? 0), 2),
    quantity: round(Number(metric.quantity ?? 0), 6),
  };
}

function emptyMetric() {
  return { factCount: 0, salesNok: 0, quantity: 0 };
}

function duplicates(values) {
  const seen = new Set();
  const duplicate = new Set();
  for (const value of values) {
    if (seen.has(value)) duplicate.add(value);
    seen.add(value);
  }
  return [...duplicate].sort(compareText);
}

function record(kind, values) {
  return `${kind}|${values.map(value => {
    const normalized = value === null || value === undefined ? "" : String(value);
    return `${normalized.length}:${normalized}`;
  }).join("|")}`;
}

function hashStrings(values, prefix) {
  let left = 0;
  let right = 0;
  for (const item of values) {
    const value = `${item}\n`;
    for (let index = 0; index < value.length; index += 1) {
      const code = value.charCodeAt(index);
      left = (left * 131 + code) % 2147483647;
      right = (right * 137 + code) % 2147483629;
    }
  }
  return `${prefix}${left.toString(16).padStart(8, "0")}${right.toString(16).padStart(8, "0")}`;
}

function requiredText(value, label) {
  const result = String(value ?? "");
  if (!result) fail("PUL-030MA-010", `${label} is blank.`);
  return result;
}

function requireArray(value, label) {
  if (!Array.isArray(value)) fail("PUL-030MA-011", `${label} must be an array.`);
}

function round(value, digits) {
  const factor = 10 ** digits;
  return Math.round((value + Number.EPSILON) * factor) / factor;
}

function compareText(left, right) {
  return String(left).localeCompare(String(right));
}

function fail(code, message) {
  throw new Error(`${code}: ${message}`);
}
