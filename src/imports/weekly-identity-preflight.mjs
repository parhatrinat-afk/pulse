import { resolveProduct } from "../mapping/hierarchical-resolver.mjs";
import { buildWeeklyCorpusManifest } from "./weekly-sales-parser.mjs";

export const WEEKLY_IDENTITY_PREFLIGHT_VERSION = "0.3.0-weekly-identity-preflight-v1";
export const IDENTITY_PENDING = "Identity Pending";
export const TEST_DEPARTMENT_SOURCE_NAMES = Object.freeze([
  "Test Department (Not for User)",
  "Test Department (Not for Users)",
]);

const MAPPING_STATES = Object.freeze([
  "Mapped",
  "Unmapped",
  "Conflict",
  "Inactive Target",
  IDENTITY_PENDING,
]);

/**
 * Resolve parsed weekly source strings against current Pulse catalogs, propose
 * stable candidates for new exact identities, and expose ambiguity/review
 * evidence. This function does not mutate catalogs, facts, mappings, or source
 * rows and does not build the weekly analytical cache.
 */
export function buildWeeklyIdentityPreflight({ parsedReports, catalogs }) {
  const sourceSystemId = requiredText(catalogs?.sourceSystemId, "catalog SourceSystemID");
  const reports = requireReports(parsedReports, sourceSystemId);
  const corpus = buildWeeklyCorpusManifest(reports);
  if (corpus.status !== "PASS") {
    throw new Error("PUL-030I-107: Weekly corpus manifest must pass before identity preflight.");
  }
  const rows = reports.flatMap(report => report.rows);
  const catalog = validateCatalogs(catalogs, sourceSystemId);
  const observed = observeSourceRows(rows, sourceSystemId);

  const restaurantPlan = planRestaurants(observed.restaurants, catalog.restaurants, sourceSystemId);
  const classificationPlan = planClassifications(
    observed.classifications,
    catalog.classifications,
    sourceSystemId,
  );
  const productPlan = planProducts(
    observed.products,
    catalog.products,
    classificationPlan.byKey,
    sourceSystemId,
  );

  const allRestaurants = [...catalog.restaurants, ...restaurantPlan.candidates];
  const allClassifications = [...catalog.classifications, ...classificationPlan.candidates];
  const allProducts = [...catalog.products, ...productPlan.candidates];
  const restaurantByKey = uniqueIndex(allRestaurants, row => restaurantKey(row.sourceSystemId, row.sourceRestaurantName));
  const classificationByKey = uniqueIndex(allClassifications, row => row.sourceClassificationKey);
  const productByKey = uniqueIndex(allProducts, row => row.productKey);
  const classificationById = uniqueIndex(allClassifications, row => row.sourceClassificationId);

  const mappingByProduct = materializeProductMapping({
    products: allProducts,
    classificationById,
    rules: catalog.mappingRules,
    groups: catalog.reportingGroups,
    asOf: catalog.catalogAsOfExcelSerial,
  });
  const hierarchyReview = buildHierarchyReview({
    observedProducts: observed.products,
    currentProducts: catalog.products,
    currentClassificationKeys: new Set(
      catalog.classifications.map(row => row.sourceClassificationKey),
    ),
    productByKey,
    classificationByKey,
    mappingByProduct,
    rules: catalog.mappingRules,
    groups: catalog.reportingGroups,
    asOf: catalog.catalogAsOfExcelSerial,
  });

  const coverage = Object.fromEntries(MAPPING_STATES.map(state => [state, emptyMetric()]));
  const identityStates = { Stable: emptyMetric(), [IDENTITY_PENDING]: emptyMetric() };
  const rowAssignments = [];
  for (const row of rows) {
    const sourceProductKey = productKey(sourceSystemId, row.item, row.salesAccount);
    const sourceClassificationKey = classificationKey(sourceSystemId, row.mainCategory, row.subCategory);
    const restaurantResolution = restaurantByKey.get(restaurantKey(sourceSystemId, row.restaurant));
    const productResolution = productByKey.get(sourceProductKey);
    const classificationResolution = classificationByKey.get(sourceClassificationKey);
    const productMapping = productResolution?.value
      ? mappingByProduct.get(productResolution.value.productId)
      : undefined;
    const pendingReasons = [];
    if (!restaurantResolution || restaurantResolution.collision) pendingReasons.push("Restaurant identity collision");
    if (!productResolution || productResolution.collision) pendingReasons.push("Product identity collision");
    if (!classificationResolution || classificationResolution.collision) {
      pendingReasons.push("Source classification identity collision");
    }
    if (productMapping?.identityPendingReason) pendingReasons.push(productMapping.identityPendingReason);

    const identityState = pendingReasons.length ? IDENTITY_PENDING : "Stable";
    const mappingStatus = identityState === IDENTITY_PENDING
      ? IDENTITY_PENDING
      : productMapping.resolutionStatus;
    addMetric(identityStates[identityState], row);
    addMetric(coverage[mappingStatus], row);
    rowAssignments.push({
      sourceRowId: row.sourceRowId,
      sourcePeriodKey: row.sourcePeriodKey,
      restaurantId: restaurantResolution?.value?.restaurantId ?? "",
      productId: productResolution?.value?.productId ?? "",
      sourceClassificationId: classificationResolution?.value?.sourceClassificationId ?? "",
      identityState,
      identityPendingReason: pendingReasons.join("; "),
      mappingStatus,
      effectiveReportingGroupId: mappingStatus === "Mapped"
        ? productMapping.effectiveReportingGroupId
        : "",
    });
  }

  const sourceTotals = totals(rows);
  const identityTotals = sumMetrics(Object.values(identityStates));
  const coverageTotals = sumMetrics(Object.values(coverage));
  const proposedStableIds = [
    ...restaurantPlan.candidates.map(row => row.restaurantId),
    ...classificationPlan.candidates.map(row => row.sourceClassificationId),
    ...productPlan.candidates.map(row => row.productId),
  ];
  const duplicateProposedStableIds = duplicateValues(proposedStableIds);
  const duplicateProposedStableKeys = [
    ...duplicateValues(restaurantPlan.candidates.map(row =>
      restaurantKey(row.sourceSystemId, row.sourceRestaurantName))).map(key => `Restaurant:${key}`),
    ...duplicateValues(classificationPlan.candidates.map(row => row.sourceClassificationKey))
      .map(key => `Source classification:${key}`),
    ...duplicateValues(productPlan.candidates.map(row => row.productKey)).map(key => `Product:${key}`),
  ];
  const reconciliationErrors = [];
  reconcileMetric("identity-state", sourceTotals, identityTotals, reconciliationErrors);
  reconcileMetric("mapping-state", sourceTotals, coverageTotals, reconciliationErrors);
  if (rowAssignments.length !== rows.length) reconciliationErrors.push("A source row was not assigned exactly once.");
  if (duplicateProposedStableIds.length) {
    reconciliationErrors.push(`Duplicate proposed stable IDs: ${duplicateProposedStableIds.join(", ")}.`);
  }
  if (duplicateProposedStableKeys.length) {
    reconciliationErrors.push(`Duplicate proposed stable keys: ${duplicateProposedStableKeys.join(", ")}.`);
  }

  const identityPendingItems = [
    ...enrichCatalogCollisions(catalog.collisions, observed),
    ...buildCurrentProductHierarchyPending({
      observedProducts: observed.products,
      currentProducts: catalog.products,
      productByKey,
      mappingByProduct,
    }),
    ...productPlan.hierarchyPending.map(row => ({
      entityType: "Product hierarchy",
      stableId: row.productId,
      sourceKey: row.productKey,
      reason: row.hierarchyPendingReason,
      observedPaths: row.observedHierarchyPaths,
      impact: row.impact,
    })),
  ].sort(compareReviewItems);
  const fingerprints = fingerprintPreflight({
    corpusFingerprint: corpus.corpusFingerprint,
    catalogs: catalog,
    restaurantCandidates: restaurantPlan.candidates,
    classificationCandidates: classificationPlan.candidates,
    productCandidates: productPlan.candidates,
    identityPendingItems,
    hierarchyReview,
    coverage,
    sourceTotals,
  });

  return {
    contractVersion: WEEKLY_IDENTITY_PREFLIGHT_VERSION,
    status: reconciliationErrors.length ? "FAIL" : identityPendingItems.length || hierarchyReview.length
      ? "PASS WITH REVIEW"
      : "PASS",
    sourceSystemId,
    catalogAsOfDate: catalog.catalogAsOfDate,
    mappingFingerprint: catalog.mappingFingerprint,
    reportCount: reports.length,
    sourceTotals,
    knownIdentitiesReused: {
      restaurants: entitySummary(observed.restaurants, restaurantPlan.knownKeys),
      products: entitySummary(observed.products, productPlan.knownKeys),
      classifications: entitySummary(observed.classifications, classificationPlan.knownKeys),
    },
    newIdentityCandidates: {
      restaurants: restaurantPlan.candidates,
      products: productPlan.candidates,
      classifications: classificationPlan.candidates,
    },
    identityPendingItems,
    hierarchyReview,
    identityStateCoverage: finalizeCoverage(identityStates),
    mappingStateCoverage: finalizeCoverage(coverage),
    reconciliation: {
      status: reconciliationErrors.length ? "FAIL" : "PASS",
      sourceFactCount: sourceTotals.factCount,
      assignedFactCount: identityTotals.factCount,
      sourceSalesNok: sourceTotals.salesNok,
      assignedSalesNok: identityTotals.salesNok,
      sourceQuantity: sourceTotals.quantity,
      assignedQuantity: identityTotals.quantity,
      duplicateProposedStableIds,
      duplicateProposedStableKeys,
      errors: reconciliationErrors,
    },
    fingerprints,
    rowAssignments,
  };
}

export function summarizeWeeklyIdentityPreflight(result, candidateLimit = 20) {
  return {
    contractVersion: result.contractVersion,
    status: result.status,
    sourceSystemId: result.sourceSystemId,
    catalogAsOfDate: result.catalogAsOfDate,
    mappingFingerprint: result.mappingFingerprint,
    reportCount: result.reportCount,
    sourceTotals: result.sourceTotals,
    knownIdentitiesReused: result.knownIdentitiesReused,
    newIdentityCandidates: {
      restaurants: result.newIdentityCandidates.restaurants,
      productCount: result.newIdentityCandidates.products.length,
      productIdRange: idRange(result.newIdentityCandidates.products.map(row => row.productId)),
      productImpact: sumMetrics(result.newIdentityCandidates.products.map(row => row.impact)),
      topProductsBySales: [...result.newIdentityCandidates.products]
        .sort((left, right) => right.impact.salesNok - left.impact.salesNok || compareText(left.productKey, right.productKey))
        .slice(0, candidateLimit),
      classificationCount: result.newIdentityCandidates.classifications.length,
      classificationIdRange: idRange(
        result.newIdentityCandidates.classifications.map(row => row.sourceClassificationId),
      ),
      classifications: result.newIdentityCandidates.classifications,
    },
    identityPendingItems: result.identityPendingItems,
    hierarchyReview: result.hierarchyReview,
    identityStateCoverage: result.identityStateCoverage,
    mappingStateCoverage: result.mappingStateCoverage,
    reconciliation: result.reconciliation,
    fingerprints: result.fingerprints,
  };
}

function validateCatalogs(catalogs, sourceSystemId) {
  const fields = ["restaurants", "products", "classifications", "reportingGroups", "mappingRules"];
  for (const field of fields) {
    if (!Array.isArray(catalogs?.[field])) throw new Error(`PUL-030I-101: Catalog ${field} is missing.`);
  }
  const restaurants = catalogs.restaurants.map(row => ({ ...row }));
  const products = catalogs.products.map(row => ({ ...row }));
  const classifications = catalogs.classifications.map(row => ({ ...row }));
  const reportingGroups = catalogs.reportingGroups.map(row => ({ ...row }));
  const mappingRules = catalogs.mappingRules.map(row => ({ ...row }));
  for (const row of [...restaurants, ...products, ...classifications]) {
    if (row.sourceSystemId !== sourceSystemId) {
      throw new Error(`PUL-030I-102: Catalog row belongs to unexpected SourceSystemID ${row.sourceSystemId}.`);
    }
  }
  const collisions = [
    ...catalogCollisions("Restaurant", restaurants, row => restaurantKey(row.sourceSystemId, row.sourceRestaurantName)),
    ...catalogCollisions("Product", products, row => row.productKey),
    ...catalogCollisions("Source classification", classifications, row => row.sourceClassificationKey),
  ];
  return {
    restaurants,
    products,
    classifications,
    reportingGroups,
    mappingRules,
    collisions,
    catalogAsOfDate: requiredText(catalogs.catalogAsOfDate, "CatalogAsOfDate"),
    catalogAsOfExcelSerial: finiteNumber(catalogs.catalogAsOfExcelSerial, "CatalogAsOfExcelSerial"),
    mappingFingerprint: requiredText(catalogs.mappingFingerprint, "mapping fingerprint"),
  };
}

function requireReports(parsedReports, sourceSystemId) {
  if (!Array.isArray(parsedReports) || !parsedReports.length) {
    throw new Error("PUL-030I-100: At least one parsed weekly report is required.");
  }
  return parsedReports.map(report => {
    if (!Array.isArray(report?.rows) || report.manifest?.contentReconciliationStatus !== "PASS") {
      throw new Error("PUL-030I-103: Parsed weekly report is incomplete or unreconciled.");
    }
    if (report.scopeContract?.sourceSystemId !== sourceSystemId) {
      throw new Error(`PUL-030I-104: Parsed report SourceSystemID differs from ${sourceSystemId}.`);
    }
    return report;
  });
}

function observeSourceRows(rows, sourceSystemId) {
  const restaurants = new Map();
  const products = new Map();
  const classifications = new Map();
  for (const row of rows) {
    const rKey = restaurantKey(sourceSystemId, row.restaurant);
    const pKey = productKey(sourceSystemId, row.item, row.salesAccount);
    const cKey = classificationKey(sourceSystemId, row.mainCategory, row.subCategory);
    addObservation(restaurants, rKey, row, { sourceRestaurantName: row.restaurant });
    const productObservation = addObservation(products, pKey, row, {
      sourceProductName: row.item,
      salesAccount: row.salesAccount,
      hierarchyPaths: new Set(),
      pathImpacts: new Map(),
    });
    productObservation.hierarchyPaths.add(cKey);
    const pathImpact = productObservation.pathImpacts.get(cKey) ?? emptyMetric();
    addMetric(pathImpact, row);
    productObservation.pathImpacts.set(cKey, pathImpact);
    addObservation(classifications, cKey, row, {
      sourceMainCategory: row.mainCategory,
      sourceSubCategory: row.subCategory,
    });
  }
  return { restaurants, products, classifications };
}

function addObservation(map, key, row, initial) {
  const value = map.get(key) ?? { key, ...initial, impact: emptyMetric(), sourcePeriodKeys: new Set() };
  addMetric(value.impact, row);
  value.sourcePeriodKeys.add(row.sourcePeriodKey);
  map.set(key, value);
  return value;
}

function planRestaurants(observed, existing, sourceSystemId) {
  const index = uniqueIndex(existing, row => restaurantKey(row.sourceSystemId, row.sourceRestaurantName));
  const knownKeys = new Set();
  const newKeys = [];
  for (const key of observed.keys()) {
    const match = index.get(key);
    if (match?.value) knownKeys.add(key);
    else if (!match?.collision) newKeys.push(key);
  }
  const ids = allocateIds(existing.map(row => row.restaurantId), "RST", 4, newKeys);
  const candidates = newKeys.sort(compareText).map(key => {
    const value = observed.get(key);
    const restaurantId = ids.get(key);
    return {
      restaurantId,
      sourceSystemId,
      sourceRestaurantName: value.sourceRestaurantName,
      displayName: value.sourceRestaurantName,
      status: "Active",
      reportingEnabled: "No",
      testDepartment: TEST_DEPARTMENT_SOURCE_NAMES.includes(value.sourceRestaurantName),
      identityState: "Proposed exact identity",
      impact: finalizeMetric(value.impact),
      firstPeriodKey: first(value.sourcePeriodKeys),
      lastPeriodKey: last(value.sourcePeriodKeys),
    };
  });
  return { knownKeys, candidates };
}

function planClassifications(observed, existing, sourceSystemId) {
  const index = uniqueIndex(existing, row => row.sourceClassificationKey);
  const knownKeys = new Set();
  const newKeys = [];
  for (const key of observed.keys()) {
    const match = index.get(key);
    if (match?.value) knownKeys.add(key);
    else if (!match?.collision) newKeys.push(key);
  }
  const ids = allocateIds(existing.map(row => row.sourceClassificationId), "SCL", 5, newKeys);
  const candidates = newKeys.sort(compareText).map(key => {
    const value = observed.get(key);
    return {
      sourceClassificationId: ids.get(key),
      sourceSystemId,
      sourceMainCategory: value.sourceMainCategory,
      sourceSubCategory: value.sourceSubCategory,
      sourceClassificationKey: key,
      status: "Active",
      identityState: "Proposed exact identity",
      impact: finalizeMetric(value.impact),
      firstPeriodKey: first(value.sourcePeriodKeys),
      lastPeriodKey: last(value.sourcePeriodKeys),
    };
  });
  const byKey = uniqueIndex([...existing, ...candidates], row => row.sourceClassificationKey);
  return { knownKeys, candidates, byKey };
}

function planProducts(observed, existing, classificationByKey, sourceSystemId) {
  const index = uniqueIndex(existing, row => row.productKey);
  const knownKeys = new Set();
  const newKeys = [];
  for (const key of observed.keys()) {
    const match = index.get(key);
    if (match?.value) knownKeys.add(key);
    else if (!match?.collision) newKeys.push(key);
  }
  const ids = allocateIds(existing.map(row => row.productId), "PRD", 6, newKeys);
  const candidates = newKeys.sort(compareText).map(key => {
    const value = observed.get(key);
    const paths = [...value.hierarchyPaths].sort(compareText);
    const resolvedPaths = paths.map(path => classificationByKey.get(path));
    const hierarchyUnambiguous = paths.length === 1 && resolvedPaths[0]?.value;
    return {
      productId: ids.get(key),
      sourceSystemId,
      sourceProductName: value.sourceProductName,
      salesAccount: value.salesAccount,
      sourceClassificationId: hierarchyUnambiguous
        ? resolvedPaths[0].value.sourceClassificationId
        : "",
      productKey: key,
      productStatus: "Active",
      identityState: "Proposed exact identity",
      hierarchyStatus: hierarchyUnambiguous ? "Proposed exact hierarchy" : IDENTITY_PENDING,
      hierarchyPendingReason: hierarchyUnambiguous ? "" : paths.length > 1
        ? "New ProductKey has multiple observed hierarchy paths."
        : "New ProductKey hierarchy classification is ambiguous.",
      observedHierarchyPaths: paths,
      impact: finalizeMetric(value.impact),
      firstPeriodKey: first(value.sourcePeriodKeys),
      lastPeriodKey: last(value.sourcePeriodKeys),
    };
  });
  return {
    knownKeys,
    candidates,
    hierarchyPending: candidates.filter(row => row.hierarchyStatus === IDENTITY_PENDING),
  };
}

function materializeProductMapping({ products, classificationById, rules, groups, asOf }) {
  const result = new Map();
  for (const product of products) {
    if (!product.sourceClassificationId) {
      result.set(product.productId, {
        resolutionStatus: IDENTITY_PENDING,
        effectiveReportingGroupId: "",
        resolutionSource: "Identity",
        resolutionState: IDENTITY_PENDING,
        ruleId: "",
        identityPendingReason: "Product hierarchy identity is pending.",
      });
      continue;
    }
    const classification = classificationById.get(product.sourceClassificationId);
    if (!classification?.value || classification.collision) {
      result.set(product.productId, {
        resolutionStatus: IDENTITY_PENDING,
        effectiveReportingGroupId: "",
        resolutionSource: "Identity",
        resolutionState: IDENTITY_PENDING,
        ruleId: "",
        identityPendingReason: "Product hierarchy classification is missing or ambiguous.",
      });
      continue;
    }
    result.set(product.productId, resolveProduct({
      product: {
        productId: product.productId,
        sourceSystemId: product.sourceSystemId,
        mainNodeId: mainNodeId(product.sourceSystemId, classification.value.sourceMainCategory),
        subNodeId: classification.value.sourceClassificationId,
      },
      rules,
      groups,
      asOf,
    }));
  }
  return result;
}

function buildHierarchyReview({
  observedProducts,
  currentProducts,
  currentClassificationKeys,
  productByKey,
  classificationByKey,
  mappingByProduct,
  rules,
  groups,
  asOf,
}) {
  const currentIds = new Set(currentProducts.map(row => row.productId));
  const reviews = [];
  for (const [key, observation] of observedProducts) {
    const product = productByKey.get(key)?.value;
    if (!product || !currentIds.has(product.productId)) continue;
    const current = mappingByProduct.get(product.productId);
    const alternatePaths = [];
    for (const path of [...observation.hierarchyPaths].sort(compareText)) {
      // The nine accepted authority-divergence cases are paths that already
      // existed in Pulse's catalog at the checkpoint. Newly observed source
      // classifications remain candidate lineage evidence; they must not be
      // treated as alternate mapping authority before the candidate catalog is
      // accepted.
      if (!currentClassificationKeys.has(path)) continue;
      const classification = classificationByKey.get(path);
      if (!classification?.value || classification.collision) continue;
      const alternate = resolveProduct({
        product: {
          productId: product.productId,
          sourceSystemId: product.sourceSystemId,
          mainNodeId: mainNodeId(product.sourceSystemId, classification.value.sourceMainCategory),
          subNodeId: classification.value.sourceClassificationId,
        },
        rules,
        groups,
        asOf,
      });
      if (mappingSignature(alternate) === mappingSignature(current)) continue;
      const pathImpact = observationByPathImpact(observation, path);
      alternatePaths.push({
        sourceClassificationId: classification.value.sourceClassificationId,
        sourceMainCategory: classification.value.sourceMainCategory,
        sourceSubCategory: classification.value.sourceSubCategory,
        alternateOutcome: compactMapping(alternate),
        impact: finalizeMetric(pathImpact),
      });
    }
    if (!alternatePaths.length) continue;
    const differentMappedGroups = alternatePaths.some(path =>
      current.resolutionStatus === "Mapped" &&
      path.alternateOutcome.resolutionStatus === "Mapped" &&
      current.effectiveReportingGroupId !== path.alternateOutcome.effectiveReportingGroupId
    );
    reviews.push({
      productId: product.productId,
      sourceProductName: product.sourceProductName,
      salesAccount: product.salesAccount,
      currentSourceClassificationId: product.sourceClassificationId,
      currentOutcome: compactMapping(current),
      reason: differentMappedGroups ? "Different mapped Reporting Groups" : "Current mapped outcome differs from observed path",
      prominent: product.productId === "PRD-000689",
      alternatePaths,
      impact: sumMetrics(alternatePaths.map(path => path.impact)),
    });
  }
  return reviews.sort((left, right) => Number(right.prominent) - Number(left.prominent) || compareText(left.productId, right.productId));
}

function observationByPathImpact(observation, path) {
  return observation.pathImpacts?.get(path) ?? emptyMetric();
}

function compactMapping(value) {
  return {
    effectiveReportingGroupId: value.effectiveReportingGroupId,
    resolutionStatus: value.resolutionStatus,
    resolutionSource: value.resolutionSource,
    resolutionState: value.resolutionState,
    ruleId: value.ruleId,
  };
}

function mappingSignature(value) {
  return `${value?.resolutionStatus ?? ""}|${value?.effectiveReportingGroupId ?? ""}`;
}

function entitySummary(observations, knownKeys) {
  return {
    distinctIdentityCount: knownKeys.size,
    impact: sumMetrics([...knownKeys].map(key => finalizeMetric(observations.get(key).impact))),
  };
}

function finalizeCoverage(value) {
  return Object.fromEntries(Object.entries(value).map(([key, metric]) => [key, finalizeMetric(metric)]));
}

function fingerprintPreflight({
  corpusFingerprint,
  catalogs,
  restaurantCandidates,
  classificationCandidates,
  productCandidates,
  identityPendingItems,
  hierarchyReview,
  coverage,
  sourceTotals,
}) {
  const catalogRecords = [
    record("CATALOG", [catalogs.catalogAsOfDate, catalogs.catalogAsOfExcelSerial, catalogs.mappingFingerprint]),
    ...catalogs.restaurants.map(row => record("RESTAURANT", [
      row.restaurantId,
      row.sourceSystemId,
      row.sourceRestaurantName,
      row.status,
      row.reportingEnabled,
    ])),
    ...catalogs.products.map(row => record("PRODUCT", [
      row.productId,
      row.productKey,
      row.sourceClassificationId,
      row.productStatus,
    ])),
    ...catalogs.classifications.map(row => record("CLASSIFICATION", [
      row.sourceClassificationId,
      row.sourceClassificationKey,
      row.status,
    ])),
    ...catalogs.reportingGroups.map(row => record("REPORTING_GROUP", [
      row.reportingGroupId,
      row.reportingGroupName,
      row.active,
      row.sortOrder,
    ])),
    ...catalogs.mappingRules.map(row => record("MAPPING_RULE", [
      row.mappingRuleId,
      row.sourceSystemId,
      row.scopeType,
      row.nodeId,
      row.targetReportingGroupId,
      row.effectiveFrom,
      row.effectiveTo,
      row.status,
      row.ruleAction,
    ])),
  ].sort(compareText);
  const candidateRecords = [
    ...restaurantCandidates.map(row => record("RESTAURANT", [
      row.restaurantId,
      row.sourceSystemId,
      row.sourceRestaurantName,
      row.status,
      row.reportingEnabled,
    ])),
    ...classificationCandidates.map(row => record("CLASSIFICATION", [
      row.sourceClassificationId,
      row.sourceClassificationKey,
      row.status,
    ])),
    ...productCandidates.map(row => record("PRODUCT", [
      row.productId,
      row.productKey,
      row.sourceClassificationId,
      row.productStatus,
      row.hierarchyStatus,
      ...row.observedHierarchyPaths,
    ])),
  ].sort(compareText);
  const reviewRecords = [
    ...identityPendingItems.map(row => record("PENDING", [
      row.entityType,
      row.stableId,
      row.sourceKey,
      row.reason,
      ...(row.candidateStableIds ?? []),
      ...(row.observedPaths ?? []),
      row.impact.factCount,
      row.impact.salesNok.toFixed(2),
      row.impact.quantity.toFixed(6),
    ])),
    ...hierarchyReview.flatMap(row => [
      record("HIERARCHY", [
        row.productId,
        row.currentSourceClassificationId,
        ...mappingFields(row.currentOutcome),
        row.reason,
      ]),
      ...row.alternatePaths.map(path => record("HIERARCHY_PATH", [
        row.productId,
        path.sourceClassificationId,
        path.sourceMainCategory,
        path.sourceSubCategory,
        ...mappingFields(path.alternateOutcome),
        path.impact.factCount,
        path.impact.salesNok.toFixed(2),
        path.impact.quantity.toFixed(6),
      ])),
    ]),
  ].sort(compareText);
  const coverageRecords = Object.keys(coverage).sort(compareText).map(state => {
    const metric = finalizeMetric(coverage[state]);
    return record("COVERAGE", [state, metric.factCount, metric.salesNok.toFixed(2), metric.quantity.toFixed(6)]);
  });
  return {
    sourceCorpusFingerprint: corpusFingerprint,
    catalogFingerprint: hashStrings(catalogRecords, "IDC-"),
    candidateFingerprint: hashStrings(candidateRecords, "IDN-"),
    reviewFingerprint: hashStrings(reviewRecords, "IDR-"),
    preflightFingerprint: hashStrings([
      corpusFingerprint,
      hashStrings(catalogRecords, "IDC-"),
      hashStrings(candidateRecords, "IDN-"),
      hashStrings(reviewRecords, "IDR-"),
      ...coverageRecords,
      record("SOURCE", [sourceTotals.factCount, sourceTotals.salesNok.toFixed(2), sourceTotals.quantity.toFixed(6)]),
    ], "IDP-"),
  };
}

function mappingFields(value) {
  return [
    value?.effectiveReportingGroupId ?? "",
    value?.resolutionStatus ?? "",
    value?.resolutionSource ?? "",
    value?.resolutionState ?? "",
    value?.ruleId ?? "",
  ];
}

function catalogCollisions(entityType, rows, keySelector) {
  const groups = groupBy(rows, keySelector);
  return [...groups.entries()].filter(([, values]) => values.length > 1).map(([sourceKey, values]) => ({
    entityType,
    stableId: "",
    sourceKey,
    reason: `Current catalog contains ${values.length} rows with the same stable key.`,
    candidateStableIds: values.map(row => row.restaurantId ?? row.productId ?? row.sourceClassificationId).sort(compareText),
    impact: emptyMetric(),
  }));
}

function enrichCatalogCollisions(collisions, observed) {
  const byEntity = {
    Restaurant: observed.restaurants,
    Product: observed.products,
    "Source classification": observed.classifications,
  };
  return collisions.map(row => ({
    ...row,
    impact: finalizeMetric(byEntity[row.entityType]?.get(row.sourceKey)?.impact ?? emptyMetric()),
  }));
}

function buildCurrentProductHierarchyPending({
  observedProducts,
  currentProducts,
  productByKey,
  mappingByProduct,
}) {
  const currentIds = new Set(currentProducts.map(row => row.productId));
  const result = [];
  for (const [key, observation] of observedProducts) {
    const resolution = productByKey.get(key);
    const product = resolution?.value;
    if (!product || resolution.collision || !currentIds.has(product.productId)) continue;
    const mapping = mappingByProduct.get(product.productId);
    if (!mapping?.identityPendingReason) continue;
    result.push({
      entityType: "Product hierarchy",
      stableId: product.productId,
      sourceKey: key,
      reason: mapping.identityPendingReason,
      observedPaths: [...observation.hierarchyPaths].sort(compareText),
      impact: finalizeMetric(observation.impact),
    });
  }
  return result;
}

function uniqueIndex(rows, keySelector) {
  const groups = groupBy(rows, keySelector);
  const result = new Map();
  for (const [key, values] of groups) {
    result.set(key, values.length === 1 ? { value: values[0], collision: false } : { value: undefined, collision: true });
  }
  return result;
}

function groupBy(values, keySelector) {
  const result = new Map();
  for (const value of values) {
    const key = keySelector(value);
    const group = result.get(key) ?? [];
    group.push(value);
    result.set(key, group);
  }
  return result;
}

function allocateIds(existingIds, prefix, digits, sourceKeys) {
  let maximum = 0;
  for (const id of existingIds) {
    const match = String(id).match(new RegExp(`^${prefix}-(\\d+)$`));
    if (match) maximum = Math.max(maximum, Number(match[1]));
  }
  const result = new Map();
  for (const key of [...sourceKeys].sort(compareText)) {
    maximum += 1;
    result.set(key, `${prefix}-${String(maximum).padStart(digits, "0")}`);
  }
  return result;
}

function restaurantKey(sourceSystemId, sourceRestaurantName) {
  return `${sourceSystemId} || ${sourceRestaurantName}`;
}

function productKey(sourceSystemId, item, salesAccount) {
  return `${sourceSystemId} || ${item} || ${salesAccount}`;
}

function classificationKey(sourceSystemId, mainCategory, subCategory) {
  return `${sourceSystemId} || ${mainCategory} || ${subCategory}`;
}

function mainNodeId(sourceSystemId, mainCategory) {
  return `${sourceSystemId} || Main || ${mainCategory}`;
}

function emptyMetric() {
  return { factCount: 0, salesNok: 0, quantity: 0 };
}

function addMetric(metric, row) {
  metric.factCount += Number(row.factCount ?? 1);
  metric.salesNok += Number(row.salesNok ?? row.salesAmount ?? 0);
  metric.quantity += Number(row.quantity ?? 0);
}

function totals(rows) {
  const result = emptyMetric();
  for (const row of rows) addMetric(result, row);
  return finalizeMetric(result);
}

function sumMetrics(metrics) {
  const result = emptyMetric();
  for (const metric of metrics) addMetric(result, metric);
  return finalizeMetric(result);
}

function finalizeMetric(metric) {
  return {
    factCount: Number(metric.factCount),
    salesNok: roundNumber(metric.salesNok, 2),
    quantity: roundNumber(metric.quantity, 6),
  };
}

function reconcileMetric(label, expected, actual, errors) {
  if (expected.factCount !== actual.factCount ||
      !almostEqual(expected.salesNok, actual.salesNok) ||
      !almostEqual(expected.quantity, actual.quantity)) {
    errors.push(`${label} coverage does not reconcile to source totals.`);
  }
}

function roundNumber(value, decimals) {
  const factor = 10 ** decimals;
  return Math.round((Number(value) + Number.EPSILON) * factor) / factor;
}

function almostEqual(left, right) {
  return Math.abs(Number(left) - Number(right)) <= 0.000001;
}

function requiredText(value, label) {
  const result = String(value ?? "");
  if (!result) throw new Error(`PUL-030I-105: ${label} is required.`);
  return result;
}

function finiteNumber(value, label) {
  const result = Number(value);
  if (!Number.isFinite(result)) throw new Error(`PUL-030I-106: ${label} must be numeric.`);
  return result;
}

function duplicateValues(values) {
  const counts = new Map();
  for (const value of values) counts.set(value, (counts.get(value) ?? 0) + 1);
  return [...counts.entries()].filter(([, count]) => count > 1).map(([value]) => value).sort(compareText);
}

function first(values) {
  return [...values].sort(compareText)[0] ?? "";
}

function last(values) {
  return [...values].sort(compareText).at(-1) ?? "";
}

function idRange(values) {
  const ordered = [...values].sort(compareText);
  return ordered.length ? `${ordered[0]}–${ordered.at(-1)}` : "";
}

function compareReviewItems(left, right) {
  return compareText(left.entityType, right.entityType) || compareText(left.stableId, right.stableId) ||
    compareText(left.sourceKey, right.sourceKey);
}

function compareText(left, right) {
  const a = String(left);
  const b = String(right);
  return a < b ? -1 : a > b ? 1 : 0;
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
