/**
 * Weekly-cache activation and consumer-freshness contract.
 *
 * This module operates on version-manifest records only. It does not mutate
 * cache rows, Performance, Reports, imports, facts, or workbook surfaces.
 */

export const WEEKLY_CACHE_ACTIVE_STATUS = "Active";
export const WEEKLY_CACHE_CANDIDATE_STATUS = "Candidate";
export const WEEKLY_CACHE_INACTIVE_STATE = "Not Active";
export const WEEKLY_CACHE_ACCEPTED_QA = "PASS";

/**
 * Validate one exact materialized candidate and return the only allowed
 * authority update. Exact reruns are idempotent.
 */
export function planWeeklyCacheActivation({
  versionManifests,
  expected,
  materialized,
  current,
}) {
  const versions = Array.isArray(versionManifests) ? versionManifests : [];
  const targetRows = versions.filter(row => row.cacheVersion === expected.cacheVersion);
  if (targetRows.length !== 1) {
    fail("PUL-030A-001", `Expected exactly one ${expected.cacheVersion} version row; found ${targetRows.length}.`);
  }
  const authorityRows = versions.filter(hasAnyAuthorityMarker);
  if (authorityRows.some(row => row.cacheVersion !== expected.cacheVersion)) {
    fail("PUL-030A-002", `Another weekly cache version has Active authority state: ${authorityRows.map(row => row.cacheVersion).join(", ")}.`);
  }
  if (authorityRows.length > 1) {
    fail("PUL-030A-003", `Weekly cache authority has ${authorityRows.length} Active-marked rows.`);
  }

  const target = targetRows[0];
  assertExpectedCandidate(target, expected, materialized);
  const freshness = validateVersionContent(target, current, { requireActive: false });
  if (freshness.length) fail("PUL-030A-004", freshness.join(" "));

  const alreadyActive = target.cacheStatus === WEEKLY_CACHE_ACTIVE_STATUS &&
    target.activationState === WEEKLY_CACHE_ACTIVE_STATUS;
  if (alreadyActive) {
    return {
      action: "Already Active",
      cacheVersion: target.cacheVersion,
      updates: Object.freeze({}),
    };
  }
  if (target.cacheStatus !== WEEKLY_CACHE_CANDIDATE_STATUS ||
      target.activationState !== WEEKLY_CACHE_INACTIVE_STATE) {
    fail(
      "PUL-030A-005",
      `Cache ${target.cacheVersion} is ${target.cacheStatus} / ${target.activationState}; expected Candidate / Not Active.`,
    );
  }
  return {
    action: "Activate",
    cacheVersion: target.cacheVersion,
    updates: Object.freeze({
      cacheStatus: WEEKLY_CACHE_ACTIVE_STATUS,
      activationState: WEEKLY_CACHE_ACTIVE_STATUS,
    }),
  };
}

/**
 * Apply a previously validated plan to plain records. Only the two authority
 * fields may change. Workbook mutation is performed by the bounded Office
 * Script after repeating the same preflight against live content.
 */
export function applyWeeklyCacheActivationPlan(versionManifests, plan) {
  if (plan.action === "Already Active") return versionManifests.map(row => ({ ...row }));
  if (plan.action !== "Activate" ||
      plan.updates.cacheStatus !== WEEKLY_CACHE_ACTIVE_STATUS ||
      plan.updates.activationState !== WEEKLY_CACHE_ACTIVE_STATUS) {
    fail("PUL-030A-006", "Activation plan contains an unsupported authority transition.");
  }
  let changed = 0;
  const next = versionManifests.map(row => {
    if (row.cacheVersion !== plan.cacheVersion) return { ...row };
    changed += 1;
    return {
      ...row,
      cacheStatus: plan.updates.cacheStatus,
      activationState: plan.updates.activationState,
    };
  });
  if (changed !== 1) fail("PUL-030A-007", `Activation target matched ${changed} version rows.`);
  return next;
}

/**
 * Minimum guard for any future weekly-cache consumer. A consumer may proceed
 * only when exactly one validated Active version matches all current semantic
 * content fingerprints. MappingAsOfDate is intentionally not an input.
 */
export function validateActiveWeeklyCacheFreshness({
  versionManifests,
  current,
}) {
  const versions = Array.isArray(versionManifests) ? versionManifests : [];
  const activeRows = versions.filter(isAuthorityActive);
  if (activeRows.length !== 1) {
    return {
      status: "Unavailable",
      activeVersion: null,
      errors: [`Weekly cache requires exactly one Active version; found ${activeRows.length}.`],
    };
  }
  const active = activeRows[0];
  const errors = validateVersionContent(active, current, { requireActive: true });
  return {
    status: errors.length ? "Stale" : "Available",
    activeVersion: active.cacheVersion,
    errors,
  };
}

function assertExpectedCandidate(target, expected, materialized) {
  const expectedFields = [
    ["cacheFingerprint", "CacheFingerprint"],
    ["mappingContentFingerprint", "MappingContentFingerprint"],
    ["catalogContentFingerprint", "CatalogContentFingerprint"],
    ["identityPreflightFingerprint", "IdentityPreflightFingerprint"],
    ["performanceRestaurantScopeFingerprint", "PerformanceRestaurantScopeFingerprint"],
  ];
  for (const [field, label] of expectedFields) {
    if (target[field] !== expected[field]) {
      fail("PUL-030A-008", `${label} ${target[field]} differs from accepted ${expected[field]}.`);
    }
  }
  if (target.validationStatus !== WEEKLY_CACHE_ACCEPTED_QA) {
    fail("PUL-030A-009", `Cache validation is ${target.validationStatus}; expected PASS.`);
  }
  const countFields = [
    ["periodRowCount", "periodRowCount"],
    ["scopeCacheRowCount", "scopeCacheRowCount"],
    ["denseRpgCacheRowCount", "denseRpgCacheRowCount"],
  ];
  for (const [manifestField, materializedField] of countFields) {
    const actual = Number(materialized[materializedField]);
    const accepted = Number(expected[manifestField]);
    if (Number(target[manifestField]) !== accepted || actual !== accepted) {
      fail(
        "PUL-030A-010",
        `${manifestField} is manifest ${target[manifestField]}, materialized ${actual}; expected ${accepted}.`,
      );
    }
  }
  if (materialized.cacheFingerprint !== expected.cacheFingerprint ||
      materialized.reconciliationStatus !== WEEKLY_CACHE_ACCEPTED_QA) {
    fail("PUL-030A-011", "Materialized cache fingerprint or reconciliation is not the accepted PASS state.");
  }
}

function validateVersionContent(version, current, { requireActive }) {
  const errors = [];
  if (requireActive && (version.cacheStatus !== WEEKLY_CACHE_ACTIVE_STATUS ||
      version.activationState !== WEEKLY_CACHE_ACTIVE_STATUS)) {
    errors.push(`Cache ${version.cacheVersion} is not Active / Active.`);
  }
  if (version.validationStatus !== WEEKLY_CACHE_ACCEPTED_QA) {
    errors.push(`Cache ${version.cacheVersion} validation is ${version.validationStatus}; expected PASS.`);
  }
  compareFingerprint(errors, version, current, "mappingContentFingerprint", "MappingContentFingerprint");
  compareFingerprint(errors, version, current, "catalogContentFingerprint", "CatalogContentFingerprint");
  compareFingerprint(errors, version, current, "identityPreflightFingerprint", "IdentityPreflightFingerprint");
  compareFingerprint(
    errors,
    version,
    current,
    "performanceRestaurantScopeFingerprint",
    "PerformanceRestaurantScopeFingerprint",
  );
  return errors;
}

function compareFingerprint(errors, version, current, field, label) {
  if (version[field] !== current[field]) {
    errors.push(`${label} ${version[field]} differs from current ${current[field]}.`);
  }
}

function isAuthorityActive(row) {
  return row.cacheStatus === WEEKLY_CACHE_ACTIVE_STATUS &&
    row.activationState === WEEKLY_CACHE_ACTIVE_STATUS;
}

function hasAnyAuthorityMarker(row) {
  return row.cacheStatus === WEEKLY_CACHE_ACTIVE_STATUS ||
    row.activationState === WEEKLY_CACHE_ACTIVE_STATUS;
}

function fail(code, message) {
  throw new Error(`${code}: ${message}`);
}
