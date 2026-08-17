import {
  validateActiveWeeklyCacheFreshness,
} from "./weekly-cache-activation.mjs";
import {
  fingerprintWeeklyCacheRows,
} from "./weekly-compact-cache.mjs";

export const WEEKLY_INTAKE_PUBLICATION_RUNTIME_VERSION =
  "0.3.0-weekly-intake-publication-runtime-v1";
export const WEEKLY_INTAKE_PUBLISHED_STATUS = "Published";
export const WEEKLY_CACHE_ROLLBACK_STATUS = "Rollback";

/**
 * Validate the repository Candidate and return the only allowed publication
 * authority transition. No cache rows are changed by this pure planner.
 */
export function planWeeklyIntakePublishedActivation({
  publicationResult,
  versionManifests,
  currentFreshness,
  processedAt = "",
}) {
  const candidate = publicationResult?.candidate;
  if (publicationResult?.outcome !== "New" ||
      !publicationResult.candidatePrepared || !candidate) {
    fail("PUL-030WPR-001", "Published activation requires an accepted New Candidate.");
  }
  const priorFreshness = validateActiveWeeklyCacheFreshness({
    versionManifests,
    current: currentFreshness,
  });
  if (priorFreshness.status !== "Available") {
    fail("PUL-030WPR-002", priorFreshness.errors.join(" "));
  }
  if (priorFreshness.activeVersion !== publicationResult.activeCacheVersion) {
    fail(
      "PUL-030WPR-003",
      `Publisher prior ${publicationResult.activeCacheVersion} differs from Active ${priorFreshness.activeVersion}.`,
    );
  }
  const manifest = candidate.versionManifest;
  if (manifest.cacheStatus !== "Candidate" || manifest.activationState !== "Not Active" ||
      manifest.validationStatus !== "PASS" || candidate.validation?.status !== "PASS") {
    fail("PUL-030WPR-004", "Candidate is not Candidate / Not Active / PASS.");
  }
  for (const field of [
    "mappingContentFingerprint",
    "catalogContentFingerprint",
    "performanceRestaurantScopeFingerprint",
  ]) {
    if (manifest[field] !== currentFreshness[field]) {
      fail("PUL-030WPR-005", `${field} ${manifest[field]} differs from current ${currentFreshness[field]}.`);
    }
  }
  const counts = {
    period: candidate.periodManifest.length,
    scope: candidate.scopeCacheRows.length,
    rpg: candidate.weeklyRpgCacheRows.length,
  };
  if (counts.period !== Number(manifest.periodRowCount) ||
      counts.scope !== Number(manifest.scopeCacheRowCount) ||
      counts.rpg !== Number(manifest.denseRpgCacheRowCount)) {
    fail("PUL-030WPR-006", "Candidate row counts differ from its version manifest.");
  }
  const fingerprint = fingerprintWeeklyCacheRows({
    cacheVersion: manifest.cacheVersion,
    periodManifest: candidate.periodManifest,
    scopeCacheRows: candidate.scopeCacheRows,
    weeklyRpgCacheRows: candidate.weeklyRpgCacheRows,
  });
  if (fingerprint !== manifest.cacheFingerprint) {
    fail(
      "PUL-030WPR-007",
      `Candidate fingerprint ${fingerprint} differs from manifest ${manifest.cacheFingerprint}.`,
    );
  }
  const priorRows = (versionManifests ?? []).filter(row =>
    row.cacheStatus === "Active" && row.activationState === "Active");
  if (priorRows.length !== 1 || priorRows[0].cacheVersion !== priorFreshness.activeVersion) {
    fail("PUL-030WPR-008", "Publication requires exactly one prior Active / Active version.");
  }
  if (manifest.cacheVersion === priorRows[0].cacheVersion) {
    fail("PUL-030WPR-009", "Candidate CacheVersion cannot equal the prior Active version.");
  }

  const priorVersion = {
    ...priorRows[0],
    cacheStatus: WEEKLY_CACHE_ROLLBACK_STATUS,
    activationState: "Not Active",
  };
  const activeVersion = {
    ...manifest,
    cacheStatus: "Active",
    activationState: "Active",
  };
  const source = publicationResult.ledgerEntry ?? {};
  const period = String(source.sourcePeriodKey ?? "");
  const ledgerEntry = {
    ...source,
    intakeEventId: stableId("WINT-", [
      period,
      source.sourceSemanticFingerprint,
      WEEKLY_INTAKE_PUBLISHED_STATUS,
      manifest.cacheVersion,
    ]),
    identityPreflightFingerprint: manifest.identityPreflightFingerprint,
    intakeStatus: WEEKLY_INTAKE_PUBLISHED_STATUS,
    statusMessage: `Published ${period} in ${manifest.cacheVersion}; prior ${priorRows[0].cacheVersion} retained for rollback.`,
    processedAt: String(processedAt || source.processedAt || ""),
    priorCacheVersion: priorRows[0].cacheVersion,
    resultingCacheVersion: manifest.cacheVersion,
    supersededCacheVersion: priorRows[0].cacheVersion,
  };
  return {
    runtimeVersion: WEEKLY_INTAKE_PUBLICATION_RUNTIME_VERSION,
    action: "Publish",
    priorCacheVersion: priorRows[0].cacheVersion,
    resultingCacheVersion: manifest.cacheVersion,
    candidateFingerprint: manifest.cacheFingerprint,
    counts,
    finalVersionManifests: [priorVersion, activeVersion],
    identityControlFingerprint: manifest.identityPreflightFingerprint,
    ledgerEntry,
    result: {
      status: WEEKLY_INTAKE_PUBLISHED_STATUS,
      period,
      cacheChanged: true,
      priorCacheVersion: priorRows[0].cacheVersion,
      resultingCacheVersion: manifest.cacheVersion,
      archiveReady: true,
      message: ledgerEntry.statusMessage,
    },
  };
}

/** Apply only the two final authority states to plain manifests. */
export function applyWeeklyIntakePublishedActivation(versionManifests, plan) {
  if (plan?.action !== "Publish" || plan.finalVersionManifests?.length !== 2) {
    fail("PUL-030WPR-010", "Publication activation plan is invalid.");
  }
  const final = plan.finalVersionManifests.map(row => ({ ...row }));
  const active = final.filter(row => row.cacheStatus === "Active" && row.activationState === "Active");
  const rollback = final.filter(row => row.cacheStatus === WEEKLY_CACHE_ROLLBACK_STATUS &&
    row.activationState === "Not Active");
  if (active.length !== 1 || rollback.length !== 1 ||
      active[0].cacheVersion !== plan.resultingCacheVersion ||
      rollback[0].cacheVersion !== plan.priorCacheVersion) {
    fail("PUL-030WPR-011", "Publication would not leave one Active and one rollback version.");
  }
  return final;
}

function stableId(prefix, values) {
  return hashStrings([record("ID", values)], prefix);
}

function record(kind, values) {
  return `${kind}|${values.map(value => {
    const normalized = value === null || value === undefined ? "" : String(value);
    return `${normalized.length}:${normalized}`;
  }).join("|")}`;
}

function hashStrings(values, prefix) {
  let left = 0; let right = 0;
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

function fail(code, message) {
  throw new Error(`${code}: ${message}`);
}
