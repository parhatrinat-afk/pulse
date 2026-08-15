import { validateActiveWeeklyCacheFreshness } from "../imports/weekly-cache-activation.mjs";
import { compareCandidateCacheRanges } from "../imports/weekly-compact-cache.mjs";

export const WEEKLY_PERFORMANCE_CONTRACT_VERSION = "0.3.0-weekly-performance-v1";

/**
 * Validate one inclusive ISO-week range against the active cache manifest.
 * A partial range is never calculated silently.
 */
export function validateWeeklyPeriodSelection(periodManifest, selection) {
  const isoYear = Number(selection?.isoYear);
  const weekStart = parseWeek(selection?.weekStart);
  const weekEnd = parseWeek(selection?.weekEnd);
  const label = rangeLabel(isoYear, weekStart, weekEnd);
  if (!Number.isInteger(isoYear) || !Number.isInteger(weekStart) ||
      !Number.isInteger(weekEnd) || weekStart < 1 || weekEnd > 53) {
    return blocked("Invalid", `${label} — invalid year/week`, 0, 0);
  }
  if (weekStart > weekEnd) {
    return blocked("Invalid", `${label} — invalid (From week is after To week)`, 0, 0);
  }

  const expected = weekEnd - weekStart + 1;
  const weeks = new Set();
  for (const row of periodManifest ?? []) {
    if (Number(row.isoYear) !== isoYear) continue;
    const week = Number(row.isoWeek);
    if (week < weekStart || week > weekEnd) continue;
    if (weeks.has(week)) {
      return blocked("Invalid", `${label} — invalid (duplicate W${padWeek(week)})`, weeks.size, expected);
    }
    weeks.add(week);
  }
  if (!weeks.size) return blocked("Invalid", `${label} — unavailable`, 0, expected);
  if (weeks.size !== expected) {
    return blocked("Incomplete", `${label} — incomplete (${weeks.size}/${expected} weeks)`, weeks.size, expected);
  }
  return {
    status: "Valid",
    available: true,
    summary: label,
    isoYear,
    weekStart,
    weekEnd,
    availableWeeks: weeks.size,
    expectedWeeks: expected,
  };
}

/**
 * Compare two independently selected complete ranges over the single active,
 * fresh weekly cache. Existing Phase 2C metric math remains authoritative.
 */
export function buildWeeklyPerformanceComparison({
  cache,
  versionManifests,
  currentFreshness,
  currentSelection,
  comparisonSelection,
  reportingGroupIds,
  restaurantIds,
}) {
  const freshness = validateActiveWeeklyCacheFreshness({
    versionManifests,
    current: currentFreshness,
  });
  if (freshness.status !== "Available") {
    return { status: freshness.status, freshness, current: null, comparison: null, warning: "" };
  }
  const current = validateWeeklyPeriodSelection(cache.periodManifest, currentSelection);
  const comparison = validateWeeklyPeriodSelection(cache.periodManifest, comparisonSelection);
  if (!current.available || !comparison.available) {
    return { status: "Unavailable", freshness, current, comparison, warning: "" };
  }
  const result = compareCandidateCacheRanges({
    cache,
    currentRange: {
      isoYear: current.isoYear,
      weekStart: current.weekStart,
      weekEnd: current.weekEnd,
    },
    comparisonRange: {
      isoYear: comparison.isoYear,
      weekStart: comparison.weekStart,
      weekEnd: comparison.weekEnd,
    },
    reportingGroupIds,
    restaurantIds,
  });
  const same = current.isoYear === comparison.isoYear &&
    current.weekStart === comparison.weekStart && current.weekEnd === comparison.weekEnd;
  const warning = same
    ? "Same Current and Compare period."
    : current.expectedWeeks === comparison.expectedWeeks
      ? ""
      : `Different complete period lengths (${current.expectedWeeks} vs ${comparison.expectedWeeks} weeks) — comparison allowed.`;
  return { status: "Available", freshness, current, comparison, warning, result };
}

function parseWeek(value) {
  if (typeof value === "string" && /^W\d{2}$/i.test(value.trim())) return Number(value.trim().slice(1));
  return Number(value);
}

function rangeLabel(year, start, end) {
  const yearText = Number.isInteger(year) ? String(year) : "?";
  const startText = Number.isInteger(start) ? padWeek(start) : "??";
  const endText = Number.isInteger(end) ? padWeek(end) : "??";
  return `${yearText} W${startText}–W${endText}`;
}

function padWeek(value) {
  return String(value).padStart(2, "0");
}

function blocked(status, summary, availableWeeks, expectedWeeks) {
  return { status, available: false, summary, availableWeeks, expectedWeeks };
}
