import { validateActiveWeeklyCacheFreshness } from "../imports/weekly-cache-activation.mjs";
import { compareCandidateCacheRanges } from "../imports/weekly-compact-cache.mjs";
import { planSelectionRows } from "./interactive-performance.mjs";

export const WEEKLY_PERFORMANCE_CONTRACT_VERSION = "0.3.0-weekly-performance-v1";

/**
 * Validate and deterministically order the current active Reporting Group
 * catalog. The weekly runtime is count-driven; nine groups are the accepted
 * business state, not a capacity contract.
 */
export function activeWeeklyPerformanceReportingGroups(reportingGroups) {
  if (!Array.isArray(reportingGroups)) {
    throw new Error("Reporting Group catalog is missing.");
  }
  const catalogIds = new Set();
  const activeNames = new Set();
  const activeSortOrders = new Set();
  const active = [];
  for (const row of reportingGroups) {
    const reportingGroupId = text(row.reportingGroupId ?? row.id);
    if (!reportingGroupId) throw new Error("Reporting Group catalog contains a blank ReportingGroupID.");
    if (catalogIds.has(reportingGroupId)) {
      throw new Error(`Reporting Group catalog repeats ReportingGroupID ${reportingGroupId}.`);
    }
    catalogIds.add(reportingGroupId);
    if (text(row.active ?? "Yes") !== "Yes") continue;
    const sortOrder = Number(row.sortOrder);
    if (!Number.isFinite(sortOrder)) {
      throw new Error(`Active Reporting Group ${reportingGroupId} has an invalid SortOrder.`);
    }
    if (activeSortOrders.has(sortOrder)) {
      throw new Error(`Active Reporting Groups repeat SortOrder ${sortOrder}.`);
    }
    const reportingGroupName = text(row.reportingGroupName ?? row.name) || reportingGroupId;
    if (activeNames.has(reportingGroupName)) {
      throw new Error(`Active Reporting Groups repeat business name ${reportingGroupName}.`);
    }
    activeNames.add(reportingGroupName);
    activeSortOrders.add(sortOrder);
    active.push({
      reportingGroupId,
      reportingGroupName,
      sortOrder,
    });
  }
  if (!active.length) throw new Error("At least one active Reporting Group is required.");
  active.sort((left, right) => left.sortOrder - right.sortOrder ||
    left.reportingGroupId.localeCompare(right.reportingGroupId));
  return active;
}

export function planWeeklyPerformanceRpgSelection({
  reportingGroups,
  priorRows = [],
  priorCatalogExists = false,
}) {
  const groups = activeWeeklyPerformanceReportingGroups(reportingGroups);
  return planSelectionRows({
    eligibleItems: groups.map(row => ({
      id: row.reportingGroupId,
      name: row.reportingGroupName,
    })),
    priorRows,
    priorCatalogExists,
  });
}

export function buildWeeklyPerformanceLayout({
  reportingGroups,
  restaurantCapacity,
}) {
  const groups = activeWeeklyPerformanceReportingGroups(reportingGroups);
  const restaurants = Number(restaurantCapacity);
  if (!Number.isInteger(restaurants) || restaurants < 1) {
    throw new Error("Weekly Performance requires a positive restaurant capacity.");
  }
  const groupCapacity = groups.length;
  const componentStartColumn = 39;
  const componentBlocks = [];
  for (let index = 0; index < 6; index += 1) {
    componentBlocks.push(componentStartColumn + index * (groupCapacity + 1));
  }
  const numericDisplayStartColumn = componentBlocks[componentBlocks.length - 1] + groupCapacity + 1;
  const totalComponentStartColumn = numericDisplayStartColumn + groupCapacity + 1;
  const totalDisplayColumn = totalComponentStartColumn + 4;
  const sortKeyColumn = totalDisplayColumn + 1;
  const sortedRestaurantIdColumn = sortKeyColumn + 1;
  const periodKeyStartColumn = sortedRestaurantIdColumn + 2;
  return {
    restaurantCapacity: restaurants,
    groupCapacity,
    componentBlocks,
    numericDisplayStartColumn,
    totalComponentStartColumn,
    totalDisplayColumn,
    sortKeyColumn,
    sortedRestaurantIdColumn,
    helperLastColumn: periodKeyStartColumn + 1,
    periodKeyStartColumn,
    matrixEndColumn: groupCapacity + 1,
    componentTotalRow: restaurants + 1,
  };
}

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

function text(value) {
  return String(value === null || value === undefined ? "" : value).trim();
}
