/**
 * Deterministic Build 0.3.0 Phase 2C interaction primitives.
 *
 * These functions consume the accepted Phase 2B Restaurant-scope result rows.
 * They do not read facts, resolve mappings, materialize new metric results, or
 * mutate configuration. The workbook formulas mirror this reference contract.
 */

export const PERFORMANCE_SORT_ORDERS = Object.freeze(["Highest first", "Lowest first"]);
export const PERFORMANCE_MATRIX_MODES = Object.freeze([
  "PP Change",
  "Current Share",
  "Comparison Share",
  "Current Sales NOK",
  "NOK Impact",
]);

export function planSelectionRows({
  eligibleItems,
  priorRows = [],
  priorCatalogExists = false,
}) {
  const eligibleIds = new Set();
  const priorById = new Map();

  for (const row of priorRows) {
    const id = text(row.id);
    if (!id) throw new Error("Prior selection contains a blank stable ID.");
    if (priorById.has(id)) throw new Error(`Prior selection repeats stable ID ${id}.`);
    priorById.set(id, normalizeInclude(row.include));
  }

  return eligibleItems.map(item => {
    const id = text(item.id);
    if (!id) throw new Error("Eligible selection item contains a blank stable ID.");
    if (eligibleIds.has(id)) throw new Error(`Eligible selection repeats stable ID ${id}.`);
    eligibleIds.add(id);
    return {
      id,
      name: text(item.name) || id,
      include: priorById.has(id)
        ? priorById.get(id)
        : priorCatalogExists ? "No" : "Yes",
    };
  });
}

export function selectedIds(rows) {
  const selected = [];
  const seen = new Set();
  for (const row of rows) {
    const id = text(row.id);
    if (!id) throw new Error("Selection row contains a blank stable ID.");
    if (seen.has(id)) throw new Error(`Selection repeats stable ID ${id}.`);
    seen.add(id);
    if (normalizeInclude(row.include) === "Yes") selected.push(id);
  }
  return selected;
}

export function aggregateSelectedRestaurantResults({
  results,
  selectedRestaurantIds,
  importId,
  reportingGroupId,
  metricId = "KPI-0001",
  channelScope = "All channels",
  publicationState = "Active Finalized",
}) {
  const index = indexRestaurantResults(results, metricId, channelScope, publicationState);
  return aggregateFromIndex({
    index,
    selectedRestaurantIds,
    importId,
    reportingGroupId,
  });
}

function aggregateFromIndex({ index, selectedRestaurantIds, importId, reportingGroupId }) {
  const seenRestaurants = new Set();
  let numeratorSalesNok = 0;
  let denominatorSalesNok = 0;

  for (const rawRestaurantId of selectedRestaurantIds) {
    const restaurantId = text(rawRestaurantId);
    if (!restaurantId) throw new Error("Selected restaurant set contains a blank RestaurantID.");
    if (seenRestaurants.has(restaurantId)) {
      throw new Error(`Selected restaurant set repeats RestaurantID ${restaurantId}.`);
    }
    seenRestaurants.add(restaurantId);
    const key = restaurantResultKey(importId, reportingGroupId, restaurantId);
    const row = index.get(key);
    if (!row) throw new Error(`Missing Restaurant result ${key}.`);
    numeratorSalesNok += number(row.numeratorSalesNok);
    denominatorSalesNok += number(row.denominatorSalesNok);
  }

  return canonicalComponents(numeratorSalesNok, denominatorSalesNok);
}

export function buildInteractivePerformanceMatrix({
  results,
  selectedRestaurantIds,
  selectedReportingGroupIds,
  currentImportId,
  comparisonImportId,
}) {
  const index = indexRestaurantResults(results, "KPI-0001", "All channels", "Active Finalized");
  const groupIds = uniqueStableIds(selectedReportingGroupIds, "ReportingGroupID");
  const restaurantIds = uniqueStableIds(selectedRestaurantIds, "RestaurantID");
  if (!restaurantIds.length || !groupIds.length) return [];
  const rows = [];
  for (const restaurantId of restaurantIds) {
    rows.push(buildMatrixRow({
      rowType: "Restaurant",
      restaurantId,
      aggregateRestaurantIds: [restaurantId],
      index,
      selectedReportingGroupIds: groupIds,
      currentImportId,
      comparisonImportId,
    }));
  }
  rows.push(buildMatrixRow({
    rowType: "Grand Total",
    restaurantId: "",
    aggregateRestaurantIds: restaurantIds,
    index,
    selectedReportingGroupIds: groupIds,
    currentImportId,
    comparisonImportId,
  }));
  return rows;
}

export function matrixDisplayValue(cell, mode) {
  if (!PERFORMANCE_MATRIX_MODES.includes(mode)) {
    throw new Error(`Unsupported Performance matrix mode ${text(mode) || "(blank)"}.`);
  }
  if (mode === "Current Sales NOK") {
    return cell.current.denominatorSalesNok === 0 ? null : cell.current.numeratorSalesNok;
  }
  if (mode === "NOK Impact") {
    if (cell.current.denominatorSalesNok === 0 || cell.comparison.denominatorSalesNok === 0) {
      return null;
    }
    return cell.current.numeratorSalesNok -
      (cell.comparison.numeratorSalesNok / cell.comparison.denominatorSalesNok) *
      cell.current.denominatorSalesNok;
  }
  if (mode === "Current Share") {
    return cell.current.denominatorSalesNok === 0 ? null : cell.current.metricValue;
  }
  if (mode === "Comparison Share") {
    return cell.comparison.denominatorSalesNok === 0 ? null : cell.comparison.metricValue;
  }
  if (cell.current.denominatorSalesNok === 0 || cell.comparison.denominatorSalesNok === 0) {
    return null;
  }
  return 100 * (cell.current.metricValue - cell.comparison.metricValue);
}

export function matrixTotalDisplayValue(row, mode) {
  const totalCell = aggregateSelectedGroupCells(row.cells || []);
  if (!totalCell) return null;
  if (mode === "Current Sales NOK") return totalCell.current.numeratorSalesNok;
  return matrixDisplayValue(totalCell, mode);
}

export function sortInteractivePerformanceRows({
  rows,
  mode,
  sortBy = "Total",
  order = "Highest first",
}) {
  if (!PERFORMANCE_MATRIX_MODES.includes(mode)) {
    throw new Error(`Unsupported Performance matrix mode ${text(mode) || "(blank)"}.`);
  }
  if (!PERFORMANCE_SORT_ORDERS.includes(order)) {
    throw new Error(`Unsupported Performance sort order ${text(order) || "(blank)"}.`);
  }
  const restaurantRows = rows.filter(row => row.rowType === "Restaurant");
  const grandTotalRows = rows.filter(row => row.rowType === "Grand Total");
  if (grandTotalRows.length > 1) throw new Error("Interactive matrix repeats Grand Total.");
  const requestedGroupId = sortBy === "Total" ? "" : text(sortBy);
  const displayedGroupIds = restaurantRows.length
    ? restaurantRows[0].cells.map(cell => text(cell.reportingGroupId))
    : [];
  const targetDisplayed = !requestedGroupId || displayedGroupIds.includes(requestedGroupId);
  const effectiveSortBy = targetDisplayed ? sortBy : "Total";
  const direction = order === "Highest first" ? -1 : 1;
  const decorated = restaurantRows.map(row => ({
    row,
    value: effectiveSortBy === "Total"
      ? matrixTotalDisplayValue(row, mode)
      : matrixDisplayValue(
        row.cells.find(cell => text(cell.reportingGroupId) === effectiveSortBy),
        mode,
      ),
  }));
  decorated.sort((left, right) => {
    const leftAvailable = typeof left.value === "number" && Number.isFinite(left.value);
    const rightAvailable = typeof right.value === "number" && Number.isFinite(right.value);
    if (leftAvailable !== rightAvailable) return leftAvailable ? -1 : 1;
    if (!leftAvailable) return left.row.restaurantId.localeCompare(right.row.restaurantId);
    if (left.value !== right.value) return direction * (left.value - right.value);
    const idOrder = left.row.restaurantId.localeCompare(right.row.restaurantId);
    return order === "Highest first" ? idOrder : -idOrder;
  });
  return {
    rows: decorated.map(value => value.row).concat(grandTotalRows),
    effectiveSortBy,
    fallbackToTotal: !targetDisplayed,
  };
}

export function validateAllSelectedAgainstCompany({
  results,
  restaurantIds,
  importIds,
  reportingGroupIds,
}) {
  const errors = [];
  for (const importId of importIds) {
    for (const reportingGroupId of reportingGroupIds) {
      const aggregate = aggregateSelectedRestaurantResults({
        results,
        selectedRestaurantIds: restaurantIds,
        importId,
        reportingGroupId,
      });
      const companyRows = results.filter(row =>
        text(row.metricId) === "KPI-0001" &&
        text(row.importId) === text(importId) &&
        text(row.reportingGroupId) === text(reportingGroupId) &&
        text(row.scopeType) === "Company" &&
        text(row.channelScope) === "All channels" &&
        text(row.publicationState) === "Active Finalized"
      );
      if (companyRows.length !== 1) {
        errors.push(`${importId}/${reportingGroupId} has ${companyRows.length} Company results; expected one.`);
        continue;
      }
      const company = companyRows[0];
      if (!almostEqual(aggregate.numeratorSalesNok, number(company.numeratorSalesNok)) ||
          !almostEqual(aggregate.denominatorSalesNok, number(company.denominatorSalesNok)) ||
          !almostEqual(aggregate.metricValue, number(company.metricValue))) {
        errors.push(`${importId}/${reportingGroupId} Restaurant aggregation differs from Company control.`);
      }
    }
  }
  return errors;
}

function buildMatrixRow({
  rowType,
  restaurantId,
  aggregateRestaurantIds,
  index,
  selectedReportingGroupIds,
  currentImportId,
  comparisonImportId,
}) {
  const cells = [];
  for (const reportingGroupId of selectedReportingGroupIds) {
    cells.push({
      reportingGroupId,
      current: aggregateFromIndex({
        index,
        selectedRestaurantIds: aggregateRestaurantIds,
        importId: currentImportId,
        reportingGroupId,
      }),
      comparison: aggregateFromIndex({
        index,
        selectedRestaurantIds: aggregateRestaurantIds,
        importId: comparisonImportId,
        reportingGroupId,
      }),
    });
  }
  return { rowType, restaurantId: text(restaurantId), cells };
}

function aggregateSelectedGroupCells(cells) {
  if (!cells.length) return null;
  let currentNumerator = 0;
  let comparisonNumerator = 0;
  const currentDenominator = number(cells[0].current.denominatorSalesNok);
  const comparisonDenominator = number(cells[0].comparison.denominatorSalesNok);
  for (const cell of cells) {
    if (!almostEqual(number(cell.current.denominatorSalesNok), currentDenominator) ||
        !almostEqual(number(cell.comparison.denominatorSalesNok), comparisonDenominator)) {
      throw new Error("Selected Reporting Group cells do not share one scope denominator.");
    }
    currentNumerator += number(cell.current.numeratorSalesNok);
    comparisonNumerator += number(cell.comparison.numeratorSalesNok);
  }
  return {
    current: canonicalComponents(currentNumerator, currentDenominator),
    comparison: canonicalComponents(comparisonNumerator, comparisonDenominator),
  };
}

function uniqueStableIds(values, label) {
  const output = [];
  const seen = new Set();
  for (const rawValue of values) {
    const value = text(rawValue);
    if (!value) throw new Error(`Selected ${label} set contains a blank ID.`);
    if (seen.has(value)) throw new Error(`Selected ${label} set repeats ${value}.`);
    seen.add(value);
    output.push(value);
  }
  return output;
}

function indexRestaurantResults(results, metricId, channelScope, publicationState) {
  const index = new Map();
  for (const row of results) {
    if (text(row.metricId) !== metricId ||
        text(row.scopeType) !== "Restaurant" ||
        text(row.channelScope) !== channelScope ||
        text(row.publicationState) !== publicationState) continue;
    const key = restaurantResultKey(row.importId, row.reportingGroupId, row.restaurantId);
    if (index.has(key)) throw new Error(`Duplicate Restaurant result ${key}.`);
    index.set(key, row);
  }
  return index;
}

function restaurantResultKey(importId, reportingGroupId, restaurantId) {
  return [importId, reportingGroupId, restaurantId].map(text).join("|");
}

function canonicalComponents(numeratorSalesNok, denominatorSalesNok) {
  return {
    numeratorSalesNok,
    denominatorSalesNok,
    metricValue: denominatorSalesNok === 0 ? 0 : numeratorSalesNok / denominatorSalesNok,
  };
}

function normalizeInclude(value) {
  return text(value) === "Yes" ? "Yes" : "No";
}

function text(value) {
  return String(value ?? "").trim();
}

function number(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function almostEqual(left, right) {
  return Math.abs(left - right) <= Math.max(0.000001, Math.abs(right) * 1e-12);
}
