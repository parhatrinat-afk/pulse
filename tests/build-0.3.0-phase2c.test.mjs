import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";
import {
  PERFORMANCE_MATRIX_MODES,
  PERFORMANCE_SORT_ORDERS,
  aggregateSelectedRestaurantResults,
  buildInteractivePerformanceMatrix,
  matrixDisplayValue,
  matrixTotalDisplayValue,
  planSelectionRows,
  selectedIds,
  sortInteractivePerformanceRows,
  validateAllSelectedAgainstCompany,
} from "../src/reporting/interactive-performance.mjs";

const restaurantIds = ["RST-1", "RST-2"];
const groupIds = ["RPG-0001", "RPG-0002"];
const importIds = ["IMP-CURRENT", "IMP-COMPARE"];
const results = buildFixtureResults();
const acceptedPhase2B = JSON.parse(fs.readFileSync(
  new URL("./expected-build-0.3.0-phase2b.json", import.meta.url),
  "utf8",
));

test("Phase 2C exposes the approved displays and sort orders without a selection mode", () => {
  assert.deepEqual(PERFORMANCE_SORT_ORDERS, ["Highest first", "Lowest first"]);
  assert.deepEqual(PERFORMANCE_MATRIX_MODES, [
    "PP Change", "Current Share", "Comparison Share", "Current Sales NOK", "NOK Impact",
  ]);
});

test("one selected RPG makes Total equal that RPG in every display mode", () => {
  const row = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: ["RST-1"],
    selectedReportingGroupIds: ["RPG-0001"],
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  })[0];
  for (const mode of PERFORMANCE_MATRIX_MODES) {
    assert.equal(matrixTotalDisplayValue(row, mode), matrixDisplayValue(row.cells[0], mode));
  }
});

test("multiple selected RPGs make Total use exactly those RPG numerators and one denominator", () => {
  const oneGroup = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: ["RST-1"],
    selectedReportingGroupIds: ["RPG-0001"],
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  })[0];
  const twoGroups = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: ["RST-1"],
    selectedReportingGroupIds: groupIds,
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  })[0];
  assert.equal(matrixTotalDisplayValue(oneGroup, "Current Sales NOK"), 20);
  assert.equal(matrixTotalDisplayValue(twoGroups, "Current Sales NOK"), 25);
  assert.equal(matrixTotalDisplayValue(twoGroups, "Current Share"), 25 / 50);
  assert.equal(matrixTotalDisplayValue(twoGroups, "Comparison Share"), 25 / 50);
  assert.equal(matrixTotalDisplayValue(twoGroups, "PP Change"), 0);
  assert.equal(matrixTotalDisplayValue(twoGroups, "NOK Impact"), 0);
  const inconsistent = structuredClone(twoGroups);
  inconsistent.cells[1].current.denominatorSalesNok += 1;
  assert.throws(
    () => matrixTotalDisplayValue(inconsistent, "Current Share"),
    /do not share one scope denominator/,
  );
});

test("All-RPG Total does not become 100 percent when denominator includes non-RPG sales", () => {
  const total = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: restaurantIds,
    selectedReportingGroupIds: groupIds,
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  }).at(-1);
  assert.equal(matrixTotalDisplayValue(total, "Current Sales NOK"), 80);
  assert.equal(matrixTotalDisplayValue(total, "Current Share"), 80 / 200);
  assert.notEqual(matrixTotalDisplayValue(total, "Current Share"), 1);
});

test("Current Sales NOK Total remains the selected numerator without a denominator dependency", () => {
  const row = displayRow("RST-ZERO", [
    displayCell("RPG-0001", 7, 0, 2, 0),
    displayCell("RPG-0002", 5, 0, 3, 0),
  ]);
  assert.equal(matrixDisplayValue(row.cells[0], "Current Sales NOK"), null);
  assert.equal(matrixTotalDisplayValue(row, "Current Sales NOK"), 12);
  assert.equal(matrixTotalDisplayValue(row, "Current Share"), null);
  assert.equal(matrixTotalDisplayValue(row, "NOK Impact"), null);
});

test("stable-ID selection preserves prior choices and defaults newly eligible IDs to No", () => {
  const planned = planSelectionRows({
    eligibleItems: [
      { id: "RST-1", name: "First renamed" },
      { id: "RST-2", name: "Second" },
      { id: "RST-3", name: "New" },
    ],
    priorCatalogExists: true,
    priorRows: [
      { id: "RST-1", name: "Old display", include: "Yes" },
      { id: "RST-2", name: "Second", include: "No" },
    ],
  });
  assert.deepEqual(planned, [
    { id: "RST-1", name: "First renamed", include: "Yes" },
    { id: "RST-2", name: "Second", include: "No" },
    { id: "RST-3", name: "New", include: "No" },
  ]);
  assert.ok(planned.every(row => !Object.hasOwn(row, "nameIdentity")));
});

test("initial restaurant and RPG catalogs default every eligible ID to Yes", () => {
  const plannedRestaurants = planSelectionRows({
    eligibleItems: restaurantIds.map(id => ({ id, name: id })),
  });
  const plannedGroups = planSelectionRows({
    eligibleItems: groupIds.map(id => ({ id, name: id })),
  });
  assert.deepEqual(selectedIds(plannedRestaurants), restaurantIds);
  assert.deepEqual(selectedIds(plannedGroups), groupIds);
  assert.ok(plannedRestaurants.concat(plannedGroups).every(row => row.include === "Yes"));
});

test("Include state directly determines selection without a separate mode", () => {
  const rows = [
    { id: "RST-1", include: "No" },
    { id: "RST-2", include: "Yes" },
  ];
  assert.deepEqual(selectedIds(rows), ["RST-2"]);
  rows[0].include = "Yes";
  assert.deepEqual(selectedIds(rows), restaurantIds);
  assert.throws(() => selectedIds([...rows, rows[0]]), /repeats/);
});

test("interactive scope sums Phase 2B Restaurant components", () => {
  const combined = aggregateSelectedRestaurantResults({
    results,
    selectedRestaurantIds: restaurantIds,
    importId: "IMP-CURRENT",
    reportingGroupId: "RPG-0001",
  });
  assert.deepEqual(combined, {
    numeratorSalesNok: 30,
    denominatorSalesNok: 200,
    metricValue: 0.15,
  });
  const custom = aggregateSelectedRestaurantResults({
    results,
    selectedRestaurantIds: ["RST-1"],
    importId: "IMP-CURRENT",
    reportingGroupId: "RPG-0001",
  });
  assert.deepEqual(custom, {
    numeratorSalesNok: 20,
    denominatorSalesNok: 50,
    metricValue: 0.4,
  });
  const sourceRestaurantRow = results.find(row =>
    row.importId === "IMP-CURRENT" &&
    row.reportingGroupId === "RPG-0001" &&
    row.scopeType === "Restaurant" &&
    row.restaurantId === "RST-1"
  );
  assert.ok(sourceRestaurantRow);
  assert.deepEqual(custom, {
    numeratorSalesNok: sourceRestaurantRow.numeratorSalesNok,
    denominatorSalesNok: sourceRestaurantRow.denominatorSalesNok,
    metricValue: sourceRestaurantRow.metricValue,
  });
});

test("deselecting and reselecting an eligible restaurant directly changes scope and restores exactly", () => {
  const selectionRows = restaurantIds.map(id => ({ id, include: "Yes" }));
  const allSelected = aggregateSelectedRestaurantResults({
    results,
    selectedRestaurantIds: selectedIds(selectionRows),
    importId: "IMP-CURRENT",
    reportingGroupId: "RPG-0001",
  });
  selectionRows[1].include = "No";
  const withoutSecond = aggregateSelectedRestaurantResults({
    results,
    selectedRestaurantIds: selectedIds(selectionRows),
    importId: "IMP-CURRENT",
    reportingGroupId: "RPG-0001",
  });
  selectionRows[1].include = "Yes";
  const restored = aggregateSelectedRestaurantResults({
    results,
    selectedRestaurantIds: selectedIds(selectionRows),
    importId: "IMP-CURRENT",
    reportingGroupId: "RPG-0001",
  });
  assert.equal(allSelected.numeratorSalesNok, 30);
  assert.equal(allSelected.denominatorSalesNok, 200);
  assert.equal(withoutSecond.numeratorSalesNok, 20);
  assert.equal(withoutSecond.denominatorSalesNok, 50);
  assert.deepEqual(restored, allSelected);
});

test("Grand Total uses summed components for weighted share and NOK Impact", () => {
  const matrix = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: restaurantIds,
    selectedReportingGroupIds: ["RPG-0001"],
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  });
  assert.equal(matrix.length, 3);
  assert.equal(matrix[0].cells[0].current.metricValue, 0.4);
  assert.equal(matrix[1].cells[0].current.metricValue, 10 / 150);
  assert.notEqual(matrix[2].cells[0].current.metricValue, (0.4 + 10 / 150) / 2);
  assert.equal(matrix[2].cells[0].current.numeratorSalesNok, 30);
  assert.equal(matrix[2].cells[0].current.denominatorSalesNok, 200);
  assert.equal(matrix[2].cells[0].current.metricValue, 30 / 200);
  assert.equal(matrix[2].cells[0].comparison.numeratorSalesNok, 40);
  assert.equal(
    matrixDisplayValue(matrix[2].cells[0], "NOK Impact"),
    30 - ((40 / 200) * 200),
  );
});

test("Grand Total NOK Impact uses aggregated shares and current sales, not SUM of row impacts", () => {
  const matrix = buildInteractivePerformanceMatrix({
    results: buildImpactAggregationResults(),
    selectedRestaurantIds: restaurantIds,
    selectedReportingGroupIds: ["RPG-0001"],
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  });
  const rowImpactSum = matrix.slice(0, -1)
    .reduce((sum, row) => sum + matrixDisplayValue(row.cells[0], "NOK Impact"), 0);
  const total = matrix.at(-1).cells[0];
  const expectedAggregateImpact = 40 - ((115 / 1000) * 200);
  assert.equal(rowImpactSum, -10);
  assert.equal(matrixDisplayValue(total, "NOK Impact"), expectedAggregateImpact);
  assert.equal(expectedAggregateImpact, 17);
  assert.notEqual(matrixDisplayValue(total, "NOK Impact"), rowImpactSum);
});

test("several direct restaurant Include selections aggregate without materializing a combination", () => {
  const threeRestaurantResults = buildGeneratedResults(
    ["RST-1", "RST-2", "RST-3"],
    ["RPG-0001"],
  );
  const selectionRows = [
    { id: "RST-1", include: "Yes" },
    { id: "RST-2", include: "Yes" },
    { id: "RST-3", include: "Yes" },
    { id: "RST-4", include: "No" },
  ];
  const combined = aggregateSelectedRestaurantResults({
    results: threeRestaurantResults,
    selectedRestaurantIds: selectedIds(selectionRows),
    importId: "IMP-CURRENT",
    reportingGroupId: "RPG-0001",
  });
  assert.equal(combined.numeratorSalesNok, 60);
  assert.equal(combined.denominatorSalesNok, 600);
  assert.equal(combined.metricValue, 0.1);
  assert.equal(threeRestaurantResults.filter(row => row.scopeType === "Company").length, 2);
  assert.equal(threeRestaurantResults.filter(row => row.scopeType === "Restaurant").length, 6);
});

test("the original four display modes remain unchanged and NOK Impact uses canonical components", () => {
  const matrix = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: ["RST-1"],
    selectedReportingGroupIds: ["RPG-0001"],
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  });
  const cell = matrix[0].cells[0];
  const unchangedComponents = structuredClone(cell);
  assert.equal(matrixDisplayValue(cell, "Current Share"), 0.4);
  assert.equal(matrixDisplayValue(cell, "Comparison Share"), 0.2);
  assert.equal(matrixDisplayValue(cell, "PP Change"), 20);
  assert.equal(matrixDisplayValue(cell, "Current Sales NOK"), 20);
  assert.equal(matrixDisplayValue(cell, "NOK Impact"), 10);
  assert.deepEqual(cell, unchangedComponents, "switching display mode must not mutate metric components");
  const zero = {
    current: { numeratorSalesNok: 0, denominatorSalesNok: 0, metricValue: 0 },
    comparison: { numeratorSalesNok: 0, denominatorSalesNok: 0, metricValue: 0 },
  };
  for (const mode of PERFORMANCE_MATRIX_MODES) {
    assert.equal(matrixDisplayValue(zero, mode), null);
  }
  const zeroCurrent = {
    current: { numeratorSalesNok: 0, denominatorSalesNok: 0, metricValue: 0 },
    comparison: { numeratorSalesNok: 10, denominatorSalesNok: 100, metricValue: 0.1 },
  };
  const zeroComparison = {
    current: { numeratorSalesNok: 10, denominatorSalesNok: 100, metricValue: 0.1 },
    comparison: { numeratorSalesNok: 0, denominatorSalesNok: 0, metricValue: 0 },
  };
  assert.equal(matrixDisplayValue(zeroCurrent, "NOK Impact"), null);
  assert.equal(matrixDisplayValue(zeroComparison, "NOK Impact"), null);
});

test("all five display modes preserve typed numeric components for the presentation facade", () => {
  const matrix = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: ["RST-1"],
    selectedReportingGroupIds: ["RPG-0001"],
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  });
  const cell = matrix[0].cells[0];
  const numericSnapshot = structuredClone(cell);
  for (const mode of PERFORMANCE_MATRIX_MODES) {
    const value = matrixDisplayValue(cell, mode);
    assert.equal(typeof value, "number");
    assert.deepEqual(cell, numericSnapshot);
  }
});

test("all-selected restaurant scope reconciles to retained Company QA controls", () => {
  assert.deepEqual(validateAllSelectedAgainstCompany({
    results,
    restaurantIds,
    importIds,
    reportingGroupIds: groupIds,
  }), []);
  const changed = results.map(row => ({ ...row }));
  changed.find(row => row.scopeType === "Company").denominatorSalesNok += 1;
  assert.match(validateAllSelectedAgainstCompany({
    results: changed,
    restaurantIds,
    importIds,
    reportingGroupIds: groupIds,
  }).join(" "), /differs from Company control/);
});

test("RPG Include directly removes and restores columns while Total follows selected RPGs", () => {
  const rows = [
    { id: "RPG-0001", name: "Add-ons", include: "Yes" },
    { id: "RPG-0002", name: "Non-Alcohol", include: "No" },
    { id: "RPG-0009", name: "Mains", include: "Yes" },
  ];
  assert.deepEqual(selectedIds(rows), ["RPG-0001", "RPG-0009"]);
  rows[2].include = "No";
  assert.deepEqual(selectedIds(rows), ["RPG-0001"]);
  rows[1].include = "Yes";
  assert.deepEqual(selectedIds(rows), ["RPG-0001", "RPG-0002"]);
});

test("deselecting one RPG immediately recalculates Total and reselection restores it", () => {
  const selectionRows = groupIds.map(id => ({ id, include: "Yes" }));
  const totalForSelection = () => {
    const row = buildInteractivePerformanceMatrix({
      results,
      selectedRestaurantIds: ["RST-1"],
      selectedReportingGroupIds: selectedIds(selectionRows),
      currentImportId: "IMP-CURRENT",
      comparisonImportId: "IMP-COMPARE",
    })[0];
    return matrixTotalDisplayValue(row, "Current Sales NOK");
  };
  assert.equal(totalForSelection(), 25);
  selectionRows[1].include = "No";
  assert.equal(totalForSelection(), 20);
  selectionRows[1].include = "Yes";
  assert.equal(totalForSelection(), 25);
});

test("positive and negative PP changes produce matching NOK Impact signs", () => {
  const normal = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: ["RST-1"],
    selectedReportingGroupIds: ["RPG-0001"],
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  })[0].cells[0];
  const swapped = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: ["RST-1"],
    selectedReportingGroupIds: ["RPG-0001"],
    currentImportId: "IMP-COMPARE",
    comparisonImportId: "IMP-CURRENT",
  })[0].cells[0];
  assert.equal(matrixDisplayValue(normal, "PP Change"), 20);
  assert.equal(matrixDisplayValue(swapped, "PP Change"), -20);
  assert.ok(matrixDisplayValue(normal, "NOK Impact") > 0);
  assert.ok(matrixDisplayValue(swapped, "NOK Impact") < 0);
});

test("same-dataset comparison returns zero NOK Impact", () => {
  const same = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: ["RST-1"],
    selectedReportingGroupIds: ["RPG-0001"],
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-CURRENT",
  })[0].cells[0];
  assert.equal(matrixDisplayValue(same, "PP Change"), 0);
  assert.equal(matrixDisplayValue(same, "NOK Impact"), 0);
});

test("accepted Add-ons fixture produces about +34,019 NOK Impact despite different period lengths", () => {
  const currentDataset = acceptedPhase2B.datasets["IMP-2026-W31"];
  const comparisonDataset = acceptedPhase2B.datasets["IMP-2025-BASELINE"];
  const currentNumerator = currentDataset.reporting_groups["RPG-0001"].sales_nok;
  const comparisonNumerator = comparisonDataset.reporting_groups["RPG-0001"].sales_nok;
  const cell = {
    current: {
      numeratorSalesNok: currentNumerator,
      denominatorSalesNok: currentDataset.sales_nok,
      metricValue: currentNumerator / currentDataset.sales_nok,
    },
    comparison: {
      numeratorSalesNok: comparisonNumerator,
      denominatorSalesNok: comparisonDataset.sales_nok,
      metricValue: comparisonNumerator / comparisonDataset.sales_nok,
    },
  };
  const impact = matrixDisplayValue(cell, "NOK Impact");
  assert.ok(Math.abs(impact - 34019.301729693412) < 0.000001);
  assert.equal(Math.round(impact), 34019);
  assert.notEqual(impact, currentNumerator - comparisonNumerator);
});

test("zero restaurant or zero RPG selection produces a blank matrix", () => {
  assert.deepEqual(buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: [],
    selectedReportingGroupIds: ["RPG-0001"],
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  }), []);
  assert.deepEqual(buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: ["RST-1"],
    selectedReportingGroupIds: [],
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  }), []);
});

test("catalog generation exposes only refresh-eligible IDs and supports all nine RPGs", () => {
  const restaurantCatalog = planSelectionRows({
    eligibleItems: [{ id: "RST-1", name: "First" }],
    priorCatalogExists: true,
    priorRows: [{ id: "RST-INELIGIBLE", name: "Hidden", include: "Yes" }],
  });
  assert.deepEqual(restaurantCatalog, [{ id: "RST-1", name: "First", include: "No" }]);
  const nineGroups = Array.from({ length: 9 }, (_, index) => ({
    id: `RPG-${String(index + 1).padStart(4, "0")}`,
    name: index === 8 ? "Mains" : `Group ${index + 1}`,
  }));
  const groupCatalog = planSelectionRows({ eligibleItems: nineGroups });
  assert.equal(groupCatalog.length, 9);
  assert.equal(selectedIds(groupCatalog).length, 9);
});

test("Highest and Lowest reverse available numeric order while unavailable rows remain last", () => {
  const rows = [
    displayRow("RST-A", [displayCell("RPG-0001", 10, 100, 10, 100)]),
    displayRow("RST-B", [displayCell("RPG-0001", 30, 100, 10, 100)]),
    displayRow("RST-C", [displayCell("RPG-0001", 20, 100, 10, 100)]),
    displayRow("RST-U2", [displayCell("RPG-0001", 0, 0, 10, 100)]),
    displayRow("RST-U1", [displayCell("RPG-0001", 0, 0, 10, 100)]),
    { rowType: "Grand Total", restaurantId: "", cells: [displayCell("RPG-0001", 60, 300, 40, 400)] },
  ];
  const highest = sortInteractivePerformanceRows({
    rows, mode: "Current Share", sortBy: "RPG-0001", order: "Highest first",
  }).rows.map(row => row.restaurantId || row.rowType);
  const lowest = sortInteractivePerformanceRows({
    rows, mode: "Current Share", sortBy: "RPG-0001", order: "Lowest first",
  }).rows.map(row => row.restaurantId || row.rowType);
  assert.deepEqual(highest, ["RST-B", "RST-C", "RST-A", "RST-U1", "RST-U2", "Grand Total"]);
  assert.deepEqual(lowest, ["RST-A", "RST-C", "RST-B", "RST-U1", "RST-U2", "Grand Total"]);
  assert.deepEqual(lowest.slice(0, 3), highest.slice(0, 3).reverse());
});

test("exact numeric ties resolve deterministically by RestaurantID", () => {
  const rows = [
    displayRow("RST-B", [displayCell("RPG-0001", 20, 100, 10, 100)]),
    displayRow("RST-A", [displayCell("RPG-0001", 20, 100, 10, 100)]),
  ];
  const highest = sortInteractivePerformanceRows({
    rows, mode: "Current Sales NOK", sortBy: "RPG-0001", order: "Highest first",
  }).rows.map(row => row.restaurantId);
  const lowest = sortInteractivePerformanceRows({
    rows, mode: "Current Sales NOK", sortBy: "RPG-0001", order: "Lowest first",
  }).rows.map(row => row.restaurantId);
  assert.deepEqual(highest, ["RST-A", "RST-B"]);
  assert.deepEqual(lowest, ["RST-B", "RST-A"]);
});

test("PP Change and NOK Impact sorting use full numeric precision rather than visible rounding", () => {
  const ppRows = [
    displayRow("RST-PP-LOW", [displayCell("RPG-0001", 10.003, 100, 10, 100)]),
    displayRow("RST-PP-HIGH", [displayCell("RPG-0001", 10.004, 100, 10, 100)]),
  ];
  const ppOrder = sortInteractivePerformanceRows({
    rows: ppRows, mode: "PP Change", sortBy: "RPG-0001", order: "Highest first",
  }).rows.map(row => row.restaurantId);
  assert.deepEqual(ppOrder, ["RST-PP-HIGH", "RST-PP-LOW"]);
  assert.equal(matrixDisplayValue(ppRows[0].cells[0], "PP Change").toFixed(2), "0.00");
  assert.equal(matrixDisplayValue(ppRows[1].cells[0], "PP Change").toFixed(2), "0.00");

  const impactRows = [
    displayRow("RST-NOK-LOW", [displayCell("RPG-0001", 20.40, 100, 10, 100)]),
    displayRow("RST-NOK-HIGH", [displayCell("RPG-0001", 20.49, 100, 10, 100)]),
  ];
  const impactOrder = sortInteractivePerformanceRows({
    rows: impactRows, mode: "NOK Impact", sortBy: "RPG-0001", order: "Highest first",
  }).rows.map(row => row.restaurantId);
  assert.deepEqual(impactOrder, ["RST-NOK-HIGH", "RST-NOK-LOW"]);
  assert.equal(matrixDisplayValue(impactRows[0].cells[0], "NOK Impact").toFixed(0), "10");
  assert.equal(matrixDisplayValue(impactRows[1].cells[0], "NOK Impact").toFixed(0), "10");
});

test("sorting changes presentation order only and preserves canonical components", () => {
  const matrix = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: restaurantIds,
    selectedReportingGroupIds: groupIds,
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  });
  const canonicalSnapshot = structuredClone(matrix);
  const highest = sortInteractivePerformanceRows({
    rows: matrix, mode: "Current Sales NOK", sortBy: "Total", order: "Highest first",
  });
  const lowest = sortInteractivePerformanceRows({
    rows: matrix, mode: "Current Sales NOK", sortBy: "Total", order: "Lowest first",
  });
  assert.deepEqual(matrix, canonicalSnapshot);
  assert.deepEqual(
    highest.rows.find(row => row.restaurantId === "RST-1"),
    lowest.rows.find(row => row.restaurantId === "RST-1"),
  );
  assert.notDeepEqual(
    highest.rows.map(row => row.restaurantId),
    lowest.rows.map(row => row.restaurantId),
  );
});

test("direct restaurant selection sorts only selected restaurants and keeps Grand Total last", () => {
  const matrix = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: ["RST-2"],
    selectedReportingGroupIds: groupIds,
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  });
  const sorted = sortInteractivePerformanceRows({
    rows: matrix, mode: "Current Sales NOK", sortBy: "Total", order: "Highest first",
  }).rows;
  assert.deepEqual(sorted.map(row => row.restaurantId || row.rowType), ["RST-2", "Grand Total"]);
});

test("changing Display or datasets immediately changes numeric presentation order", () => {
  const displayRows = [
    displayRow("RST-A", [displayCell("RPG-0001", 100, 1000, 20, 100)]),
    displayRow("RST-B", [displayCell("RPG-0001", 50, 100, 40, 100)]),
  ];
  const salesOrder = sortInteractivePerformanceRows({
    rows: displayRows, mode: "Current Sales NOK", sortBy: "Total", order: "Highest first",
  }).rows.map(row => row.restaurantId);
  const ppOrder = sortInteractivePerformanceRows({
    rows: displayRows, mode: "PP Change", sortBy: "Total", order: "Highest first",
  }).rows.map(row => row.restaurantId);
  assert.deepEqual(salesOrder, ["RST-A", "RST-B"]);
  assert.deepEqual(ppOrder, ["RST-B", "RST-A"]);

  const currentMatrix = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: restaurantIds,
    selectedReportingGroupIds: ["RPG-0001"],
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  });
  const swappedMatrix = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: restaurantIds,
    selectedReportingGroupIds: ["RPG-0001"],
    currentImportId: "IMP-COMPARE",
    comparisonImportId: "IMP-CURRENT",
  });
  assert.deepEqual(sortInteractivePerformanceRows({
    rows: currentMatrix, mode: "Current Sales NOK", sortBy: "RPG-0001", order: "Highest first",
  }).rows.slice(0, 2).map(row => row.restaurantId), ["RST-1", "RST-2"]);
  assert.deepEqual(sortInteractivePerformanceRows({
    rows: swappedMatrix, mode: "Current Sales NOK", sortBy: "RPG-0001", order: "Highest first",
  }).rows.slice(0, 2).map(row => row.restaurantId), ["RST-2", "RST-1"]);
});

test("hidden RPG sort target falls back to Total and automatically resumes when restored", () => {
  const matrix = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: restaurantIds,
    selectedReportingGroupIds: ["RPG-0001"],
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  });
  const sorted = sortInteractivePerformanceRows({
    rows: matrix, mode: "Current Share", sortBy: "RPG-0002", order: "Highest first",
  });
  assert.equal(sorted.effectiveSortBy, "Total");
  assert.equal(sorted.fallbackToTotal, true);
  assert.deepEqual(
    sorted.rows.map(row => row.restaurantId),
    sortInteractivePerformanceRows({
      rows: matrix, mode: "Current Share", sortBy: "Total", order: "Highest first",
    }).rows.map(row => row.restaurantId),
  );
  const restoredMatrix = buildInteractivePerformanceMatrix({
    results,
    selectedRestaurantIds: restaurantIds,
    selectedReportingGroupIds: groupIds,
    currentImportId: "IMP-CURRENT",
    comparisonImportId: "IMP-COMPARE",
  });
  const restored = sortInteractivePerformanceRows({
    rows: restoredMatrix, mode: "Current Share", sortBy: "RPG-0002", order: "Highest first",
  });
  assert.equal(restored.effectiveSortBy, "RPG-0002");
  assert.equal(restored.fallbackToTotal, false);
});

test("Phase 2C Office Script uses centralized additive rows and formula-driven interaction", () => {
  const source = fs.readFileSync(
    new URL("../office-scripts/Build_0_3_0_Phase2C.ts", import.meta.url),
    "utf8",
  );
  const componentWriter = source.slice(
    source.indexOf("function writeComponentBlock("),
    source.indexOf("function writeShareBlock("),
  );
  assert.match(componentWriter, /SUMIFS\(tblMetricRPGResults\[/);
  assert.match(componentWriter, /ScopeType\],\"Restaurant\"/);
  assert.doesNotMatch(componentWriter, /ScopeType\],\"Company\"/);
  assert.doesNotMatch(componentWriter, /\$G\$6|Matrix display/);
  assert.doesNotMatch(componentWriter, /tblSalesFacts|tblMetricRPGFacts|ReportingCategoryID/);
  assert.match(source, /tblPerformanceRestaurantSelection/);
  assert.match(source, /tblPerformanceRPGSelection/);
  assert.match(source, /"Sort by", sortBy/);
  assert.match(source, /"Order", prior\.sortOrder/);
  assert.match(source, /\["Restaurants", ""\]/);
  assert.match(source, /\["Reporting Groups", ""\]/);
  assert.match(source, /"Restaurant", "Total"/);
  assert.match(source, /FILTER\(tblPerformanceRestaurantSelection/);
  assert.match(source, /FILTER\(tblPerformanceRPGSelection/);
  assert.match(source, /tblPerformanceRestaurantSelection\[Include\]=\"Yes\"/);
  assert.match(source, /tblPerformanceRPGSelection\[Include\]=\"Yes\"/);
  assert.doesNotMatch(source, /Restaurant mode|Reporting Group mode|SELECTION_MODES/);
  const interactionWriter = source.slice(
    source.indexOf("function writeInteractionCalc("),
    source.indexOf("function writeNumericDisplayBlock("),
  );
  assert.doesNotMatch(interactionWriter, /Performance!\$G\$7=\"All\"|Performance!\$G\$8=\"All\"/);
  assert.match(source, /Grand Total/);
  assert.match(source, /Performance!B16/);
  assert.match(source, /\["Detail Reporting Group", detailGroup\.name\]/);
  assert.match(source, /MATCH\('_Metric_Calc'!\$AL\$6/);
  const groupSelectionWriter = source.slice(
    source.indexOf("const selectedGroupFormulas"),
    source.indexOf("sheet.getRangeByIndexes(0, layout.controlStartColumn"),
  );
  assert.doesNotMatch(groupSelectionWriter, /Performance!\$B\$7/);
  assert.match(source, /Selected Numeric Display \| RPG-ID/);
  assert.match(source, /function writeNumericDisplayBlock\(/);
  assert.match(source, /function writeTotalAndSortHelpers\(/);
  assert.match(source, /function matrixTotalPresentationFormula\(/);
  assert.match(source, /function matrixPresentationFormula\(/);
  assert.match(source, /FIXED\(\$\{numericValue\},2,TRUE\)/);
  assert.match(source, /FIXED\(\$\{numericValue\}\*100,2,TRUE\)/);
  assert.match(source, /FIXED\(\$\{numericValue\},0,FALSE\)/);
  assert.match(source, /\$G\$6="NOK Impact"/);
  assert.match(source, /\$\{currentNumerator\}-\(\(\$\{comparisonNumerator\}\/\$\{comparisonDenominator\}\)\*\$\{currentDenominator\}\)/);
  assert.match(source, /Different period lengths — comparison allowed/);
  assert.match(source, /Interactive dropdown validation ready \(8\/8\)/);
  const totalAndSortWriter = source.slice(
    source.indexOf("function writeTotalAndSortHelpers("),
    source.indexOf("function componentCellReference("),
  );
  assert.match(totalAndSortWriter, /Current Sales NOK",\$\{currentNumerator\}/);
  assert.match(totalAndSortWriter, /\$\{currentDenominator\}\)\),""\)\)\)\)\)\)/);
  assert.match(totalAndSortWriter, /selectedNumeratorExpression/);
  assert.match(totalAndSortWriter, /SORTBY\(FILTER\(/);
  assert.match(totalAndSortWriter, /FILTER\(--\(\$\{keyRange\}=""\)/);
  assert.match(totalAndSortWriter, /IF\(\$AL\$14="Highest first",1,-1\)/);
  assert.match(totalAndSortWriter, /FILTER\(IF\(\$\{keyRange\}="","",\$\{idRange\}\)/);
  assert.doesNotMatch(totalAndSortWriter, /FIXED\(|matrixFacadeFormula|Performance!\$?[A-Z]+\$?2[4-9]/);
  assert.match(source, /Using Total — \"&Performance!\$I\$6&\" hidden/);
  assert.match(source, /All versus Custom is derived, not user-selected/);
  assert.match(source, /Sorting uses numeric values; Grand Total remains fixed\./);
  assert.match(source, /performance\.getRange\("7:7"\).*setRowHeight\(30\)/s);
  assert.match(source, /performance\.getRange\("8:8"\).*setRowHeight\(45\)/s);
  assert.match(source, /performance\.getRange\("18:19"\).*setRowHeight\(30\)/s);
  assert.match(source, /function applyLegacyAllMode\(/);
  assert.match(source, /COUNTIF\(\$AI\$2:\$AI\$/);
  assert.doesNotMatch(source, /matrixNumberFormatsFromDetailCells|MatrixNumberFormats/);
  assert.doesNotMatch(source, /getNumberFormatLocal\(\)/);
  assert.doesNotMatch(source, /getCultureInfo\(\)|getNumberDecimalSeparator\(\)/);
  const conditionalWriter = source.slice(
    source.indexOf("function applyMatrixConditionalFormats("),
    source.indexOf("function writeInteractionCalc("),
  );
  assert.doesNotMatch(conditionalWriter, /setNumberFormat/);
  assert.match(conditionalWriter, /getFont\(\)\.setColor/);
  assert.match(conditionalWriter, /OR\(\$G\$6="PP Change",\$G\$6="NOK Impact"\)/);
  assert.doesNotMatch(source, /Attach Rate|Pulse ♥|labour|reviews/i);
});

test("visible matrix text facade cannot feed authoritative metric calculations or Reports", () => {
  const source = fs.readFileSync(
    new URL("../office-scripts/Build_0_3_0_Phase2C.ts", import.meta.url),
    "utf8",
  );
  const numericWriter = source.slice(
    source.indexOf("function writeNumericDisplayBlock("),
    source.indexOf("function componentCellReference("),
  );
  const presentationWriter = source.slice(
    source.indexOf("function matrixPresentationFormula("),
    source.indexOf("function numericDisplayRangeExpression("),
  );
  const reportsWriter = source.slice(
    source.indexOf("function writeReports("),
    source.indexOf("function writeInteractionQA("),
  );
  assert.doesNotMatch(numericWriter, /FIXED\(|" pp"|" NOK"/);
  assert.match(numericWriter, /componentCellReference/);
  assert.match(numericWriter, /currentNumerator\}-\(\(\$\{comparisonNumerator\}\/\$\{comparisonDenominator\}\)\*\$\{currentDenominator\}\)/);
  assert.doesNotMatch(numericWriter, /SUM\(|AVERAGE\(|matrixPresentationFormula|Period|G16|G18/);
  assert.match(presentationWriter, /numericDisplayRangeExpression/);
  assert.match(presentationWriter, /FIXED\(/);
  assert.doesNotMatch(presentationWriter, /TEXT\(|SUBSTITUTE\(|NUMBERVALUE\(/);
  assert.doesNotMatch(presentationWriter, /SUMIFS\(|tblMetricRPGResults|tblSalesFacts/);
  assert.doesNotMatch(reportsWriter, /FIXED\(|Interactive Sales Performance/);
  assert.match(reportsWriter, /Performance!B7/);
  assert.match(reportsWriter, /Performance!B16/);
  assert.match(reportsWriter, /Performance!B19/);
  assert.doesNotMatch(reportsWriter, /Performance!\$?[A-Z]+\$?(?:2[4-9]|3[0-9]|40)/);
  assert.equal((source.match(/matrixPresentationFormula\(/g) || []).length, 2);
  assert.equal((source.match(/matrixTotalPresentationFormula\(/g) || []).length, 2);
});

test("QA-0302C-09 distinguishes the Current Sales NOK mode name from NOK presentation text", () => {
  const source = fs.readFileSync(
    new URL("../office-scripts/Build_0_3_0_Phase2C.ts", import.meta.url),
    "utf8",
  );
  const qaWriter = source.slice(
    source.indexOf("const presentationFacadeIsolated"),
    source.indexOf("const dashVisible"),
  );
  assert.match(qaWriter, /numericDisplayText\.indexOf\('\" NOK\"'\) < 0/);
  assert.match(qaWriter, /totalHelperText\.indexOf\('\" NOK\"'\) < 0/);
  assert.match(qaWriter, /visibleMatrixText\.indexOf\('\" NOK\"'\) >= 0/);
  assert.match(qaWriter, /visibleMatrixText\.indexOf\('\" pp\"'\) >= 0/);
  assert.doesNotMatch(qaWriter, /numericDisplayText\.indexOf\(" NOK"\)/);
  assert.match(source, /\["QA-0302C-09", "Five matrix display modes"/);
  assert.match(source, /"PP Change", "Current Share", "Comparison Share", "Current Sales NOK", "NOK Impact"/);

  const numericFormula = '=IF($AL$7="Current Sales NOK",$DF2,"")';
  const visibleFormula = '=IF($G$6="Current Sales NOK",FIXED($DF2,0,FALSE)&" NOK","")';
  assert.equal(numericFormula.includes(" NOK"), true, "the former broad check produces the false positive");
  assert.equal(numericFormula.includes('" NOK"'), false);
  assert.equal(visibleFormula.includes('" NOK"'), true);
});

test("Phase 2C Office Script avoids unsupported iterator constructs and looped workbook reads", () => {
  const source = fs.readFileSync(
    new URL("../office-scripts/Build_0_3_0_Phase2C.ts", import.meta.url),
    "utf8",
  );
  assert.doesNotMatch(source, /new\s+(Map|Set)\s*</);
  assert.doesNotMatch(source, /\.(entries|keys|values)\(\)/);
  assert.doesNotMatch(source, /Array\.from\(/);
  assert.doesNotMatch(source, /\.\.\.(?:new\s+)?(?:Map|Set)/);
  const loops = [...source.matchAll(/(?:for\s*\([^)]*\)|for\s*\([^)]*\sof\s[^)]*\))\s*\{[\s\S]*?\n\s*\}/g)]
    .map(match => match[0]);
  assert.ok(loops.every(loop => !/\.get(?:Values?|Text|Formulas?)\s*\(/.test(loop)));
});

function buildFixtureResults() {
  const components = {
    "IMP-CURRENT|RPG-0001|RST-1": [20, 50],
    "IMP-CURRENT|RPG-0001|RST-2": [10, 150],
    "IMP-CURRENT|RPG-0002|RST-1": [5, 50],
    "IMP-CURRENT|RPG-0002|RST-2": [45, 150],
    "IMP-COMPARE|RPG-0001|RST-1": [10, 50],
    "IMP-COMPARE|RPG-0001|RST-2": [30, 150],
    "IMP-COMPARE|RPG-0002|RST-1": [15, 50],
    "IMP-COMPARE|RPG-0002|RST-2": [15, 150],
  };
  const output = [];
  for (const importId of importIds) {
    for (const reportingGroupId of groupIds) {
      let companyNumerator = 0;
      let companyDenominator = 0;
      for (const restaurantId of restaurantIds) {
        const [numeratorSalesNok, denominatorSalesNok] =
          components[`${importId}|${reportingGroupId}|${restaurantId}`];
        companyNumerator += numeratorSalesNok;
        companyDenominator += denominatorSalesNok;
        output.push(resultRow({
          importId, reportingGroupId, restaurantId, scopeType: "Restaurant",
          numeratorSalesNok, denominatorSalesNok,
        }));
      }
      output.push(resultRow({
        importId, reportingGroupId, restaurantId: "", scopeType: "Company",
        numeratorSalesNok: companyNumerator,
        denominatorSalesNok: companyDenominator,
      }));
    }
  }
  return output;
}

function buildImpactAggregationResults() {
  const components = {
    "IMP-CURRENT|RST-1": [30, 100],
    "IMP-CURRENT|RST-2": [10, 100],
    "IMP-COMPARE|RST-1": [20, 50],
    "IMP-COMPARE|RST-2": [95, 950],
  };
  const output = [];
  for (const importId of importIds) {
    for (const restaurantId of restaurantIds) {
      const [numeratorSalesNok, denominatorSalesNok] =
        components[`${importId}|${restaurantId}`];
      output.push(resultRow({
        importId,
        reportingGroupId: "RPG-0001",
        restaurantId,
        scopeType: "Restaurant",
        numeratorSalesNok,
        denominatorSalesNok,
      }));
    }
  }
  return output;
}

function displayRow(restaurantId, cells) {
  return { rowType: "Restaurant", restaurantId, cells };
}

function displayCell(reportingGroupId, currentNumerator, currentDenominator, comparisonNumerator, comparisonDenominator) {
  return {
    reportingGroupId,
    current: {
      numeratorSalesNok: currentNumerator,
      denominatorSalesNok: currentDenominator,
      metricValue: currentDenominator === 0 ? 0 : currentNumerator / currentDenominator,
    },
    comparison: {
      numeratorSalesNok: comparisonNumerator,
      denominatorSalesNok: comparisonDenominator,
      metricValue: comparisonDenominator === 0 ? 0 : comparisonNumerator / comparisonDenominator,
    },
  };
}

function resultRow({
  importId,
  reportingGroupId,
  restaurantId,
  scopeType,
  numeratorSalesNok,
  denominatorSalesNok,
}) {
  return {
    metricId: "KPI-0001",
    importId,
    reportingGroupId,
    scopeType,
    restaurantId,
    channelScope: "All channels",
    publicationState: "Active Finalized",
    numeratorSalesNok,
    denominatorSalesNok,
    metricValue: denominatorSalesNok === 0 ? 0 : numeratorSalesNok / denominatorSalesNok,
  };
}

function buildGeneratedResults(restaurants, groups) {
  const output = [];
  for (const importId of ["IMP-CURRENT", "IMP-COMPARE"]) {
    for (const reportingGroupId of groups) {
      let companyNumerator = 0;
      let companyDenominator = 0;
      for (let index = 0; index < restaurants.length; index++) {
        const numeratorSalesNok = (index + 1) * 10;
        const denominatorSalesNok = (index + 1) * 100;
        companyNumerator += numeratorSalesNok;
        companyDenominator += denominatorSalesNok;
        output.push(resultRow({
          importId,
          reportingGroupId,
          restaurantId: restaurants[index],
          scopeType: "Restaurant",
          numeratorSalesNok,
          denominatorSalesNok,
        }));
      }
      output.push(resultRow({
        importId,
        reportingGroupId,
        restaurantId: "",
        scopeType: "Company",
        numeratorSalesNok: companyNumerator,
        denominatorSalesNok: companyDenominator,
      }));
    }
  }
  return output;
}
