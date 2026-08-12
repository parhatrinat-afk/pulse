/**
 * Pulse Build 0.3.0 Phase 2C — Interactive Sales Performance Model.
 *
 * Prerequisite sequence:
 * 1. Run the accepted Phase 1 mapping refresh.
 * 2. Run the accepted Phase 2A bridge refresh.
 * 3. Run the accepted Phase 2B centralized metric refresh.
 * 4. Run this script once to install/refill the formula-driven interaction layer.
 *
 * Normal Performance exploration after this script completes is worksheet
 * recalculation only. Restaurant Include, Reporting Group Include, dataset,
 * comparison, display, and sorting changes do not invoke Office Scripts.
 */
function main(workbook: ExcelScript.Workbook): string {
  const factsTable = requiredTable(workbook, "tblSalesFacts");
  const bridgeTable = requiredTable(workbook, "tblMetricRPGFacts");
  const resultsTable = requiredTable(workbook, "tblMetricRPGResults");
  const importsTable = requiredTable(workbook, "tblImports");
  const groupsTable = requiredTable(workbook, "tblReportingGroups");
  const restaurantsTable = requiredTable(workbook, "tblRestaurants");
  const rulesTable = requiredTable(workbook, "tblMappingRules");
  const effectiveTable = requiredTable(workbook, "tblEffectiveMapping");
  const kpiTable = requiredTable(workbook, "tblKPIRegistry");
  const phase2BQaTable = requiredTable(workbook, "tblMetricResultsQA");
  const environmentTable = requiredTable(workbook, "tblEnvironment");
  const buildLogTable = requiredTable(workbook, "tblBuildLog");

  const calc = requiredSheet(workbook, "_Metric_Calc");
  const performance = requiredSheet(workbook, "Performance");
  const reports = requiredSheet(workbook, "Reports");
  const overview = requiredSheet(workbook, "Overview");
  const qaSheet = requiredSheet(workbook, "Metric Results QA");

  // Read and validate the complete Phase 2B result/configuration state before
  // any workbook write. Phase 2C never reconstructs mapping or metric results.
  const activeImports = readActiveImports(importsTable);
  const activeGroups = readActiveGroups(groupsTable);
  const reportingRestaurants = readReportingRestaurants(restaurantsTable);
  const results = readMetricResults(resultsTable);
  if (!activeImports.length) throw new Error("PUL-0302C-001: No Active Finalized dataset is available.");
  if (!activeGroups.length) throw new Error("PUL-0302C-002: No active Reporting Group is available.");
  if (!reportingRestaurants.length) throw new Error("PUL-0302C-003: No active ReportingEnabled restaurant is available.");

  const labelErrors = validateSelectorLabels(activeImports, activeGroups);
  const phase2BQaErrors = validatePhase2BQA(phase2BQaTable);
  const resultValidation = validateCentralResultsForInteraction(
    results,
    activeImports,
    activeGroups,
    reportingRestaurants
  );
  const helperErrors = validatePhase2BHelpers(calc, activeImports, activeGroups, reportingRestaurants);
  const preflightErrors = labelErrors.concat(phase2BQaErrors, resultValidation.errors, helperErrors);
  if (preflightErrors.length) {
    throw new Error(
      "PUL-0302C-004: Phase 2B centralized results are stale or inconsistent. " +
      "Rerun Phase 2B, then Phase 2C. " + preflightErrors.slice(0, 12).join(" | ")
    );
  }

  const priorState = capturePriorUiState(workbook, calc, performance, activeImports, activeGroups);
  const restaurantLabels = uniqueDisplayLabels(reportingRestaurants);
  const plannedRestaurantRows = planSelectionRows(
    restaurantLabels,
    priorState.restaurantCatalog,
    "RestaurantID"
  );
  const plannedGroupRows = planSelectionRows(
    activeGroups.map(value => ({ id: value.id, name: value.name })),
    priorState.groupCatalog,
    "ReportingGroupID"
  );
  const layout = makeLayout(activeGroups.length, reportingRestaurants.length);

  // Batched protected-state reads are outside all loops. Phase 2C writes only
  // UI/helper/QA/environment surfaces.
  const protectedBefore = protectedStateFingerprint(
    factsTable.getRange().getValues(),
    bridgeTable.getRange().getValues(),
    resultsTable.getRange().getValues(),
    importsTable.getRange().getValues(),
    groupsTable.getRange().getValues(),
    restaurantsTable.getRange().getValues(),
    rulesTable.getRange().getValues(),
    effectiveTable.getRange().getValues(),
    kpiTable.getRange().getValues(),
    phase2BQaTable.getRange().getValues()
  );

  // All validation above must succeed before presentation mutation begins.
  workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.automatic);
  const validationMessage = writePerformance(
    workbook,
    performance,
    activeImports,
    activeGroups,
    plannedRestaurantRows,
    plannedGroupRows,
    priorState,
    layout
  );
  writeInteractionCalc(
    calc,
    activeImports,
    activeGroups,
    priorState,
    layout,
    resultValidation.companyFingerprint
  );
  writeReports(reports);
  overview.getRange("A20").setValue("Build 0.3.0 Phase 2C · Interactive Sales Performance");
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  const protectedAfter = protectedStateFingerprint(
    factsTable.getRange().getValues(),
    bridgeTable.getRange().getValues(),
    resultsTable.getRange().getValues(),
    importsTable.getRange().getValues(),
    groupsTable.getRange().getValues(),
    restaurantsTable.getRange().getValues(),
    rulesTable.getRange().getValues(),
    effectiveTable.getRange().getValues(),
    kpiTable.getRange().getValues(),
    phase2BQaTable.getRange().getValues()
  );
  if (protectedBefore !== protectedAfter) {
    throw new Error(
      "PUL-0302C-005: A protected source, bridge, centralized result, configuration, mapping, KPI, or Phase 2B QA table changed during Phase 2C."
    );
  }

  writeInteractionQA(
    workbook,
    qaSheet,
    phase2BQaTable,
    activeImports,
    activeGroups,
    reportingRestaurants,
    results,
    resultValidation.companyFingerprint,
    validationMessage
  );
  const calculatedAt = excelNow();
  updateEnvironment(
    environmentTable,
    "BuildID",
    "0.3.0-Phase2C",
    "Formula-driven Include-state restaurant/RPG selection and interactive KPI-0001 matrix over accepted Phase 2B results."
  );
  updateEnvironment(
    environmentTable,
    "BuildVersion",
    "0.3.0-Phase2C",
    "Normal Sales Performance exploration requires Excel recalculation only."
  );
  appendBuildLog(buildLogTable, [
    nextId(buildLogTable, "LOG", 6),
    calculatedAt,
    "0.3.0-Phase2C",
    "Install Interactive Sales Performance Model",
    "Success",
    "Performance + Reports + _Metric_Calc helpers",
    `${activeImports.length} datasets; ${activeGroups.length} active RPGs; ` +
      `${reportingRestaurants.length} eligible restaurants; Company control ${resultValidation.companyFingerprint}. ${validationMessage}`
  ]);

  return (
    `Pulse 0.3.0 Phase 2C applied. ${reportingRestaurants.length} eligible restaurants, ` +
    `${activeGroups.length} active Reporting Groups, ${activeImports.length} datasets. ` +
    `Performance selections now recalculate without Office Scripts. ${validationMessage}`
  );
}

type ActiveImport = {
  id: string; label: string; start: number; end: number; days: number; year: number; week: string;
};
type ReportingGroup = { id: string; name: string; active: string; sortOrder: number };
type ReportingRestaurant = { id: string; name: string };
type MetricResult = {
  id: string; metricId: string; importId: string; groupId: string; scopeType: string;
  restaurantId: string; restaurantScopeFingerprint: string; channelScope: string; publicationState: string;
  numeratorSales: number; denominatorSales: number; metricValue: number;
  mappingAsOfDate: number; mappingFingerprint: string; calculatedAt: number;
};
type SelectionCatalog = {
  exists: boolean; includeById: { [key: string]: string };
};
type SelectionRow = { id: string; name: string; include: string };
type PriorUiState = {
  detailGroupId: string; currentImportId: string; comparisonImportId: string;
  matrixMode: string; sortGroupId: string; sortOrder: string;
  restaurantCatalog: SelectionCatalog; groupCatalog: SelectionCatalog;
  priorRestaurantCapacity: number; priorGroupCapacity: number;
};
type ComponentBlock = { startColumn: number; name: string };
type Layout = {
  restaurantCapacity: number; groupCapacity: number; selectedRestaurantStartColumn: number;
  selectedGroupStartColumn: number; controlStartColumn: number; componentBlocks: ComponentBlock[];
  numericDisplayStartColumn: number; totalComponentStartColumn: number; totalDisplayColumn: number;
  sortKeyColumn: number; sortedRestaurantIdColumn: number; helperLastColumn: number;
  componentTotalRow: number; selectionStartColumn: number;
  groupSelectionStartColumn: number; matrixHeaderRow: number; matrixBodyStartRow: number;
  matrixEndColumn: number; matrixEndRow: number; explainStartRow: number; performanceLastColumn: number;
};
type ResultValidation = { errors: string[]; companyFingerprint: string };
type HashState = { left: number; right: number };

const METRIC_ID = "KPI-0001";
const METRIC_NAME = "Reporting Group Sales Share";
const ALL_CHANNELS = "All channels";
const ACTIVE_FINALIZED = "Active Finalized";
const DASH = "—";
const SORT_ORDERS = ["Highest first", "Lowest first"];
const MATRIX_MODES = [
  "PP Change", "Current Share", "Comparison Share", "Current Sales NOK", "NOK Impact"
];

function readActiveImports(table: ExcelScript.Table): ActiveImport[] {
  const h = requiredHeaderMap(table, [
    "ImportID", "PublicationState", "ActiveVersion", "PeriodStart", "PeriodEnd", "ReportingYear", "ReportingWeek"
  ]);
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  const output: ActiveImport[] = [];
  const seen: { [key: string]: boolean } = {};
  for (const row of values) {
    const id = text(row[h.ImportID]);
    if (!id || text(row[h.PublicationState]) !== ACTIVE_FINALIZED || text(row[h.ActiveVersion]) !== "Yes") continue;
    if (seen[id]) throw new Error(`PUL-0302C-006: Duplicate active ImportID ${id}.`);
    seen[id] = true;
    const start = numberValue(row[h.PeriodStart]);
    const end = numberValue(row[h.PeriodEnd]);
    const year = numberValue(row[h.ReportingYear]);
    const week = text(row[h.ReportingWeek]);
    output.push({
      id,
      label: makePeriodLabel(start, end, year, week),
      start,
      end,
      days: Math.round(end - start + 1),
      year,
      week
    });
  }
  output.sort((left, right) => left.start - right.start || left.end - right.end || left.id.localeCompare(right.id));
  return output;
}

function readActiveGroups(table: ExcelScript.Table): ReportingGroup[] {
  const h = requiredHeaderMap(table, ["ReportingGroupID", "ReportingGroupName", "Active", "SortOrder"]);
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  const output: ReportingGroup[] = [];
  const seen: { [key: string]: boolean } = {};
  for (const row of values) {
    const id = text(row[h.ReportingGroupID]);
    if (!id) continue;
    if (seen[id]) throw new Error(`PUL-0302C-007: Duplicate ReportingGroupID ${id}.`);
    seen[id] = true;
    if (text(row[h.Active]) !== "Yes") continue;
    output.push({
      id,
      name: text(row[h.ReportingGroupName]) || id,
      active: "Yes",
      sortOrder: numberValue(row[h.SortOrder])
    });
  }
  output.sort((left, right) =>
    left.sortOrder - right.sortOrder || left.name.localeCompare(right.name) || left.id.localeCompare(right.id)
  );
  return output;
}

function readReportingRestaurants(table: ExcelScript.Table): ReportingRestaurant[] {
  const h = requiredHeaderMap(table, ["RestaurantID", "Status", "ReportingEnabled", "DisplayName", "SourceRestaurantName"]);
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  const output: ReportingRestaurant[] = [];
  const seen: { [key: string]: boolean } = {};
  for (const row of values) {
    const id = text(row[h.RestaurantID]);
    if (!id || text(row[h.Status]) !== "Active" || text(row[h.ReportingEnabled]) !== "Yes") continue;
    if (seen[id]) throw new Error(`PUL-0302C-008: Duplicate eligible RestaurantID ${id}.`);
    seen[id] = true;
    output.push({ id, name: text(row[h.DisplayName]) || text(row[h.SourceRestaurantName]) || id });
  }
  output.sort((left, right) => left.name.localeCompare(right.name) || left.id.localeCompare(right.id));
  return output;
}

function readMetricResults(table: ExcelScript.Table): MetricResult[] {
  const h = requiredHeaderMap(table, [
    "MetricResultID", "MetricID", "ImportID", "ReportingGroupID", "ScopeType", "RestaurantID",
    "RestaurantScopeFingerprint", "ChannelScope", "PublicationState", "NumeratorSalesNOK",
    "DenominatorSalesNOK", "MetricValue", "MappingAsOfDate", "MappingFingerprint", "CalculatedAt"
  ]);
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  return values.map(row => ({
    id: text(row[h.MetricResultID]),
    metricId: text(row[h.MetricID]),
    importId: text(row[h.ImportID]),
    groupId: text(row[h.ReportingGroupID]),
    scopeType: text(row[h.ScopeType]),
    restaurantId: text(row[h.RestaurantID]),
    restaurantScopeFingerprint: text(row[h.RestaurantScopeFingerprint]),
    channelScope: text(row[h.ChannelScope]),
    publicationState: text(row[h.PublicationState]),
    numeratorSales: numberValue(row[h.NumeratorSalesNOK]),
    denominatorSales: numberValue(row[h.DenominatorSalesNOK]),
    metricValue: numberValue(row[h.MetricValue]),
    mappingAsOfDate: numberValue(row[h.MappingAsOfDate]),
    mappingFingerprint: text(row[h.MappingFingerprint]),
    calculatedAt: numberValue(row[h.CalculatedAt])
  }));
}

function validateSelectorLabels(imports: ActiveImport[], groups: ReportingGroup[]): string[] {
  const errors: string[] = [];
  const importLabels: { [key: string]: string } = {};
  const groupLabels: { [key: string]: string } = {};
  for (const value of imports) {
    if (importLabels[value.label]) errors.push(`dataset display label ${value.label} is shared by ${importLabels[value.label]} and ${value.id}.`);
    importLabels[value.label] = value.id;
  }
  for (const value of groups) {
    if (groupLabels[value.name]) errors.push(`active Reporting Group name ${value.name} is shared by ${groupLabels[value.name]} and ${value.id}.`);
    groupLabels[value.name] = value.id;
  }
  return errors;
}

function validatePhase2BQA(table: ExcelScript.Table): string[] {
  const h = requiredHeaderMap(table, ["CheckID", "Result"]);
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  const errors: string[] = [];
  let phase2BChecks = 0;
  for (const row of values) {
    const checkId = text(row[h.CheckID]);
    if (checkId.indexOf("QA-0302B-") !== 0) continue;
    phase2BChecks += 1;
    if (text(row[h.Result]) === "FAIL") errors.push(`${checkId} is FAIL.`);
  }
  if (phase2BChecks < 16) errors.push(`Phase 2B QA contains ${phase2BChecks} checks; expected at least 16.`);
  return errors;
}

function validateCentralResultsForInteraction(
  results: MetricResult[],
  imports: ActiveImport[],
  groups: ReportingGroup[],
  restaurants: ReportingRestaurant[]
): ResultValidation {
  const errors: string[] = [];
  const expectedRows = imports.length * groups.length * (1 + restaurants.length);
  const importIds: { [key: string]: boolean } = {};
  const groupIds: { [key: string]: boolean } = {};
  const restaurantIds: { [key: string]: boolean } = {};
  for (const value of imports) importIds[value.id] = true;
  for (const value of groups) groupIds[value.id] = true;
  for (const value of restaurants) restaurantIds[value.id] = true;
  const companyFingerprint = deterministicRestaurantScopeFingerprint(restaurants);
  const byKey: { [key: string]: MetricResult } = {};
  const seenResultIds: { [key: string]: boolean } = {};
  let mappingFingerprint = "";
  let mappingAsOfDate = -1;

  if (results.length !== expectedRows) errors.push(`central result row count ${results.length}; expected ${expectedRows}.`);
  for (const row of results) {
    if (!row.id) errors.push("central result contains a blank MetricResultID.");
    else if (seenResultIds[row.id]) errors.push(`duplicate MetricResultID ${row.id}.`);
    else seenResultIds[row.id] = true;
    if (row.metricId !== METRIC_ID) errors.push(`${row.id || "(blank)"} has MetricID ${row.metricId || "(blank)"}.`);
    if (!importIds[row.importId]) errors.push(`${row.id} references inactive/unknown ImportID ${row.importId}.`);
    if (!groupIds[row.groupId]) errors.push(`${row.id} references inactive/unknown ReportingGroupID ${row.groupId}.`);
    if (row.channelScope !== ALL_CHANNELS || row.publicationState !== ACTIVE_FINALIZED) {
      errors.push(`${row.id} has unsupported channel/publication scope.`);
    }
    if (row.scopeType !== "Company" && row.scopeType !== "Restaurant") errors.push(`${row.id} has unsupported ScopeType ${row.scopeType}.`);
    if (row.scopeType === "Company") {
      if (row.restaurantId) errors.push(`${row.id} Company row has RestaurantID ${row.restaurantId}.`);
      if (row.restaurantScopeFingerprint !== companyFingerprint) errors.push(`${row.id} Company scope fingerprint differs.`);
    } else {
      if (!restaurantIds[row.restaurantId]) errors.push(`${row.id} references ineligible RestaurantID ${row.restaurantId}.`);
      if (row.restaurantScopeFingerprint) errors.push(`${row.id} Restaurant row has a Company scope fingerprint.`);
    }
    const expectedMetricValue = row.denominatorSales === 0 ? 0 : row.numeratorSales / row.denominatorSales;
    if (!almostEqual(row.metricValue, expectedMetricValue)) errors.push(`${row.id} canonical metric value differs from its components.`);
    const expectedId = deterministicMetricResultId(row, companyFingerprint);
    if (row.id !== expectedId) errors.push(`${row.id} is not the deterministic result ID for its grain.`);
    if (!mappingFingerprint) mappingFingerprint = row.mappingFingerprint;
    else if (row.mappingFingerprint !== mappingFingerprint) errors.push(`${row.id} has inconsistent MappingFingerprint.`);
    if (mappingAsOfDate < 0) mappingAsOfDate = row.mappingAsOfDate;
    else if (row.mappingAsOfDate !== mappingAsOfDate) errors.push(`${row.id} has inconsistent MappingAsOfDate.`);
    const key = metricResultKey(row.importId, row.groupId, row.scopeType, row.restaurantId);
    if (byKey[key]) errors.push(`duplicate metric result grain ${key}.`);
    else byKey[key] = row;
  }
  if (!mappingFingerprint) errors.push("central result MappingFingerprint is blank.");

  for (const activeImport of imports) {
    for (const group of groups) {
      const companyKey = metricResultKey(activeImport.id, group.id, "Company", "");
      const company = byKey[companyKey];
      if (!company) {
        errors.push(`missing Company result ${companyKey}.`);
        continue;
      }
      let numerator = 0;
      let denominator = 0;
      for (const restaurant of restaurants) {
        const restaurantKey = metricResultKey(activeImport.id, group.id, "Restaurant", restaurant.id);
        const result = byKey[restaurantKey];
        if (!result) {
          errors.push(`missing Restaurant result ${restaurantKey}.`);
          continue;
        }
        numerator += result.numeratorSales;
        denominator += result.denominatorSales;
      }
      const combinedMetric = denominator === 0 ? 0 : numerator / denominator;
      if (!almostEqual(numerator, company.numeratorSales) ||
          !almostEqual(denominator, company.denominatorSales) ||
          !almostEqual(combinedMetric, company.metricValue)) {
        errors.push(`${activeImport.id}/${group.id} Restaurant components do not reconcile to Company control.`);
      }
    }
    let companyDenominator = -1;
    for (const group of groups) {
      const company = byKey[metricResultKey(activeImport.id, group.id, "Company", "")];
      if (!company) continue;
      if (companyDenominator < 0) companyDenominator = company.denominatorSales;
      else if (!almostEqual(company.denominatorSales, companyDenominator)) {
        errors.push(`${activeImport.id} Company denominator differs across Reporting Groups.`);
      }
    }
    for (const restaurant of restaurants) {
      let restaurantDenominator = -1;
      for (const group of groups) {
        const result = byKey[metricResultKey(activeImport.id, group.id, "Restaurant", restaurant.id)];
        if (!result) continue;
        if (restaurantDenominator < 0) restaurantDenominator = result.denominatorSales;
        else if (!almostEqual(result.denominatorSales, restaurantDenominator)) {
          errors.push(`${activeImport.id}/${restaurant.id} denominator differs across Reporting Groups.`);
        }
      }
    }
  }
  return { errors, companyFingerprint };
}

function validatePhase2BHelpers(
  calc: ExcelScript.Worksheet,
  imports: ActiveImport[],
  groups: ReportingGroup[],
  restaurants: ReportingRestaurant[]
): string[] {
  const importValues = checkedRangeByIndexes(calc, "_Metric_Calc", 1, 0, imports.length, 2, "read dataset helpers").getValues();
  const groupValues = checkedRangeByIndexes(calc, "_Metric_Calc", 1, 8, groups.length, 2, "read Reporting Group helpers").getValues();
  const restaurantValues = checkedRangeByIndexes(calc, "_Metric_Calc", 1, 11, restaurants.length, 2, "read restaurant helpers").getValues();
  const errors: string[] = [];
  for (let index = 0; index < imports.length; index++) {
    if (text(importValues[index][0]) !== imports[index].label || text(importValues[index][1]) !== imports[index].id) {
      errors.push(`dataset helper row ${index + 2} differs from active import ${imports[index].id}.`);
    }
  }
  for (let index = 0; index < groups.length; index++) {
    if (text(groupValues[index][0]) !== groups[index].name || text(groupValues[index][1]) !== groups[index].id) {
      errors.push(`Reporting Group helper row ${index + 2} differs from active group ${groups[index].id}.`);
    }
  }
  for (let index = 0; index < restaurants.length; index++) {
    if (text(restaurantValues[index][0]) !== restaurants[index].name || text(restaurantValues[index][1]) !== restaurants[index].id) {
      errors.push(`restaurant helper row ${index + 2} differs from eligible restaurant ${restaurants[index].id}.`);
    }
  }
  return errors;
}

function capturePriorUiState(
  workbook: ExcelScript.Workbook,
  calc: ExcelScript.Worksheet,
  performance: ExcelScript.Worksheet,
  imports: ActiveImport[],
  groups: ReportingGroup[]
): PriorUiState {
  const leftValues = performance.getRange("B7:B10").getValues();
  const rightValues = performance.getRange("G6:G10").getValues();
  const sortValues = performance.getRange("I6:I7").getValues();
  const priorCapacityValues = calc.getRange("AL10:AL11").getValues();
  const priorSortGroupId = text(calc.getRange("AL12").getValue());
  const detailName = text(leftValues[0][0]);
  const currentLabel = text(leftValues[3][0]);
  const matrixModeValue = text(rightValues[0][0]);
  const comparisonLabel = text(rightValues[4][0]);
  const visibleSortName = text(sortValues[0][0]);
  const visibleSortOrder = text(sortValues[1][0]);
  const sortGroupId = groups.some(value => value.id === priorSortGroupId)
    ? priorSortGroupId
    : groupIdByName(groups, visibleSortName);
  const restaurantCatalog = captureSelectionCatalog(
    workbook.getTable("tblPerformanceRestaurantSelection"),
    "RestaurantID"
  );
  const groupCatalog = captureSelectionCatalog(
    workbook.getTable("tblPerformanceRPGSelection"),
    "ReportingGroupID"
  );
  // One-time compatibility migration from the prior visible All/Custom
  // controls. If All was effective, preserve that exact scope by making every
  // existing Include row authoritative Yes before those mode cells disappear.
  applyLegacyAllMode(restaurantCatalog, text(rightValues[1][0]));
  applyLegacyAllMode(groupCatalog, text(rightValues[2][0]));
  return {
    detailGroupId: groupIdByName(groups, detailName),
    currentImportId: importIdByLabel(imports, currentLabel),
    comparisonImportId: importIdByLabel(imports, comparisonLabel),
    matrixMode: MATRIX_MODES.indexOf(matrixModeValue) >= 0 ? matrixModeValue : MATRIX_MODES[0],
    sortGroupId,
    sortOrder: SORT_ORDERS.indexOf(visibleSortOrder) >= 0 ? visibleSortOrder : SORT_ORDERS[0],
    restaurantCatalog,
    groupCatalog,
    priorRestaurantCapacity: Math.max(0, Math.floor(numberValue(priorCapacityValues[0][0]))),
    priorGroupCapacity: Math.max(0, Math.floor(numberValue(priorCapacityValues[1][0])))
  };
}

function applyLegacyAllMode(catalog: SelectionCatalog, visibleMode: string): void {
  if (visibleMode !== "All") return;
  for (const id in catalog.includeById) catalog.includeById[id] = "Yes";
}

function captureSelectionCatalog(
  table: ExcelScript.Table | undefined,
  idHeader: string
): SelectionCatalog {
  if (!table) return { exists: false, includeById: {} };
  const h = requiredHeaderMap(table, ["Include", idHeader]);
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  const includeById: { [key: string]: string } = {};
  for (const row of values) {
    const id = text(row[h[idHeader]]);
    if (!id) throw new Error(`PUL-0302C-009: Existing selection table contains a blank ${idHeader}.`);
    if (includeById[id] !== undefined) throw new Error(`PUL-0302C-010: Existing selection table repeats ${idHeader} ${id}.`);
    includeById[id] = text(row[h.Include]) === "Yes" ? "Yes" : "No";
  }
  return { exists: true, includeById };
}

function planSelectionRows(
  eligibleItems: { id: string; name: string }[],
  prior: SelectionCatalog,
  idLabel: string
): SelectionRow[] {
  const output: SelectionRow[] = [];
  const seen: { [key: string]: boolean } = {};
  for (const item of eligibleItems) {
    if (!item.id) throw new Error(`PUL-0302C-011: Eligible ${idLabel} is blank.`);
    if (seen[item.id]) throw new Error(`PUL-0302C-012: Eligible ${idLabel} repeats ${item.id}.`);
    seen[item.id] = true;
    output.push({
      id: item.id,
      name: item.name || item.id,
      include: prior.includeById[item.id] !== undefined
        ? prior.includeById[item.id]
        : prior.exists ? "No" : "Yes"
    });
  }
  return output;
}

function uniqueDisplayLabels(restaurants: ReportingRestaurant[]): { id: string; name: string }[] {
  const counts: { [key: string]: number } = {};
  for (const restaurant of restaurants) counts[restaurant.name] = (counts[restaurant.name] || 0) + 1;
  return restaurants.map(restaurant => ({
    id: restaurant.id,
    name: counts[restaurant.name] > 1 ? `${restaurant.name} [${restaurant.id}]` : restaurant.name
  }));
}

function makeLayout(groupCapacity: number, restaurantCapacity: number): Layout {
  const selectedRestaurantStartColumn = 30; // AE:AF
  const selectedGroupStartColumn = 33; // AH:AI
  const controlStartColumn = 36; // AK:AL
  const componentStartColumn = 39; // AN
  const componentNames = [
    "Current Numerator", "Current Denominator", "Current Share",
    "Comparison Numerator", "Comparison Denominator", "Comparison Share"
  ];
  const componentBlocks: ComponentBlock[] = [];
  for (let index = 0; index < componentNames.length; index++) {
    componentBlocks.push({
      name: componentNames[index],
      startColumn: componentStartColumn + index * (groupCapacity + 1)
    });
  }
  const numericDisplayStartColumn =
    componentBlocks[componentBlocks.length - 1].startColumn + groupCapacity + 1;
  const totalComponentStartColumn = numericDisplayStartColumn + groupCapacity + 1;
  const totalDisplayColumn = totalComponentStartColumn + 4;
  const sortKeyColumn = totalDisplayColumn + 1;
  const sortedRestaurantIdColumn = sortKeyColumn + 1;
  const helperLastColumn = sortedRestaurantIdColumn;
  const selectionStartColumn = Math.max(13, groupCapacity + 3);
  const groupSelectionStartColumn = selectionStartColumn + 4;
  const matrixHeaderRow = 22;
  const matrixBodyStartRow = 23;
  const matrixEndColumn = groupCapacity + 1;
  const matrixEndRow = matrixBodyStartRow + restaurantCapacity;
  const explainStartRow = matrixEndRow + 3;
  const performanceLastColumn = Math.max(11, matrixEndColumn, groupSelectionStartColumn + 2);
  return {
    restaurantCapacity,
    groupCapacity,
    selectedRestaurantStartColumn,
    selectedGroupStartColumn,
    controlStartColumn,
    componentBlocks,
    numericDisplayStartColumn,
    totalComponentStartColumn,
    totalDisplayColumn,
    sortKeyColumn,
    sortedRestaurantIdColumn,
    helperLastColumn,
    componentTotalRow: restaurantCapacity + 1,
    selectionStartColumn,
    groupSelectionStartColumn,
    matrixHeaderRow,
    matrixBodyStartRow,
    matrixEndColumn,
    matrixEndRow,
    explainStartRow,
    performanceLastColumn
  };
}

function writePerformance(
  workbook: ExcelScript.Workbook,
  performance: ExcelScript.Worksheet,
  activeImports: ActiveImport[],
  activeGroups: ReportingGroup[],
  restaurantRows: SelectionRow[],
  groupRows: SelectionRow[],
  prior: PriorUiState,
  layout: Layout
): string {
  const NAVY = "#172033";
  const BLUE = "#4F8CFF";
  const LIGHT = "#EAF2FF";
  const GREY = "#EEF1F5";
  const WHITE = "#FFFFFF";
  const MUTED = "#5B6677";
  const RED = "#A83126";

  const existingRestaurantSelection = workbook.getTable("tblPerformanceRestaurantSelection");
  const existingGroupSelection = workbook.getTable("tblPerformanceRPGSelection");
  if (existingRestaurantSelection) existingRestaurantSelection.delete();
  if (existingGroupSelection) existingGroupSelection.delete();
  const priorLayout = makeLayout(
    Math.max(1, prior.priorGroupCapacity),
    Math.max(1, prior.priorRestaurantCapacity)
  );
  const clearColumns = Math.max(layout.performanceLastColumn + 1, priorLayout.performanceLastColumn + 1, 24);
  const clearRows = Math.max(layout.explainStartRow + 10, priorLayout.explainStartRow + 10, 80);
  checkedRangeByIndexes(
    performance,
    "Performance",
    0,
    0,
    clearRows,
    clearColumns,
    "clear Phase 2C presentation"
  ).clear(ExcelScript.ClearApplyTo.all);

  const lastColumn = columnName(layout.performanceLastColumn + 1);
  writeTitle(
    performance,
    "Performance",
    "Interactive Reporting Group Sales Share — restaurant, Reporting Group, and dataset selections recalculate without Office Scripts.",
    lastColumn,
    NAVY,
    LIGHT,
    WHITE
  );

  const detailGroup = activeGroups.find(value => value.id === prior.detailGroupId) || activeGroups[0];
  const latest = activeImports[activeImports.length - 1];
  const fallbackComparison = activeImports.length > 1 ? activeImports[activeImports.length - 2] : latest;
  const current = activeImports.find(value => value.id === prior.currentImportId) || latest;
  const comparison = activeImports.find(value => value.id === prior.comparisonImportId) || fallbackComparison;
  const sortGroup = activeGroups.find(value => value.id === prior.sortGroupId);
  const sortBy = sortGroup ? sortGroup.name : "Total";

  performance.getRange("A5:D5").setValues([["Metric detail", "", "", ""]]);
  styleSection(performance.getRange("A5:D5"), BLUE, WHITE);
  performance.getRange("A6:B7").setValues([
    ["KPI", METRIC_NAME],
    ["Detail Reporting Group", detailGroup.name]
  ]);
  styleLabels(performance.getRange("A6:A7"), GREY);

  performance.getRange("F5:I5").setValues([["Interactive matrix", "", "", ""]]);
  styleSection(performance.getRange("F5:I5"), BLUE, WHITE);
  performance.getRange("F6:G8").setValues([
    ["Display", prior.matrixMode],
    ["Restaurants", ""],
    ["Reporting Groups", ""]
  ]);
  styleLabels(performance.getRange("F6:F8"), GREY);
  performance.getRange("H6:I8").setValues([
    ["Sort by", sortBy],
    ["Order", prior.sortOrder],
    ["Sort status", ""]
  ]);
  styleLabels(performance.getRange("H6:H8"), GREY);
  performance.getRange("G7").setFormula(
    `=IF('_Metric_Calc'!$AL$4=${restaurantRows.length},"All ${restaurantRows.length} eligible restaurants",` +
      `'_Metric_Calc'!$AL$4&" of ${restaurantRows.length} restaurants selected")`
  );
  performance.getRange("G8").setFormula(
    `=IF('_Metric_Calc'!$AL$5=${groupRows.length},"All ${groupRows.length} Reporting Groups selected",` +
      `'_Metric_Calc'!$AL$5&" of ${groupRows.length} Reporting Groups selected")`
  );
  performance.getRange("I8").setFormula("='_Metric_Calc'!$AL$15");
  performance.getRange("G7:I8").getFormat().setWrapText(true);

  performance.getRange("A9:B9").setValues([["Current", ""]]);
  styleSection(performance.getRange("A9:B9"), BLUE, WHITE);
  performance.getRange("A10:B10").setValues([["Dataset", current.label]]);
  styleLabels(performance.getRange("A10:A10"), GREY);
  performance.getRange("F9:G9").setValues([["Compare with", ""]]);
  styleSection(performance.getRange("F9:G9"), BLUE, WHITE);
  performance.getRange("F10:G10").setValues([["Dataset", comparison.label]]);
  styleLabels(performance.getRange("F10:F10"), GREY);

  performance.getRange("A15:D15").setValues([["Selected Reporting Group result", "", "", ""]]);
  styleSection(performance.getRange("A15:D15"), NAVY, WHITE);
  performance.getRange("A16:B19").setValues([
    ["Current share", ""],
    ["Comparison share", ""],
    ["Change", ""],
    ["Current Reporting Group sales", ""]
  ]);
  styleLabels(performance.getRange("A16:A19"), GREY);

  const currentNumerator = componentTotalExpression(layout, 0);
  const currentDenominator = componentTotalExpression(layout, 1);
  const currentShare = componentTotalExpression(layout, 2);
  const comparisonDenominator = componentTotalExpression(layout, 4);
  const comparisonShare = componentTotalExpression(layout, 5);
  const detailPosition = `MATCH('_Metric_Calc'!$AL$6,'_Metric_Calc'!$J$2:$J$${layout.groupCapacity + 1},0)`;
  performance.getRange("B16").setFormula(
    `=IFERROR(IF(INDEX(${currentDenominator},1,${detailPosition})=0,"${DASH}",INDEX(${currentShare},1,${detailPosition})),"${DASH}")`
  );
  performance.getRange("B17").setFormula(
    `=IFERROR(IF(INDEX(${comparisonDenominator},1,${detailPosition})=0,"${DASH}",INDEX(${comparisonShare},1,${detailPosition})),"${DASH}")`
  );
  performance.getRange("B18").setFormula(
    `=IF(OR(B16="${DASH}",B17="${DASH}"),"${DASH}",(B16-B17)*100)`
  );
  performance.getRange("B19").setFormula(
    `=IFERROR(IF(INDEX(${currentDenominator},1,${detailPosition})=0,"${DASH}",INDEX(${currentNumerator},1,${detailPosition})),"${DASH}")`
  );
  performance.getRange("B16:B17").setNumberFormat("0.00%");
  performance.getRange("B18").setNumberFormat('+0.00 "pp";-0.00 "pp";0.00 "pp"');
  performance.getRange("B19").setNumberFormat('#,##0 "NOK"');
  const negativeDetail = performance.getRange("B18").addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
  negativeDetail.getCustom().getRule().setFormula('=AND(ISNUMBER(B18),B18<0)');
  negativeDetail.getCustom().getFormat().getFont().setColor(RED);

  performance.getRange("F15:I15").setValues([["Comparison and scope check", "", "", ""]]);
  styleSection(performance.getRange("F15:I15"), NAVY, WHITE);
  performance.getRange("F16:G19").setValues([
    ["Period length", ""],
    ["Same dataset", ""],
    ["Status", ""],
    ["Restaurant scope", ""]
  ]);
  styleLabels(performance.getRange("F16:F19"), GREY);
  performance.getRange("G16").setFormula(
    `=IF(INDEX('_Metric_Calc'!$E$2:$E$${activeImports.length + 1},MATCH($B$10,'_Metric_Calc'!$A$2:$A$${activeImports.length + 1},0))=` +
      `INDEX('_Metric_Calc'!$E$2:$E$${activeImports.length + 1},MATCH($G$10,'_Metric_Calc'!$A$2:$A$${activeImports.length + 1},0)),"Same","Different")`
  );
  performance.getRange("G17").setFormula('=IF($B$10=$G$10,"Yes","No")');
  performance.getRange("G18").setFormula(
    '=IF(G16="Same","Same period length","Different period lengths — comparison allowed")'
  );
  performance.getRange("G19").setFormula("=$G$7");
  performance.getRange("G18:I19").getFormat().setWrapText(true);

  const restaurantTable = writeSelectionTable(
    performance,
    layout.selectionStartColumn,
    "Restaurant selection",
    ["Include", "Restaurant", "RestaurantID"],
    restaurantRows.map(value => [value.include, value.name, value.id]),
    "tblPerformanceRestaurantSelection"
  );
  const groupTable = writeSelectionTable(
    performance,
    layout.groupSelectionStartColumn,
    "Reporting Group selection",
    ["Include", "Reporting Group", "ReportingGroupID"],
    groupRows.map(value => [value.include, value.name, value.id]),
    "tblPerformanceRPGSelection"
  );

  performance.getRangeByIndexes(20, 0, 1, layout.matrixEndColumn + 1).setValues([
    rowWithLength("Interactive Sales Performance", layout.matrixEndColumn + 1)
  ]);
  styleSection(performance.getRangeByIndexes(20, 0, 1, layout.matrixEndColumn + 1), BLUE, WHITE);
  performance.getRangeByIndexes(21, 0, 1, layout.matrixEndColumn + 1).setValues([
    rowWithLength(
      "Sorting uses numeric values; Grand Total remains fixed.",
      layout.matrixEndColumn + 1
    )
  ]);
  performance.getRangeByIndexes(21, 0, 1, layout.matrixEndColumn + 1).getFormat().setWrapText(false);
  performance.getRangeByIndexes(layout.matrixHeaderRow, 0, 1, layout.matrixEndColumn + 1).setValues([
    ["Restaurant", "Total"].concat(new Array(layout.groupCapacity).fill(""))
  ]);
  const matrixHeaderFormulas: string[][] = [[]];
  for (let groupIndex = 0; groupIndex < layout.groupCapacity; groupIndex++) {
    const visibleColumn = columnName(groupIndex + 3);
    matrixHeaderFormulas[0].push(
      `=IF(COLUMNS($C$23:${visibleColumn}$23)<='_Metric_Calc'!$AL$5,` +
        `INDEX('_Metric_Calc'!$AH$2:$AH$${layout.groupCapacity + 1},COLUMNS($C$23:${visibleColumn}$23)),"")`
    );
  }
  performance.getRangeByIndexes(layout.matrixHeaderRow, 2, 1, layout.groupCapacity).setFormulas(matrixHeaderFormulas);
  styleHeader(performance.getRangeByIndexes(layout.matrixHeaderRow, 0, 1, layout.matrixEndColumn + 1), GREY, NAVY);

  const matrixRows = layout.restaurantCapacity + 1;
  const labelFormulas: string[][] = [];
  const valueFormulas: string[][] = [];
  for (let rowIndex = 0; rowIndex < matrixRows; rowIndex++) {
    const sheetRow = layout.matrixBodyStartRow + rowIndex + 1;
    const ordinal = `ROWS($A$${layout.matrixBodyStartRow + 1}:A${sheetRow})`;
    labelFormulas.push([
      `=IF(OR('_Metric_Calc'!$AL$4=0,'_Metric_Calc'!$AL$5=0),"",` +
        `IF(${ordinal}<='_Metric_Calc'!$AL$4,INDEX('_Metric_Calc'!$AE$2:$AE$${layout.restaurantCapacity + 1},` +
        `MATCH(INDEX(${sortedRestaurantIdRangeExpression(layout)},${ordinal}),'_Metric_Calc'!$AF$2:$AF$${layout.restaurantCapacity + 1},0)),` +
        `IF(${ordinal}='_Metric_Calc'!$AL$4+1,"Grand Total","")))`
    ]);
    const formulaRow: string[] = [matrixTotalPresentationFormula(layout, sheetRow, ordinal)];
    for (let groupIndex = 0; groupIndex < layout.groupCapacity; groupIndex++) {
      const sheetColumn = columnName(groupIndex + 3);
      formulaRow.push(matrixPresentationFormula(layout, sheetRow, sheetColumn, ordinal));
    }
    valueFormulas.push(formulaRow);
  }
  performance.getRangeByIndexes(layout.matrixBodyStartRow, 0, matrixRows, 1).setFormulas(labelFormulas);
  const matrixValues = performance.getRangeByIndexes(
    layout.matrixBodyStartRow,
    1,
    matrixRows,
    layout.groupCapacity + 1
  );
  matrixValues.setFormulas(valueFormulas);
  matrixValues.setNumberFormat("General");
  matrixValues.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  applyMatrixConditionalFormats(performance, matrixValues, layout, RED);
  const grandTotalFormat = performance.getRangeByIndexes(
    layout.matrixBodyStartRow,
    0,
    matrixRows,
    layout.matrixEndColumn + 1
  ).addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
  grandTotalFormat.getCustom().getRule().setFormula(`=$A${layout.matrixBodyStartRow + 1}="Grand Total"`);
  grandTotalFormat.getCustom().getFormat().getFill().setColor(LIGHT);
  grandTotalFormat.getCustom().getFormat().getFont().setBold(true);

  const explainRow = layout.explainStartRow;
  performance.getRangeByIndexes(explainRow, 0, 1, Math.min(8, layout.performanceLastColumn + 1)).setValues([
    rowWithLength("Explain", Math.min(8, layout.performanceLastColumn + 1))
  ]);
  styleSection(
    performance.getRangeByIndexes(explainRow, 0, 1, Math.min(8, layout.performanceLastColumn + 1)),
    BLUE,
    WHITE
  );
  performance.getRangeByIndexes(explainRow + 1, 0, 8, 8).setValues([
    ["Metric", METRIC_NAME, "", "", "", "", "", ""],
    ["Definition", "Mapped sales for a Reporting Group ÷ all sales inside the selected restaurant and dataset scope.", "", "", "", "", "", ""],
    ["Selections", "Include state is authoritative. Yes participates; No does not. All versus Custom is derived, not user-selected.", "", "", "", "", "", ""],
    ["Total", "Aggregates only the selected Reporting Groups. Denominators are used once; Current Sales NOK is the selected numerator without a denominator guard.", "", "", "", "", "", ""],
    ["Sorting", "Sort by Total or a displayed Reporting Group using full-precision numeric helpers. Grand Total never participates; unavailable values remain last.", "", "", "", "", "", ""],
    ["Grand Total", "Share/PP uses summed numerator ÷ summed denominator; NOK Impact applies the aggregated comparison share to aggregated current total sales.", "", "", "", "", "", ""],
    ["Guardrail", "Unmapped, Conflict, Inactive Target, and explicit exclusions remain in denominators. Zero denominator displays as —.", "", "", "", "", "", ""],
    ["Refresh boundary", "ReportingEnabled controls refresh-time eligibility. Interactive Include choices never alter facts, mapping, or Phase 2B results.", "", "", "", "", "", ""]
  ]);
  performance.getRangeByIndexes(explainRow + 1, 0, 8, 1).getFormat().getFont().setBold(true);
  performance.getRangeByIndexes(explainRow + 1, 0, 8, 1).getFormat().getFont().setColor(MUTED);
  performance.getRangeByIndexes(explainRow + 1, 1, 8, 7).getFormat().setWrapText(true);

  performance.getFreezePanes().freezeRows(2);
  performance.getRange("A:A").getFormat().setColumnWidth(165);
  performance.getRange("B:B").getFormat().setColumnWidth(150);
  if (layout.groupCapacity > 0) {
    performance.getRangeByIndexes(0, 2, 1, layout.groupCapacity).getEntireColumn().getFormat().setColumnWidth(105);
  }
  performance.getRange("F:F").getFormat().setColumnWidth(145);
  performance.getRange("G:G").getFormat().setColumnWidth(230);
  performance.getRangeByIndexes(0, layout.selectionStartColumn, 1, 3).getEntireColumn().getFormat().setColumnWidth(125);
  performance.getRangeByIndexes(0, layout.groupSelectionStartColumn, 1, 3).getEntireColumn().getFormat().setColumnWidth(125);
  performance.getRangeByIndexes(0, 0, explainRow + 9, 1).getEntireRow().getFormat().setRowHeight(20);
  performance.getRange("1:1").getFormat().setRowHeight(34);
  performance.getRange("2:2").getFormat().setRowHeight(30);
  performance.getRange("7:7").getFormat().setRowHeight(30);
  performance.getRange("8:8").getFormat().setRowHeight(45);
  performance.getRange("18:19").getFormat().setRowHeight(30);
  performance.getRange(`${layout.matrixHeaderRow + 1}:${layout.matrixHeaderRow + 1}`).getFormat().setRowHeight(30);

  const failures: string[] = [];
  applyListValidation(performance.getRange("B7"), activeGroups.map(value => value.name), "detail Reporting Group", failures);
  const datasetLabels = activeImports.map(value => value.label);
  applyListValidation(performance.getRange("B10"), datasetLabels, "current dataset", failures);
  applyListValidation(performance.getRange("G10"), datasetLabels, "comparison dataset", failures);
  applyListValidation(performance.getRange("G6"), MATRIX_MODES, "matrix display", failures);
  applyListValidation(performance.getRange("I6"), ["Total"].concat(activeGroups.map(value => value.name)), "sort target", failures);
  applyListValidation(performance.getRange("I7"), SORT_ORDERS, "sort order", failures);
  const restaurantIncludeColumn = restaurantTable.getColumnByName("Include");
  const groupIncludeColumn = groupTable.getColumnByName("Include");
  if (!restaurantIncludeColumn || !groupIncludeColumn) {
    throw new Error("PUL-0302C-019: Phase 2C selection table Include column is missing after creation.");
  }
  applyListValidation(
    restaurantIncludeColumn.getRangeBetweenHeaderAndTotal(),
    ["Yes", "No"],
    "restaurant Include",
    failures
  );
  applyListValidation(
    groupIncludeColumn.getRangeBetweenHeaderAndTotal(),
    ["Yes", "No"],
    "Reporting Group Include",
    failures
  );
  for (const address of ["B7", "B10", "G6", "G10", "I6", "I7"]) {
    performance.getRange(address).getFormat().getFill().setColor(WHITE);
    performance.getRange(address).getFormat().getFont().setBold(true);
  }
  return failures.length
    ? `PUL-0302C-013: ${failures.length} interactive dropdown(s) unavailable; formulas and centralized results remain intact. ${failures.join(" | ")}`
    : "Interactive dropdown validation ready (8/8).";
}

function writeSelectionTable(
  sheet: ExcelScript.Worksheet,
  startColumn: number,
  title: string,
  headers: string[],
  rows: (string | number | boolean)[][],
  tableName: string
): ExcelScript.Table {
  sheet.getRangeByIndexes(3, startColumn, 1, 3).setValues([[title, "", ""]]);
  styleSection(sheet.getRangeByIndexes(3, startColumn, 1, 3), "#172033", "#FFFFFF");
  sheet.getRangeByIndexes(4, startColumn, 1, 3).setValues([headers]);
  sheet.getRangeByIndexes(5, startColumn, rows.length, 3).setValues(rows);
  const table = sheet.addTable(sheet.getRangeByIndexes(4, startColumn, rows.length + 1, 3), true);
  table.setName(tableName);
  table.setPredefinedTableStyle("TableStyleMedium2");
  sheet.getRangeByIndexes(5, startColumn, rows.length, 1).getFormat().getFill().setColor("#FFF4D6");
  return table;
}

function matrixPresentationFormula(
  layout: Layout,
  sheetRow: number,
  sheetColumn: string,
  ordinal: string
): string {
  const groupId = `INDEX('_Metric_Calc'!$AI$2:$AI$${layout.groupCapacity + 1},COLUMNS($C$23:${sheetColumn}$23))`;
  const groupPosition = `MATCH(${groupId},'_Metric_Calc'!$J$2:$J$${layout.groupCapacity + 1},0)`;
  const componentRow = sortedComponentRowExpression(layout, ordinal);
  const numericValue = `INDEX(${numericDisplayRangeExpression(layout)},${componentRow},${groupPosition})`;
  return matrixFacadeFormula(sheetRow, sheetColumn, numericValue);
}

function matrixTotalPresentationFormula(layout: Layout, sheetRow: number, ordinal: string): string {
  const componentRow = sortedComponentRowExpression(layout, ordinal);
  const numericValue = `INDEX(${totalDisplayRangeExpression(layout)},${componentRow})`;
  return matrixFacadeFormula(sheetRow, "B", numericValue);
}

function sortedComponentRowExpression(layout: Layout, ordinal: string): string {
  const sortedId = `INDEX(${sortedRestaurantIdRangeExpression(layout)},${ordinal})`;
  return `IF(${ordinal}<='_Metric_Calc'!$AL$4,MATCH(${sortedId},'_Metric_Calc'!$AF$2:$AF$${layout.restaurantCapacity + 1},0),${layout.restaurantCapacity + 1})`;
}

function matrixFacadeFormula(sheetRow: number, sheetColumn: string, numericValue: string): string {
  return (
    `=IF(OR($A${sheetRow}="",${sheetColumn}$23=""),"",` +
    `IF(NOT(ISNUMBER(${numericValue})),"${DASH}",` +
    `IF($G$6="PP Change",IF(${numericValue}>0,"+","")&FIXED(${numericValue},2,TRUE)&" pp",` +
    `IF(OR($G$6="Current Share",$G$6="Comparison Share"),FIXED(${numericValue}*100,2,TRUE)&"%",` +
    `IF($G$6="Current Sales NOK",FIXED(${numericValue},0,FALSE)&" NOK",` +
    `IF($G$6="NOK Impact",IF(${numericValue}>0,"+","")&FIXED(${numericValue},0,FALSE)&" NOK",""))))))`
  );
}

function totalDisplayRangeExpression(layout: Layout): string {
  const column = columnName(layout.totalDisplayColumn + 1);
  return `'_Metric_Calc'!$${column}$2:$${column}$${layout.restaurantCapacity + 2}`;
}

function sortedRestaurantIdRangeExpression(layout: Layout): string {
  const column = columnName(layout.sortedRestaurantIdColumn + 1);
  return `'_Metric_Calc'!$${column}$2:$${column}$${layout.restaurantCapacity + 1}`;
}

function numericDisplayRangeExpression(layout: Layout): string {
  const firstColumn = columnName(layout.numericDisplayStartColumn + 1);
  const lastColumn = columnName(layout.numericDisplayStartColumn + layout.groupCapacity);
  return `'_Metric_Calc'!$${firstColumn}$2:$${lastColumn}$${layout.restaurantCapacity + 2}`;
}

function componentTotalExpression(layout: Layout, blockIndex: number): string {
  const block = layout.componentBlocks[blockIndex];
  const firstColumn = columnName(block.startColumn + 1);
  const lastColumn = columnName(block.startColumn + layout.groupCapacity);
  const totalExcelRow = layout.componentTotalRow + 1;
  return `'_Metric_Calc'!$${firstColumn}$${totalExcelRow}:$${lastColumn}$${totalExcelRow}`;
}

function applyMatrixConditionalFormats(
  performance: ExcelScript.Worksheet,
  matrixValues: ExcelScript.Range,
  layout: Layout,
  red: string
): void {
  matrixValues.clearAllConditionalFormats();
  const firstBodyRow = layout.matrixBodyStartRow + 1;
  const negative = matrixValues.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
  negative.getCustom().getRule().setFormula(
    `=AND(OR($G$6="PP Change",$G$6="NOK Impact"),LEFT(B${firstBodyRow},1)="-")`
  );
  negative.getCustom().getFormat().getFont().setColor(red);
  performance.getRangeByIndexes(
    layout.matrixBodyStartRow,
    0,
    layout.restaurantCapacity + 1,
    1
  ).getFormat().getFont().setBold(false);
}

function writeInteractionCalc(
  sheet: ExcelScript.Worksheet,
  activeImports: ActiveImport[],
  activeGroups: ReportingGroup[],
  prior: PriorUiState,
  layout: Layout,
  companyFingerprint: string
): void {
  const priorLayout = makeLayout(
    Math.max(1, prior.priorGroupCapacity),
    Math.max(1, prior.priorRestaurantCapacity)
  );
  const clearRows = Math.max(layout.restaurantCapacity, priorLayout.restaurantCapacity) + 16;
  const clearLastColumn = Math.max(layout.helperLastColumn, priorLayout.helperLastColumn);
  checkedRangeByIndexes(
    sheet,
    "_Metric_Calc",
    0,
    layout.selectedRestaurantStartColumn,
    clearRows,
    clearLastColumn - layout.selectedRestaurantStartColumn + 1,
    "clear Phase 2C helper surface"
  ).clear(ExcelScript.ClearApplyTo.all);

  sheet.getRangeByIndexes(0, layout.selectedRestaurantStartColumn, 1, 2).setValues([[
    "Selected Restaurant", "Selected RestaurantID"
  ]]);
  sheet.getRangeByIndexes(0, layout.selectedGroupStartColumn, 1, 2).setValues([[
    "Selected Reporting Group", "Selected ReportingGroupID"
  ]]);
  styleHeader(sheet.getRangeByIndexes(0, layout.selectedRestaurantStartColumn, 1, 2), "#172033", "#FFFFFF");
  styleHeader(sheet.getRangeByIndexes(0, layout.selectedGroupStartColumn, 1, 2), "#172033", "#FFFFFF");

  const selectedRestaurantFormulas: string[][] = [];
  for (let index = 0; index < layout.restaurantCapacity; index++) {
    const excelRow = index + 2;
    selectedRestaurantFormulas.push([
      `=IFERROR(INDEX(FILTER(tblPerformanceRestaurantSelection[Restaurant],` +
        `tblPerformanceRestaurantSelection[Include]="Yes"),ROWS($AE$2:AE${excelRow})),"")`,
      `=IFERROR(INDEX(FILTER(tblPerformanceRestaurantSelection[RestaurantID],` +
        `tblPerformanceRestaurantSelection[Include]="Yes"),ROWS($AF$2:AF${excelRow})),"")`
    ]);
  }
  sheet.getRangeByIndexes(1, layout.selectedRestaurantStartColumn, layout.restaurantCapacity, 2)
    .setFormulas(selectedRestaurantFormulas);

  const selectedGroupFormulas: string[][] = [];
  for (let index = 0; index < layout.groupCapacity; index++) {
    const excelRow = index + 2;
    selectedGroupFormulas.push([
      `=IFERROR(INDEX(FILTER(tblPerformanceRPGSelection[Reporting Group],` +
        `tblPerformanceRPGSelection[Include]="Yes"),ROWS($AH$2:AH${excelRow})),"")`,
      `=IFERROR(INDEX(FILTER(tblPerformanceRPGSelection[ReportingGroupID],` +
        `tblPerformanceRPGSelection[Include]="Yes"),ROWS($AI$2:AI${excelRow})),"")`
    ]);
  }
  sheet.getRangeByIndexes(1, layout.selectedGroupStartColumn, layout.groupCapacity, 2)
    .setFormulas(selectedGroupFormulas);

  sheet.getRangeByIndexes(0, layout.controlStartColumn, 15, 2).setValues([
    ["Phase 2C control", "Value"],
    ["Current ImportID", ""],
    ["Comparison ImportID", ""],
    ["Selected restaurant count", ""],
    ["Selected RPG count", ""],
    ["Detail ReportingGroupID", ""],
    ["Matrix display", ""],
    ["Company scope fingerprint", companyFingerprint],
    ["Interaction contract", "RPG result components only"],
    ["Restaurant capacity", layout.restaurantCapacity],
    ["RPG capacity", layout.groupCapacity],
    ["Requested sort ReportingGroupID", ""],
    ["Effective sort ReportingGroupID", ""],
    ["Sort order", ""],
    ["Sort status", ""]
  ]);
  styleHeader(sheet.getRangeByIndexes(0, layout.controlStartColumn, 1, 2), "#172033", "#FFFFFF");
  const controlValueColumn = columnName(layout.controlStartColumn + 2);
  sheet.getRange(`${controlValueColumn}2`).setFormula(
    `=IFERROR(INDEX($B$2:$B$${activeImports.length + 1},MATCH(Performance!$B$10,$A$2:$A$${activeImports.length + 1},0)),"")`
  );
  sheet.getRange(`${controlValueColumn}3`).setFormula(
    `=IFERROR(INDEX($B$2:$B$${activeImports.length + 1},MATCH(Performance!$G$10,$A$2:$A$${activeImports.length + 1},0)),"")`
  );
  sheet.getRange(`${controlValueColumn}4`).setFormula(
    '=COUNTIF(tblPerformanceRestaurantSelection[Include],"Yes")'
  );
  sheet.getRange(`${controlValueColumn}5`).setFormula(
    '=COUNTIF(tblPerformanceRPGSelection[Include],"Yes")'
  );
  sheet.getRange(`${controlValueColumn}6`).setFormula(
    `=IFERROR(INDEX($J$2:$J$${activeGroups.length + 1},MATCH(Performance!$B$7,$I$2:$I$${activeGroups.length + 1},0)),"")`
  );
  sheet.getRange(`${controlValueColumn}7`).setFormula("=Performance!$G$6");
  sheet.getRange(`${controlValueColumn}12`).setFormula(
    `=IF(Performance!$I$6="Total","",IFERROR(INDEX($J$2:$J$${activeGroups.length + 1},` +
      `MATCH(Performance!$I$6,$I$2:$I$${activeGroups.length + 1},0)),""))`
  );
  sheet.getRange(`${controlValueColumn}13`).setFormula(
    `=IF(OR(${controlValueColumn}12="",COUNTIF($AI$2:$AI$${layout.groupCapacity + 1},${controlValueColumn}12)=0),"",${controlValueColumn}12)`
  );
  sheet.getRange(`${controlValueColumn}14`).setFormula("=Performance!$I$7");
  sheet.getRange(`${controlValueColumn}15`).setFormula(
    `=IF(Performance!$I$6="Total","",` +
      `IF(${controlValueColumn}12="","Using Total — target unavailable",` +
      `IF(${controlValueColumn}13="","Using Total — "&Performance!$I$6&" hidden","")))`
  );

  for (let blockIndex = 0; blockIndex < layout.componentBlocks.length; blockIndex++) {
    const block = layout.componentBlocks[blockIndex];
    const headers = [`${block.name} | RPG-ID`].concat(activeGroups.map(value => value.id));
    sheet.getRangeByIndexes(0, block.startColumn - 1, 1, headers.length).setValues([headers]);
    styleHeader(sheet.getRangeByIndexes(0, block.startColumn - 1, 1, headers.length), "#EEF1F5", "#172033");
  }
  const numericDisplayHeaders = ["Selected Numeric Display | RPG-ID"].concat(
    activeGroups.map(value => value.id)
  );
  sheet.getRangeByIndexes(
    0,
    layout.numericDisplayStartColumn - 1,
    1,
    numericDisplayHeaders.length
  ).setValues([numericDisplayHeaders]);
  styleHeader(
    sheet.getRangeByIndexes(
      0,
      layout.numericDisplayStartColumn - 1,
      1,
      numericDisplayHeaders.length
    ),
    "#EEF1F5",
    "#172033"
  );
  const totalHeaders = [
    "Selected RPG Current Numerator", "Current Scope Denominator",
    "Selected RPG Comparison Numerator", "Comparison Scope Denominator",
    "Selected RPG Total Numeric Display", "Restaurant Numeric Sort Key", "Sorted RestaurantID"
  ];
  sheet.getRangeByIndexes(0, layout.totalComponentStartColumn, 1, totalHeaders.length).setValues([totalHeaders]);
  styleHeader(
    sheet.getRangeByIndexes(0, layout.totalComponentStartColumn, 1, totalHeaders.length),
    "#EEF1F5",
    "#172033"
  );

  writeComponentBlock(sheet, layout, 0, activeGroups, true, true);
  writeComponentBlock(sheet, layout, 1, activeGroups, true, false);
  writeShareBlock(sheet, layout, 2, 0, 1);
  writeComponentBlock(sheet, layout, 3, activeGroups, false, true);
  writeComponentBlock(sheet, layout, 4, activeGroups, false, false);
  writeShareBlock(sheet, layout, 5, 3, 4);
  writeNumericDisplayBlock(sheet, layout, activeGroups);
  writeTotalAndSortHelpers(sheet, layout, activeGroups);

  sheet.getRangeByIndexes(1, layout.componentBlocks[0].startColumn, layout.restaurantCapacity + 1, layout.groupCapacity)
    .setNumberFormat('#,##0.00 "NOK"');
  sheet.getRangeByIndexes(1, layout.componentBlocks[1].startColumn, layout.restaurantCapacity + 1, layout.groupCapacity)
    .setNumberFormat('#,##0.00 "NOK"');
  sheet.getRangeByIndexes(1, layout.componentBlocks[2].startColumn, layout.restaurantCapacity + 1, layout.groupCapacity)
    .setNumberFormat("0.00%");
  sheet.getRangeByIndexes(1, layout.componentBlocks[3].startColumn, layout.restaurantCapacity + 1, layout.groupCapacity)
    .setNumberFormat('#,##0.00 "NOK"');
  sheet.getRangeByIndexes(1, layout.componentBlocks[4].startColumn, layout.restaurantCapacity + 1, layout.groupCapacity)
    .setNumberFormat('#,##0.00 "NOK"');
  sheet.getRangeByIndexes(1, layout.componentBlocks[5].startColumn, layout.restaurantCapacity + 1, layout.groupCapacity)
    .setNumberFormat("0.00%");
  sheet.getRangeByIndexes(
    1,
    layout.numericDisplayStartColumn,
    layout.restaurantCapacity + 1,
    layout.groupCapacity
  ).setNumberFormat("General");
  sheet.getRangeByIndexes(
    1,
    layout.totalComponentStartColumn,
    layout.restaurantCapacity + 1,
    4
  ).setNumberFormat('#,##0.00 "NOK"');
  sheet.getRangeByIndexes(
    1,
    layout.totalDisplayColumn,
    layout.restaurantCapacity + 1,
    2
  ).setNumberFormat("General");
  sheet.getRangeByIndexes(0, layout.selectedRestaurantStartColumn, 1, clearLastColumn - layout.selectedRestaurantStartColumn + 1)
    .getEntireColumn().getFormat().setColumnWidth(115);
}

function writeNumericDisplayBlock(
  sheet: ExcelScript.Worksheet,
  layout: Layout,
  groups: ReportingGroup[]
): void {
  const formulas: string[][] = [];
  const displayCell = `$${columnName(layout.controlStartColumn + 2)}$7`;
  for (let rowIndex = 0; rowIndex < layout.restaurantCapacity + 1; rowIndex++) {
    const excelRow = rowIndex + 2;
    const row: string[] = [];
    for (let groupIndex = 0; groupIndex < groups.length; groupIndex++) {
      const currentNumerator = componentCellReference(layout, 0, excelRow, groupIndex);
      const currentDenominator = componentCellReference(layout, 1, excelRow, groupIndex);
      const currentShare = componentCellReference(layout, 2, excelRow, groupIndex);
      const comparisonNumerator = componentCellReference(layout, 3, excelRow, groupIndex);
      const comparisonDenominator = componentCellReference(layout, 4, excelRow, groupIndex);
      const comparisonShare = componentCellReference(layout, 5, excelRow, groupIndex);
      row.push(
        `=IF(${displayCell}="Current Sales NOK",IF(${currentDenominator}=0,"",${currentNumerator}),` +
          `IF(${displayCell}="Current Share",IF(${currentDenominator}=0,"",${currentShare}),` +
          `IF(${displayCell}="Comparison Share",IF(${comparisonDenominator}=0,"",${comparisonShare}),` +
          `IF(${displayCell}="PP Change",IF(OR(${currentDenominator}=0,${comparisonDenominator}=0),"",` +
          `(${currentShare}-${comparisonShare})*100),` +
          `IF(${displayCell}="NOK Impact",IF(OR(${currentDenominator}=0,${comparisonDenominator}=0),"",` +
          `${currentNumerator}-((${comparisonNumerator}/${comparisonDenominator})*${currentDenominator})),"")))))`
      );
    }
    formulas.push(row);
  }
  sheet.getRangeByIndexes(
    1,
    layout.numericDisplayStartColumn,
    layout.restaurantCapacity + 1,
    layout.groupCapacity
  ).setFormulas(formulas);
}

function writeTotalAndSortHelpers(
  sheet: ExcelScript.Worksheet,
  layout: Layout,
  groups: ReportingGroup[]
): void {
  const currentNumeratorColumn = columnName(layout.totalComponentStartColumn + 1);
  const currentDenominatorColumn = columnName(layout.totalComponentStartColumn + 2);
  const comparisonNumeratorColumn = columnName(layout.totalComponentStartColumn + 3);
  const comparisonDenominatorColumn = columnName(layout.totalComponentStartColumn + 4);
  const totalDisplayColumn = columnName(layout.totalDisplayColumn + 1);
  const sortKeyColumn = columnName(layout.sortKeyColumn + 1);
  const sortedIdColumn = columnName(layout.sortedRestaurantIdColumn + 1);
  const formulas: string[][] = [];
  const displayCell = `$${columnName(layout.controlStartColumn + 2)}$7`;
  for (let rowIndex = 0; rowIndex < layout.restaurantCapacity + 1; rowIndex++) {
    const excelRow = rowIndex + 2;
    const currentNumerator = `$${currentNumeratorColumn}${excelRow}`;
    const currentDenominator = `$${currentDenominatorColumn}${excelRow}`;
    const comparisonNumerator = `$${comparisonNumeratorColumn}${excelRow}`;
    const comparisonDenominator = `$${comparisonDenominatorColumn}${excelRow}`;
    formulas.push([
      `=IF($AL$5=0,0,${selectedNumeratorExpression(layout, 0, excelRow, groups.length)})`,
      `=${componentCellReference(layout, 1, excelRow, 0)}`,
      `=IF($AL$5=0,0,${selectedNumeratorExpression(layout, 3, excelRow, groups.length)})`,
      `=${componentCellReference(layout, 4, excelRow, 0)}`,
      `=IF($AL$5=0,"",IF(${displayCell}="Current Sales NOK",${currentNumerator},` +
        `IF(${displayCell}="Current Share",IF(${currentDenominator}=0,"",${currentNumerator}/${currentDenominator}),` +
        `IF(${displayCell}="Comparison Share",IF(${comparisonDenominator}=0,"",${comparisonNumerator}/${comparisonDenominator}),` +
        `IF(${displayCell}="PP Change",IF(OR(${currentDenominator}=0,${comparisonDenominator}=0),"",` +
        `((${currentNumerator}/${currentDenominator})-(${comparisonNumerator}/${comparisonDenominator}))*100),` +
        `IF(${displayCell}="NOK Impact",IF(OR(${currentDenominator}=0,${comparisonDenominator}=0),"",` +
        `${currentNumerator}-((${comparisonNumerator}/${comparisonDenominator})*${currentDenominator})),""))))))`
    ]);
  }
  sheet.getRangeByIndexes(
    1,
    layout.totalComponentStartColumn,
    layout.restaurantCapacity + 1,
    5
  ).setFormulas(formulas);

  const sortKeyFormulas: string[][] = [];
  const sortedIdFormulas: string[][] = [];
  const idRange = `$AF$2:$AF$${layout.restaurantCapacity + 1}`;
  const keyRange = `$${sortKeyColumn}$2:$${sortKeyColumn}$${layout.restaurantCapacity + 1}`;
  for (let rowIndex = 0; rowIndex < layout.restaurantCapacity; rowIndex++) {
    const excelRow = rowIndex + 2;
    const numericFirstColumn = columnName(layout.numericDisplayStartColumn + 1);
    const numericLastColumn = columnName(layout.numericDisplayStartColumn + layout.groupCapacity);
    sortKeyFormulas.push([
      `=IF($AF${excelRow}="","",IF($AL$13="",$${totalDisplayColumn}${excelRow},` +
        `IFERROR(INDEX($${numericFirstColumn}${excelRow}:$${numericLastColumn}${excelRow},1,` +
        `MATCH($AL$13,$J$2:$J$${layout.groupCapacity + 1},0)),"")))`
    ]);
    sortedIdFormulas.push([
      `=IFERROR(INDEX(SORTBY(FILTER(${idRange},${idRange}<>""),` +
        `FILTER(--(${keyRange}=""),${idRange}<>""),1,` +
        `FILTER(IF(${keyRange}="",0,${keyRange}),${idRange}<>""),IF($AL$14="Highest first",-1,1),` +
        `FILTER(IF(${keyRange}="","",${idRange}),${idRange}<>""),IF($AL$14="Highest first",1,-1),` +
        `FILTER(${idRange},${idRange}<>""),1),` +
        `ROWS($${sortedIdColumn}$2:${sortedIdColumn}${excelRow})),"")`
    ]);
  }
  sheet.getRangeByIndexes(1, layout.sortKeyColumn, layout.restaurantCapacity, 1).setFormulas(sortKeyFormulas);
  sheet.getRangeByIndexes(1, layout.sortedRestaurantIdColumn, layout.restaurantCapacity, 1)
    .setFormulas(sortedIdFormulas);
}

function selectedNumeratorExpression(
  layout: Layout,
  blockIndex: number,
  excelRow: number,
  groupCount: number
): string {
  const terms: string[] = [];
  for (let groupIndex = 0; groupIndex < groupCount; groupIndex++) {
    terms.push(
      `IF(COUNTIF($AI$2:$AI$${layout.groupCapacity + 1},$J$${groupIndex + 2})>0,` +
        `${componentCellReference(layout, blockIndex, excelRow, groupIndex)},0)`
    );
  }
  return terms.join("+") || "0";
}

function componentCellReference(
  layout: Layout,
  blockIndex: number,
  excelRow: number,
  groupIndex: number
): string {
  const column = columnName(layout.componentBlocks[blockIndex].startColumn + groupIndex + 1);
  return `$${column}${excelRow}`;
}

function writeComponentBlock(
  sheet: ExcelScript.Worksheet,
  layout: Layout,
  blockIndex: number,
  groups: ReportingGroup[],
  current: boolean,
  numerator: boolean
): void {
  const block = layout.componentBlocks[blockIndex];
  const formulas: string[][] = [];
  const importCell = current ? "$AL$2" : "$AL$3";
  const valueColumn = numerator ? "NumeratorSalesNOK" : "DenominatorSalesNOK";
  for (let restaurantIndex = 0; restaurantIndex < layout.restaurantCapacity; restaurantIndex++) {
    const restaurantExcelRow = restaurantIndex + 2;
    const row: string[] = [];
    for (let groupIndex = 0; groupIndex < groups.length; groupIndex++) {
      const groupExcelRow = groupIndex + 2;
      row.push(
        `=IF(OR($AF${restaurantExcelRow}="",${importCell}=""),0,` +
          `SUMIFS(tblMetricRPGResults[${valueColumn}],` +
          `tblMetricRPGResults[MetricID],"${METRIC_ID}",` +
          `tblMetricRPGResults[ImportID],${importCell},` +
          `tblMetricRPGResults[ReportingGroupID],$J$${groupExcelRow},` +
          `tblMetricRPGResults[ScopeType],"Restaurant",` +
          `tblMetricRPGResults[RestaurantID],$AF${restaurantExcelRow},` +
          `tblMetricRPGResults[ChannelScope],"${ALL_CHANNELS}",` +
          `tblMetricRPGResults[PublicationState],"${ACTIVE_FINALIZED}"))`
      );
    }
    formulas.push(row);
  }
  sheet.getRangeByIndexes(1, block.startColumn, layout.restaurantCapacity, layout.groupCapacity).setFormulas(formulas);
  const totalFormulas: string[][] = [[]];
  for (let groupIndex = 0; groupIndex < groups.length; groupIndex++) {
    const column = columnName(block.startColumn + groupIndex + 1);
    totalFormulas[0].push(`=SUM(${column}$2:${column}$${layout.restaurantCapacity + 1})`);
  }
  sheet.getRangeByIndexes(layout.componentTotalRow, block.startColumn, 1, layout.groupCapacity).setFormulas(totalFormulas);
}

function writeShareBlock(
  sheet: ExcelScript.Worksheet,
  layout: Layout,
  shareBlockIndex: number,
  numeratorBlockIndex: number,
  denominatorBlockIndex: number
): void {
  const shareBlock = layout.componentBlocks[shareBlockIndex];
  const numeratorBlock = layout.componentBlocks[numeratorBlockIndex];
  const denominatorBlock = layout.componentBlocks[denominatorBlockIndex];
  const formulas: string[][] = [];
  for (let rowIndex = 0; rowIndex < layout.restaurantCapacity + 1; rowIndex++) {
    const excelRow = rowIndex + 2;
    const row: string[] = [];
    for (let groupIndex = 0; groupIndex < layout.groupCapacity; groupIndex++) {
      const numeratorColumn = columnName(numeratorBlock.startColumn + groupIndex + 1);
      const denominatorColumn = columnName(denominatorBlock.startColumn + groupIndex + 1);
      row.push(`=IF(${denominatorColumn}${excelRow}=0,0,${numeratorColumn}${excelRow}/${denominatorColumn}${excelRow})`);
    }
    formulas.push(row);
  }
  sheet.getRangeByIndexes(1, shareBlock.startColumn, layout.restaurantCapacity + 1, layout.groupCapacity)
    .setFormulas(formulas);
}

function writeReports(reports: ExcelScript.Worksheet): void {
  const NAVY = "#172033";
  const BLUE = "#4F8CFF";
  const LIGHT = "#EAF2FF";
  const GREY = "#EEF1F5";
  const WHITE = "#FFFFFF";
  reports.getRange("A1:H35").clear(ExcelScript.ClearApplyTo.all);
  writeTitle(
    reports,
    "Reports",
    "Preview uses the same interactive restaurant scope and selected Reporting Group result shown in Performance.",
    "H",
    NAVY,
    LIGHT,
    WHITE
  );
  reports.getRange("A5:D5").setValues([["Report context", "", "", ""]]);
  styleSection(reports.getRange("A5:D5"), BLUE, WHITE);
  reports.getRange("A6:B11").setValues([
    ["KPI", METRIC_NAME],
    ["Reporting Group", ""],
    ["Current", ""],
    ["Compare with", ""],
    ["Scope check", ""],
    ["Restaurant scope", ""]
  ]);
  styleLabels(reports.getRange("A6:A11"), GREY);
  reports.getRange("B7").setFormula("=Performance!B7");
  reports.getRange("B8").setFormula("=Performance!B10");
  reports.getRange("B9").setFormula("=Performance!G10");
  reports.getRange("B10").setFormula("=Performance!G18");
  reports.getRange("B11").setFormula("=Performance!G19");
  reports.getRange("B10:D11").getFormat().setWrapText(true);

  reports.getRange("A14:D14").setValues([["Selected result", "", "", ""]]);
  styleSection(reports.getRange("A14:D14"), NAVY, WHITE);
  reports.getRange("A15:B18").setValues([
    ["Current share", ""],
    ["Comparison share", ""],
    ["Change", ""],
    ["Current Reporting Group sales", ""]
  ]);
  styleLabels(reports.getRange("A15:A18"), GREY);
  reports.getRange("B15").setFormula("=Performance!B16");
  reports.getRange("B16").setFormula("=Performance!B17");
  reports.getRange("B17").setFormula("=Performance!B18");
  reports.getRange("B18").setFormula("=Performance!B19");
  reports.getRange("B15:B16").setNumberFormat("0.00%");
  reports.getRange("B17").setNumberFormat('+0.00 "pp";-0.00 "pp";0.00 "pp"');
  reports.getRange("B18").setNumberFormat('#,##0 "NOK"');
  reports.getRange("A:H").getFormat().setColumnWidth(18);
  reports.getRange("A:A").getFormat().setColumnWidth(26);
  reports.getRange("B:B").getFormat().setColumnWidth(42);
  reports.getFreezePanes().freezeRows(2);
}

function writeInteractionQA(
  workbook: ExcelScript.Workbook,
  sheet: ExcelScript.Worksheet,
  phase2BQaTable: ExcelScript.Table,
  imports: ActiveImport[],
  groups: ReportingGroup[],
  restaurants: ReportingRestaurant[],
  results: MetricResult[],
  companyFingerprint: string,
  validationMessage: string
): void {
  const existing = workbook.getTable("tblPerformanceInteractionQA");
  let clearStartRow = phase2BQaTable.getRange().getRowIndex() + phase2BQaTable.getRange().getRowCount() + 2;
  let clearRowCount = 30;
  if (existing) {
    clearStartRow = Math.min(clearStartRow, existing.getRange().getRowIndex() - 1);
    clearRowCount = Math.max(clearRowCount, existing.getRange().getRowCount() + 4);
    existing.delete();
  }
  const phase2BEndRow = phase2BQaTable.getRange().getRowIndex() + phase2BQaTable.getRange().getRowCount();
  const startRow = Math.max(phase2BEndRow + 2, clearStartRow);
  checkedRangeByIndexes(
    sheet,
    "Metric Results QA",
    startRow,
    0,
    clearRowCount,
    5,
    "clear Phase 2C QA surface"
  ).clear(ExcelScript.ClearApplyTo.all);

  const performance = requiredSheet(workbook, "Performance");
  const calc = requiredSheet(workbook, "_Metric_Calc");
  const reports = requiredSheet(workbook, "Reports");
  const restaurantSelection = requiredTable(workbook, "tblPerformanceRestaurantSelection");
  const groupSelection = requiredTable(workbook, "tblPerformanceRPGSelection");
  const layout = makeLayout(groups.length, restaurants.length);
  const helperFormulas = checkedRangeByIndexes(
    calc,
    "_Metric_Calc",
    0,
    layout.selectedRestaurantStartColumn,
    restaurants.length + 2,
    layout.helperLastColumn - layout.selectedRestaurantStartColumn + 1,
    "read Phase 2C helper formulas"
  ).getFormulas();
  const performanceFormulas = checkedRangeByIndexes(
    performance,
    "Performance",
    0,
    0,
    layout.matrixEndRow + 1,
    layout.matrixEndColumn + 1,
    "read Phase 2C Performance formulas"
  ).getFormulas();
  const reportsFormulas = reports.getRange("A1:H20").getFormulas();
  const helperText = matrixText(helperFormulas);
  const performanceText = matrixText(performanceFormulas);
  const reportsText = matrixText(reportsFormulas);
  const expectedRows = imports.length * groups.length * (1 + restaurants.length);
  const restaurantSelectionCount = restaurantSelection.getRangeBetweenHeaderAndTotal().getRowCount();
  const groupSelectionCount = groupSelection.getRangeBetweenHeaderAndTotal().getRowCount();
  const additiveOnly = helperText.indexOf("tblMetricRPGResults") >= 0 &&
    helperText.indexOf("tblSalesFacts") < 0 &&
    helperText.indexOf("ReportingCategoryID") < 0;
  const presentationOnly = performanceText.indexOf("tblSalesFacts") < 0 &&
    performanceText.indexOf("ReportingCategoryID") < 0;
  const reportsLinked = reportsText.indexOf("Performance!B16") >= 0 &&
    reportsText.indexOf("Performance!B19") >= 0 && reportsText.indexOf("tblSalesFacts") < 0;
  const grandTotalFormulaText = matrixText(
    calc.getRangeByIndexes(layout.componentTotalRow, layout.componentBlocks[0].startColumn, 1, layout.groupCapacity).getFormulas()
  );
  const canonicalShareText = matrixText(
    calc.getRangeByIndexes(1, layout.componentBlocks[2].startColumn, restaurants.length + 1, groups.length).getFormulas()
  );
  const numericDisplayText = matrixText(
    calc.getRangeByIndexes(
      1,
      layout.numericDisplayStartColumn,
      restaurants.length + 1,
      groups.length
    ).getFormulas()
  );
  const totalHelperText = matrixText(
    calc.getRangeByIndexes(
      1,
      layout.totalComponentStartColumn,
      restaurants.length + 1,
      5
    ).getFormulas()
  );
  const grandTotalTotalDisplayText = matrixText(
    calc.getRangeByIndexes(layout.componentTotalRow, layout.totalDisplayColumn, 1, 1).getFormulas()
  );
  const sortHelperText = matrixText(
    calc.getRangeByIndexes(1, layout.sortKeyColumn, restaurants.length, 2).getFormulas()
  );
  const visibleMatrixText = matrixText(
    performance.getRangeByIndexes(
      layout.matrixBodyStartRow,
      1,
      restaurants.length + 1,
      groups.length + 1
    ).getFormulas()
  );
  const presentationFacadeIsolated =
    numericDisplayText.indexOf("FIXED(") < 0 &&
    numericDisplayText.indexOf('" pp"') < 0 &&
    numericDisplayText.indexOf('" NOK"') < 0 &&
    totalHelperText.indexOf("FIXED(") < 0 &&
    totalHelperText.indexOf('" pp"') < 0 &&
    totalHelperText.indexOf('" NOK"') < 0 &&
    visibleMatrixText.indexOf("FIXED(") >= 0 &&
    visibleMatrixText.indexOf('" pp"') >= 0 &&
    visibleMatrixText.indexOf('" NOK"') >= 0 &&
    visibleMatrixText.indexOf("ISNUMBER(") >= 0;
  const dashVisible = performanceText.indexOf(DASH) >= 0;

  const rows: (string | number | boolean)[][] = [
    ["QA-0302C-01", "Accepted Phase 2B result grain", results.length === expectedRows ? "PASS" : "FAIL", results.length, `${imports.length} datasets × ${groups.length} RPGs × ${1 + restaurants.length} Phase 2B scopes = ${expectedRows}.`],
    ["QA-0302C-02", "All-selected restaurant components equal Company controls", "PASS", companyFingerprint, "Preflight reconciled numerator, denominator, and canonical metric for every dataset/RPG."],
    ["QA-0302C-03", "Stable-ID restaurant selection catalog", restaurantSelectionCount === restaurants.length ? "PASS" : "FAIL", restaurantSelectionCount, `${restaurants.length} refresh-eligible RestaurantIDs; newly eligible IDs default No after initial catalog creation.`],
    ["QA-0302C-04", "Stable-ID Reporting Group selection catalog", groupSelectionCount === groups.length ? "PASS" : "FAIL", groupSelectionCount, `${groups.length} active ReportingGroupIDs; detail selector remains separate.`],
    ["QA-0302C-05", "Shared selected-restaurant scope contract", additiveOnly ? "PASS" : "FAIL", restaurants.length, "Phase 2C component formulas aggregate Phase 2B Restaurant rows only; Company rows remain QA controls."],
    ["QA-0302C-06", "No legacy/fact calculation in Performance", additiveOnly && presentationOnly ? "PASS" : "FAIL", 0, "No tblSalesFacts or ReportingCategoryID dependency exists in Phase 2C helpers or Performance."],
    ["QA-0302C-07", "Grand Total uses summed components", grandTotalFormulaText.indexOf("SUM(") >= 0 && grandTotalTotalDisplayText.indexOf("NOK Impact") >= 0 && grandTotalTotalDisplayText.indexOf("SUM(") < 0 ? "PASS" : "FAIL", groups.length, "NOK Impact uses aggregated current numerator minus aggregated comparison share × aggregated current denominator. It never sums restaurant impacts or parses visible values."],
    ["QA-0302C-08", "Zero-denominator contract", canonicalShareText.indexOf("=0,0,") >= 0 && dashVisible ? "PASS" : "FAIL", DASH, "Canonical helper value remains zero; Performance displays an em dash."],
    ["QA-0302C-09", "Five matrix display modes", MATRIX_MODES.length === 5 && presentationFacadeIsolated ? "PASS" : "FAIL", MATRIX_MODES.join(" | "), "A bounded numeric helper remains authoritative; the visible matrix is an isolated FIXED-based text facade with no conditional number format."],
    ["QA-0302C-10", "Authoritative Include-state selection", helperText.indexOf('Performance!$G$7="All"') < 0 && helperText.indexOf('Performance!$G$8="All"') < 0 ? "PASS" : "FAIL", "Yes / No", "Selection tables directly determine restaurant and RPG scope; All versus Custom is derived, not user-selected."],
    ["QA-0302C-11", "Reports linkage", reportsLinked ? "PASS" : "FAIL", 0, "Reports links to the same interactive Performance detail result and selected restaurant scope."],
    ["QA-0302C-12", "Interactive validation", validationMessage.indexOf("ready") >= 0 ? "PASS" : "WARN", 0, validationMessage],
    ["QA-0302C-13", "Protected authoritative surfaces", "PASS", 0, "Facts, bridge, centralized results, mappings, Reporting Groups, restaurants, KPI Registry, and Phase 2B QA fingerprints were unchanged."],
    ["QA-0302C-14", "Recalculation-only exploration", "PASS", 0, "Restaurant, RPG, current dataset, comparison dataset, display, sort target, and sort order changes require no Office Script rerun."],
    ["QA-0302C-15", "Selected-RPG Total contract", totalHelperText.indexOf("COUNTIF($AI$2:$AI$") >= 0 && totalHelperText.indexOf("Current Sales NOK") >= 0 ? "PASS" : "FAIL", groups.length, "Total sums only selected RPG numerators, uses each scope denominator once, and keeps Current Sales NOK equal to the selected numerator."],
    ["QA-0302C-16", "Numeric presentation sorting", sortHelperText.indexOf("SORTBY(") >= 0 && sortHelperText.indexOf("FIXED(") < 0 && sortHelperText.indexOf("$AF$2:$AF$") >= 0 ? "PASS" : "FAIL", restaurants.length, "Only the visible restaurant lookup order changes. Full-precision numeric keys drive sorting; unavailable keys remain last and Grand Total is excluded."]
  ];
  sheet.getRangeByIndexes(startRow, 0, 1, 5).setValues([["Phase 2C Interactive Performance QA", "", "", "", ""]]);
  styleSection(sheet.getRangeByIndexes(startRow, 0, 1, 5), "#4F8CFF", "#FFFFFF");
  sheet.getRangeByIndexes(startRow + 1, 0, 1, 5).setValues([[
    "CheckID", "Check", "Result", "Observed", "Explanation"
  ]]);
  sheet.getRangeByIndexes(startRow + 2, 0, rows.length, 5).setValues(rows);
  const table = sheet.addTable(sheet.getRangeByIndexes(startRow + 1, 0, rows.length + 1, 5), true);
  table.setName("tblPerformanceInteractionQA");
  table.setPredefinedTableStyle("TableStyleMedium2");
  styleHeader(sheet.getRangeByIndexes(startRow + 1, 0, 1, 5), "#EEF1F5", "#172033");
  setWidths(sheet, [115, 280, 80, 185, 430]);
  sheet.getRangeByIndexes(startRow, 0, rows.length + 2, 5).getFormat().setWrapText(true);
}

function deterministicMetricResultId(row: MetricResult, companyFingerprint: string): string {
  const scopeFingerprint = row.scopeType === "Company" ? companyFingerprint : "";
  const key = [
    row.metricId,
    row.importId,
    row.groupId,
    row.scopeType,
    row.restaurantId,
    scopeFingerprint,
    ALL_CHANNELS,
    ACTIVE_FINALIZED
  ].join("|");
  return hashStrings([key], "MRR-");
}

function deterministicRestaurantScopeFingerprint(restaurants: ReportingRestaurant[]): string {
  const ids = restaurants.map(value => value.id);
  ids.sort();
  const serialized = ids.map(id => `${id.length}:${id}`).join("|");
  return hashStrings([`ENABLED-RESTAURANTS|${serialized}`], "RSC-");
}

function metricResultKey(
  importId: string,
  groupId: string,
  scopeType: string,
  restaurantId: string
): string {
  return [importId, groupId, scopeType, restaurantId].join("|");
}

function protectedStateFingerprint(
  facts: (string | number | boolean)[][],
  bridge: (string | number | boolean)[][],
  results: (string | number | boolean)[][],
  imports: (string | number | boolean)[][],
  groups: (string | number | boolean)[][],
  restaurants: (string | number | boolean)[][],
  rules: (string | number | boolean)[][],
  effective: (string | number | boolean)[][],
  kpis: (string | number | boolean)[][],
  phase2BQa: (string | number | boolean)[][]
): string {
  const state = newHashState();
  updateNamedMatrix(state, "tblSalesFacts", facts);
  updateNamedMatrix(state, "tblMetricRPGFacts", bridge);
  updateNamedMatrix(state, "tblMetricRPGResults", results);
  updateNamedMatrix(state, "tblImports", imports);
  updateNamedMatrix(state, "tblReportingGroups", groups);
  updateNamedMatrix(state, "tblRestaurants", restaurants);
  updateNamedMatrix(state, "tblMappingRules", rules);
  updateNamedMatrix(state, "tblEffectiveMapping", effective);
  updateNamedMatrix(state, "tblKPIRegistry", kpis);
  updateNamedMatrix(state, "tblMetricResultsQA", phase2BQa);
  return finishHash(state, "PROTECTED-");
}

function updateNamedMatrix(
  state: HashState,
  name: string,
  values: (string | number | boolean)[][]
): void {
  updateHash(state, `${name}\n`);
  updateHashMatrix(state, values);
}

function applyListValidation(
  target: ExcelScript.Range,
  items: string[],
  label: string,
  failures: string[]
): void {
  try {
    const values = items.map(value => text(value)).filter(value => value.length > 0);
    if (!values.length) throw new Error("source list is empty");
    if (values.some(value => value.indexOf(",") >= 0)) throw new Error("source item contains a comma");
    const validation = target.getDataValidation();
    validation.clear();
    validation.setRule({ list: { inCellDropDown: true, source: values.join(",") } });
  } catch (error) {
    failures.push(`${label}: ${errorMessage(error)}`);
  }
}

function requiredSheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  const sheet = workbook.getWorksheet(name);
  if (!sheet) throw new Error(`PUL-0302C-014: Required worksheet missing: ${name}`);
  return sheet;
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-0302C-015: Required table missing: ${name}`);
  return table;
}

function requiredHeaderMap(
  table: ExcelScript.Table,
  requiredHeaders: string[]
): { [key: string]: number } {
  const headers = table.getHeaderRowRange().getValues()[0];
  const output: { [key: string]: number } = {};
  for (let index = 0; index < headers.length; index++) output[text(headers[index])] = index;
  for (const header of requiredHeaders) {
    if (output[header] === undefined) {
      throw new Error(`PUL-0302C-016: ${table.getName()} is missing required column ${header}.`);
    }
  }
  return output;
}

function checkedRangeByIndexes(
  sheet: ExcelScript.Worksheet,
  sheetName: string,
  startRow: number,
  startColumn: number,
  rowCount: number,
  columnCount: number,
  operation: string
): ExcelScript.Range {
  const startAddress = `${columnName(startColumn + 1)}${startRow + 1}`;
  const endAddress = `${columnName(startColumn + columnCount)}${startRow + rowCount}`;
  if (startRow < 0 || startColumn < 0 || rowCount < 1 || columnCount < 1) {
    throw new Error(
      `PUL-0302C-017: Invalid generated range for ${operation}: ${sheetName}!${startAddress}:${endAddress} ` +
      `(rowCount ${rowCount}, columnCount ${columnCount}).`
    );
  }
  try {
    return sheet.getRangeByIndexes(startRow, startColumn, rowCount, columnCount);
  } catch (error) {
    throw new Error(
      `PUL-0302C-018: Worksheet range failed for ${operation}: ${sheetName}!${startAddress}:${endAddress}. ${errorMessage(error)}`
    );
  }
}

function writeTitle(
  sheet: ExcelScript.Worksheet,
  title: string,
  subtitle: string,
  endColumn: string,
  navy: string,
  light: string,
  white: string
): void {
  const columns = columnNumber(endColumn);
  const titleRow = rowWithLength(title, columns);
  const subtitleRow = rowWithLength(subtitle, columns);
  sheet.getRange(`A1:${endColumn}1`).setValues([titleRow]);
  sheet.getRange(`A1:${endColumn}1`).getFormat().getFill().setColor(navy);
  sheet.getRange(`A1:${endColumn}1`).getFormat().getFont().setColor(white);
  sheet.getRange(`A1:${endColumn}1`).getFormat().getFont().setBold(true);
  sheet.getRange(`A1:${endColumn}1`).getFormat().getFont().setSize(19);
  sheet.getRange(`A1:${endColumn}1`).getFormat().setRowHeight(34);
  sheet.getRange(`A2:${endColumn}2`).setValues([subtitleRow]);
  sheet.getRange(`A2:${endColumn}2`).getFormat().getFill().setColor(light);
  sheet.getRange(`A2:${endColumn}2`).getFormat().setWrapText(true);
  sheet.getRange(`A2:${endColumn}2`).getFormat().setRowHeight(30);
}

function rowWithLength(firstValue: string, length: number): string[] {
  const output = [firstValue];
  while (output.length < length) output.push("");
  return output;
}

function styleSection(range: ExcelScript.Range, fill: string, fontColor: string): void {
  range.getFormat().getFill().setColor(fill);
  range.getFormat().getFont().setColor(fontColor);
  range.getFormat().getFont().setBold(true);
}

function styleHeader(range: ExcelScript.Range, fill: string, fontColor: string): void {
  range.getFormat().getFill().setColor(fill);
  range.getFormat().getFont().setColor(fontColor);
  range.getFormat().getFont().setBold(true);
  range.getFormat().setWrapText(true);
}

function styleLabels(range: ExcelScript.Range, fill: string): void {
  range.getFormat().getFill().setColor(fill);
  range.getFormat().getFont().setBold(true);
}

function setWidths(sheet: ExcelScript.Worksheet, widths: number[]): void {
  for (let index = 0; index < widths.length; index++) {
    sheet.getRangeByIndexes(0, index, 1, 1).getEntireColumn().getFormat().setColumnWidth(widths[index]);
  }
}

function updateEnvironment(table: ExcelScript.Table, key: string, value: string, note: string): void {
  const body = table.getRangeBetweenHeaderAndTotal();
  const values = body.getValues();
  for (let index = 0; index < values.length; index++) {
    if (text(values[index][0]) !== key) continue;
    body.getCell(index, 1).setValue(value);
    body.getCell(index, 2).setValue(note);
    return;
  }
  table.addRow(-1, [key, value, note]);
}

function appendBuildLog(table: ExcelScript.Table, row: (string | number | boolean)[]): void {
  table.addRow(-1, row);
}

function nextId(table: ExcelScript.Table, prefix: string, digits: number): string {
  const values = table.getRangeBetweenHeaderAndTotal().getValues();
  let maximum = 0;
  for (const row of values) {
    const match = text(row[0]).match(new RegExp(`^${prefix}-(\\d+)$`));
    if (match) maximum = Math.max(maximum, Number(match[1]));
  }
  return `${prefix}-${String(maximum + 1).padStart(digits, "0")}`;
}

function groupIdByName(groups: ReportingGroup[], name: string): string {
  const match = groups.find(value => value.name === name);
  return match ? match.id : "";
}

function importIdByLabel(imports: ActiveImport[], label: string): string {
  const match = imports.find(value => value.label === label);
  return match ? match.id : "";
}

function makePeriodLabel(start: number, end: number, year: number, week: string): string {
  if (week) return `Week ${week} ${year}`;
  const startDate = excelSerialToDate(start);
  const endDate = excelSerialToDate(end);
  if (startDate.getUTCMonth() === 0 && startDate.getUTCDate() === 1 &&
      endDate.getUTCMonth() === 11 && endDate.getUTCDate() === 31 &&
      startDate.getUTCFullYear() === endDate.getUTCFullYear()) {
    return `${startDate.getUTCFullYear()} full year`;
  }
  return `${dateLabel(start)}–${dateLabel(end)}`;
}

function excelSerialToDate(serial: number): Date {
  return new Date(Date.UTC(1899, 11, 30) + serial * 86400000);
}

function dateLabel(serial: number): string {
  const value = excelSerialToDate(serial);
  return `${String(value.getUTCDate()).padStart(2, "0")}.${String(value.getUTCMonth() + 1).padStart(2, "0")}.${value.getUTCFullYear()}`;
}

function matrixText(values: (string | number | boolean)[][]): string {
  let output = "";
  for (const row of values) {
    for (const value of row) output += `${text(value)}\n`;
  }
  return output;
}

function hashStrings(values: string[], prefix: string): string {
  const state = newHashState();
  for (const value of values) updateHash(state, `${value}\n`);
  return finishHash(state, prefix);
}

function newHashState(): HashState {
  return { left: 0, right: 0 };
}

function updateHash(state: HashState, value: string): void {
  for (let index = 0; index < value.length; index++) {
    const code = value.charCodeAt(index);
    state.left = (state.left * 131 + code) % 2147483647;
    state.right = (state.right * 137 + code) % 2147483629;
  }
}

function updateHashMatrix(state: HashState, values: (string | number | boolean)[][]): void {
  for (const row of values) {
    for (const value of row) {
      const normalized = text(value);
      updateHash(state, `${normalized.length}:${normalized}|`);
    }
    updateHash(state, "\n");
  }
}

function finishHash(state: HashState, prefix: string): string {
  return `${prefix}${state.left.toString(16).padStart(8, "0")}${state.right.toString(16).padStart(8, "0")}`;
}

function columnName(columnNumberValue: number): string {
  let value = columnNumberValue;
  let result = "";
  while (value > 0) {
    const remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }
  return result;
}

function columnNumber(column: string): number {
  let result = 0;
  for (let index = 0; index < column.length; index++) result = result * 26 + column.charCodeAt(index) - 64;
  return result;
}

function almostEqual(left: number, right: number): boolean {
  return Math.abs(left - right) <= Math.max(0.000001, Math.abs(right) * 1e-12);
}

function excelNow(): number {
  return Date.now() / 86400000 + 25569;
}

function errorMessage(error: unknown): string {
  if (error instanceof Error) return error.message;
  return text(error);
}

function numberValue(value: unknown): number {
  const converted = Number(value);
  return Number.isFinite(converted) ? converted : 0;
}

function text(value: unknown): string {
  return String(value ?? "").trim();
}
