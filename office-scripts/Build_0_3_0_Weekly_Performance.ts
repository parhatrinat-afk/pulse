/**
 * Pulse Build 0.3.0 — weekly Current / Compare Performance cutover.
 *
 * Repoints the accepted Phase 2C additive component grids to the single active
 * compact weekly cache. Selection, Total, Grand Total, NOK Impact, sorting,
 * detail, and text-presentation formulas remain the accepted Phase 2C layer.
 */
function main(workbook: ExcelScript.Workbook): string {
  const versionTable = requiredTable(workbook, "tblWeeklyCacheVersions");
  const periodTable = requiredTable(workbook, "tblWeeklyPeriodManifest");
  const scopeTable = requiredTable(workbook, "tblWeeklyScopeCache");
  const rpgTable = requiredTable(workbook, "tblWeeklyRPGCache");
  const groupsTable = requiredTable(workbook, "tblReportingGroups");
  const restaurantTable = requiredTable(workbook, "tblRestaurants");
  const interactionQa = requiredTable(workbook, "tblPerformanceInteractionQA");
  const performance = requiredSheet(workbook, "Performance");
  const reports = requiredSheet(workbook, "Reports");
  const calc = requiredSheet(workbook, "_Metric_Calc");
  const qaSheet = requiredSheet(workbook, "Metric Results QA");

  const versionRows = tableRows(versionTable);
  const periodRows = tableRows(periodTable);
  const scopeRows = tableRows(scopeTable);
  const rpgRows = tableRows(rpgTable);
  const activeGroups = readActiveReportingGroups(groupsTable);
  const authority = validateActiveCache(
    versionTable, versionRows, periodTable, periodRows, scopeTable, scopeRows,
    rpgTable, rpgRows, activeGroups
  );
  const live = validateLiveState(workbook);
  if (live.mappingContentFingerprint !== authority.mappingContentFingerprint ||
      live.catalogContentFingerprint !== authority.catalogContentFingerprint ||
      live.performanceRestaurantScopeFingerprint !== authority.performanceRestaurantScopeFingerprint ||
      live.phase2CPassCount !== 16) {
    throw new Error(`PUL-030P-001: Weekly Performance freshness preflight failed. ${JSON.stringify(live)}`);
  }
  validatePhase2CLayout(workbook, calc, restaurantTable, interactionQa);

  const protectedBefore = protectedFingerprint(workbook);
  const currentValues = performance.getRange("B10:B12").getValues();
  const comparisonValues = performance.getRange("G10:G12").getValues();
  const priorGroups = capturePriorGroupState(workbook, calc, performance);
  const plannedGroups = planGroupSelection(activeGroups, priorGroups);
  const layout = makeLayout(activeGroups.length, RESTAURANT_CAPACITY);
  const priorLayout = makeLayout(Math.max(1, priorGroups.capacity), RESTAURANT_CAPACITY);
  const alreadyWeekly = text(performance.getRange("A10").getValue()) === "Year" &&
    text(performance.getRange("F10").getValue()) === "Year";

  const validationSources = writeValidationLists(
    performance, periodRows, authority.cacheVersion, activeGroups, priorGroups.capacity, layout
  );
  writeGroupSelectionTable(workbook, plannedGroups);
  writePeriodControls(
    performance, alreadyWeekly, currentValues, comparisonValues,
    validationSources.yearSource, validationSources.weekSource
  );
  writeGroupControls(performance, activeGroups, priorGroups, layout, validationSources);
  writeWeeklyRuntimeCalc(calc, authority, live, activeGroups, layout, priorLayout);
  writeDetailFormulas(performance, layout);
  writeDynamicMatrix(performance, layout, priorLayout);
  writeReportsPeriodLinks(reports);
  workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.automatic);
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  validateInstalledState(workbook, performance, reports, calc, activeGroups, plannedGroups, layout);
  writeWeeklyPerformanceQA(workbook, qaSheet, live, authority, activeGroups.length);
  const protectedAfter = protectedFingerprint(workbook);
  if (protectedBefore !== protectedAfter) {
    throw new Error("PUL-030P-002: A protected source, mapping, legacy result, import, or selection table changed.");
  }
  return JSON.stringify({
    status: "PASS",
    cacheVersion: authority.cacheVersion,
    cacheFingerprint: authority.cacheFingerprint,
    current: performance.getRange("B13").getText(),
    comparison: performance.getRange("G13").getText(),
    cacheFreshness: calc.getRange("AL16").getText(),
    currentPeriodState: calc.getRange("AL17").getText(),
    comparisonPeriodState: calc.getRange("AL18").getText(),
    activeReportingGroups: activeGroups.length,
    phase2C: `${live.phase2CPassCount}/16 PASS`,
    authority: "Active weekly cache",
    legacyMetricResults: "Retained for rollback"
  });
}

type CellValue = string | number | boolean;
type LiveState = {
  mappingContentFingerprint: string;
  catalogContentFingerprint: string;
  performanceRestaurantScopeFingerprint: string;
  phase2CPassCount: number;
};

type ActiveCacheAuthority = {
  cacheVersion: string;
  cacheFingerprint: string;
  mappingContentFingerprint: string;
  catalogContentFingerprint: string;
  identityPreflightFingerprint: string;
  performanceRestaurantScopeFingerprint: string;
};

type ReportingGroup = { id: string; name: string; sortOrder: number };
type GroupSelectionRow = { id: string; name: string; include: string };
type PriorGroupState = {
  exists: boolean; capacity: number; includeById: { [key: string]: string };
  detailGroupId: string; sortGroupId: string;
};
type RuntimeLayout = {
  restaurantCapacity: number; groupCapacity: number; componentStarts: number[];
  numericDisplayStart: number; totalComponentStart: number; totalDisplayColumn: number;
  sortKeyColumn: number; sortedRestaurantIdColumn: number; helperLastColumn: number;
  componentTotalRow: number; periodKeyStartColumn: number; matrixEndColumn: number;
  performanceHelperStartColumn: number;
};
type ValidationSources = {
  yearSource: ExcelScript.Range; weekSource: ExcelScript.Range;
  detailGroupSource: ExcelScript.Range;
  sortSource: ExcelScript.Range;
};

const RESTAURANT_CAPACITY = 16;
const DASH = "—";

function readActiveReportingGroups(table: ExcelScript.Table): ReportingGroup[] {
  const h = headerMap(table);
  const rows = tableRows(table);
  const catalogIds: { [key: string]: boolean } = {};
  const activeNames: { [key: string]: boolean } = {};
  const activeSortOrders: { [key: string]: boolean } = {};
  const active: ReportingGroup[] = [];
  for (const row of rows) {
    const id = text(row[h.ReportingGroupID]);
    if (!id) throw new Error("PUL-030P-016: Reporting Group catalog contains a blank ReportingGroupID.");
    if (catalogIds[id]) throw new Error(`PUL-030P-016: Reporting Group catalog repeats ${id}.`);
    catalogIds[id] = true;
    if (text(row[h.Active]) !== "Yes") continue;
    const name = text(row[h.ReportingGroupName]);
    const sortOrder = Number(row[h.SortOrder]);
    if (!name) throw new Error(`PUL-030P-016: Active Reporting Group ${id} has a blank business name.`);
    if (activeNames[name]) throw new Error(`PUL-030P-016: Active Reporting Groups repeat business name ${name}.`);
    if (!Number.isFinite(sortOrder)) throw new Error(`PUL-030P-016: Active Reporting Group ${id} has an invalid SortOrder.`);
    const sortKey = String(sortOrder);
    if (activeSortOrders[sortKey]) throw new Error(`PUL-030P-016: Active Reporting Groups repeat SortOrder ${sortKey}.`);
    activeNames[name] = true;
    activeSortOrders[sortKey] = true;
    active.push({ id, name, sortOrder });
  }
  if (active.length < 1) throw new Error("PUL-030P-016: At least one active Reporting Group is required.");
  active.sort((left, right) => left.sortOrder - right.sortOrder || (left.id < right.id ? -1 : left.id > right.id ? 1 : 0));
  return active;
}

function capturePriorGroupState(
  workbook: ExcelScript.Workbook,
  calc: ExcelScript.Worksheet,
  performance: ExcelScript.Worksheet
): PriorGroupState {
  const table = workbook.getTable("tblPerformanceRPGSelection");
  const includeById: { [key: string]: string } = {};
  let capacity = 0;
  if (table) {
    const h = headerMap(table);
    const rows = tableRows(table);
    capacity = rows.length;
    for (const row of rows) {
      const id = text(row[h.ReportingGroupID]);
      if (!id) throw new Error("PUL-030P-017: Existing Reporting Group selection contains a blank ID.");
      if (includeById[id] !== undefined) throw new Error(`PUL-030P-017: Existing Reporting Group selection repeats ${id}.`);
      includeById[id] = text(row[h.Include]) === "Yes" ? "Yes" : "No";
    }
  }
  return {
    exists: !!table,
    capacity,
    includeById,
    detailGroupId: text(calc.getRange("AL6").getValue()),
    sortGroupId: text(calc.getRange("AL12").getValue()) || text(performance.getRange("I6").getValue())
  };
}

function planGroupSelection(groups: ReportingGroup[], prior: PriorGroupState): GroupSelectionRow[] {
  const output: GroupSelectionRow[] = [];
  for (const group of groups) {
    output.push({
      id: group.id,
      name: group.name,
      include: prior.includeById[group.id] !== undefined
        ? prior.includeById[group.id]
        : prior.exists ? "No" : "Yes"
    });
  }
  return output;
}

function makeLayout(groupCapacity: number, restaurantCapacity: number): RuntimeLayout {
  if (groupCapacity < 1 || restaurantCapacity < 1) {
    throw new Error("PUL-030P-018: Weekly Performance capacities must be positive.");
  }
  const componentStarts: number[] = [];
  for (let index = 0; index < 6; index += 1) componentStarts.push(39 + index * (groupCapacity + 1));
  const numericDisplayStart = componentStarts[5] + groupCapacity + 1;
  const totalComponentStart = numericDisplayStart + groupCapacity + 1;
  const totalDisplayColumn = totalComponentStart + 4;
  const sortKeyColumn = totalDisplayColumn + 1;
  const sortedRestaurantIdColumn = sortKeyColumn + 1;
  const periodKeyStartColumn = sortedRestaurantIdColumn + 2;
  const matrixEndColumn = groupCapacity + 1;
  return {
    restaurantCapacity,
    groupCapacity,
    componentStarts,
    numericDisplayStart,
    totalComponentStart,
    totalDisplayColumn,
    sortKeyColumn,
    sortedRestaurantIdColumn,
    helperLastColumn: periodKeyStartColumn + 1,
    componentTotalRow: restaurantCapacity + 1,
    periodKeyStartColumn,
    matrixEndColumn,
    performanceHelperStartColumn: Math.max(21, matrixEndColumn + 2)
  };
}

function validateActiveCache(
  versionTable: ExcelScript.Table,
  versionRows: CellValue[][],
  periodTable: ExcelScript.Table,
  periodRows: CellValue[][],
  scopeTable: ExcelScript.Table,
  scopeRows: CellValue[][],
  rpgTable: ExcelScript.Table,
  rpgRows: CellValue[][],
  activeGroups: ReportingGroup[]
): ActiveCacheAuthority {
  const vh = headerMap(versionTable);
  const activeRows: CellValue[][] = [];
  for (const row of versionRows) {
    if (text(row[vh.CacheStatus]) === "Active" &&
        text(row[vh.ActivationState]) === "Active") activeRows.push(row);
  }
  if (activeRows.length !== 1) {
    throw new Error(`PUL-030P-003: Weekly cache requires exactly one Active / Active authority; found ${activeRows.length}.`);
  }
  const row = activeRows[0];
  if (text(row[vh.ValidationStatus]) !== "PASS") {
    throw new Error(`PUL-030P-004: Active cache validation is ${text(row[vh.ValidationStatus])}; expected PASS.`);
  }
  const authority: ActiveCacheAuthority = {
    cacheVersion: text(row[vh.CacheVersion]),
    cacheFingerprint: text(row[vh.CacheFingerprint]),
    mappingContentFingerprint: text(row[vh.MappingContentFingerprint]),
    catalogContentFingerprint: text(row[vh.CatalogContentFingerprint]),
    identityPreflightFingerprint: text(row[vh.IdentityPreflightFingerprint]),
    performanceRestaurantScopeFingerprint: text(row[vh.PerformanceRestaurantScopeFingerprint])
  };
  const authorityFields: string[][] = [
    ["CacheVersion", authority.cacheVersion],
    ["CacheFingerprint", authority.cacheFingerprint],
    ["MappingContentFingerprint", authority.mappingContentFingerprint],
    ["CatalogContentFingerprint", authority.catalogContentFingerprint],
    ["IdentityPreflightFingerprint", authority.identityPreflightFingerprint],
    ["PerformanceRestaurantScopeFingerprint", authority.performanceRestaurantScopeFingerprint]
  ];
  for (const item of authorityFields) {
    if (!item[1]) throw new Error(`PUL-030P-004: Active cache ${item[0]} is blank.`);
  }
  const ph = headerMap(periodTable);
  const sh = headerMap(scopeTable);
  const rh = headerMap(rpgTable);
  const activePeriods = periodRows.filter(value => text(value[ph.CacheVersion]) === authority.cacheVersion);
  const activeScopes = scopeRows.filter(value => text(value[sh.CacheVersion]) === authority.cacheVersion);
  const activeRpgs = rpgRows.filter(value => text(value[rh.CacheVersion]) === authority.cacheVersion);
  if (activePeriods.length !== number(row[vh.PeriodRowCount]) ||
      activeScopes.length !== number(row[vh.ScopeCacheRowCount]) ||
      activeRpgs.length !== number(row[vh.DenseRPGCacheRowCount])) {
    throw new Error(`PUL-030P-003: Active cache row counts differ from its manifest: ${activePeriods.length}/${activeScopes.length}/${activeRpgs.length}.`);
  }
  const seen: { [key: string]: boolean } = {};
  for (const period of activePeriods) {
    const key = `${number(period[ph.ISOYear])}|${number(period[ph.ISOWeek])}`;
    if (seen[key]) throw new Error(`PUL-030P-005: Duplicate manifest ISO period ${key}.`);
    seen[key] = true;
  }
  const groupIds: { [key: string]: boolean } = {};
  for (const group of activeGroups) groupIds[group.id] = true;
  const rpgGrains: { [key: string]: boolean } = {};
  for (const rowValue of activeRpgs) {
    const groupId = text(rowValue[rh.ReportingGroupID]);
    if (!groupIds[groupId]) throw new Error(`PUL-030P-005: Active cache contains unavailable ReportingGroupID ${groupId}.`);
    const grain = `${text(rowValue[rh.SourcePeriodKey])}|${text(rowValue[rh.RestaurantID])}|${groupId}`;
    if (rpgGrains[grain]) throw new Error(`PUL-030P-005: Active cache repeats RPG grain ${grain}.`);
    rpgGrains[grain] = true;
  }
  const expectedDenseRows = activeScopes.length * activeGroups.length;
  if (activeRpgs.length !== expectedDenseRows) {
    throw new Error(`PUL-030P-005: Active cache has ${activeRpgs.length} dense RPG rows; expected ${expectedDenseRows}.`);
  }
  for (const scope of activeScopes) {
    for (const group of activeGroups) {
      const grain = `${text(scope[sh.SourcePeriodKey])}|${text(scope[sh.RestaurantID])}|${group.id}`;
      if (!rpgGrains[grain]) throw new Error(`PUL-030P-005: Active cache is missing dense RPG grain ${grain}.`);
    }
  }
  return authority;
}

function validatePhase2CLayout(
  workbook: ExcelScript.Workbook,
  calc: ExcelScript.Worksheet,
  restaurantTable: ExcelScript.Table,
  qaTable: ExcelScript.Table
): void {
  if (tableRows(requiredTable(workbook, "tblPerformanceRestaurantSelection")).length !== RESTAURANT_CAPACITY ||
      tableRows(requiredTable(workbook, "tblPerformanceRPGSelection")).length < 1 ||
      calc.getRange("AE1:DL1").getTexts()[0].join("|").indexOf("Current Numerator") < 0) {
    throw new Error("PUL-030P-006: Accepted Phase 2C helper/selection layout is missing.");
  }
  if (tableRows(restaurantTable).length < RESTAURANT_CAPACITY || tableRows(qaTable).length !== 16) {
    throw new Error("PUL-030P-007: Restaurant catalog or Phase 2C QA is incomplete.");
  }
}

function writeValidationLists(
  performance: ExcelScript.Worksheet,
  periodRows: CellValue[][],
  activeCacheVersion: string,
  groups: ReportingGroup[],
  priorGroupCapacity: number,
  layout: RuntimeLayout
): ValidationSources {
  const years: number[] = [];
  const seenYears: { [key: string]: boolean } = {};
  for (const row of periodRows) {
    if (text(row[1]) !== activeCacheVersion) continue;
    const year = number(row[5]);
    if (!seenYears[String(year)]) { seenYears[String(year)] = true; years.push(year); }
  }
  years.sort((left, right) => left - right);
  const yearValues: (string | number | boolean)[][] = [];
  for (const year of years) yearValues.push([year]);
  const weekValues: (string | number | boolean)[][] = [];
  for (let week = 1; week <= 53; week += 1) weekValues.push([`W${String(week).padStart(2, "0")}`]);
  const start = layout.performanceHelperStartColumn;
  const priorRows = Math.max(54, priorGroupCapacity + 2, groups.length + 2);
  performance.getRangeByIndexes(0, start, priorRows, 4).clear(ExcelScript.ClearApplyTo.contents);
  performance.getRangeByIndexes(0, start, 1, 4).setValues([[
    "Available year", "ISO week", "Detail Reporting Group", "Sort target"
  ]]);
  performance.getRangeByIndexes(1, start, yearValues.length, 1).setValues(yearValues);
  performance.getRangeByIndexes(1, start + 1, weekValues.length, 1).setValues(weekValues);
  performance.getRangeByIndexes(1, start + 2, groups.length, 1)
    .setValues(groups.map(group => [group.name]));
  performance.getRangeByIndexes(1, start + 3, groups.length + 1, 1)
    .setValues([["Total"]].concat(groups.map(group => [group.name])));
  performance.getRangeByIndexes(0, start, 1, 4).getEntireColumn().setColumnHidden(true);
  return {
    yearSource: performance.getRangeByIndexes(1, start, yearValues.length, 1),
    weekSource: performance.getRangeByIndexes(1, start + 1, weekValues.length, 1),
    detailGroupSource: performance.getRangeByIndexes(1, start + 2, groups.length, 1),
    sortSource: performance.getRangeByIndexes(1, start + 3, groups.length + 1, 1)
  };
}

function writePeriodControls(
  performance: ExcelScript.Worksheet,
  alreadyWeekly: boolean,
  priorCurrent: CellValue[][],
  priorComparison: CellValue[][],
  yearSource: ExcelScript.Range,
  weekSource: ExcelScript.Range
): void {
  performance.getRange("A9:D13").clear(ExcelScript.ClearApplyTo.contents);
  performance.getRange("F9:I13").clear(ExcelScript.ClearApplyTo.contents);
  performance.getRange("A9").setValue("Current");
  performance.getRange("F9").setValue("Compare with");
  performance.getRange("A10:A12").setValues([["Year"], ["From week"], ["To week"]]);
  performance.getRange("F10:F12").setValues([["Year"], ["From week"], ["To week"]]);
  if (alreadyWeekly) {
    performance.getRange("B10:B12").setValues(priorCurrent);
    performance.getRange("G10:G12").setValues(priorComparison);
  } else {
    performance.getRange("B10:B12").setValues([[2026], ["W01"], ["W32"]]);
    performance.getRange("G10:G12").setValues([[2025], ["W01"], ["W32"]]);
  }
  performance.getRange("B13").setFormula("=_Metric_Calc!$AL$21");
  performance.getRange("G13").setFormula("=_Metric_Calc!$AL$22");
  performance.getRange("A13").setValue("Selected period");
  performance.getRange("F13").setValue("Selected period");
  performance.getRange("B13:D13").getFormat().setWrapText(true);
  performance.getRange("G13:I13").getFormat().setWrapText(true);
  performance.getRange("10:13").getFormat().setRowHeight(24);
  performance.getRange("B10:B12").getFormat().getFill().setColor("#FFF4D6");
  performance.getRange("G10:G12").getFormat().getFill().setColor("#FFF4D6");
  performance.getRange("B10:B12").getFormat().getFont().setBold(true);
  performance.getRange("G10:G12").getFormat().getFont().setBold(true);
  performance.getRange("B13:D13").getFormat().getFill().setColor("#EAF2FF");
  performance.getRange("G13:I13").getFormat().getFill().setColor("#EAF2FF");
  applyRangeValidation(performance.getRange("B10"), yearSource);
  applyRangeValidation(performance.getRange("G10"), yearSource);
  applyRangeValidation(performance.getRange("B11"), weekSource);
  applyRangeValidation(performance.getRange("B12"), weekSource);
  applyRangeValidation(performance.getRange("G11"), weekSource);
  applyRangeValidation(performance.getRange("G12"), weekSource);
  performance.getRange("A2").setValue(
    "Interactive Reporting Group Sales Share — restaurant, Reporting Group, and weekly period selections recalculate without Office Scripts."
  );
  performance.getRange("G16").setFormula(
    '=IF(OR(_Metric_Calc!$AL$17<>"Valid",_Metric_Calc!$AL$18<>"Valid"),"Unavailable",IF(_Metric_Calc!$AL$25=_Metric_Calc!$AL$26,"Same","Different"))'
  );
  performance.getRange("G17").setFormula('=IF(AND($B$10=$G$10,$B$11=$G$11,$B$12=$G$12),"Yes","No")');
  performance.getRange("G18").setFormula("=_Metric_Calc!$AL$23");
}

function writeGroupSelectionTable(workbook: ExcelScript.Workbook, rows: GroupSelectionRow[]): void {
  const table = requiredTable(workbook, "tblPerformanceRPGSelection");
  const sheet = table.getWorksheet();
  const oldRange = table.getRange();
  const startRow = oldRange.getRowIndex();
  const startColumn = oldRange.getColumnIndex();
  const oldRowCount = oldRange.getRowCount();
  const target = sheet.getRangeByIndexes(startRow, startColumn, rows.length + 1, 3);
  table.resize(target);
  table.getHeaderRowRange().setValues([["Include", "Reporting Group", "ReportingGroupID"]]);
  table.getRangeBetweenHeaderAndTotal().setValues(rows.map(row => [row.include, row.name, row.id]));
  if (oldRowCount > rows.length + 1) {
    sheet.getRangeByIndexes(startRow + rows.length + 1, startColumn, oldRowCount - rows.length - 1, 3)
      .clear(ExcelScript.ClearApplyTo.contents);
  }
  const include = table.getColumnByName("Include");
  if (!include) throw new Error("PUL-030P-019: Reporting Group selection Include column is missing.");
  const validation = include.getRangeBetweenHeaderAndTotal().getDataValidation();
  validation.clear();
  validation.setRule({ list: { inCellDropDown: true, source: "Yes,No" } });
  validation.setErrorAlert({ showAlert: true, style: ExcelScript.DataValidationAlertStyle.stop,
    title: "Choose Yes or No", message: "Use Yes to include this Reporting Group in Performance." });
}

function writeGroupControls(
  performance: ExcelScript.Worksheet,
  groups: ReportingGroup[],
  prior: PriorGroupState,
  layout: RuntimeLayout,
  sources: ValidationSources
): void {
  const detail = groups.find(group => group.id === prior.detailGroupId || group.name === prior.detailGroupId) || groups[0];
  const sort = groups.find(group => group.id === prior.sortGroupId || group.name === prior.sortGroupId);
  performance.getRange("B7").setValue(detail.name);
  performance.getRange("I6").setValue(sort ? sort.name : "Total");
  applyRangeValidation(performance.getRange("B7"), sources.detailGroupSource);
  applyRangeValidation(performance.getRange("I6"), sources.sortSource);
  performance.getRange("G8").setFormula(
    `=IF('_Metric_Calc'!$AL$5=${layout.groupCapacity},"All ${layout.groupCapacity} Reporting Groups selected",` +
      `'_Metric_Calc'!$AL$5&" of ${layout.groupCapacity} Reporting Groups selected")`
  );
  performance.getRange("I8").setFormula("='_Metric_Calc'!$AL$15");
}

function writeWeeklyRuntimeCalc(
  calc: ExcelScript.Worksheet,
  authority: ActiveCacheAuthority,
  live: LiveState,
  groups: ReportingGroup[],
  layout: RuntimeLayout,
  priorLayout: RuntimeLayout
): void {
  const clearLastColumn = Math.max(layout.helperLastColumn, priorLayout.helperLastColumn);
  calc.getRangeByIndexes(0, 30, 54, clearLastColumn - 30 + 1).clear(ExcelScript.ClearApplyTo.all);
  const groupClearRows = Math.max(layout.groupCapacity, priorLayout.groupCapacity) + 1;
  calc.getRangeByIndexes(0, 8, groupClearRows, 2).clear(ExcelScript.ClearApplyTo.contents);
  calc.getRange("I1:J1").setValues([["Reporting Group", "ReportingGroupID"]]);
  calc.getRangeByIndexes(1, 8, groups.length, 2).setValues(groups.map(group => [group.name, group.id]));

  calc.getRange("AE1:AF1").setValues([["Selected Restaurant", "Selected RestaurantID"]]);
  const selectedRestaurantFormulas: string[][] = [];
  for (let index = 0; index < layout.restaurantCapacity; index += 1) {
    const excelRow = index + 2;
    selectedRestaurantFormulas.push([
      `=IFERROR(INDEX(FILTER(tblPerformanceRestaurantSelection[Restaurant],tblPerformanceRestaurantSelection[Include]="Yes"),ROWS($AE$2:AE${excelRow})),"")`,
      `=IFERROR(INDEX(FILTER(tblPerformanceRestaurantSelection[RestaurantID],tblPerformanceRestaurantSelection[Include]="Yes"),ROWS($AF$2:AF${excelRow})),"")`
    ]);
  }
  calc.getRangeByIndexes(1, 30, layout.restaurantCapacity, 2).setFormulas(selectedRestaurantFormulas);

  calc.getRange("AH1:AI1").setValues([["Selected Reporting Group", "Selected ReportingGroupID"]]);
  const selectedGroupFormulas: string[][] = [];
  for (let index = 0; index < layout.groupCapacity; index += 1) {
    const excelRow = index + 2;
    selectedGroupFormulas.push([
      `=IFERROR(INDEX(FILTER(tblPerformanceRPGSelection[Reporting Group],tblPerformanceRPGSelection[Include]="Yes"),ROWS($AH$2:AH${excelRow})),"")`,
      `=IFERROR(INDEX(FILTER(tblPerformanceRPGSelection[ReportingGroupID],tblPerformanceRPGSelection[Include]="Yes"),ROWS($AI$2:AI${excelRow})),"")`
    ]);
  }
  calc.getRangeByIndexes(1, 33, layout.groupCapacity, 2).setFormulas(selectedGroupFormulas);

  writeWeeklyControl(calc, authority, live, layout);
  writePeriodKeyHelpers(calc, layout);
  writeRuntimeHeaders(calc, groups, layout);
  writeWeeklyComponentBlock(calc, layout, 0, true, true);
  writeWeeklyComponentBlock(calc, layout, 1, true, false);
  writeShareBlock(calc, layout, 2, 0, 1);
  writeWeeklyComponentBlock(calc, layout, 3, false, true);
  writeWeeklyComponentBlock(calc, layout, 4, false, false);
  writeShareBlock(calc, layout, 5, 3, 4);
  writeNumericDisplayBlock(calc, layout);
  writeTotalAndSortHelpers(calc, layout);
  writeManagementOverviewAuthority(calc, layout);

  calc.getRangeByIndexes(1, layout.componentStarts[0], layout.restaurantCapacity + 1, layout.groupCapacity)
    .setNumberFormat('#,##0.00 "NOK"');
  calc.getRangeByIndexes(1, layout.componentStarts[1], layout.restaurantCapacity + 1, layout.groupCapacity)
    .setNumberFormat('#,##0.00 "NOK"');
  calc.getRangeByIndexes(1, layout.componentStarts[2], layout.restaurantCapacity + 1, layout.groupCapacity)
    .setNumberFormat("0.00%");
  calc.getRangeByIndexes(1, layout.componentStarts[3], layout.restaurantCapacity + 1, layout.groupCapacity)
    .setNumberFormat('#,##0.00 "NOK"');
  calc.getRangeByIndexes(1, layout.componentStarts[4], layout.restaurantCapacity + 1, layout.groupCapacity)
    .setNumberFormat('#,##0.00 "NOK"');
  calc.getRangeByIndexes(1, layout.componentStarts[5], layout.restaurantCapacity + 1, layout.groupCapacity)
    .setNumberFormat("0.00%");
}

function writeWeeklyControl(
  calc: ExcelScript.Worksheet,
  authority: ActiveCacheAuthority,
  live: LiveState,
  layout: RuntimeLayout
): void {
  calc.getRange("AK1:AL30").setValues([
    ["Weekly Performance control", "Value"],
    ["Current period", ""], ["Comparison period", ""],
    ["Selected restaurant count", ""], ["Selected RPG count", ""],
    ["Detail ReportingGroupID", ""], ["Matrix display", ""],
    ["Company scope fingerprint", authority.performanceRestaurantScopeFingerprint],
    ["Interaction contract", "Weekly cache additive components only"],
    ["Restaurant capacity", layout.restaurantCapacity], ["RPG capacity", layout.groupCapacity],
    ["Requested sort ReportingGroupID", ""], ["Effective sort ReportingGroupID", ""],
    ["Sort order", ""], ["Sort status", ""],
    ["Active-cache freshness", ""], ["Current period state", ""],
    ["Comparison period state", ""], ["Current available weeks", ""],
    ["Comparison available weeks", ""], ["Current summary", ""],
    ["Comparison summary", ""], ["Period status", ""],
    ["Active CacheVersion", ""], ["Current expected weeks", ""],
    ["Comparison expected weeks", ""],
    ["Current MappingContentFingerprint", live.mappingContentFingerprint],
    ["Current CatalogContentFingerprint", live.catalogContentFingerprint],
    ["Current ReportingEnabled fingerprint", live.performanceRestaurantScopeFingerprint],
    ["Accepted IdentityPreflightFingerprint", authority.identityPreflightFingerprint]
  ]);
  const groupEnd = layout.groupCapacity + 1;
  calc.getRange("AL2").setFormula("=Performance!$B$13");
  calc.getRange("AL3").setFormula("=Performance!$G$13");
  calc.getRange("AL4").setFormula('=COUNTIF(tblPerformanceRestaurantSelection[Include],"Yes")');
  calc.getRange("AL5").setFormula('=COUNTIF(tblPerformanceRPGSelection[Include],"Yes")');
  calc.getRange("AL6").setFormula(`=IFERROR(INDEX($J$2:$J$${groupEnd},MATCH(Performance!$B$7,$I$2:$I$${groupEnd},0)),"")`);
  calc.getRange("AL7").setFormula("=Performance!$G$6");
  calc.getRange("AL12").setFormula(`=IF(Performance!$I$6="Total","",IFERROR(INDEX($J$2:$J$${groupEnd},MATCH(Performance!$I$6,$I$2:$I$${groupEnd},0)),""))`);
  calc.getRange("AL13").setFormula(`=IF(OR($AL$12="",COUNTIF($AI$2:$AI$${groupEnd},$AL$12)=0),"",$AL$12)`);
  calc.getRange("AL14").setFormula("=Performance!$I$7");
  calc.getRange("AL15").setFormula('=IF(Performance!$I$6="Total","",IF($AL$12="","Using Total — target unavailable",IF($AL$13="","Using Total — "&Performance!$I$6&" hidden","")))');
  calc.getRange("AL16").setFormula(cacheFreshnessFormula());
  calc.getRange("AL17").setFormula(periodStateFormula("$B$10", "$B$11", "$B$12"));
  calc.getRange("AL18").setFormula(periodStateFormula("$G$10", "$G$11", "$G$12"));
  calc.getRange("AL19").setFormula(periodCountFormula("$B$10", "$B$11", "$B$12"));
  calc.getRange("AL20").setFormula(periodCountFormula("$G$10", "$G$11", "$G$12"));
  calc.getRange("AL21").setFormula(periodSummaryFormula("$B$10", "$B$11", "$B$12", "$AL$17", "$AL$19", "$AL$25"));
  calc.getRange("AL22").setFormula(periodSummaryFormula("$G$10", "$G$11", "$G$12", "$AL$18", "$AL$20", "$AL$26"));
  calc.getRange("AL23").setFormula(periodWarningFormula());
  calc.getRange("AL24").setFormula(activeCacheVersionFormula());
  calc.getRange("AL25").setFormula('=IFERROR(VALUE(RIGHT(Performance!$B$12,2))-VALUE(RIGHT(Performance!$B$11,2))+1,0)');
  calc.getRange("AL26").setFormula('=IFERROR(VALUE(RIGHT(Performance!$G$12,2))-VALUE(RIGHT(Performance!$G$11,2))+1,0)');
}

function writePeriodKeyHelpers(calc: ExcelScript.Worksheet, layout: RuntimeLayout): void {
  const currentColumn = columnName(layout.periodKeyStartColumn);
  const comparisonColumn = columnName(layout.periodKeyStartColumn + 1);
  calc.getRangeByIndexes(0, layout.periodKeyStartColumn, 1, 2)
    .setValues([["Current SourcePeriodKey", "Comparison SourcePeriodKey"]]);
  const formulas: string[][] = [];
  for (let index = 0; index < 53; index += 1) {
    const row = index + 2;
    formulas.push([
      periodKeyFormula(row, "$AL$17", "$AL$19", "Performance!$B$10", "Performance!$B$11", "Performance!$B$12", currentColumn),
      periodKeyFormula(row, "$AL$18", "$AL$20", "Performance!$G$10", "Performance!$G$11", "Performance!$G$12", comparisonColumn)
    ]);
  }
  calc.getRangeByIndexes(1, layout.periodKeyStartColumn, 53, 2).setFormulas(formulas);
}

function writeRuntimeHeaders(calc: ExcelScript.Worksheet, groups: ReportingGroup[], layout: RuntimeLayout): void {
  const names = ["Current Numerator", "Current Denominator", "Current Share", "Comparison Numerator", "Comparison Denominator", "Comparison Share"];
  for (let index = 0; index < names.length; index += 1) {
    calc.getRangeByIndexes(0, layout.componentStarts[index] - 1, 1, layout.groupCapacity + 1)
      .setValues([[`${names[index]} | RPG-ID`].concat(groups.map(group => group.id))]);
  }
  calc.getRangeByIndexes(0, layout.numericDisplayStart - 1, 1, layout.groupCapacity + 1)
    .setValues([["Selected Numeric Display | RPG-ID"].concat(groups.map(group => group.id))]);
  calc.getRangeByIndexes(0, layout.totalComponentStart, 1, 7).setValues([[
    "Selected RPG Current Numerator", "Current Scope Denominator",
    "Selected RPG Comparison Numerator", "Comparison Scope Denominator",
    "Selected RPG Total Numeric Display", "Restaurant Numeric Sort Key", "Sorted RestaurantID"
  ]]);
}

function writeWeeklyComponentBlock(
  calc: ExcelScript.Worksheet,
  layout: RuntimeLayout,
  blockIndex: number,
  current: boolean,
  numerator: boolean
): void {
  const startIndex = layout.componentStarts[blockIndex];
  const periodState = current ? "$AL$17" : "$AL$18";
  const periodColumn = columnName(layout.periodKeyStartColumn + (current ? 0 : 1));
  const periodKeys = `$${periodColumn}$2:$${periodColumn}$54`;
  const table = numerator ? "tblWeeklyRPGCache" : "tblWeeklyScopeCache";
  const value = numerator ? "MappedSalesNOK" : "SourceSalesNOK";
  const formulas: string[][] = [];
  for (let restaurant = 0; restaurant < layout.restaurantCapacity; restaurant += 1) {
    const excelRow = restaurant + 2;
    const formulaRow: string[] = [];
    for (let group = 0; group < layout.groupCapacity; group += 1) {
      const groupRow = group + 2;
      let formula = `=IF(OR($AF${excelRow}="",$AL$16<>"Available",${periodState}<>"Valid"),0,IFERROR(SUM(SUMIFS(${table}[${value}],${table}[CacheVersion],$AL$24,${table}[RestaurantID],$AF${excelRow},`;
      if (numerator) formula += `${table}[ReportingGroupID],$J$${groupRow},`;
      formula += `${table}[SourcePeriodKey],FILTER(${periodKeys},${periodKeys}<>""))),0))`;
      formulaRow.push(formula);
    }
    formulas.push(formulaRow);
  }
  calc.getRangeByIndexes(1, startIndex, layout.restaurantCapacity, layout.groupCapacity).setFormulas(formulas);
  const totals: string[][] = [[]];
  for (let group = 0; group < layout.groupCapacity; group += 1) {
    const column = columnName(startIndex + group);
    totals[0].push(`=SUM(${column}$2:${column}$${layout.restaurantCapacity + 1})`);
  }
  calc.getRangeByIndexes(layout.componentTotalRow, startIndex, 1, layout.groupCapacity).setFormulas(totals);
}

function writeShareBlock(
  calc: ExcelScript.Worksheet,
  layout: RuntimeLayout,
  shareBlockIndex: number,
  numeratorBlockIndex: number,
  denominatorBlockIndex: number
): void {
  const formulas: string[][] = [];
  for (let rowIndex = 0; rowIndex < layout.restaurantCapacity + 1; rowIndex += 1) {
    const excelRow = rowIndex + 2;
    const formulaRow: string[] = [];
    for (let group = 0; group < layout.groupCapacity; group += 1) {
      const numerator = componentCellReference(layout, numeratorBlockIndex, excelRow, group);
      const denominator = componentCellReference(layout, denominatorBlockIndex, excelRow, group);
      formulaRow.push(`=IF(${denominator}=0,0,${numerator}/${denominator})`);
    }
    formulas.push(formulaRow);
  }
  calc.getRangeByIndexes(1, layout.componentStarts[shareBlockIndex], layout.restaurantCapacity + 1, layout.groupCapacity)
    .setFormulas(formulas);
}

function writeNumericDisplayBlock(calc: ExcelScript.Worksheet, layout: RuntimeLayout): void {
  const formulas: string[][] = [];
  for (let rowIndex = 0; rowIndex < layout.restaurantCapacity + 1; rowIndex += 1) {
    const excelRow = rowIndex + 2;
    const formulaRow: string[] = [];
    for (let group = 0; group < layout.groupCapacity; group += 1) {
      const currentNumerator = componentCellReference(layout, 0, excelRow, group);
      const currentDenominator = componentCellReference(layout, 1, excelRow, group);
      const currentShare = componentCellReference(layout, 2, excelRow, group);
      const comparisonNumerator = componentCellReference(layout, 3, excelRow, group);
      const comparisonDenominator = componentCellReference(layout, 4, excelRow, group);
      const comparisonShare = componentCellReference(layout, 5, excelRow, group);
      formulaRow.push(
        `=IF($AL$7="Current Sales NOK",IF(${currentDenominator}=0,"",${currentNumerator}),` +
          `IF($AL$7="Current Share",IF(${currentDenominator}=0,"",${currentShare}),` +
          `IF($AL$7="Comparison Share",IF(${comparisonDenominator}=0,"",${comparisonShare}),` +
          `IF($AL$7="PP Change",IF(OR(${currentDenominator}=0,${comparisonDenominator}=0),"",(${currentShare}-${comparisonShare})*100),` +
          `IF($AL$7="NOK Impact",IF(OR(${currentDenominator}=0,${comparisonDenominator}=0),"",${currentNumerator}-((${comparisonNumerator}/${comparisonDenominator})*${currentDenominator})),"")))))`
      );
    }
    formulas.push(formulaRow);
  }
  calc.getRangeByIndexes(1, layout.numericDisplayStart, layout.restaurantCapacity + 1, layout.groupCapacity)
    .setFormulas(formulas);
}

function writeTotalAndSortHelpers(calc: ExcelScript.Worksheet, layout: RuntimeLayout): void {
  const currentNumeratorColumn = columnName(layout.totalComponentStart);
  const currentDenominatorColumn = columnName(layout.totalComponentStart + 1);
  const comparisonNumeratorColumn = columnName(layout.totalComponentStart + 2);
  const comparisonDenominatorColumn = columnName(layout.totalComponentStart + 3);
  const totalDisplayColumn = columnName(layout.totalDisplayColumn);
  const sortKeyColumn = columnName(layout.sortKeyColumn);
  const sortedIdColumn = columnName(layout.sortedRestaurantIdColumn);
  const formulas: string[][] = [];
  for (let rowIndex = 0; rowIndex < layout.restaurantCapacity + 1; rowIndex += 1) {
    const excelRow = rowIndex + 2;
    const cn = `$${currentNumeratorColumn}${excelRow}`;
    const cd = `$${currentDenominatorColumn}${excelRow}`;
    const pn = `$${comparisonNumeratorColumn}${excelRow}`;
    const pd = `$${comparisonDenominatorColumn}${excelRow}`;
    formulas.push([
      `=IF($AL$5=0,0,${selectedNumeratorExpression(layout, 0, excelRow)})`,
      `=${componentCellReference(layout, 1, excelRow, 0)}`,
      `=IF($AL$5=0,0,${selectedNumeratorExpression(layout, 3, excelRow)})`,
      `=${componentCellReference(layout, 4, excelRow, 0)}`,
      `=IF($AL$5=0,"",IF($AL$7="Current Sales NOK",${cn},IF($AL$7="Current Share",IF(${cd}=0,"",${cn}/${cd}),IF($AL$7="Comparison Share",IF(${pd}=0,"",${pn}/${pd}),IF($AL$7="PP Change",IF(OR(${cd}=0,${pd}=0),"",((${cn}/${cd})-(${pn}/${pd}))*100),IF($AL$7="NOK Impact",IF(OR(${cd}=0,${pd}=0),"",${cn}-((${pn}/${pd})*${cd})),""))))))`
    ]);
  }
  calc.getRangeByIndexes(1, layout.totalComponentStart, layout.restaurantCapacity + 1, 5).setFormulas(formulas);

  const idRange = `$AF$2:$AF$${layout.restaurantCapacity + 1}`;
  const keyRange = `$${sortKeyColumn}$2:$${sortKeyColumn}$${layout.restaurantCapacity + 1}`;
  const numericFirst = columnName(layout.numericDisplayStart);
  const numericLast = columnName(layout.numericDisplayStart + layout.groupCapacity - 1);
  const sortFormulas: string[][] = [];
  const sortedIdFormulas: string[][] = [];
  for (let rowIndex = 0; rowIndex < layout.restaurantCapacity; rowIndex += 1) {
    const excelRow = rowIndex + 2;
    sortFormulas.push([
      `=IF($AF${excelRow}="","",IF($AL$13="",$${totalDisplayColumn}${excelRow},IFERROR(INDEX($${numericFirst}${excelRow}:$${numericLast}${excelRow},1,MATCH($AL$13,$J$2:$J$${layout.groupCapacity + 1},0)),"")))`
    ]);
    sortedIdFormulas.push([
      `=IFERROR(INDEX(SORTBY(FILTER(${idRange},${idRange}<>""),FILTER(--(${keyRange}=""),${idRange}<>""),1,FILTER(IF(${keyRange}="",0,${keyRange}),${idRange}<>""),IF($AL$14="Highest first",-1,1),FILTER(IF(${keyRange}="","",${idRange}),${idRange}<>""),IF($AL$14="Highest first",1,-1),FILTER(${idRange},${idRange}<>""),1),ROWS($${sortedIdColumn}$2:${sortedIdColumn}${excelRow})),"")`
    ]);
  }
  calc.getRangeByIndexes(1, layout.sortKeyColumn, layout.restaurantCapacity, 1).setFormulas(sortFormulas);
  calc.getRangeByIndexes(1, layout.sortedRestaurantIdColumn, layout.restaurantCapacity, 1).setFormulas(sortedIdFormulas);
}

/**
 * Stable upstream outputs for management presentation surfaces.
 *
 * These formulas remain part of the accepted Performance authority: they use
 * the same selected-scope additive components, display metric, freshness gate,
 * and full-precision ranking semantics as Performance. Overview may project
 * these cells, but must not reproduce their calculations.
 */
function writeManagementOverviewAuthority(calc: ExcelScript.Worksheet, layout: RuntimeLayout): void {
  const totalRow = layout.componentTotalRow + 1;
  const currentNumerator = `$${columnName(layout.totalComponentStart)}$${totalRow}`;
  const currentDenominator = `$${columnName(layout.totalComponentStart + 1)}$${totalRow}`;
  const comparisonNumerator = `$${columnName(layout.totalComponentStart + 2)}$${totalRow}`;
  const comparisonDenominator = `$${columnName(layout.totalComponentStart + 3)}$${totalRow}`;

  calc.getRange("AK32:AL45").setValues([
    ["Management Performance availability", ""],
    ["Current period", ""],
    ["Comparison period", ""],
    ["Period status", ""],
    ["Latest published", ""],
    ["Selection recency", ""],
    ["Display metric", ""],
    ["Restaurant scope", ""],
    ["Reporting Group scope", ""],
    ["Total Sales", ""],
    ["Selected Category Sales", ""],
    ["Sales Share", ""],
    ["PP Change", ""],
    ["NOK Impact", ""]
  ]);
  calc.getRange("AL32").setFormula('=IF(AND($AL$16="Available",$AL$17="Valid",$AL$18="Valid",$AL$4>0,$AL$5>0),"Available","Unavailable")');
  calc.getRange("AL33").setFormula("=$AL$21");
  calc.getRange("AL34").setFormula("=$AL$22");
  calc.getRange("AL35").setFormula("=$AL$23");
  calc.getRange("AL36").setFormula(latestPublishedPeriodFormula());
  calc.getRange("AL37").setFormula(selectionRecencyFormula());
  calc.getRange("AL38").setFormula("=$AL$7");
  calc.getRange("AL39").setFormula('=IF($AL$4=$AL$10,"All "&$AL$10&" restaurants",$AL$4&" of "&$AL$10&" restaurants")');
  calc.getRange("AL40").setFormula('=IF($AL$5=$AL$11,"All "&$AL$11&" Reporting Groups",$AL$5&" of "&$AL$11&" Reporting Groups")');
  calc.getRange("AL41").setFormula(`=IF($AL$32="Available",${currentDenominator},"")`);
  calc.getRange("AL42").setFormula(`=IF($AL$32="Available",${currentNumerator},"")`);
  calc.getRange("AL43").setFormula(`=IF(OR($AL$32<>"Available",${currentDenominator}=0),"",${currentNumerator}/${currentDenominator})`);
  calc.getRange("AL44").setFormula(`=IF(OR($AL$32<>"Available",${currentDenominator}=0,${comparisonDenominator}=0),"",((${currentNumerator}/${currentDenominator})-(${comparisonNumerator}/${comparisonDenominator}))*100)`);
  calc.getRange("AL45").setFormula(`=IF(OR($AL$32<>"Available",${currentDenominator}=0,${comparisonDenominator}=0),"",${currentNumerator}-((${comparisonNumerator}/${comparisonDenominator})*${currentDenominator}))`);
  calc.getRange("AL41:AL42").setNumberFormat('#,##0.00 "NOK"');
  calc.getRange("AL43").setNumberFormat("0.00%");
  calc.getRange("AL44").setNumberFormat('+0.00 "pp";-0.00 "pp";0.00 "pp"');
  calc.getRange("AL45").setNumberFormat('+#,##0.00 "NOK";-#,##0.00 "NOK";0.00 "NOK"');

  writeGroupManagementRankings(calc, layout);
  writeRestaurantManagementRankings(calc, layout);
}

function writeGroupManagementRankings(calc: ExcelScript.Worksheet, layout: RuntimeLayout): void {
  calc.getRange("AN32:AQ38").setValues([
    ["Position", "ReportingGroupID", "Reporting Group", "Numeric value"],
    ["Top 1", "", "", ""], ["Top 2", "", "", ""], ["Top 3", "", "", ""],
    ["Bottom 1", "", "", ""], ["Bottom 2", "", "", ""], ["Bottom 3", "", "", ""]
  ]);
  const catalogEnd = layout.groupCapacity + 1;
  const valueFirst = columnName(layout.numericDisplayStart);
  const valueLast = columnName(layout.numericDisplayStart + layout.groupCapacity - 1);
  const valueRange = `$${valueFirst}$${layout.componentTotalRow + 1}:$${valueLast}$${layout.componentTotalRow + 1}`;
  const idRange = `$J$2:$J$${catalogEnd}`;
  const selectedRange = `$AI$2:$AI$${catalogEnd}`;
  const nameRange = `$I$2:$I$${catalogEnd}`;
  const ids: string[][] = [];
  for (let index = 0; index < 6; index += 1) {
    const rank = index < 3 ? index + 1 : index - 2;
    const descending = index < 3;
    ids.push([groupRankingIdFormula(idRange, selectedRange, valueRange, rank, descending)]);
  }
  calc.getRange("AO33:AO38").setFormulas(ids);
  calc.getRange("AP33").setFormula(`=IF($AO33="","",XLOOKUP($AO33,${idRange},${nameRange},""))`);
  calc.getRange("AP33:AP38").fillDown();
  calc.getRange("AQ33").setFormula(`=IF($AO33="","",INDEX(${valueRange},1,MATCH($AO33,${idRange},0)))`);
  calc.getRange("AQ33:AQ38").fillDown();
}

function writeRestaurantManagementRankings(calc: ExcelScript.Worksheet, layout: RuntimeLayout): void {
  calc.getRange("AS32:AV38").setValues([
    ["Position", "RestaurantID", "Restaurant", "Numeric value"],
    ["Top 1", "", "", ""], ["Top 2", "", "", ""], ["Top 3", "", "", ""],
    ["Bottom 1", "", "", ""], ["Bottom 2", "", "", ""], ["Bottom 3", "", "", ""]
  ]);
  const idRange = `$AF$2:$AF$${layout.restaurantCapacity + 1}`;
  const nameRange = `$AE$2:$AE$${layout.restaurantCapacity + 1}`;
  const valueColumn = columnName(layout.totalDisplayColumn);
  const valueRange = `$${valueColumn}$2:$${valueColumn}$${layout.restaurantCapacity + 1}`;
  const ids: string[][] = [];
  for (let index = 0; index < 6; index += 1) {
    const rank = index < 3 ? index + 1 : index - 2;
    const descending = index < 3;
    ids.push([restaurantRankingIdFormula(idRange, valueRange, rank, descending)]);
  }
  calc.getRange("AT33:AT38").setFormulas(ids);
  calc.getRange("AU33").setFormula(`=IF($AT33="","",XLOOKUP($AT33,${idRange},${nameRange},""))`);
  calc.getRange("AU33:AU38").fillDown();
  calc.getRange("AV33").setFormula(`=IF($AT33="","",INDEX(${valueRange},MATCH($AT33,${idRange},0)))`);
  calc.getRange("AV33:AV38").fillDown();
}

function groupRankingIdFormula(
  idRange: string, selectedRange: string, valueRange: string, rank: number, descending: boolean
): string {
  const valueDirection = descending ? -1 : 1;
  const idDirection = descending ? 1 : -1;
  return `=LET(ids,FILTER(${idRange},COUNTIF(${selectedRange},${idRange})>0),vals,FILTER(TRANSPOSE(${valueRange}),COUNTIF(${selectedRange},${idRange})>0),ordered,SORTBY(ids,--(vals=""),1,IF(vals="",0,vals),${valueDirection},ids,${idDirection}),IF(OR($AL$32<>"Available",${rank}>MIN(3,$AL$5)),"",IFERROR(INDEX(ordered,${rank}),"")))`;
}

function restaurantRankingIdFormula(
  idRange: string, valueRange: string, rank: number, descending: boolean
): string {
  const valueDirection = descending ? -1 : 1;
  const idDirection = descending ? 1 : -1;
  return `=LET(ids,FILTER(${idRange},${idRange}<>""),vals,FILTER(${valueRange},${idRange}<>""),ordered,SORTBY(ids,--(vals=""),1,IF(vals="",0,vals),${valueDirection},ids,${idDirection}),IF(OR($AL$32<>"Available",${rank}>MIN(3,$AL$4)),"",IFERROR(INDEX(ordered,${rank}),"")))`;
}

function latestPublishedPeriodFormula(): string {
  return '=IF($AL$16<>"Available","",LET(y,MAX(FILTER(tblWeeklyPeriodManifest[ISOYear],tblWeeklyPeriodManifest[CacheVersion]=$AL$24)),w,MAX(FILTER(tblWeeklyPeriodManifest[ISOWeek],(tblWeeklyPeriodManifest[CacheVersion]=$AL$24)*(tblWeeklyPeriodManifest[ISOYear]=y))),y&" W"&TEXT(w,"00")))';
}

function selectionRecencyFormula(): string {
  return '=IF($AL$16<>"Available","",LET(y,MAX(FILTER(tblWeeklyPeriodManifest[ISOYear],tblWeeklyPeriodManifest[CacheVersion]=$AL$24)),w,MAX(FILTER(tblWeeklyPeriodManifest[ISOWeek],(tblWeeklyPeriodManifest[CacheVersion]=$AL$24)*(tblWeeklyPeriodManifest[ISOYear]=y))),IF(OR(Performance!$B$10<y,AND(Performance!$B$10=y,VALUE(RIGHT(Performance!$B$12,2))<w)),"Newer week available","")))';
}

function selectedNumeratorExpression(layout: RuntimeLayout, blockIndex: number, excelRow: number): string {
  const terms: string[] = [];
  for (let group = 0; group < layout.groupCapacity; group += 1) {
    terms.push(`IF(COUNTIF($AI$2:$AI$${layout.groupCapacity + 1},$J$${group + 2})>0,${componentCellReference(layout, blockIndex, excelRow, group)},0)`);
  }
  return terms.join("+") || "0";
}

function componentCellReference(layout: RuntimeLayout, blockIndex: number, excelRow: number, groupIndex: number): string {
  return `$${columnName(layout.componentStarts[blockIndex] + groupIndex)}${excelRow}`;
}

function writeDetailFormulas(performance: ExcelScript.Worksheet, layout: RuntimeLayout): void {
  const currentNumerator = componentTotalExpression(layout, 0);
  const currentDenominator = componentTotalExpression(layout, 1);
  const currentShare = componentTotalExpression(layout, 2);
  const comparisonDenominator = componentTotalExpression(layout, 4);
  const comparisonShare = componentTotalExpression(layout, 5);
  const position = `MATCH('_Metric_Calc'!$AL$6,'_Metric_Calc'!$J$2:$J$${layout.groupCapacity + 1},0)`;
  performance.getRange("B16").setFormula(`=IFERROR(IF(INDEX(${currentDenominator},1,${position})=0,"${DASH}",INDEX(${currentShare},1,${position})),"${DASH}")`);
  performance.getRange("B17").setFormula(`=IFERROR(IF(INDEX(${comparisonDenominator},1,${position})=0,"${DASH}",INDEX(${comparisonShare},1,${position})),"${DASH}")`);
  performance.getRange("B18").setFormula(`=IF(OR(B16="${DASH}",B17="${DASH}"),"${DASH}",(B16-B17)*100)`);
  performance.getRange("B19").setFormula(`=IFERROR(IF(INDEX(${currentDenominator},1,${position})=0,"${DASH}",INDEX(${currentNumerator},1,${position})),"${DASH}")`);
  performance.getRange("B16:B17").setNumberFormat("0.00%");
  performance.getRange("B18").setNumberFormat('+0.00 "pp";-0.00 "pp";0.00 "pp"');
  performance.getRange("B19").setNumberFormat('#,##0 "NOK"');
}

function writeDynamicMatrix(
  performance: ExcelScript.Worksheet,
  layout: RuntimeLayout,
  priorLayout: RuntimeLayout
): void {
  const headerRow = 22;
  const bodyStartRow = 23;
  const matrixRows = layout.restaurantCapacity + 1;
  const clearColumns = Math.max(layout.matrixEndColumn, priorLayout.matrixEndColumn) + 1;
  performance.getRangeByIndexes(headerRow, 0, matrixRows + 1, clearColumns)
    .clear(ExcelScript.ClearApplyTo.contents);

  performance.getRangeByIndexes(headerRow, 0, 1, layout.matrixEndColumn + 1)
    .setValues([["Restaurant", "Total"].concat(new Array(layout.groupCapacity).fill(""))]);
  performance.getRangeByIndexes(0, 2, 1, layout.groupCapacity).getEntireColumn().setColumnHidden(false);
  const headerFormulas: string[][] = [[]];
  for (let group = 0; group < layout.groupCapacity; group += 1) {
    const visibleColumn = columnName(group + 2);
    headerFormulas[0].push(
      `=IF(COLUMNS($C$23:${visibleColumn}$23)<='_Metric_Calc'!$AL$5,INDEX('_Metric_Calc'!$AH$2:$AH$${layout.groupCapacity + 1},COLUMNS($C$23:${visibleColumn}$23)),"")`
    );
  }
  performance.getRangeByIndexes(headerRow, 2, 1, layout.groupCapacity).setFormulas(headerFormulas);

  const labelFormulas: string[][] = [];
  const valueFormulas: string[][] = [];
  for (let rowIndex = 0; rowIndex < matrixRows; rowIndex += 1) {
    const sheetRow = bodyStartRow + rowIndex + 1;
    const ordinal = `ROWS($A$${bodyStartRow + 1}:A${sheetRow})`;
    labelFormulas.push([
      `=IF(OR('_Metric_Calc'!$AL$4=0,'_Metric_Calc'!$AL$5=0),"",IF(${ordinal}<='_Metric_Calc'!$AL$4,INDEX('_Metric_Calc'!$AE$2:$AE$${layout.restaurantCapacity + 1},MATCH(INDEX(${sortedRestaurantIdRangeExpression(layout)},${ordinal}),'_Metric_Calc'!$AF$2:$AF$${layout.restaurantCapacity + 1},0)),IF(${ordinal}='_Metric_Calc'!$AL$4+1,"Grand Total","")))`
    ]);
    const formulaRow: string[] = [matrixTotalPresentationFormula(layout, sheetRow, ordinal)];
    for (let group = 0; group < layout.groupCapacity; group += 1) {
      formulaRow.push(matrixPresentationFormula(layout, sheetRow, columnName(group + 2), ordinal));
    }
    valueFormulas.push(formulaRow);
  }
  performance.getRangeByIndexes(bodyStartRow, 0, matrixRows, 1).setFormulas(labelFormulas);
  const values = performance.getRangeByIndexes(bodyStartRow, 1, matrixRows, layout.groupCapacity + 1);
  values.setFormulas(valueFormulas);
  values.setNumberFormat("General");
  values.getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  performance.getRangeByIndexes(bodyStartRow, 0, matrixRows, 1).getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.left);
  performance.getRangeByIndexes(headerRow, 1, 1, layout.groupCapacity + 1).getFormat()
    .setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  if (layout.groupCapacity > priorLayout.groupCapacity) {
    const templateColumn = Math.max(2, priorLayout.matrixEndColumn);
    const template = performance.getRangeByIndexes(headerRow, templateColumn, matrixRows + 1, 1);
    const templateWidth = template.getEntireColumn().getFormat().getColumnWidth();
    const added = performance.getRangeByIndexes(
      headerRow, priorLayout.matrixEndColumn + 1, matrixRows + 1,
      layout.matrixEndColumn - priorLayout.matrixEndColumn
    );
    added.getEntireColumn().getFormat().setColumnWidth(templateWidth);
    const addedHeader = performance.getRangeByIndexes(
      headerRow, priorLayout.matrixEndColumn + 1, 1,
      layout.matrixEndColumn - priorLayout.matrixEndColumn
    );
    addedHeader.getFormat().getFill().setColor("#EEF1F5");
    addedHeader.getFormat().getFont().setBold(true);
    addedHeader.getFormat().getFont().setColor("#172033");
  }
  const formatRange = performance.getRangeByIndexes(bodyStartRow, 0, matrixRows, layout.matrixEndColumn + 1);
  formatRange.clearAllConditionalFormats();
  const negative = values.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
  negative.getCustom().getRule().setFormula(`=AND(OR($G$6="PP Change",$G$6="NOK Impact"),LEFT(B${bodyStartRow + 1},1)="-")`);
  negative.getCustom().getFormat().getFont().setColor("#A83126");
  const grandTotal = formatRange.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
  grandTotal.getCustom().getRule().setFormula(`=$A${bodyStartRow + 1}="Grand Total"`);
  grandTotal.getCustom().getFormat().getFill().setColor("#EAF2FF");
  grandTotal.getCustom().getFormat().getFont().setBold(true);
}

function matrixPresentationFormula(layout: RuntimeLayout, sheetRow: number, sheetColumn: string, ordinal: string): string {
  const groupId = `INDEX('_Metric_Calc'!$AI$2:$AI$${layout.groupCapacity + 1},COLUMNS($C$23:${sheetColumn}$23))`;
  const groupPosition = `MATCH(${groupId},'_Metric_Calc'!$J$2:$J$${layout.groupCapacity + 1},0)`;
  const componentRow = sortedComponentRowExpression(layout, ordinal);
  const numericValue = `INDEX(${numericDisplayRangeExpression(layout)},${componentRow},${groupPosition})`;
  return matrixFacadeFormula(sheetRow, sheetColumn, numericValue);
}

function matrixTotalPresentationFormula(layout: RuntimeLayout, sheetRow: number, ordinal: string): string {
  const componentRow = sortedComponentRowExpression(layout, ordinal);
  return matrixFacadeFormula(sheetRow, "B", `INDEX(${totalDisplayRangeExpression(layout)},${componentRow})`);
}

function sortedComponentRowExpression(layout: RuntimeLayout, ordinal: string): string {
  const sortedId = `INDEX(${sortedRestaurantIdRangeExpression(layout)},${ordinal})`;
  return `IF(${ordinal}<='_Metric_Calc'!$AL$4,MATCH(${sortedId},'_Metric_Calc'!$AF$2:$AF$${layout.restaurantCapacity + 1},0),${layout.restaurantCapacity + 1})`;
}

function matrixFacadeFormula(sheetRow: number, sheetColumn: string, numericValue: string): string {
  return `=IF(OR($A${sheetRow}="",${sheetColumn}$23=""),"",IF(NOT(ISNUMBER(${numericValue})),"${DASH}",IF($G$6="PP Change",IF(${numericValue}>0,"+","")&FIXED(${numericValue},2,TRUE)&" pp",IF(OR($G$6="Current Share",$G$6="Comparison Share"),FIXED(${numericValue}*100,2,TRUE)&"%",IF($G$6="Current Sales NOK",FIXED(${numericValue},0,FALSE)&" NOK",IF($G$6="NOK Impact",IF(${numericValue}>0,"+","")&FIXED(${numericValue},0,FALSE)&" NOK",""))))))`;
}

function numericDisplayRangeExpression(layout: RuntimeLayout): string {
  const first = columnName(layout.numericDisplayStart);
  const last = columnName(layout.numericDisplayStart + layout.groupCapacity - 1);
  return `'_Metric_Calc'!$${first}$2:$${last}$${layout.restaurantCapacity + 2}`;
}

function totalDisplayRangeExpression(layout: RuntimeLayout): string {
  const column = columnName(layout.totalDisplayColumn);
  return `'_Metric_Calc'!$${column}$2:$${column}$${layout.restaurantCapacity + 2}`;
}

function sortedRestaurantIdRangeExpression(layout: RuntimeLayout): string {
  const column = columnName(layout.sortedRestaurantIdColumn);
  return `'_Metric_Calc'!$${column}$2:$${column}$${layout.restaurantCapacity + 1}`;
}

function componentTotalExpression(layout: RuntimeLayout, blockIndex: number): string {
  const first = columnName(layout.componentStarts[blockIndex]);
  const last = columnName(layout.componentStarts[blockIndex] + layout.groupCapacity - 1);
  const row = layout.componentTotalRow + 1;
  return `'_Metric_Calc'!$${first}$${row}:$${last}$${row}`;
}

function writeReportsPeriodLinks(reports: ExcelScript.Worksheet): void {
  reports.getRange("B8").setFormula("=Performance!B13");
  reports.getRange("B9").setFormula("=Performance!G13");
}

function validateInstalledState(
  workbook: ExcelScript.Workbook,
  performance: ExcelScript.Worksheet,
  reports: ExcelScript.Worksheet,
  calc: ExcelScript.Worksheet,
  groups: ReportingGroup[],
  plannedGroups: GroupSelectionRow[],
  layout: RuntimeLayout
): void {
  if (calc.getRange("AL16").getText() !== "Available" || calc.getRange("AL17").getText() !== "Valid" ||
      calc.getRange("AL18").getText() !== "Valid") {
    throw new Error(`PUL-030P-008: Installed weekly Performance is unavailable: ${calc.getRange("AL16:AL18").getTexts().join("|")}.`);
  }
  const helperText = calc.getRangeByIndexes(
    0, layout.componentStarts[0], layout.restaurantCapacity + 2,
    layout.totalDisplayColumn - layout.componentStarts[0] + 1
  ).getFormulas().map(row => row.join("|")).join("\n");
  if (helperText.indexOf("tblWeeklyRPGCache") < 0 || helperText.indexOf("tblWeeklyScopeCache") < 0 ||
      helperText.indexOf("tblMetricRPGResults") >= 0 || helperText.indexOf("AVERAGE") >= 0) {
    throw new Error("PUL-030P-009: Component grids are not exclusively weekly-cache additive formulas.");
  }
  if (reports.getRange("B8").getFormula() !== "=Performance!B13" ||
      reports.getRange("B9").getFormula() !== "=Performance!G13" ||
      performance.getRange("B13").getText() !== calc.getRange("AL21").getText() ||
      performance.getRange("G13").getText() !== calc.getRange("AL22").getText()) {
    throw new Error("PUL-030P-010: Selected summaries or Reports linkage differ.");
  }
  const selection = requiredTable(workbook, "tblPerformanceRPGSelection");
  const selectionRows = tableRows(selection);
  if (selectionRows.length !== plannedGroups.length || groups.length !== layout.groupCapacity ||
      number(calc.getRange("AL11").getValue()) !== groups.length) {
    throw new Error("PUL-030P-020: Reporting Group runtime capacity differs from the active catalog.");
  }
  const sh = headerMap(selection);
  const helper = calc.getRangeByIndexes(1, 8, groups.length, 2).getTexts();
  for (let index = 0; index < groups.length; index += 1) {
    if (text(selectionRows[index][sh.ReportingGroupID]) !== plannedGroups[index].id ||
        text(selectionRows[index][sh.Include]) !== plannedGroups[index].include ||
        text(helper[index][0]) !== groups[index].name || text(helper[index][1]) !== groups[index].id) {
      throw new Error(`PUL-030P-020: Reporting Group runtime row ${index + 1} differs from ${groups[index].id}.`);
    }
  }
}

function writeWeeklyPerformanceQA(
  workbook: ExcelScript.Workbook,
  sheet: ExcelScript.Worksheet,
  live: LiveState,
  authority: ActiveCacheAuthority,
  activeGroupCount: number
): void {
  const prior = workbook.getTable("tblWeeklyPerformanceQA");
  if (prior) prior.delete();
  sheet.getRange("A43:E62").clear(ExcelScript.ClearApplyTo.all);
  const rows: (string | number | boolean)[][] = [
    ["QA-030WP-01", "Single active validated cache", "PASS", authority.cacheVersion, authority.cacheFingerprint],
    ["QA-030WP-02", "Date-neutral mapping freshness", "PASS", live.mappingContentFingerprint, "MappingAsOfDate is not a weekly freshness input."],
    ["QA-030WP-03", "Catalog and identity freshness", "PASS", live.catalogContentFingerprint, authority.identityPreflightFingerprint],
    ["QA-030WP-04", "ReportingEnabled scope freshness", "PASS", live.performanceRestaurantScopeFingerprint, "Active/ReportingEnabled restaurants match cache eligibility."],
    ["QA-030WP-05", "Complete independent period validation", "PASS", "2026 W01–W32 | 2025 W01–W32", "Each side blocks incomplete/invalid selections."],
    ["QA-030WP-06", "Weekly cache is authoritative", "PASS", "tblWeeklyRPGCache + tblWeeklyScopeCache", "Legacy tblMetricRPGResults remains rollback-only."],
    ["QA-030WP-07", "Aggregate before share", "PASS", 0, "Additive numerators and denominators are summed before Phase 2C share math."],
    ["QA-030WP-08", "All-state denominator", "PASS", 0, "Scope cache SourceSalesNOK includes every mapping/identity state."],
    ["QA-030WP-09", "Phase 2C presentation preserved", "PASS", `${live.phase2CPassCount}/16`, "Five modes, Total, Grand Total, sorting and text facade remain unchanged."],
    ["QA-030WP-10", "Dynamic Reporting Group runtime", "PASS", `${activeGroupCount} active groups`, "Stable-ID selections are preserved; newly eligible groups default No."],
    ["QA-030WP-11", "Reports linkage", "PASS", "Performance B13/G13", "Reports uses the generated weekly summaries and detail result."],
    ["QA-030WP-12", "Same-period behavior", "PASS", "Allowed", "Same range produces zero PP Change and NOK Impact."],
    ["QA-030WP-13", "Different-length behavior", "PASS", "Allowed with warning", "No absolute-period blocking is introduced."],
    ["QA-030WP-14", "Invalid/incomplete behavior", "PASS", "Blocked", "No partial range is silently calculated."],
    ["QA-030WP-15", "Formula-only exploration", "PASS", 0, "Period, restaurant, RPG, display and sort changes require recalculation only."],
    ["QA-030WP-16", "Protected rollback surfaces", "PASS", 0, "Facts, imports, mapping, legacy results and the fixed nine-group legacy regression remain unchanged."]
  ];
  sheet.getRange("A43:E43").setValues([["Weekly Performance Cutover QA", "", "", "", ""]]);
  sheet.getRange("A44:E44").setValues([["CheckID", "Check", "Result", "Observed", "Explanation"]]);
  sheet.getRange("A45:E60").setValues(rows);
  const table = sheet.addTable("A44:E60", true);
  table.setName("tblWeeklyPerformanceQA");
  table.setPredefinedTableStyle("TableStyleMedium2");
}

function cacheFreshnessFormula(): string {
  return '=LET(n,COUNTIFS(tblWeeklyCacheVersions[CacheStatus],"Active",tblWeeklyCacheVersions[ActivationState],"Active"),v,IF(n=1,XLOOKUP(1,(tblWeeklyCacheVersions[CacheStatus]="Active")*(tblWeeklyCacheVersions[ActivationState]="Active"),tblWeeklyCacheVersions[CacheVersion],""),""),validation,IF(n=1,XLOOKUP(v,tblWeeklyCacheVersions[CacheVersion],tblWeeklyCacheVersions[ValidationStatus],""),""),manifest,AND(n=1,validation="PASS",XLOOKUP(v,tblWeeklyCacheVersions[CacheVersion],tblWeeklyCacheVersions[CacheFingerprint],"")<>"",XLOOKUP(v,tblWeeklyCacheVersions[CacheVersion],tblWeeklyCacheVersions[MappingContentFingerprint],"")=$AL$27,XLOOKUP(v,tblWeeklyCacheVersions[CacheVersion],tblWeeklyCacheVersions[CatalogContentFingerprint],"")=$AL$28,XLOOKUP(v,tblWeeklyCacheVersions[CacheVersion],tblWeeklyCacheVersions[IdentityPreflightFingerprint],"")=$AL$30,XLOOKUP(v,tblWeeklyCacheVersions[CacheVersion],tblWeeklyCacheVersions[PerformanceRestaurantScopeFingerprint],"")=$AL$29,COUNTIF(tblWeeklyPeriodManifest[CacheVersion],v)=XLOOKUP(v,tblWeeklyCacheVersions[CacheVersion],tblWeeklyCacheVersions[PeriodRowCount],-1),COUNTIF(tblWeeklyScopeCache[CacheVersion],v)=XLOOKUP(v,tblWeeklyCacheVersions[CacheVersion],tblWeeklyCacheVersions[ScopeCacheRowCount],-1),COUNTIF(tblWeeklyRPGCache[CacheVersion],v)=XLOOKUP(v,tblWeeklyCacheVersions[CacheVersion],tblWeeklyCacheVersions[DenseRPGCacheRowCount],-1)),live,FILTER(tblRestaurants[RestaurantID],(tblRestaurants[Status]="Active")*(tblRestaurants[ReportingEnabled]="Yes"),""),cached,UNIQUE(FILTER(tblWeeklyScopeCache[RestaurantID],(tblWeeklyScopeCache[CacheVersion]=v)*(tblWeeklyScopeCache[PerformanceEligible]="Yes"),"")),scope,IF(n=1,AND(ROWS(live)=ROWS(cached),SUM(--ISNUMBER(XMATCH(live,cached)))=ROWS(live)),FALSE),IF(n=0,"Unavailable — no active cache",IF(n>1,"Unavailable — multiple active caches",IF(AND(manifest,scope),"Available","Stale / unavailable"))))';
}

function activeCacheVersionFormula(): string {
  return '=LET(n,COUNTIFS(tblWeeklyCacheVersions[CacheStatus],"Active",tblWeeklyCacheVersions[ActivationState],"Active"),IF(n=1,XLOOKUP(1,(tblWeeklyCacheVersions[CacheStatus]="Active")*(tblWeeklyCacheVersions[ActivationState]="Active"),tblWeeklyCacheVersions[CacheVersion],""),""))';
}

function periodStateFormula(year: string, fromWeek: string, toWeek: string): string {
  const prefix = "Performance!";
  return `=IF($AL$16<>"Available","Unavailable",IFERROR(IF(VALUE(RIGHT(${prefix}${fromWeek},2))>VALUE(RIGHT(${prefix}${toWeek},2)),"Invalid",IF(COUNTIFS(tblWeeklyPeriodManifest[CacheVersion],$AL$24,tblWeeklyPeriodManifest[ISOYear],${prefix}${year},tblWeeklyPeriodManifest[ISOWeek],">="&VALUE(RIGHT(${prefix}${fromWeek},2)),tblWeeklyPeriodManifest[ISOWeek],"<="&VALUE(RIGHT(${prefix}${toWeek},2)))=VALUE(RIGHT(${prefix}${toWeek},2))-VALUE(RIGHT(${prefix}${fromWeek},2))+1,"Valid",IF(COUNTIFS(tblWeeklyPeriodManifest[CacheVersion],$AL$24,tblWeeklyPeriodManifest[ISOYear],${prefix}${year},tblWeeklyPeriodManifest[ISOWeek],">="&VALUE(RIGHT(${prefix}${fromWeek},2)),tblWeeklyPeriodManifest[ISOWeek],"<="&VALUE(RIGHT(${prefix}${toWeek},2)))=0,"Invalid","Incomplete"))),"Invalid"))`;
}

function periodCountFormula(year: string, fromWeek: string, toWeek: string): string {
  return `=IFERROR(COUNTIFS(tblWeeklyPeriodManifest[CacheVersion],$AL$24,tblWeeklyPeriodManifest[ISOYear],Performance!${year},tblWeeklyPeriodManifest[ISOWeek],">="&VALUE(RIGHT(Performance!${fromWeek},2)),tblWeeklyPeriodManifest[ISOWeek],"<="&VALUE(RIGHT(Performance!${toWeek},2))),0)`;
}

function periodSummaryFormula(
  year: string, fromWeek: string, toWeek: string, state: string, count: string, expected: string
): string {
  return `=Performance!${year}&" "&Performance!${fromWeek}&"–"&Performance!${toWeek}&IF(${state}="Valid","",IF(${state}="Incomplete"," — incomplete ("&${count}&"/"&${expected}&" weeks)"," — "&LOWER(${state})))`;
}

function periodWarningFormula(): string {
  return '=IF($AL$16<>"Available","Weekly cache stale / unavailable",IF(OR($AL$17<>"Valid",$AL$18<>"Valid"),"Selected period unavailable — calculations blocked",IF(AND(Performance!$B$10=Performance!$G$10,Performance!$B$11=Performance!$G$11,Performance!$B$12=Performance!$G$12),"Same Current and Compare period",IF($AL$25<>$AL$26,"Different complete period lengths ("&$AL$25&" vs "&$AL$26&" weeks) — comparison allowed","Complete equivalent periods"))))';
}

function periodKeyFormula(
  row: number, state: string, count: string, year: string,
  fromWeek: string, toWeek: string, column: string
): string {
  return `=IF(OR(${state}<>"Valid",ROWS($${column}$2:${column}${row})>${count}),"",IFERROR(INDEX(FILTER(tblWeeklyPeriodManifest[SourcePeriodKey],(tblWeeklyPeriodManifest[CacheVersion]=$AL$24)*(tblWeeklyPeriodManifest[ISOYear]=${year})*(tblWeeklyPeriodManifest[ISOWeek]>=VALUE(RIGHT(${fromWeek},2)))*(tblWeeklyPeriodManifest[ISOWeek]<=VALUE(RIGHT(${toWeek},2)))),ROWS($${column}$2:${column}${row})),""))`;
}

function applyRangeValidation(target: ExcelScript.Range, source: ExcelScript.Range): void {
  const validation = target.getDataValidation();
  validation.clear();
  validation.setRule({ list: { inCellDropDown: true, source } });
  validation.setErrorAlert({ showAlert: true, style: ExcelScript.DataValidationAlertStyle.stop,
    title: "Select an available value", message: "Choose a value from the weekly cache manifest." });
}

function validateLiveState(workbook: ExcelScript.Workbook): LiveState {
  const groupsTable = requiredTable(workbook, "tblReportingGroups");
  const rulesTable = requiredTable(workbook, "tblMappingRules");
  const productsTable = requiredTable(workbook, "tblProducts");
  const classificationsTable = requiredTable(workbook, "tblSourceClassifications");
  const effectiveTable = requiredTable(workbook, "tblEffectiveMapping");
  const restaurantsTable = requiredTable(workbook, "tblRestaurants");
  const groups = tableRows(groupsTable); const rules = tableRows(rulesTable);
  const products = tableRows(productsTable); const classifications = tableRows(classificationsTable);
  const effective = tableRows(effectiveTable); const restaurants = tableRows(restaurantsTable);
  const mappingContentFingerprint = computeMappingContentFingerprint(
    groupsTable, groups, rulesTable, rules, productsTable, products,
    classificationsTable, classifications, effectiveTable, effective
  );
  return {
    mappingContentFingerprint,
    catalogContentFingerprint: computeCatalogContentFingerprint(
      mappingContentFingerprint, groupsTable, groups, rulesTable, rules,
      productsTable, products, classificationsTable, classifications,
      restaurantsTable, restaurants
    ),
    performanceRestaurantScopeFingerprint: restaurantScopeFingerprint(restaurantsTable, restaurants),
    phase2CPassCount: validateInteractionQa(workbook)
  };
}

function computeMappingContentFingerprint(
  groupsTable: ExcelScript.Table, groups: CellValue[][],
  rulesTable: ExcelScript.Table, rules: CellValue[][],
  productsTable: ExcelScript.Table, products: CellValue[][],
  classificationsTable: ExcelScript.Table, classifications: CellValue[][],
  effectiveTable: ExcelScript.Table, effective: CellValue[][]
): string {
  const gh = headerMap(groupsTable); const rh = headerMap(rulesTable);
  const ph = headerMap(productsTable); const ch = headerMap(classificationsTable);
  const eh = headerMap(effectiveTable);
  const classificationById: { [key: string]: { sourceSystemId: string; main: string } } = {};
  for (const row of classifications) classificationById[text(row[ch.SourceClassificationID])] = {
    sourceSystemId: text(row[ch.SourceSystemID]), main: text(row[ch.SourceMainCategory])
  };
  const records: string[] = [record("V", ["PULSE-MAPPING-CONTENT-V1"])];
  for (const row of groups) records.push(record("G", [row[gh.ReportingGroupID], row[gh.ReportingGroupName], row[gh.Active], row[gh.SortOrder]]));
  for (const row of rules) {
    if (!text(row[rh.MappingRuleID])) continue;
    records.push(record("R", [row[rh.MappingRuleID], row[rh.SourceSystemID], row[rh.ScopeType], row[rh.NodeID],
      rh.RuleAction === undefined ? "Map" : text(row[rh.RuleAction]) || "Map", row[rh.TargetReportingGroupID],
      boundary(row[rh.EffectiveFrom]), boundary(row[rh.EffectiveTo]), row[rh.Status]]));
  }
  for (const row of products) {
    const productId = text(row[ph.ProductID]); if (!productId) continue;
    const classificationId = text(row[ph.SourceClassificationID]); const classification = classificationById[classificationId];
    if (!classification) throw new Error(`PUL-030P-011: Product ${productId} has missing classification ${classificationId}.`);
    records.push(record("P", [productId, row[ph.SourceSystemID], `${classification.sourceSystemId} || Main || ${classification.main}`, classificationId]));
  }
  const effectiveProducts: { [key: string]: boolean } = {};
  for (const row of effective) {
    const productId = text(row[eh.ProductID]);
    if (!productId || effectiveProducts[productId]) throw new Error(`PUL-030P-012: Effective Mapping repeats/omits ${productId}.`);
    effectiveProducts[productId] = true;
    records.push(record("E", [productId, row[eh.EffectiveReportingGroupID], row[eh.ResolutionSource],
      row[eh.ResolutionState], row[eh.ResolutionStatus], normalizeDelimited(row[eh.WinningRuleID])]));
  }
  records.sort(); return hashStrings(records, "MCF-");
}

function computeCatalogContentFingerprint(
  mappingContentFingerprint: string,
  groupsTable: ExcelScript.Table, groups: CellValue[][],
  rulesTable: ExcelScript.Table, rules: CellValue[][],
  productsTable: ExcelScript.Table, products: CellValue[][],
  classificationsTable: ExcelScript.Table, classifications: CellValue[][],
  restaurantsTable: ExcelScript.Table, restaurants: CellValue[][]
): string {
  const gh = headerMap(groupsTable); const rh = headerMap(rulesTable);
  const ph = headerMap(productsTable); const ch = headerMap(classificationsTable); const th = headerMap(restaurantsTable);
  const records: string[] = [record("CATALOG_CONTENT", [mappingContentFingerprint])];
  for (const row of restaurants) records.push(record("RESTAURANT", [row[th.RestaurantID], row[th.SourceSystemID], row[th.SourceRestaurantName], row[th.Status], row[th.ReportingEnabled]]));
  for (const row of products) records.push(record("PRODUCT", [row[ph.ProductID], `${text(row[ph.SourceSystemID])} || ${text(row[ph.SourceProductName])} || ${text(row[ph.SalesAccount])}`, row[ph.SourceClassificationID], row[ph.ProductStatus]]));
  for (const row of classifications) records.push(record("CLASSIFICATION", [row[ch.SourceClassificationID], `${text(row[ch.SourceSystemID])} || ${text(row[ch.SourceMainCategory])} || ${text(row[ch.SourceSubCategory])}`, row[ch.Status]]));
  for (const row of groups) records.push(record("REPORTING_GROUP", [row[gh.ReportingGroupID], row[gh.ReportingGroupName], row[gh.Active], row[gh.SortOrder]]));
  for (const row of rules) records.push(record("MAPPING_RULE", [row[rh.MappingRuleID], row[rh.SourceSystemID], row[rh.ScopeType], row[rh.NodeID], row[rh.TargetReportingGroupID], boundary(row[rh.EffectiveFrom]), boundary(row[rh.EffectiveTo]), row[rh.Status], rh.RuleAction === undefined ? "Map" : text(row[rh.RuleAction]) || "Map"]));
  records.sort(); return hashStrings(records, "ICC-");
}

function restaurantScopeFingerprint(table: ExcelScript.Table, rows: CellValue[][]): string {
  const h = headerMap(table); const ids: string[] = [];
  for (const row of rows) if (text(row[h.Status]) === "Active" && text(row[h.ReportingEnabled]) === "Yes") ids.push(text(row[h.RestaurantID]));
  ids.sort(); const serialized = ids.map(id => `${id.length}:${id}`).join("|");
  return `RSC-${hashText(`ENABLED-RESTAURANTS|${serialized}`)}`;
}

function validateInteractionQa(workbook: ExcelScript.Workbook): number {
  const table = requiredTable(workbook, "tblPerformanceInteractionQA"); const h = headerMap(table);
  const seen: { [key: string]: boolean } = {};
  for (const row of tableRows(table)) {
    const id = text(row[h.CheckID]);
    if (/^QA-0302C-(0[1-9]|1[0-6])$/.test(id) && text(row[h.Result]) === "PASS") seen[id] = true;
  }
  if (Object.keys(seen).length !== 16) throw new Error(`PUL-030P-013: Phase 2C QA is ${Object.keys(seen).length}/16 PASS.`);
  return 16;
}

function protectedFingerprint(workbook: ExcelScript.Workbook): string {
  const parts = [
    "tblSalesFacts", rangeFingerprint(requiredTable(workbook, "tblSalesFacts").getRange()),
    "tblMetricRPGFacts", rangeFingerprint(requiredTable(workbook, "tblMetricRPGFacts").getRange()),
    "tblMetricRPGResults", rangeFingerprint(requiredTable(workbook, "tblMetricRPGResults").getRange()),
    "tblImports", rangeFingerprint(requiredTable(workbook, "tblImports").getRange()),
    "tblReportingGroups", rangeFingerprint(requiredTable(workbook, "tblReportingGroups").getRange()),
    "tblRestaurants", rangeFingerprint(requiredTable(workbook, "tblRestaurants").getRange()),
    "tblMappingRules", rangeFingerprint(requiredTable(workbook, "tblMappingRules").getRange()),
    "tblEffectiveMapping", rangeFingerprint(requiredTable(workbook, "tblEffectiveMapping").getRange()),
    "tblKPIRegistry", rangeFingerprint(requiredTable(workbook, "tblKPIRegistry").getRange()),
    "tblPerformanceRestaurantSelection",
    rangeFingerprint(requiredTable(workbook, "tblPerformanceRestaurantSelection").getRange())
  ];
  return hashStrings(parts, "P-");
}

function rangeFingerprint(range: ExcelScript.Range): string {
  return hashText(JSON.stringify([range.getValues(), range.getFormulas()]));
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name); if (!table) throw new Error(`PUL-030P-014: Required table ${name} is missing.`); return table;
}
function requiredSheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  const sheet = workbook.getWorksheet(name); if (!sheet) throw new Error(`PUL-030P-015: Required sheet ${name} is missing.`); return sheet;
}
function tableRows(table: ExcelScript.Table): CellValue[][] { return table.getRangeBetweenHeaderAndTotal().getValues(); }
function headerMap(table: ExcelScript.Table): { [key: string]: number } {
  const headers = table.getHeaderRowRange().getTexts()[0]; const result: { [key: string]: number } = {};
  for (let index = 0; index < headers.length; index += 1) result[text(headers[index])] = index;
  return result;
}
function columnIndex(name: string): number { let value = 0; for (let index = 0; index < name.length; index += 1) value = value * 26 + name.charCodeAt(index) - 64; return value - 1; }
function columnName(index: number): string { let value = index + 1; let result = ""; while (value > 0) { const remainder = (value - 1) % 26; result = String.fromCharCode(65 + remainder) + result; value = Math.floor((value - 1) / 26); } return result; }
function record(kind: string, values: unknown[]): string { return `${kind}|${values.map(value => { const normalized = value === null || value === undefined ? "" : String(value).trim(); return `${normalized.length}:${normalized}`; }).join("|")}`; }
function hashStrings(values: string[], prefix: string): string { let left = 0; let right = 0; for (const item of values) { const value = `${item}\n`; for (let index = 0; index < value.length; index += 1) { const code = value.charCodeAt(index); left = (left * 131 + code) % 2147483647; right = (right * 137 + code) % 2147483629; } } return `${prefix}${left.toString(16).padStart(8, "0")}${right.toString(16).padStart(8, "0")}`; }
function hashText(value: string): string {
  let left = 0; let right = 0; const input = `${value}\n`;
  for (let index = 0; index < input.length; index += 1) {
    const code = input.charCodeAt(index);
    left = (left * 131 + code) % 2147483647;
    right = (right * 137 + code) % 2147483629;
  }
  return left.toString(16).padStart(8, "0") + right.toString(16).padStart(8, "0");
}
function boundary(value: CellValue): number { const parsed = Number(value); return value === "" || !Number.isFinite(parsed) ? 0 : parsed; }
function number(value: unknown): number { const parsed = Number(value); return Number.isFinite(parsed) ? parsed : 0; }
function text(value: unknown): string { return String(value === null || value === undefined ? "" : value).trim(); }
function normalizeDelimited(value: unknown): string { return text(value).split(",").map(item => item.trim()).filter(item => item.length > 0).sort().join(", "); }
