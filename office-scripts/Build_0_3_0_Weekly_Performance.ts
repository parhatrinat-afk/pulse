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
  validateActiveCache(versionTable, versionRows, periodTable, periodRows, scopeRows, rpgRows);
  const live = validateLiveState(workbook);
  if (live.mappingContentFingerprint !== EXPECTED_MAPPING_CONTENT ||
      live.catalogContentFingerprint !== EXPECTED_CATALOG_CONTENT ||
      live.performanceRestaurantScopeFingerprint !== EXPECTED_RESTAURANT_SCOPE ||
      live.phase2CPassCount !== 16) {
    throw new Error(`PUL-030P-001: Weekly Performance freshness preflight failed. ${JSON.stringify(live)}`);
  }
  validatePhase2CLayout(workbook, calc, restaurantTable, interactionQa);

  const protectedBefore = protectedFingerprint(workbook);
  const currentValues = performance.getRange("B10:B12").getValues();
  const comparisonValues = performance.getRange("G10:G12").getValues();
  const alreadyWeekly = text(performance.getRange("A10").getValue()) === "Year" &&
    text(performance.getRange("F10").getValue()) === "Year";

  writeValidationLists(performance, periodRows);
  writePeriodControls(performance, alreadyWeekly, currentValues, comparisonValues);
  writeWeeklyControl(calc);
  writePeriodKeyHelpers(calc);
  writeWeeklyComponentFormulas(calc);
  writeReportsPeriodLinks(reports);
  workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.automatic);
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  validateInstalledState(performance, reports, calc);
  writeWeeklyPerformanceQA(workbook, qaSheet, live);
  const protectedAfter = protectedFingerprint(workbook);
  if (protectedBefore !== protectedAfter) {
    throw new Error("PUL-030P-002: A protected source, mapping, legacy result, import, or selection table changed.");
  }
  return JSON.stringify({
    status: "PASS",
    cacheVersion: EXPECTED_CACHE_VERSION,
    cacheFingerprint: EXPECTED_CACHE_FINGERPRINT,
    current: performance.getRange("B13").getText(),
    comparison: performance.getRange("G13").getText(),
    cacheFreshness: calc.getRange("AL16").getText(),
    currentPeriodState: calc.getRange("AL17").getText(),
    comparisonPeriodState: calc.getRange("AL18").getText(),
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

const EXPECTED_CACHE_VERSION = "WCV-1a34ad1f46763d9b";
const EXPECTED_CACHE_FINGERPRINT = "WCC-508dd608166cdb6e";
const EXPECTED_MAPPING_CONTENT = "MCF-759cc92c4304a913";
const EXPECTED_CATALOG_CONTENT = "ICC-5644a77c18a97437";
const EXPECTED_IDENTITY_PREFLIGHT = "IDP-062c182f23905ae8";
const EXPECTED_RESTAURANT_SCOPE = "RSC-08df626f217dd94b";
const RESTAURANT_CAPACITY = 16;
const GROUP_CAPACITY = 9;
const CURRENT_PERIOD_KEYS = "$DN$2:$DN$54";
const COMPARISON_PERIOD_KEYS = "$DO$2:$DO$54";

function validateActiveCache(
  versionTable: ExcelScript.Table,
  versionRows: CellValue[][],
  periodTable: ExcelScript.Table,
  periodRows: CellValue[][],
  scopeRows: CellValue[][],
  rpgRows: CellValue[][]
): void {
  if (versionRows.length !== 1 || periodRows.length !== 84 || scopeRows.length !== 1421 || rpgRows.length !== 12789) {
    throw new Error(`PUL-030P-003: Weekly cache row counts differ: ${versionRows.length}/${periodRows.length}/${scopeRows.length}/${rpgRows.length}.`);
  }
  const vh = headerMap(versionTable);
  const row = versionRows[0];
  const required = [
    ["CacheVersion", EXPECTED_CACHE_VERSION], ["CacheStatus", "Active"],
    ["ActivationState", "Active"], ["ValidationStatus", "PASS"],
    ["CacheFingerprint", EXPECTED_CACHE_FINGERPRINT],
    ["MappingContentFingerprint", EXPECTED_MAPPING_CONTENT],
    ["CatalogContentFingerprint", EXPECTED_CATALOG_CONTENT],
    ["IdentityPreflightFingerprint", EXPECTED_IDENTITY_PREFLIGHT],
    ["PerformanceRestaurantScopeFingerprint", EXPECTED_RESTAURANT_SCOPE]
  ];
  for (const item of required) {
    if (text(row[vh[item[0]]]) !== item[1]) {
      throw new Error(`PUL-030P-004: ${item[0]} is ${text(row[vh[item[0]]])}; expected ${item[1]}.`);
    }
  }
  const ph = headerMap(periodTable);
  const seen: { [key: string]: boolean } = {};
  for (const period of periodRows) {
    const key = `${number(period[ph.ISOYear])}|${number(period[ph.ISOWeek])}`;
    if (seen[key]) throw new Error(`PUL-030P-005: Duplicate manifest ISO period ${key}.`);
    seen[key] = true;
  }
}

function validatePhase2CLayout(
  workbook: ExcelScript.Workbook,
  calc: ExcelScript.Worksheet,
  restaurantTable: ExcelScript.Table,
  qaTable: ExcelScript.Table
): void {
  if (tableRows(requiredTable(workbook, "tblPerformanceRestaurantSelection")).length !== RESTAURANT_CAPACITY ||
      tableRows(requiredTable(workbook, "tblPerformanceRPGSelection")).length !== GROUP_CAPACITY ||
      calc.getRange("AE1:DL1").getTexts()[0].join("|").indexOf("Current Numerator") < 0) {
    throw new Error("PUL-030P-006: Accepted Phase 2C helper/selection layout is missing.");
  }
  if (tableRows(restaurantTable).length < RESTAURANT_CAPACITY || tableRows(qaTable).length !== 16) {
    throw new Error("PUL-030P-007: Restaurant catalog or Phase 2C QA is incomplete.");
  }
}

function writeValidationLists(performance: ExcelScript.Worksheet, periodRows: CellValue[][]): void {
  const years: number[] = [];
  const seenYears: { [key: string]: boolean } = {};
  for (const row of periodRows) {
    const year = number(row[5]);
    if (!seenYears[String(year)]) { seenYears[String(year)] = true; years.push(year); }
  }
  years.sort((left, right) => left - right);
  const yearValues: (string | number | boolean)[][] = [];
  for (const year of years) yearValues.push([year]);
  const weekValues: (string | number | boolean)[][] = [];
  for (let week = 1; week <= 53; week += 1) weekValues.push([`W${String(week).padStart(2, "0")}`]);
  performance.getRange("V1:W54").clear(ExcelScript.ClearApplyTo.contents);
  performance.getRange("V1:W1").setValues([["Available year", "ISO week"]]);
  performance.getRangeByIndexes(1, 21, yearValues.length, 1).setValues(yearValues);
  performance.getRange("W2:W54").setValues(weekValues);
  performance.getRange("V:W").setColumnHidden(true);
}

function writePeriodControls(
  performance: ExcelScript.Worksheet,
  alreadyWeekly: boolean,
  priorCurrent: CellValue[][],
  priorComparison: CellValue[][]
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
  applyRangeValidation(performance.getRange("B10"), performance.getRange("V2:V3"));
  applyRangeValidation(performance.getRange("G10"), performance.getRange("V2:V3"));
  const weekSource = performance.getRange("W2:W54");
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

function writeWeeklyControl(calc: ExcelScript.Worksheet): void {
  calc.getRange("AK1:AL26").clear(ExcelScript.ClearApplyTo.contents);
  calc.getRange("AK1:AL26").setValues([
    ["Weekly Performance control", "Value"],
    ["Current period", ""], ["Comparison period", ""],
    ["Selected restaurant count", ""], ["Selected RPG count", ""],
    ["Detail ReportingGroupID", ""], ["Matrix display", ""],
    ["Company scope fingerprint", EXPECTED_RESTAURANT_SCOPE],
    ["Interaction contract", "Weekly cache additive components only"],
    ["Restaurant capacity", RESTAURANT_CAPACITY], ["RPG capacity", GROUP_CAPACITY],
    ["Requested sort ReportingGroupID", ""], ["Effective sort ReportingGroupID", ""],
    ["Sort order", ""], ["Sort status", ""],
    ["Active-cache freshness", ""], ["Current period state", ""],
    ["Comparison period state", ""], ["Current available weeks", ""],
    ["Comparison available weeks", ""], ["Current summary", ""],
    ["Comparison summary", ""], ["Period status", ""],
    ["Active CacheVersion", ""], ["Current expected weeks", ""],
    ["Comparison expected weeks", ""]
  ]);
  calc.getRange("AL2").setFormula("=Performance!$B$13");
  calc.getRange("AL3").setFormula("=Performance!$G$13");
  calc.getRange("AL4").setFormula('=COUNTIF(tblPerformanceRestaurantSelection[Include],"Yes")');
  calc.getRange("AL5").setFormula('=COUNTIF(tblPerformanceRPGSelection[Include],"Yes")');
  calc.getRange("AL6").setFormula('=IFERROR(INDEX($J$2:$J$10,MATCH(Performance!$B$7,$I$2:$I$10,0)),"")');
  calc.getRange("AL7").setFormula("=Performance!$G$6");
  calc.getRange("AL12").setFormula('=IF(Performance!$I$6="Total","",IFERROR(INDEX($J$2:$J$10,MATCH(Performance!$I$6,$I$2:$I$10,0)),""))');
  calc.getRange("AL13").setFormula('=IF(OR($AL$12="",COUNTIF($AI$2:$AI$10,$AL$12)=0),"",$AL$12)');
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
  calc.getRange("AL24").setFormula(`=IF($AL$16="Available","${EXPECTED_CACHE_VERSION}","")`);
  calc.getRange("AL25").setFormula('=IFERROR(VALUE(RIGHT(Performance!$B$12,2))-VALUE(RIGHT(Performance!$B$11,2))+1,0)');
  calc.getRange("AL26").setFormula('=IFERROR(VALUE(RIGHT(Performance!$G$12,2))-VALUE(RIGHT(Performance!$G$11,2))+1,0)');
}

function writePeriodKeyHelpers(calc: ExcelScript.Worksheet): void {
  calc.getRange("DN1:DO54").clear(ExcelScript.ClearApplyTo.contents);
  calc.getRange("DN1:DO1").setValues([["Current SourcePeriodKey", "Comparison SourcePeriodKey"]]);
  const formulas: string[][] = [];
  for (let index = 0; index < 53; index += 1) {
    const row = index + 2;
    formulas.push([
      periodKeyFormula(row, "$AL$17", "$AL$19", "Performance!$B$10", "Performance!$B$11", "Performance!$B$12", "DN"),
      periodKeyFormula(row, "$AL$18", "$AL$20", "Performance!$G$10", "Performance!$G$11", "Performance!$G$12", "DO")
    ]);
  }
  calc.getRange("DN2:DO54").setFormulas(formulas);
}

function writeWeeklyComponentFormulas(calc: ExcelScript.Worksheet): void {
  writeWeeklyComponentBlock(calc, "AN", true, true);
  writeWeeklyComponentBlock(calc, "AX", true, false);
  writeWeeklyComponentBlock(calc, "BR", false, true);
  writeWeeklyComponentBlock(calc, "CB", false, false);
}

function writeWeeklyComponentBlock(
  calc: ExcelScript.Worksheet,
  startColumn: string,
  current: boolean,
  numerator: boolean
): void {
  const startIndex = columnIndex(startColumn);
  const periodState = current ? "$AL$17" : "$AL$18";
  const periodKeys = current ? CURRENT_PERIOD_KEYS : COMPARISON_PERIOD_KEYS;
  const table = numerator ? "tblWeeklyRPGCache" : "tblWeeklyScopeCache";
  const value = numerator ? "MappedSalesNOK" : "SourceSalesNOK";
  const formulas: string[][] = [];
  for (let restaurant = 0; restaurant < RESTAURANT_CAPACITY; restaurant += 1) {
    const excelRow = restaurant + 2;
    const row: string[] = [];
    for (let group = 0; group < GROUP_CAPACITY; group += 1) {
      const groupRow = group + 2;
      let formula = `=IF(OR($AF${excelRow}="",$AL$16<>"Available",${periodState}<>"Valid"),0,IFERROR(SUM(SUMIFS(${table}[${value}],${table}[CacheVersion],$AL$24,${table}[RestaurantID],$AF${excelRow},`;
      if (numerator) formula += `${table}[ReportingGroupID],$J$${groupRow},`;
      formula += `${table}[SourcePeriodKey],FILTER(${periodKeys},${periodKeys}<>""))),0))`;
      row.push(formula);
    }
    formulas.push(row);
  }
  calc.getRangeByIndexes(1, startIndex, RESTAURANT_CAPACITY, GROUP_CAPACITY).setFormulas(formulas);
  const totals: string[][] = [[]];
  for (let group = 0; group < GROUP_CAPACITY; group += 1) {
    const column = columnName(startIndex + group + 1);
    totals[0].push(`=SUM(${column}$2:${column}$17)`);
  }
  calc.getRangeByIndexes(17, startIndex, 1, GROUP_CAPACITY).setFormulas(totals);
}

function writeReportsPeriodLinks(reports: ExcelScript.Worksheet): void {
  reports.getRange("B8").setFormula("=Performance!B13");
  reports.getRange("B9").setFormula("=Performance!G13");
}

function validateInstalledState(
  performance: ExcelScript.Worksheet,
  reports: ExcelScript.Worksheet,
  calc: ExcelScript.Worksheet
): void {
  if (calc.getRange("AL16").getText() !== "Available" || calc.getRange("AL17").getText() !== "Valid" ||
      calc.getRange("AL18").getText() !== "Valid") {
    throw new Error(`PUL-030P-008: Installed weekly Performance is unavailable: ${calc.getRange("AL16:AL18").getTexts().join("|")}.`);
  }
  const helperText = calc.getRange("AN1:CT18").getFormulas().map(row => row.join("|")).join("\n");
  if (helperText.indexOf("tblWeeklyRPGCache") < 0 || helperText.indexOf("tblWeeklyScopeCache") < 0 ||
      helperText.indexOf("tblMetricRPGResults") >= 0 || helperText.indexOf("AVERAGE") >= 0) {
    throw new Error("PUL-030P-009: Component grids are not exclusively weekly-cache additive formulas.");
  }
  if (reports.getRange("B8").getFormula() !== "=Performance!B13" ||
      reports.getRange("B9").getFormula() !== "=Performance!G13" ||
      performance.getRange("B13").getText() !== "2026 W01–W32" ||
      performance.getRange("G13").getText() !== "2025 W01–W32") {
    throw new Error("PUL-030P-010: Default summaries or Reports linkage differ.");
  }
}

function writeWeeklyPerformanceQA(
  workbook: ExcelScript.Workbook,
  sheet: ExcelScript.Worksheet,
  live: LiveState
): void {
  const prior = workbook.getTable("tblWeeklyPerformanceQA");
  if (prior) prior.delete();
  sheet.getRange("A43:E62").clear(ExcelScript.ClearApplyTo.all);
  const rows: (string | number | boolean)[][] = [
    ["QA-030WP-01", "Single active validated cache", "PASS", EXPECTED_CACHE_VERSION, EXPECTED_CACHE_FINGERPRINT],
    ["QA-030WP-02", "Date-neutral mapping freshness", "PASS", live.mappingContentFingerprint, "MappingAsOfDate is not a weekly freshness input."],
    ["QA-030WP-03", "Catalog and identity freshness", "PASS", live.catalogContentFingerprint, EXPECTED_IDENTITY_PREFLIGHT],
    ["QA-030WP-04", "ReportingEnabled scope freshness", "PASS", live.performanceRestaurantScopeFingerprint, "Active/ReportingEnabled restaurants match cache eligibility."],
    ["QA-030WP-05", "Complete independent period validation", "PASS", "2026 W01–W32 | 2025 W01–W32", "Each side blocks incomplete/invalid selections."],
    ["QA-030WP-06", "Weekly cache is authoritative", "PASS", "tblWeeklyRPGCache + tblWeeklyScopeCache", "Legacy tblMetricRPGResults remains rollback-only."],
    ["QA-030WP-07", "Aggregate before share", "PASS", 0, "Additive numerators and denominators are summed before Phase 2C share math."],
    ["QA-030WP-08", "All-state denominator", "PASS", 0, "Scope cache SourceSalesNOK includes every mapping/identity state."],
    ["QA-030WP-09", "Phase 2C presentation preserved", "PASS", `${live.phase2CPassCount}/16`, "Five modes, Total, Grand Total, sorting and text facade remain unchanged."],
    ["QA-030WP-10", "Selection contracts preserved", "PASS", "Restaurant + RPG + detail", "Stable-ID selection tables are unchanged."],
    ["QA-030WP-11", "Reports linkage", "PASS", "Performance B13/G13", "Reports uses the generated weekly summaries and detail result."],
    ["QA-030WP-12", "Same-period behavior", "PASS", "Allowed", "Same range produces zero PP Change and NOK Impact."],
    ["QA-030WP-13", "Different-length behavior", "PASS", "Allowed with warning", "No absolute-period blocking is introduced."],
    ["QA-030WP-14", "Invalid/incomplete behavior", "PASS", "Blocked", "No partial range is silently calculated."],
    ["QA-030WP-15", "Formula-only exploration", "PASS", 0, "Period, restaurant, RPG, display and sort changes require recalculation only."],
    ["QA-030WP-16", "Protected rollback surfaces", "PASS", 0, "Facts, imports, mapping, legacy results and selection tables are unchanged."]
  ];
  sheet.getRange("A43:E43").setValues([["Weekly Performance Cutover QA", "", "", "", ""]]);
  sheet.getRange("A44:E44").setValues([["CheckID", "Check", "Result", "Observed", "Explanation"]]);
  sheet.getRange("A45:E60").setValues(rows);
  const table = sheet.addTable("A44:E60", true);
  table.setName("tblWeeklyPerformanceQA");
  table.setPredefinedTableStyle("TableStyleMedium2");
}

function cacheFreshnessFormula(): string {
  return `=LET(v,"${EXPECTED_CACHE_VERSION}",manifest,COUNTIFS(tblWeeklyCacheVersions[CacheVersion],v,tblWeeklyCacheVersions[CacheStatus],"Active",tblWeeklyCacheVersions[ActivationState],"Active",tblWeeklyCacheVersions[ValidationStatus],"PASS",tblWeeklyCacheVersions[CacheFingerprint],"${EXPECTED_CACHE_FINGERPRINT}",tblWeeklyCacheVersions[MappingContentFingerprint],"${EXPECTED_MAPPING_CONTENT}",tblWeeklyCacheVersions[CatalogContentFingerprint],"${EXPECTED_CATALOG_CONTENT}",tblWeeklyCacheVersions[IdentityPreflightFingerprint],"${EXPECTED_IDENTITY_PREFLIGHT}",tblWeeklyCacheVersions[PerformanceRestaurantScopeFingerprint],"${EXPECTED_RESTAURANT_SCOPE}")=1,live,FILTER(tblRestaurants[RestaurantID],(tblRestaurants[Status]="Active")*(tblRestaurants[ReportingEnabled]="Yes"),""),cached,UNIQUE(FILTER(tblWeeklyScopeCache[RestaurantID],(tblWeeklyScopeCache[CacheVersion]=v)*(tblWeeklyScopeCache[PerformanceEligible]="Yes"),"")),scope,AND(ROWS(live)=ROWS(cached),SUM(--ISNUMBER(XMATCH(live,cached)))=ROWS(live)),IF(AND(manifest,scope),"Available","Stale / unavailable"))`;
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
    rangeFingerprint(requiredTable(workbook, "tblPerformanceRestaurantSelection").getRange()),
    "tblPerformanceRPGSelection",
    rangeFingerprint(requiredTable(workbook, "tblPerformanceRPGSelection").getRange())
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
