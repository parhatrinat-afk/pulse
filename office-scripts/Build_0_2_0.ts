/**
 * Pulse Build 0.2.0
 * First deterministic KPI vertical slice: Category Sales Share.
 *
 * Architecture:
 * - Category Sales Share = category sales / total sales inside the selected scope.
 * - KPI is channel-neutral.
 * - Current and comparison datasets are independently selectable.
 * - Current and comparison channels are independently selectable.
 * - A dataset is not permanently a "baseline"; it becomes the comparison reference when selected.
 * - Scope differences are surfaced as information, never blocked.
 * - Pulse does not make business recommendations.
 *
 * Run this script against Pulse Build 0.1.1 or a later compatible workbook.
 */
function main(workbook: ExcelScript.Workbook): string {
  const performance = requiredSheet(workbook, "Performance");
  const reports = requiredSheet(workbook, "Reports");
  const overview = requiredSheet(workbook, "Overview");
  const context = requiredSheet(workbook, "Context");
  const kpiRegistry = requiredSheet(workbook, "KPI Registry");
  const reportingCategories = requiredSheet(workbook, "Reporting Categories");
  const restaurants = requiredSheet(workbook, "Restaurants");
  const environment = requiredSheet(workbook, "_Environment");
  const buildLog = requiredSheet(workbook, "_Build_Log");

  const facts = requiredTable(workbook, "tblSalesFacts");
  const imports = requiredTable(workbook, "tblImports");
  const categoryTable = requiredTable(workbook, "tblReportingCategories");
  const restaurantTable = requiredTable(workbook, "tblRestaurants");
  const kpiTable = requiredTable(workbook, "tblKPIRegistry");
  const environmentTable = requiredTable(workbook, "tblEnvironment");
  const buildLogTable = requiredTable(workbook, "tblBuildLog");

  const calc = workbook.getWorksheet("_Metric_Calc") ?? workbook.addWorksheet("_Metric_Calc");

  const NAVY = "#172033";
  const BLUE = "#4F8CFF";
  const LIGHT = "#EAF2FF";
  const GREY = "#EEF1F5";
  const WHITE = "#FFFFFF";
  const MUTED = "#5B6677";
  const GREEN = "#116B46";
  const RED = "#A83126";

  workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.automatic);

  // ---------------------------------------------------------------------------
  // 1. Build dynamic selector lists from ACTIVE FINALIZED data already in Pulse.
  // ---------------------------------------------------------------------------
  const importRows = imports.getRangeBetweenHeaderAndTotal().getValues();
  const activeImports: {
    label: string;
    id: string;
    start: number;
    end: number;
    days: number;
    year: number;
    week: string;
  }[] = [];

  for (const row of importRows) {
    const importId = text(row[0]);
    if (!importId) continue;

    const publicationState = text(row[14]);
    const activeVersion = text(row[15]);
    if (publicationState !== "Active Finalized" || activeVersion !== "Yes") continue;

    const start = numberValue(row[5]);
    const end = numberValue(row[6]);
    const year = numberValue(row[7]);
    const weekRaw = text(row[8]);

    if (!Number.isFinite(start) || !Number.isFinite(end)) continue;

    activeImports.push({
      label: makePeriodLabel(start, end, year, weekRaw),
      id: importId,
      start,
      end,
      days: Math.round(end - start + 1),
      year,
      week: weekRaw
    });
  }

  if (activeImports.length === 0) {
    throw new Error("PUL-5001: No Active Finalized imports are available.");
  }

  const factRows = facts.getRangeBetweenHeaderAndTotal().getValues();

  const channels = new Set<string>();
  for (const row of factRows) {
    if (text(row[16]) !== "Active Finalized") continue;
    const channel = text(row[10]);
    if (channel) channels.add(channel);
  }

  const categoryRows = categoryTable.getRangeBetweenHeaderAndTotal().getValues();
  const categories: { id: string; name: string; order: number }[] = [];
  for (const row of categoryRows) {
    const id = text(row[0]);
    if (!id || text(row[4]) !== "Yes") continue;
    categories.push({
      id,
      name: text(row[1]) || "Uncategorized",
      order: numberValue(row[5]) || 999999
    });
  }
  categories.sort((a, b) => a.order - b.order || a.name.localeCompare(b.name));

  const restaurantRows = restaurantTable.getRangeBetweenHeaderAndTotal().getValues();
  const reportingRestaurants: { id: string; name: string }[] = [];
  for (const row of restaurantRows) {
    const id = text(row[0]);
    if (!id) continue;
    if (text(row[6]) !== "Active" || text(row[7]) !== "Yes") continue;
    reportingRestaurants.push({
      id,
      name: text(row[3]) || text(row[2]) || id
    });
  }
  reportingRestaurants.sort((a, b) => a.name.localeCompare(b.name));

  // ---------------------------------------------------------------------------
  // 2. _Metric_Calc holds selector lists and formula helpers.
  //    It is technical infrastructure, not a user-facing destination.
  // ---------------------------------------------------------------------------
  calc.getRange("A1:Q1000").clear(ExcelScript.ClearApplyTo.all);

  calc.getRange("A1:E1").setValues([[
    "Dataset", "ImportID", "PeriodStart", "PeriodEnd", "Days"
  ]]);
  if (activeImports.length > 0) {
    calc.getRange(`A2:E${activeImports.length + 1}`).setValues(
      activeImports.map(x => [x.label, x.id, x.start, x.end, x.days])
    );
  }

  const channelList = ["All channels", ...Array.from(channels).sort()];
  calc.getRange("G1").setValue("Channel");
  calc.getRange(`G2:G${channelList.length + 1}`).setValues(channelList.map(x => [x]));

  calc.getRange("I1:J1").setValues([["Category", "ReportingCategoryID"]]);
  if (categories.length > 0) {
    calc.getRange(`I2:J${categories.length + 1}`).setValues(
      categories.map(x => [x.name, x.id])
    );
  }

  calc.getRange("L1:M1").setValues([["Restaurant", "RestaurantID"]]);
  if (reportingRestaurants.length > 0) {
    calc.getRange(`L2:M${reportingRestaurants.length + 1}`).setValues(
      reportingRestaurants.map(x => [x.name, x.id])
    );
  }

  calc.getRange("A1:M1").getFormat().getFont().setBold(true);
  calc.getRange("A1:M1").getFormat().getFill().setColor(NAVY);
  calc.getRange("A1:M1").getFormat().getFont().setColor(WHITE);
  calc.getRange("C2:D100").setNumberFormatLocal("dd.mm.yyyy");

  // ---------------------------------------------------------------------------
  // 3. KPI Registry: only the implemented KPI is Active in Build 0.2.0.
  // ---------------------------------------------------------------------------
  const kpiBody = kpiTable.getRangeBetweenHeaderAndTotal();
  const kpiRows = kpiBody.getValues();
  for (let i = 0; i < kpiRows.length; i++) {
    const kpiId = text(kpiRows[i][0]);
    if (!kpiId) continue;

    if (kpiId === "KPI-0001") {
      kpiBody.getCell(i, 1).setValue("Category Sales Share");
      kpiBody.getCell(i, 4).setValue("Active");
      kpiBody.getCell(i, 5).setValue("Yes");
      kpiBody.getCell(i, 7).setValue("Overview");
      kpiBody.getCell(i, 8).setValue("Yes");
      kpiBody.getCell(i, 9).setValue("No");
      kpiBody.getCell(i, 10).setValue("No");
      kpiBody.getCell(i, 11).setValue("Yes");
    } else {
      // Preserve the definitions but do not present unfinished KPIs as available.
      kpiBody.getCell(i, 4).setValue("Draft");
      kpiBody.getCell(i, 5).setValue("No");
    }
  }

  // ---------------------------------------------------------------------------
  // 4. Performance — first real interactive KPI experience.
  // ---------------------------------------------------------------------------
  performance.getRange("A1:L90").clear(ExcelScript.ClearApplyTo.all);

  writeTitle(
    performance,
    "Performance",
    "Category Sales Share — current and comparison scopes are independently selectable.",
    "L",
    NAVY,
    LIGHT,
    WHITE
  );

  performance.getRange("A5:D5").setValues([["Metric", "", "", ""]]);
  styleSection(performance.getRange("A5:D5"), BLUE, WHITE);

  const defaultCategory =
    categories.find(x => x.name === "Add-ons")?.name ??
    categories[0]?.name ??
    "Uncategorized";

  performance.getRange("A6:B7").setValues([
    ["KPI", "Category Sales Share"],
    ["Category", defaultCategory]
  ]);
  styleLabels(performance.getRange("A6:A7"), GREY);
  performance.getRange("B7").getDataValidation().setRule({
    list: {
      inCellDropDown: true,
      source: `='_Metric_Calc'!$I$2:$I$${Math.max(2, categories.length + 1)}`
    }
  });

  performance.getRange("A9:B9").setValues([["Current", ""]]);
  styleSection(performance.getRange("A9:B9"), BLUE, WHITE);

  const latest = activeImports[activeImports.length - 1];
  const comparisonDefault =
    activeImports.length > 1 ? activeImports[activeImports.length - 2] : latest;

  const firstChannel =
    channels.has("In-house") ? "In-house" :
    channels.values().next().value ?? "All channels";

  performance.getRange("A10:B11").setValues([
    ["Dataset", latest.label],
    ["Channel", firstChannel]
  ]);
  styleLabels(performance.getRange("A10:A11"), GREY);

  performance.getRange("F9:G9").setValues([["Compare with", ""]]);
  styleSection(performance.getRange("F9:G9"), BLUE, WHITE);
  performance.getRange("F10:G11").setValues([
    ["Dataset", comparisonDefault.label],
    ["Channel", firstChannel]
  ]);
  styleLabels(performance.getRange("F10:F11"), GREY);

  const periodSource = `='_Metric_Calc'!$A$2:$A$${activeImports.length + 1}`;
  const channelSource = `='_Metric_Calc'!$G$2:$G$${channelList.length + 1}`;

  for (const address of ["B10", "G10"]) {
    performance.getRange(address).getDataValidation().setRule({
      list: { inCellDropDown: true, source: periodSource }
    });
  }

  for (const address of ["B11", "G11"]) {
    performance.getRange(address).getDataValidation().setRule({
      list: { inCellDropDown: true, source: channelSource }
    });
  }

  // Selected result.
  performance.getRange("A15:D15").setValues([["Selected category result", "", "", ""]]);
  styleSection(performance.getRange("A15:D15"), NAVY, WHITE);
  performance.getRange("A16:B19").setValues([
    ["Current share", ""],
    ["Comparison share", ""],
    ["Change", ""],
    ["Current category sales", ""]
  ]);
  styleLabels(performance.getRange("A16:A19"), GREY);

  const selectedCategoryId =
    `INDEX('_Metric_Calc'!$J$2:$J$${categories.length + 1},MATCH($B$7,'_Metric_Calc'!$I$2:$I$${categories.length + 1},0))`;
  const currentImportId =
    `INDEX('_Metric_Calc'!$B$2:$B$${activeImports.length + 1},MATCH($B$10,'_Metric_Calc'!$A$2:$A$${activeImports.length + 1},0))`;
  const compareImportId =
    `INDEX('_Metric_Calc'!$B$2:$B$${activeImports.length + 1},MATCH($G$10,'_Metric_Calc'!$A$2:$A$${activeImports.length + 1},0))`;

  performance.getRange("B16").setFormula(
    categoryShareFormula(selectedCategoryId, currentImportId, "$B$11")
  );
  performance.getRange("B17").setFormula(
    categoryShareFormula(selectedCategoryId, compareImportId, "$G$11")
  );
  performance.getRange("B18").setFormula("=B16-B17");
  performance.getRange("B19").setFormula(
    categorySalesFormula(selectedCategoryId, currentImportId, "$B$11")
  );
  performance.getRange("B16:B18").setNumberFormat("0.00%");
  performance.getRange("B19").setNumberFormat('#,##0 "NOK"');

  performance.getRange("B18").addConditionalFormat(ExcelScript.ConditionalFormatType.custom)
    .getCustom().getRule().setFormula("=B18<0");
  performance.getRange("B18").getConditionalFormats().getItemAt(0)
    .getCustom().getFormat().getFont().setColor(RED);

  // Comparison check. Informational only.
  performance.getRange("F15:I15").setValues([["Comparison check", "", "", ""]]);
  styleSection(performance.getRange("F15:I15"), NAVY, WHITE);
  performance.getRange("F16:G19").setValues([
    ["Channel", ""],
    ["Period length", ""],
    ["Same dataset", ""],
    ["Status", ""]
  ]);
  styleLabels(performance.getRange("F16:F19"), GREY);

  performance.getRange("G16").setFormula('=IF($B$11=$G$11,"Same","Different")');
  performance.getRange("G17").setFormula(
    `=IF(` +
    `INDEX('_Metric_Calc'!$E$2:$E$${activeImports.length + 1},MATCH($B$10,'_Metric_Calc'!$A$2:$A$${activeImports.length + 1},0))=` +
    `INDEX('_Metric_Calc'!$E$2:$E$${activeImports.length + 1},MATCH($G$10,'_Metric_Calc'!$A$2:$A$${activeImports.length + 1},0)),` +
    `"Same","Different")`
  );
  performance.getRange("G18").setFormula('=IF($B$10=$G$10,"Yes","No")');
  performance.getRange("G19").setFormula(
    '=IF(AND(G16="Same",G17="Same"),"Comparable scope","Scope differs — intentional comparisons are allowed")'
  );
  performance.getRange("G19:I19").getFormat().setWrapText(true);

  // Category overview uses the same current/comparison selectors.
  performance.getRange("A23:E23").setValues([[
    "Category", "Current", "Comparison", "Change", "Current sales"
  ]]);
  styleHeader(performance.getRange("A23:E23"), GREY, "#172033");

  const categoryStartRow = 24;
  for (let i = 0; i < categories.length; i++) {
    const row = categoryStartRow + i;
    const calcRow = i + 2;
    performance.getRange(`A${row}`).setFormula(`='_Metric_Calc'!I${calcRow}`);
    const categoryIdExpr = `'_Metric_Calc'!J${calcRow}`;
    performance.getRange(`B${row}`).setFormula(
      categoryShareFormula(categoryIdExpr, currentImportId, "$B$11")
    );
    performance.getRange(`C${row}`).setFormula(
      categoryShareFormula(categoryIdExpr, compareImportId, "$G$11")
    );
    performance.getRange(`D${row}`).setFormula(`=B${row}-C${row}`);
    performance.getRange(`E${row}`).setFormula(
      categorySalesFormula(categoryIdExpr, currentImportId, "$B$11")
    );
  }

  if (categories.length > 0) {
    const end = categoryStartRow + categories.length - 1;
    performance.getRange(`B${categoryStartRow}:D${end}`).setNumberFormat("0.00%");
    performance.getRange(`E${categoryStartRow}:E${end}`).setNumberFormat('#,##0 "NOK"');
  }

  // Restaurant breakdown for the selected category.
  const restaurantHeaderRow = categoryStartRow + categories.length + 3;
  performance.getRange(`A${restaurantHeaderRow}:E${restaurantHeaderRow}`).setValues([[
    "Restaurant", "Current", "Comparison", "Change", "Current sales"
  ]]);
  styleHeader(
    performance.getRange(`A${restaurantHeaderRow}:E${restaurantHeaderRow}`),
    GREY,
    "#172033"
  );

  for (let i = 0; i < reportingRestaurants.length; i++) {
    const row = restaurantHeaderRow + 1 + i;
    const calcRow = i + 2;
    performance.getRange(`A${row}`).setFormula(`='_Metric_Calc'!L${calcRow}`);
    const restaurantIdExpr = `'_Metric_Calc'!M${calcRow}`;

    performance.getRange(`B${row}`).setFormula(
      categoryShareFormula(selectedCategoryId, currentImportId, "$B$11", restaurantIdExpr)
    );
    performance.getRange(`C${row}`).setFormula(
      categoryShareFormula(selectedCategoryId, compareImportId, "$G$11", restaurantIdExpr)
    );
    performance.getRange(`D${row}`).setFormula(`=B${row}-C${row}`);
    performance.getRange(`E${row}`).setFormula(
      categorySalesFormula(selectedCategoryId, currentImportId, "$B$11", restaurantIdExpr)
    );
  }

  if (reportingRestaurants.length > 0) {
    const start = restaurantHeaderRow + 1;
    const end = restaurantHeaderRow + reportingRestaurants.length;
    performance.getRange(`B${start}:D${end}`).setNumberFormat("0.00%");
    performance.getRange(`E${start}:E${end}`).setNumberFormat('#,##0 "NOK"');
  }

  // Explainability.
  const explainRow = restaurantHeaderRow + reportingRestaurants.length + 4;
  performance.getRange(`A${explainRow}:H${explainRow}`).setValues([[
    "Explain", "", "", "", "", "", "", ""
  ]]);
  styleSection(performance.getRange(`A${explainRow}:H${explainRow}`), BLUE, WHITE);

  performance.getRange(`A${explainRow + 1}:H${explainRow + 5}`).setValues([
    ["Metric", "Category Sales Share", "", "", "", "", "", ""],
    ["Definition", "Category sales ÷ total sales inside the selected scope.", "", "", "", "", "", ""],
    ["Current / comparison", "Datasets and channels are selected independently.", "", "", "", "", "", ""],
    ["Guardrail", "Scope differences are shown but never block an intentional comparison.", "", "", "", "", "", ""],
    ["Decision ownership", "Pulse presents evidence. Interpretation and decisions remain human.", "", "", "", "", "", ""]
  ]);
  performance.getRange(`A${explainRow + 1}:A${explainRow + 5}`)
    .getFormat().getFont().setBold(true);
  performance.getRange(`A${explainRow + 1}:A${explainRow + 5}`)
    .getFormat().getFont().setColor(MUTED);
  performance.getRange(`B${explainRow + 1}:H${explainRow + 5}`)
    .getFormat().setWrapText(true);

  performance.getFreezePanes().freezeRows(2);

  performance.getRange("A:L").getFormat().setColumnWidth(16);
  performance.getRange("A:A").getFormat().setColumnWidth(24);
  performance.getRange("B:B").getFormat().setColumnWidth(28);
  performance.getRange("F:F").getFormat().setColumnWidth(22);
  performance.getRange("G:G").getFormat().setColumnWidth(34);

  // ---------------------------------------------------------------------------
  // 5. Reports consumes exactly the same selected context/results.
  // ---------------------------------------------------------------------------
  reports.getRange("A1:H35").clear(ExcelScript.ClearApplyTo.all);
  writeTitle(
    reports,
    "Reports",
    "Preview uses the same KPI result and the same current/comparison scopes selected in Performance.",
    "H",
    NAVY,
    LIGHT,
    WHITE
  );

  reports.getRange("A5:D5").setValues([["Report context", "", "", ""]]);
  styleSection(reports.getRange("A5:D5"), BLUE, WHITE);
  reports.getRange("A6:B10").setValues([
    ["KPI", "Category Sales Share"],
    ["Category", ""],
    ["Current", ""],
    ["Compare with", ""],
    ["Scope check", ""]
  ]);
  styleLabels(reports.getRange("A6:A10"), GREY);

  reports.getRange("B7").setFormula("=Performance!B7");
  reports.getRange("B8").setFormula('=Performance!B10&" · "&Performance!B11');
  reports.getRange("B9").setFormula('=Performance!G10&" · "&Performance!G11');
  reports.getRange("B10").setFormula("=Performance!G19");
  reports.getRange("B10:D10").getFormat().setWrapText(true);

  reports.getRange("A14:D14").setValues([["Selected result", "", "", ""]]);
  styleSection(reports.getRange("A14:D14"), NAVY, WHITE);
  reports.getRange("A15:B18").setValues([
    ["Current share", ""],
    ["Comparison share", ""],
    ["Change", ""],
    ["Current category sales", ""]
  ]);
  styleLabels(reports.getRange("A15:A18"), GREY);
  reports.getRange("B15").setFormula("=Performance!B16");
  reports.getRange("B16").setFormula("=Performance!B17");
  reports.getRange("B17").setFormula("=Performance!B18");
  reports.getRange("B18").setFormula("=Performance!B19");
  reports.getRange("B15:B17").setNumberFormat("0.00%");
  reports.getRange("B18").setNumberFormat('#,##0 "NOK"');

  reports.getRange("A:H").getFormat().setColumnWidth(18);
  reports.getRange("A:A").getFormat().setColumnWidth(26);
  reports.getRange("B:B").getFormat().setColumnWidth(38);
  reports.getFreezePanes().freezeRows(2);

  // ---------------------------------------------------------------------------
  // 6. Context language: dataset ≠ baseline; comparison is independently selected.
  // ---------------------------------------------------------------------------
  context.getRange("A8:C13").setValues([
    ["PeriodType", "Selected Dataset", "Selected Dataset / Week / Month / YTD / Custom"],
    ["PeriodStart", "", "Derived from the selected current dataset or period."],
    ["PeriodEnd", "", "Derived from the selected current dataset or period."],
    ["ComparisonType", "Selected Dataset", "Any available valid dataset or period may be selected."],
    ["ComparisonStart", "", "Derived independently from the comparison selection."],
    ["ComparisonEnd", "", "Derived independently from the comparison selection."]
  ]);

  context.getRange("F8:H10").setValues([
    ["4", "Period determines current facts.", "Current and comparison do not have to use the same period type."],
    ["5", "Comparison is independent.", "Pulse may flag scope differences but does not block intentional comparisons."],
    ["6", "Dataset is not baseline.", "A dataset becomes the baseline only when selected as the comparison reference."]
  ]);

  // ---------------------------------------------------------------------------
  // 7. Overview / build metadata.
  // ---------------------------------------------------------------------------
  overview.getRange("A15:B18").setValues([
    ["Latest published dataset", latest.label],
    ["Comparison", "User selected"],
    ["Available KPI", "Category Sales Share"],
    ["Channel", "User selected"]
  ]);
  overview.getRange("A20").setValue("Build 0.2.0 · Category Sales Share");

  updateEnvironment(environmentTable, "BuildID", "0.2.0", "First live KPI");
  updateEnvironment(environmentTable, "BuildVersion", "0.2.0", "Category Sales Share");
  updateEnvironment(
    environmentTable,
    "BuildDate",
    excelNow(),
    "Build generated by office-scripts/Build_0_2_0.ts"
  );

  appendBuildLog(buildLogTable, [
    nextLogId(buildLogTable),
    excelNow(),
    "0.2.0",
    "Implement Category Sales Share",
    "Success",
    "Metric vertical slice",
    "Channel-neutral KPI with independent current/comparison dataset and channel selection. Scope differences warn but do not block."
  ]);

  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  return (
    `Pulse Build 0.2.0 applied. ${activeImports.length} published dataset(s), ` +
    `${categories.length} category/categories, ${reportingRestaurants.length} reporting restaurant(s).`
  );
}

// =============================================================================
// Formula helpers
// =============================================================================

function categoryShareFormula(
  categoryIdExpr: string,
  importIdExpr: string,
  channelCell: string,
  restaurantIdExpr?: string
): string {
  const baseCriteria =
    `tblSalesFacts[ReportingCategoryID],${categoryIdExpr},` +
    `tblSalesFacts[ImportID],${importIdExpr},` +
    `tblSalesFacts[PublicationState],"Active Finalized"`;

  const denominatorCriteria =
    `tblSalesFacts[ImportID],${importIdExpr},` +
    `tblSalesFacts[PublicationState],"Active Finalized"`;

  const restaurantNumerator =
    restaurantIdExpr ? `,tblSalesFacts[RestaurantID],${restaurantIdExpr}` : "";
  const restaurantDenominator =
    restaurantIdExpr ? `,tblSalesFacts[RestaurantID],${restaurantIdExpr}` : "";

  const allChannels =
    `IFERROR(` +
    `SUMIFS(tblSalesFacts[SalesAmount],${baseCriteria}${restaurantNumerator})/` +
    `SUMIFS(tblSalesFacts[SalesAmount],${denominatorCriteria}${restaurantDenominator}),0)`;

  const selectedChannel =
    `IFERROR(` +
    `SUMIFS(tblSalesFacts[SalesAmount],${baseCriteria},tblSalesFacts[ReportingChannel],${channelCell}${restaurantNumerator})/` +
    `SUMIFS(tblSalesFacts[SalesAmount],${denominatorCriteria},tblSalesFacts[ReportingChannel],${channelCell}${restaurantDenominator}),0)`;

  return `=IF(${channelCell}="All channels",${allChannels},${selectedChannel})`;
}

function categorySalesFormula(
  categoryIdExpr: string,
  importIdExpr: string,
  channelCell: string,
  restaurantIdExpr?: string
): string {
  const restaurant =
    restaurantIdExpr ? `,tblSalesFacts[RestaurantID],${restaurantIdExpr}` : "";

  const allChannels =
    `SUMIFS(` +
    `tblSalesFacts[SalesAmount],` +
    `tblSalesFacts[ReportingCategoryID],${categoryIdExpr},` +
    `tblSalesFacts[ImportID],${importIdExpr},` +
    `tblSalesFacts[PublicationState],"Active Finalized"${restaurant})`;

  const selectedChannel =
    `SUMIFS(` +
    `tblSalesFacts[SalesAmount],` +
    `tblSalesFacts[ReportingCategoryID],${categoryIdExpr},` +
    `tblSalesFacts[ImportID],${importIdExpr},` +
    `tblSalesFacts[ReportingChannel],${channelCell},` +
    `tblSalesFacts[PublicationState],"Active Finalized"${restaurant})`;

  return `=IF(${channelCell}="All channels",${allChannels},${selectedChannel})`;
}

// =============================================================================
// Workbook helpers
// =============================================================================

function requiredSheet(
  workbook: ExcelScript.Workbook,
  name: string
): ExcelScript.Worksheet {
  const sheet = workbook.getWorksheet(name);
  if (!sheet) throw new Error(`PUL-5002: Required worksheet missing: ${name}`);
  return sheet;
}

function requiredTable(
  workbook: ExcelScript.Workbook,
  name: string
): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-5003: Required table missing: ${name}`);
  return table;
}

function makePeriodLabel(
  start: number,
  end: number,
  year: number,
  week: string
): string {
  if (week) return `Week ${week}, ${year}`;

  const startDate = excelSerialToDate(start);
  const endDate = excelSerialToDate(end);

  if (
    startDate.getUTCMonth() === 0 &&
    startDate.getUTCDate() === 1 &&
    endDate.getUTCMonth() === 11 &&
    endDate.getUTCDate() === 31 &&
    startDate.getUTCFullYear() === endDate.getUTCFullYear()
  ) {
    return `${startDate.getUTCFullYear()} full year`;
  }

  return `${dateLabel(start)}–${dateLabel(end)}`;
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
  sheet.getRange(`A1:${endColumn}1`).setValues([
    [title, ...new Array(columnNumber(endColumn) - 1).fill("")]
  ]);
  sheet.getRange(`A1:${endColumn}1`).getFormat().getFill().setColor(navy);
  sheet.getRange(`A1:${endColumn}1`).getFormat().getFont().setColor(white);
  sheet.getRange(`A1:${endColumn}1`).getFormat().getFont().setBold(true);
  sheet.getRange(`A1:${endColumn}1`).getFormat().getFont().setSize(19);
  sheet.getRange(`A1:${endColumn}1`).getFormat().setRowHeight(34);

  sheet.getRange(`A2:${endColumn}2`).setValues([
    [subtitle, ...new Array(columnNumber(endColumn) - 1).fill("")]
  ]);
  sheet.getRange(`A2:${endColumn}2`).getFormat().getFill().setColor(light);
  sheet.getRange(`A2:${endColumn}2`).getFormat().setWrapText(true);
  sheet.getRange(`A2:${endColumn}2`).getFormat().setRowHeight(30);
}

function styleSection(
  range: ExcelScript.Range,
  fill: string,
  fontColor: string
): void {
  range.getFormat().getFill().setColor(fill);
  range.getFormat().getFont().setColor(fontColor);
  range.getFormat().getFont().setBold(true);
}

function styleHeader(
  range: ExcelScript.Range,
  fill: string,
  fontColor: string
): void {
  range.getFormat().getFill().setColor(fill);
  range.getFormat().getFont().setColor(fontColor);
  range.getFormat().getFont().setBold(true);
  range.getFormat().setWrapText(true);
}

function styleLabels(range: ExcelScript.Range, fill: string): void {
  range.getFormat().getFill().setColor(fill);
  range.getFormat().getFont().setBold(true);
}

function updateEnvironment(
  table: ExcelScript.Table,
  key: string,
  value: string | number,
  note: string
): void {
  const body = table.getRangeBetweenHeaderAndTotal();
  const rows = body.getValues();

  for (let i = 0; i < rows.length; i++) {
    if (text(rows[i][0]) === key) {
      body.getCell(i, 1).setValue(value);
      body.getCell(i, 2).setValue(note);
      return;
    }
  }

  table.addRow(-1, [key, value, note]);
}

function appendBuildLog(
  table: ExcelScript.Table,
  row: (string | number | boolean)[]
): void {
  table.addRow(-1, row);
}

function nextLogId(table: ExcelScript.Table): string {
  let max = 0;
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  for (const row of rows) {
    const id = text(row[0]);
    const match = id.match(/^LOG-(\d+)$/);
    if (!match) continue;
    const n = Number(match[1]);
    if (Number.isFinite(n)) max = Math.max(max, n);
  }
  return `LOG-${String(max + 1).padStart(6, "0")}`;
}

function excelNow(): number {
  return Date.now() / 86400000 + 25569;
}

function excelSerialToDate(serial: number): Date {
  return new Date((serial - 25569) * 86400000);
}

function dateLabel(serial: number): string {
  const d = excelSerialToDate(serial);
  const day = String(d.getUTCDate()).padStart(2, "0");
  const month = String(d.getUTCMonth() + 1).padStart(2, "0");
  return `${day}.${month}.${d.getUTCFullYear()}`;
}

function columnNumber(column: string): number {
  let result = 0;
  for (const char of column.toUpperCase()) {
    result = result * 26 + char.charCodeAt(0) - 64;
  }
  return result;
}

function numberValue(value: unknown): number {
  const n = Number(value);
  return Number.isFinite(n) ? n : 0;
}

function text(value: unknown): string {
  return String(value ?? "").trim();
}
