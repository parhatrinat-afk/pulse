/**
 * Pulse Build 0.3.0 — live management Overview facade.
 *
 * Overview deliberately contains no business calculation engine. It projects
 * stable outputs from the weekly Performance authority plus the existing
 * Imports and Mapping operational facades. Periods, scope, metrics, rankings,
 * freshness, mapping coverage, and import readiness remain owned upstream.
 */
function main(workbook: ExcelScript.Workbook): string {
  const overview = requiredSheet(workbook, "Overview");
  const performance = requiredSheet(workbook, "Performance");
  const reports = requiredSheet(workbook, "Reports");
  const imports = requiredSheet(workbook, "Imports");
  const mapping = requiredSheet(workbook, "Mapping");
  const settings = requiredSheet(workbook, "Settings");
  const calc = requiredSheet(workbook, "_Metric_Calc");

  if (overview.getTables().length !== 0) {
    throw new Error("PUL-030OV-001: Overview contains an unexpected table; no facade changes were applied.");
  }
  validateQa(requiredTable(workbook, "tblMappingQA"), /^QA-0301-/, 9, "Mapping");
  validateQa(requiredTable(workbook, "tblPerformanceInteractionQA"), /^QA-0302C-/, 16, "Phase 2C");
  validateQa(requiredTable(workbook, "tblWeeklyPerformanceQA"), /^QA-030WP-/, 16, "Weekly Performance");
  validatePrimarySheets(workbook, overview, performance, reports, imports, mapping, settings);
  validateManagementAuthority(calc);
  const protectedBefore = protectedState(workbook, performance, reports, imports, mapping, calc);

  buildOverviewFacade(overview);
  workbook.getApplication().setCalculationMode(ExcelScript.CalculationMode.automatic);
  workbook.getApplication().calculate(ExcelScript.CalculationType.full);

  validateOverviewPostconditions(overview, calc, imports, mapping);
  const protectedAfter = protectedState(workbook, performance, reports, imports, mapping, calc);
  if (protectedBefore !== protectedAfter) {
    throw new Error("PUL-030OV-002: An upstream authority or accepted QA surface changed while building Overview.");
  }

  overview.activate();
  overview.getRange("A1").select();
  return JSON.stringify({
    status: "PASS",
    current: overview.getRange("A7").getText(),
    comparison: overview.getRange("D7").getText(),
    latest: overview.getRange("G7").getText(),
    performance: calc.getRange("AL32").getText(),
    display: calc.getRange("AL38").getText(),
    navigation: "Performance | Reports | Imports | Mapping | Settings",
    authority: "Weekly Performance + Imports + Mapping facades"
  });
}

const NAVY = "#17365D";
const BLUE = "#1F4E78";
const LIGHT_BLUE = "#D9EAF7";
const PALE_BLUE = "#EAF2FF";
const PALE_YELLOW = "#FFF4CE";
const WHITE = "#FFFFFF";
const VERY_LIGHT = "#F7F9FC";
const INK = "#172033";
const MUTED = "#5B6573";
const RED = "#A83126";
const GREEN = "#E2F0D9";
const DASH = "—";

function buildOverviewFacade(sheet: ExcelScript.Worksheet): void {
  sheet.getRange("A1:J43").unmerge();
  sheet.getRange("A1:J43").clear(ExcelScript.ClearApplyTo.all);
  sheet.setShowGridlines(false);

  sheet.getRange("A1:J1").merge();
  sheet.getRange("A2:J2").merge();
  sheet.getRange("A3:B3").merge();
  sheet.getRange("C3:D3").merge();
  sheet.getRange("E3:F3").merge();
  sheet.getRange("G3:H3").merge();
  sheet.getRange("I3:J3").merge();
  sheet.getRange("A5:J5").merge();
  sheet.getRange("A6:C6").merge();
  sheet.getRange("D6:F6").merge();
  sheet.getRange("G6:J6").merge();
  sheet.getRange("A7:C7").merge();
  sheet.getRange("D7:F7").merge();
  sheet.getRange("G7:J7").merge();
  sheet.getRange("A8:C8").merge();
  sheet.getRange("D8:F8").merge();
  sheet.getRange("G8:J8").merge();
  sheet.getRange("A9:J9").merge();
  sheet.getRange("A10:J10").merge();
  sheet.getRange("A12:J12").merge();
  mergeCard(sheet, "A13:B13", "A14:B16");
  mergeCard(sheet, "C13:D13", "C14:D16");
  mergeCard(sheet, "E13:F13", "E14:F16");
  mergeCard(sheet, "G13:H13", "G14:H16");
  mergeCard(sheet, "I13:J13", "I14:J16");
  sheet.getRange("A17:J17").merge();
  sheet.getRange("A19:E19").merge();
  sheet.getRange("F19:J19").merge();
  sheet.getRange("A20:E20").merge();
  sheet.getRange("F20:J20").merge();
  mergeRankingRow(sheet, 21);
  mergeRankingRow(sheet, 22);
  mergeRankingRow(sheet, 23);
  sheet.getRange("A24:E24").merge();
  sheet.getRange("F24:J24").merge();
  mergeRankingRow(sheet, 25);
  mergeRankingRow(sheet, 26);
  mergeRankingRow(sheet, 27);
  sheet.getRange("A29:E29").merge();
  sheet.getRange("F29:J29").merge();
  mergeAttentionRow(sheet, 30);
  mergeAttentionRow(sheet, 31);
  mergeAttentionRow(sheet, 32);
  mergeAttentionRow(sheet, 33);
  mergeAttentionRow(sheet, 34);
  sheet.getRange("A36:E36").merge();
  sheet.getRange("F30:H30").merge(); sheet.getRange("I30:J30").merge();
  sheet.getRange("F31:H31").merge(); sheet.getRange("I31:J31").merge();
  sheet.getRange("F32:H32").merge(); sheet.getRange("I32:J32").merge();
  sheet.getRange("F33:H33").merge(); sheet.getRange("I33:J33").merge();
  sheet.getRange("F34:H34").merge(); sheet.getRange("I34:J34").merge();
  sheet.getRange("F36:J36").merge();

  sheet.getRange("A1").setValue("Pulse");
  sheet.getRange("A2").setValue("Operational performance at a glance.");
  setInternalLink(sheet.getRange("A3"), "Performance", "Change analytical periods, scope, Display, and sorting.");
  setInternalLink(sheet.getRange("C3"), "Reports", "Open the linked reporting view.");
  setInternalLink(sheet.getRange("E3"), "Imports", "Review publication and historical coverage.");
  setInternalLink(sheet.getRange("G3"), "Mapping", "Review mapping coverage and classifications.");
  setInternalLink(sheet.getRange("I3"), "Settings", "Manage Pulse configuration.");

  sheet.getRange("A5").setValue("Analysis Context");
  sheet.getRange("A6").setValue("Current");
  sheet.getRange("D6").setValue("Compare");
  sheet.getRange("G6").setValue("Latest available");
  sheet.getRange("A7").setFormula("='_Metric_Calc'!$AL$33");
  sheet.getRange("D7").setFormula("='_Metric_Calc'!$AL$34");
  sheet.getRange("G7").setFormula("='_Metric_Calc'!$AL$36");
  sheet.getRange("A8").setFormula('=IF(\'_Metric_Calc\'!$AL$17="Valid",TEXT(\'_Metric_Calc\'!$AL$25,"0")&" weeks","—")');
  sheet.getRange("D8").setFormula('=IF(\'_Metric_Calc\'!$AL$18="Valid",TEXT(\'_Metric_Calc\'!$AL$26,"0")&" weeks","—")');
  sheet.getRange("G8").setFormula("='_Metric_Calc'!$AL$37");
  sheet.getRange("A9").setFormula("='_Metric_Calc'!$AL$35");
  sheet.getRange("A10").setValue("Change periods, scope, Display, and sorting on Performance.");

  sheet.getRange("A12").setValue("Selected Performance");
  sheet.getRange("A13").setValue("Total Sales");
  sheet.getRange("C13").setValue("Selected Category Sales");
  sheet.getRange("E13").setValue("Sales Share");
  sheet.getRange("G13").setValue("PP Change");
  sheet.getRange("I13").setValue("NOK Impact");
  sheet.getRange("A14").setFormula("='_Metric_Calc'!$AL$41");
  sheet.getRange("C14").setFormula("='_Metric_Calc'!$AL$42");
  sheet.getRange("E14").setFormula("='_Metric_Calc'!$AL$43");
  sheet.getRange("G14").setFormula("='_Metric_Calc'!$AL$44");
  sheet.getRange("I14").setFormula("='_Metric_Calc'!$AL$45");
  sheet.getRange("A17").setFormula('="Scope: "&\'_Metric_Calc\'!$AL$39&" · "&\'_Metric_Calc\'!$AL$40&" · Display: "&\'_Metric_Calc\'!$AL$38');

  sheet.getRange("A19").setFormula('="Reporting Groups — "&\'_Metric_Calc\'!$AL$38');
  sheet.getRange("F19").setFormula('="Restaurants — "&\'_Metric_Calc\'!$AL$38');
  sheet.getRange("A20").setValue("Top 3");
  sheet.getRange("F20").setValue("Top 3");
  sheet.getRange("A24").setValue("Bottom 3");
  sheet.getRange("F24").setValue("Bottom 3");
  writeRankingProjection(sheet, "A21", "E21", "AP33", "AQ33");
  writeRankingProjection(sheet, "A22", "E22", "AP34", "AQ34");
  writeRankingProjection(sheet, "A23", "E23", "AP35", "AQ35");
  writeRankingProjection(sheet, "A25", "E25", "AP36", "AQ36");
  writeRankingProjection(sheet, "A26", "E26", "AP37", "AQ37");
  writeRankingProjection(sheet, "A27", "E27", "AP38", "AQ38");
  writeRankingProjection(sheet, "F21", "J21", "AU33", "AV33");
  writeRankingProjection(sheet, "F22", "J22", "AU34", "AV34");
  writeRankingProjection(sheet, "F23", "J23", "AU35", "AV35");
  writeRankingProjection(sheet, "F25", "J25", "AU36", "AV36");
  writeRankingProjection(sheet, "F26", "J26", "AU37", "AV37");
  writeRankingProjection(sheet, "F27", "J27", "AU38", "AV38");

  sheet.getRange("A29").setValue("Attention");
  sheet.getRange("F29").setValue("Data Status");
  sheet.getRange("A30").setFormula('=IF(Mapping!$D$8>0,"Unmapped historical sales","")');
  sheet.getRange("D30").setFormula('=IF(Mapping!$D$8>0,Mapping!$D$8,"")');
  sheet.getRange("A31").setFormula('=IF(Mapping!$F$6>0,"Identity Pending","")');
  sheet.getRange("D31").setFormula('=IF(Mapping!$F$6>0,FIXED(Mapping!$F$6,0,TRUE)&" products · "&FIXED(Mapping!$F$8,0,FALSE)&" NOK","")');
  sheet.getRange("A32").setFormula('=IF(\'_Metric_Calc\'!$AL$32<>"Available","Performance refresh required","")');
  sheet.getRange("D32").setFormula('=IF(\'_Metric_Calc\'!$AL$32<>"Available","Performance values are unavailable until refreshed.","")');
  sheet.getRange("A33").setFormula('=IF(Mapping!$H$6>0,"Mapping conflict","")');
  sheet.getRange("D33").setFormula('=IF(Mapping!$H$6>0,FIXED(Mapping!$H$6,0,TRUE)&" products · "&FIXED(Mapping!$H$8,0,FALSE)&" NOK","")');
  sheet.getRange("A34").setFormula('=IF(Mapping!$J$6>0,"Inactive target","")');
  sheet.getRange("D34").setFormula('=IF(Mapping!$J$6>0,FIXED(Mapping!$J$6,0,TRUE)&" products · "&FIXED(Mapping!$J$8,0,FALSE)&" NOK","")');
  sheet.getRange("A36").setValue("Mapping amounts are historical coverage across the active weekly history.");
  sheet.getRange("F30").setValue("Latest published");
  sheet.getRange("F31").setValue("Historical coverage");
  sheet.getRange("F32").setValue("Imports");
  sheet.getRange("F33").setValue("Classifications");
  sheet.getRange("F34").setValue("Performance");
  sheet.getRange("I30").setFormula("=Imports!$A$9");
  sheet.getRange("I31").setFormula('=SUBSTITUTE(Imports!$D$9,CHAR(10)," · ")');
  sheet.getRange("I32").setFormula("=Imports!$G$9");
  sheet.getRange("I33").setFormula("=Mapping!$A$9");
  sheet.getRange("I34").setFormula("='_Metric_Calc'!$AL$32");
  sheet.getRange("F36").setValue("Open Performance for detailed analysis and scope changes.");

  formatOverview(sheet);
}

function mergeCard(sheet: ExcelScript.Worksheet, labelAddress: string, valueAddress: string): void {
  sheet.getRange(labelAddress).merge();
  sheet.getRange(valueAddress).merge();
}

function mergeRankingRow(sheet: ExcelScript.Worksheet, row: number): void {
  sheet.getRange(`A${row}:D${row}`).merge();
  sheet.getRange(`F${row}:I${row}`).merge();
}

function mergeAttentionRow(sheet: ExcelScript.Worksheet, row: number): void {
  sheet.getRange(`A${row}:C${row}`).merge();
  sheet.getRange(`D${row}:E${row}`).merge();
}

function writeRankingProjection(
  sheet: ExcelScript.Worksheet, nameCell: string, valueCell: string,
  authorityNameCell: string, authorityValueCell: string
): void {
  sheet.getRange(nameCell).setFormula(`='_Metric_Calc'!$${authorityNameCell}`);
  sheet.getRange(valueCell).setFormula(rankingPresentationFormula(authorityValueCell));
}

function rankingPresentationFormula(authorityValueCell: string): string {
  const value = `'_Metric_Calc'!$${authorityValueCell}`;
  return `=IF(${value}="","${DASH}",IF('_Metric_Calc'!$AL$38="PP Change",IF(${value}>0,"+","")&FIXED(${value},2,TRUE)&" pp",IF(OR('_Metric_Calc'!$AL$38="Current Share",'_Metric_Calc'!$AL$38="Comparison Share"),FIXED(${value}*100,2,TRUE)&"%",IF('_Metric_Calc'!$AL$38="Current Sales NOK",FIXED(${value},0,FALSE)&" NOK",IF('_Metric_Calc'!$AL$38="NOK Impact",IF(${value}>0,"+","")&FIXED(${value},0,FALSE)&" NOK","${DASH}")))))`;
}

function formatOverview(sheet: ExcelScript.Worksheet): void {
  sheet.getRange("A1:J36").getFormat().getFont().setName("Carlito");
  sheet.getRange("A1:J36").getFormat().getFont().setSize(11);
  sheet.getRange("A1:J36").getFormat().getFont().setColor(INK);
  sheet.getRange("A1:J36").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  sheet.getRange("A1:J36").getFormat().setWrapText(false);
  sheet.getRange("A1:J36").getFormat().getFill().setColor(WHITE);

  sheet.getRange("A1:J1").getFormat().getFill().setColor(NAVY);
  sheet.getRange("A1:J1").getFormat().getFont().setColor(WHITE);
  sheet.getRange("A1:J1").getFormat().getFont().setBold(true);
  sheet.getRange("A1:J1").getFormat().getFont().setSize(20);
  sheet.getRange("A2:J2").getFormat().getFill().setColor(LIGHT_BLUE);
  sheet.getRange("A2:J2").getFormat().getFont().setColor(NAVY);

  sheet.getRange("A3:J3").getFormat().getFill().setColor(PALE_BLUE);
  sheet.getRange("A3:J3").getFormat().getFont().setColor(BLUE);
  sheet.getRange("A3:J3").getFormat().getFont().setBold(true);
  sheet.getRange("A3:J3").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  formatSectionHeader(sheet.getRange("A5:J5"));
  formatSectionHeader(sheet.getRange("A12:J12"));
  formatSectionHeader(sheet.getRange("A19:E19"));
  formatSectionHeader(sheet.getRange("F19:J19"));
  formatSectionHeader(sheet.getRange("A29:E29"));
  formatSectionHeader(sheet.getRange("F29:J29"));

  sheet.getRange("A6:J6").getFormat().getFill().setColor(LIGHT_BLUE);
  sheet.getRange("A6:J6").getFormat().getFont().setColor(NAVY);
  sheet.getRange("A6:J6").getFormat().getFont().setBold(true);
  sheet.getRange("A6:J6").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  sheet.getRange("A7:J8").getFormat().getFill().setColor(VERY_LIGHT);
  sheet.getRange("A7:J7").getFormat().getFont().setBold(true);
  sheet.getRange("A7:J7").getFormat().getFont().setSize(14);
  sheet.getRange("A7:J8").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  sheet.getRange("A8:J8").getFormat().getFont().setColor(MUTED);
  sheet.getRange("A9:J9").getFormat().getFill().setColor(PALE_YELLOW);
  sheet.getRange("A9:J9").getFormat().getFont().setColor(MUTED);
  sheet.getRange("A9:J9").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  sheet.getRange("A9:J9").getFormat().setWrapText(true);
  sheet.getRange("A10:J10").getFormat().getFont().setColor(MUTED);
  sheet.getRange("A10:J10").getFormat().getFont().setItalic(true);
  sheet.getRange("A10:J10").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  sheet.getRange("A13:J13").getFormat().getFill().setColor(LIGHT_BLUE);
  sheet.getRange("A13:J13").getFormat().getFont().setColor(NAVY);
  sheet.getRange("A13:J13").getFormat().getFont().setBold(true);
  sheet.getRange("A13:J13").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  sheet.getRange("A14:J16").getFormat().getFill().setColor(VERY_LIGHT);
  sheet.getRange("A14:J16").getFormat().getFont().setBold(true);
  sheet.getRange("A14:J16").getFormat().getFont().setSize(16);
  sheet.getRange("A14:B16").getFormat().getFont().setSize(19);
  sheet.getRange("A14:J16").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);
  sheet.getRange("A14:B16").setNumberFormat('#,##0 "NOK"');
  sheet.getRange("C14:D16").setNumberFormat('#,##0 "NOK"');
  sheet.getRange("E14:F16").setNumberFormat("0.00%");
  sheet.getRange("G14:H16").setNumberFormat('+0.00 "pp";-0.00 "pp";0.00 "pp"');
  sheet.getRange("I14:J16").setNumberFormat('+#,##0 "NOK";-#,##0 "NOK";0 "NOK"');
  sheet.getRange("A17:J17").getFormat().getFill().setColor(PALE_BLUE);
  sheet.getRange("A17:J17").getFormat().getFont().setColor(MUTED);
  sheet.getRange("A17:J17").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.center);

  formatRankingBlock(sheet, "A20:E27");
  formatRankingBlock(sheet, "F20:J27");
  sheet.getRange("A20:E20").getFormat().getFill().setColor(LIGHT_BLUE);
  sheet.getRange("F20:J20").getFormat().getFill().setColor(LIGHT_BLUE);
  sheet.getRange("A24:E24").getFormat().getFill().setColor(LIGHT_BLUE);
  sheet.getRange("F24:J24").getFormat().getFill().setColor(LIGHT_BLUE);
  sheet.getRange("A20:J20").getFormat().getFont().setBold(true);
  sheet.getRange("A24:J24").getFormat().getFont().setBold(true);
  sheet.getRange("E21:E27").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  sheet.getRange("J21:J27").getFormat().setHorizontalAlignment(ExcelScript.HorizontalAlignment.right);
  sheet.getRange("F19:F27").getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setStyle(ExcelScript.BorderLineStyle.continuous);
  sheet.getRange("F19:F27").getFormat().getRangeBorder(ExcelScript.BorderIndex.edgeLeft).setColor("#AAB7C4");
  addNegativeTextFormat(sheet.getRange("E21:E27"), "E21");
  addNegativeTextFormat(sheet.getRange("J21:J27"), "J21");

  sheet.getRange("A30:E34").getFormat().getFill().setColor(VERY_LIGHT);
  sheet.getRange("A30:C34").getFormat().getFont().setBold(true);
  sheet.getRange("A30:E34").getFormat().setWrapText(true);
  sheet.getRange("D30:E30").setNumberFormat('#,##0 "NOK"');
  sheet.getRange("A36:E36").getFormat().getFont().setColor(MUTED);
  sheet.getRange("A36:E36").getFormat().getFont().setItalic(true);
  sheet.getRange("A36:E36").getFormat().setWrapText(true);
  sheet.getRange("F30:H34").getFormat().getFont().setBold(true);
  sheet.getRange("I30:J34").getFormat().getFill().setColor(VERY_LIGHT);
  sheet.getRange("I30:J34").getFormat().setWrapText(true);
  sheet.getRange("F36:J36").getFormat().getFont().setColor(MUTED);
  sheet.getRange("F36:J36").getFormat().getFont().setItalic(true);

  const negativePp = sheet.getRange("G14").addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
  negativePp.getCellValue().setRule({ formula1: "0", operator: ExcelScript.ConditionalCellValueOperator.lessThan });
  negativePp.getCellValue().getFormat().getFont().setColor(RED);
  const negativeImpact = sheet.getRange("I14").addConditionalFormat(ExcelScript.ConditionalFormatType.cellValue);
  negativeImpact.getCellValue().setRule({ formula1: "0", operator: ExcelScript.ConditionalCellValueOperator.lessThan });
  negativeImpact.getCellValue().getFormat().getFont().setColor(RED);
  const availableStatus = sheet.getRange("I34:J34").addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
  availableStatus.getCustom().getRule().setFormula('=$I$34="Available"');
  availableStatus.getCustom().getFormat().getFill().setColor(GREEN);

  sheet.getRange("1:1").getFormat().setRowHeight(36);
  sheet.getRange("2:2").getFormat().setRowHeight(28);
  sheet.getRange("3:3").getFormat().setRowHeight(26);
  sheet.getRange("4:4").getFormat().setRowHeight(8);
  sheet.getRange("5:6").getFormat().setRowHeight(25);
  sheet.getRange("7:8").getFormat().setRowHeight(25);
  sheet.getRange("9:10").getFormat().setRowHeight(24);
  sheet.getRange("11:11").getFormat().setRowHeight(8);
  sheet.getRange("12:13").getFormat().setRowHeight(25);
  sheet.getRange("14:16").getFormat().setRowHeight(23);
  sheet.getRange("17:17").getFormat().setRowHeight(26);
  sheet.getRange("18:18").getFormat().setRowHeight(8);
  sheet.getRange("19:20").getFormat().setRowHeight(25);
  sheet.getRange("21:27").getFormat().setRowHeight(23);
  sheet.getRange("28:28").getFormat().setRowHeight(8);
  sheet.getRange("29:29").getFormat().setRowHeight(25);
  sheet.getRange("30:34").getFormat().setRowHeight(25);
  sheet.getRange("35:35").getFormat().setRowHeight(5);
  sheet.getRange("36:36").getFormat().setRowHeight(27);

  sheet.getRange("A:A").getFormat().setColumnWidth(118);
  sheet.getRange("B:B").getFormat().setColumnWidth(82);
  sheet.getRange("C:C").getFormat().setColumnWidth(108);
  sheet.getRange("D:D").getFormat().setColumnWidth(92);
  sheet.getRange("E:E").getFormat().setColumnWidth(112);
  sheet.getRange("F:F").getFormat().setColumnWidth(112);
  sheet.getRange("G:G").getFormat().setColumnWidth(82);
  sheet.getRange("H:H").getFormat().setColumnWidth(98);
  sheet.getRange("I:I").getFormat().setColumnWidth(110);
  sheet.getRange("J:J").getFormat().setColumnWidth(112);
  sheet.getFreezePanes().freezeRows(2);
}

function formatSectionHeader(range: ExcelScript.Range): void {
  range.getFormat().getFill().setColor(BLUE);
  range.getFormat().getFont().setColor(WHITE);
  range.getFormat().getFont().setBold(true);
  range.getFormat().getFont().setSize(12);
}

function formatRankingBlock(sheet: ExcelScript.Worksheet, address: string): void {
  sheet.getRange(address).getFormat().getFill().setColor(VERY_LIGHT);
  sheet.getRange(address).getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
}

function addNegativeTextFormat(range: ExcelScript.Range, anchor: string): void {
  const negative = range.addConditionalFormat(ExcelScript.ConditionalFormatType.custom);
  negative.getCustom().getRule().setFormula(`=LEFT(${anchor},1)="-"`);
  negative.getCustom().getFormat().getFont().setColor(RED);
}

function validateManagementAuthority(calc: ExcelScript.Worksheet): void {
  const labels = calc.getRange("AK32:AK45").getTexts().map(row => row[0]).join("|");
  const required = [
    "Management Performance availability", "Current period", "Comparison period", "Period status",
    "Latest published", "Selection recency", "Display metric", "Restaurant scope",
    "Reporting Group scope", "Total Sales", "Selected Category Sales", "Sales Share",
    "PP Change", "NOK Impact"
  ].join("|");
  if (labels !== required || calc.getRange("AL32").getFormula() === "") {
    throw new Error("PUL-030OV-003: Weekly Performance management authority is missing.");
  }
  const groupLabels = calc.getRange("AN32:AQ32").getTexts()[0].join("|");
  const restaurantLabels = calc.getRange("AS32:AV32").getTexts()[0].join("|");
  if (groupLabels !== "Position|ReportingGroupID|Reporting Group|Numeric value" ||
      restaurantLabels !== "Position|RestaurantID|Restaurant|Numeric value") {
    throw new Error("PUL-030OV-004: Performance Top/Bottom ranking authority is missing.");
  }
}

function validateOverviewPostconditions(
  overview: ExcelScript.Worksheet,
  calc: ExcelScript.Worksheet,
  imports: ExcelScript.Worksheet,
  mapping: ExcelScript.Worksheet
): void {
  if (overview.getShowGridlines() || overview.getRange("A1").getText() !== "Pulse" ||
      overview.getRange("A2").getText() !== "Operational performance at a glance." ||
      overview.getRange("A7").getText() !== calc.getRange("AL33").getText() ||
      overview.getRange("D7").getText() !== calc.getRange("AL34").getText() ||
      overview.getRange("G7").getText() !== calc.getRange("AL36").getText() ||
      overview.getRange("I30").getText() !== imports.getRange("A9").getText() ||
      overview.getRange("I33").getText() !== mapping.getRange("A9").getText()) {
    throw new Error("PUL-030OV-005: Overview facade does not match its accepted upstream authorities.");
  }
  const formulas = overview.getRange("A1:J43").getFormulas().map(row => row.join("|")).join("\n");
  if (/tblSalesFacts|tblMetricRPGFacts|tblWeeklyRPGCache|tblWeeklyScopeCache|SUM\(|SUMIFS|AVERAGE\(|SORTBY\(|tblWeeklyPeriodManifest/i.test(formulas)) {
    throw new Error("PUL-030OV-006: Overview contains a prohibited independent business calculation or ranking formula.");
  }
  validateNavigation(overview);
}

function validateNavigation(sheet: ExcelScript.Worksheet): void {
  const performance = sheet.getRange("A3").getHyperlink();
  const reports = sheet.getRange("C3").getHyperlink();
  const imports = sheet.getRange("E3").getHyperlink();
  const mapping = sheet.getRange("G3").getHyperlink();
  const settings = sheet.getRange("I3").getHyperlink();
  if (performance.documentReference !== "Performance!A1" || reports.documentReference !== "Reports!A1" ||
      imports.documentReference !== "Imports!A1" || mapping.documentReference !== "Mapping!A1" ||
      settings.documentReference !== "Settings!A1") {
    throw new Error("PUL-030OV-007: Overview navigation is incomplete.");
  }
}

function validatePrimarySheets(
  workbook: ExcelScript.Workbook,
  overview: ExcelScript.Worksheet,
  performance: ExcelScript.Worksheet,
  reports: ExcelScript.Worksheet,
  imports: ExcelScript.Worksheet,
  mapping: ExcelScript.Worksheet,
  settings: ExcelScript.Worksheet
): void {
  const visibleSheetCount = workbook.getWorksheets()
    .filter(sheet => sheet.getVisibility() === ExcelScript.SheetVisibility.visible).length;
  if (visibleSheetCount !== 6 ||
      overview.getVisibility() !== ExcelScript.SheetVisibility.visible || overview.getPosition() !== 0 ||
      performance.getVisibility() !== ExcelScript.SheetVisibility.visible || performance.getPosition() !== 1 ||
      reports.getVisibility() !== ExcelScript.SheetVisibility.visible || reports.getPosition() !== 2 ||
      imports.getVisibility() !== ExcelScript.SheetVisibility.visible || imports.getPosition() !== 3 ||
      mapping.getVisibility() !== ExcelScript.SheetVisibility.visible || mapping.getPosition() !== 4 ||
      settings.getVisibility() !== ExcelScript.SheetVisibility.visible || settings.getPosition() !== 5) {
    throw new Error("PUL-030OV-008: Accepted six-sheet primary architecture is not present.");
  }
}

function validateQa(table: ExcelScript.Table, pattern: RegExp, expected: number, label: string): void {
  const headers = headerMap(table);
  const rows = table.getRangeBetweenHeaderAndTotal().getValues();
  let pass = 0;
  for (const row of rows) {
    if (pattern.test(text(row[headers.CheckID])) && text(row[headers.Result]) === "PASS") pass += 1;
  }
  if (pass !== expected) throw new Error(`PUL-030OV-009: ${label} QA is ${pass}/${expected} PASS.`);
}

function protectedState(
  workbook: ExcelScript.Workbook,
  performance: ExcelScript.Worksheet,
  reports: ExcelScript.Worksheet,
  imports: ExcelScript.Worksheet,
  mapping: ExcelScript.Worksheet,
  calc: ExcelScript.Worksheet
): string {
  return JSON.stringify([
    performance.getRange("B6:I19").getValues(), performance.getRange("B6:I19").getFormulas(),
    reports.getRange("B7:B18").getValues(), reports.getRange("B7:B18").getFormulas(),
    imports.getRange("A9:G9").getValues(), imports.getRange("A9:G9").getFormulas(),
    mapping.getRange("A4:J9").getValues(), mapping.getRange("A4:J9").getFormulas(),
    calc.getRange("AK1:AV45").getValues(), calc.getRange("AK1:AV45").getFormulas(),
    requiredTable(workbook, "tblPerformanceRestaurantSelection").getRange().getValues(),
    requiredTable(workbook, "tblPerformanceRPGSelection").getRange().getValues(),
    requiredTable(workbook, "tblWeeklyCacheVersions").getRange().getValues(),
    requiredTable(workbook, "tblMappingRules").getRange().getValues(),
    requiredTable(workbook, "tblEffectiveMapping").getRange().getValues()
  ]);
}

function setInternalLink(target: ExcelScript.Range, sheetName: string, screenTip: string): void {
  target.setHyperlink({
    documentReference: `${sheetName}!A1`,
    screenTip,
    textToDisplay: sheetName
  });
}

function requiredSheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  const sheet = workbook.getWorksheet(name);
  if (!sheet) throw new Error(`PUL-030OV-010: Required worksheet ${name} is missing.`);
  return sheet;
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-030OV-011: Required table ${name} is missing.`);
  return table;
}

function headerMap(table: ExcelScript.Table): { [key: string]: number } {
  const headers = table.getHeaderRowRange().getTexts()[0];
  const result: { [key: string]: number } = {};
  for (let index = 0; index < headers.length; index += 1) result[text(headers[index])] = index;
  return result;
}

function text(value: unknown): string {
  return String(value === null || value === undefined ? "" : value).trim();
}
