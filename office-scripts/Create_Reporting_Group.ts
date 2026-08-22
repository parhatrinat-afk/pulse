/**
 * Pulse 0.3.0 — Create Reporting Group.
 *
 * User entry point: Excel for the web -> Automate -> Create Reporting Group.
 * tblReportingGroups remains the sole authority. Creation adds no Mapping Rule
 * and deliberately makes weekly Performance unavailable until a complete cache
 * version is rebuilt and activated for the new active-group catalog.
 */
function main(workbook: ExcelScript.Workbook): string {
  const settings = requiredSheet(workbook, "Settings");
  const groupsTable = requiredTable(workbook, "tblReportingGroups");
  const initialRows = tableRows(groupsTable);
  const initialGroups = readAuthorityRows(groupsTable, initialRows);
  ensureSettingsSurface(workbook, settings, initialGroups);

  const requestedName = text(settings.getRange("G17").getValue());
  if (!requestedName) {
    writeSettingsStatus(settings, "Ready — enter a business name, then run Create Reporting Group.", false);
    return JSON.stringify({ status: "Ready", created: false, activeReportingGroups: activeGroups(initialGroups).length });
  }

  const description = text(settings.getRange("G18").getValue());
  const notes = text(settings.getRange("G20").getValue());
  const authority = validateAuthority(initialGroups);
  const plan = planCreation(authority, requestedName, description, notes);
  const protectedBefore = captureProtectedState(workbook);
  validateDownstreamSurfaces(workbook);
  const priorGroupCount = initialRows.length;
  const priorRuleCount = tableRows(requiredTable(workbook, "tblMappingRules")).length;
  const priorEffectiveCount = tableRows(requiredTable(workbook, "tblEffectiveMapping")).length;
  const priorIncludeById = capturePriorPerformanceSelections(workbook);
  let appended = false;

  try {
    groupsTable.addRow(-1, [[
      plan.id, plan.name, plan.domainId, "Yes", plan.sortOrder, plan.description, plan.notes
    ]]);
    appended = true;
    const currentRows = tableRows(groupsTable);
    const currentGroups = readAuthorityRows(groupsTable, currentRows);
    validateAuthority(currentGroups);
    const currentActive = activeGroups(currentGroups);
    const fingerprints = computeLiveFingerprints(workbook);

    refreshSettingsOverview(workbook, settings, currentGroups);
    refreshMappingSurfaces(workbook, currentActive);
    refreshPerformanceSurfaces(workbook, currentActive, priorIncludeById, fingerprints);
    workbook.getApplication().calculate(ExcelScript.CalculationType.full);

    validateCreatedState(
      workbook, plan, priorGroupCount, priorRuleCount, priorEffectiveCount,
      priorIncludeById, protectedBefore
    );
    settings.getRange("G17:K17").clear(ExcelScript.ClearApplyTo.contents);
    settings.getRange("G18:K19").clear(ExcelScript.ClearApplyTo.contents);
    settings.getRange("G20:K21").clear(ExcelScript.ClearApplyTo.contents);
    writeSettingsStatus(settings,
      `Created ${plan.id} — ${plan.name}. Performance refresh required.`, false);
    return JSON.stringify({
      status: "Created",
      created: true,
      reportingGroupId: plan.id,
      reportingGroupName: plan.name,
      activeReportingGroups: currentActive.length,
      mappingRulesCreated: 0,
      performance: "Refresh required"
    });
  } catch (error) {
    if (appended) {
      try {
        groupsTable.deleteRowsAt(priorGroupCount, 1);
        const restoredRows = tableRows(groupsTable);
        const restoredGroups = readAuthorityRows(groupsTable, restoredRows);
        const restoredActive = activeGroups(restoredGroups);
        const restoredFingerprints = computeLiveFingerprints(workbook);
        refreshSettingsOverview(workbook, settings, restoredGroups);
        refreshMappingSurfaces(workbook, restoredActive);
        refreshPerformanceSurfaces(workbook, restoredActive, priorIncludeById, restoredFingerprints);
        workbook.getApplication().calculate(ExcelScript.CalculationType.full);
      } catch (rollbackError) {
        writeSettingsStatus(settings,
          `Creation failed and rollback needs review: ${errorText(rollbackError)}`, true);
        throw new Error(`PUL-030RG-099: Creation failed (${errorText(error)}); rollback failed (${errorText(rollbackError)}).`);
      }
    }
    writeSettingsStatus(settings, `Not created — ${errorText(error)}`, true);
    throw error;
  }
}

type CellValue = string | number | boolean;
type AuthorityRow = {
  id: string; name: string; domainId: string; active: string; sortOrder: number;
  description: string; notes: string;
};
type AuthorityPlan = {
  rows: AuthorityRow[]; id: string; sortOrder: number; domainId: string;
  activeNameKeys: { [key: string]: boolean };
};
type CreationPlan = {
  id: string; name: string; domainId: string; sortOrder: number;
  description: string; notes: string;
};
type ActiveGroup = { id: string; name: string; sortOrder: number };
type LiveFingerprints = { mappingContent: string; catalogContent: string };

function readAuthorityRows(table: ExcelScript.Table, values: CellValue[][]): AuthorityRow[] {
  const h = headerMap(table);
  const required = ["ReportingGroupID", "ReportingGroupName", "DomainID", "Active",
    "SortOrder", "Description", "Notes"];
  for (let index = 0; index < required.length; index += 1) {
    if (h[required[index]] === undefined) {
      throw new Error(`PUL-030RG-001: tblReportingGroups is missing ${required[index]}.`);
    }
  }
  const rows: AuthorityRow[] = [];
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index];
    rows.push({
      id: text(value[h.ReportingGroupID]),
      name: text(value[h.ReportingGroupName]),
      domainId: text(value[h.DomainID]),
      active: text(value[h.Active]),
      sortOrder: number(value[h.SortOrder]),
      description: text(value[h.Description]),
      notes: text(value[h.Notes])
    });
  }
  return rows;
}

function validateAuthority(rows: AuthorityRow[]): AuthorityPlan {
  if (rows.length < 1) throw new Error("PUL-030RG-002: Reporting Group authority is empty.");
  const ids: { [key: string]: boolean } = {};
  const sortOrders: { [key: string]: boolean } = {};
  const activeNameKeys: { [key: string]: boolean } = {};
  let highestId = 0;
  let idWidth = 4;
  let highestSortOrder = 0;
  let activeCount = 0;
  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];
    const match = /^RPG-(\d+)$/.exec(row.id);
    if (!match || Number(match[1]) < 1) {
      throw new Error(`PUL-030RG-002: Invalid stable ReportingGroupID ${row.id || "(blank)"}.`);
    }
    if (ids[row.id]) throw new Error(`PUL-030RG-002: Duplicate ReportingGroupID ${row.id}.`);
    ids[row.id] = true;
    highestId = Math.max(highestId, Number(match[1]));
    idWidth = Math.max(idWidth, match[1].length);
    if (!row.name) throw new Error(`PUL-030RG-002: Reporting Group ${row.id} has a blank name.`);
    if (!Number.isInteger(row.sortOrder) || row.sortOrder < 1) {
      throw new Error(`PUL-030RG-002: Reporting Group ${row.id} has invalid SortOrder.`);
    }
    const sortKey = String(row.sortOrder);
    if (sortOrders[sortKey]) throw new Error(`PUL-030RG-002: Duplicate SortOrder ${sortKey}.`);
    sortOrders[sortKey] = true;
    highestSortOrder = Math.max(highestSortOrder, row.sortOrder);
    if (row.active === "Yes") {
      activeCount += 1;
      const nameKey = businessNameKey(row.name);
      if (activeNameKeys[nameKey]) {
        throw new Error(`PUL-030RG-002: Duplicate active business name ${row.name}.`);
      }
      activeNameKeys[nameKey] = true;
      if (row.domainId !== "DOM-SALES") {
        throw new Error(`PUL-030RG-002: Active Reporting Group ${row.id} is outside DOM-SALES.`);
      }
    } else if (row.active !== "No") {
      throw new Error(`PUL-030RG-002: Reporting Group ${row.id} has invalid Active value.`);
    }
  }
  if (activeCount < 1) throw new Error("PUL-030RG-002: At least one active Reporting Group is required.");
  const nextNumber = highestId + 1;
  const nextId = `RPG-${String(nextNumber).padStart(idWidth, "0")}`;
  if (ids[nextId]) throw new Error(`PUL-030RG-002: Next stable ID ${nextId} already exists.`);
  let nextSort = (Math.floor(highestSortOrder / 10) + 1) * 10;
  while (sortOrders[String(nextSort)]) nextSort += 10;
  return { rows, id: nextId, sortOrder: nextSort, domainId: "DOM-SALES", activeNameKeys };
}

function planCreation(
  authority: AuthorityPlan, nameValue: string, description: string, notes: string
): CreationPlan {
  const name = text(nameValue);
  if (!name) throw new Error("PUL-030RG-003: Reporting Group name is required.");
  if (authority.activeNameKeys[businessNameKey(name)]) {
    throw new Error(`PUL-030RG-004: Active Reporting Group ${name} already exists.`);
  }
  return {
    id: authority.id,
    name,
    domainId: authority.domainId,
    sortOrder: authority.sortOrder,
    description: text(description),
    notes: text(notes)
  };
}

function activeGroups(rows: AuthorityRow[]): ActiveGroup[] {
  const groups: ActiveGroup[] = [];
  for (let index = 0; index < rows.length; index += 1) {
    const row = rows[index];
    if (row.active === "Yes") groups.push({ id: row.id, name: row.name, sortOrder: row.sortOrder });
  }
  groups.sort((left, right) => left.sortOrder - right.sortOrder ||
    (left.id < right.id ? -1 : left.id > right.id ? 1 : 0));
  if (groups.length < 1) throw new Error("PUL-030RG-002: No active Reporting Groups are available.");
  return groups;
}

function ensureSettingsSurface(
  workbook: ExcelScript.Workbook, settings: ExcelScript.Worksheet, groups: AuthorityRow[]
): void {
  refreshSettingsOverview(workbook, settings, groups);
  const installed = text(settings.getRange("F15").getValue()) === "New Reporting Group";
  if (!installed) {
    settings.getRange("F15:K24").unmerge();
    settings.getRange("F15:K24").clear(ExcelScript.ClearApplyTo.all);
    settings.getRange("F15:K15").merge();
    settings.getRange("F15").setValue("New Reporting Group");
    settings.getRange("F16:K16").merge();
    settings.getRange("F16").setValue("Create one active business classification. Pulse assigns the stable ID.");
    settings.getRange("F17").setValue("Name");
    settings.getRange("G17:K17").merge();
    settings.getRange("F18").setValue("Description");
    settings.getRange("G18:K19").merge();
    settings.getRange("F20").setValue("Notes");
    settings.getRange("G20:K21").merge();
    settings.getRange("F22:K22").merge();
    settings.getRange("F22").setValue("Run Create Reporting Group from Automate to validate and create.");
    settings.getRange("F23").setValue("Status");
    settings.getRange("G23:K24").merge();
  }
  const navy = "#17365D";
  const paleBlue = "#D9EAF7";
  const input = "#FFF4CE";
  settings.getRange("F15:K15").getFormat().getFill().setColor(navy);
  settings.getRange("F15:K15").getFormat().getFont().setColor("#FFFFFF");
  settings.getRange("F15:K15").getFormat().getFont().setBold(true);
  settings.getRange("F16:K16").getFormat().getFill().setColor(paleBlue);
  settings.getRange("G17:K21").getFormat().getFill().setColor(input);
  settings.getRange("F17:F23").getFormat().getFont().setBold(true);
  settings.getRange("F15:K24").getFormat().setVerticalAlignment(ExcelScript.VerticalAlignment.center);
  settings.getRange("F16:K24").getFormat().setWrapText(true);
  settings.getRange("15:17").getFormat().setRowHeight(26);
  settings.getRange("18:21").getFormat().setRowHeight(22);
  settings.getRange("22:24").getFormat().setRowHeight(26);
  settings.getRange("F:F").getFormat().setColumnWidth(96);
  settings.getRange("G:K").getFormat().setColumnWidth(82);
  settings.getFreezePanes().freezeRows(3);
}

function refreshSettingsOverview(
  workbook: ExcelScript.Workbook, settings: ExcelScript.Worksheet, groups: AuthorityRow[]
): void {
  const impacts = reportingGroupImpacts(workbook);
  const sorted = groups.slice();
  sorted.sort((left, right) => left.sortOrder - right.sortOrder ||
    (left.id < right.id ? -1 : left.id > right.id ? 1 : 0));
  settings.getRange("A15:D15").unmerge();
  settings.getRange("A15:D15").merge();
  settings.getRange("A15").setValue("Reporting Groups");
  settings.getRange("A15:D15").getFormat().getFill().setColor("#17365D");
  settings.getRange("A15:D15").getFormat().getFont().setColor("#FFFFFF");
  settings.getRange("A15:D15").getFormat().getFont().setBold(true);
  const rows: CellValue[][] = [];
  for (let index = 0; index < sorted.length; index += 1) {
    const group = sorted[index];
    const impact = impacts[group.id] || { products: 0, facts: 0, sales: 0 };
    rows.push([group.name, group.active === "Yes" ? "Active" : "Inactive", impact.products, impact.sales]);
  }
  let table = workbook.getTable("tblSettingsReportingGroups");
  const target = settings.getRangeByIndexes(15, 0, rows.length + 1, 4);
  if (!table) {
    settings.getRange("A16:D16").setValues([[
      "Reporting Group", "Status", "Products", "Sales NOK"
    ]]);
    table = settings.addTable(target, true);
    table.setName("tblSettingsReportingGroups");
  } else {
    table.resize(target);
    table.getHeaderRowRange().setValues([[
      "Reporting Group", "Status", "Products", "Sales NOK"
    ]]);
  }
  if (rows.length) table.getRangeBetweenHeaderAndTotal().setValues(rows);
  table.setPredefinedTableStyle("TableStyleMedium2");
  const products = table.getColumnByName("Products");
  const sales = table.getColumnByName("Sales NOK");
  if (!products || !sales) throw new Error("PUL-030RG-005: Settings overview columns are missing.");
  products.getRangeBetweenHeaderAndTotal().setNumberFormat("#,##0");
  sales.getRangeBetweenHeaderAndTotal().setNumberFormat("#,##0.00");
  settings.getRange("A:A").getFormat().setColumnWidth(142);
  settings.getRange("B:B").getFormat().setColumnWidth(72);
  settings.getRange("C:C").getFormat().setColumnWidth(72);
  settings.getRange("D:D").getFormat().setColumnWidth(104);
}

function reportingGroupImpacts(workbook: ExcelScript.Workbook): { [key: string]: { products: number; facts: number; sales: number } } {
  const table = requiredTable(workbook, "tblMappingReportingGroupCatalog");
  const h = headerMap(table);
  const rows = tableRows(table);
  const result: { [key: string]: { products: number; facts: number; sales: number } } = {};
  for (let index = 0; index < rows.length; index += 1) {
    const reportingGroupId = text(rows[index][h.ReportingGroupID]);
    if (!reportingGroupId) continue;
    const current = result[reportingGroupId] || { products: 0, facts: 0, sales: 0 };
    current.products += 1;
    current.facts += number(rows[index][h.Facts]);
    current.sales += number(rows[index][h["Sales NOK"]]);
    result[reportingGroupId] = current;
  }
  return result;
}

function refreshMappingSurfaces(workbook: ExcelScript.Workbook, groups: ActiveGroup[]): void {
  const mapping = requiredSheet(workbook, "Mapping");
  const lists = requiredSheet(workbook, "_Mapping_Lists");
  const impacts = reportingGroupImpacts(workbook);
  const values: CellValue[][] = [];
  for (let index = 0; index < groups.length; index += 1) {
    const group = groups[index];
    const impact = impacts[group.id] || { products: 0, facts: 0, sales: 0 };
    values.push([group.name, impact.products, impact.facts, impact.sales]);
  }
  const overview = requiredTable(workbook, "tblMappingReportingGroupOverview");
  const overviewRange = overview.getRange();
  overview.resize(mapping.getRangeByIndexes(
    overviewRange.getRowIndex(), overviewRange.getColumnIndex(), groups.length + 1, 4
  ));
  overview.getHeaderRowRange().setValues([["Reporting Group", "Products", "Facts", "Sales NOK"]]);
  overview.getRangeBetweenHeaderAndTotal().setValues(values);
  const productsColumn = overview.getColumnByName("Products");
  const factsColumn = overview.getColumnByName("Facts");
  const salesColumn = overview.getColumnByName("Sales NOK");
  if (!productsColumn || !factsColumn || !salesColumn) {
    throw new Error("PUL-030RG-005: Mapping overview columns are missing.");
  }
  productsColumn.getRangeBetweenHeaderAndTotal().setNumberFormat("#,##0");
  factsColumn.getRangeBetweenHeaderAndTotal().setNumberFormat("#,##0");
  salesColumn.getRangeBetweenHeaderAndTotal().setNumberFormat("#,##0.00");

  mapping.getRange("AS2:AS200").clear(ExcelScript.ClearApplyTo.contents);
  mapping.getRange("AN2:AN200").clear(ExcelScript.ClearApplyTo.contents);
  lists.getRange("A2:A1000").clear(ExcelScript.ClearApplyTo.contents);
  const names: CellValue[][] = [];
  const ids: CellValue[][] = [];
  const formulas: string[][] = [];
  for (let index = 0; index < groups.length; index += 1) {
    names.push([groups[index].name]);
    ids.push([groups[index].id]);
    formulas.push([`=IF(AND($AL$8,$B$15="Assign Reporting Group"),$AS$${index + 2},"")`]);
  }
  mapping.getRangeByIndexes(1, 44, groups.length, 1).setValues(names);
  mapping.getRangeByIndexes(1, 39, groups.length, 1).setFormulas(formulas);
  lists.getRangeByIndexes(1, 0, groups.length, 1).setValues(ids);
  applyRangeValidation(mapping.getRange("E10:F10"), mapping.getRangeByIndexes(1, 44, groups.length, 1));
  applyRangeValidation(mapping.getRange("E15:F15"), mapping.getRangeByIndexes(1, 39, groups.length, 1));
  mapping.getRange("A9").setValue("Performance refresh required");
  mapping.getRange("A9:K9").getFormat().getFill().setColor("#FFF4CE");
  mapping.getRange("A9:K9").getFormat().getFont().setBold(true);
}

function capturePriorPerformanceSelections(workbook: ExcelScript.Workbook): { [key: string]: string } {
  const table = requiredTable(workbook, "tblPerformanceRPGSelection");
  const h = headerMap(table);
  const rows = tableRows(table);
  const result: { [key: string]: string } = {};
  for (let index = 0; index < rows.length; index += 1) {
    const id = text(rows[index][h.ReportingGroupID]);
    if (!id || result[id] !== undefined) {
      throw new Error(`PUL-030RG-006: Performance Reporting Group selection repeats or omits ${id || "an ID"}.`);
    }
    result[id] = text(rows[index][h.Include]) === "Yes" ? "Yes" : "No";
  }
  return result;
}

function refreshPerformanceSurfaces(
  workbook: ExcelScript.Workbook,
  groups: ActiveGroup[],
  priorIncludeById: { [key: string]: string },
  fingerprints: LiveFingerprints
): void {
  const performance = requiredSheet(workbook, "Performance");
  const calc = requiredSheet(workbook, "_Metric_Calc");
  const selection = requiredTable(workbook, "tblPerformanceRPGSelection");
  const selectionRange = selection.getRange();
  selection.resize(performance.getRangeByIndexes(
    selectionRange.getRowIndex(), selectionRange.getColumnIndex(), groups.length + 1, 3
  ));
  selection.getHeaderRowRange().setValues([["Include", "Reporting Group", "ReportingGroupID"]]);
  const selectionValues: CellValue[][] = [];
  for (let index = 0; index < groups.length; index += 1) {
    const group = groups[index];
    selectionValues.push([
      priorIncludeById[group.id] !== undefined ? priorIncludeById[group.id] : "No",
      group.name,
      group.id
    ]);
  }
  selection.getRangeBetweenHeaderAndTotal().setValues(selectionValues);
  const includeColumn = selection.getColumnByName("Include");
  if (!includeColumn) throw new Error("PUL-030RG-006: Performance Include column is missing.");
  const includeValidation = includeColumn.getRangeBetweenHeaderAndTotal().getDataValidation();
  includeValidation.clear();
  includeValidation.setRule({ list: { inCellDropDown: true, source: "Yes,No" } });
  includeValidation.setErrorAlert({ showAlert: true, style: ExcelScript.DataValidationAlertStyle.stop,
    title: "Choose Yes or No", message: "New Reporting Groups default to No until selected." });

  calc.getRange("I2:J200").clear(ExcelScript.ClearApplyTo.contents);
  const groupValues: CellValue[][] = [];
  for (let index = 0; index < groups.length; index += 1) groupValues.push([groups[index].name, groups[index].id]);
  calc.getRangeByIndexes(1, 8, groups.length, 2).setValues(groupValues);
  calc.getRange("AH2:AI200").clear(ExcelScript.ClearApplyTo.contents);
  const selectedFormulas: string[][] = [];
  for (let index = 0; index < groups.length; index += 1) {
    const row = index + 2;
    selectedFormulas.push([
      `=IFERROR(INDEX(FILTER(tblPerformanceRPGSelection[Reporting Group],tblPerformanceRPGSelection[Include]="Yes"),ROWS($AH$2:AH${row})),"")`,
      `=IFERROR(INDEX(FILTER(tblPerformanceRPGSelection[ReportingGroupID],tblPerformanceRPGSelection[Include]="Yes"),ROWS($AI$2:AI${row})),"")`
    ]);
  }
  calc.getRangeByIndexes(1, 33, groups.length, 2).setFormulas(selectedFormulas);

  const helperStart = 21;
  performance.getRangeByIndexes(1, helperStart + 2, 100, 2).clear(ExcelScript.ClearApplyTo.contents);
  performance.getRangeByIndexes(1, helperStart + 2, groups.length, 1)
    .setValues(groups.map(group => [group.name]));
  performance.getRangeByIndexes(1, helperStart + 3, groups.length + 1, 1)
    .setValues([["Total"]].concat(groups.map(group => [group.name])));
  const currentDetail = text(performance.getRange("B7").getValue());
  const currentSort = text(performance.getRange("I6").getValue());
  if (!containsGroupName(groups, currentDetail)) performance.getRange("B7").setValue(groups[0].name);
  if (currentSort !== "Total" && !containsGroupName(groups, currentSort)) performance.getRange("I6").setValue("Total");
  applyRangeValidation(performance.getRange("B7"),
    performance.getRangeByIndexes(1, helperStart + 2, groups.length, 1));
  applyRangeValidation(performance.getRange("I6"),
    performance.getRangeByIndexes(1, helperStart + 3, groups.length + 1, 1));

  const groupEnd = groups.length + 1;
  calc.getRange("AL5").setFormula('=COUNTIF(tblPerformanceRPGSelection[Include],"Yes")');
  calc.getRange("AL6").setFormula(`=IFERROR(INDEX($J$2:$J$${groupEnd},MATCH(Performance!$B$7,$I$2:$I$${groupEnd},0)),"")`);
  calc.getRange("AL11").setValue(groups.length);
  calc.getRange("AL12").setFormula(`=IF(Performance!$I$6="Total","",IFERROR(INDEX($J$2:$J$${groupEnd},MATCH(Performance!$I$6,$I$2:$I$${groupEnd},0)),""))`);
  calc.getRange("AL13").setFormula(`=IF(OR($AL$12="",COUNTIF($AI$2:$AI$${groupEnd},$AL$12)=0),"",$AL$12)`);
  calc.getRange("AL27").setValue(fingerprints.mappingContent);
  calc.getRange("AL28").setValue(fingerprints.catalogContent);
  performance.getRange("G8").setFormula(
    `=IF('_Metric_Calc'!$AL$5=${groups.length},"All ${groups.length} Reporting Groups selected",` +
      `'_Metric_Calc'!$AL$5&" of ${groups.length} Reporting Groups selected")`
  );
}

function containsGroupName(groups: ActiveGroup[], name: string): boolean {
  for (let index = 0; index < groups.length; index += 1) if (groups[index].name === name) return true;
  return false;
}

function computeLiveFingerprints(workbook: ExcelScript.Workbook): LiveFingerprints {
  const groupsTable = requiredTable(workbook, "tblReportingGroups");
  const rulesTable = requiredTable(workbook, "tblMappingRules");
  const productsTable = requiredTable(workbook, "tblProducts");
  const classificationsTable = requiredTable(workbook, "tblSourceClassifications");
  const effectiveTable = requiredTable(workbook, "tblEffectiveMapping");
  const restaurantsTable = requiredTable(workbook, "tblRestaurants");
  const groups = tableRows(groupsTable); const rules = tableRows(rulesTable);
  const products = tableRows(productsTable); const classifications = tableRows(classificationsTable);
  const effective = tableRows(effectiveTable); const restaurants = tableRows(restaurantsTable);
  const mappingContent = computeMappingContentFingerprint(
    groupsTable, groups, rulesTable, rules, productsTable, products,
    classificationsTable, classifications, effectiveTable, effective
  );
  return {
    mappingContent,
    catalogContent: computeCatalogContentFingerprint(
      mappingContent, groupsTable, groups, rulesTable, rules, productsTable, products,
      classificationsTable, classifications, restaurantsTable, restaurants
    )
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
  for (let index = 0; index < classifications.length; index += 1) {
    const row = classifications[index];
    classificationById[text(row[ch.SourceClassificationID])] = {
      sourceSystemId: text(row[ch.SourceSystemID]), main: text(row[ch.SourceMainCategory])
    };
  }
  const records: string[] = [record("V", ["PULSE-MAPPING-CONTENT-V1"])];
  for (let index = 0; index < groups.length; index += 1) {
    const row = groups[index];
    records.push(record("G", [row[gh.ReportingGroupID], row[gh.ReportingGroupName], row[gh.Active], row[gh.SortOrder]]));
  }
  for (let index = 0; index < rules.length; index += 1) {
    const row = rules[index];
    if (!text(row[rh.MappingRuleID])) continue;
    records.push(record("R", [row[rh.MappingRuleID], row[rh.SourceSystemID], row[rh.ScopeType], row[rh.NodeID],
      rh.RuleAction === undefined ? "Map" : text(row[rh.RuleAction]) || "Map", row[rh.TargetReportingGroupID],
      boundary(row[rh.EffectiveFrom]), boundary(row[rh.EffectiveTo]), row[rh.Status]]));
  }
  for (let index = 0; index < products.length; index += 1) {
    const row = products[index];
    const productId = text(row[ph.ProductID]);
    if (!productId) continue;
    const classificationId = text(row[ph.SourceClassificationID]);
    const classification = classificationById[classificationId];
    if (!classification) throw new Error(`PUL-030RG-007: Product ${productId} has missing classification ${classificationId}.`);
    records.push(record("P", [productId, row[ph.SourceSystemID],
      `${classification.sourceSystemId} || Main || ${classification.main}`, classificationId]));
  }
  const effectiveProducts: { [key: string]: boolean } = {};
  for (let index = 0; index < effective.length; index += 1) {
    const row = effective[index];
    const productId = text(row[eh.ProductID]);
    if (!productId || effectiveProducts[productId]) {
      throw new Error(`PUL-030RG-007: Effective Mapping repeats or omits ${productId}.`);
    }
    effectiveProducts[productId] = true;
    records.push(record("E", [productId, row[eh.EffectiveReportingGroupID], row[eh.ResolutionSource],
      row[eh.ResolutionState], row[eh.ResolutionStatus], normalizeDelimited(row[eh.WinningRuleID])]));
  }
  records.sort();
  return hashStrings(records, "MCF-");
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
  const ph = headerMap(productsTable); const ch = headerMap(classificationsTable);
  const th = headerMap(restaurantsTable);
  const records: string[] = [record("CATALOG_CONTENT", [mappingContentFingerprint])];
  for (let index = 0; index < restaurants.length; index += 1) {
    const row = restaurants[index];
    records.push(record("RESTAURANT", [row[th.RestaurantID], row[th.SourceSystemID],
      row[th.SourceRestaurantName], row[th.Status], row[th.ReportingEnabled]]));
  }
  for (let index = 0; index < products.length; index += 1) {
    const row = products[index];
    records.push(record("PRODUCT", [row[ph.ProductID],
      `${text(row[ph.SourceSystemID])} || ${text(row[ph.SourceProductName])} || ${text(row[ph.SalesAccount])}`,
      row[ph.SourceClassificationID], row[ph.ProductStatus]]));
  }
  for (let index = 0; index < classifications.length; index += 1) {
    const row = classifications[index];
    records.push(record("CLASSIFICATION", [row[ch.SourceClassificationID],
      `${text(row[ch.SourceSystemID])} || ${text(row[ch.SourceMainCategory])} || ${text(row[ch.SourceSubCategory])}`,
      row[ch.Status]]));
  }
  for (let index = 0; index < groups.length; index += 1) {
    const row = groups[index];
    records.push(record("REPORTING_GROUP", [row[gh.ReportingGroupID], row[gh.ReportingGroupName],
      row[gh.Active], row[gh.SortOrder]]));
  }
  for (let index = 0; index < rules.length; index += 1) {
    const row = rules[index];
    records.push(record("MAPPING_RULE", [row[rh.MappingRuleID], row[rh.SourceSystemID], row[rh.ScopeType],
      row[rh.NodeID], row[rh.TargetReportingGroupID], boundary(row[rh.EffectiveFrom]),
      boundary(row[rh.EffectiveTo]), row[rh.Status],
      rh.RuleAction === undefined ? "Map" : text(row[rh.RuleAction]) || "Map"]));
  }
  records.sort();
  return hashStrings(records, "ICC-");
}

function validateDownstreamSurfaces(workbook: ExcelScript.Workbook): void {
  requiredTable(workbook, "tblReportingGroups");
  requiredTable(workbook, "tblMappingRules");
  requiredTable(workbook, "tblEffectiveMapping");
  requiredTable(workbook, "tblMappingReportingGroupOverview");
  requiredTable(workbook, "tblMappingReportingGroupCatalog");
  requiredTable(workbook, "tblPerformanceRPGSelection");
  requiredTable(workbook, "tblWeeklyCacheVersions");
  requiredTable(workbook, "tblWeeklyPeriodManifest");
  requiredTable(workbook, "tblWeeklyScopeCache");
  requiredTable(workbook, "tblWeeklyRPGCache");
  requiredTable(workbook, "tblMetricRPGResults");
  requiredTable(workbook, "tblSalesFacts");
  requiredTable(workbook, "tblImports");
  requiredTable(workbook, "tblPerformanceInteractionQA");
  requiredTable(workbook, "tblWeeklyPerformanceQA");
  requiredTable(workbook, "tblMappingQA");
  if (passCount(requiredTable(workbook, "tblMappingQA"), /^QA-0301-/, 9) !== 9 ||
      passCount(requiredTable(workbook, "tblPerformanceInteractionQA"), /^QA-0302C-/, 16) !== 16 ||
      passCount(requiredTable(workbook, "tblWeeklyPerformanceQA"), /^QA-030WP-/, 16) !== 16) {
    throw new Error("PUL-030RG-008: Accepted Mapping, Phase 2C, or Weekly Performance QA is incomplete.");
  }
  const versions = requiredTable(workbook, "tblWeeklyCacheVersions");
  const h = headerMap(versions);
  const rows = tableRows(versions);
  let activeCount = 0;
  for (let index = 0; index < rows.length; index += 1) {
    if (text(rows[index][h.CacheStatus]) === "Active" && text(rows[index][h.ActivationState]) === "Active") {
      activeCount += 1;
    }
  }
  if (activeCount !== 1) throw new Error(`PUL-030RG-008: Expected one Active / Active cache; found ${activeCount}.`);
}

function passCount(table: ExcelScript.Table, idPattern: RegExp, expected: number): number {
  const h = headerMap(table);
  const rows = tableRows(table);
  let count = 0;
  for (let index = 0; index < rows.length; index += 1) {
    const id = text(rows[index][h.CheckID]);
    if (idPattern.test(id) && text(rows[index][h.Result]) === "PASS") count += 1;
  }
  return count;
}

function validateCreatedState(
  workbook: ExcelScript.Workbook,
  plan: CreationPlan,
  priorGroupCount: number,
  priorRuleCount: number,
  priorEffectiveCount: number,
  priorIncludeById: { [key: string]: string },
  protectedBefore: string
): void {
  const groupsTable = requiredTable(workbook, "tblReportingGroups");
  const groups = readAuthorityRows(groupsTable, tableRows(groupsTable));
  if (groups.length !== priorGroupCount + 1) {
    throw new Error("PUL-030RG-009: Creation did not append exactly one authority row.");
  }
  let createdCount = 0;
  for (let index = 0; index < groups.length; index += 1) {
    if (groups[index].id === plan.id && groups[index].name === plan.name && groups[index].active === "Yes") {
      createdCount += 1;
    }
  }
  if (createdCount !== 1) throw new Error(`PUL-030RG-009: ${plan.id} was not created exactly once.`);
  if (tableRows(requiredTable(workbook, "tblMappingRules")).length !== priorRuleCount ||
      tableRows(requiredTable(workbook, "tblEffectiveMapping")).length !== priorEffectiveCount) {
    throw new Error("PUL-030RG-009: Mapping Rules or Effective Mapping changed during group creation.");
  }
  const selection = requiredTable(workbook, "tblPerformanceRPGSelection");
  const sh = headerMap(selection);
  const selectionRows = tableRows(selection);
  if (selectionRows.length !== activeGroups(groups).length) {
    throw new Error("PUL-030RG-009: Performance selection capacity differs from active groups.");
  }
  for (let index = 0; index < selectionRows.length; index += 1) {
    const id = text(selectionRows[index][sh.ReportingGroupID]);
    const include = text(selectionRows[index][sh.Include]);
    if (id === plan.id && include !== "No") {
      throw new Error("PUL-030RG-009: Newly created group did not default to No.");
    }
    if (priorIncludeById[id] !== undefined && include !== priorIncludeById[id]) {
      throw new Error(`PUL-030RG-009: Existing selection ${id} changed.`);
    }
  }
  const catalog = requiredTable(workbook, "tblMappingReportingGroupCatalog");
  const ch = headerMap(catalog);
  const catalogRows = tableRows(catalog);
  let mappedProducts = 0;
  for (let index = 0; index < catalogRows.length; index += 1) {
    if (text(catalogRows[index][ch.ReportingGroupID]) === plan.id) mappedProducts += 1;
  }
  if (mappedProducts !== 0) throw new Error("PUL-030RG-009: New group unexpectedly received Product mappings.");
  if (requiredSheet(workbook, "_Metric_Calc").getRange("AL16").getText() === "Available" ||
      requiredSheet(workbook, "Mapping").getRange("A9").getText() !== "Performance refresh required") {
    throw new Error("PUL-030RG-009: Performance did not become truthfully stale after creation.");
  }
  if (captureProtectedState(workbook) !== protectedBefore) {
    throw new Error("PUL-030RG-009: A protected source, mapping, cache, import, or legacy result changed.");
  }
}

function captureProtectedState(workbook: ExcelScript.Workbook): string {
  const parts: string[] = [
    "tblSalesFacts", rangeFingerprint(requiredTable(workbook, "tblSalesFacts").getRange()),
    "tblMetricRPGFacts", rangeFingerprint(requiredTable(workbook, "tblMetricRPGFacts").getRange()),
    "tblMetricRPGResults", rangeFingerprint(requiredTable(workbook, "tblMetricRPGResults").getRange()),
    "tblImports", rangeFingerprint(requiredTable(workbook, "tblImports").getRange()),
    "tblRestaurants", rangeFingerprint(requiredTable(workbook, "tblRestaurants").getRange()),
    "tblProducts", rangeFingerprint(requiredTable(workbook, "tblProducts").getRange()),
    "tblSourceClassifications", rangeFingerprint(requiredTable(workbook, "tblSourceClassifications").getRange()),
    "tblMappingRules", rangeFingerprint(requiredTable(workbook, "tblMappingRules").getRange()),
    "tblEffectiveMapping", rangeFingerprint(requiredTable(workbook, "tblEffectiveMapping").getRange()),
    "tblKPIRegistry", rangeFingerprint(requiredTable(workbook, "tblKPIRegistry").getRange()),
    "tblWeeklyCacheVersions", rangeFingerprint(requiredTable(workbook, "tblWeeklyCacheVersions").getRange()),
    "tblWeeklyPeriodManifest", rangeFingerprint(requiredTable(workbook, "tblWeeklyPeriodManifest").getRange()),
    "tblWeeklyScopeCache", rangeFingerprint(requiredTable(workbook, "tblWeeklyScopeCache").getRange()),
    "tblWeeklyRPGCache", rangeFingerprint(requiredTable(workbook, "tblWeeklyRPGCache").getRange()),
    "tblWeeklyIdentityRegistryControl", rangeFingerprint(requiredTable(workbook, "tblWeeklyIdentityRegistryControl").getRange()),
    "tblWeeklyIdentityRestaurants", rangeFingerprint(requiredTable(workbook, "tblWeeklyIdentityRestaurants").getRange()),
    "tblWeeklyIdentityClassifications", rangeFingerprint(requiredTable(workbook, "tblWeeklyIdentityClassifications").getRange()),
    "tblWeeklyIdentityProducts", rangeFingerprint(requiredTable(workbook, "tblWeeklyIdentityProducts").getRange()),
    "tblPerformanceRestaurantSelection", rangeFingerprint(requiredTable(workbook, "tblPerformanceRestaurantSelection").getRange())
  ];
  return hashStrings(parts, "RG-PROTECTED-");
}

function applyRangeValidation(target: ExcelScript.Range, source: ExcelScript.Range): void {
  const validation = target.getDataValidation();
  validation.clear();
  validation.setRule({ list: { inCellDropDown: true, source } });
  validation.setErrorAlert({ showAlert: true, style: ExcelScript.DataValidationAlertStyle.stop,
    title: "Choose an available value", message: "Select a current active Reporting Group." });
}

function writeSettingsStatus(settings: ExcelScript.Worksheet, message: string, error: boolean): void {
  settings.getRange("G23").setValue(message);
  settings.getRange("G23:K24").getFormat().getFill().setColor(error ? "#FCE8E6" : "#E2F0D9");
  settings.getRange("G23:K24").getFormat().getFont().setColor(error ? "#A83126" : "#17365D");
  settings.getRange("G23:K24").getFormat().getFont().setBold(true);
}

function rangeFingerprint(range: ExcelScript.Range): string {
  return hashText(JSON.stringify([range.getValues(), range.getFormulas()]));
}

function requiredTable(workbook: ExcelScript.Workbook, name: string): ExcelScript.Table {
  const table = workbook.getTable(name);
  if (!table) throw new Error(`PUL-030RG-010: Required table ${name} is missing.`);
  return table;
}

function requiredSheet(workbook: ExcelScript.Workbook, name: string): ExcelScript.Worksheet {
  const sheet = workbook.getWorksheet(name);
  if (!sheet) throw new Error(`PUL-030RG-011: Required sheet ${name} is missing.`);
  return sheet;
}

function tableRows(table: ExcelScript.Table): CellValue[][] {
  return table.getRangeBetweenHeaderAndTotal().getValues();
}

function headerMap(table: ExcelScript.Table): { [key: string]: number } {
  const headers = table.getHeaderRowRange().getTexts()[0];
  const result: { [key: string]: number } = {};
  for (let index = 0; index < headers.length; index += 1) result[text(headers[index])] = index;
  return result;
}

function record(kind: string, values: unknown[]): string {
  const parts: string[] = [];
  for (let index = 0; index < values.length; index += 1) {
    const normalized = text(values[index]);
    parts.push(`${normalized.length}:${normalized}`);
  }
  return `${kind}|${parts.join("|")}`;
}

function hashStrings(values: string[], prefix: string): string {
  let left = 0; let right = 0;
  for (let itemIndex = 0; itemIndex < values.length; itemIndex += 1) {
    const value = `${values[itemIndex]}\n`;
    for (let index = 0; index < value.length; index += 1) {
      const code = value.charCodeAt(index);
      left = (left * 131 + code) % 2147483647;
      right = (right * 137 + code) % 2147483629;
    }
  }
  return `${prefix}${left.toString(16).padStart(8, "0")}${right.toString(16).padStart(8, "0")}`;
}

function hashText(value: string): string {
  let left = 0; let right = 0; const input = `${value}\n`;
  for (let index = 0; index < input.length; index += 1) {
    const code = input.charCodeAt(index);
    left = (left * 131 + code) % 2147483647;
    right = (right * 137 + code) % 2147483629;
  }
  return left.toString(16).padStart(8, "0") + right.toString(16).padStart(8, "0");
}

function boundary(value: CellValue): number {
  const parsed = Number(value);
  return value === "" || !Number.isFinite(parsed) ? 0 : parsed;
}

function normalizeDelimited(value: unknown): string {
  return text(value).split(",").map(item => item.trim()).filter(item => item.length > 0).sort().join(", ");
}

function businessNameKey(value: unknown): string { return text(value).replace(/\s+/g, " ").toLocaleLowerCase(); }
function number(value: unknown): number { const parsed = Number(value); return Number.isFinite(parsed) ? parsed : 0; }
function text(value: unknown): string { return String(value === null || value === undefined ? "" : value).trim(); }
function errorText(error: unknown): string { return error instanceof Error ? error.message : String(error); }
