/**
 * Reporting Group creation contract.
 *
 * tblReportingGroups remains the sole authority. This module plans one safe
 * append; it does not create mappings, rebuild caches, or mutate Performance.
 */

export const REPORTING_GROUP_CREATE_CONTRACT_VERSION = "0.3.0-reporting-group-create-v1";
export const SALES_DOMAIN_ID = "DOM-SALES";

export function validateReportingGroupAuthority(reportingGroups) {
  if (!Array.isArray(reportingGroups) || reportingGroups.length < 1) {
    throw new Error("Reporting Group authority must contain at least one row.");
  }
  const ids = new Set();
  const sortOrders = new Set();
  const activeNames = new Set();
  const activeDomains = new Set();
  let highestId = 0;
  let idWidth = 4;
  let highestSortOrder = 0;

  for (const source of reportingGroups) {
    const row = normalizeAuthorityRow(source);
    const match = /^RPG-(\d+)$/.exec(row.reportingGroupId);
    if (!match || Number(match[1]) < 1) {
      throw new Error(`Reporting Group authority contains invalid stable ID ${row.reportingGroupId || "(blank)"}.`);
    }
    if (ids.has(row.reportingGroupId)) {
      throw new Error(`Reporting Group authority repeats ReportingGroupID ${row.reportingGroupId}.`);
    }
    ids.add(row.reportingGroupId);
    highestId = Math.max(highestId, Number(match[1]));
    idWidth = Math.max(idWidth, match[1].length);

    if (!row.reportingGroupName) {
      throw new Error(`Reporting Group ${row.reportingGroupId} has a blank business name.`);
    }
    if (!Number.isInteger(row.sortOrder) || row.sortOrder < 1) {
      throw new Error(`Reporting Group ${row.reportingGroupId} has an invalid SortOrder.`);
    }
    if (sortOrders.has(row.sortOrder)) {
      throw new Error(`Reporting Group authority repeats SortOrder ${row.sortOrder}.`);
    }
    sortOrders.add(row.sortOrder);
    highestSortOrder = Math.max(highestSortOrder, row.sortOrder);

    if (row.active === "Yes") {
      const nameKey = businessNameKey(row.reportingGroupName);
      if (activeNames.has(nameKey)) {
        throw new Error(`Active Reporting Groups repeat business name ${row.reportingGroupName}.`);
      }
      activeNames.add(nameKey);
      if (!row.domainId) {
        throw new Error(`Active Reporting Group ${row.reportingGroupId} has a blank DomainID.`);
      }
      activeDomains.add(row.domainId);
    } else if (row.active !== "No") {
      throw new Error(`Reporting Group ${row.reportingGroupId} has invalid Active value ${row.active || "(blank)"}.`);
    }
  }
  if (!activeNames.size) throw new Error("At least one active Reporting Group is required.");
  if (activeDomains.size !== 1 || !activeDomains.has(SALES_DOMAIN_ID)) {
    throw new Error(`Active Reporting Groups must use the existing ${SALES_DOMAIN_ID} sales domain.`);
  }

  const nextNumber = highestId + 1;
  const nextReportingGroupId = `RPG-${String(nextNumber).padStart(idWidth, "0")}`;
  if (ids.has(nextReportingGroupId)) {
    throw new Error(`Next ReportingGroupID ${nextReportingGroupId} is already present.`);
  }
  let nextSortOrder = (Math.floor(highestSortOrder / 10) + 1) * 10;
  while (sortOrders.has(nextSortOrder)) nextSortOrder += 10;
  return {
    rows: reportingGroups.map(normalizeAuthorityRow),
    nextReportingGroupId,
    nextSortOrder,
    domainId: SALES_DOMAIN_ID,
    activeNameKeys: activeNames,
  };
}

export function planReportingGroupCreation({
  reportingGroups,
  reportingGroupName,
  description = "",
  notes = "",
}) {
  const authority = validateReportingGroupAuthority(reportingGroups);
  const name = text(reportingGroupName);
  if (!name) throw new Error("Reporting Group name is required.");
  if (authority.activeNameKeys.has(businessNameKey(name))) {
    throw new Error(`Active Reporting Group name ${name} already exists.`);
  }
  return {
    contractVersion: REPORTING_GROUP_CREATE_CONTRACT_VERSION,
    reportingGroup: {
      reportingGroupId: authority.nextReportingGroupId,
      reportingGroupName: name,
      domainId: authority.domainId,
      active: "Yes",
      sortOrder: authority.nextSortOrder,
      description: text(description),
      notes: text(notes),
    },
    performanceRefreshRequired: true,
    mappingRulesCreated: 0,
  };
}

export function buildReportingGroupBusinessOverview(reportingGroups, impacts = []) {
  const authority = validateReportingGroupAuthority(reportingGroups);
  const impactById = new Map();
  for (const source of impacts ?? []) {
    const id = text(source.reportingGroupId ?? source.id);
    if (!id) continue;
    if (impactById.has(id)) throw new Error(`Reporting Group impact repeats ${id}.`);
    impactById.set(id, {
      products: finiteNumber(source.products),
      salesNok: finiteNumber(source.salesNok ?? source.sales),
    });
  }
  return authority.rows
    .slice()
    .sort((left, right) => left.sortOrder - right.sortOrder ||
      left.reportingGroupId.localeCompare(right.reportingGroupId))
    .map(group => {
      const impact = impactById.get(group.reportingGroupId) ?? { products: 0, salesNok: 0 };
      return {
        reportingGroupId: group.reportingGroupId,
        reportingGroupName: group.reportingGroupName,
        status: group.active === "Yes" ? "Active" : "Inactive",
        products: impact.products,
        salesNok: impact.salesNok,
      };
    });
}

function normalizeAuthorityRow(source) {
  return {
    reportingGroupId: text(source?.reportingGroupId ?? source?.ReportingGroupID ?? source?.id),
    reportingGroupName: text(source?.reportingGroupName ?? source?.ReportingGroupName ?? source?.name),
    domainId: text(source?.domainId ?? source?.DomainID),
    active: text(source?.active ?? source?.Active),
    sortOrder: Number(source?.sortOrder ?? source?.SortOrder),
    description: text(source?.description ?? source?.Description),
    notes: text(source?.notes ?? source?.Notes),
  };
}

function businessNameKey(value) {
  return text(value).replace(/\s+/g, " ").toLocaleLowerCase("en");
}

function finiteNumber(value) {
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : 0;
}

function text(value) {
  return String(value === null || value === undefined ? "" : value).trim();
}
