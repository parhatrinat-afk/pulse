#!/usr/bin/env node

/**
 * Read-only development audit for the local weekly fixture corpus.
 *
 * Usage:
 *   node src/imports/audit-weekly-corpus.mjs /exact/corpus/path
 *
 * The exact path is mandatory. This command never searches Desktop, Downloads,
 * OneDrive, or the repository, and it never writes to the source workbooks.
 * Production intake remains OneDrive + Power Automate + Office Scripts.
 */

import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { readdir, readFile, stat } from "node:fs/promises";
import path from "node:path";
import {
  WEEKLY_SALES_SHEET,
  buildWeeklyCorpusManifest,
  parseWeeklySalesMatrix,
} from "./weekly-sales-parser.mjs";

const root = process.argv[2];
if (!root || root.startsWith("--")) {
  throw new Error("Provide one exact fixture corpus path. The audit will not search for source files.");
}
const expectedFlag = process.argv.indexOf("--expected");
const expectedPath = expectedFlag >= 0 ? process.argv[expectedFlag + 1] : "";
if (expectedFlag >= 0 && !expectedPath) throw new Error("--expected requires one exact JSON path.");

const files = await listXlsxFiles(path.resolve(root));
if (files.length === 0) throw new Error(`No .xlsx reports found under ${root}.`);

const parsedReports = [];
const binaryBytes = [];
for (const file of files) {
  const fileBytes = await readFile(file);
  const workbook = readWorkbookMatrix(file);
  if (workbook.sheetNames.length !== 1 || workbook.sheetNames[0] !== WEEKLY_SALES_SHEET) {
    throw new Error(`${file} must contain exactly one '${WEEKLY_SALES_SHEET}' sheet.`);
  }
  const binaryFingerprint = `SHA256-${createHash("sha256").update(fileBytes).digest("hex")}`;
  const parsed = parseWeeklySalesMatrix(workbook.matrix, {
    sourceLocator: path.relative(path.resolve(root), file),
    sourceBinaryFingerprint: binaryFingerprint,
  });
  parsedReports.push(parsed);
  binaryBytes.push((await stat(file)).size);
}

const corpus = buildWeeklyCorpusManifest(parsedReports);
const rows = parsedReports.flatMap(value => value.rows);
const result = {
  root: path.resolve(root),
  readOnly: true,
  reportCount: corpus.reportCount,
  uniquePeriodCount: corpus.uniquePeriodCount,
  firstPeriodStart: corpus.firstPeriodStart,
  lastPeriodEnd: corpus.lastPeriodEnd,
  sourceRowCount: corpus.sourceRowCount,
  totalSalesNok: corpus.totalSalesNok,
  totalQuantity: corpus.totalQuantity,
  totalSourceBytes: binaryBytes.reduce((total, value) => total + value, 0),
  uniqueRestaurantCount: new Set(rows.map(value => value.restaurant)).size,
  uniqueItemCount: new Set(rows.map(value => value.item)).size,
  scopeFingerprints: corpus.scopeFingerprints,
  schemaVersions: corpus.schemaVersions,
  coverage: corpus.coverage,
  duplicatePeriods: corpus.duplicatePeriods,
  duplicateFileIds: corpus.duplicateFileIds,
  overlaps: corpus.overlaps,
  gaps: corpus.gaps,
  restaurantSetSignatures: corpus.restaurantSetSignatures,
  salesAccountSetSignatureCount: corpus.salesAccountSetSignatures.length,
  corpusFingerprint: corpus.corpusFingerprint,
  status: corpus.status,
};

if (expectedPath) {
  const expected = JSON.parse(await readFile(path.resolve(expectedPath), "utf8"));
  const actual = {
    report_count: result.reportCount,
    unique_period_count: result.uniquePeriodCount,
    first_period_start: result.firstPeriodStart,
    last_period_end: result.lastPeriodEnd,
    source_row_count: result.sourceRowCount,
    total_sales_nok: result.totalSalesNok,
    total_quantity: result.totalQuantity,
    total_source_bytes: result.totalSourceBytes,
    unique_restaurant_source_values: result.uniqueRestaurantCount,
    unique_item_source_values: result.uniqueItemCount,
    scope_fingerprint: result.scopeFingerprints[0] ?? "",
    schema_version: result.schemaVersions[0] ?? "",
    corpus_fingerprint: result.corpusFingerprint,
    duplicate_period_count: result.duplicatePeriods.length,
    duplicate_file_id_count: result.duplicateFileIds.length,
    overlap_count: result.overlaps.length,
    gap_count: result.gaps.length,
    status: result.status,
  };
  for (const [key, value] of Object.entries(actual)) {
    if (expected[key] !== value) {
      throw new Error(`Expected corpus checkpoint mismatch for ${key}: expected ${expected[key]}, received ${value}.`);
    }
  }
  result.expectedCheckpointStatus = "PASS";
}

process.stdout.write(`${JSON.stringify(result, null, 2)}\n`);

async function listXlsxFiles(directory) {
  const found = [];
  const entries = await readdir(directory, { withFileTypes: true });
  for (const entry of entries) {
    const location = path.join(directory, entry.name);
    if (entry.isDirectory()) found.push(...await listXlsxFiles(location));
    else if (entry.isFile() && entry.name.toLowerCase().endsWith(".xlsx")) found.push(location);
  }
  return found.sort((left, right) => left.localeCompare(right));
}

function readWorkbookMatrix(file) {
  const workbookXml = unzipEntry(file, "xl/workbook.xml");
  const relationshipXml = unzipEntry(file, "xl/_rels/workbook.xml.rels");
  const sharedStrings = readSharedStrings(file);
  const relationshipTargets = parseRelationships(relationshipXml);
  const sheetMatches = [...workbookXml.matchAll(/<sheet\b[^>]*\bname="([^"]+)"[^>]*\br:id="([^"]+)"[^>]*\/?\s*>/g)];
  const sheets = sheetMatches.map(match => ({
    name: decodeXml(match[1]),
    target: normalizeSheetTarget(relationshipTargets[match[2]]),
  }));
  const sourceSheet = sheets.find(value => value.name === WEEKLY_SALES_SHEET);
  if (!sourceSheet) throw new Error(`${file} has no '${WEEKLY_SALES_SHEET}' sheet.`);
  return {
    sheetNames: sheets.map(value => value.name),
    matrix: parseWorksheet(unzipEntry(file, sourceSheet.target), sharedStrings),
  };
}

function readSharedStrings(file) {
  const result = spawnSync("unzip", ["-p", file, "xl/sharedStrings.xml"], { encoding: "utf8" });
  if (result.status !== 0) return [];
  return [...result.stdout.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>/g)].map(match => {
    const parts = [...match[1].matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(value => decodeXml(value[1]));
    return parts.join("");
  });
}

function parseRelationships(xml) {
  const targets = Object.create(null);
  for (const match of xml.matchAll(/<Relationship\b[^>]*\bId="([^"]+)"[^>]*\bTarget="([^"]+)"[^>]*\/?\s*>/g)) {
    targets[match[1]] = decodeXml(match[2]);
  }
  return targets;
}

function normalizeSheetTarget(target) {
  if (!target) throw new Error("Workbook sheet relationship is missing.");
  const normalized = target.replace(/^\//, "");
  return normalized.startsWith("xl/") ? normalized : `xl/${normalized}`;
}

function parseWorksheet(xml, sharedStrings) {
  const cells = [];
  let maxRow = 0;
  let maxColumn = 0;
  for (const match of xml.matchAll(/<c\b([^>]*)>([\s\S]*?)<\/c>/g)) {
    const attributes = match[1];
    const body = match[2];
    const ref = /\br="([A-Z]+)(\d+)"/.exec(attributes);
    if (!ref) continue;
    const row = Number(ref[2]) - 1;
    const column = columnIndex(ref[1]);
    const type = /\bt="([^"]+)"/.exec(attributes)?.[1] ?? "n";
    const valueMatch = /<v\b[^>]*>([\s\S]*?)<\/v>/.exec(body);
    let value = "";
    if (type === "inlineStr") {
      value = [...body.matchAll(/<t\b[^>]*>([\s\S]*?)<\/t>/g)].map(item => decodeXml(item[1])).join("");
    } else if (!valueMatch) {
      value = "";
    } else if (type === "s") {
      value = sharedStrings[Number(valueMatch[1])] ?? "";
    } else if (type === "str") {
      value = decodeXml(valueMatch[1]);
    } else if (type === "b") {
      value = valueMatch[1] === "1";
    } else {
      const numeric = Number(valueMatch[1]);
      value = Number.isFinite(numeric) ? numeric : decodeXml(valueMatch[1]);
    }
    cells.push({ row, column, value });
    maxRow = Math.max(maxRow, row);
    maxColumn = Math.max(maxColumn, column);
  }
  const matrix = Array.from({ length: maxRow + 1 }, () => Array(maxColumn + 1).fill(""));
  for (const cell of cells) matrix[cell.row][cell.column] = cell.value;
  return matrix;
}

function columnIndex(letters) {
  let value = 0;
  for (const letter of letters) value = value * 26 + letter.charCodeAt(0) - 64;
  return value - 1;
}

function decodeXml(value) {
  return String(value)
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

function unzipEntry(file, entry) {
  const result = spawnSync("unzip", ["-p", file, entry], {
    encoding: "utf8",
    maxBuffer: 32 * 1024 * 1024,
  });
  if (result.status !== 0) {
    throw new Error(`Unable to read ${entry} from ${file}: ${result.stderr.trim()}`);
  }
  return result.stdout;
}
