import { spawnSync } from "node:child_process";
import { readdir } from "node:fs/promises";
import path from "node:path";
import { WEEKLY_SALES_SHEET } from "./weekly-sales-parser.mjs";

/**
 * Read-only local fixture helpers shared by weekly parser/preflight audits.
 * Production intake remains OneDrive + Power Automate + Office Scripts.
 */
export async function listWeeklyXlsxFiles(directory) {
  const found = [];
  const entries = await readdir(directory, { withFileTypes: true });
  for (const entry of entries) {
    const location = path.join(directory, entry.name);
    if (entry.isDirectory()) found.push(...await listWeeklyXlsxFiles(location));
    else if (entry.isFile() && entry.name.toLowerCase().endsWith(".xlsx")) found.push(location);
  }
  return found.sort(compareText);
}

export function readWeeklyWorkbookMatrix(file) {
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

function compareText(left, right) {
  const a = String(left);
  const b = String(right);
  return a < b ? -1 : a > b ? 1 : 0;
}
