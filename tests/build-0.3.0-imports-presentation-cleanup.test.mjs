import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";

const scriptPath = new URL("../office-scripts/Build_0_3_0_Imports_Presentation_Cleanup.ts", import.meta.url);
const source = fs.readFileSync(scriptPath, "utf8");

function functionBody(name) {
  const start = source.indexOf(`function ${name}`);
  assert.notEqual(start, -1, `missing function ${name}`);
  let depth = 0;
  let opened = false;
  for (let index = source.indexOf("{", start); index < source.length; index++) {
    if (source[index] === "{") { depth++; opened = true; }
    if (source[index] === "}") depth--;
    if (opened && depth === 0) return source.slice(start, index + 1);
  }
  throw new Error(`unterminated function ${name}`);
}

test("Imports presentation is driven by weekly authority and genuine ledger tables", () => {
  assert.match(source, /tblWeeklyCacheVersions/);
  assert.match(source, /tblWeeklyPeriodManifest/);
  assert.match(source, /tblWeeklyIntakeLog/);
  assert.match(source, /tblWeeklyIntakeLog\[IntakeEventID\]<>""/);
  assert.match(source, /SORTBY\(HSTACK\(period,friendly,stamp,report,rows,sales,message\),stamp,-1\)/);
  assert.doesNotMatch(source, /tblWeeklyScopeCache|tblWeeklyRPGCache/);
});

test("visible Imports copy and columns are operational rather than technical", () => {
  for (const label of [
    "Weekly sales reports processed by Pulse.", "Latest published", "Coverage", "Status",
    "Weekly Import Activity", "Period", "Processed", "Source report", "Rows", "Sales NOK", "Message"
  ]) assert.match(source, new RegExp(label.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")));
  assert.match(source, /Duplicate — no data change/);
  assert.match(source, /Review required/);
  assert.match(source, /Blocked — cache stale/);
  assert.match(source, /Already processed; no data changed\./);
});

test("technical locators and cache identifiers are not exposed in the visible view", () => {
  assert.match(source, /RIGHT\(LOWER\(locators\),5\)="\.xlsx"/);
  assert.match(source, /"Source report"/);
  const install = functionBody("installPresentation");
  assert.doesNotMatch(install, /CacheVersion|CacheFingerprint|MappingContentFingerprint|SourcePeriodKey/);
  assert.doesNotMatch(ACTIVE_VIEW_TEXT(), /WCV-|WCC-|MCF-|IDP-/);
});

test("legacy Imports table stays intact but outside the normal view", () => {
  assert.match(source, /Imports!A4:S6/);
  assert.match(source, /sheet\.getRange\("4:7"\)\.setRowHidden\(true\)/);
  assert.match(source, /sheet\.getRange\("J:S"\)\.setColumnHidden\(true\)/);
  assert.match(source, /tableFingerprint\(imports\)/);
  assert.doesNotMatch(source, /delete\(.*tblImports|importsTable\.delete/);
});

test("summary validates the single active cache, rollback and accepted weekly coverage", () => {
  assert.match(source, /activeRows\.length !== 1/);
  assert.match(source, /rollbackRows\.length !== 1/);
  assert.match(source, /PeriodRowCount"\)\) !== 85/);
  assert.match(source, /2025 W01/);
  assert.match(source, /2025 W52/);
  assert.match(source, /2026 W33/);
  assert.match(source, /COUNTIFS\(tblWeeklyIntakeLog\[SourcePeriodKey\],k,tblWeeklyIntakeLog\[IntakeStatus\],"Published"\)/);
});

test("Performance, Reports, Mapping, QA and six-sheet IA are protected", () => {
  assert.match(source, /performance\.getUsedRange\(true\)/);
  assert.match(source, /reports\.getUsedRange\(true\)/);
  assert.match(source, /mapping\.getUsedRange\(true\)/);
  assert.match(source, /QA-030WP-/);
  assert.match(source, /QA-0302C-/);
  assert.match(source, /Overview", "Performance", "Reports", "Imports", "Mapping", "Settings/);
  assert.doesNotMatch(source, /getWorksheet\("Performance"\)\.[\s\S]*setFormula/);
});

test("presentation uses bounded formatting and Office Scripts-compatible iteration", () => {
  assert.match(source, /A8:I53/);
  assert.match(source, /A14:G53/);
  assert.match(source, /Row 7 is intentionally left blank/);
  assert.match(source, /A7:S7/);
  assert.match(source, /CalculationType\.full/);
  assert.doesNotMatch(source, /autofitColumns|autofitRows/);
  assert.doesNotMatch(source, /new Map|new Set|\.entries\(\)|\.keys\(\)|\.values\(\)|\.forEach\(/);
  assert.doesNotMatch(source, /for \(const \[[^\]]+\] of/);
});

test("Office Script static delimiter and string scan is balanced", () => {
  let braces = 0;
  let parentheses = 0;
  let brackets = 0;
  let quote = "";
  let escaped = false;
  let lineComment = false;
  let blockComment = false;
  for (let index = 0; index < source.length; index++) {
    const char = source[index];
    const next = source[index + 1];
    if (lineComment) {
      if (char === "\n") lineComment = false;
      continue;
    }
    if (blockComment) {
      if (char === "*" && next === "/") { blockComment = false; index++; }
      continue;
    }
    if (quote) {
      if (escaped) escaped = false;
      else if (char === "\\") escaped = true;
      else if (char === quote) quote = "";
      continue;
    }
    if (char === "/" && next === "/") { lineComment = true; index++; continue; }
    if (char === "/" && next === "*") { blockComment = true; index++; continue; }
    if (char === '"' || char === "'" || char === "`") quote = char;
    else if (char === "{") braces++;
    else if (char === "}") braces--;
    else if (char === "(") parentheses++;
    else if (char === ")") parentheses--;
    else if (char === "[") brackets++;
    else if (char === "]") brackets--;
    assert.ok(braces >= 0 && parentheses >= 0 && brackets >= 0, `Unbalanced delimiter at ${index}.`);
  }
  assert.equal(quote, "");
  assert.equal(blockComment, false);
  assert.equal(braces, 0);
  assert.equal(parentheses, 0);
  assert.equal(brackets, 0);
  assert.match(source, /function main\(workbook: ExcelScript\.Workbook\): string/);
});

function ACTIVE_VIEW_TEXT() {
  return [LATEST_FORMULA_TEXT(), COVERAGE_FORMULA_TEXT(), STATUS_FORMULA_TEXT(), ACTIVITY_FORMULA_TEXT()].join("\n");
}

function constantText(name) {
  const match = source.match(new RegExp(`const ${name} =([\\s\\S]*?);\\n`));
  assert.ok(match, `missing constant ${name}`);
  return match[1];
}

function LATEST_FORMULA_TEXT() { return constantText("LATEST_FORMULA"); }
function COVERAGE_FORMULA_TEXT() { return constantText("COVERAGE_FORMULA"); }
function STATUS_FORMULA_TEXT() { return constantText("STATUS_FORMULA"); }
function ACTIVITY_FORMULA_TEXT() { return constantText("ACTIVITY_FORMULA"); }
