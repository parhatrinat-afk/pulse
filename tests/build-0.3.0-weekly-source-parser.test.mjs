import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";
import {
  WEEKLY_SALES_HEADERS,
  WEEKLY_SALES_SCOPE_CONTRACT,
  buildWeeklyCorpusManifest,
  parseWeeklySalesMatrix,
} from "../src/imports/weekly-sales-parser.mjs";

const officeScriptPath = new URL("../office-scripts/Parse_Weekly_Sales_Report.ts", import.meta.url);
const parserPath = new URL("../src/imports/weekly-sales-parser.mjs", import.meta.url);
const auditPath = new URL("../src/imports/audit-weekly-corpus.mjs", import.meta.url);
const expectedCorpusPath = new URL("./expected-build-0.3.0-weekly-source.json", import.meta.url);

test("weekly parser derives period identity only from A1 and preserves exact source strings", () => {
  const parsed = parseWeeklySalesMatrix(matrix(
    "2026-07-27", "2026-08-02",
    [
      ["Swift By Sumo - Oslo S", "Bao Buns*", "Hot & Spicy", "3000 - Salg Mat 25%", "Product*", 2.5, 400],
      ["Bjørvika", "Delivery", "Fees", "3990 - Leveringsgebyr 25%", "Delivery F 100", 1, 100],
    ],
  ), { sourceLocator: "anything/banana.xlsx", sourceBinaryFingerprint: "SHA256-test" });

  assert.equal(parsed.manifest.sourcePeriodKey, "PERIOD-2026-07-27-2026-08-02");
  assert.equal(parsed.manifest.isoYear, 2026);
  assert.equal(parsed.manifest.isoWeek, 31);
  assert.equal(parsed.manifest.sourceRowCount, 2);
  assert.equal(parsed.manifest.totalSalesNok, 500);
  assert.equal(parsed.manifest.totalQuantity, 3.5);
  assert.equal(parsed.rows[0].mainCategory, "Bao Buns*");
  assert.equal(parsed.rows[0].item, "Product*");
  assert.equal(parsed.manifest.sourceLocator, "anything/banana.xlsx");
  assert.equal(parsed.manifest.sourceBinaryFingerprint, "SHA256-test");
  assert.equal(parsed.manifest.contentReconciliationStatus, "PASS");
});

test("filename and locator changes do not change semantic identity", () => {
  const source = matrix("2025-12-29", "2026-01-04", [sampleRow("W01")]);
  const actual = parseWeeklySalesMatrix(source, { sourceLocator: "original.xlsx" });
  const banana = parseWeeklySalesMatrix(source, { sourceLocator: "banana.xlsx" });

  assert.equal(actual.manifest.semanticFingerprint, banana.manifest.semanticFingerprint);
  assert.equal(actual.manifest.sourceFileId, banana.manifest.sourceFileId);
  assert.equal(actual.manifest.sourcePeriodKey, banana.manifest.sourcePeriodKey);
  assert.notEqual(actual.manifest.sourceLocator, banana.manifest.sourceLocator);
});

test("semantic fingerprint is independent of source row order while rows retain lineage", () => {
  const first = sampleRow("First");
  const second = sampleRow("Second");
  const normal = parseWeeklySalesMatrix(matrix("2026-01-05", "2026-01-11", [first, second]));
  const reordered = parseWeeklySalesMatrix(matrix("2026-01-05", "2026-01-11", [second, first]));

  assert.equal(normal.manifest.semanticFingerprint, reordered.manifest.semanticFingerprint);
  assert.equal(normal.manifest.sourceFileId, reordered.manifest.sourceFileId);
  assert.deepEqual(
    normal.rows.map(value => value.sourceRowId).sort(),
    reordered.rows.map(value => value.sourceRowId).sort(),
  );
  assert.deepEqual(normal.rows.map(value => value.sourceRowNumber), [3, 4]);
});

test("equivalent Excel and OOXML floating values have identical semantic identity", () => {
  const ooxml = matrix("2026-08-03", "2026-08-09", [sampleRow("Float")]);
  const excel = matrix("2026-08-03", "2026-08-09", [sampleRow("Float")]);
  ooxml[2][5] = 0.30000000000000004;
  excel[2][5] = 0.3;

  const parsedOoxml = parseWeeklySalesMatrix(ooxml);
  const parsedExcel = parseWeeklySalesMatrix(excel);

  assert.equal(parsedOoxml.manifest.semanticFingerprint, parsedExcel.manifest.semanticFingerprint);
  assert.equal(parsedOoxml.manifest.sourceFileId, parsedExcel.manifest.sourceFileId);
  assert.equal(parsedOoxml.rows[0].sourceRowId, parsedExcel.rows[0].sourceRowId);
  assert.equal(parsedOoxml.rows[0].quantity, 0.30000000000000004);
  assert.equal(parsedExcel.rows[0].quantity, 0.3);
});

test("scope contract never infers channel from delivery products or Sales Account text", () => {
  const parsed = parseWeeklySalesMatrix(matrix("2026-01-12", "2026-01-18", [
    ["Bjørvika", "Takeaway-looking label", "Delivery", "3990 - Leveringsgebyr 15%", "Delivery Fee 500", 1, 500],
  ]));

  assert.equal(parsed.scopeContract.scopeId, WEEKLY_SALES_SCOPE_CONTRACT.scopeId);
  assert.equal(parsed.manifest.channelScope, "Source-defined scope; channel not encoded");
  assert.equal(parsed.manifest.channelFieldEncoded, false);
  assert.equal(parsed.manifest.exportFilterMetadataEncoded, false);
});

test("strict schema, period, text, numeric and bounded-column validation rejects malformed exports", () => {
  assert.throws(
    () => parseWeeklySalesMatrix([["not a period"], WEEKLY_SALES_HEADERS, sampleRow("Bad")]),
    /PUL-030I-004/,
  );
  assert.throws(
    () => parseWeeklySalesMatrix(matrix("2026-01-05", "2026-01-10", [sampleRow("Short")])),
    /PUL-030I-005/,
  );
  const badHeader = matrix("2026-01-05", "2026-01-11", [sampleRow("Header")]);
  badHeader[1][1] = "Category";
  assert.throws(() => parseWeeklySalesMatrix(badHeader), /PUL-030I-008/);
  const badText = matrix("2026-01-05", "2026-01-11", [sampleRow("Text")]);
  badText[2][3] = "";
  assert.throws(() => parseWeeklySalesMatrix(badText), /PUL-030I-012/);
  const badNumber = matrix("2026-01-05", "2026-01-11", [sampleRow("Number")]);
  badNumber[2][5] = "1.00";
  assert.throws(() => parseWeeklySalesMatrix(badNumber), /PUL-030I-013/);
  const extraColumn = matrix("2026-01-05", "2026-01-11", [sampleRow("Extra")]);
  extraColumn[2].push("unexpected");
  assert.throws(() => parseWeeklySalesMatrix(extraColumn), /PUL-030I-011/);
});

test("blank rows inside the source body fail rather than silently changing lineage", () => {
  const source = matrix("2026-02-02", "2026-02-08", [sampleRow("Before"), sampleRow("After")]);
  source.splice(3, 0, ["", "", "", "", "", "", ""]);
  assert.throws(() => parseWeeklySalesMatrix(source), /PUL-030I-010/);
});

test("corpus manifest reconciles consecutive weeks and crosses ISO-year boundaries", () => {
  const reports = [
    parseWeeklySalesMatrix(matrix("2025-12-22", "2025-12-28", [sampleRow("2025-W52")])),
    parseWeeklySalesMatrix(matrix("2025-12-29", "2026-01-04", [sampleRow("2026-W01")])),
    parseWeeklySalesMatrix(matrix("2026-01-05", "2026-01-11", [sampleRow("2026-W02")])),
  ];
  const corpus = buildWeeklyCorpusManifest(reports);

  assert.equal(corpus.status, "PASS");
  assert.equal(corpus.reportCount, 3);
  assert.equal(corpus.uniquePeriodCount, 3);
  assert.equal(corpus.sourceRowCount, 3);
  assert.deepEqual(corpus.gaps, []);
  assert.deepEqual(corpus.overlaps, []);
  assert.deepEqual(corpus.duplicatePeriods, []);
  assert.deepEqual(corpus.coverage, [
    { isoYear: 2025, weeks: [52], reportCount: 1 },
    { isoYear: 2026, weeks: [1, 2], reportCount: 2 },
  ]);
});

test("corpus manifest surfaces duplicate periods, overlaps, gaps and inconsistent scope", () => {
  const first = parseWeeklySalesMatrix(matrix("2026-01-05", "2026-01-11", [sampleRow("A")]));
  const duplicate = parseWeeklySalesMatrix(matrix("2026-01-05", "2026-01-11", [sampleRow("B")]));
  const afterGap = parseWeeklySalesMatrix(matrix("2026-01-19", "2026-01-25", [sampleRow("C")]));
  const duplicated = buildWeeklyCorpusManifest([first, duplicate, afterGap]);
  assert.equal(duplicated.status, "FAIL");
  assert.equal(duplicated.duplicatePeriods.length, 1);
  assert.equal(duplicated.gaps.length, 1);

  const alternateScope = {
    ...WEEKLY_SALES_SCOPE_CONTRACT,
    scopeId: "SCOPE-DIFFERENT",
  };
  const inconsistent = buildWeeklyCorpusManifest([
    first,
    parseWeeklySalesMatrix(matrix("2026-01-12", "2026-01-18", [sampleRow("D")]), {
      scopeContract: alternateScope,
    }),
  ]);
  assert.equal(inconsistent.status, "FAIL");
  assert.equal(inconsistent.scopeFingerprints.length, 2);
});

test("a coverage gap alone prevents the corpus from receiving PASS", () => {
  const corpus = buildWeeklyCorpusManifest([
    parseWeeklySalesMatrix(matrix("2026-01-05", "2026-01-11", [sampleRow("W02")])),
    parseWeeklySalesMatrix(matrix("2026-01-19", "2026-01-25", [sampleRow("W04")])),
  ]);

  assert.equal(corpus.gaps.length, 1);
  assert.equal(corpus.status, "FAIL");
});

test("old accepted baseline fixtures are not an acceptance dependency of the weekly parser", async () => {
  const parser = await readFile(parserPath, "utf8");
  assert.doesNotMatch(parser, /expected-build-0\.3\.0-phase2b|IMP-2025-BASELINE|IMP-2026-W31/);
});

test("Office Script adapter is read-only, batched and Office-Scripts-compatible", async () => {
  const script = await readFile(officeScriptPath, "utf8");
  assert.match(script, /^function main\(/m);
  assert.doesNotMatch(script, /^export\s+function\s+main\(/m);
  assert.match(script, /getUsedRange\(true\)/);
  assert.equal((script.match(/\.getValues\(\)/g) ?? []).length, 1);
  assert.doesNotMatch(script, /\.setValue\(|\.setValues\(|\.clear\(|\.delete\(|\.addTable\(/);
  assert.doesNotMatch(script, /new Map|new Set|\.entries\(|\.keys\(|\.values\(|\.\.\./);
  assert.doesNotMatch(script, /getFileName|sourceLocator.*semanticRecords|sourceLocator.*sourceFileId/);
  assert.match(script, /channelScope: "Source-defined scope; channel not encoded"/);
  assert.match(script, /canonicalFingerprintNumber\(row\.quantity, 6\)/);
  assert.match(script, /canonicalFingerprintNumber\(row\.salesNok, 2\)/);
});

test("frozen corpus evidence records the complete candidate weekly checkpoint", async () => {
  const expected = JSON.parse(await readFile(expectedCorpusPath, "utf8"));
  assert.equal(expected.report_count, 84);
  assert.equal(expected.unique_period_count, 84);
  assert.equal(expected.coverage["2025"].weeks, "W01-W52");
  assert.equal(expected.coverage["2026"].weeks, "W01-W32");
  assert.equal(expected.source_row_count, 245632);
  assert.equal(expected.total_sales_nok, 484728367.25);
  assert.equal(expected.total_quantity, 2469988.09);
  assert.equal(expected.scope_id, WEEKLY_SALES_SCOPE_CONTRACT.scopeId);
  assert.match(expected.scope_note, /must not be inferred/);
  assert.equal(expected.status, "PASS");
});

test("development corpus audit requires an exact caller-supplied path and remains read-only", async () => {
  const audit = await readFile(auditPath, "utf8");
  assert.match(audit, /Provide one exact fixture corpus path/);
  assert.match(audit, /readOnly: true/);
  assert.doesNotMatch(audit, /\/Users\/|process\.env\.(HOME|CODEX_HOME)|homedir\(/);
  assert.doesNotMatch(audit, /writeFile|rename\(|unlink\(|rm\(|copyFile/);
});

function matrix(periodStart, periodEnd, rows) {
  return [
    [`Period: ${periodStart} - ${periodEnd}`],
    [...WEEKLY_SALES_HEADERS],
    ...rows.map(value => [...value]),
  ];
}

function sampleRow(label) {
  return [
    "Bjørvika",
    `Main ${label}`,
    `Sub ${label}`,
    "3000 - Salg Mat 25%",
    `Item ${label}`,
    1,
    100,
  ];
}
