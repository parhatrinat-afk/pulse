import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs";
import test from "node:test";

const scriptUrl = new URL("../office-scripts/Build_0_3_0_UX_Visual_Slice2A.ts", import.meta.url);
const phase2CScriptUrl = new URL("../office-scripts/Build_0_3_0_Phase2C.ts", import.meta.url);
const phase2CModuleUrl = new URL("../src/reporting/interactive-performance.mjs", import.meta.url);
const iaScriptUrl = new URL("../office-scripts/Build_0_3_0_UX_IA_Slice1.ts", import.meta.url);
const source = fs.readFileSync(scriptUrl, "utf8");

function sha256(url) {
  return crypto.createHash("sha256").update(fs.readFileSync(url)).digest("hex");
}

function functionBody(name) {
  const start = source.indexOf(`function ${name}(`);
  assert.ok(start >= 0, `Missing ${name}.`);
  const bodyStart = source.indexOf("{", start);
  let depth = 0;
  let quote = "";
  let escaped = false;
  for (let index = bodyStart; index < source.length; index++) {
    const char = source[index];
    if (quote) {
      if (escaped) escaped = false;
      else if (char === "\\") escaped = true;
      else if (char === quote) quote = "";
      continue;
    }
    if (char === '"' || char === "'" || char === "`") {
      quote = char;
      continue;
    }
    if (char === "{") depth++;
    if (char === "}") depth--;
    if (depth === 0) return source.slice(bodyStart + 1, index);
  }
  assert.fail(`Unbalanced body for ${name}.`);
}

test("accepted Phase 2C, interaction module, and IA Slice 1 remain byte-for-byte unchanged", () => {
  assert.equal(
    sha256(phase2CScriptUrl),
    "c6662a0911c67f570d2bdd5212808d7636d783a0879caa312bebb97d5d58b166"
  );
  assert.equal(
    sha256(phase2CModuleUrl),
    "b85ab143255cdeec5915e0e40586eb906be611c79313e623ddc7dc274a7976a5"
  );
  assert.equal(
    sha256(iaScriptUrl),
    "c3df8d907e3c21e8ccd785c2e5cd2d1c53f55c462bc17edaebb8a2af8d516a0b"
  );
});

test("Visual Slice 2A validates the accepted 47-sheet IA and 16-of-16 QA before mutation", () => {
  const firstMutation = source.indexOf("overview.setShowGridlines(false)");
  assert.ok(firstMutation > source.indexOf("validateAcceptedPhase2C(interactionQa)"));
  assert.ok(firstMutation > source.indexOf("validateIaState(overview"));
  assert.ok(firstMutation > source.indexOf("validateOverviewNavigation(overview)"));
  assert.ok(firstMutation > source.indexOf("validateReportsLinkage(reports)"));
  assert.match(source, /worksheetCountBefore !== 47/);
  assert.match(source, /supportingSheets\.length !== 41/);
  assert.match(source, /overview\.getPosition\(\) !== 0/);
  assert.match(source, /settings\.getPosition\(\) !== 5/);
  for (let index = 1; index <= 16; index++) {
    assert.match(source, new RegExp(`QA-0302C-${String(index).padStart(2, "0")}`));
  }
});

test("gridlines are disabled only on the exact six primary sheets", () => {
  const calls = Array.from(
    source.matchAll(/\b([a-zA-Z]+)\.setShowGridlines\(false\)/g),
    match => match[1]
  );
  assert.deepEqual(calls, ["overview", "performance", "reports", "imports", "mapping", "settings"]);
  assert.doesNotMatch(source, /setShowGridlines\(true\)/);
  assert.doesNotMatch(source, /supportingSheets\[[^\]]+\]\.setShowGridlines/);
});

test("presentation mutation cannot introduce metric formulas or table/schema changes", () => {
  assert.doesNotMatch(source, /\.setFormula|\.setFormulas|\.setFormulaLocal|\.setFormulasLocal/);
  assert.doesNotMatch(source, /\.addTable|\.delete\(|\.resize\(|\.setName\(|\.addRow|\.addRows/);
  assert.match(source, /performanceFormulaAfter !== performanceFormulaBefore/);
  assert.match(source, /reportsFormulaAfter !== reportsFormulaBefore/);
  assert.equal((source.match(/getUsedRange\(true\)\.getFormulas\(\)/g) || []).length, 4);
  assert.match(source, /primaryTableContentAfter !== primaryTableContentBefore/);
  assert.match(source, /primaryTableStructureAfter !== primaryTableStructureBefore/);
  assert.match(source, /worksheetCountAfter !== worksheetCountBefore/);
  assert.match(source, /tableCountAfter !== tableCountBefore/);
});

test("Reports retains exact Phase 2C linkage and receives structural clipping safeguards", () => {
  const reportsBody = functionBody("formatReports");
  for (const formula of [
    "=Performance!B7", "=Performance!B10", "=Performance!G10", "=Performance!G18",
    "=Performance!G19", "=Performance!B16", "=Performance!B17",
    "=Performance!B18", "=Performance!B19"
  ]) {
    assert.match(source, new RegExp(formula.replace(/[!.$]/g, "\\$&")));
  }
  assert.match(reportsBody, /getRange\("A:A"\).*setColumnWidth\(205\)/s);
  assert.match(reportsBody, /getRange\("B:B"\).*setColumnWidth\(315\)/s);
  assert.match(reportsBody, /getRange\("10:11"\).*setRowHeight\(42\)/s);
  assert.match(reportsBody, /getRange\("18:18"\).*setRowHeight\(34\)/s);
  assert.match(reportsBody, /getRange\("B10:D11"\).*setWrapText\(true\)/s);
  assert.doesNotMatch(reportsBody, /setNumberFormat/);
});

test("Performance Explain is reduced to four operational concepts with explicit sizing", () => {
  const body = functionBody("formatPerformance");
  assert.match(body, /\["Metric", "Reporting Group Sales Share"/);
  assert.match(body, /\["Definition", "Selected Reporting Group sales as a share of total sales/);
  assert.match(body, /\["Total", "Total combines the currently selected Reporting Groups\."/);
  assert.match(body, /\["Comparison", "PP Change compares current share with comparison share\. NOK Impact/);
  assert.doesNotMatch(body, /numeric helpers|FIXED facade|fingerprint|formula isolation|SORTBY|Phase 2B results/);
  assert.match(body, /setRowHeight\(48\)/);
  assert.match(body, /setRowHeight\(52\)/);
  assert.equal((body.match(/\.merge\(\)/g) || []).length, 4);
  assert.equal((body.match(/\.unmerge\(\)/g) || []).length, 1);
});

test("Detail selector clarification and editable-state styling are presentation-only", () => {
  const body = functionBody("formatPerformance");
  assert.match(source, /Detail selection controls the result card and Reports; matrix Reporting Group selection is separate\./);
  assert.match(body, /getRange\("A8"\)\.setValue\(DETAIL_NOTE\)/);
  for (const address of ["B7", "B10", "G6", "G10", "I6", "I7"]) {
    assert.match(body, new RegExp(`getRange\\("${address}"\\).*setColor\\(PALE_YELLOW\\)`, "s"));
  }
  assert.match(body, /requiredColumn\(restaurantSelection, "Include"\)/);
  assert.match(body, /requiredColumn\(groupSelection, "Include"\)/);
});

test("Performance matrix calculation and structure are not rewritten by the visual script", () => {
  const valueWrites = Array.from(source.matchAll(/\.setValue(?:s)?\s*\(/g), match => match[0]);
  assert.equal(valueWrites.length, 2);
  assert.match(source, /getRange\("A8"\)\.setValue\(DETAIL_NOTE\)/);
  assert.match(source, /explainBody\.setValues\(/);
  assert.doesNotMatch(source, /getRange\("A23:K40"\)|matrixPresentationFormula|componentBlocks|tblMetricRPGResults/);
  assert.doesNotMatch(source, /setPosition\(|setVisibility\(/);
});

test("Imports and Mapping logic are fingerprinted while only safe presentation ranges are formatted", () => {
  assert.match(source, /tblImports/);
  assert.match(source, /tblMappingMainNodes/);
  assert.match(source, /tblMappingSubcategoryNodes/);
  assert.match(source, /tblMappingProducts/);
  const mappingBody = functionBody("formatMapping");
  const importsBody = functionBody("formatImports");
  assert.match(mappingBody, /getRange\("B5:B12"\).*setColor\(PALE_YELLOW\)/s);
  assert.match(mappingBody, /getRange\("E5:N7"\).*setWrapText\(true\)/s);
  assert.match(importsBody, /getRange\("A4:S4"\).*setWrapText\(true\)/s);
  assert.doesNotMatch(mappingBody + importsBody, /setValue|setFormula|setNumberFormat|addTable|delete/);
});

test("the visual script is rerunnable and adds no protection, drawings, charts, or Phase 3 work", () => {
  assert.match(source, /detailNoteBefore && detailNoteBefore !== DETAIL_NOTE/);
  assert.match(source, /\.unmerge\(\)/);
  assert.doesNotMatch(source, /addConditionalFormat|conditionalFormats|\.protect\(|getProtection|sheetProtection/);
  assert.doesNotMatch(source, /addChart|addShape|addImage|Phase 3|Pulse ♥|labour|review KPI/i);
  assert.doesNotMatch(source, /clear\(ExcelScript\.ClearApplyTo\.all\)/);
});

test("Office Scripts compatibility guard rejects Map/Set iterators and workbook reads inside loops", () => {
  assert.doesNotMatch(source, /new\s+Map\s*</);
  assert.doesNotMatch(source, /new\s+Set\s*</);
  assert.doesNotMatch(source, /\.(entries|keys|values)\s*\(\)/);
  assert.doesNotMatch(source, /Array\.from\s*\(/);
  assert.doesNotMatch(source, /\.\.\./);

  const loopBodies = Array.from(
    source.matchAll(/for\s*\([^)]*\)\s*\{([\s\S]*?)\n  \}/g),
    match => match[1]
  );
  for (const body of loopBodies) {
    assert.doesNotMatch(body, /workbook\.get|sheet\.get|table\.get/);
  }
});

test("static TypeScript delimiter and string scan is balanced", () => {
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
      if (char === "*" && next === "/") {
        blockComment = false;
        index++;
      }
      continue;
    }
    if (quote) {
      if (escaped) escaped = false;
      else if (char === "\\") escaped = true;
      else if (char === quote) quote = "";
      continue;
    }
    if (char === "/" && next === "/") {
      lineComment = true;
      index++;
      continue;
    }
    if (char === "/" && next === "*") {
      blockComment = true;
      index++;
      continue;
    }
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
