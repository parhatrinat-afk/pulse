import assert from "node:assert/strict";
import fs from "node:fs";
import test from "node:test";

const scriptUrl = new URL("../office-scripts/Build_0_3_0_Performance_Presentation_Cleanup.ts", import.meta.url);
const docUrl = new URL("../docs/BUILD_0_3_0_PERFORMANCE_PRESENTATION_CLEANUP.md", import.meta.url);
const source = fs.readFileSync(scriptUrl, "utf8");
const docs = fs.readFileSync(docUrl, "utf8");

test("cleanup requires both accepted 16-of-16 QA gates before mutation", () => {
  const mutation = source.indexOf("moveSelectionTables(performance");
  assert.ok(mutation > source.indexOf('validateQa(interactionQa, "QA-0302C-"'));
  assert.ok(mutation > source.indexOf('validateQa(weeklyQa, "QA-030WP-"'));
  assert.match(source, /worksheetCountBefore !== 48/);
  assert.match(source, /visibleCount !== 6/);
});

test("existing selection tables move natively below Explain without reconstruction", () => {
  assert.match(source, /getRange\("N4:T21"\)\.moveTo\(performance\.getRange\("B50:H67"\)\)/);
  assert.match(source, /Performance!B51:D67/);
  assert.match(source, /Performance!F51:H60/);
  assert.doesNotMatch(source, /addTable|\.delete\(|\.resize\(|setName\(|addRow|addRows/);
  assert.match(docs, /native range move from\s+`N4:T21` to `B50:H67`/);
});

test("selection content, stable IDs, formulas, and validation are protected", () => {
  assert.match(source, /tableContentFingerprint\(restaurantSelection/);
  assert.match(source, /tableContentFingerprint\(groupSelection/);
  assert.match(source, /selectionValidationFingerprint/);
  assert.match(source, /Selection values, stable IDs, or formulas changed/);
  assert.match(source, /A Performance formula changed/);
  assert.match(source, /A Current\/Compare selector value or formula changed/);
  assert.doesNotMatch(source, /setRule|clear\(ExcelScript\.ClearApplyTo/);
});

test("matrix presentation changes alignment and bounded widths only", () => {
  assert.match(source, /getRange\("A23:A40"\).*HorizontalAlignment\.left/);
  assert.match(source, /getRange\("B23:K40"\).*HorizontalAlignment\.center/);
  assert.match(source, /getRange\("B23:K23"\).*setWrapText\(true\)/);
  assert.match(source, /165, 98, 98, 102, 108, 100, 105, 110, 100, 95, 95/);
  assert.match(source, /getRange\("51:51"\).*setRowHeight\(32\)/);
  assert.match(source, /getHeaderRowRange\(\).*setWrapText\(true\)/);
  assert.match(docs, /Beer & Cider is reduced\s+from 180 to 105/);
  assert.doesNotMatch(source, /setNumberFormat|addConditionalFormat|clearAllConditionalFormats/);
});

test("Explain and weekly selector label cleanup are exact and bounded", () => {
  assert.match(source, /selected restaurants and period/);
  assert.match(source, /\["A10:A12", "F10:F12"\]/);
  assert.equal(
    (source.match(/\.setValue\(/g) || []).length,
    1,
    "Only the approved Explain wording should write a cell value."
  );
  assert.doesNotMatch(source, /setFormula|setFormulas|setValues/);
});

test("Reports, rollback results, Imports, and accepted matrix results are guarded", () => {
  assert.match(source, /rangeFingerprint\(reports\.getUsedRange\(true\)/);
  assert.match(source, /tableContentFingerprint\(metricResults/);
  assert.match(source, /tableContentFingerprint\(imports/);
  assert.match(source, /rangeFingerprint\(performance\.getRange\("A23:K40"\)/);
});

test("Office Scripts compatibility guards remain clean", () => {
  assert.doesNotMatch(source, /new\s+Map\s*</);
  assert.doesNotMatch(source, /new\s+Set\s*</);
  assert.doesNotMatch(source, /\.(entries|keys|values)\s*\(\)/);
  assert.doesNotMatch(source, /Array\.from\s*\(/);
  assert.doesNotMatch(source, /\.\.\./);

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
});
