import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs";
import test from "node:test";

const scriptUrl = new URL("../office-scripts/Build_0_3_0_UX_Visual_Slice2B.ts", import.meta.url);
const slice2AUrl = new URL("../office-scripts/Build_0_3_0_UX_Visual_Slice2A.ts", import.meta.url);
const phase2CUrl = new URL("../office-scripts/Build_0_3_0_Phase2C.ts", import.meta.url);
const iaUrl = new URL("../office-scripts/Build_0_3_0_UX_IA_Slice1.ts", import.meta.url);
const source = fs.readFileSync(scriptUrl, "utf8");

function sha256(url) {
  return crypto.createHash("sha256").update(fs.readFileSync(url)).digest("hex");
}

test("accepted Phase 2C, IA, and Visual Slice 2A sources remain unchanged", () => {
  assert.equal(sha256(phase2CUrl), "c6662a0911c67f570d2bdd5212808d7636d783a0879caa312bebb97d5d58b166");
  assert.equal(sha256(iaUrl), "c3df8d907e3c21e8ccd785c2e5cd2d1c53f55c462bc17edaebb8a2af8d516a0b");
  assert.equal(sha256(slice2AUrl), "1b8461bec855d2aa2457838989461ad51a0d744748682d35b8f03f9bdfeff508");
});

test("Slice 2B requires the accepted layout and 16-of-16 QA before mutation", () => {
  const firstMutation = source.indexOf("formatPerformance(performance");
  assert.ok(firstMutation > source.indexOf("validatePrimaryIa("));
  assert.ok(firstMutation > source.indexOf("validateAcceptedPhase2C(interactionQa)"));
  assert.ok(firstMutation > source.indexOf("validateSlice2AState("));
  assert.match(source, /worksheetCountBefore !== 47/);
  assert.match(source, /Performance!N5:P21/);
  assert.match(source, /Performance!R5:T14/);
  for (let index = 1; index <= 16; index++) {
    assert.match(source, new RegExp(`QA-0302C-${String(index).padStart(2, "0")}`));
  }
});

test("mutation surface is limited to bounded Performance and Mapping presentation", () => {
  assert.match(source, /getRange\("A2:T2"\)/);
  assert.match(source, /setRowHeight\(28\)/);
  assert.match(source, /getRange\("6:14"\).*setRowHeight\(19\)/);
  assert.match(source, /getRange\("16:20"\).*setRowHeight\(19\)/);
  assert.match(source, /getRange\("24:40"\).*setRowHeight\(19\)/);
  assert.match(source, /getRange\("A2:N2"\)/);
  assert.match(source, /getRange\("5:8"\).*setRowHeight\(32\)/);
  assert.match(source, /getRange\("E5:N8"\).*setWrapText\(false\)/);
  assert.match(source, /getRange\("B10"\)\.setNumberFormat\(\[\["dd\.mm\.yyyy"\]\]\)/);
  assert.doesNotMatch(source, /setValue|setValues|setFormula|setFormulas/);
  assert.doesNotMatch(source, /autofit/);
});

test("visual hierarchy uses restrained semantic colors and compact widths", () => {
  assert.match(source, /const NAVY = "#17233A"/);
  assert.match(source, /const BLUE = "#4F86F7"/);
  assert.match(source, /const PALE_YELLOW = "#FFF2CC"/);
  assert.match(source, /const PALE_GREEN = "#E2F0D9"/);
  assert.match(source, /getRange\("N:N"\).*setColumnWidth\(65\)/);
  assert.match(source, /getRange\("O:O"\).*setColumnWidth\(145\)/);
  assert.match(source, /getRange\("R:R"\).*setColumnWidth\(65\)/);
  assert.match(source, /getRange\("T:T"\).*setColumnWidth\(110\)/);
  assert.match(source, /getRange\("G16:J19"\).*setWrapText\(false\)/);
  assert.match(source, /getRange\("A40:K40"\).*setColor\(LIGHT_BLUE\)/);
  assert.match(source, /getRange\("E8:N8"\).*setColor\(PALE_GREEN\)/);
});

test("formula, value, table, validation, QA, and IA guards run after repair", () => {
  assert.match(source, /performanceValuesAfter !== performanceValuesBefore/);
  assert.match(source, /mappingValuesAfter !== mappingValuesBefore/);
  assert.match(source, /reportsValuesAfter !== reportsValuesBefore/);
  assert.match(source, /performanceFormulasAfter !== performanceFormulasBefore/);
  assert.match(source, /mappingFormulasAfter !== mappingFormulasBefore/);
  assert.match(source, /reportsFormulasAfter !== reportsFormulasBefore/);
  assert.match(source, /protectedContentAfter !== protectedContentBefore/);
  assert.match(source, /protectedStructureAfter !== protectedStructureBefore/);
  assert.match(source, /validationAfter !== validationBefore/);
  assert.match(source, /validateAcceptedPhase2C\(interactionQa\);\s*validateSlice2AState/s);
  assert.match(source, /validateRepairPostconditions/);
});

test("tables, schemas, validations, sheet state, and functionality cannot be mutated", () => {
  assert.doesNotMatch(source, /addTable|\.delete\(|\.resize\(|setName\(|addRow|addRows/);
  assert.doesNotMatch(source, /setRule|clear\(ExcelScript\.ClearApplyTo|setVisibility|setPosition/);
  assert.doesNotMatch(source, /setHyperlink|addConditionalFormat|clearAllConditionalFormats/);
  assert.doesNotMatch(source, /activate\(|select\(|setSelected/);
  assert.doesNotMatch(source, /_Metric_Calc|tblMetricRPGResults|NOK Impact|Grand Total/);
});

test("Office Scripts compatibility and static syntax guards remain clean", () => {
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
