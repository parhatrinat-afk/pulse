import assert from "node:assert/strict";
import crypto from "node:crypto";
import fs from "node:fs";
import test from "node:test";

const scriptUrl = new URL("../office-scripts/Build_0_3_0_UX_IA_Slice1.ts", import.meta.url);
const phase2CScriptUrl = new URL("../office-scripts/Build_0_3_0_Phase2C.ts", import.meta.url);
const phase2CModuleUrl = new URL("../src/reporting/interactive-performance.mjs", import.meta.url);
const source = fs.readFileSync(scriptUrl, "utf8");

const acceptedCheckpointOrder = [
  "Overview", "Performance", "Reports", "Imports", "Settings", "KPI Registry",
  "Context", "Views", "Restaurants", "Reporting Categories", "Products",
  "Source Classifications", "Remap Assistant", "Remap Rules", "Import Exclusions",
  "Expected Coverage", "Publication Control", "Import Actions", "Import Certificates",
  "Domains", "Source Systems", "Adapters", "Adapter Contract", "Import Control",
  "Test Run Control", "_Raw_2025_Baseline", "_Raw_2026_Week31", "_Sales_Facts",
  "_Standard_Staging", "_Remap_Audit", "_Import_Action_Audit", "_Build_Log",
  "_Environment", "_Lists", "Effective Categories", "_Metric_Calc",
  "Reporting Groups", "Mapping Rules", "Effective Mapping", "Mapping",
  "_Mapping_Lists", "Mapping QA", "Metric Contract", "Metric Equivalence",
  "_Metric_RPG_Facts", "Metric Migration QA", "Metric Results QA"
];

const expectedPrimarySheets = [
  "Overview", "Performance", "Reports", "Imports", "Mapping", "Settings"
];

function namesFromWorksheetArray(variableName) {
  const match = source.match(new RegExp(
    `const ${variableName}: ExcelScript\\.Worksheet\\[\\] = \\[([\\s\\S]*?)\\n  \\];`
  ));
  assert.ok(match, `Missing ${variableName} worksheet array.`);
  return Array.from(
    match[1].matchAll(/requiredSheet\(workbook, "([^"]+)"\)/g),
    value => value[1]
  );
}

function sha256(url) {
  return crypto.createHash("sha256").update(fs.readFileSync(url)).digest("hex");
}

test("IA Slice 1 exposes exactly the six approved primary sheets", () => {
  const primary = namesFromWorksheetArray("primarySheets");
  const supporting = namesFromWorksheetArray("supportingSheets");
  assert.deepEqual(primary, expectedPrimarySheets);
  assert.equal(supporting.length, 41);
  assert.equal(primary.length + supporting.length, 47);
  assert.equal(new Set(primary.concat(supporting)).size, 47);
  assert.deepEqual(
    primary.concat(supporting).slice().sort(),
    acceptedCheckpointOrder.slice().sort()
  );
  assert.match(source, /sheet\.setVisibility\(ExcelScript\.SheetVisibility\.hidden\)/);
  assert.doesNotMatch(source, /SheetVisibility\.veryHidden/);
});

test("the primary tabs move first in the exact approved order without deleting sheets", () => {
  const simulated = acceptedCheckpointOrder.slice();
  for (let index = 0; index < expectedPrimarySheets.length; index++) {
    const sheetName = expectedPrimarySheets[index];
    const oldIndex = simulated.indexOf(sheetName);
    assert.notEqual(oldIndex, -1);
    simulated.splice(oldIndex, 1);
    simulated.splice(index, 0, sheetName);
  }
  assert.deepEqual(simulated.slice(0, 6), expectedPrimarySheets);
  assert.equal(simulated.length, acceptedCheckpointOrder.length);
  assert.deepEqual(simulated.slice().sort(), acceptedCheckpointOrder.slice().sort());
  assert.match(source, /primarySheets\[index\]\.setPosition\(index\)/);
});

test("Overview navigation targets valid primary sheets and ordinary A1 document links", () => {
  assert.match(source, /primarySheets\[0\]\.getRange\("E8"\)/);
  assert.match(source, /primarySheets\[0\]\.getRange\("E9"\)/);
  assert.match(source, /primarySheets\[0\]\.getRange\("E10"\)/);
  assert.match(source, /primarySheets\[0\]\.getRange\("E11"\)/);
  assert.match(source, /primarySheets\[0\]\.getRange\("E12"\)/);
  assert.match(source, /setInternalLink\(overviewNavigation\[0\], "Performance"/);
  assert.match(source, /setInternalLink\(overviewNavigation\[1\], "Reports"/);
  assert.match(source, /setInternalLink\(overviewNavigation\[2\], "Imports"/);
  assert.match(source, /setInternalLink\(overviewNavigation\[3\], "Mapping"/);
  assert.match(source, /setInternalLink\(overviewNavigation\[4\], "Settings"/);
  assert.match(source, /documentReference: `\$\{sheetName\}!A1`/);
  assert.match(source, /target\.setHyperlink\(/);
});

test("all six primary saved views are reset to A1 and Overview is final", () => {
  const anchorBlock = source.match(
    /const primaryAnchors: ExcelScript\.Range\[\] = \[([\s\S]*?)\n  \];/
  );
  assert.ok(anchorBlock);
  assert.equal((anchorBlock[1].match(/\.getRange\("A1"\)/g) || []).length, 6);
  assert.match(source, /primarySheets\[index\]\.activate\(\);\s*primaryAnchors\[index\]\.select\(\)/);
  assert.match(source, /primarySheets\[0\]\.activate\(\);\s*primaryAnchors\[0\]\.select\(\);/);
});

test("the script requires the exact accepted 16-of-16 Phase 2C checkpoint before mutation", () => {
  const validationCall = source.indexOf("validateAcceptedPhase2C(interactionQa)");
  const firstMutation = source.indexOf(".setVisibility(");
  assert.ok(validationCall > 0);
  assert.ok(firstMutation > validationCall);
  assert.match(source, /requiredTable\(workbook, "tblPerformanceInteractionQA"\)/);
  assert.match(source, /rows\.length !== expectedIds\.length/);
  assert.match(source, /observed\[checkId\] !== "PASS"/);
  for (let index = 1; index <= 16; index++) {
    assert.match(source, new RegExp(`QA-0302C-${String(index).padStart(2, "0")}`));
  }
});

test("IA mutation surface is restricted to visibility, position, and hyperlinks", () => {
  const mutators = Array.from(
    source.matchAll(/\.(set[A-Z][A-Za-z0-9]*|delete|clear|addTable|copyFrom)\s*\(/g),
    match => match[1]
  );
  assert.deepEqual(
    Array.from(new Set(mutators)).sort(),
    ["setHyperlink", "setPosition", "setVisibility"]
  );
  assert.doesNotMatch(source, /\.setFormula|\.setFormulas|\.setValue|\.setValues/);
  assert.doesNotMatch(source, /\.addTable|\.delete\(|\.clear\(|\.resize\(|\.setName\(/);
  assert.match(source, /tableCountAfter !== tableCountBefore/);
  assert.match(source, /sheetCountAfter !== sheetCountBefore/);
});

test("accepted Phase 2C calculation and interaction sources remain byte-for-byte unchanged", () => {
  assert.equal(
    sha256(phase2CScriptUrl),
    "c6662a0911c67f570d2bdd5212808d7636d783a0879caa312bebb97d5d58b166"
  );
  assert.equal(
    sha256(phase2CModuleUrl),
    "b85ab143255cdeec5915e0e40586eb906be611c79313e623ddc7dc274a7976a5"
  );
});

test("IA Office Script avoids unsupported iterator patterns and repeated workbook reads in loops", () => {
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
