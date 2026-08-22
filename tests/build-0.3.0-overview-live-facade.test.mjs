import test from "node:test";
import assert from "node:assert/strict";
import { readFile } from "node:fs/promises";

const overviewUrl = new URL("../office-scripts/Build_0_3_0_Overview.ts", import.meta.url);
const performanceUrl = new URL("../office-scripts/Build_0_3_0_Weekly_Performance.ts", import.meta.url);
const documentationUrl = new URL("../docs/BUILD_0_3_0_OVERVIEW_LIVE_FACADE.md", import.meta.url);
const overview = await readFile(overviewUrl, "utf8");
const performance = await readFile(performanceUrl, "utf8");
const documentation = await readFile(documentationUrl, "utf8");
const facade = overview.slice(
  overview.indexOf("function buildOverviewFacade"),
  overview.indexOf("function mergeCard"),
);

test("Overview is a facade over explicit upstream management authority", () => {
  assert.match(performance, /writeManagementOverviewAuthority\(calc, layout\)/);
  assert.match(performance, /Management Performance availability/);
  assert.match(performance, /Selected Category Sales/);
  assert.match(performance, /writeGroupManagementRankings/);
  assert.match(performance, /writeRestaurantManagementRankings/);
  for (const cell of ["AL33", "AL34", "AL36", "AL41", "AL42", "AL43", "AL44", "AL45"]) {
    assert.match(facade, new RegExp(`_Metric_Calc'!\\$${cell.slice(0, 2)}\\$${cell.slice(2)}`));
  }
});

test("Overview contains no raw-fact, cache, aggregation, ranking, or period engine", () => {
  assert.doesNotMatch(facade, /tblSalesFacts|tblMetricRPGFacts|tblMetricRPGResults/);
  assert.doesNotMatch(facade, /tblWeeklyRPGCache|tblWeeklyScopeCache|tblWeeklyPeriodManifest/);
  assert.doesNotMatch(facade, /SUM\s*\(|SUMIFS|AVERAGE\s*\(|AVERAGEIFS|SORTBY\s*\(/);
  assert.doesNotMatch(facade, /COUNTIFS\s*\(|MAX\s*\(FILTER|SourcePeriodKey|CacheVersionID/);
  assert.doesNotMatch(facade, /addTable\s*\(|setRule\s*\(|getDataValidation\s*\(/);
  assert.match(overview, /Overview deliberately contains no business calculation engine/);
});

test("selected-scope metrics are calculated only in Performance authority", () => {
  assert.match(performance, /\$\{currentNumerator\}\/\$\{currentDenominator\}/);
  assert.match(performance, /\$\{comparisonNumerator\}\/\$\{comparisonDenominator\}/);
  assert.match(performance, /NOK Impact/);
  assert.match(overview, /rankingPresentationFormula/);
  assert.doesNotMatch(overview, /currentNumerator|currentDenominator|comparisonNumerator|comparisonDenominator/);
});

test("Top and Bottom 3 use Performance full-precision rankings and dynamic selection counts", () => {
  assert.match(performance, /SORTBY\(ids,--\(vals=""\),1,IF\(vals="",0,vals\)/);
  assert.match(performance, /MIN\(3,\$AL\$5\)/);
  assert.match(performance, /MIN\(3,\$AL\$4\)/);
  assert.match(performance, /layout\.groupCapacity/);
  assert.doesNotMatch(performance, /GROUP_CAPACITY\s*=\s*9|activeGroups\.length\s*!==\s*9/);
  assert.match(overview, /AP33/);
  assert.match(overview, /AQ38/);
  assert.match(overview, /AU33/);
  assert.match(overview, /AV38/);
});

test("stale Performance suppresses metrics and rankings upstream", () => {
  assert.match(performance, /\$AL\$32="Available"/);
  assert.match(performance, /\$AL\$32<>"Available"/);
  assert.match(performance, /Management Performance availability/);
  assert.match(overview, /Performance refresh required/);
  assert.match(overview, /Performance values are unavailable until refreshed/);
});

test("Overview projects current operational status and preserves compact navigation", () => {
  assert.match(facade, /Imports!\$A\$9/);
  assert.match(facade, /Imports!\$D\$9/);
  assert.match(facade, /SUBSTITUTE\(Imports!\$D\$9,CHAR\(10\)," · "\)/);
  assert.match(facade, /Imports!\$G\$9/);
  assert.match(facade, /Mapping!\$A\$9/);
  assert.match(facade, /Mapping!\$D\$8/);
  assert.match(facade, /Mapping!\$F\$6/);
  assert.match(facade, /A3[\s\S]*Performance/);
  assert.match(facade, /I3[\s\S]*Settings/);
  assert.match(facade, /A7[\s\S]*AL\$33/);
  assert.match(facade, /D7[\s\S]*AL\$34/);
  assert.match(facade, /G7[\s\S]*AL\$36/);
  assert.doesNotMatch(facade, /A33[\s\S]*Newer week available/);
  assert.doesNotMatch(facade, /WCV-|WCC-|MCF-|ICC-|QA-030/);
  assert.match(overview, /visibleSheetCount\s*!==\s*6/);
  assert.doesNotMatch(overview, /getWorksheets\(\)\.length\s*!==\s*47/);
});

test("final Overview presentation is compact and keeps recency only in Analysis Context", () => {
  assert.match(facade, /A3[\s\S]*C3[\s\S]*E3[\s\S]*G3[\s\S]*I3/);
  assert.match(facade, /A6[\s\S]*Current[\s\S]*D6[\s\S]*Compare[\s\S]*G6[\s\S]*Latest available/);
  assert.match(facade, /A8[\s\S]*AL\$25[\s\S]*D8[\s\S]*AL\$26[\s\S]*G8[\s\S]*AL\$37/);
  assert.match(facade, /A29[\s\S]*Attention[\s\S]*F29[\s\S]*Data Status/);
  assert.doesNotMatch(facade, /A3[0-4][\s\S]*Newer week available/);
  assert.match(overview, /getRange\("A1:J36"\)/);
});

test("authority documentation maps every management field to its owner", () => {
  for (const field of [
    "Current", "Compare", "Latest available", "Total Sales", "Selected Category Sales",
    "Sales Share", "PP Change", "NOK Impact", "Reporting Group Top/Bottom 3",
    "Restaurant Top/Bottom 3", "Mapping attention", "Data Status",
  ]) assert.match(documentation, new RegExp(field.replace("/", "\\/")));
  assert.match(documentation, /Overview never reads raw facts or weekly cache rows/);
  assert.match(documentation, /Performance authority/);
});

test("Overview and upstream extension remain Office Scripts compatible", () => {
  for (const source of [overview, performance]) {
    assert.match(source, /^function main\(/m);
    assert.doesNotMatch(source, /^(export|async|public|private|protected)\s+function\s+main/m);
    assert.doesNotMatch(source, /new\s+Map\s*</);
    assert.doesNotMatch(source, /new\s+Set\s*</);
    assert.doesNotMatch(source, /\.(entries|keys|values)\s*\(\)/);
    assert.doesNotMatch(source, /\.\.\./);
    assertBalanced(source);
  }
  const loopBodies = Array.from(overview.matchAll(/for\s*\([^)]*\)\s*\{([\s\S]*?)\n  \}/g), match => match[1]);
  for (const body of loopBodies) assert.doesNotMatch(body, /\.getValues\(|\.getTexts\(|workbook\.get|sheet\.get|table\.get/);
});

function assertBalanced(source) {
  for (const [open, close] of [["{", "}"], ["[", "]"]]) {
    assert.equal([...source].filter(value => value === open).length,
      [...source].filter(value => value === close).length, `${open}${close} balance`);
  }
}
