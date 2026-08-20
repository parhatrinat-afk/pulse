#!/usr/bin/env node

/**
 * Read-only 85-week Mapping-attention evidence/materialization payload audit.
 * The caller supplies exact corpus and catalog paths; this command never
 * searches for workbooks and never writes source or workbook files.
 */

import { createHash } from "node:crypto";
import { readFile } from "node:fs/promises";
import path from "node:path";
import { buildCandidateWeeklyCache } from "../imports/weekly-compact-cache.mjs";
import { buildWeeklyIdentityPreflight } from "../imports/weekly-identity-preflight.mjs";
import {
  listWeeklyXlsxFiles,
  readWeeklyWorkbookMatrix,
} from "../imports/weekly-sales-ooxml.mjs";
import { parseWeeklySalesMatrix } from "../imports/weekly-sales-parser.mjs";
import {
  buildWeeklyMappingAttentionMaterializationPlan,
  weeklyMappingAttentionChunk,
} from "./weekly-mapping-attention-materialization.mjs";

const root = exactPath(process.argv[2], "Corpus path");
const catalogPath = exactPath(flag("--catalog"), "Catalog path");
const authorityPath = exactPath(flag("--authority"), "Active authority evidence path");
const files = await listWeeklyXlsxFiles(root);
if (!files.length) throw new Error("PUL-030MA-301: Exact corpus path contains no .xlsx reports.");
const reports = [];
for (const file of files) {
  const bytes = await readFile(file);
  const workbook = readWeeklyWorkbookMatrix(file);
  reports.push(parseWeeklySalesMatrix(workbook.matrix, {
    sourceLocator: path.relative(root, file),
    sourceBinaryFingerprint: `SHA256-${createHash("sha256").update(bytes).digest("hex")}`,
  }));
}
const catalogs = JSON.parse(await readFile(catalogPath, "utf8"));
const authorityEvidence = JSON.parse(await readFile(authorityPath, "utf8"));
const preflight = buildWeeklyIdentityPreflight({ parsedReports: reports, catalogs });
const cache = buildCandidateWeeklyCache({
  parsedReports: reports,
  catalogs,
  expectedMappingContentFingerprint: preflight.mappingContentFingerprint,
  expectedIdentityPreflightFingerprint: preflight.fingerprints.preflightFingerprint,
});
cache.versionManifest = {
  ...cache.versionManifest,
  cacheVersion: authorityEvidence.candidate.cacheVersion,
  cacheFingerprint: authorityEvidence.candidate.cacheFingerprint,
  sourceCorpusFingerprint: authorityEvidence.candidate.sourceCorpusFingerprint,
  identityPreflightFingerprint: authorityEvidence.candidate.identityPreflightFingerprint,
  cacheStatus: "Active",
  activationState: "Active",
};
const plan = buildWeeklyMappingAttentionMaterializationPlan(cache);
const sectionName = flag("--section");
if (sectionName) {
  const chunk = weeklyMappingAttentionChunk(
    plan,
    sectionName,
    optionalInteger("--offset", 0),
    optionalInteger("--limit", 400),
  );
  process.stdout.write(`${JSON.stringify(chunk)}\n`);
} else {
  process.stdout.write(`${JSON.stringify({
    readOnly: true,
    reportCount: reports.length,
    projectionFingerprint: plan.projectionFingerprint,
    activeCacheVersion: plan.activeCacheVersion,
    activeCacheFingerprint: plan.activeCacheFingerprint,
    throughPeriodLabel: plan.throughPeriodLabel,
    stateCoverage: plan.stateCoverage,
    rows: {
      control: plan.sections.control.rowCount,
      products: plan.sections.products.rowCount,
    },
    preparePayload: identityPayload(plan),
    finalizePayload: identityPayload(plan),
  }, null, 2)}\n`);
}

function identityPayload(plan) {
  return {
    projectionFingerprint: plan.projectionFingerprint,
    activeCacheVersion: plan.activeCacheVersion,
    activeCacheFingerprint: plan.activeCacheFingerprint,
  };
}

function flag(name) {
  const index = process.argv.indexOf(name);
  if (index < 0) return "";
  const value = process.argv[index + 1];
  if (!value || value.startsWith("--")) throw new Error(`${name} requires a value.`);
  return value;
}

function exactPath(value, label) {
  if (!value || value.startsWith("--")) throw new Error(`${label} must be supplied exactly.`);
  return path.resolve(value);
}

function optionalInteger(name, fallback) {
  const value = flag(name);
  if (!value) return fallback;
  const result = Number(value);
  if (!Number.isInteger(result) || result < 0) throw new Error(`${name} must be a non-negative integer.`);
  return result;
}
