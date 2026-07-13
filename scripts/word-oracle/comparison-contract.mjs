import fs from "node:fs/promises";
import path from "node:path";

import {
  formatValidationIssues,
  isPortableRelativePath,
  resolveManifestPath,
  sha256File,
} from "./contract.mjs";

export const WORD_ORACLE_COMPARISON_SCHEMA_VERSION = 1;

const ID_PATTERN = /^[a-z0-9][a-z0-9._-]{0,127}$/;
const SHA256_PATTERN = /^[a-f0-9]{64}$/;

export const METRIC_THRESHOLDS = Object.freeze([
  Object.freeze({
    metric: "meanAbsoluteDiff",
    threshold: "maxMeanAbsoluteDiff",
  }),
  Object.freeze({
    metric: "rootMeanSquareDiff",
    threshold: "maxRootMeanSquareDiff",
  }),
  Object.freeze({
    metric: "mismatchRatio",
    threshold: "maxMismatchRatio",
  }),
  Object.freeze({
    metric: "layoutStructureDiff",
    threshold: "maxLayoutStructureDiff",
  }),
]);

const REFERENCE_ROOT_KEYS = new Set([
  "schemaVersion",
  "corpus",
  "sourceManifestSha256",
  "rasterizer",
  "references",
]);
const VIEWER_ROOT_KEYS = new Set([
  "schemaVersion",
  "corpus",
  "sourceManifestSha256",
  "renderer",
  "references",
]);
const REFERENCE_KEYS = new Set([
  "caseId",
  "referenceId",
  "provider",
  "captureProfileId",
  "rasterization",
  "source",
  "oraclePdf",
  "pages",
]);
const VIEWER_REFERENCE_KEYS = new Set(["caseId", "referenceId", "pages"]);
const REFERENCE_PAGE_KEYS = new Set([
  "pageNumber",
  "widthPoints",
  "heightPoints",
  "rotation",
  "imagePath",
  "imageSha256",
]);
const VIEWER_PAGE_KEYS = new Set([
  "pageNumber",
  "widthPoints",
  "heightPoints",
  "imagePath",
  "imageSha256",
]);

export const RENDERER_ENVIRONMENT_FIELDS = Object.freeze([
  "browser",
  "browserVersion",
  "platform",
  "hostPlatform",
  "hostArchitecture",
  "viewport",
  "deviceScaleFactor",
  "fontSetFingerprintSha256",
  "locale",
  "timezone",
]);

const RENDERER_ENVIRONMENT_KEYS = new Set(RENDERER_ENVIRONMENT_FIELDS);

export const METRIC_SETTING_FIELDS = Object.freeze([
  "comparisonWidth",
  "comparisonHeight",
  "tolerance",
  "inkThreshold",
  "verticalBands",
  "horizontalBands",
  "gridColumns",
  "gridRows",
]);

export const PAGE_GEOMETRY_THRESHOLDS = Object.freeze([
  "maxPageDimensionRelativeDiff",
  "maxPageAspectRatioRelativeDiff",
]);

function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function addIssue(issues, issuePath, message) {
  issues.push({ path: issuePath, message });
}

function requireRecord(issues, value, issuePath) {
  if (!isRecord(value)) {
    addIssue(issues, issuePath, "must be an object");
    return false;
  }
  return true;
}

function checkAllowedKeys(issues, value, issuePath, allowedKeys) {
  if (!isRecord(value)) {
    return;
  }
  for (const key of Object.keys(value)) {
    if (!allowedKeys.has(key)) {
      addIssue(issues, `${issuePath}/${key}`, "is not a supported property");
    }
  }
}

function requireString(issues, value, issuePath) {
  if (typeof value !== "string" || value.length === 0) {
    addIssue(issues, issuePath, "must be a non-empty string");
    return false;
  }
  return true;
}

function requireId(issues, value, issuePath) {
  if (!requireString(issues, value, issuePath)) {
    return;
  }
  if (!ID_PATTERN.test(value)) {
    addIssue(
      issues,
      issuePath,
      "must contain only lowercase letters, digits, dots, underscores, or hyphens"
    );
  }
}

function requireSha256(issues, value, issuePath) {
  if (!requireString(issues, value, issuePath)) {
    return;
  }
  if (!SHA256_PATTERN.test(value)) {
    addIssue(issues, issuePath, "must be a lowercase SHA-256 digest");
  }
}

function requirePortablePngPath(issues, value, issuePath) {
  if (!requireString(issues, value, issuePath)) {
    return;
  }
  if (!isPortableRelativePath(value)) {
    addIssue(
      issues,
      issuePath,
      "must be a normalized POSIX path relative to the manifest"
    );
  }
  if (!value.toLowerCase().endsWith(".png")) {
    addIssue(issues, issuePath, "must end with .png");
  }
}

function requirePortableTracePath(issues, value, issuePath, extension) {
  if (!requireString(issues, value, issuePath)) {
    return;
  }
  if (!isPortableRelativePath(value)) {
    addIssue(issues, issuePath, "must be a normalized POSIX relative path");
  }
  if (!value.toLowerCase().endsWith(extension)) {
    addIssue(issues, issuePath, `must end with ${extension}`);
  }
}

function requirePositiveInteger(issues, value, issuePath) {
  if (!Number.isInteger(value) || value < 1) {
    addIssue(issues, issuePath, "must be an integer greater than zero");
  }
}

function validateRendererEnvironment(
  issues,
  value,
  issuePath,
  options = {}
) {
  if (!requireRecord(issues, value, issuePath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    value,
    issuePath,
    options.allowRendererMetadata
      ? new Set(["name", "version", "build", ...RENDERER_ENVIRONMENT_FIELDS])
      : RENDERER_ENVIRONMENT_KEYS
  );
  for (const field of [
    "browser",
    "browserVersion",
    "platform",
    "hostPlatform",
    "hostArchitecture",
    "locale",
    "timezone",
  ]) {
    requireString(issues, value[field], `${issuePath}/${field}`);
  }
  if (requireString(issues, value.viewport, `${issuePath}/viewport`)) {
    const match = /^(\d+)x(\d+)$/.exec(value.viewport);
    const width = Number(match?.[1]);
    const height = Number(match?.[2]);
    if (
      !match ||
      width < 320 ||
      width > 16384 ||
      height < 320 ||
      height > 16384
    ) {
      addIssue(
        issues,
        `${issuePath}/viewport`,
        "must use WIDTHxHEIGHT with dimensions between 320 and 16384"
      );
    }
  }
  if (
    typeof value.deviceScaleFactor !== "number" ||
    !Number.isFinite(value.deviceScaleFactor) ||
    value.deviceScaleFactor <= 0 ||
    value.deviceScaleFactor > 8
  ) {
    addIssue(
      issues,
      `${issuePath}/deviceScaleFactor`,
      "must be a finite number greater than zero and at most 8"
    );
  }
  requireSha256(
    issues,
    value.fontSetFingerprintSha256,
    `${issuePath}/fontSetFingerprintSha256`
  );
}

function validateCorpus(issues, corpus, issuePath) {
  if (!requireRecord(issues, corpus, issuePath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    corpus,
    issuePath,
    new Set(["id", "revision", "description"])
  );
  requireId(issues, corpus.id, `${issuePath}/id`);
  requireString(issues, corpus.revision, `${issuePath}/revision`);
  if (corpus.description !== undefined) {
    requireString(issues, corpus.description, `${issuePath}/description`);
  }
}

function validatePage(
  issues,
  page,
  issuePath,
  allowedKeys,
  geometryRequired,
  rotationRequired
) {
  if (!requireRecord(issues, page, issuePath)) {
    return;
  }
  checkAllowedKeys(issues, page, issuePath, allowedKeys);
  requirePositiveInteger(issues, page.pageNumber, `${issuePath}/pageNumber`);
  requirePortablePngPath(issues, page.imagePath, `${issuePath}/imagePath`);
  requireSha256(issues, page.imageSha256, `${issuePath}/imageSha256`);

  if (geometryRequired) {
    for (const dimension of ["widthPoints", "heightPoints"]) {
      if (
        typeof page[dimension] !== "number" ||
        !Number.isFinite(page[dimension]) ||
        page[dimension] <= 0
      ) {
        addIssue(
          issues,
          `${issuePath}/${dimension}`,
          "must be a finite number greater than zero"
        );
      }
    }
  }
  if (rotationRequired) {
    if (![0, 90, 180, 270].includes(page.rotation)) {
      addIssue(
        issues,
        `${issuePath}/rotation`,
        "must be one of 0, 90, 180, or 270"
      );
    }
  }
}

function validatePages(
  issues,
  pages,
  issuePath,
  allowedKeys,
  geometryRequired,
  rotationRequired,
  allowEmpty = false
) {
  if (!Array.isArray(pages)) {
    addIssue(issues, issuePath, "must be an array");
    return;
  }
  if (pages.length === 0 && !allowEmpty) {
    addIssue(issues, issuePath, "must be a non-empty array");
    return;
  }
  const pageNumbers = new Map();
  pages.forEach((page, pageIndex) => {
    validatePage(
      issues,
      page,
      `${issuePath}/${pageIndex}`,
      allowedKeys,
      geometryRequired,
      rotationRequired
    );
    if (Number.isInteger(page?.pageNumber)) {
      if (pageNumbers.has(page.pageNumber)) {
        addIssue(
          issues,
          `${issuePath}/${pageIndex}/pageNumber`,
          `duplicates ${issuePath}/${pageNumbers.get(
            page.pageNumber
          )}/pageNumber`
        );
      } else {
        pageNumbers.set(page.pageNumber, pageIndex);
      }
    }
  });
}

function validateReferenceIdentity(issues, reference, issuePath) {
  requireId(issues, reference.caseId, `${issuePath}/caseId`);
  requireId(issues, reference.referenceId, `${issuePath}/referenceId`);
}

function checkReferenceIdentities(issues, references, issuePath) {
  const seen = new Map();
  references.forEach((reference, index) => {
    if (
      !isRecord(reference) ||
      typeof reference.caseId !== "string" ||
      typeof reference.referenceId !== "string"
    ) {
      return;
    }
    const identity = referenceKey(reference.caseId, reference.referenceId);
    if (seen.has(identity)) {
      addIssue(
        issues,
        `${issuePath}/${index}`,
        `duplicates ${issuePath}/${seen.get(identity)}`
      );
    } else {
      seen.set(identity, index);
    }
  });
}

export function validateReferencePagesManifest(manifest) {
  const issues = [];
  if (!requireRecord(issues, manifest, "/")) {
    return issues;
  }
  checkAllowedKeys(issues, manifest, "", REFERENCE_ROOT_KEYS);
  if (manifest.schemaVersion !== WORD_ORACLE_COMPARISON_SCHEMA_VERSION) {
    addIssue(
      issues,
      "/schemaVersion",
      `must equal ${WORD_ORACLE_COMPARISON_SCHEMA_VERSION}`
    );
  }
  validateCorpus(issues, manifest.corpus, "/corpus");
  requireSha256(issues, manifest.sourceManifestSha256, "/sourceManifestSha256");

  if (requireRecord(issues, manifest.rasterizer, "/rasterizer")) {
    checkAllowedKeys(
      issues,
      manifest.rasterizer,
      "/rasterizer",
      new Set(["name", "version"])
    );
    requireString(issues, manifest.rasterizer.name, "/rasterizer/name");
    requireString(issues, manifest.rasterizer.version, "/rasterizer/version");
  }

  if (!Array.isArray(manifest.references) || manifest.references.length === 0) {
    addIssue(issues, "/references", "must be a non-empty array");
    return issues;
  }

  manifest.references.forEach((reference, referenceIndex) => {
    const issuePath = `/references/${referenceIndex}`;
    if (!requireRecord(issues, reference, issuePath)) {
      return;
    }
    checkAllowedKeys(issues, reference, issuePath, REFERENCE_KEYS);
    validateReferenceIdentity(issues, reference, issuePath);
    requireId(
      issues,
      reference.captureProfileId,
      `${issuePath}/captureProfileId`
    );
    if (requireRecord(issues, reference.provider, `${issuePath}/provider`)) {
      checkAllowedKeys(
        issues,
        reference.provider,
        `${issuePath}/provider`,
        new Set(["id", "displayName"])
      );
      requireId(issues, reference.provider.id, `${issuePath}/provider/id`);
      requireString(
        issues,
        reference.provider.displayName,
        `${issuePath}/provider/displayName`
      );
    }
    if (
      requireRecord(
        issues,
        reference.rasterization,
        `${issuePath}/rasterization`
      )
    ) {
      checkAllowedKeys(
        issues,
        reference.rasterization,
        `${issuePath}/rasterization`,
        new Set(["dpi", "pageBox", "format", "background"])
      );
      if (
        !Number.isInteger(reference.rasterization.dpi) ||
        reference.rasterization.dpi < 72 ||
        reference.rasterization.dpi > 600
      ) {
        addIssue(
          issues,
          `${issuePath}/rasterization/dpi`,
          "must be an integer between 72 and 600"
        );
      }
      if (!["crop", "media"].includes(reference.rasterization.pageBox)) {
        addIssue(
          issues,
          `${issuePath}/rasterization/pageBox`,
          'must be "crop" or "media"'
        );
      }
      if (reference.rasterization.format !== "png") {
        addIssue(issues, `${issuePath}/rasterization/format`, 'must be "png"');
      }
      if (reference.rasterization.background !== "#ffffff") {
        addIssue(
          issues,
          `${issuePath}/rasterization/background`,
          'must be "#ffffff"'
        );
      }
    }
    if (requireRecord(issues, reference.source, `${issuePath}/source`)) {
      checkAllowedKeys(
        issues,
        reference.source,
        `${issuePath}/source`,
        new Set(["path", "sha256"])
      );
      requirePortableTracePath(
        issues,
        reference.source.path,
        `${issuePath}/source/path`,
        ".docx"
      );
      requireSha256(
        issues,
        reference.source.sha256,
        `${issuePath}/source/sha256`
      );
    }
    if (requireRecord(issues, reference.oraclePdf, `${issuePath}/oraclePdf`)) {
      checkAllowedKeys(
        issues,
        reference.oraclePdf,
        `${issuePath}/oraclePdf`,
        new Set(["path", "sha256"])
      );
      requirePortableTracePath(
        issues,
        reference.oraclePdf.path,
        `${issuePath}/oraclePdf/path`,
        ".pdf"
      );
      requireSha256(
        issues,
        reference.oraclePdf.sha256,
        `${issuePath}/oraclePdf/sha256`
      );
    }
    validatePages(
      issues,
      reference.pages,
      `${issuePath}/pages`,
      REFERENCE_PAGE_KEYS,
      true,
      true
    );
  });
  checkReferenceIdentities(issues, manifest.references, "/references");
  return issues;
}

export function validateViewerPagesManifest(manifest) {
  const issues = [];
  if (!requireRecord(issues, manifest, "/")) {
    return issues;
  }
  checkAllowedKeys(issues, manifest, "", VIEWER_ROOT_KEYS);
  if (manifest.schemaVersion !== WORD_ORACLE_COMPARISON_SCHEMA_VERSION) {
    addIssue(
      issues,
      "/schemaVersion",
      `must equal ${WORD_ORACLE_COMPARISON_SCHEMA_VERSION}`
    );
  }
  validateCorpus(issues, manifest.corpus, "/corpus");
  requireSha256(issues, manifest.sourceManifestSha256, "/sourceManifestSha256");

  if (requireRecord(issues, manifest.renderer, "/renderer")) {
    checkAllowedKeys(
      issues,
      manifest.renderer,
      "/renderer",
      new Set([
        "name",
        "version",
        "build",
        ...RENDERER_ENVIRONMENT_FIELDS,
      ])
    );
    requireString(issues, manifest.renderer.name, "/renderer/name");
    requireString(issues, manifest.renderer.version, "/renderer/version");
    if (manifest.renderer.build !== undefined) {
      requireString(issues, manifest.renderer.build, "/renderer/build");
    }
    validateRendererEnvironment(issues, manifest.renderer, "/renderer", {
      allowRendererMetadata: true,
    });
  }

  if (!Array.isArray(manifest.references)) {
    addIssue(issues, "/references", "must be an array");
    return issues;
  }
  manifest.references.forEach((reference, referenceIndex) => {
    const issuePath = `/references/${referenceIndex}`;
    if (!requireRecord(issues, reference, issuePath)) {
      return;
    }
    checkAllowedKeys(issues, reference, issuePath, VIEWER_REFERENCE_KEYS);
    validateReferenceIdentity(issues, reference, issuePath);
    validatePages(
      issues,
      reference.pages,
      `${issuePath}/pages`,
      VIEWER_PAGE_KEYS,
      true,
      false,
      true
    );
  });
  checkReferenceIdentities(issues, manifest.references, "/references");
  return issues;
}

export function validateThresholds(thresholds) {
  const issues = [];
  if (!requireRecord(issues, thresholds, "/")) {
    return issues;
  }
  checkAllowedKeys(
    issues,
    thresholds,
    "",
    new Set([
      "schemaVersion",
      "metricScriptSha256",
      "expectedRendererEnvironment",
      ...METRIC_SETTING_FIELDS,
      ...PAGE_GEOMETRY_THRESHOLDS,
      ...METRIC_THRESHOLDS.map(({ threshold }) => threshold),
    ])
  );
  if (thresholds.schemaVersion !== WORD_ORACLE_COMPARISON_SCHEMA_VERSION) {
    addIssue(
      issues,
      "/schemaVersion",
      `must equal ${WORD_ORACLE_COMPARISON_SCHEMA_VERSION}`
    );
  }
  requireSha256(issues, thresholds.metricScriptSha256, "/metricScriptSha256");
  validateRendererEnvironment(
    issues,
    thresholds.expectedRendererEnvironment,
    "/expectedRendererEnvironment"
  );
  for (const setting of METRIC_SETTING_FIELDS) {
    const value = thresholds[setting];
    const minimum = setting === "tolerance" || setting === "inkThreshold" ? 0 : 1;
    const maximum =
      setting === "tolerance" || setting === "inkThreshold" ? 255 : 16384;
    if (!Number.isInteger(value) || value < minimum || value > maximum) {
      addIssue(
        issues,
        `/${setting}`,
        `must be an integer between ${minimum} and ${maximum}`
      );
    }
  }
  for (const threshold of PAGE_GEOMETRY_THRESHOLDS) {
    const value = thresholds[threshold];
    if (
      typeof value !== "number" ||
      !Number.isFinite(value) ||
      value < 0 ||
      value > 1
    ) {
      addIssue(
        issues,
        `/${threshold}`,
        "must be a finite number between 0 and 1"
      );
    }
  }
  for (const { threshold } of METRIC_THRESHOLDS) {
    const value = thresholds[threshold];
    if (
      typeof value !== "number" ||
      !Number.isFinite(value) ||
      value < 0 ||
      value > 1
    ) {
      addIssue(
        issues,
        `/${threshold}`,
        "must be a finite number between 0 and 1"
      );
    }
  }
  for (const [countField, dimensionField] of [
    ["verticalBands", "comparisonHeight"],
    ["horizontalBands", "comparisonWidth"],
    ["gridColumns", "comparisonWidth"],
    ["gridRows", "comparisonHeight"],
  ]) {
    if (
      Number.isInteger(thresholds[countField]) &&
      Number.isInteger(thresholds[dimensionField]) &&
      thresholds[countField] > thresholds[dimensionField]
    ) {
      addIssue(
        issues,
        `/${countField}`,
        `must not exceed ${dimensionField}`
      );
    }
  }
  return issues;
}

async function readJsonWithValidator(filePath, label, validator) {
  let parsed;
  try {
    parsed = JSON.parse(await fs.readFile(filePath, "utf8"));
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`Unable to read ${label} ${filePath}: ${message}`);
  }
  const issues = validator(parsed);
  if (issues.length > 0) {
    throw new Error(
      `Invalid ${label} ${filePath}:\n${formatValidationIssues(issues)}`
    );
  }
  return parsed;
}

export function readReferencePagesManifest(filePath) {
  return readJsonWithValidator(
    filePath,
    "Word reference pages manifest",
    validateReferencePagesManifest
  );
}

export function readViewerPagesManifest(filePath) {
  return readJsonWithValidator(
    filePath,
    "viewer pages manifest",
    validateViewerPagesManifest
  );
}

export function readThresholds(filePath) {
  return readJsonWithValidator(
    filePath,
    "Word comparison thresholds",
    validateThresholds
  );
}

export function referenceKey(caseId, referenceId) {
  return `${caseId}\u0000${referenceId}`;
}

export function pageKey(caseId, referenceId, pageNumber) {
  return `${referenceKey(caseId, referenceId)}\u0000${pageNumber}`;
}

function compareIdentity(left, right) {
  return (
    left.caseId.localeCompare(right.caseId) ||
    left.referenceId.localeCompare(right.referenceId) ||
    (left.pageNumber ?? 0) - (right.pageNumber ?? 0)
  );
}

function sortedPageNumbers(pages) {
  return pages
    .map((page) => page.pageNumber)
    .sort((left, right) => left - right);
}

function sameNumbers(left, right) {
  return (
    left.length === right.length &&
    left.every((value, index) => value === right[index])
  );
}

export function buildComparisonPlan(referenceManifest, viewerManifest) {
  const referenceGroups = new Map(
    referenceManifest.references.map((reference) => [
      referenceKey(reference.caseId, reference.referenceId),
      reference,
    ])
  );
  const viewerGroups = new Map(
    viewerManifest.references.map((reference) => [
      referenceKey(reference.caseId, reference.referenceId),
      reference,
    ])
  );
  const identities = [
    ...new Set([...referenceGroups.keys(), ...viewerGroups.keys()]),
  ];
  const pageCounts = identities
    .map((identity) => {
      const reference = referenceGroups.get(identity);
      const viewer = viewerGroups.get(identity);
      const exemplar = reference ?? viewer;
      const expectedPageNumbers = sortedPageNumbers(reference?.pages ?? []);
      const actualPageNumbers = sortedPageNumbers(viewer?.pages ?? []);
      const passed = sameNumbers(expectedPageNumbers, actualPageNumbers);
      return {
        caseId: exemplar.caseId,
        referenceId: exemplar.referenceId,
        expectedPageCount: expectedPageNumbers.length,
        actualPageCount: actualPageNumbers.length,
        expectedPageNumbers,
        actualPageNumbers,
        passed,
      };
    })
    .sort(compareIdentity);

  const pageCountPassed = pageCounts.every((check) => check.passed);
  const pairs = [];
  if (pageCountPassed) {
    for (const reference of referenceManifest.references) {
      const viewer = viewerGroups.get(
        referenceKey(reference.caseId, reference.referenceId)
      );
      const viewerPages = new Map(
        viewer.pages.map((page) => [page.pageNumber, page])
      );
      for (const referencePage of reference.pages) {
        pairs.push({
          caseId: reference.caseId,
          referenceId: reference.referenceId,
          pageNumber: referencePage.pageNumber,
          groundTruth: referencePage,
          viewer: viewerPages.get(referencePage.pageNumber),
        });
      }
    }
    pairs.sort(compareIdentity);
  }

  return { pageCountPassed, pageCounts, pairs };
}

function roundedRatio(value) {
  return Math.round(value * 1_000_000) / 1_000_000;
}

function relativeDifference(expected, actual) {
  return Math.abs(actual - expected) / expected;
}

export function evaluatePageGeometry(plan, thresholds) {
  if (!plan.pageCountPassed) {
    return { passed: false, comparisons: [] };
  }
  const comparisons = plan.pairs.map((pair) => {
    const expectedWidth = pair.groundTruth.widthPoints;
    const expectedHeight = pair.groundTruth.heightPoints;
    const actualWidth = pair.viewer.widthPoints;
    const actualHeight = pair.viewer.heightPoints;
    const widthRelativeDiff = relativeDifference(expectedWidth, actualWidth);
    const heightRelativeDiff = relativeDifference(expectedHeight, actualHeight);
    const dimensionRelativeDiff = Math.max(
      widthRelativeDiff,
      heightRelativeDiff
    );
    const expectedAspectRatio = expectedWidth / expectedHeight;
    const actualAspectRatio = actualWidth / actualHeight;
    const aspectRatioRelativeDiff = relativeDifference(
      expectedAspectRatio,
      actualAspectRatio
    );
    const failures = [];
    if (dimensionRelativeDiff > thresholds.maxPageDimensionRelativeDiff) {
      failures.push({
        metric: "pageDimensionRelativeDiff",
        actual: roundedRatio(dimensionRelativeDiff),
        maximum: thresholds.maxPageDimensionRelativeDiff,
      });
    }
    if (aspectRatioRelativeDiff > thresholds.maxPageAspectRatioRelativeDiff) {
      failures.push({
        metric: "pageAspectRatioRelativeDiff",
        actual: roundedRatio(aspectRatioRelativeDiff),
        maximum: thresholds.maxPageAspectRatioRelativeDiff,
      });
    }
    return {
      caseId: pair.caseId,
      referenceId: pair.referenceId,
      pageNumber: pair.pageNumber,
      expected: {
        widthPoints: expectedWidth,
        heightPoints: expectedHeight,
        aspectRatio: roundedRatio(expectedAspectRatio),
      },
      actual: {
        widthPoints: actualWidth,
        heightPoints: actualHeight,
        aspectRatio: roundedRatio(actualAspectRatio),
      },
      widthRelativeDiff: roundedRatio(widthRelativeDiff),
      heightRelativeDiff: roundedRatio(heightRelativeDiff),
      dimensionRelativeDiff: roundedRatio(dimensionRelativeDiff),
      aspectRatioRelativeDiff: roundedRatio(aspectRatioRelativeDiff),
      passed: failures.length === 0,
      failures,
    };
  });
  return {
    passed: comparisons.every((comparison) => comparison.passed),
    comparisons,
  };
}

export function manifestsDescribeSameCorpus(referenceManifest, viewerManifest) {
  const issues = [];
  if (referenceManifest.corpus.id !== viewerManifest.corpus.id) {
    issues.push({
      path: "/corpus/id",
      message: `viewer value ${viewerManifest.corpus.id} does not match reference value ${referenceManifest.corpus.id}`,
    });
  }
  if (referenceManifest.corpus.revision !== viewerManifest.corpus.revision) {
    issues.push({
      path: "/corpus/revision",
      message: `viewer value ${viewerManifest.corpus.revision} does not match reference value ${referenceManifest.corpus.revision}`,
    });
  }
  if (
    referenceManifest.sourceManifestSha256 !==
    viewerManifest.sourceManifestSha256
  ) {
    issues.push({
      path: "/sourceManifestSha256",
      message:
        "viewer and reference pages were not derived from the same source manifest",
    });
  }
  return issues;
}

export function compareRendererEnvironment(viewerManifest, thresholds) {
  const issues = [];
  const expected = thresholds.expectedRendererEnvironment;
  for (const field of RENDERER_ENVIRONMENT_FIELDS) {
    if (viewerManifest.renderer[field] !== expected[field]) {
      issues.push({
        path: `/renderer/${field}`,
        message: `viewer value ${JSON.stringify(
          viewerManifest.renderer[field]
        )} does not match pinned value ${JSON.stringify(expected[field])}`,
      });
    }
  }
  return issues;
}

export function buildMetricPairs(
  plan,
  referenceManifestPath,
  viewerManifestPath
) {
  return plan.pairs.map((pair) => ({
    caseId: pair.caseId,
    referenceId: pair.referenceId,
    pageNumber: pair.pageNumber,
    groundTruthPath: resolveManifestPath(
      referenceManifestPath,
      pair.groundTruth.imagePath
    ),
    viewerPath: resolveManifestPath(viewerManifestPath, pair.viewer.imagePath),
    groundTruthSha256: pair.groundTruth.imageSha256,
    viewerSha256: pair.viewer.imageSha256,
  }));
}

export async function verifyComparisonArtifacts(
  plan,
  referenceManifestPath,
  viewerManifestPath
) {
  const issues = [];
  for (const pair of plan.pairs) {
    for (const artifact of [
      {
        kind: "groundTruth",
        manifestPath: referenceManifestPath,
        page: pair.groundTruth,
      },
      {
        kind: "viewer",
        manifestPath: viewerManifestPath,
        page: pair.viewer,
      },
    ]) {
      const issuePath = `/${pair.caseId}/${pair.referenceId}/${pair.pageNumber}/${artifact.kind}`;
      let resolved;
      try {
        resolved = resolveManifestPath(
          artifact.manifestPath,
          artifact.page.imagePath
        );
        const actualHash = await sha256File(resolved);
        if (actualHash !== artifact.page.imageSha256) {
          addIssue(
            issues,
            `${issuePath}/imageSha256`,
            `does not match ${artifact.page.imagePath}: expected ${artifact.page.imageSha256}, got ${actualHash}`
          );
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        addIssue(
          issues,
          `${issuePath}/imagePath`,
          `cannot be read: ${message}`
        );
      }
    }
  }
  return issues;
}

function requireMetricNumber(issues, value, issuePath) {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value < 0 ||
    value > 1
  ) {
    addIssue(issues, issuePath, "must be a finite number between 0 and 1");
  }
}

export function validateAndIndexMetrics(metricsOutput, expectedPairs, thresholds) {
  const issues = [];
  if (!requireRecord(issues, metricsOutput, "/")) {
    return { issues, settings: undefined, results: [] };
  }
  const settings = {};
  for (const setting of METRIC_SETTING_FIELDS) {
    const value = metricsOutput[setting];
    const minimum = setting === "tolerance" || setting === "inkThreshold" ? 0 : 1;
    if (!Number.isInteger(value) || value < minimum) {
      addIssue(
        issues,
        `/${setting}`,
        `must be an integer greater than or equal to ${minimum}`
      );
    } else {
      settings[setting] = value;
      if (thresholds && value !== thresholds[setting]) {
        addIssue(
          issues,
          `/${setting}`,
          `must equal the pinned value ${thresholds[setting]}`
        );
      }
    }
  }

  if (!Array.isArray(metricsOutput.results)) {
    addIssue(issues, "/results", "must be an array");
    return { issues, settings, results: [] };
  }

  const expected = new Map(
    expectedPairs.map((pair) => [
      pageKey(pair.caseId, pair.referenceId, pair.pageNumber),
      pair,
    ])
  );
  const seen = new Set();
  const normalized = [];
  metricsOutput.results.forEach((result, resultIndex) => {
    const issuePath = `/results/${resultIndex}`;
    if (!requireRecord(issues, result, issuePath)) {
      return;
    }
    const identity = pageKey(
      result.caseId,
      result.referenceId,
      result.pageNumber
    );
    const pair = expected.get(identity);
    if (!pair) {
      addIssue(
        issues,
        issuePath,
        "does not match an expected case/reference/page comparison"
      );
      return;
    }
    if (seen.has(identity)) {
      addIssue(issues, issuePath, "duplicates an earlier comparison result");
      return;
    }
    seen.add(identity);
    for (const hashField of ["groundTruthSha256", "viewerSha256"]) {
      if (result[hashField] !== pair[hashField]) {
        addIssue(
          issues,
          `${issuePath}/${hashField}`,
          "does not match the page manifest"
        );
      }
    }
    for (const { metric } of METRIC_THRESHOLDS) {
      requireMetricNumber(issues, result[metric], `${issuePath}/${metric}`);
    }
    normalized.push({
      caseId: pair.caseId,
      referenceId: pair.referenceId,
      pageNumber: pair.pageNumber,
      groundTruthSha256: pair.groundTruthSha256,
      viewerSha256: pair.viewerSha256,
      meanAbsoluteDiff: result.meanAbsoluteDiff,
      rootMeanSquareDiff: result.rootMeanSquareDiff,
      mismatchRatio: result.mismatchRatio,
      layoutStructureDiff: result.layoutStructureDiff,
    });
  });

  for (const [identity, pair] of expected) {
    if (!seen.has(identity)) {
      addIssue(
        issues,
        `/results`,
        `is missing ${pair.caseId}/${pair.referenceId}/page-${pair.pageNumber}`
      );
    }
  }
  normalized.sort(compareIdentity);
  return { issues, settings, results: normalized };
}

export function evaluateMetricThresholds(metricResults, thresholds) {
  const pages = metricResults.map((result) => {
    const failures = [];
    for (const { metric, threshold } of METRIC_THRESHOLDS) {
      if (result[metric] > thresholds[threshold]) {
        failures.push({
          metric,
          actual: result[metric],
          maximum: thresholds[threshold],
        });
      }
    }
    return { ...result, passed: failures.length === 0, failures };
  });

  const maxima = Object.fromEntries(
    METRIC_THRESHOLDS.map(({ metric }) => [
      metric,
      pages.length === 0
        ? null
        : Math.max(...pages.map((page) => page[metric])),
    ])
  );
  return {
    passed: pages.every((page) => page.passed),
    summary: {
      pagesCompared: pages.length,
      failedPages: pages.filter((page) => !page.passed).length,
      maxima,
    },
    pages,
  };
}

export async function writeJson(filePath, value) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  const temporaryPath = `${filePath}.tmp-${process.pid}`;
  await fs.writeFile(
    temporaryPath,
    `${JSON.stringify(value, null, 2)}\n`,
    "utf8"
  );
  await fs.rename(temporaryPath, filePath);
}

export { sha256File };
