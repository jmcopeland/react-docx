import path from "node:path";

import { validateViewerPagesManifest } from "./comparison-contract.mjs";
import { formatValidationIssues } from "./contract.mjs";

const SHA256_PATTERN = /^[a-f0-9]{64}$/;
const VIEWPORT_PATTERN = /^(\d+)x(\d+)$/;

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

function isLoopbackHostname(hostname) {
  const normalized = hostname.toLowerCase();
  if (
    normalized === "localhost" ||
    normalized === "[::1]" ||
    normalized === "::1"
  ) {
    return true;
  }
  const ipv4 = /^(\d{1,3})\.(\d{1,3})\.(\d{1,3})\.(\d{1,3})$/.exec(
    normalized
  );
  return Boolean(
    ipv4 &&
      ipv4.slice(1).every((part) => Number(part) <= 255) &&
      Number(ipv4[1]) === 127
  );
}

export function parseCaptureBaseUrl(value, options = {}) {
  let parsed;
  try {
    parsed = new URL(value);
  } catch {
    throw new Error("base URL must be an absolute HTTP(S) URL");
  }
  if (!new Set(["http:", "https:"]).has(parsed.protocol)) {
    throw new Error("base URL must be an absolute HTTP(S) URL");
  }
  if (!options.allowRemote && !isLoopbackHostname(parsed.hostname)) {
    throw new Error(
      "base URL must use a loopback host unless --allow-remote is supplied"
    );
  }
  return parsed.href;
}

export function assertCapturePageUrl(expectedBaseUrl, actualUrl, options = {}) {
  const expected = new URL(
    parseCaptureBaseUrl(expectedBaseUrl, { allowRemote: options.allowRemote })
  );
  const actual = new URL(
    parseCaptureBaseUrl(actualUrl, { allowRemote: options.allowRemote })
  );
  if (actual.origin !== expected.origin) {
    throw new Error(
      `playground navigation changed origin from ${expected.origin} to ${actual.origin}`
    );
  }
  return actual.href;
}

function compareIdentity(left, right) {
  return (
    left.caseId.localeCompare(right.caseId) ||
    left.referenceId.localeCompare(right.referenceId)
  );
}

function referenceKey(caseId, referenceId) {
  return `${caseId}\u0000${referenceId}`;
}

function assertSha256(value, label) {
  if (!SHA256_PATTERN.test(value ?? "")) {
    throw new Error(`${label} must be a lowercase 64-character SHA-256 digest`);
  }
}

export function parseViewport(value) {
  const match = VIEWPORT_PATTERN.exec(value ?? "");
  if (!match) {
    throw new Error("viewport must use WIDTHxHEIGHT, for example 1440x1600");
  }
  const width = Number(match[1]);
  const height = Number(match[2]);
  if (width < 320 || width > 16384 || height < 320 || height > 16384) {
    throw new Error("viewport width and height must be between 320 and 16384");
  }
  return { width, height };
}

function nonEmptyMeasuredString(value, label) {
  if (typeof value !== "string" || value.length === 0) {
    throw new Error(`${label} was not reported by the capture runtime`);
  }
  return value;
}

export function buildMeasuredRendererEnvironment({
  browser,
  browserVersion,
  platform,
  hostPlatform,
  hostArchitecture,
  viewportWidth,
  viewportHeight,
  deviceScaleFactor,
  fontSetFingerprintSha256,
  locale,
  timezone,
}) {
  if (
    !Number.isInteger(viewportWidth) ||
    !Number.isInteger(viewportHeight)
  ) {
    throw new Error("browser viewport dimensions must be measured integers");
  }
  const viewport = `${viewportWidth}x${viewportHeight}`;
  parseViewport(viewport);
  if (
    typeof deviceScaleFactor !== "number" ||
    !Number.isFinite(deviceScaleFactor) ||
    deviceScaleFactor <= 0
  ) {
    throw new Error(
      "browser device scale factor must be a measured positive number"
    );
  }
  assertSha256(
    fontSetFingerprintSha256,
    "measured fontSetFingerprintSha256"
  );
  return {
    browser: nonEmptyMeasuredString(browser, "browser user agent"),
    browserVersion: nonEmptyMeasuredString(
      browserVersion,
      "browser binary version"
    ),
    platform: nonEmptyMeasuredString(platform, "browser platform"),
    hostPlatform: nonEmptyMeasuredString(hostPlatform, "host platform"),
    hostArchitecture: nonEmptyMeasuredString(
      hostArchitecture,
      "host architecture"
    ),
    viewport,
    deviceScaleFactor,
    fontSetFingerprintSha256,
    locale: nonEmptyMeasuredString(locale, "browser locale"),
    timezone: nonEmptyMeasuredString(timezone, "browser timezone"),
  };
}

export function assertRequestedCaptureEnvironment(actual, requested) {
  const expected = {
    viewport: `${requested.viewport.width}x${requested.viewport.height}`,
    deviceScaleFactor: requested.deviceScaleFactor,
    locale: requested.locale,
    timezone: requested.timezone,
    fontSetFingerprintSha256: requested.fontSetFingerprintSha256,
  };
  const mismatches = Object.entries(expected).filter(
    ([field, value]) => actual[field] !== value
  );
  if (mismatches.length > 0) {
    throw new Error(
      `Capture runtime did not honor the requested environment: ${mismatches
        .map(
          ([field, value]) =>
            `${field} expected ${JSON.stringify(value)}, measured ${JSON.stringify(
              actual[field]
            )}`
        )
        .join("; ")}`
    );
  }
}

function selectedReferenceRecords(sourceManifest, referencePagesManifest) {
  const sourceCases = new Map(
    sourceManifest.cases.map((testCase) => [testCase.id, testCase])
  );
  if (!referencePagesManifest) {
    return sourceManifest.cases.flatMap((testCase) =>
      testCase.references.map((reference) => ({
        caseId: testCase.id,
        referenceId: reference.id,
        source: testCase.source,
      }))
    );
  }

  if (referencePagesManifest.corpus.id !== sourceManifest.corpus.id) {
    throw new Error(
      `reference-pages corpus ${referencePagesManifest.corpus.id} does not match source corpus ${sourceManifest.corpus.id}`
    );
  }
  if (
    referencePagesManifest.corpus.revision !== sourceManifest.corpus.revision
  ) {
    throw new Error(
      `reference-pages revision ${referencePagesManifest.corpus.revision} does not match source revision ${sourceManifest.corpus.revision}`
    );
  }

  return referencePagesManifest.references.map((reference) => {
    const testCase = sourceCases.get(reference.caseId);
    if (!testCase) {
      throw new Error(
        `reference-pages selects unknown source case: ${reference.caseId}`
      );
    }
    if (
      !testCase.references.some(
        (candidate) => candidate.id === reference.referenceId
      )
    ) {
      throw new Error(
        `reference-pages selects unknown source reference: ${reference.caseId}/${reference.referenceId}`
      );
    }
    if (reference.source.sha256 !== testCase.source.sha256) {
      throw new Error(
        `reference-pages source digest does not match ${reference.caseId}`
      );
    }
    return {
      caseId: reference.caseId,
      referenceId: reference.referenceId,
      source: testCase.source,
    };
  });
}

export function buildCaptureTargets(
  sourceManifest,
  referencePagesManifest,
  options = {}
) {
  const caseFilter = new Set(options.caseIds ?? []);
  const referenceFilter = new Set(options.referenceIds ?? []);
  const available = selectedReferenceRecords(
    sourceManifest,
    referencePagesManifest
  );
  const selected = available.filter(
    (reference) =>
      (caseFilter.size === 0 || caseFilter.has(reference.caseId)) &&
      (referenceFilter.size === 0 ||
        referenceFilter.has(reference.referenceId))
  );

  if (selected.length === 0) {
    throw new Error("No viewer capture targets matched the supplied selectors");
  }
  for (const caseId of caseFilter) {
    if (!selected.some((reference) => reference.caseId === caseId)) {
      throw new Error(`Unknown or empty case selector: ${caseId}`);
    }
  }
  for (const referenceId of referenceFilter) {
    if (
      !selected.some((reference) => reference.referenceId === referenceId)
    ) {
      throw new Error(`Unknown reference selector: ${referenceId}`);
    }
  }

  const identities = new Set();
  for (const reference of selected) {
    const identity = referenceKey(reference.caseId, reference.referenceId);
    if (identities.has(identity)) {
      throw new Error(
        `Duplicate viewer capture target: ${reference.caseId}/${reference.referenceId}`
      );
    }
    identities.add(identity);
  }

  const groups = new Map();
  for (const reference of selected.sort(compareIdentity)) {
    const current = groups.get(reference.caseId);
    if (current) {
      current.referenceIds.push(reference.referenceId);
    } else {
      groups.set(reference.caseId, {
        caseId: reference.caseId,
        source: reference.source,
        referenceIds: [reference.referenceId],
      });
    }
  }
  return [...groups.values()];
}

function roundedMetric(value) {
  return Number.isFinite(value) ? Math.round(value * 1000) / 1000 : null;
}

export function readinessIssues(snapshot) {
  const issues = [];
  if (snapshot.loadedStatus !== true) {
    issues.push("the playground has not reported the imported file as loaded");
  }
  if (snapshot.importError) {
    issues.push(`the viewer reported an import error: ${snapshot.importError}`);
  }
  if (snapshot.loadingOverlayCount !== 0) {
    issues.push("the viewer loading overlay is still present");
  }
  if (snapshot.fontStatus !== "loaded") {
    issues.push(`document fonts are ${snapshot.fontStatus ?? "not ready"}`);
  }
  if (snapshot.pendingImageCount !== 0) {
    issues.push(`${snapshot.pendingImageCount} viewer image(s) are not ready`);
  }
  if (snapshot.runningAnimationCount !== 0) {
    issues.push(
      `${snapshot.runningAnimationCount} viewer animation(s) are still running`
    );
  }
  if (!Array.isArray(snapshot.pages) || snapshot.pages.length === 0) {
    issues.push("the viewer has no rendered pages");
    return issues;
  }

  const indexes = snapshot.pages.map((page) => page.index);
  const expectedIndexes = snapshot.pages.map((_, index) => index);
  if (
    indexes.length !== expectedIndexes.length ||
    indexes.some((value, index) => value !== expectedIndexes[index])
  ) {
    issues.push("rendered page indexes must be contiguous and zero-based");
  }
  snapshot.pages.forEach((page, index) => {
    if (
      !Number.isFinite(page.width) ||
      page.width <= 0 ||
      !Number.isFinite(page.height) ||
      page.height <= 0 ||
      !Number.isFinite(page.surfaceWidth) ||
      page.surfaceWidth <= 0 ||
      !Number.isFinite(page.surfaceHeight) ||
      page.surfaceHeight <= 0
    ) {
      issues.push(`page ${index + 1} has invalid geometry`);
    }
    if (page.visible !== true) {
      issues.push(`page ${index + 1} is not visible`);
    }
    if (page.surfaceCount !== 1) {
      issues.push(`page ${index + 1} must contain exactly one page surface`);
    }
  });
  return issues;
}

export function readinessSignature(snapshot) {
  return JSON.stringify({
    fontStatus: snapshot.fontStatus,
    imageCount: snapshot.imageCount,
    pages: snapshot.pages.map((page) => ({
      index: page.index,
      width: roundedMetric(page.width),
      height: roundedMetric(page.height),
      surfaceWidth: roundedMetric(page.surfaceWidth),
      surfaceHeight: roundedMetric(page.surfaceHeight),
      scrollWidth: roundedMetric(page.scrollWidth),
      scrollHeight: roundedMetric(page.scrollHeight),
      contentHash: page.contentHash,
    })),
  });
}

export function advanceReadinessState(
  previousState,
  snapshot,
  requiredStableFrames
) {
  if (!Number.isInteger(requiredStableFrames) || requiredStableFrames < 2) {
    throw new Error("requiredStableFrames must be an integer of at least 2");
  }
  const issues = readinessIssues(snapshot);
  if (issues.length > 0) {
    return { signature: null, stableFrames: 0, ready: false, issues };
  }
  const signature = readinessSignature(snapshot);
  const stableFrames =
    previousState?.signature === signature
      ? previousState.stableFrames + 1
      : 1;
  return {
    signature,
    stableFrames,
    ready: stableFrames >= requiredStableFrames,
    issues: [],
  };
}

export function buildViewerPagesManifest({
  corpus,
  sourceManifestSha256,
  renderer,
  references,
}) {
  assertSha256(sourceManifestSha256, "sourceManifestSha256");
  if (renderer.fontSetFingerprintSha256 !== undefined) {
    assertSha256(
      renderer.fontSetFingerprintSha256,
      "renderer.fontSetFingerprintSha256"
    );
  }
  const manifest = {
    schemaVersion: 1,
    corpus: structuredClone(corpus),
    sourceManifestSha256,
    renderer: structuredClone(renderer),
    references: references
      .map((reference) => ({
        caseId: reference.caseId,
        referenceId: reference.referenceId,
        pages: reference.pages
          .map((page) => ({ ...page }))
          .sort((left, right) => left.pageNumber - right.pageNumber),
      }))
      .sort(compareIdentity),
  };
  const issues = validateViewerPagesManifest(manifest);
  if (issues.length > 0) {
    throw new Error(
      `Invalid generated viewer-pages manifest:\n${formatValidationIssues(
        issues
      )}`
    );
  }
  return manifest;
}

export function portablePagePath(caseId, referenceId, pageNumber) {
  return path.posix.join(
    caseId,
    referenceId,
    `page-${String(pageNumber).padStart(4, "0")}.png`
  );
}
