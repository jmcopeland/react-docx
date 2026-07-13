import assert from "node:assert/strict";
import { spawnSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";

import {
  buildComparisonPlan,
  compareRendererEnvironment,
  evaluateMetricThresholds,
  evaluatePageGeometry,
  validateAndIndexMetrics,
  validateReferencePagesManifest,
  validateThresholds,
  validateViewerPagesManifest,
} from "../../scripts/word-oracle/comparison-contract.mjs";
import { sha256File } from "../../scripts/word-oracle/contract.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "../..");
const compareCli = path.join(repoRoot, "scripts/word-oracle/compare-pages.mjs");
const SOURCE_MANIFEST_HASH = "a".repeat(64);
const SOURCE_HASH = "b".repeat(64);
const PDF_HASH = "c".repeat(64);
const METRIC_SCRIPT_HASH = "d".repeat(64);
const RENDERER_ENVIRONMENT = Object.freeze({
  browser:
    "Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 HeadlessChrome/138.0.0.0 Safari/537.36",
  browserVersion: "138.0.0.0",
  platform: "Linux x86_64",
  hostPlatform: "linux",
  hostArchitecture: "x64",
  viewport: "816x1056",
  deviceScaleFactor: 1,
  fontSetFingerprintSha256: "e".repeat(64),
  locale: "en-US",
  timezone: "UTC",
});

function referenceManifest(referenceImageHash, pageOverrides = {}) {
  return {
    schemaVersion: 1,
    corpus: { id: "word-fidelity", revision: "fixture-v1" },
    sourceManifestSha256: SOURCE_MANIFEST_HASH,
    rasterizer: { name: "pdftoppm", version: "fixture-1.0" },
    references: [
      {
        caseId: "simple-paragraph",
        referenceId: "word-win-current",
        provider: {
          id: "microsoft-word-desktop",
          displayName: "Microsoft Word Desktop",
        },
        captureProfileId: "final-print-v1",
        rasterization: {
          dpi: 144,
          pageBox: "crop",
          format: "png",
          background: "#ffffff",
        },
        source: { path: "sources/simple.docx", sha256: SOURCE_HASH },
        oraclePdf: { path: "oracles/simple.pdf", sha256: PDF_HASH },
        pages: [
          {
            pageNumber: 1,
            widthPoints: 612,
            heightPoints: 792,
            rotation: 0,
            imagePath: "images/reference-page-1.png",
            imageSha256: referenceImageHash,
            ...pageOverrides,
          },
        ],
      },
    ],
  };
}

function viewerManifest(viewerImageHash, pageOverrides = {}) {
  return {
    schemaVersion: 1,
    corpus: { id: "word-fidelity", revision: "fixture-v1" },
    sourceManifestSha256: SOURCE_MANIFEST_HASH,
    renderer: {
      name: "react-docx",
      version: "fixture",
      ...RENDERER_ENVIRONMENT,
    },
    references: [
      {
        caseId: "simple-paragraph",
        referenceId: "word-win-current",
        pages: [
          {
            pageNumber: 1,
            widthPoints: 612,
            heightPoints: 792,
            imagePath: "images/viewer-page-1.png",
            imageSha256: viewerImageHash,
            ...pageOverrides,
          },
        ],
      },
    ],
  };
}

function thresholds(overrides = {}) {
  return {
    schemaVersion: 1,
    metricScriptSha256: METRIC_SCRIPT_HASH,
    expectedRendererEnvironment: { ...RENDERER_ENVIRONMENT },
    comparisonWidth: 816,
    comparisonHeight: 1056,
    tolerance: 18,
    inkThreshold: 24,
    verticalBands: 12,
    horizontalBands: 8,
    gridColumns: 6,
    gridRows: 8,
    maxPageDimensionRelativeDiff: 0.01,
    maxPageAspectRatioRelativeDiff: 0.01,
    maxMeanAbsoluteDiff: 0.1,
    maxRootMeanSquareDiff: 0.15,
    maxMismatchRatio: 0.2,
    maxLayoutStructureDiff: 0.05,
    ...overrides,
  };
}

function metricOutput(referenceImageHash, viewerImageHash, overrides = {}) {
  return {
    comparisonWidth: 816,
    comparisonHeight: 1056,
    tolerance: 18,
    inkThreshold: 24,
    verticalBands: 12,
    horizontalBands: 8,
    gridColumns: 6,
    gridRows: 8,
    results: [
      {
        caseId: "simple-paragraph",
        referenceId: "word-win-current",
        pageNumber: 1,
        groundTruthSha256: referenceImageHash,
        viewerSha256: viewerImageHash,
        meanAbsoluteDiff: 0.02,
        rootMeanSquareDiff: 0.04,
        mismatchRatio: 0.06,
        layoutStructureDiff: 0.01,
        ...overrides,
      },
    ],
  };
}

async function writeJson(filePath, value) {
  await fs.mkdir(path.dirname(filePath), { recursive: true });
  await fs.writeFile(filePath, `${JSON.stringify(value, null, 2)}\n`, "utf8");
}

async function createFixture(metricOverrides = {}) {
  const root = await fs.mkdtemp(
    path.join(os.tmpdir(), "word-oracle-comparison-")
  );
  const referenceRoot = path.join(root, "reference");
  const viewerRoot = path.join(root, "viewer");
  const referenceImagePath = path.join(
    referenceRoot,
    "images/reference-page-1.png"
  );
  const viewerImagePath = path.join(viewerRoot, "images/viewer-page-1.png");
  await fs.mkdir(path.dirname(referenceImagePath), { recursive: true });
  await fs.mkdir(path.dirname(viewerImagePath), { recursive: true });
  // A hash-pinned live fixture comparator consumes only the pair metadata in
  // these contract tests, so binary image fixtures are unnecessary here.
  await fs.writeFile(referenceImagePath, "reference page fixture\n", "utf8");
  await fs.writeFile(viewerImagePath, "viewer page fixture\n", "utf8");
  const referenceImageHash = await sha256File(referenceImagePath);
  const viewerImageHash = await sha256File(viewerImagePath);

  const paths = {
    root,
    referenceManifest: path.join(referenceRoot, "reference-pages.json"),
    viewerManifest: path.join(viewerRoot, "viewer-pages.json"),
    thresholdFile: path.join(root, "thresholds.json"),
    metricScript: path.join(root, "fixture-metrics.mjs"),
    reportFile: path.join(root, "report.json"),
  };
  await fs.writeFile(
    paths.metricScript,
    `import fs from "node:fs";
const [pairsPath, outputPath, ...args] = process.argv.slice(2);
const value = (name) => Number(args[args.indexOf(name) + 1]);
const pairs = JSON.parse(fs.readFileSync(pairsPath, "utf8"));
const results = pairs.map((pair) => ({
  ...pair,
  meanAbsoluteDiff: ${metricOverrides.meanAbsoluteDiff ?? 0.02},
  rootMeanSquareDiff: ${metricOverrides.rootMeanSquareDiff ?? 0.04},
  mismatchRatio: ${metricOverrides.mismatchRatio ?? 0.06},
  layoutStructureDiff: ${metricOverrides.layoutStructureDiff ?? 0.01},
}));
fs.writeFileSync(outputPath, JSON.stringify({
  comparisonWidth: value("--width"),
  comparisonHeight: value("--height"),
  tolerance: value("--tolerance"),
  inkThreshold: value("--ink-threshold"),
  verticalBands: value("--vertical-bands"),
  horizontalBands: value("--horizontal-bands"),
  gridColumns: value("--grid-columns"),
  gridRows: value("--grid-rows"),
  results,
}) + "\\n");
`,
    "utf8"
  );
  const metricScriptSha256 = await sha256File(paths.metricScript);
  await writeJson(
    paths.referenceManifest,
    referenceManifest(referenceImageHash)
  );
  await writeJson(paths.viewerManifest, viewerManifest(viewerImageHash));
  await writeJson(
    paths.thresholdFile,
    thresholds({ metricScriptSha256 })
  );
  return {
    paths,
    referenceImageHash,
    viewerImageHash,
    metricScriptSha256,
  };
}

function runGate(paths, metricScript = paths.metricScript) {
  return spawnSync(
    process.execPath,
    [
      compareCli,
      "--references",
      paths.referenceManifest,
      "--viewer",
      paths.viewerManifest,
      "--thresholds",
      paths.thresholdFile,
      "--metric-script",
      metricScript,
      "--python-bin",
      process.execPath,
      "--out",
      paths.reportFile,
    ],
    { encoding: "utf8" }
  );
}

test("validates viewer, reference, and threshold comparison contracts", () => {
  const hash = "0".repeat(64);
  assert.deepEqual(validateReferencePagesManifest(referenceManifest(hash)), []);
  assert.deepEqual(validateViewerPagesManifest(viewerManifest(hash)), []);
  assert.deepEqual(validateThresholds(thresholds()), []);

  const emptyViewer = viewerManifest(hash);
  emptyViewer.references = [];
  assert.deepEqual(validateViewerPagesManifest(emptyViewer), []);

  const invalidViewer = viewerManifest(hash);
  invalidViewer.references.push(structuredClone(invalidViewer.references[0]));
  invalidViewer.references[0].pages[0].imagePath = "../escape.png";
  const issues = validateViewerPagesManifest(invalidViewer);
  assert.ok(issues.some((issue) => issue.message.includes("relative")));
  assert.ok(issues.some((issue) => issue.message.includes("duplicates")));
  const missingGeometry = viewerManifest(hash);
  delete missingGeometry.references[0].pages[0].widthPoints;
  assert.ok(
    validateViewerPagesManifest(missingGeometry).some(
      (issue) => issue.path.endsWith("/widthPoints")
    )
  );
  const missingRendererPin = viewerManifest(hash);
  delete missingRendererPin.renderer.locale;
  assert.ok(
    validateViewerPagesManifest(missingRendererPin).some(
      (issue) => issue.path === "/renderer/locale"
    )
  );
  const missingExpectedEnvironment = thresholds();
  delete missingExpectedEnvironment.expectedRendererEnvironment;
  assert.ok(
    validateThresholds(missingExpectedEnvironment).some(
      (issue) => issue.path === "/expectedRendererEnvironment"
    )
  );
  assert.ok(
    validateThresholds(thresholds({ maxMismatchRatio: 1.1 })).some(
      (issue) => issue.path === "/maxMismatchRatio"
    )
  );
});

test("rejects every unpinned renderer environment substitution", () => {
  const substitutions = {
    browser: "arbitrary browser",
    browserVersion: "999.0",
    platform: "arbitrary platform",
    hostPlatform: "arbitrary-host",
    hostArchitecture: "arbitrary-arch",
    viewport: "1024x768",
    deviceScaleFactor: 2,
    fontSetFingerprintSha256: "f".repeat(64),
    locale: "fr-FR",
    timezone: "Pacific/Honolulu",
  };

  for (const [field, value] of Object.entries(substitutions)) {
    const viewer = viewerManifest("0".repeat(64));
    viewer.renderer[field] = value;
    assert.deepEqual(validateViewerPagesManifest(viewer), []);
    assert.deepEqual(compareRendererEnvironment(viewer, thresholds()), [
      {
        path: `/renderer/${field}`,
        message: `viewer value ${JSON.stringify(
          value
        )} does not match pinned value ${JSON.stringify(
          RENDERER_ENVIRONMENT[field]
        )}`,
      },
    ]);
  }
});

test("fails physical page geometry before normalized pixel comparison", () => {
  const hash = "0".repeat(64);
  const references = referenceManifest(hash);
  const viewer = viewerManifest(hash);
  const matching = evaluatePageGeometry(
    buildComparisonPlan(references, viewer),
    thresholds()
  );
  assert.equal(matching.passed, true);

  viewer.references[0].pages[0].widthPoints = 792;
  viewer.references[0].pages[0].heightPoints = 612;
  const mismatched = evaluatePageGeometry(
    buildComparisonPlan(references, viewer),
    thresholds()
  );
  assert.equal(mismatched.passed, false);
  assert.deepEqual(
    mismatched.comparisons[0].failures.map((failure) => failure.metric),
    ["pageDimensionRelativeDiff", "pageAspectRatioRelativeDiff"]
  );
});

test("joins pages by case, reference, and page rather than array position", () => {
  const hash = "0".repeat(64);
  const references = referenceManifest(hash);
  references.references[0].pages.push({
    ...references.references[0].pages[0],
    pageNumber: 2,
    imagePath: "images/reference-page-2.png",
  });
  references.references[0].pages.reverse();
  const viewer = viewerManifest(hash);
  viewer.references[0].pages.push({
    ...viewer.references[0].pages[0],
    pageNumber: 2,
    imagePath: "images/viewer-page-2.png",
  });

  const plan = buildComparisonPlan(references, viewer);
  assert.equal(plan.pageCountPassed, true);
  assert.deepEqual(
    plan.pairs.map((pair) => pair.pageNumber),
    [1, 2]
  );
  assert.equal(
    plan.pairs[0].groundTruth.imagePath,
    "images/reference-page-1.png"
  );
  assert.equal(plan.pairs[0].viewer.imagePath, "images/viewer-page-1.png");
});

test("binds live metric results to page hashes and pinned settings", () => {
  const groundTruthSha256 = "1".repeat(64);
  const viewerSha256 = "2".repeat(64);
  const pairs = [
    {
      caseId: "simple-paragraph",
      referenceId: "word-win-current",
      pageNumber: 1,
      groundTruthSha256,
      viewerSha256,
    },
  ];
  const normalized = validateAndIndexMetrics(
    metricOutput(groundTruthSha256, viewerSha256),
    pairs,
    thresholds()
  );
  assert.deepEqual(normalized.issues, []);
  assert.equal(
    evaluateMetricThresholds(normalized.results, thresholds()).passed,
    true
  );

  const swapped = metricOutput(viewerSha256, groundTruthSha256);
  assert.ok(
    validateAndIndexMetrics(swapped, pairs, thresholds()).issues.some((issue) =>
      issue.path.endsWith("/groundTruthSha256")
    )
  );
  const wrongResolution = metricOutput(groundTruthSha256, viewerSha256);
  wrongResolution.comparisonWidth = 1;
  assert.ok(
    validateAndIndexMetrics(wrongResolution, pairs, thresholds()).issues.some(
      (issue) =>
        issue.path === "/comparisonWidth" && issue.message.includes("pinned")
    )
  );
});

test("CLI writes a traceable passing report from synthetic metadata", async () => {
  const fixture = await createFixture();
  try {
    const result = runGate(fixture.paths);
    assert.equal(result.status, 0, result.stderr);
    const report = JSON.parse(
      await fs.readFile(fixture.paths.reportFile, "utf8")
    );
    assert.equal(report.result, "passed");
    assert.equal(report.pageCount.passed, true);
    assert.equal(report.inputIntegrity.artifactsChecked, 2);
    assert.equal(report.inputIntegrity.status, "passed");
    assert.equal(report.pageGeometry.status, "passed");
    assert.equal(report.metrics.provenance.mode, "invoked");
    assert.equal(report.metrics.summary.pagesCompared, 1);
    assert.equal(
      report.metrics.pages[0].groundTruthSha256,
      fixture.referenceImageHash
    );
    assert.equal(report.metrics.pages[0].viewerSha256, fixture.viewerImageHash);
    assert.match(report.inputs.referencePages.sha256, /^[a-f0-9]{64}$/);
    assert.match(report.inputs.thresholds.sha256, /^[a-f0-9]{64}$/);
  } finally {
    await fs.rm(fixture.paths.root, { recursive: true, force: true });
  }
});

test("CLI rejects a structurally valid but unpinned renderer environment", async () => {
  const fixture = await createFixture();
  try {
    const viewer = JSON.parse(
      await fs.readFile(fixture.paths.viewerManifest, "utf8")
    );
    viewer.renderer.viewport = "1024x768";
    viewer.renderer.deviceScaleFactor = 2;
    viewer.renderer.fontSetFingerprintSha256 = "f".repeat(64);
    viewer.renderer.locale = "fr-FR";
    viewer.renderer.timezone = "Pacific/Honolulu";
    await writeJson(fixture.paths.viewerManifest, viewer);

    const result = runGate(fixture.paths);
    assert.equal(result.status, 1);
    const report = JSON.parse(
      await fs.readFile(fixture.paths.reportFile, "utf8")
    );
    assert.equal(report.rendererEnvironment.status, "failed");
    assert.deepEqual(
      report.rendererEnvironment.issues.map((issue) => issue.path),
      [
        "/renderer/viewport",
        "/renderer/deviceScaleFactor",
        "/renderer/fontSetFingerprintSha256",
        "/renderer/locale",
        "/renderer/timezone",
      ]
    );
    assert.equal(report.metrics.status, "not-run");
    assert.ok(
      report.failures.every(
        (failure) => failure.stage === "renderer-environment"
      )
    );
  } finally {
    await fs.rm(fixture.paths.root, { recursive: true, force: true });
  }
});

test("CLI rejects a metric process whose bytes are not pinned", async () => {
  const fixture = await createFixture();
  try {
    await fs.appendFile(fixture.paths.metricScript, "// changed\n", "utf8");
    const result = runGate(fixture.paths);
    assert.equal(result.status, 1);
    const report = JSON.parse(
      await fs.readFile(fixture.paths.reportFile, "utf8")
    );
    assert.equal(report.result, "failed");
    assert.equal(report.metrics.status, "execution-failed");
    assert.match(report.failures[0].message, /digest mismatch/);
  } finally {
    await fs.rm(fixture.paths.root, { recursive: true, force: true });
  }
});

test("CLI exits nonzero and identifies every hard metric threshold failure", async () => {
  const fixture = await createFixture({
    meanAbsoluteDiff: 0.11,
    rootMeanSquareDiff: 0.16,
    mismatchRatio: 0.21,
    layoutStructureDiff: 0.06,
  });
  try {
    const result = runGate(fixture.paths);
    assert.equal(result.status, 1);
    const report = JSON.parse(
      await fs.readFile(fixture.paths.reportFile, "utf8")
    );
    assert.equal(report.result, "failed");
    assert.equal(report.metrics.status, "failed");
    assert.equal(report.metrics.summary.failedPages, 1);
    assert.deepEqual(
      report.metrics.pages[0].failures.map((failure) => failure.metric),
      [
        "meanAbsoluteDiff",
        "rootMeanSquareDiff",
        "mismatchRatio",
        "layoutStructureDiff",
      ]
    );
  } finally {
    await fs.rm(fixture.paths.root, { recursive: true, force: true });
  }
});

test("CLI fails on page count before invoking the metric process", async () => {
  const fixture = await createFixture();
  try {
    const references = referenceManifest(fixture.referenceImageHash);
    references.references[0].pages.push({
      ...references.references[0].pages[0],
      pageNumber: 2,
      imagePath: "images/reference-page-2.png",
    });
    await writeJson(fixture.paths.referenceManifest, references);

    const missingMetricScript = path.join(
      fixture.paths.root,
      "does-not-exist.json"
    );
    const result = runGate(fixture.paths, missingMetricScript);
    assert.equal(result.status, 1);
    const report = JSON.parse(
      await fs.readFile(fixture.paths.reportFile, "utf8")
    );
    assert.equal(report.pageCount.passed, false);
    assert.equal(report.pageCount.comparisons[0].expectedPageCount, 2);
    assert.equal(report.pageCount.comparisons[0].actualPageCount, 1);
    assert.equal(report.inputIntegrity.status, "skipped-page-count-mismatch");
    assert.equal(report.metrics.status, "skipped-page-count-mismatch");
    assert.equal(report.failures[0].stage, "page-count");
  } finally {
    await fs.rm(fixture.paths.root, { recursive: true, force: true });
  }
});

test("CLI fails on page geometry before invoking the metric process", async () => {
  const fixture = await createFixture();
  try {
    const viewer = viewerManifest(fixture.viewerImageHash, {
      widthPoints: 792,
      heightPoints: 612,
    });
    await writeJson(fixture.paths.viewerManifest, viewer);
    const missingMetricScript = path.join(
      fixture.paths.root,
      "does-not-exist.mjs"
    );
    const result = runGate(fixture.paths, missingMetricScript);
    assert.equal(result.status, 1);
    const report = JSON.parse(
      await fs.readFile(fixture.paths.reportFile, "utf8")
    );
    assert.equal(report.pageGeometry.status, "failed");
    assert.equal(report.metrics.status, "skipped-page-geometry-mismatch");
    assert.equal(report.failures[0].stage, "page-geometry");
  } finally {
    await fs.rm(fixture.paths.root, { recursive: true, force: true });
  }
});

test("CLI rejects a page whose bytes do not match its manifest hash", async () => {
  const fixture = await createFixture();
  try {
    const viewer = viewerManifest(fixture.viewerImageHash);
    viewer.references[0].pages[0].imageSha256 = "f".repeat(64);
    await writeJson(fixture.paths.viewerManifest, viewer);
    const result = runGate(fixture.paths);
    assert.equal(result.status, 1);
    const report = JSON.parse(
      await fs.readFile(fixture.paths.reportFile, "utf8")
    );
    assert.equal(report.pageCount.passed, true);
    assert.equal(report.inputIntegrity.status, "failed");
    assert.equal(report.metrics.status, "skipped-input-integrity-failure");
    assert.equal(report.failures[0].stage, "input-integrity");
    assert.match(report.failures[0].message, /expected f{64}/);
  } finally {
    await fs.rm(fixture.paths.root, { recursive: true, force: true });
  }
});
