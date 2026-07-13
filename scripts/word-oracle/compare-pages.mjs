#!/usr/bin/env node

import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  buildComparisonPlan,
  buildMetricPairs,
  compareRendererEnvironment,
  evaluateMetricThresholds,
  evaluatePageGeometry,
  manifestsDescribeSameCorpus,
  readReferencePagesManifest,
  readThresholds,
  readViewerPagesManifest,
  sha256File,
  validateAndIndexMetrics,
  verifyComparisonArtifacts,
  writeJson,
} from "./comparison-contract.mjs";
import { formatValidationIssues, runCommand } from "./contract.mjs";

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const defaultMetricScript = path.resolve(
  scriptDirectory,
  "../measure_png_visual_diff.py"
);

function printUsage() {
  console.log(`Usage:
  node scripts/word-oracle/compare-pages.mjs \\
    --references <reference-pages.json> \\
    --viewer <viewer-pages.json> \\
    --thresholds <thresholds.json> \\
    --out <comparison-report.json> [options]

Options:
  --metric-script <path>   Metric script override. Defaults to
                           scripts/measure_png_visual_diff.py. Its SHA-256 must
                           match the value pinned in --thresholds.
  --python-bin <path>      Python executable. Defaults to PYTHON_BIN or python3.
  -h, --help               Show this help.

The command is offline and deterministic. Page sets must match exactly before
images are hashed or pixel/structure metrics are evaluated.`);
}

function nextValue(argv, index, option) {
  const value = argv[index + 1];
  if (!value || value.startsWith("--")) {
    throw new Error(`${option} requires a value`);
  }
  return value;
}

function parseArgs(argv) {
  const options = {
    referenceManifestPath: undefined,
    viewerManifestPath: undefined,
    thresholdsPath: undefined,
    outputPath: undefined,
    metricScriptPath: defaultMetricScript,
    pythonBin: process.env.PYTHON_BIN || "python3",
  };
  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    switch (arg) {
      case "--references":
        options.referenceManifestPath = path.resolve(
          nextValue(argv, index, "--references")
        );
        index += 1;
        break;
      case "--viewer":
        options.viewerManifestPath = path.resolve(
          nextValue(argv, index, "--viewer")
        );
        index += 1;
        break;
      case "--thresholds":
        options.thresholdsPath = path.resolve(
          nextValue(argv, index, "--thresholds")
        );
        index += 1;
        break;
      case "--out":
        options.outputPath = path.resolve(nextValue(argv, index, "--out"));
        index += 1;
        break;
      case "--metric-script":
        options.metricScriptPath = path.resolve(
          nextValue(argv, index, "--metric-script")
        );
        index += 1;
        break;
      case "--python-bin":
        options.pythonBin = nextValue(argv, index, "--python-bin");
        index += 1;
        break;
      case "-h":
      case "--help":
        printUsage();
        process.exit(0);
      default:
        throw new Error(`Unknown argument: ${arg}`);
    }
  }

  for (const [property, option] of [
    ["referenceManifestPath", "--references"],
    ["viewerManifestPath", "--viewer"],
    ["thresholdsPath", "--thresholds"],
    ["outputPath", "--out"],
  ]) {
    if (!options[property]) {
      throw new Error(`${option} is required`);
    }
  }
  return options;
}

function tracePath(outputPath, inputPath) {
  const relative = path.relative(path.dirname(outputPath), inputPath) || ".";
  return relative.split(path.sep).join("/");
}

async function readJson(filePath, label) {
  try {
    return JSON.parse(await fs.readFile(filePath, "utf8"));
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`Unable to read ${label} ${filePath}: ${message}`);
  }
}

function pythonVersion(pythonBin) {
  const result = runCommand(pythonBin, ["--version"]);
  return `${result.stdout || result.stderr}`.trim();
}

function pageCountFailures(plan) {
  return plan.pageCounts
    .filter((check) => !check.passed)
    .map((check) => ({
      stage: "page-count",
      caseId: check.caseId,
      referenceId: check.referenceId,
      message: `expected pages [${check.expectedPageNumbers.join(
        ", "
      )}], got [${check.actualPageNumbers.join(", ")}]`,
    }));
}

function baseReport({
  options,
  referenceManifest,
  viewerManifest,
  thresholds,
  inputHashes,
  plan,
}) {
  return {
    schemaVersion: 1,
    gate: "word-oracle-page-comparison",
    result: "failed",
    corpus: referenceManifest.corpus,
    inputs: {
      referencePages: {
        path: tracePath(options.outputPath, options.referenceManifestPath),
        sha256: inputHashes.referenceManifest,
        sourceManifestSha256: referenceManifest.sourceManifestSha256,
        rasterizer: referenceManifest.rasterizer,
      },
      viewerPages: {
        path: tracePath(options.outputPath, options.viewerManifestPath),
        sha256: inputHashes.viewerManifest,
        sourceManifestSha256: viewerManifest.sourceManifestSha256,
        renderer: viewerManifest.renderer,
      },
      thresholds: {
        path: tracePath(options.outputPath, options.thresholdsPath),
        sha256: inputHashes.thresholds,
        values: thresholds,
      },
    },
    pageCount: {
      requirement: "exact-page-number-set",
      passed: plan.pageCountPassed,
      comparisons: plan.pageCounts,
    },
    rendererEnvironment: {
      status: "passed",
      expected: thresholds.expectedRendererEnvironment,
      actual: Object.fromEntries(
        Object.keys(thresholds.expectedRendererEnvironment).map((field) => [
          field,
          viewerManifest.renderer[field],
        ])
      ),
      issues: [],
    },
    pageGeometry: {
      status: "not-run",
      thresholds: {
        maxPageDimensionRelativeDiff:
          thresholds.maxPageDimensionRelativeDiff,
        maxPageAspectRatioRelativeDiff:
          thresholds.maxPageAspectRatioRelativeDiff,
      },
      comparisons: [],
    },
    inputIntegrity: {
      status: "not-run",
      artifactsChecked: 0,
      issues: [],
    },
    metrics: {
      status: "not-run",
      provenance: null,
      settings: null,
      summary: null,
      pages: [],
    },
    failures: pageCountFailures(plan),
  };
}

async function writeReport(options, report) {
  await writeJson(options.outputPath, report);
  const summary =
    report.result === "passed"
      ? `Word oracle comparison passed: ${report.metrics.summary.pagesCompared} page(s).`
      : `Word oracle comparison failed: ${report.failures.length} failure(s).`;
  console.log(`${summary} Report: ${options.outputPath}`);
}

async function invokeMetricScript(options, metricPairs, thresholds) {
  const temporaryDirectory = await fs.mkdtemp(
    path.join(os.tmpdir(), "react-docx-word-oracle-")
  );
  try {
    const pairsPath = path.join(temporaryDirectory, "pairs.json");
    const outputPath = path.join(temporaryDirectory, "metrics.json");
    await writeJson(pairsPath, metricPairs);
    const scriptSha256 = await sha256File(options.metricScriptPath);
    if (scriptSha256 !== thresholds.metricScriptSha256) {
      throw new Error(
        `Metric script digest mismatch: expected ${thresholds.metricScriptSha256}, got ${scriptSha256}`
      );
    }
    runCommand(options.pythonBin, [
      options.metricScriptPath,
      pairsPath,
      outputPath,
      "--width",
      String(thresholds.comparisonWidth),
      "--height",
      String(thresholds.comparisonHeight),
      "--tolerance",
      String(thresholds.tolerance),
      "--ink-threshold",
      String(thresholds.inkThreshold),
      "--vertical-bands",
      String(thresholds.verticalBands),
      "--horizontal-bands",
      String(thresholds.horizontalBands),
      "--grid-columns",
      String(thresholds.gridColumns),
      "--grid-rows",
      String(thresholds.gridRows),
    ]);
    return {
      metricsOutput: await readJson(outputPath, "generated metric output"),
      provenance: {
        mode: "invoked",
        script: {
          path: tracePath(options.outputPath, options.metricScriptPath),
          sha256: scriptSha256,
        },
        runtime: {
          command: options.pythonBin,
          version: pythonVersion(options.pythonBin),
        },
      },
    };
  } finally {
    await fs.rm(temporaryDirectory, { recursive: true, force: true });
  }
}

function decorateMetricPages(evaluation, plan) {
  const pagePaths = new Map(
    plan.pairs.map((pair) => [
      `${pair.caseId}\u0000${pair.referenceId}\u0000${pair.pageNumber}`,
      {
        groundTruthImagePath: pair.groundTruth.imagePath,
        viewerImagePath: pair.viewer.imagePath,
      },
    ])
  );
  return evaluation.pages.map((page) => ({
    ...page,
    ...pagePaths.get(
      `${page.caseId}\u0000${page.referenceId}\u0000${page.pageNumber}`
    ),
  }));
}

function metricFailures(pages) {
  return pages.flatMap((page) =>
    page.failures.map((failure) => ({
      stage: "metrics",
      caseId: page.caseId,
      referenceId: page.referenceId,
      pageNumber: page.pageNumber,
      message: `${failure.metric} ${failure.actual} exceeds ${failure.maximum}`,
    }))
  );
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const [referenceManifest, viewerManifest, thresholds] = await Promise.all([
    readReferencePagesManifest(options.referenceManifestPath),
    readViewerPagesManifest(options.viewerManifestPath),
    readThresholds(options.thresholdsPath),
  ]);
  const inputHashes = {
    referenceManifest: await sha256File(options.referenceManifestPath),
    viewerManifest: await sha256File(options.viewerManifestPath),
    thresholds: await sha256File(options.thresholdsPath),
  };

  const identityIssues = manifestsDescribeSameCorpus(
    referenceManifest,
    viewerManifest
  );
  const rendererEnvironmentIssues = compareRendererEnvironment(
    viewerManifest,
    thresholds
  );
  if (identityIssues.length > 0 || rendererEnvironmentIssues.length > 0) {
    const report = {
      schemaVersion: 1,
      gate: "word-oracle-page-comparison",
      result: "failed",
      corpus: referenceManifest.corpus,
      inputs: {
        referencePages: {
          path: tracePath(options.outputPath, options.referenceManifestPath),
          sha256: inputHashes.referenceManifest,
        },
        viewerPages: {
          path: tracePath(options.outputPath, options.viewerManifestPath),
          sha256: inputHashes.viewerManifest,
        },
        thresholds: {
          path: tracePath(options.outputPath, options.thresholdsPath),
          sha256: inputHashes.thresholds,
          values: thresholds,
        },
      },
      pageCount: { status: "not-run", comparisons: [] },
      rendererEnvironment: {
        status:
          rendererEnvironmentIssues.length === 0 ? "passed" : "failed",
        expected: thresholds.expectedRendererEnvironment,
        actual: Object.fromEntries(
          Object.keys(thresholds.expectedRendererEnvironment).map((field) => [
            field,
            viewerManifest.renderer[field],
          ])
        ),
        issues: rendererEnvironmentIssues,
      },
      pageGeometry: { status: "not-run", comparisons: [] },
      inputIntegrity: { status: "not-run", artifactsChecked: 0, issues: [] },
      metrics: {
        status: "not-run",
        provenance: null,
        settings: null,
        summary: null,
        pages: [],
      },
      failures: [
        ...identityIssues.map((issue) => ({
          stage: "manifest-identity",
          path: issue.path,
          message: issue.message,
        })),
        ...rendererEnvironmentIssues.map((issue) => ({
          stage: "renderer-environment",
          path: issue.path,
          message: issue.message,
        })),
      ],
    };
    await writeReport(options, report);
    process.exitCode = 1;
    return;
  }

  const plan = buildComparisonPlan(referenceManifest, viewerManifest);
  const report = baseReport({
    options,
    referenceManifest,
    viewerManifest,
    thresholds,
    inputHashes,
    plan,
  });
  if (!plan.pageCountPassed) {
    report.pageGeometry.status = "skipped-page-count-mismatch";
    report.metrics.status = "skipped-page-count-mismatch";
    report.inputIntegrity.status = "skipped-page-count-mismatch";
    await writeReport(options, report);
    process.exitCode = 1;
    return;
  }

  const geometry = evaluatePageGeometry(plan, thresholds);
  report.pageGeometry = {
    status: geometry.passed ? "passed" : "failed",
    thresholds: {
      maxPageDimensionRelativeDiff:
        thresholds.maxPageDimensionRelativeDiff,
      maxPageAspectRatioRelativeDiff:
        thresholds.maxPageAspectRatioRelativeDiff,
    },
    comparisons: geometry.comparisons,
  };
  if (!geometry.passed) {
    report.inputIntegrity.status = "skipped-page-geometry-mismatch";
    report.metrics.status = "skipped-page-geometry-mismatch";
    report.failures = geometry.comparisons.flatMap((comparison) =>
      comparison.failures.map((failure) => ({
        stage: "page-geometry",
        caseId: comparison.caseId,
        referenceId: comparison.referenceId,
        pageNumber: comparison.pageNumber,
        message: `${failure.metric} ${failure.actual} exceeds ${failure.maximum}`,
      }))
    );
    await writeReport(options, report);
    process.exitCode = 1;
    return;
  }

  const artifactIssues = await verifyComparisonArtifacts(
    plan,
    options.referenceManifestPath,
    options.viewerManifestPath
  );
  report.inputIntegrity = {
    status: artifactIssues.length === 0 ? "passed" : "failed",
    artifactsChecked: plan.pairs.length * 2,
    issues: artifactIssues,
  };
  if (artifactIssues.length > 0) {
    report.metrics.status = "skipped-input-integrity-failure";
    report.failures = artifactIssues.map((issue) => ({
      stage: "input-integrity",
      path: issue.path,
      message: issue.message,
    }));
    await writeReport(options, report);
    process.exitCode = 1;
    return;
  }

  const metricPairs = buildMetricPairs(
    plan,
    options.referenceManifestPath,
    options.viewerManifestPath
  );
  let metricRun;
  try {
    metricRun = await invokeMetricScript(options, metricPairs, thresholds);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    report.metrics.status = "execution-failed";
    report.failures = [{ stage: "metric-execution", message }];
    await writeReport(options, report);
    process.exitCode = 1;
    return;
  }

  const normalized = validateAndIndexMetrics(
    metricRun.metricsOutput,
    metricPairs,
    thresholds
  );
  report.metrics.provenance = metricRun.provenance;
  report.metrics.settings = normalized.settings;
  if (normalized.issues.length > 0) {
    report.metrics.status = "invalid-output";
    report.failures = normalized.issues.map((issue) => ({
      stage: "metric-output",
      path: issue.path,
      message: issue.message,
    }));
    await writeReport(options, report);
    process.exitCode = 1;
    return;
  }

  const evaluation = evaluateMetricThresholds(normalized.results, thresholds);
  const pages = decorateMetricPages(evaluation, plan);
  report.metrics = {
    status: evaluation.passed ? "passed" : "failed",
    provenance: metricRun.provenance,
    settings: normalized.settings,
    summary: evaluation.summary,
    pages,
  };
  report.failures = metricFailures(pages);
  report.result = evaluation.passed ? "passed" : "failed";
  await writeReport(options, report);
  if (!evaluation.passed) {
    process.exitCode = 1;
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
