#!/usr/bin/env node

import fs from "node:fs/promises";
import path from "node:path";

import {
  commandVersion,
  formatValidationIssues,
  readManifest,
  resolveManifestPath,
  runCommand,
  sha256File,
  verifyManifestArtifacts,
} from "./contract.mjs";

function printUsage() {
  console.log(`Usage:
  node scripts/word-oracle/materialize-references.mjs \\
    --manifest <path> --out-dir <path> [options]

Options:
  --manifest <path>         Valid Word oracle manifest.
  --out-dir <path>          New directory tree for reference page PNGs.
  --case <id>               Include only this case. Can be repeated.
  --reference <id>          Include only this reference ID. Can be repeated.
  --pdfinfo-bin <path>      pdfinfo override. Defaults to PDFINFO_BIN or pdfinfo.
  --pdftoppm-bin <path>     pdftoppm override. Defaults to PDFTOPPM_BIN or pdftoppm.
  --force                   Replace selected generated reference directories/report.
  -h, --help                Show this help.

All inputs are local. This command verifies hashes and PDF geometry before rasterizing.`);
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
    manifestPath: undefined,
    outputDir: undefined,
    caseIds: [],
    referenceIds: [],
    pdfinfoBin: process.env.PDFINFO_BIN || "pdfinfo",
    pdftoppmBin: process.env.PDFTOPPM_BIN || "pdftoppm",
    force: false,
  };

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    switch (arg) {
      case "--manifest":
        options.manifestPath = path.resolve(
          nextValue(argv, index, "--manifest")
        );
        index += 1;
        break;
      case "--out-dir":
        options.outputDir = path.resolve(nextValue(argv, index, "--out-dir"));
        index += 1;
        break;
      case "--case":
        options.caseIds.push(nextValue(argv, index, "--case"));
        index += 1;
        break;
      case "--reference":
        options.referenceIds.push(nextValue(argv, index, "--reference"));
        index += 1;
        break;
      case "--pdfinfo-bin":
        options.pdfinfoBin = nextValue(argv, index, "--pdfinfo-bin");
        index += 1;
        break;
      case "--pdftoppm-bin":
        options.pdftoppmBin = nextValue(argv, index, "--pdftoppm-bin");
        index += 1;
        break;
      case "--force":
        options.force = true;
        break;
      case "-h":
      case "--help":
        printUsage();
        process.exit(0);
      default:
        throw new Error(`Unknown argument: ${arg}`);
    }
  }

  if (!options.manifestPath) {
    throw new Error("--manifest is required");
  }
  if (!options.outputDir) {
    throw new Error("--out-dir is required");
  }
  return options;
}

function selectCases(manifest, options) {
  const caseIdFilter = new Set(options.caseIds);
  const referenceIdFilter = new Set(options.referenceIds);
  const selectedCases = manifest.cases
    .filter(
      (testCase) => caseIdFilter.size === 0 || caseIdFilter.has(testCase.id)
    )
    .map((testCase) => ({
      ...testCase,
      references: testCase.references.filter(
        (reference) =>
          referenceIdFilter.size === 0 || referenceIdFilter.has(reference.id)
      ),
    }))
    .filter((testCase) => testCase.references.length > 0);

  if (selectedCases.length === 0) {
    throw new Error("No Word oracle references matched the supplied filters");
  }

  const selectedCaseIds = new Set(selectedCases.map((testCase) => testCase.id));
  const selectedReferenceIds = new Set(
    selectedCases.flatMap((testCase) =>
      testCase.references.map((reference) => reference.id)
    )
  );
  for (const caseId of caseIdFilter) {
    if (!selectedCaseIds.has(caseId)) {
      throw new Error(`Unknown or empty case selector: ${caseId}`);
    }
  }
  for (const referenceId of referenceIdFilter) {
    if (!selectedReferenceIds.has(referenceId)) {
      throw new Error(`Unknown reference selector: ${referenceId}`);
    }
  }
  return selectedCases;
}

async function pathExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

function numericPngSort(left, right) {
  const leftPage = Number(left.match(/-(\d+)\.png$/)?.[1]);
  const rightPage = Number(right.match(/-(\d+)\.png$/)?.[1]);
  return leftPage - rightPage;
}

async function rasterizeReference({
  artifactPath,
  artifact,
  outputDir,
  rasterization,
  pdftoppmBin,
  force,
}) {
  if (await pathExists(outputDir)) {
    if (!force) {
      throw new Error(
        `Generated reference directory already exists: ${outputDir}. Pass --force to replace it.`
      );
    }
    await fs.rm(outputDir, { recursive: true, force: true });
  }
  await fs.mkdir(outputDir, { recursive: true });

  const generatedPrefix = path.join(outputDir, "rendered");
  const args = [
    "-f",
    "1",
    "-l",
    String(artifact.pageCount),
    "-r",
    String(rasterization.dpi),
    "-png",
    "-aa",
    "yes",
    "-aaVector",
    "yes",
  ];
  if (rasterization.pageBox === "crop") {
    args.push("-cropbox");
  }
  args.push(artifactPath, generatedPrefix);
  runCommand(pdftoppmBin, args);

  const generatedFiles = (await fs.readdir(outputDir))
    .filter((fileName) => /^rendered-\d+\.png$/.test(fileName))
    .sort(numericPngSort);
  if (generatedFiles.length !== artifact.pageCount) {
    throw new Error(
      `pdftoppm produced ${generatedFiles.length} page(s), expected ${artifact.pageCount}`
    );
  }

  const pages = [];
  for (let index = 0; index < generatedFiles.length; index += 1) {
    const pageNumber = index + 1;
    const outputName = `page-${String(pageNumber).padStart(4, "0")}.png`;
    const outputPath = path.join(outputDir, outputName);
    await fs.rename(path.join(outputDir, generatedFiles[index]), outputPath);
    pages.push({
      pageNumber,
      ...artifact.pages[index],
      imageFile: outputName,
      imageSha256: await sha256File(outputPath),
    });
  }
  return pages;
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const manifest = await readManifest(options.manifestPath);
  const selectedCases = selectCases(manifest, options);
  const selectedManifest = { ...manifest, cases: selectedCases };
  const artifactIssues = await verifyManifestArtifacts(
    selectedManifest,
    options.manifestPath,
    {
      inspectPdfMetadata: true,
      pdfinfoBin: options.pdfinfoBin,
    }
  );
  if (artifactIssues.length > 0) {
    throw new Error(
      `Invalid Word oracle artifacts:\n${formatValidationIssues(
        artifactIssues
      )}`
    );
  }

  const reportPath = path.join(options.outputDir, "reference-pages.json");
  if ((await pathExists(reportPath)) && !options.force) {
    throw new Error(
      `Generated report already exists: ${reportPath}. Pass --force to replace it.`
    );
  }
  await fs.mkdir(options.outputDir, { recursive: true });

  const profiles = new Map(
    manifest.captureProfiles.map((profile) => [profile.id, profile])
  );
  const references = [];
  for (const testCase of selectedCases) {
    for (const reference of testCase.references) {
      const profile = profiles.get(reference.captureProfileId);
      if (profile.rasterization.background.toLowerCase() !== "#ffffff") {
        throw new Error(
          `Only a white raster background is currently supported: ${profile.rasterization.background}`
        );
      }
      const artifactPath = resolveManifestPath(
        options.manifestPath,
        reference.artifact.path
      );
      const relativeOutputDir = path.posix.join(testCase.id, reference.id);
      const absoluteOutputDir = path.join(
        options.outputDir,
        testCase.id,
        reference.id
      );
      const pages = await rasterizeReference({
        artifactPath,
        artifact: reference.artifact,
        outputDir: absoluteOutputDir,
        rasterization: profile.rasterization,
        pdftoppmBin: options.pdftoppmBin,
        force: options.force,
      });
      references.push({
        caseId: testCase.id,
        referenceId: reference.id,
        provider: reference.provider,
        captureProfileId: profile.id,
        rasterization: profile.rasterization,
        source: testCase.source,
        oraclePdf: {
          path: reference.artifact.path,
          sha256: reference.artifact.sha256,
        },
        pages: pages.map(({ imageFile, ...pageMetadata }) => ({
          ...pageMetadata,
          imagePath: path.posix.join(relativeOutputDir, imageFile),
        })),
      });
    }
  }

  const output = {
    schemaVersion: 1,
    corpus: manifest.corpus,
    sourceManifestSha256: await sha256File(options.manifestPath),
    rasterizer: {
      name: "pdftoppm",
      version: commandVersion(options.pdftoppmBin),
    },
    references,
  };
  await fs.writeFile(
    reportPath,
    `${JSON.stringify(output, null, 2)}\n`,
    "utf8"
  );
  console.log(
    `Materialized ${references.length} Word oracle reference(s) in ${options.outputDir}`
  );
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
