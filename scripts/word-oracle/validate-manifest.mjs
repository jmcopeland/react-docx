#!/usr/bin/env node

import path from "node:path";

import {
  formatValidationIssues,
  readManifest,
  verifyManifestArtifacts,
} from "./contract.mjs";

function printUsage() {
  console.log(`Usage:
  node scripts/word-oracle/validate-manifest.mjs --manifest <path> [options]

Options:
  --manifest <path>         Word oracle manifest to validate.
  --metadata-only           Validate manifest metadata without reading DOCX/PDF files.
  --skip-pdf-metadata       Verify file hashes, but do not invoke pdfinfo.
  --pdfinfo-bin <path>      pdfinfo executable override. Defaults to PDFINFO_BIN or pdfinfo.
  --json                    Print a machine-readable success result.
  -h, --help                Show this help.

The command is offline: it never calls Microsoft Graph or another provider.`);
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
    metadataOnly: false,
    inspectPdfMetadata: true,
    pdfinfoBin: process.env.PDFINFO_BIN || "pdfinfo",
    json: false,
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
      case "--metadata-only":
        options.metadataOnly = true;
        break;
      case "--skip-pdf-metadata":
        options.inspectPdfMetadata = false;
        break;
      case "--pdfinfo-bin":
        options.pdfinfoBin = nextValue(argv, index, "--pdfinfo-bin");
        index += 1;
        break;
      case "--json":
        options.json = true;
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
  return options;
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const manifest = await readManifest(options.manifestPath);
  let artifactIssues = [];
  if (!options.metadataOnly) {
    artifactIssues = await verifyManifestArtifacts(
      manifest,
      options.manifestPath,
      {
        inspectPdfMetadata: options.inspectPdfMetadata,
        pdfinfoBin: options.pdfinfoBin,
      }
    );
  }
  if (artifactIssues.length > 0) {
    throw new Error(
      `Invalid Word oracle artifacts:\n${formatValidationIssues(
        artifactIssues
      )}`
    );
  }

  const referenceCount = manifest.cases.reduce(
    (total, testCase) => total + testCase.references.length,
    0
  );
  const result = {
    valid: true,
    manifest: options.manifestPath,
    corpusId: manifest.corpus.id,
    corpusRevision: manifest.corpus.revision,
    cases: manifest.cases.length,
    references: referenceCount,
    verification: options.metadataOnly
      ? "metadata"
      : options.inspectPdfMetadata
      ? "metadata+hashes+pdf"
      : "metadata+hashes",
  };

  if (options.json) {
    console.log(JSON.stringify(result, null, 2));
  } else {
    console.log(
      `Valid Word oracle manifest: ${result.cases} case(s), ${result.references} reference(s), ${result.verification}.`
    );
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
