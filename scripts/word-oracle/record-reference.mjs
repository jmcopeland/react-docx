#!/usr/bin/env node

import fs from "node:fs/promises";
import { realpathSync } from "node:fs";
import path from "node:path";

import {
  formatValidationIssues,
  inspectPdf,
  sha256File,
  validateManifest,
} from "./contract.mjs";

function printUsage() {
  console.log(`Usage:
  node scripts/word-oracle/record-reference.mjs [options]

Required:
  --root <dir>                 Directory the final manifest will live in.
  --source <docx>              Frozen source DOCX.
  --pdf <pdf>                  Pre-generated Microsoft reference PDF.
  --case-id <id>               Stable corpus case ID.
  --reference-id <id>          Stable reference capture ID.
  --provider-id <id>           Adapter/provider ID, e.g. microsoft-word-desktop.
  --provider-name <name>       Human-readable provider name.
  --profile-id <id>            Capture profile used to produce the PDF.
  --captured-at <timestamp>    ISO 8601 UTC timestamp ending in Z.
  --renderer-name <name>       Renderer name, e.g. Microsoft Word.
  --renderer-version <version> Renderer or service API version.
  --locale <locale>            Capture locale, e.g. en-US.
  --timezone <timezone>        Capture timezone, e.g. America/New_York.
  --font-policy <policy>       pinned, document-embedded, or provider-managed.

Optional:
  --renderer-build <build>
  --renderer-channel <channel>
  --platform-name <name>
  --platform-version <version>
  --platform-architecture <architecture>
  --font-fingerprint <sha256>  Required for a pinned font policy.
  --font-description <text>
  --notes <text>
  --pdfinfo-bin <path>         Defaults to PDFINFO_BIN or pdfinfo.
  --output <path>              Write the case fragment instead of stdout.
  -h, --help

This command only inspects local files. It does not acquire credentials or call a provider.`);
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
    pdfinfoBin: process.env.PDFINFO_BIN || "pdfinfo",
  };
  const valueOptions = new Map([
    ["--root", "root"],
    ["--source", "source"],
    ["--pdf", "pdf"],
    ["--case-id", "caseId"],
    ["--reference-id", "referenceId"],
    ["--provider-id", "providerId"],
    ["--provider-name", "providerName"],
    ["--profile-id", "profileId"],
    ["--captured-at", "capturedAt"],
    ["--renderer-name", "rendererName"],
    ["--renderer-version", "rendererVersion"],
    ["--renderer-build", "rendererBuild"],
    ["--renderer-channel", "rendererChannel"],
    ["--platform-name", "platformName"],
    ["--platform-version", "platformVersion"],
    ["--platform-architecture", "platformArchitecture"],
    ["--locale", "locale"],
    ["--timezone", "timezone"],
    ["--font-policy", "fontPolicy"],
    ["--font-fingerprint", "fontFingerprint"],
    ["--font-description", "fontDescription"],
    ["--notes", "notes"],
    ["--pdfinfo-bin", "pdfinfoBin"],
    ["--output", "output"],
  ]);

  for (let index = 0; index < argv.length; index += 1) {
    const arg = argv[index];
    if (arg === "-h" || arg === "--help") {
      printUsage();
      process.exit(0);
    }
    const key = valueOptions.get(arg);
    if (!key) {
      throw new Error(`Unknown argument: ${arg}`);
    }
    options[key] = nextValue(argv, index, arg);
    index += 1;
  }

  const requiredKeys = [
    "root",
    "source",
    "pdf",
    "caseId",
    "referenceId",
    "providerId",
    "providerName",
    "profileId",
    "capturedAt",
    "rendererName",
    "rendererVersion",
    "locale",
    "timezone",
    "fontPolicy",
  ];
  for (const key of requiredKeys) {
    if (!options[key]) {
      throw new Error(`Missing required option for ${key}`);
    }
  }
  options.root = path.resolve(options.root);
  options.source = path.resolve(options.source);
  options.pdf = path.resolve(options.pdf);
  if (options.output) {
    options.output = path.resolve(options.output);
  }
  return options;
}

function portablePathWithinRoot(root, filePath) {
  const relativePath = path.relative(root, filePath);
  if (
    relativePath.length === 0 ||
    relativePath === ".." ||
    relativePath.startsWith(`..${path.sep}`) ||
    path.isAbsolute(relativePath)
  ) {
    throw new Error(`${filePath} must be a file beneath --root ${root}`);
  }
  const realRoot = realpathSync(root);
  const realFile = realpathSync(filePath);
  const realRelative = path.relative(realRoot, realFile);
  if (
    realRelative === ".." ||
    realRelative.startsWith(`..${path.sep}`) ||
    path.isAbsolute(realRelative)
  ) {
    throw new Error(
      `${filePath} resolves outside --root ${root} through a symlink`
    );
  }
  return relativePath.split(path.sep).join("/");
}

function compactObject(value) {
  return Object.fromEntries(
    Object.entries(value).filter(([, entryValue]) => entryValue !== undefined)
  );
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const sourcePath = portablePathWithinRoot(options.root, options.source);
  const pdfPath = portablePathWithinRoot(options.root, options.pdf);
  const pdf = inspectPdf(options.pdf, options.pdfinfoBin);

  const renderer = compactObject({
    name: options.rendererName,
    version: options.rendererVersion,
    build: options.rendererBuild,
    channel: options.rendererChannel,
  });
  const platform = options.platformName
    ? compactObject({
        name: options.platformName,
        version: options.platformVersion,
        architecture: options.platformArchitecture,
      })
    : undefined;
  if (
    !options.platformName &&
    (options.platformVersion || options.platformArchitecture)
  ) {
    throw new Error(
      "--platform-name is required when platform version or architecture is provided"
    );
  }
  const fontSet = compactObject({
    policy: options.fontPolicy,
    fingerprintSha256: options.fontFingerprint,
    description: options.fontDescription,
  });
  const reference = compactObject({
    id: options.referenceId,
    provider: {
      id: options.providerId,
      displayName: options.providerName,
    },
    captureProfileId: options.profileId,
    capturedAt: options.capturedAt,
    artifact: {
      path: pdfPath,
      sha256: await sha256File(options.pdf),
      mediaType: "application/pdf",
      pageCount: pdf.pageCount,
      pages: pdf.pages,
    },
    environment: compactObject({
      renderer,
      platform,
      locale: options.locale,
      timezone: options.timezone,
      fontSet,
    }),
    notes: options.notes,
  });
  const testCase = {
    id: options.caseId,
    source: {
      path: sourcePath,
      sha256: await sha256File(options.source),
    },
    references: [reference],
  };

  const validationManifest = {
    schemaVersion: 1,
    corpus: { id: "capture-validation", revision: "local" },
    captureProfiles: [
      {
        id: options.profileId,
        pdfExport: {
          mode: "print",
          markup: "final",
          updateFields: false,
          pdfa: false,
        },
        rasterization: {
          dpi: 144,
          pageBox: "crop",
          format: "png",
          background: "#ffffff",
        },
      },
    ],
    cases: [testCase],
  };
  const issues = validateManifest(validationManifest);
  if (issues.length > 0) {
    throw new Error(
      `Invalid capture metadata:\n${formatValidationIssues(issues)}`
    );
  }

  const json = `${JSON.stringify(testCase, null, 2)}\n`;
  if (options.output) {
    await fs.writeFile(options.output, json, "utf8");
    console.log(`Wrote Word oracle case fragment to ${options.output}`);
  } else {
    process.stdout.write(json);
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
