#!/usr/bin/env node

import { createHash } from "node:crypto";
import fs from "node:fs/promises";
import path from "node:path";
import { spawnSync } from "node:child_process";
import { fileURLToPath } from "node:url";

const SHA256_PATTERN = /^[a-f0-9]{64}$/;
const FONTCONFIG_FORMAT =
  "%{postscriptname}\t%{family}\t%{style}\t%{index}\t%{fontversion}\t%{file}\\n";

function sha256Bytes(value) {
  return createHash("sha256").update(value).digest("hex");
}

function parseFontconfigOutput(output) {
  const records = [];
  for (const [index, line] of output.split(/\r?\n/).entries()) {
    if (!line) {
      continue;
    }
    const fields = line.split("\t");
    if (fields.length !== 6 || fields[5].length === 0) {
      throw new Error(
        `fc-list returned an incomplete font record on line ${index + 1}`
      );
    }
    const [postScriptName, family, style, faceIndex, version, filePath] =
      fields;
    const fontIdentifier =
      postScriptName ||
      `fontconfig:${family || "unknown"}/${style || "unknown"}#${
        faceIndex || "0"
      }`;
    if (
      /[\r\n\t]/.test(fontIdentifier) ||
      /[\r\n\t]/.test(version) ||
      /[\r\n\t]/.test(filePath)
    ) {
      throw new Error(
        `fc-list returned an unsafe font record on line ${index + 1}`
      );
    }
    records.push({
      postScriptName: fontIdentifier,
      version: version || "unknown",
      filePath,
    });
  }
  if (records.length === 0) {
    throw new Error("fc-list did not report any installed fonts");
  }
  return records;
}

export async function fingerprintFontRecords(records, options = {}) {
  const readFile = options.readFile ?? fs.readFile;
  const fileDigests = new Map();
  const lines = new Set();

  for (const record of records) {
    if (
      typeof record?.postScriptName !== "string" ||
      record.postScriptName.length === 0 ||
      typeof record?.version !== "string" ||
      record.version.length === 0 ||
      typeof record?.filePath !== "string" ||
      record.filePath.length === 0
    ) {
      throw new Error("font inventory records require name, version, and path");
    }
    if (
      /[\r\n\t]/.test(record.postScriptName) ||
      /[\r\n\t]/.test(record.version)
    ) {
      throw new Error("font inventory names and versions must be single fields");
    }
    let fileSha256 = fileDigests.get(record.filePath);
    if (!fileSha256) {
      fileSha256 = sha256Bytes(await readFile(record.filePath));
      fileDigests.set(record.filePath, fileSha256);
    }
    lines.add(
      `${record.postScriptName}\t${record.version}\t${fileSha256}`
    );
  }

  const normalizedInventory = `${[...lines]
    .sort((left, right) =>
      Buffer.compare(Buffer.from(left, "utf8"), Buffer.from(right, "utf8"))
    )
    .join("\n")}\n`;
  return {
    fingerprintSha256: sha256Bytes(
      Buffer.from(normalizedInventory, "utf8")
    ),
    fontFaceCount: lines.size,
    fontFileCount: fileDigests.size,
    normalizedInventory,
  };
}

export async function computeInstalledFontFingerprint(options = {}) {
  const fcListBin = options.fcListBin ?? process.env.FC_LIST_BIN ?? "fc-list";
  const result = spawnSync(fcListBin, ["--format", FONTCONFIG_FORMAT], {
    encoding: "utf8",
    env: { ...process.env, LC_ALL: "C", LANG: "C" },
    maxBuffer: 64 * 1024 * 1024,
  });
  if (result.error) {
    throw new Error(`Unable to execute ${fcListBin}: ${result.error.message}`);
  }
  if (result.status !== 0) {
    throw new Error(
      `${fcListBin} exited with status ${result.status}: ${String(
        result.stderr ?? ""
      ).trim()}`
    );
  }
  return fingerprintFontRecords(parseFontconfigOutput(result.stdout));
}

export function assertExpectedFontFingerprint(actual, expected) {
  if (!SHA256_PATTERN.test(expected ?? "")) {
    throw new Error(
      "expected font fingerprint must be a lowercase 64-character SHA-256 digest"
    );
  }
  if (actual !== expected) {
    throw new Error(
      `Installed font fingerprint mismatch: expected ${expected}, got ${actual}`
    );
  }
}

function printUsage() {
  console.log(`Usage:
  node scripts/word-oracle/font-fingerprint.mjs [options]

Options:
  --expected <sha256>       Fail unless the computed fingerprint matches.
  --fc-list-bin <path>      Fontconfig fc-list override.
  --inventory-out <path>    Write the normalized, path-free inventory.
  --json                    Print structured metadata instead of only the hash.
  -h, --help                Show this help.`);
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
    expected: undefined,
    fcListBin: process.env.FC_LIST_BIN ?? "fc-list",
    inventoryOut: undefined,
    json: false,
  };
  for (let index = 0; index < argv.length; index += 1) {
    switch (argv[index]) {
      case "--expected":
        options.expected = nextValue(argv, index, "--expected");
        index += 1;
        break;
      case "--fc-list-bin":
        options.fcListBin = nextValue(argv, index, "--fc-list-bin");
        index += 1;
        break;
      case "--inventory-out":
        options.inventoryOut = path.resolve(
          nextValue(argv, index, "--inventory-out")
        );
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
        throw new Error(`Unknown argument: ${argv[index]}`);
    }
  }
  return options;
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  const result = await computeInstalledFontFingerprint({
    fcListBin: options.fcListBin,
  });
  if (options.expected !== undefined) {
    assertExpectedFontFingerprint(result.fingerprintSha256, options.expected);
  }
  if (options.inventoryOut) {
    await fs.mkdir(path.dirname(options.inventoryOut), { recursive: true });
    await fs.writeFile(
      options.inventoryOut,
      result.normalizedInventory,
      "utf8"
    );
  }
  if (options.json) {
    console.log(
      JSON.stringify({
        fingerprintSha256: result.fingerprintSha256,
        fontFaceCount: result.fontFaceCount,
        fontFileCount: result.fontFileCount,
      })
    );
  } else {
    console.log(result.fingerprintSha256);
  }
}

if (
  process.argv[1] &&
  path.resolve(process.argv[1]) === fileURLToPath(import.meta.url)
) {
  main().catch((error) => {
    console.error(error instanceof Error ? error.message : String(error));
    process.exit(1);
  });
}
