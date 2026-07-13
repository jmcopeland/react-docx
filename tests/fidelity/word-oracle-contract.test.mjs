import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";

import {
  isPortableRelativePath,
  parsePdfInfoOutput,
  resolveManifestPath,
  sha256File,
  validateManifest,
} from "../../scripts/word-oracle/contract.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "../..");
const ZERO_HASH = "0".repeat(64);
const ONE_HASH = "1".repeat(64);

function validManifest() {
  return {
    $schema: "../../scripts/word-oracle/manifest.schema.json",
    schemaVersion: 1,
    corpus: {
      id: "word-fidelity",
      revision: "2026-07-09",
      description: "Small deterministic Word reference corpus",
    },
    captureProfiles: [
      {
        id: "final-print-v1",
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
    cases: [
      {
        id: "mixed-sections",
        tags: ["sections", "pagination"],
        source: {
          path: "sources/mixed-sections.docx",
          sha256: ZERO_HASH,
        },
        references: [
          {
            id: "word-win-current",
            provider: {
              id: "microsoft-word-desktop",
              displayName: "Microsoft Word Desktop",
            },
            captureProfileId: "final-print-v1",
            capturedAt: "2026-07-09T16:00:00Z",
            artifact: {
              path: "oracles/mixed-sections-word.pdf",
              sha256: ONE_HASH,
              mediaType: "application/pdf",
              pageCount: 2,
              pages: [
                { widthPoints: 612, heightPoints: 792, rotation: 0 },
                { widthPoints: 792, heightPoints: 612, rotation: 90 },
              ],
            },
            environment: {
              renderer: {
                name: "Microsoft Word",
                version: "16.0",
                build: "19127.20154",
                channel: "current",
              },
              platform: {
                name: "Windows 11",
                version: "24H2",
                architecture: "x64",
              },
              locale: "en-US",
              timezone: "America/New_York",
              fontSet: {
                policy: "pinned",
                fingerprintSha256: ZERO_HASH,
              },
            },
          },
        ],
      },
    ],
  };
}

test("accepts a complete provider-neutral Word oracle manifest", () => {
  assert.deepEqual(validateManifest(validManifest()), []);
});

test("rejects ambiguous IDs, escaping paths, and unpinned capture inputs", () => {
  const manifest = validManifest();
  manifest.captureProfiles.push(structuredClone(manifest.captureProfiles[0]));
  manifest.cases[0].source.path = "../outside.docx";
  manifest.cases[0].references[0].captureProfileId = "missing-profile";
  delete manifest.cases[0].references[0].environment.fontSet.fingerprintSha256;

  const issues = validateManifest(manifest);
  assert.ok(
    issues.some(
      (issue) =>
        issue.path === "/captureProfiles/1/id" &&
        issue.message.includes("duplicates")
    )
  );
  assert.ok(
    issues.some(
      (issue) =>
        issue.path === "/cases/0/source/path" &&
        issue.message.includes("relative")
    )
  );
  assert.ok(
    issues.some(
      (issue) =>
        issue.path === "/cases/0/references/0/captureProfileId" &&
        issue.message.includes("missing-profile")
    )
  );
  assert.ok(
    issues.some(
      (issue) =>
        issue.path.endsWith("/fontSet/fingerprintSha256") &&
        issue.message.includes("required")
    )
  );
});

test("portable paths are POSIX, normalized, and manifest-relative", () => {
  assert.equal(isPortableRelativePath("sources/case one.docx"), true);
  assert.equal(isPortableRelativePath("../case.docx"), false);
  assert.equal(isPortableRelativePath("sources//case.docx"), false);
  assert.equal(isPortableRelativePath("sources\\case.docx"), false);
  assert.equal(isPortableRelativePath("/tmp/case.docx"), false);
  assert.equal(isPortableRelativePath("C:/temp/case.docx"), false);
  assert.equal(isPortableRelativePath("https://example.test/case.docx"), false);
});

test("manifest-relative inputs cannot escape through symlinks", async (t) => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "word-oracle-path-"));
  try {
    const corpusDir = path.join(tempDir, "corpus");
    const outsidePath = path.join(tempDir, "outside.docx");
    await fs.mkdir(corpusDir);
    await fs.writeFile(outsidePath, "private corpus bytes", "utf8");
    try {
      await fs.symlink(outsidePath, path.join(corpusDir, "linked.docx"));
    } catch (error) {
      if (error?.code === "EPERM" || error?.code === "EACCES") {
        t.skip("creating symlinks is not permitted on this platform");
        return;
      }
      throw error;
    }
    assert.throws(
      () =>
        resolveManifestPath(
          path.join(corpusDir, "word-oracle.json"),
          "linked.docx"
        ),
      /outside the manifest directory/
    );
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
});

test("parses single-page and multi-page pdfinfo geometry", () => {
  assert.deepEqual(
    parsePdfInfoOutput(`Pages:           1
Page size:       612 x 792 pts (letter)
Page rot:        0
`),
    {
      pageCount: 1,
      pages: [{ widthPoints: 612, heightPoints: 792, rotation: 0 }],
    }
  );

  assert.deepEqual(
    parsePdfInfoOutput(`Pages:           2
Page    1 size:  612 x 792 pts (letter)
Page    1 rot:   0
Page    2 size:  792 x 612 pts (letter)
Page    2 rot:   90
`),
    {
      pageCount: 2,
      pages: [
        { widthPoints: 612, heightPoints: 792, rotation: 0 },
        { widthPoints: 792, heightPoints: 612, rotation: 90 },
      ],
    }
  );
});

test("hashes artifacts without loading their complete contents into memory", async () => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "word-oracle-test-"));
  try {
    const artifactPath = path.join(tempDir, "artifact.txt");
    await fs.writeFile(artifactPath, "word oracle\n", "utf8");
    assert.equal(
      await sha256File(artifactPath),
      "34fa80f09e9d39d58942c3812353e84fed89108903521393a18e510a116ad9ad"
    );
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
});

test("metadata-only CLI validates without corpus files or credentials", async () => {
  const tempDir = await fs.mkdtemp(path.join(os.tmpdir(), "word-oracle-cli-"));
  try {
    const manifestPath = path.join(tempDir, "word-oracle.json");
    await fs.writeFile(
      manifestPath,
      `${JSON.stringify(validManifest(), null, 2)}\n`,
      "utf8"
    );
    const stdout = execFileSync(
      process.execPath,
      [
        path.join(repoRoot, "scripts/word-oracle/validate-manifest.mjs"),
        "--manifest",
        manifestPath,
        "--metadata-only",
        "--json",
      ],
      { encoding: "utf8" }
    );
    assert.deepEqual(JSON.parse(stdout), {
      valid: true,
      manifest: manifestPath,
      corpusId: "word-fidelity",
      corpusRevision: "2026-07-09",
      cases: 1,
      references: 1,
      verification: "metadata",
    });
  } finally {
    await fs.rm(tempDir, { recursive: true, force: true });
  }
});
