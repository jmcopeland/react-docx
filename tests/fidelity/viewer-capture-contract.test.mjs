import assert from "node:assert/strict";
import { test } from "node:test";

import {
  advanceReadinessState,
  assertCapturePageUrl,
  assertRequestedCaptureEnvironment,
  buildCaptureTargets,
  buildMeasuredRendererEnvironment,
  buildViewerPagesManifest,
  parseCaptureBaseUrl,
  parseViewport,
  portablePagePath,
} from "../../scripts/word-oracle/viewer-capture-contract.mjs";

const ZERO_HASH = "0".repeat(64);
const ONE_HASH = "1".repeat(64);
const TWO_HASH = "2".repeat(64);

function sourceManifest() {
  return {
    corpus: { id: "word-fidelity", revision: "2026-07-09" },
    cases: [
      {
        id: "alpha",
        source: { path: "sources/alpha.docx", sha256: ZERO_HASH },
        references: [{ id: "word-a" }, { id: "word-b" }],
      },
      {
        id: "beta",
        source: { path: "sources/beta.docx", sha256: ONE_HASH },
        references: [{ id: "word-a" }],
      },
    ],
  };
}

function referencePagesManifest() {
  return {
    corpus: { id: "word-fidelity", revision: "2026-07-09" },
    references: [
      {
        caseId: "alpha",
        referenceId: "word-b",
        source: { path: "../sources/alpha.docx", sha256: ZERO_HASH },
      },
    ],
  };
}

function readinessSnapshot(overrides = {}) {
  return {
    loadedStatus: true,
    importError: "",
    loadingOverlayCount: 0,
    fontStatus: "loaded",
    imageCount: 1,
    pendingImageCount: 0,
    runningAnimationCount: 0,
    pages: [
      {
        index: 0,
        width: 816,
        height: 1056,
        surfaceWidth: 816,
        surfaceHeight: 1056,
        scrollWidth: 816,
        scrollHeight: 1056,
        visible: true,
        surfaceCount: 1,
        contentHash: "abc12345",
      },
    ],
    ...overrides,
  };
}

test("parses bounded WIDTHxHEIGHT viewport metadata", () => {
  assert.deepEqual(parseViewport("1440x1600"), {
    width: 1440,
    height: 1600,
  });
  assert.throws(() => parseViewport("1440,1600"), /WIDTHxHEIGHT/);
  assert.throws(() => parseViewport("120x1600"), /between 320 and 16384/);
});

test("capture URLs are loopback-only by default and cannot redirect origins", () => {
  assert.equal(
    parseCaptureBaseUrl("http://127.0.0.1:4173"),
    "http://127.0.0.1:4173/"
  );
  assert.equal(
    parseCaptureBaseUrl("http://localhost:4173/viewer"),
    "http://localhost:4173/viewer"
  );
  assert.throws(
    () => parseCaptureBaseUrl("https://example.com/viewer"),
    /loopback/
  );
  assert.equal(
    parseCaptureBaseUrl("https://example.com/viewer", { allowRemote: true }),
    "https://example.com/viewer"
  );
  assert.throws(
    () =>
      assertCapturePageUrl(
        "http://127.0.0.1:4173",
        "http://localhost:4173/viewer"
      ),
    /changed origin/
  );
});

test("builds one capture per source while retaining every reference identity", () => {
  assert.deepEqual(buildCaptureTargets(sourceManifest()), [
    {
      caseId: "alpha",
      source: { path: "sources/alpha.docx", sha256: ZERO_HASH },
      referenceIds: ["word-a", "word-b"],
    },
    {
      caseId: "beta",
      source: { path: "sources/beta.docx", sha256: ONE_HASH },
      referenceIds: ["word-a"],
    },
  ]);
});

test("materialized references constrain capture identities and bind source hashes", () => {
  assert.deepEqual(
    buildCaptureTargets(sourceManifest(), referencePagesManifest()),
    [
      {
        caseId: "alpha",
        source: { path: "sources/alpha.docx", sha256: ZERO_HASH },
        referenceIds: ["word-b"],
      },
    ]
  );

  const mismatched = referencePagesManifest();
  mismatched.references[0].source.sha256 = ONE_HASH;
  assert.throws(
    () => buildCaptureTargets(sourceManifest(), mismatched),
    /source digest does not match/
  );
});

test("unknown filters fail instead of silently producing a partial manifest", () => {
  assert.throws(
    () =>
      buildCaptureTargets(sourceManifest(), undefined, {
        caseIds: ["missing"],
      }),
    /No viewer capture targets/
  );
  assert.throws(
    () =>
      buildCaptureTargets(sourceManifest(), undefined, {
        referenceIds: ["missing"],
      }),
    /No viewer capture targets/
  );
});

test("readiness requires consecutive equal, fully ready geometry snapshots", () => {
  const first = advanceReadinessState(undefined, readinessSnapshot(), 3);
  assert.equal(first.ready, false);
  assert.equal(first.stableFrames, 1);

  const second = advanceReadinessState(first, readinessSnapshot(), 3);
  assert.equal(second.stableFrames, 2);

  const changed = readinessSnapshot({
    pages: [
      {
        ...readinessSnapshot().pages[0],
        surfaceHeight: 1060,
      },
    ],
  });
  const reset = advanceReadinessState(second, changed, 3);
  assert.equal(reset.ready, false);
  assert.equal(reset.stableFrames, 1);

  const ready = advanceReadinessState(
    advanceReadinessState(reset, changed, 3),
    changed,
    3
  );
  assert.equal(ready.ready, true);
  assert.equal(ready.stableFrames, 3);

  const pending = advanceReadinessState(
    ready,
    readinessSnapshot({ pendingImageCount: 1 }),
    3
  );
  assert.equal(pending.stableFrames, 0);
  assert.match(pending.issues[0], /image/);
});

test("constructs a comparison-compatible viewer-pages manifest", () => {
  const manifest = buildViewerPagesManifest({
    corpus: sourceManifest().corpus,
    sourceManifestSha256: ZERO_HASH,
    renderer: {
      name: "react-docx",
      version: "0.8.0",
      browser: "Mozilla/5.0 HeadlessChrome/138.0.0.0",
      browserVersion: "138.0.0.0",
      platform: "Linux x86_64",
      hostPlatform: "linux",
      hostArchitecture: "x64",
      viewport: "1440x1600",
      deviceScaleFactor: 1,
      fontSetFingerprintSha256: ONE_HASH,
      locale: "en-US",
      timezone: "UTC",
    },
    references: [
      {
        caseId: "alpha",
        referenceId: "word-a",
        pages: [
          {
            pageNumber: 1,
            widthPoints: 612,
            heightPoints: 792,
            imagePath: portablePagePath("alpha", "word-a", 1),
            imageSha256: TWO_HASH,
          },
        ],
      },
    ],
  });

  assert.equal(manifest.schemaVersion, 1);
  assert.equal(
    manifest.references[0].pages[0].imagePath,
    "alpha/word-a/page-0001.png"
  );
  assert.equal(manifest.renderer.deviceScaleFactor, 1);
});

test("builds renderer metadata only from measured runtime values", () => {
  const actual = buildMeasuredRendererEnvironment({
    browser: "measured user agent",
    browserVersion: "138.0.0.0",
    platform: "measured browser platform",
    hostPlatform: "linux",
    hostArchitecture: "x64",
    viewportWidth: 1440,
    viewportHeight: 1600,
    deviceScaleFactor: 1,
    fontSetFingerprintSha256: ONE_HASH,
    locale: "en-US",
    timezone: "UTC",
  });
  assert.deepEqual(actual, {
    browser: "measured user agent",
    browserVersion: "138.0.0.0",
    platform: "measured browser platform",
    hostPlatform: "linux",
    hostArchitecture: "x64",
    viewport: "1440x1600",
    deviceScaleFactor: 1,
    fontSetFingerprintSha256: ONE_HASH,
    locale: "en-US",
    timezone: "UTC",
  });
  assert.doesNotThrow(() =>
    assertRequestedCaptureEnvironment(actual, {
      viewport: { width: 1440, height: 1600 },
      deviceScaleFactor: 1,
      fontSetFingerprintSha256: ONE_HASH,
      locale: "en-US",
      timezone: "UTC",
    })
  );

  for (const override of [
    { viewport: { width: 1024, height: 768 } },
    { deviceScaleFactor: 2 },
    { fontSetFingerprintSha256: TWO_HASH },
    { locale: "fr-FR" },
    { timezone: "Pacific/Honolulu" },
  ]) {
    assert.throws(
      () =>
        assertRequestedCaptureEnvironment(actual, {
          viewport: { width: 1440, height: 1600 },
          deviceScaleFactor: 1,
          fontSetFingerprintSha256: ONE_HASH,
          locale: "en-US",
          timezone: "UTC",
          ...override,
        }),
      /did not honor/
    );
  }
});
