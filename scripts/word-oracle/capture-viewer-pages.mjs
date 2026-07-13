#!/usr/bin/env node

import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import {
  readReferencePagesManifest,
  writeJson,
} from "./comparison-contract.mjs";
import {
  readManifest,
  resolveManifestPath,
  sha256File,
} from "./contract.mjs";
import {
  assertExpectedFontFingerprint,
  computeInstalledFontFingerprint,
} from "./font-fingerprint.mjs";
import {
  assertCapturePageUrl,
  assertRequestedCaptureEnvironment,
  buildMeasuredRendererEnvironment,
  buildCaptureTargets,
  buildViewerPagesManifest,
  parseCaptureBaseUrl,
  parseViewport,
  portablePagePath,
} from "./viewer-capture-contract.mjs";

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const repositoryRoot = path.resolve(scriptDirectory, "../..");
const DEFAULT_VIEWPORT = "1440x1600";
const DEFAULT_TIMEOUT_MS = 120_000;
const DEFAULT_STABLE_FRAMES = 8;
const SHA256_PATTERN = /^[a-f0-9]{64}$/;

function printUsage() {
  console.log(`Usage:
  node scripts/word-oracle/capture-viewer-pages.mjs \\
    --manifest <word-oracle.json> \\
    --base-url <running-playground-url> \\
    --out-dir <path> \\
    --font-fingerprint <sha256> [options]

Options:
  --references <path>       Materialized reference-pages.json. When supplied,
                            capture only its case/reference identities.
  --case <id>               Include only this case. Can be repeated.
  --reference <id>          Include only this reference ID. Can be repeated.
  --browser <name>          chromium, firefox, or webkit. Default: chromium.
  --viewport <WxH>          Browser viewport. Default: ${DEFAULT_VIEWPORT}.
  --device-scale-factor <n> Screenshot DPR. Default: 1.
  --font-fingerprint <hash> Expected SHA-256 of the installed font inventory.
                            Capture recomputes it with fc-list and fails on a
                            mismatch. Defaults to WORD_ORACLE_FONT_FINGERPRINT.
  --renderer-version <v>    react-docx version override.
  --renderer-build <id>     Optional build or commit identifier.
  --locale <locale>         Browser locale. Default: en-US.
  --timezone <zone>         Browser timezone. Default: UTC.
  --timeout-ms <n>          Per-document readiness deadline. Default: ${DEFAULT_TIMEOUT_MS}.
  --stable-frames <n>       Matching animation-frame snapshots required. Default: ${DEFAULT_STABLE_FRAMES}.
  --headed                  Show the browser window.
  --allow-remote            Permit a non-loopback playground URL. Cross-origin
                            redirects are still rejected.
  --force                   Replace an existing output directory.
  -h, --help                Show this help.

The command reads only local DOCX/reference metadata and the supplied running
playground. It never contacts Microsoft or another oracle provider.`);
}

function nextValue(argv, index, option) {
  const value = argv[index + 1];
  if (!value || value.startsWith("--")) {
    throw new Error(`${option} requires a value`);
  }
  return value;
}

function positiveNumber(value, option) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) {
    throw new Error(`${option} must be a finite number greater than zero`);
  }
  return number;
}

function positiveInteger(value, option) {
  const number = Number(value);
  if (!Number.isInteger(number) || number < 1) {
    throw new Error(`${option} must be an integer greater than zero`);
  }
  return number;
}

function parseArgs(argv) {
  const options = {
    manifestPath: undefined,
    referencePagesPath: undefined,
    baseUrl: undefined,
    outputDir: undefined,
    caseIds: [],
    referenceIds: [],
    browserName: "chromium",
    viewport: parseViewport(DEFAULT_VIEWPORT),
    deviceScaleFactor: 1,
    fontFingerprint:
      process.env.WORD_ORACLE_FONT_FINGERPRINT || undefined,
    rendererVersion: undefined,
    rendererBuild:
      process.env.REACT_DOCX_BUILD || process.env.GITHUB_SHA || undefined,
    locale: "en-US",
    timezoneId: "UTC",
    timeoutMs: DEFAULT_TIMEOUT_MS,
    stableFrames: DEFAULT_STABLE_FRAMES,
    headed: false,
    allowRemote: false,
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
      case "--references":
        options.referencePagesPath = path.resolve(
          nextValue(argv, index, "--references")
        );
        index += 1;
        break;
      case "--base-url":
        options.baseUrl = nextValue(argv, index, "--base-url");
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
      case "--browser":
        options.browserName = nextValue(argv, index, "--browser");
        index += 1;
        break;
      case "--viewport":
        options.viewport = parseViewport(
          nextValue(argv, index, "--viewport")
        );
        index += 1;
        break;
      case "--device-scale-factor":
        options.deviceScaleFactor = positiveNumber(
          nextValue(argv, index, "--device-scale-factor"),
          "--device-scale-factor"
        );
        index += 1;
        break;
      case "--font-fingerprint":
        options.fontFingerprint = nextValue(
          argv,
          index,
          "--font-fingerprint"
        );
        index += 1;
        break;
      case "--renderer-version":
        options.rendererVersion = nextValue(
          argv,
          index,
          "--renderer-version"
        );
        index += 1;
        break;
      case "--renderer-build":
        options.rendererBuild = nextValue(
          argv,
          index,
          "--renderer-build"
        );
        index += 1;
        break;
      case "--locale":
        options.locale = nextValue(argv, index, "--locale");
        index += 1;
        break;
      case "--timezone":
        options.timezoneId = nextValue(argv, index, "--timezone");
        index += 1;
        break;
      case "--timeout-ms":
        options.timeoutMs = positiveInteger(
          nextValue(argv, index, "--timeout-ms"),
          "--timeout-ms"
        );
        index += 1;
        break;
      case "--stable-frames":
        options.stableFrames = positiveInteger(
          nextValue(argv, index, "--stable-frames"),
          "--stable-frames"
        );
        index += 1;
        break;
      case "--headed":
        options.headed = true;
        break;
      case "--allow-remote":
        options.allowRemote = true;
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

  for (const [property, option] of [
    ["manifestPath", "--manifest"],
    ["baseUrl", "--base-url"],
    ["outputDir", "--out-dir"],
    ["fontFingerprint", "--font-fingerprint"],
  ]) {
    if (!options[property]) {
      throw new Error(`${option} is required`);
    }
  }
  if (!new Set(["chromium", "firefox", "webkit"]).has(options.browserName)) {
    throw new Error("--browser must be chromium, firefox, or webkit");
  }
  if (!SHA256_PATTERN.test(options.fontFingerprint)) {
    throw new Error(
      "--font-fingerprint must be a lowercase 64-character SHA-256 digest"
    );
  }
  if (options.stableFrames < 2 || options.stableFrames > 120) {
    throw new Error("--stable-frames must be between 2 and 120");
  }
  options.baseUrl = parseCaptureBaseUrl(options.baseUrl, {
    allowRemote: options.allowRemote,
  });
  return options;
}

async function pathExists(filePath) {
  try {
    await fs.access(filePath);
    return true;
  } catch {
    return false;
  }
}

async function readRendererVersion() {
  const packagePath = path.join(
    repositoryRoot,
    "packages/react-viewer/package.json"
  );
  const packageJson = JSON.parse(await fs.readFile(packagePath, "utf8"));
  if (typeof packageJson.version !== "string" || !packageJson.version) {
    throw new Error(`Unable to determine react-docx version from ${packagePath}`);
  }
  return packageJson.version;
}

async function verifySources(targets, manifestPath) {
  for (const target of targets) {
    const sourcePath = resolveManifestPath(manifestPath, target.source.path);
    let actualHash;
    try {
      actualHash = await sha256File(sourcePath);
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      throw new Error(
        `Cannot read source DOCX ${target.caseId}/${target.source.path}: ${message}`
      );
    }
    if (actualHash !== target.source.sha256) {
      throw new Error(
        `Source DOCX digest mismatch for ${target.caseId}: expected ${target.source.sha256}, got ${actualHash}`
      );
    }
  }
}

async function installDeterministicCaptureStyles(page) {
  await page.addStyleTag({
    content: `
      *, *::before, *::after {
        animation-delay: 0s !important;
        animation-duration: 0s !important;
        caret-color: transparent !important;
        scroll-behavior: auto !important;
        transition-delay: 0s !important;
        transition-duration: 0s !important;
      }
    `,
  });
}

async function waitForStableViewer(page, fileName, options) {
  return page.evaluate(
    async ({ expectedStatus, stableFramesRequired, timeoutMs }) => {
      const startedAt = performance.now();
      let previousSignature = null;
      let stableFrames = 0;
      let lastIssues = [];

      const hashString = (value) => {
        let hash = 2166136261;
        for (let index = 0; index < value.length; index += 1) {
          hash ^= value.charCodeAt(index);
          hash = Math.imul(hash, 16777619);
        }
        return (hash >>> 0).toString(16).padStart(8, "0");
      };
      const rounded = (value) => Math.round(value * 1000) / 1000;
      const nextFrame = () =>
        new Promise((resolve) => requestAnimationFrame(resolve));

      while (performance.now() - startedAt <= timeoutMs) {
        const viewer = document.querySelector(
          '[data-testid="docx-editor-viewer"]'
        );
        const wrappers = viewer
          ? Array.from(
              viewer.querySelectorAll(
                '[data-docx-page-wrapper="true"][data-docx-page-index]'
              )
            )
          : [];
        const statusLoaded = Array.from(document.querySelectorAll("span")).some(
          (element) => element.textContent?.trim() === expectedStatus
        );
        const importError = document.querySelector(
          '[data-docx-import-error="true"]'
        );
        const images = viewer ? Array.from(viewer.querySelectorAll("img")) : [];
        const pendingImages = images.filter(
          (image) => !image.complete || image.naturalWidth <= 0
        );
        const runningAnimations = document
          .getAnimations({ subtree: true })
          .filter(
            (animation) =>
              animation.playState === "running" ||
              animation.playState === "pending"
          );
        const pages = wrappers
          .map((wrapper) => {
            const index = Number(
              wrapper.getAttribute("data-docx-page-index")
            );
            const surfaces = wrapper.querySelectorAll(
              '[data-docx-page-surface="true"]'
            );
            const surface = surfaces[0];
            const wrapperRect = wrapper.getBoundingClientRect();
            const surfaceRect = surface?.getBoundingClientRect();
            const computed = getComputedStyle(wrapper);
            return {
              index,
              width: rounded(wrapperRect.width),
              height: rounded(wrapperRect.height),
              surfaceWidth: rounded(surfaceRect?.width ?? 0),
              surfaceHeight: rounded(surfaceRect?.height ?? 0),
              scrollWidth: wrapper.scrollWidth,
              scrollHeight: wrapper.scrollHeight,
              surfaceCount: surfaces.length,
              visible:
                computed.display !== "none" &&
                computed.visibility !== "hidden" &&
                Number(computed.opacity) !== 0,
              contentHash: surface
                ? hashString(
                    `${surface.innerHTML}\u0000${surface.textContent ?? ""}`
                  )
                : "",
            };
          })
          .sort((left, right) => left.index - right.index);

        const issues = [];
        if (!statusLoaded) {
          issues.push("the playground has not reported the file as loaded");
        }
        if (importError) {
          issues.push(
            `import failed: ${importError.textContent?.trim() || "unknown error"}`
          );
        }
        if (
          document.querySelectorAll('[data-docx-initial-layout-overlay="true"]')
            .length > 0
        ) {
          issues.push("the initial layout overlay is present");
        }
        if (document.fonts.status !== "loaded") {
          issues.push(`fonts are ${document.fonts.status}`);
        }
        if (pendingImages.length > 0) {
          issues.push(`${pendingImages.length} image(s) are pending`);
        }
        if (runningAnimations.length > 0) {
          issues.push(`${runningAnimations.length} animation(s) are running`);
        }
        if (pages.length === 0) {
          issues.push("no rendered pages are present");
        }
        pages.forEach((candidate, index) => {
          if (candidate.index !== index) {
            issues.push("page indexes are not contiguous and zero-based");
          }
          if (
            candidate.width <= 0 ||
            candidate.height <= 0 ||
            candidate.surfaceWidth <= 0 ||
            candidate.surfaceHeight <= 0
          ) {
            issues.push(`page ${index + 1} has invalid geometry`);
          }
          if (candidate.surfaceCount !== 1 || !candidate.visible) {
            issues.push(`page ${index + 1} is not capture-ready`);
          }
        });

        if (issues.length === 0) {
          const signature = JSON.stringify({
            fontStatus: document.fonts.status,
            imageCount: images.length,
            pages: pages.map((candidate) => ({
              index: candidate.index,
              width: candidate.width,
              height: candidate.height,
              surfaceWidth: candidate.surfaceWidth,
              surfaceHeight: candidate.surfaceHeight,
              scrollWidth: candidate.scrollWidth,
              scrollHeight: candidate.scrollHeight,
              contentHash: candidate.contentHash,
            })),
          });
          stableFrames =
            signature === previousSignature ? stableFrames + 1 : 1;
          previousSignature = signature;
          if (stableFrames >= stableFramesRequired) {
            return {
              pages: pages.map((candidate) => ({
                pageIndex: candidate.index,
                widthCssPixels: candidate.surfaceWidth,
                heightCssPixels: candidate.surfaceHeight,
              })),
              stableFrames,
            };
          }
        } else {
          stableFrames = 0;
          previousSignature = null;
          lastIssues = [...new Set(issues)];
        }
        await nextFrame();
      }

      throw new Error(
        `viewer did not become stable: ${
          lastIssues.join("; ") || "page geometry kept changing"
        }`
      );
    },
    {
      expectedStatus: `Loaded ${fileName}`,
      stableFramesRequired: options.stableFrames,
      timeoutMs: options.timeoutMs,
    }
  );
}

async function captureCase(page, target, sourcePath, stagingDir, options) {
  await page.goto(options.baseUrl, {
    waitUntil: "domcontentloaded",
    timeout: options.timeoutMs,
  });
  assertCapturePageUrl(options.baseUrl, page.url(), {
    allowRemote: options.allowRemote,
  });
  await installDeterministicCaptureStyles(page);
  const fileInput = page.locator('input[type="file"][accept*=".doc"]');
  await fileInput.waitFor({ state: "attached", timeout: options.timeoutMs });
  await fileInput.setInputFiles(sourcePath);
  const readiness = await waitForStableViewer(
    page,
    path.basename(sourcePath),
    options
  );

  const canonicalDir = path.join(stagingDir, ".capture", target.caseId);
  await fs.mkdir(canonicalDir, { recursive: true });
  const captured = [];
  for (const pageGeometry of readiness.pages) {
    const pageIndex = pageGeometry.pageIndex;
    const pageNumber = pageIndex + 1;
    const surface = page.locator(
      `[data-docx-page-index="${pageIndex}"] [data-docx-page-surface="true"]`
    );
    if ((await surface.count()) !== 1) {
      throw new Error(
        `Expected one page surface for ${target.caseId} page ${pageNumber}`
      );
    }
    await surface.scrollIntoViewIfNeeded({ timeout: options.timeoutMs });
    const outputPath = path.join(
      canonicalDir,
      `page-${String(pageNumber).padStart(4, "0")}.png`
    );
    await surface.screenshot({
      path: outputPath,
      animations: "disabled",
      caret: "hide",
      scale: "device",
      timeout: options.timeoutMs,
    });
    captured.push({
      pageNumber,
      widthPoints:
        Math.round(pageGeometry.widthCssPixels * 0.75 * 1000) / 1000,
      heightPoints:
        Math.round(pageGeometry.heightCssPixels * 0.75 * 1000) / 1000,
      canonicalPath: outputPath,
      imageSha256: await sha256File(outputPath),
    });
  }

  const references = [];
  for (const referenceId of target.referenceIds) {
    const pages = [];
    for (const capturedPage of captured) {
      const imagePath = portablePagePath(
        target.caseId,
        referenceId,
        capturedPage.pageNumber
      );
      const absolutePath = path.join(
        stagingDir,
        ...imagePath.split(path.posix.sep)
      );
      await fs.mkdir(path.dirname(absolutePath), { recursive: true });
      await fs.copyFile(capturedPage.canonicalPath, absolutePath);
      pages.push({
        pageNumber: capturedPage.pageNumber,
        widthPoints: capturedPage.widthPoints,
        heightPoints: capturedPage.heightPoints,
        imagePath,
        imageSha256: capturedPage.imageSha256,
      });
    }
    references.push({
      caseId: target.caseId,
      referenceId,
      pages,
    });
  }
  return references;
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  if ((await pathExists(options.outputDir)) && !options.force) {
    throw new Error(
      `Output directory already exists: ${options.outputDir}. Pass --force to replace it.`
    );
  }

  const sourceManifest = await readManifest(options.manifestPath);
  const sourceManifestSha256 = await sha256File(options.manifestPath);
  const referencePagesManifest = options.referencePagesPath
    ? await readReferencePagesManifest(options.referencePagesPath)
    : undefined;
  if (
    referencePagesManifest &&
    referencePagesManifest.sourceManifestSha256 !== sourceManifestSha256
  ) {
    throw new Error(
      "reference-pages.json was not materialized from the supplied source manifest"
    );
  }
  const targets = buildCaptureTargets(
    sourceManifest,
    referencePagesManifest,
    options
  );
  await verifySources(targets, options.manifestPath);
  const installedFonts = await computeInstalledFontFingerprint();
  assertExpectedFontFingerprint(
    installedFonts.fingerprintSha256,
    options.fontFingerprint
  );

  const outputParent = path.dirname(options.outputDir);
  await fs.mkdir(outputParent, { recursive: true });
  const stagingDir = await fs.mkdtemp(
    path.join(
      outputParent,
      `.${path.basename(options.outputDir)}-${process.pid}-`
    )
  );

  let browser;
  try {
    const playwright = await import("@playwright/test");
    const browserType = playwright[options.browserName];
    browser = await browserType.launch({ headless: !options.headed });
    const context = await browser.newContext({
      viewport: options.viewport,
      deviceScaleFactor: options.deviceScaleFactor,
      locale: options.locale,
      timezoneId: options.timezoneId,
      colorScheme: "light",
      reducedMotion: "reduce",
    });
    const page = await context.newPage();
    page.setDefaultTimeout(options.timeoutMs);

    const measuredBrowser = await page.evaluate(() => {
      const dateTime = Intl.DateTimeFormat().resolvedOptions();
      return {
        browser: navigator.userAgent,
        platform:
          navigator.platform || navigator.userAgentData?.platform || "",
        viewportWidth: window.innerWidth,
        viewportHeight: window.innerHeight,
        deviceScaleFactor: window.devicePixelRatio,
        locale: dateTime.locale,
        timezone: dateTime.timeZone,
      };
    });
    const measuredEnvironment = buildMeasuredRendererEnvironment({
      ...measuredBrowser,
      browserVersion: browser.version(),
      hostPlatform: os.platform(),
      hostArchitecture: os.arch(),
      fontSetFingerprintSha256: installedFonts.fingerprintSha256,
    });
    assertRequestedCaptureEnvironment(measuredEnvironment, {
      viewport: options.viewport,
      deviceScaleFactor: options.deviceScaleFactor,
      locale: options.locale,
      timezone: options.timezoneId,
      fontSetFingerprintSha256: options.fontFingerprint,
    });

    const references = [];
    for (const target of targets) {
      const sourcePath = resolveManifestPath(
        options.manifestPath,
        target.source.path
      );
      console.log(
        `Capturing ${target.caseId} for ${target.referenceIds.length} reference(s)...`
      );
      references.push(
        ...(await captureCase(
          page,
          target,
          sourcePath,
          stagingDir,
          options
        ))
      );
    }

    const renderer = {
      name: "react-docx",
      version: options.rendererVersion ?? (await readRendererVersion()),
      ...(options.rendererBuild ? { build: options.rendererBuild } : {}),
      ...measuredEnvironment,
    };
    const viewerManifest = buildViewerPagesManifest({
      corpus: sourceManifest.corpus,
      sourceManifestSha256,
      renderer,
      references,
    });
    await fs.rm(path.join(stagingDir, ".capture"), {
      recursive: true,
      force: true,
    });
    await writeJson(path.join(stagingDir, "viewer-pages.json"), viewerManifest);
    await context.close();
    await browser.close();
    browser = undefined;

    if (options.force) {
      await fs.rm(options.outputDir, { recursive: true, force: true });
    }
    await fs.rename(stagingDir, options.outputDir);
    const pageCount = references.reduce(
      (total, reference) => total + reference.pages.length,
      0
    );
    console.log(
      `Captured ${references.length} reference set(s), ${pageCount} page artifact(s). Manifest: ${path.join(
        options.outputDir,
        "viewer-pages.json"
      )}`
    );
  } catch (error) {
    await browser?.close().catch(() => {});
    await fs.rm(stagingDir, { recursive: true, force: true });
    throw error;
  }
}

main().catch((error) => {
  console.error(error instanceof Error ? error.message : String(error));
  process.exit(1);
});
