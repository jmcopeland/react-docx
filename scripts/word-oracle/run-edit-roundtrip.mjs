#!/usr/bin/env node

import { createHash } from "node:crypto";
import fs from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { writeJson } from "./comparison-contract.mjs";
import { readManifest, resolveManifestPath, sha256File } from "./contract.mjs";
import {
  assertSemanticModelExpectation,
  bindEditScenarios,
  buildEditResultsManifest,
  comparableEditSummary,
  formatEditActionIssues,
  validateEditActionManifest,
} from "./edit-action-contract.mjs";
import {
  assertCapturePageUrl,
  parseCaptureBaseUrl,
  parseViewport,
} from "./viewer-capture-contract.mjs";

const scriptDirectory = path.dirname(fileURLToPath(import.meta.url));
const repositoryRoot = path.resolve(scriptDirectory, "../..");
const DEFAULT_VIEWPORT = "1440x1600";
const DEFAULT_TIMEOUT_MS = 120_000;
const DEFAULT_SETTLE_FRAMES = 6;
const SHA256_PATTERN = /^[a-f0-9]{64}$/;

function printUsage() {
  console.log(`Usage:
  node scripts/word-oracle/run-edit-roundtrip.mjs \\
    --manifest <word-oracle.json> \\
    --actions <edit-actions.json> \\
    --base-url <running-development-playground-url> \\
    --out-dir <path> [options]

Options:
  --scenario <id>           Run only this scenario. Can be repeated.
  --browser <name>          chromium, firefox, or webkit. Default: chromium.
  --viewport <WxH>          Browser viewport. Default: ${DEFAULT_VIEWPORT}.
  --device-scale-factor <n> Browser DPR. Default: 1.
  --font-fingerprint <hash> Optional normalized browser-font SHA-256.
  --renderer-version <v>    react-docx version override.
  --renderer-build <id>     Optional build or commit identifier.
  --locale <locale>         Browser locale. Default: en-US.
  --timezone <zone>         Browser timezone. Default: UTC.
  --timeout-ms <n>          Import/action/download deadline. Default: ${DEFAULT_TIMEOUT_MS}.
  --settle-frames <n>       Equal action-state frames required. Default: ${DEFAULT_SETTLE_FRAMES}.
  --headed                  Show the browser window.
  --allow-remote            Permit a non-loopback playground URL. Cross-origin
                            redirects are still rejected.
  --force                   Replace an existing output directory.
  -h, --help                Show this help.

The supplied playground must be a development build exposing
window.__DOCX_TEST_HOOKS__. The command never contacts Microsoft.`);
}

function nextValue(argv, index, option) {
  const value = argv[index + 1];
  if (!value || value.startsWith("--")) {
    throw new Error(`${option} requires a value`);
  }
  return value;
}

function positiveInteger(value, option) {
  const parsed = Number(value);
  if (!Number.isInteger(parsed) || parsed < 1) {
    throw new Error(`${option} must be an integer greater than zero`);
  }
  return parsed;
}

function positiveNumber(value, option) {
  const parsed = Number(value);
  if (!Number.isFinite(parsed) || parsed <= 0) {
    throw new Error(`${option} must be a finite number greater than zero`);
  }
  return parsed;
}

function parseArgs(argv) {
  const options = {
    manifestPath: undefined,
    actionManifestPath: undefined,
    baseUrl: undefined,
    outputDir: undefined,
    scenarioIds: [],
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
    settleFrames: DEFAULT_SETTLE_FRAMES,
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
      case "--actions":
        options.actionManifestPath = path.resolve(
          nextValue(argv, index, "--actions")
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
      case "--scenario":
        options.scenarioIds.push(nextValue(argv, index, "--scenario"));
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
      case "--settle-frames":
        options.settleFrames = positiveInteger(
          nextValue(argv, index, "--settle-frames"),
          "--settle-frames"
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
    ["actionManifestPath", "--actions"],
    ["baseUrl", "--base-url"],
    ["outputDir", "--out-dir"],
  ]) {
    if (!options[property]) {
      throw new Error(`${option} is required`);
    }
  }
  if (!new Set(["chromium", "firefox", "webkit"]).has(options.browserName)) {
    throw new Error("--browser must be chromium, firefox, or webkit");
  }
  if (
    options.fontFingerprint !== undefined &&
    !SHA256_PATTERN.test(options.fontFingerprint)
  ) {
    throw new Error(
      "--font-fingerprint must be a lowercase 64-character SHA-256 digest"
    );
  }
  if (options.settleFrames < 2 || options.settleFrames > 120) {
    throw new Error("--settle-frames must be between 2 and 120");
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

async function readJson(filePath, label) {
  try {
    return JSON.parse(await fs.readFile(filePath, "utf8"));
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`Unable to read ${label} ${filePath}: ${message}`);
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

function sha256Json(value) {
  return createHash("sha256").update(JSON.stringify(value)).digest("hex");
}

async function verifySources(scenarios, manifestPath) {
  const verified = new Set();
  for (const scenario of scenarios) {
    const sourcePath = resolveManifestPath(manifestPath, scenario.source.path);
    if (verified.has(sourcePath)) {
      continue;
    }
    const actualHash = await sha256File(sourcePath);
    if (actualHash !== scenario.source.sha256) {
      throw new Error(
        `Source DOCX digest mismatch for ${scenario.sourceCaseId}: expected ${scenario.source.sha256}, got ${actualHash}`
      );
    }
    verified.add(sourcePath);
  }
}

async function preparePage(page, options) {
  await page.goto(options.baseUrl, {
    waitUntil: "domcontentloaded",
    timeout: options.timeoutMs,
  });
  assertCapturePageUrl(options.baseUrl, page.url(), {
    allowRemote: options.allowRemote,
  });
  await page.addStyleTag({
    content: `
      *, *::before, *::after {
        animation-duration: 0s !important;
        caret-color: transparent !important;
        scroll-behavior: auto !important;
        transition-duration: 0s !important;
      }
    `,
  });
  await page.waitForFunction(
    () => Boolean(window.__DOCX_TEST_HOOKS__),
    undefined,
    { timeout: options.timeoutMs, polling: "raf" }
  );
}

async function waitForImportedDocument(page, fileName, options) {
  return page.evaluate(
    async ({ expectedFileName, timeoutMs, stableFramesRequired }) => {
      const startTime = performance.now();
      let previousSignature = null;
      let stableFrames = 0;
      let lastIssue = "test hooks are unavailable";

      const hashString = (value) => {
        let hash = 2166136261;
        for (let index = 0; index < value.length; index += 1) {
          hash ^= value.charCodeAt(index);
          hash = Math.imul(hash, 16777619);
        }
        return (hash >>> 0).toString(16).padStart(8, "0");
      };
      const nextFrame = () =>
        new Promise((resolve) => requestAnimationFrame(resolve));

      while (performance.now() - startTime <= timeoutMs) {
        const hooks = window.__DOCX_TEST_HOOKS__;
        const importError = document.querySelector(
          '[data-docx-import-error="true"]'
        );
        const images = Array.from(
          document.querySelectorAll('[data-testid="docx-editor-viewer"] img')
        );
        const imagesReady = images.every(
          (image) => image.complete && image.naturalWidth > 0
        );
        const pages = Array.from(
          document.querySelectorAll('[data-docx-page-surface="true"]')
        );
        const summary = hooks?.getSummary();
        const ready =
          hooks &&
          summary?.fileName === expectedFileName &&
          summary.status === `Loaded ${expectedFileName}` &&
          !importError &&
          document.fonts.status === "loaded" &&
          imagesReady &&
          pages.length > 0;

        if (ready) {
          const signature = JSON.stringify({
            summary,
            images: images.map((image) => [
              image.currentSrc || image.src,
              image.naturalWidth,
              image.naturalHeight,
            ]),
            pages: pages.map((surface) => {
              const rect = surface.getBoundingClientRect();
              return [
                Math.round(rect.width * 1000) / 1000,
                Math.round(rect.height * 1000) / 1000,
                hashString(`${surface.innerHTML}\u0000${surface.textContent ?? ""}`),
              ];
            }),
          });
          stableFrames =
            signature === previousSignature ? stableFrames + 1 : 1;
          previousSignature = signature;
          if (stableFrames >= stableFramesRequired) {
            return summary;
          }
        } else {
          stableFrames = 0;
          previousSignature = null;
          lastIssue = importError
            ? importError.textContent?.trim() || "the viewer reported an import error"
            : !hooks
            ? "test hooks are unavailable"
            : summary?.fileName !== expectedFileName
            ? `waiting for ${expectedFileName}`
            : summary.status !== `Loaded ${expectedFileName}`
            ? `viewer status is ${summary.status}`
            : document.fonts.status !== "loaded"
            ? `fonts are ${document.fonts.status}`
            : !imagesReady
            ? "viewer images are not ready"
            : "no rendered pages are present";
        }
        await nextFrame();
      }
      throw new Error(`document import did not settle: ${lastIssue}`);
    },
    {
      expectedFileName: fileName,
      timeoutMs: options.timeoutMs,
      stableFramesRequired: options.settleFrames,
    }
  );
}

async function importDocument(page, filePath, options) {
  const input = page.locator('input[type="file"][accept*=".doc"]');
  await input.waitFor({ state: "attached", timeout: options.timeoutMs });
  await input.setInputFiles(filePath);
  return waitForImportedDocument(page, path.basename(filePath), options);
}

async function replayAction(page, action, options) {
  return page.evaluate(
    async ({ actionToRun, timeoutMs, stableFramesRequired }) => {
      const hooks = window.__DOCX_TEST_HOOKS__;
      if (!hooks) {
        throw new Error("window.__DOCX_TEST_HOOKS__ is unavailable");
      }

      const assertParagraph = (nodeIndex) => {
        const paragraph = document.querySelector(
          `[data-docx-paragraph-kind="paragraph"][data-docx-paragraph-node-index="${nodeIndex}"]`
        );
        if (!paragraph) {
          throw new Error(`paragraph node ${nodeIndex} does not exist`);
        }
      };
      const tableShape = (tableIndex) => {
        const shape = hooks.getTableShape(tableIndex);
        if (!shape) {
          throw new Error(`table node ${tableIndex} does not exist`);
        }
        return shape;
      };
      const assertCell = (tableIndex, rowIndex, cellIndex) => {
        const shape = tableShape(tableIndex);
        const columnCount = shape.columnCounts[rowIndex];
        if (columnCount === undefined) {
          throw new Error(`table ${tableIndex} row ${rowIndex} does not exist`);
        }
        if (cellIndex >= columnCount) {
          throw new Error(
            `table ${tableIndex} row ${rowIndex} cell ${cellIndex} does not exist`
          );
        }
      };
      const assertLocation = (location) => {
        if (location.kind === "paragraph") {
          assertParagraph(location.nodeIndex);
          return;
        }
        assertCell(location.tableIndex, location.rowIndex, location.cellIndex);
        const paragraph = document.querySelector(
          `[data-docx-paragraph-kind="table-cell"][data-docx-table-index="${location.tableIndex}"][data-docx-row-index="${location.rowIndex}"][data-docx-cell-index="${location.cellIndex}"][data-docx-paragraph-index="${location.paragraphIndex}"]`
        );
        if (!paragraph) {
          throw new Error(
            `table ${location.tableIndex} cell paragraph ${location.paragraphIndex} does not exist`
          );
        }
      };

      switch (actionToRun.type) {
        case "select-paragraph":
          assertParagraph(actionToRun.nodeIndex);
          hooks.selectParagraph(actionToRun.nodeIndex);
          break;
        case "select-table-cell":
          assertCell(
            actionToRun.tableIndex,
            actionToRun.rowIndex,
            actionToRun.cellIndex
          );
          hooks.selectTableCell(
            actionToRun.tableIndex,
            actionToRun.rowIndex,
            actionToRun.cellIndex
          );
          break;
        case "set-text-range":
          if (actionToRun.range) {
            assertLocation(actionToRun.range.start.location);
            assertLocation(actionToRun.range.end.location);
          }
          hooks.setActiveTextRange(actionToRun.range ?? undefined);
          break;
        case "commit-paragraph-text":
          assertParagraph(actionToRun.nodeIndex);
          hooks.commitParagraphText(actionToRun.nodeIndex, actionToRun.text);
          break;
        case "commit-table-cell-text":
          assertCell(
            actionToRun.tableIndex,
            actionToRun.rowIndex,
            actionToRun.cellIndex
          );
          hooks.commitTableCellText(
            actionToRun.tableIndex,
            actionToRun.rowIndex,
            actionToRun.cellIndex,
            actionToRun.text
          );
          break;
        case "toggle-bold":
          hooks.toggleBold();
          break;
        case "toggle-italic":
          hooks.toggleItalic();
          break;
        case "toggle-underline":
          hooks.toggleUnderline();
          break;
        case "toggle-strike":
          hooks.toggleStrike();
          break;
        case "set-text-color":
          hooks.setTextColor(actionToRun.color ?? undefined);
          break;
        case "set-highlight":
          hooks.setHighlight(actionToRun.color ?? undefined);
          break;
        case "set-font-family":
          hooks.setFontFamily(actionToRun.fontFamily);
          break;
        case "set-font-size":
          hooks.setFontSize(actionToRun.fontSizePt);
          break;
        case "set-alignment":
          hooks.setAlignment(actionToRun.alignment ?? undefined);
          break;
        case "toggle-list":
          hooks.toggleList(actionToRun.listType);
          break;
        case "set-line-spacing":
          hooks.setLineSpacing(actionToRun.lineMultiple);
          break;
        case "insert-table-row": {
          const shape = tableShape(actionToRun.tableIndex);
          if (actionToRun.rowIndex >= shape.rowCount) {
            throw new Error(
              `table ${actionToRun.tableIndex} row ${actionToRun.rowIndex} does not exist`
            );
          }
          hooks.insertTableRow(
            actionToRun.tableIndex,
            actionToRun.rowIndex,
            actionToRun.direction
          );
          break;
        }
        case "insert-table-column": {
          const shape = tableShape(actionToRun.tableIndex);
          const rowIndex = actionToRun.rowIndex ?? 0;
          const columnCount = shape.columnCounts[rowIndex];
          if (columnCount === undefined || actionToRun.cellIndex >= columnCount) {
            throw new Error(
              `table ${actionToRun.tableIndex} cell ${actionToRun.cellIndex} does not exist in row ${rowIndex}`
            );
          }
          hooks.insertTableColumn(
            actionToRun.tableIndex,
            actionToRun.cellIndex,
            actionToRun.direction,
            actionToRun.rowIndex
          );
          break;
        }
        case "delete-table-row": {
          const shape = tableShape(actionToRun.tableIndex);
          if (actionToRun.rowIndex >= shape.rowCount) {
            throw new Error(
              `table ${actionToRun.tableIndex} row ${actionToRun.rowIndex} does not exist`
            );
          }
          hooks.deleteTableRow(actionToRun.tableIndex, actionToRun.rowIndex);
          break;
        }
        case "delete-table-column": {
          const shape = tableShape(actionToRun.tableIndex);
          const rowIndex = actionToRun.rowIndex ?? 0;
          const columnCount = shape.columnCounts[rowIndex];
          if (columnCount === undefined || actionToRun.cellIndex >= columnCount) {
            throw new Error(
              `table ${actionToRun.tableIndex} cell ${actionToRun.cellIndex} does not exist in row ${rowIndex}`
            );
          }
          hooks.deleteTableColumn(
            actionToRun.tableIndex,
            actionToRun.cellIndex,
            actionToRun.rowIndex
          );
          break;
        }
        case "undo":
          if (!hooks.getSummary().canUndo) {
            throw new Error("undo is unavailable");
          }
          hooks.undo();
          break;
        case "redo":
          if (!hooks.getSummary().canRedo) {
            throw new Error("redo is unavailable");
          }
          hooks.redo();
          break;
        case "accept-tracked-change":
          if (actionToRun.index >= hooks.getSummary().trackedChangeCount) {
            throw new Error(
              `tracked change ${actionToRun.index} does not exist`
            );
          }
          hooks.acceptTrackedChange(actionToRun.index);
          break;
        case "reject-tracked-change":
          if (actionToRun.index >= hooks.getSummary().trackedChangeCount) {
            throw new Error(
              `tracked change ${actionToRun.index} does not exist`
            );
          }
          hooks.rejectTrackedChange(actionToRun.index);
          break;
        case "create-comment":
          hooks.createComment(actionToRun.text, actionToRun.options);
          break;
        case "set-comment-resolved":
          hooks.setCommentResolved(
            actionToRun.commentId,
            actionToRun.resolved
          );
          break;
        default:
          throw new Error(`unsupported edit action: ${actionToRun.type}`);
      }

      const startTime = performance.now();
      let previousHooks;
      let previousSignature;
      let stableFrames = 0;
      const nextFrame = () =>
        new Promise((resolve) => requestAnimationFrame(resolve));
      const hashString = (value) => {
        let hash = 2166136261;
        for (let index = 0; index < value.length; index += 1) {
          hash ^= value.charCodeAt(index);
          hash = Math.imul(hash, 16777619);
        }
        return (hash >>> 0).toString(16).padStart(8, "0");
      };
      const pageSignature = () =>
        Array.from(
          document.querySelectorAll('[data-docx-page-surface="true"]')
        ).map((surface) =>
          hashString(`${surface.innerHTML}\u0000${surface.textContent ?? ""}`)
        );

      while (performance.now() - startTime <= timeoutMs) {
        await nextFrame();
        const currentHooks = window.__DOCX_TEST_HOOKS__;
        if (!currentHooks) {
          stableFrames = 0;
          continue;
        }
        const signature = JSON.stringify({
          summary: currentHooks.getSummary(),
          pages: pageSignature(),
        });
        stableFrames =
          currentHooks === previousHooks && signature === previousSignature
            ? stableFrames + 1
            : 1;
        previousHooks = currentHooks;
        previousSignature = signature;
        if (stableFrames >= stableFramesRequired) {
          return currentHooks.getSummary();
        }
      }
      throw new Error("editor state did not settle after the action");
    },
    {
      actionToRun: action,
      timeoutMs: options.timeoutMs,
      stableFramesRequired: options.settleFrames,
    }
  );
}

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function assertExact(actual, expected, label) {
  if (!sameJson(actual, expected)) {
    throw new Error(
      `${label} is ${JSON.stringify(actual)}, expected ${JSON.stringify(expected)}`
    );
  }
}

function requireObserved(value, label) {
  if (value === undefined || value === null) {
    throw new Error(`${label} could not be read back from the editor model`);
  }
  return value;
}

async function readActionEffectState(page, effect) {
  return page.evaluate((expectedEffect) => {
    const hooks = window.__DOCX_TEST_HOOKS__;
    if (!hooks) {
      throw new Error("window.__DOCX_TEST_HOOKS__ is unavailable");
    }
    const state = {
      action: hooks.getActionState(),
    };
    switch (expectedEffect.kind) {
      case "paragraph-text":
        state.location = hooks.getLocationState({
          kind: "paragraph",
          nodeIndex: expectedEffect.nodeIndex,
        });
        break;
      case "table-cell-text":
        state.tableCellText = hooks.getTableCellText(
          expectedEffect.tableIndex,
          expectedEffect.rowIndex,
          expectedEffect.cellIndex
        );
        break;
      case "text-style":
        state.range = hooks.getRangeState(expectedEffect.range);
        break;
      case "paragraph-style":
        state.location = hooks.getLocationState(expectedEffect.location);
        break;
      case "table-shape":
        state.tableShape = hooks.getTableShape(expectedEffect.tableIndex);
        break;
      case "tracked-change":
        state.location = hooks.getLocationState(expectedEffect.location);
        state.trackedChanges = hooks.getTrackedChanges();
        break;
      case "created-comment":
        state.range = hooks.getRangeState(expectedEffect.range);
        state.location = hooks.getLocationState(expectedEffect.location);
        state.comments = hooks.getComments();
        break;
      case "resolved-comment":
        state.location = hooks.getLocationState(expectedEffect.location);
        state.comments = hooks.getComments();
        break;
      default:
        break;
    }
    return state;
  }, effect);
}

function observedStyleValue(rangeState, property) {
  if (!Array.isArray(rangeState.styles) || rangeState.styles.length === 0) {
    throw new Error("target text range contains no readable text-run styles");
  }
  const values = rangeState.styles.map((style) => style[property]);
  if (!values.every((value) => sameJson(value, values[0]))) {
    throw new Error(
      `target text range has mixed ${property} values: ${JSON.stringify(values)}`
    );
  }
  return values[0];
}

function findExpectedComment(comments, effect) {
  return comments.find(
    (comment) => comment.commentId === effect.commentId && comment.id === effect.id
  );
}

export async function verifyActionPostcondition(
  action,
  beforeSummary,
  afterSummary,
  beforeEffectState,
  afterEffectState
) {
  const semanticTransition = assertSemanticModelExpectation(
    action.expect.semanticModel,
    beforeSummary,
    afterSummary
  );
  const effect = action.expect.effect;
  let targetBeforeDigest;
  let targetAfterDigest;
  let observed;

  switch (effect.kind) {
    case "selection":
      assertExact(
        afterEffectState.action.selection,
        effect.selection,
        "editor selection"
      );
      observed = afterEffectState.action.selection;
      break;
    case "active-text-range":
      assertExact(
        afterEffectState.action.activeTextRange,
        effect.range,
        "active text range"
      );
      observed = afterEffectState.action.activeTextRange;
      break;
    case "paragraph-text": {
      const before = requireObserved(
        beforeEffectState.location,
        `paragraph ${effect.nodeIndex} before state`
      );
      const after = requireObserved(
        afterEffectState.location,
        `paragraph ${effect.nodeIndex} after state`
      );
      if (before.text === effect.text) {
        throw new Error("paragraph target already contained the expected text before the action");
      }
      assertExact(after.text, effect.text, `paragraph ${effect.nodeIndex} text`);
      targetBeforeDigest = before.digest;
      targetAfterDigest = after.digest;
      if (targetBeforeDigest === targetAfterDigest) {
        throw new Error("paragraph target digest did not change");
      }
      observed = { nodeIndex: effect.nodeIndex, text: after.text };
      break;
    }
    case "table-cell-text":
      if (beforeEffectState.tableCellText === effect.text) {
        throw new Error("table-cell target already contained the expected text before the action");
      }
      assertExact(
        afterEffectState.tableCellText,
        effect.text,
        `table ${effect.tableIndex} row ${effect.rowIndex} cell ${effect.cellIndex} text`
      );
      observed = {
        tableIndex: effect.tableIndex,
        rowIndex: effect.rowIndex,
        cellIndex: effect.cellIndex,
        text: afterEffectState.tableCellText,
      };
      break;
    case "text-style": {
      const before = requireObserved(beforeEffectState.range, "target text range before state");
      const after = requireObserved(afterEffectState.range, "target text range after state");
      assertExact(afterEffectState.action.activeTextRange, effect.range, "active text range");
      assertExact(after.range, effect.range, "styled text range");
      assertExact(after.text, before.text, "styled text range text");
      targetBeforeDigest = before.digest;
      targetAfterDigest = after.digest;
      if (targetBeforeDigest === targetAfterDigest) {
        throw new Error("target text range style digest did not change");
      }
      const beforeValue = observedStyleValue(before, effect.property);
      if (sameJson(beforeValue, effect.value)) {
        throw new Error(
          `target text range ${effect.property} already equaled the expected value before the action`
        );
      }
      if (
        action.type.startsWith("toggle-") &&
        effect.value !== !Boolean(beforeValue)
      ) {
        throw new Error(
          `toggle expectation for ${effect.property} is not the inverse of the uniform before value`
        );
      }
      const value = observedStyleValue(after, effect.property);
      assertExact(value, effect.value, `target text range ${effect.property}`);
      observed = { range: after.range, property: effect.property, value };
      break;
    }
    case "paragraph-style": {
      const before = requireObserved(beforeEffectState.location, "target paragraph before state");
      const after = requireObserved(afterEffectState.location, "target paragraph after state");
      assertExact(
        afterEffectState.action.selectedParagraphLocation,
        effect.location,
        "selected paragraph location"
      );
      if (sameJson(before[effect.property], effect.value)) {
        throw new Error(
          `target paragraph ${effect.property} already equaled the expected value before the action`
        );
      }
      if (action.type === "toggle-list") {
        const expectedListType =
          before.listType === action.listType ? null : action.listType;
        if (effect.value !== expectedListType) {
          throw new Error(
            `toggle-list expectation must be ${JSON.stringify(
              expectedListType
            )} for the observed before state`
          );
        }
      }
      assertExact(
        after[effect.property],
        effect.value,
        `target paragraph ${effect.property}`
      );
      targetBeforeDigest = before.digest;
      targetAfterDigest = after.digest;
      if (targetBeforeDigest === targetAfterDigest) {
        throw new Error("target paragraph style digest did not change");
      }
      observed = {
        location: after.location,
        property: effect.property,
        value: after[effect.property],
      };
      break;
    }
    case "table-shape": {
      const expectedShape = {
        rowCount: effect.rowCount,
        columnCounts: effect.columnCounts,
      };
      const before = requireObserved(beforeEffectState.tableShape, "target table before shape");
      const after = requireObserved(afterEffectState.tableShape, "target table after shape");
      if (sameJson(before, after)) {
        throw new Error("target table shape did not change");
      }
      assertExact(after, expectedShape, `table ${effect.tableIndex} shape`);
      observed = { tableIndex: effect.tableIndex, ...after };
      break;
    }
    case "history":
      assertExact(
        afterEffectState.action.semanticModelDigest,
        effect.semanticModelDigest,
        "history-restored semantic model digest"
      );
      assertExact(afterEffectState.action.canUndo, effect.canUndo, "history canUndo");
      assertExact(afterEffectState.action.canRedo, effect.canRedo, "history canRedo");
      observed = {
        semanticModelDigest: afterEffectState.action.semanticModelDigest,
        canUndo: afterEffectState.action.canUndo,
        canRedo: afterEffectState.action.canRedo,
      };
      break;
    case "tracked-change": {
      const beforeChanges = beforeEffectState.trackedChanges ?? [];
      const beforeChange = beforeChanges[action.index];
      if (!beforeChange) {
        throw new Error(`tracked change ${action.index} was missing before the action`);
      }
      assertExact(beforeChange.id, effect.changeId, "tracked change id");
      assertExact(beforeChange.revisionId, effect.revisionId, "tracked change revisionId");
      assertExact(beforeChange.kind, effect.changeKind, "tracked change kind");
      assertExact(beforeChange.location, effect.location, "tracked change location");
      assertExact(beforeChange.text ?? "", effect.changeText, "tracked change text");
      const afterChanges = afterEffectState.trackedChanges ?? [];
      if (
        afterChanges.some(
          (change) =>
            change.id === effect.changeId || change.revisionId === effect.revisionId
        )
      ) {
        throw new Error("the exact tracked change still exists after accept/reject");
      }
      assertExact(afterChanges.length, effect.remainingCount, "remaining tracked-change count");
      const beforeLocation = requireObserved(beforeEffectState.location, "tracked-change target before state");
      const afterLocation = requireObserved(afterEffectState.location, "tracked-change target after state");
      assertExact(afterLocation.text, effect.resultText, "tracked-change target result text");
      targetBeforeDigest = beforeLocation.digest;
      targetAfterDigest = afterLocation.digest;
      observed = {
        changeId: effect.changeId,
        revisionId: effect.revisionId,
        remainingCount: afterChanges.length,
        location: afterLocation.location,
        resultText: afterLocation.text,
      };
      break;
    }
    case "created-comment": {
      const beforeComments = beforeEffectState.comments ?? [];
      if (
        beforeComments.some(
          (comment) => comment.commentId === effect.commentId || comment.id === effect.id
        )
      ) {
        throw new Error("expected created comment identity already existed before the action");
      }
      assertExact(beforeEffectState.action.activeTextRange, effect.range, "comment creation range");
      const afterComments = afterEffectState.comments ?? [];
      const comment = requireObserved(
        findExpectedComment(afterComments, effect),
        `created comment ${effect.commentId}`
      );
      for (const key of [
        "id",
        "commentId",
        "location",
        "text",
        "anchorText",
        "author",
        "initials",
        "date",
      ]) {
        assertExact(comment[key], effect[key], `created comment ${key}`);
      }
      assertExact(comment.resolved ?? false, effect.resolved, "created comment resolved");
      assertExact(afterComments.length, effect.commentCount, "created comment total count");
      assertExact(
        afterComments.filter((entry) => entry.resolved === true).length,
        effect.resolvedCommentCount,
        "created comment resolved count"
      );
      observed = {
        id: comment.id,
        commentId: comment.commentId,
        location: comment.location,
        anchorText: comment.anchorText,
        commentCount: afterComments.length,
      };
      break;
    }
    case "resolved-comment": {
      const beforeComments = beforeEffectState.comments ?? [];
      const beforeComment = requireObserved(
        findExpectedComment(beforeComments, effect),
        `comment ${effect.commentId} before state`
      );
      assertExact(beforeComment.location, effect.location, "comment location before resolution");
      assertExact(beforeComment.text, effect.text, "comment text before resolution");
      if ((beforeComment.resolved ?? false) === effect.resolved) {
        throw new Error("comment already had the expected resolution state before the action");
      }
      const afterComments = afterEffectState.comments ?? [];
      const afterComment = requireObserved(
        findExpectedComment(afterComments, effect),
        `comment ${effect.commentId} after state`
      );
      assertExact(afterComment.location, effect.location, "resolved comment location");
      assertExact(afterComment.text, effect.text, "resolved comment text");
      assertExact(afterComment.resolved ?? false, effect.resolved, "comment resolved state");
      assertExact(afterComments.length, effect.commentCount, "resolved comment total count");
      assertExact(
        afterComments.filter((entry) => entry.resolved === true).length,
        effect.resolvedCommentCount,
        "resolved comment count"
      );
      observed = {
        id: afterComment.id,
        commentId: afterComment.commentId,
        location: afterComment.location,
        text: afterComment.text,
        resolved: afterComment.resolved ?? false,
      };
      break;
    }
    default:
      throw new Error(`unsupported action effect: ${effect.kind}`);
  }

  return {
    semanticModel: action.expect.semanticModel,
    beforeDigest: semanticTransition.beforeDigest,
    afterDigest: semanticTransition.afterDigest,
    effectKind: effect.kind,
    ...(targetBeforeDigest !== undefined ? { targetBeforeDigest } : {}),
    ...(targetAfterDigest !== undefined ? { targetAfterDigest } : {}),
    observed,
  };
}

async function verifyScenarioPostcondition(
  page,
  expected,
  sourceSummary,
  editedSummary
) {
  assertSemanticModelExpectation(
    expected.semanticModel,
    sourceSummary,
    editedSummary
  );
  if (
    expected.semanticModelDigest !== undefined &&
    editedSummary.semanticModelDigest !== expected.semanticModelDigest
  ) {
    throw new Error(
      `final semantic model digest ${editedSummary.semanticModelDigest} does not match expected ${expected.semanticModelDigest}`
    );
  }
  for (const field of [
    "trackedChangeCount",
    "commentCount",
    "resolvedCommentCount",
  ]) {
    if (expected[field] !== undefined && editedSummary[field] !== expected[field]) {
      throw new Error(
        `${field} is ${editedSummary[field]}, expected ${expected[field]}`
      );
    }
  }

  const observations = await page.evaluate((postcondition) => {
    const hooks = window.__DOCX_TEST_HOOKS__;
    if (!hooks) {
      throw new Error("window.__DOCX_TEST_HOOKS__ is unavailable");
    }
    return {
      paragraphTexts: (postcondition.paragraphTexts ?? []).map((entry) => {
        const exists = Boolean(
          document.querySelector(
            `[data-docx-paragraph-kind="paragraph"][data-docx-paragraph-node-index="${entry.nodeIndex}"]`
          )
        );
        return { ...entry, exists, actual: hooks.getParagraphText(entry.nodeIndex) };
      }),
      tableCellTexts: (postcondition.tableCellTexts ?? []).map((entry) => ({
        ...entry,
        tableExists: Boolean(hooks.getTableShape(entry.tableIndex)),
        actual: hooks.getTableCellText(
          entry.tableIndex,
          entry.rowIndex,
          entry.cellIndex
        ),
      })),
      tableShapes: (postcondition.tableShapes ?? []).map((entry) => ({
        ...entry,
        actual: hooks.getTableShape(entry.tableIndex),
      })),
    };
  }, expected);

  for (const observation of observations.paragraphTexts) {
    if (!observation.exists || observation.actual !== observation.text) {
      throw new Error(
        `paragraph ${observation.nodeIndex} postcondition failed: expected ${JSON.stringify(
          observation.text
        )}, got ${JSON.stringify(observation.actual)}`
      );
    }
  }
  for (const observation of observations.tableCellTexts) {
    if (!observation.tableExists || observation.actual !== observation.text) {
      throw new Error(
        `table ${observation.tableIndex} row ${observation.rowIndex} cell ${
          observation.cellIndex
        } postcondition failed: expected ${JSON.stringify(
          observation.text
        )}, got ${JSON.stringify(observation.actual)}`
      );
    }
  }
  for (const observation of observations.tableShapes) {
    if (
      !observation.actual ||
      observation.actual.rowCount !== observation.rowCount ||
      JSON.stringify(observation.actual.columnCounts) !==
        JSON.stringify(observation.columnCounts)
    ) {
      throw new Error(
        `table ${observation.tableIndex} shape postcondition failed: expected ${JSON.stringify(
          {
            rowCount: observation.rowCount,
            columnCounts: observation.columnCounts,
          }
        )}, got ${JSON.stringify(observation.actual)}`
      );
    }
  }

  return {
    semanticModel: expected.semanticModel,
    sourceDigest: sourceSummary.semanticModelDigest,
    finalDigest: editedSummary.semanticModelDigest,
    ...(expected.semanticModelDigest
      ? { expectedDigest: expected.semanticModelDigest }
      : {}),
    observations,
  };
}

async function exportDocument(page, outputPath, options) {
  const downloadPromise = page.waitForEvent("download", {
    timeout: options.timeoutMs,
  });
  await page.evaluate(() => {
    const hooks = window.__DOCX_TEST_HOOKS__;
    if (!hooks) {
      throw new Error("window.__DOCX_TEST_HOOKS__ is unavailable");
    }
    hooks.exportDocx();
  });
  const download = await downloadPromise;
  await fs.mkdir(path.dirname(outputPath), { recursive: true });
  await download.saveAs(outputPath);
  const failure = await download.failure();
  if (failure) {
    throw new Error(`DOCX download failed: ${failure}`);
  }
  const stat = await fs.stat(outputPath);
  if (stat.size <= 0) {
    throw new Error("Exported DOCX is empty");
  }
  return {
    suggestedFileName: download.suggestedFilename(),
    sizeBytes: stat.size,
  };
}

async function runScenario(context, scenario, stagingDir, options) {
  const sourcePath = resolveManifestPath(options.manifestPath, scenario.source.path);
  const page = await context.newPage();
  await preparePage(page, options);
  const sourceSummary = await importDocument(page, sourcePath, options);
  const actionOutcomes = [];
  for (let index = 0; index < scenario.actions.length; index += 1) {
    try {
      const beforeSummary = await page.evaluate(() =>
        window.__DOCX_TEST_HOOKS__.getSummary()
      );
      const action = scenario.actions[index];
      const beforeEffectState = await readActionEffectState(
        page,
        action.expect.effect
      );
      const summary = await replayAction(page, action, options);
      const afterEffectState = await readActionEffectState(
        page,
        action.expect.effect
      );
      const postcondition = await verifyActionPostcondition(
        action,
        beforeSummary,
        summary,
        beforeEffectState,
        afterEffectState
      );
      actionOutcomes.push({
        actionIndex: index,
        type: scenario.actions[index].type,
        status: summary.status,
        canUndo: summary.canUndo,
        canRedo: summary.canRedo,
        postcondition,
      });
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      throw new Error(
        `Scenario ${scenario.id} action ${index} (${scenario.actions[index].type}) failed: ${message}`
      );
    }
  }
  const editedSummary = await page.evaluate(() =>
    window.__DOCX_TEST_HOOKS__.getSummary()
  );
  let scenarioPostcondition;
  try {
    scenarioPostcondition = await verifyScenarioPostcondition(
      page,
      scenario.expected,
      sourceSummary,
      editedSummary
    );
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(`Scenario ${scenario.id} final postcondition failed: ${message}`);
  }
  const relativeExportPath = path.posix.join(
    scenario.id,
    `${scenario.id}.docx`
  );
  const absoluteExportPath = path.join(
    stagingDir,
    ...relativeExportPath.split(path.posix.sep)
  );
  const download = await exportDocument(page, absoluteExportPath, options);
  await page.close();

  const reopenPage = await context.newPage();
  await preparePage(reopenPage, options);
  const reopenedSummary = await importDocument(
    reopenPage,
    absoluteExportPath,
    options
  );
  await reopenPage.close();
  const editedStructure = comparableEditSummary(editedSummary);
  const reopenedStructure = comparableEditSummary(reopenedSummary);
  if (JSON.stringify(editedStructure) !== JSON.stringify(reopenedStructure)) {
    throw new Error(
      `Scenario ${scenario.id} changed structure after export/reopen:\nexpected ${JSON.stringify(
        editedStructure
      )}\nreceived ${JSON.stringify(reopenedStructure)}`
    );
  }

  return {
    id: scenario.id,
    sourceCaseId: scenario.sourceCaseId,
    ...(scenario.description ? { description: scenario.description } : {}),
    actions: {
      count: scenario.actions.length,
      sha256: sha256Json(scenario.actions),
      outcomes: actionOutcomes,
    },
    inputSource: {
      sha256: scenario.source.sha256,
    },
    exportedDocx: {
      path: relativeExportPath,
      sha256: await sha256File(absoluteExportPath),
      mediaType:
        "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      sizeBytes: download.sizeBytes,
      suggestedFileName: download.suggestedFileName,
    },
    sourceSummary,
    editedSummary,
    reopenedSummary,
    scenarioPostcondition,
    reopenVerification: {
      importSucceeded: true,
      structuralSummaryMatched: true,
    },
  };
}

async function main() {
  const options = parseArgs(process.argv.slice(2));
  if ((await pathExists(options.outputDir)) && !options.force) {
    throw new Error(
      `Output directory already exists: ${options.outputDir}. Pass --force to replace it.`
    );
  }
  const [sourceManifest, actionManifest] = await Promise.all([
    readManifest(options.manifestPath),
    readJson(options.actionManifestPath, "edit action manifest"),
  ]);
  const actionIssues = validateEditActionManifest(actionManifest);
  if (actionIssues.length > 0) {
    throw new Error(
      `Invalid edit action manifest ${options.actionManifestPath}:\n${formatEditActionIssues(
        actionIssues
      )}`
    );
  }
  const [sourceManifestSha256, actionManifestSha256] = await Promise.all([
    sha256File(options.manifestPath),
    sha256File(options.actionManifestPath),
  ]);
  const scenarios = bindEditScenarios(
    actionManifest,
    sourceManifest,
    sourceManifestSha256,
    options.scenarioIds
  );
  await verifySources(scenarios, options.manifestPath);

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
    browser = await playwright[options.browserName].launch({
      headless: !options.headed,
    });
    const context = await browser.newContext({
      viewport: options.viewport,
      deviceScaleFactor: options.deviceScaleFactor,
      locale: options.locale,
      timezoneId: options.timezoneId,
      colorScheme: "light",
      reducedMotion: "reduce",
      acceptDownloads: true,
    });
    const results = [];
    for (const scenario of scenarios) {
      console.log(
        `Running ${scenario.id}: ${scenario.actions.length} action(s) on ${scenario.sourceCaseId}...`
      );
      results.push(
        await runScenario(context, scenario, stagingDir, options)
      );
    }

    const runner = {
      name: "react-docx-playground-edit-runner",
      version: "1",
      rendererVersion:
        options.rendererVersion ?? (await readRendererVersion()),
      ...(options.rendererBuild ? { rendererBuild: options.rendererBuild } : {}),
      browser: `${options.browserName} ${browser.version()}`,
      platform: `${os.platform()}-${os.arch()}`,
      viewport: `${options.viewport.width}x${options.viewport.height}`,
      deviceScaleFactor: options.deviceScaleFactor,
      locale: options.locale,
      timezone: options.timezoneId,
      ...(options.fontFingerprint
        ? { fontSetFingerprintSha256: options.fontFingerprint }
        : {}),
    };
    const resultManifest = buildEditResultsManifest({
      corpus: sourceManifest.corpus,
      sourceManifestSha256,
      actionManifestSha256,
      runner,
      scenarios: results,
    });
    await writeJson(
      path.join(stagingDir, "edit-results.json"),
      resultManifest
    );
    await context.close();
    await browser.close();
    browser = undefined;

    if (options.force) {
      await fs.rm(options.outputDir, { recursive: true, force: true });
    }
    await fs.rename(stagingDir, options.outputDir);
    console.log(
      `Completed ${results.length} edit scenario(s). Manifest: ${path.join(
        options.outputDir,
        "edit-results.json"
      )}`
    );
  } catch (error) {
    await browser?.close().catch(() => {});
    await fs.rm(stagingDir, { recursive: true, force: true });
    throw error;
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
