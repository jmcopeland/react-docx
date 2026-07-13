import assert from "node:assert/strict";
import fs from "node:fs/promises";
import path from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";

import {
  assertSemanticModelExpectation,
  bindEditScenarios,
  buildEditResultsManifest,
  comparableEditSummary,
  SUPPORTED_EDIT_ACTION_TYPES,
  validateEditActionManifest,
} from "../../scripts/word-oracle/edit-action-contract.mjs";
import { verifyActionPostcondition } from "../../scripts/word-oracle/run-edit-roundtrip.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(__dirname, "../..");
const ZERO_HASH = "0".repeat(64);
const ONE_HASH = "1".repeat(64);
const TWO_HASH = "2".repeat(64);

function paragraphBoundary(nodeIndex, offset) {
  return {
    location: { kind: "paragraph", nodeIndex },
    offset,
  };
}

const DEFAULT_RANGE = {
  start: paragraphBoundary(0, 0),
  end: paragraphBoundary(0, 5),
};

function actionEffect(action, actionIndex) {
  switch (action.type) {
    case "select-paragraph":
      return {
        kind: "selection",
        selection: { kind: "paragraph", nodeIndex: action.nodeIndex },
      };
    case "select-table-cell":
      return {
        kind: "selection",
        selection: {
          kind: "table-cell",
          tableIndex: action.tableIndex,
          rowIndex: action.rowIndex,
          cellIndex: action.cellIndex,
        },
      };
    case "set-text-range":
      return { kind: "active-text-range", range: action.range };
    case "commit-paragraph-text":
      return {
        kind: "paragraph-text",
        nodeIndex: action.nodeIndex,
        text: action.text,
      };
    case "commit-table-cell-text":
      return {
        kind: "table-cell-text",
        tableIndex: action.tableIndex,
        rowIndex: action.rowIndex,
        cellIndex: action.cellIndex,
        text: action.text,
      };
    case "toggle-bold":
    case "toggle-italic":
    case "toggle-underline":
    case "toggle-strike":
    case "set-text-color":
    case "set-highlight":
    case "set-font-family":
    case "set-font-size": {
      const properties = {
        "toggle-bold": "bold",
        "toggle-italic": "italic",
        "toggle-underline": "underline",
        "toggle-strike": "strike",
        "set-text-color": "color",
        "set-highlight": "highlight",
        "set-font-family": "fontFamily",
        "set-font-size": "fontSizePt",
      };
      const value = action.type.startsWith("toggle-")
        ? true
        : action.type === "set-font-family"
        ? action.fontFamily
        : action.type === "set-font-size"
        ? action.fontSizePt
        : action.color;
      return {
        kind: "text-style",
        range: DEFAULT_RANGE,
        property: properties[action.type],
        value,
      };
    }
    case "set-alignment":
      return {
        kind: "paragraph-style",
        location: { kind: "paragraph", nodeIndex: 0 },
        property: "alignment",
        value: action.alignment,
      };
    case "toggle-list":
      return {
        kind: "paragraph-style",
        location: { kind: "paragraph", nodeIndex: 0 },
        property: "listType",
        value: action.listType,
      };
    case "set-line-spacing":
      return {
        kind: "paragraph-style",
        location: { kind: "paragraph", nodeIndex: 0 },
        property: "lineMultiple",
        value: action.lineMultiple,
      };
    case "insert-table-row":
    case "insert-table-column":
    case "delete-table-row":
    case "delete-table-column":
      return {
        kind: "table-shape",
        tableIndex: action.tableIndex,
        rowCount: 1,
        columnCounts: [1],
      };
    case "undo":
    case "redo":
      return {
        kind: "history",
        semanticModelDigest: actionIndex.toString(16).padStart(32, "0"),
        canUndo: action.type === "redo",
        canRedo: action.type === "undo",
      };
    case "accept-tracked-change":
    case "reject-tracked-change":
      return {
        kind: "tracked-change",
        changeId: `change-${actionIndex}`,
        revisionId: String(actionIndex),
        changeKind: "insertion",
        location: { kind: "paragraph", nodeIndex: 0 },
        changeText: "tracked",
        resultText: "Edited",
        remainingCount: 0,
      };
    case "create-comment":
      return {
        kind: "created-comment",
        range: DEFAULT_RANGE,
        commentId: 1,
        id: "comment-1-paragraph-0",
        location: { kind: "paragraph", nodeIndex: 0 },
        text: action.text,
        anchorText: "Edite",
        author: action.options.author,
        initials: action.options.initials,
        date: action.options.date,
        resolved: false,
        commentCount: 1,
        resolvedCommentCount: 0,
      };
    case "set-comment-resolved":
      return {
        kind: "resolved-comment",
        commentId: action.commentId,
        id: "comment-1-paragraph-0",
        location: { kind: "paragraph", nodeIndex: 0 },
        text: "Please review",
        resolved: action.resolved,
        commentCount: 1,
        resolvedCommentCount: action.resolved ? 1 : 0,
      };
    default:
      throw new Error(`missing test effect for ${action.type}`);
  }
}

function validManifest() {
  return {
    $schema: "../../scripts/word-oracle/edit-actions.schema.json",
    schemaVersion: 1,
    corpus: { id: "word-fidelity", revision: "2026-07-09" },
    sourceManifestSha256: ZERO_HASH,
    scenarios: [
      {
        id: "edited-alpha",
        sourceCaseId: "alpha",
        expected: {
          semanticModel: "changed-from-source",
          trackedChangeCount: 0,
          commentCount: 1,
          resolvedCommentCount: 1,
          paragraphTexts: [{ nodeIndex: 0, text: "Edited" }],
        },
        actions: [
          { type: "select-paragraph", nodeIndex: 0 },
          {
            type: "set-text-range",
            range: {
              start: paragraphBoundary(0, 0),
              end: paragraphBoundary(0, 5),
            },
          },
          { type: "toggle-bold" },
          { type: "toggle-italic" },
          { type: "toggle-underline" },
          { type: "toggle-strike" },
          { type: "set-text-color", color: "#112233" },
          { type: "set-highlight", color: "#ffee00" },
          { type: "set-font-family", fontFamily: "Aptos" },
          { type: "set-font-size", fontSizePt: 12.5 },
          { type: "set-alignment", alignment: "justify" },
          { type: "toggle-list", listType: "ordered" },
          { type: "set-line-spacing", lineMultiple: 1.15 },
          { type: "commit-paragraph-text", nodeIndex: 0, text: "Edited" },
          {
            type: "select-table-cell",
            tableIndex: 1,
            rowIndex: 0,
            cellIndex: 0,
          },
          {
            type: "commit-table-cell-text",
            tableIndex: 1,
            rowIndex: 0,
            cellIndex: 0,
            text: "Cell",
          },
          {
            type: "insert-table-row",
            tableIndex: 1,
            rowIndex: 0,
            direction: "below",
          },
          {
            type: "insert-table-column",
            tableIndex: 1,
            rowIndex: 0,
            cellIndex: 0,
            direction: "right",
          },
          { type: "delete-table-row", tableIndex: 1, rowIndex: 1 },
          {
            type: "delete-table-column",
            tableIndex: 1,
            cellIndex: 1,
          },
          { type: "accept-tracked-change", index: 0 },
          { type: "reject-tracked-change", index: 0 },
          {
            type: "create-comment",
            text: "Please review",
            options: {
              author: "Ada",
              initials: "AL",
              date: "2026-07-09T12:00:00.000Z",
            },
          },
          { type: "set-comment-resolved", commentId: 1, resolved: true },
          { type: "undo" },
          { type: "redo" },
          { type: "set-text-range", range: null },
          { type: "set-text-color", color: null },
          { type: "set-highlight", color: null },
          { type: "set-alignment", alignment: null },
        ].map((action, actionIndex) => ({
          ...action,
          expect: {
            semanticModel: new Set([
              "select-paragraph",
              "select-table-cell",
              "set-text-range",
            ]).has(action.type)
              ? "unchanged"
              : "changed",
            effect: actionEffect(action, actionIndex),
          },
        })),
      },
    ],
  };
}

function sourceManifest() {
  return {
    corpus: { id: "word-fidelity", revision: "2026-07-09" },
    cases: [
      {
        id: "alpha",
        source: { path: "sources/alpha.docx", sha256: ONE_HASH },
      },
    ],
  };
}

test("accepts the complete safe edit-action subset", () => {
  const manifest = validManifest();
  assert.deepEqual(
    [...new Set(manifest.scenarios[0].actions.map((action) => action.type))].sort(),
    [...SUPPORTED_EDIT_ACTION_TYPES].sort()
  );
  assert.deepEqual(validateEditActionManifest(manifest), []);
});

test("the published JSON schema is valid JSON and versioned", async () => {
  const schema = JSON.parse(
    await fs.readFile(
      path.join(repoRoot, "scripts/word-oracle/edit-actions.schema.json"),
      "utf8"
    )
  );
  assert.equal(schema.$schema, "https://json-schema.org/draft/2020-12/schema");
  assert.equal(schema.properties.schemaVersion.const, 1);
  assert.equal(schema.additionalProperties, false);
});

test("rejects unknown actions, extra properties, and incomplete coordinates", () => {
  const manifest = validManifest();
  manifest.scenarios[0].actions = [
    { type: "run-javascript", source: "alert(1)" },
    { type: "toggle-bold", unexpected: true },
    { type: "select-table-cell", tableIndex: 1, cellIndex: 0 },
  ];
  const issues = validateEditActionManifest(manifest);
  assert.ok(issues.some((issue) => issue.path.endsWith("/0/type")));
  assert.ok(issues.some((issue) => issue.path.endsWith("/1/unexpected")));
  assert.ok(issues.some((issue) => issue.path.endsWith("/2/rowIndex")));
});

test("rejects noncanonical colors, unsafe font names, and reversed ranges", () => {
  const manifest = validManifest();
  manifest.scenarios[0].actions = [
    { type: "set-text-color", color: "red" },
    { type: "set-highlight", color: "#AABBCC" },
    { type: "set-font-family", fontFamily: " Aptos" },
    {
      type: "set-text-range",
      range: {
        start: paragraphBoundary(2, 4),
        end: paragraphBoundary(1, 4),
      },
    },
  ];
  const issues = validateEditActionManifest(manifest);
  assert.ok(issues.some((issue) => issue.path.endsWith("/0/color")));
  assert.ok(issues.some((issue) => issue.path.endsWith("/1/color")));
  assert.ok(issues.some((issue) => issue.path.endsWith("/2/fontFamily")));
  assert.ok(
    issues.some(
      (issue) => issue.path.endsWith("/3/range") && issue.message.includes("start")
    )
  );
});

test("requires explicit action and final semantic postconditions", () => {
  const manifest = validManifest();
  delete manifest.scenarios[0].expected;
  manifest.scenarios[0].actions[0].expect.semanticModel = "changed";
  delete manifest.scenarios[0].actions[2].expect;
  const issues = validateEditActionManifest(manifest);
  assert.ok(issues.some((issue) => issue.path.endsWith("/expected")));
  assert.ok(
    issues.some(
      (issue) =>
        issue.path.endsWith("/actions/0/expect/semanticModel") &&
        issue.message.includes("unchanged")
    )
  );
  assert.ok(
    issues.some((issue) => issue.path.endsWith("/actions/2/expect"))
  );
});

test("rejects missing, generic, and wrong-kind action effects", () => {
  const missing = validManifest();
  delete missing.scenarios[0].actions[0].expect.effect;
  assert.ok(
    validateEditActionManifest(missing).some((issue) =>
      issue.path.endsWith("/actions/0/expect/effect")
    )
  );

  const wrongKind = validManifest();
  wrongKind.scenarios[0].actions[0].expect.effect = {
    kind: "history",
    semanticModelDigest: "0".repeat(32),
    canUndo: false,
    canRedo: false,
  };
  assert.ok(
    validateEditActionManifest(wrongKind).some(
      (issue) =>
        issue.path.endsWith("/actions/0/expect/effect/kind") &&
        issue.message.includes("selection")
    )
  );
});

test("rejects commit and table effects that assert the wrong target", () => {
  const paragraph = validManifest();
  paragraph.scenarios[0].actions[13].expect.effect.nodeIndex = 9;
  assert.ok(
    validateEditActionManifest(paragraph).some(
      (issue) =>
        issue.path.endsWith("/actions/13/expect/effect") &&
        issue.message.includes("paragraph edit target")
    )
  );

  const table = validManifest();
  table.scenarios[0].actions[16].expect.effect.tableIndex = 4;
  assert.ok(
    validateEditActionManifest(table).some(
      (issue) =>
        issue.path.endsWith("/actions/16/expect/effect/tableIndex") &&
        issue.message.includes("action tableIndex")
    )
  );
});

test("rejects text-style effects without an exact range, property, and value", () => {
  const missingRange = validManifest();
  delete missingRange.scenarios[0].actions[2].expect.effect.range;
  assert.ok(
    validateEditActionManifest(missingRange).some((issue) =>
      issue.path.endsWith("/actions/2/expect/effect/range")
    )
  );

  const wrongProperty = validManifest();
  wrongProperty.scenarios[0].actions[2].expect.effect.property = "italic";
  assert.ok(
    validateEditActionManifest(wrongProperty).some((issue) =>
      issue.path.endsWith("/actions/2/expect/effect/property")
    )
  );

  const wrongSetterValue = validManifest();
  wrongSetterValue.scenarios[0].actions[6].expect.effect.value = "#000000";
  assert.ok(
    validateEditActionManifest(wrongSetterValue).some((issue) =>
      issue.path.endsWith("/actions/6/expect/effect/value")
    )
  );
});

test("rejects incomplete tracked-change identity and result assertions", () => {
  const manifest = validManifest();
  const effect = manifest.scenarios[0].actions[20].expect.effect;
  delete effect.revisionId;
  delete effect.resultText;
  delete effect.location;
  const issues = validateEditActionManifest(manifest);
  assert.ok(
    issues.some((issue) =>
      issue.path.endsWith("/actions/20/expect/effect/revisionId")
    )
  );
  assert.ok(
    issues.some((issue) =>
      issue.path.endsWith("/actions/20/expect/effect/resultText")
    )
  );
  assert.ok(
    issues.some((issue) =>
      issue.path.endsWith("/actions/20/expect/effect/location")
    )
  );
});

test("rejects nondeterministic comment creation and wrong resolution identity", () => {
  const creation = validManifest();
  delete creation.scenarios[0].actions[22].options.date;
  creation.scenarios[0].actions[22].expect.effect.author = "Grace";
  const creationIssues = validateEditActionManifest(creation);
  assert.ok(
    creationIssues.some((issue) =>
      issue.path.endsWith("/actions/22/options/date")
    )
  );
  assert.ok(
    creationIssues.some(
      (issue) =>
        issue.path.endsWith("/actions/22/expect/effect") &&
        issue.message.includes("requested comment")
    )
  );

  const resolution = validManifest();
  resolution.scenarios[0].actions[23].expect.effect.commentId = 99;
  assert.ok(
    validateEditActionManifest(resolution).some(
      (issue) =>
        issue.path.endsWith("/actions/23/expect/effect") &&
        issue.message.includes("resolution target")
    )
  );
});

test("rejects history effects without exact digest and capability state", () => {
  const manifest = validManifest();
  manifest.scenarios[0].actions[24].expect.effect.semanticModelDigest = "abc";
  delete manifest.scenarios[0].actions[24].expect.effect.canRedo;
  const issues = validateEditActionManifest(manifest);
  assert.ok(
    issues.some((issue) =>
      issue.path.endsWith("/actions/24/expect/effect/semanticModelDigest")
    )
  );
  assert.ok(
    issues.some((issue) =>
      issue.path.endsWith("/actions/24/expect/effect/canRedo")
    )
  );
});

function semanticSummary(digest) {
  return {
    semanticModelVersion: 1,
    semanticModelDigest: digest,
  };
}

function syntheticActionState(digest) {
  return {
    selection: { kind: "paragraph", nodeIndex: 0 },
    activeTextRange: null,
    selectedParagraphLocation: { kind: "paragraph", nodeIndex: 0 },
    semanticModelDigest: digest,
    canUndo: true,
    canRedo: false,
  };
}

test("runner rejects a wrong-target paragraph mutation despite a changed global digest", async () => {
  const action = {
    type: "commit-paragraph-text",
    nodeIndex: 0,
    text: "Expected",
    expect: {
      semanticModel: "changed",
      effect: { kind: "paragraph-text", nodeIndex: 0, text: "Expected" },
    },
  };
  await assert.rejects(
    verifyActionPostcondition(
      action,
      semanticSummary("0".repeat(32)),
      semanticSummary("1".repeat(32)),
      {
        action: syntheticActionState("0".repeat(32)),
        location: {
          location: { kind: "paragraph", nodeIndex: 0 },
          text: "Before",
          digest: "a".repeat(32),
        },
      },
      {
        action: syntheticActionState("1".repeat(32)),
        location: {
          location: { kind: "paragraph", nodeIndex: 0 },
          text: "Before",
          digest: "a".repeat(32),
        },
      }
    ),
    /paragraph 0 text/
  );
});

test("runner derives the exact toggle-list result from the before state", async () => {
  const location = { kind: "paragraph", nodeIndex: 0 };
  const action = {
    type: "toggle-list",
    listType: "ordered",
    expect: {
      semanticModel: "changed",
      effect: {
        kind: "paragraph-style",
        location,
        property: "listType",
        value: null,
      },
    },
  };
  await assert.rejects(
    verifyActionPostcondition(
      action,
      semanticSummary("0".repeat(32)),
      semanticSummary("1".repeat(32)),
      {
        action: syntheticActionState("0".repeat(32)),
        location: {
          location,
          text: "Item",
          digest: "a".repeat(32),
          listType: "unordered",
        },
      },
      {
        action: syntheticActionState("1".repeat(32)),
        location: {
          location,
          text: "Item",
          digest: "b".repeat(32),
          listType: null,
        },
      }
    ),
    /toggle-list expectation must be "ordered"/
  );
});

test("runner rejects wrong tracked-change and comment identities", async () => {
  const beforeSummary = semanticSummary("0".repeat(32));
  const afterSummary = semanticSummary("1".repeat(32));
  const location = { kind: "paragraph", nodeIndex: 0 };
  const locationState = { location, text: "Body", digest: "a".repeat(32) };

  const trackedAction = {
    type: "accept-tracked-change",
    index: 0,
    expect: {
      semanticModel: "changed",
      effect: {
        kind: "tracked-change",
        changeId: "expected-change",
        revisionId: "7",
        changeKind: "insertion",
        location,
        changeText: "Body",
        resultText: "Body",
        remainingCount: 0,
      },
    },
  };
  await assert.rejects(
    verifyActionPostcondition(
      trackedAction,
      beforeSummary,
      afterSummary,
      {
        action: syntheticActionState("0".repeat(32)),
        location: locationState,
        trackedChanges: [
          {
            id: "wrong-change",
            revisionId: "7",
            kind: "insertion",
            location,
            text: "Body",
          },
        ],
      },
      {
        action: syntheticActionState("1".repeat(32)),
        location: { ...locationState, digest: "b".repeat(32) },
        trackedChanges: [],
      }
    ),
    /tracked change id/
  );

  const commentAction = {
    type: "set-comment-resolved",
    commentId: 1,
    resolved: true,
    expect: {
      semanticModel: "changed",
      effect: {
        kind: "resolved-comment",
        commentId: 1,
        id: "expected-comment",
        location,
        text: "Review",
        resolved: true,
        commentCount: 1,
        resolvedCommentCount: 1,
      },
    },
  };
  await assert.rejects(
    verifyActionPostcondition(
      commentAction,
      beforeSummary,
      afterSummary,
      {
        action: syntheticActionState("0".repeat(32)),
        location: locationState,
        comments: [
          {
            id: "expected-comment",
            commentId: 1,
            location,
            text: "Review",
            resolved: false,
          },
        ],
      },
      {
        action: syntheticActionState("1".repeat(32)),
        location: locationState,
        comments: [
          {
            id: "wrong-comment",
            commentId: 1,
            location,
            text: "Review",
            resolved: true,
          },
        ],
      }
    ),
    /comment 1 after state/
  );
});

test("semantic postconditions reject no-op mutations", () => {
  const before = {
    semanticModelVersion: 1,
    semanticModelDigest: "0".repeat(32),
  };
  const unchanged = structuredClone(before);
  const changed = {
    semanticModelVersion: 1,
    semanticModelDigest: "1".repeat(32),
  };
  assert.throws(
    () => assertSemanticModelExpectation("changed", before, unchanged),
    /expected semantic model to change/
  );
  assert.throws(
    () => assertSemanticModelExpectation("unchanged", before, changed),
    /remain unchanged/
  );
  assert.equal(
    assertSemanticModelExpectation("changed-from-source", before, changed)
      .changed,
    true
  );
});

test("binds scenarios to exact source-manifest bytes and case IDs", () => {
  assert.deepEqual(
    bindEditScenarios(validManifest(), sourceManifest(), ZERO_HASH),
    [
      {
        ...validManifest().scenarios[0],
        source: { path: "sources/alpha.docx", sha256: ONE_HASH },
      },
    ]
  );

  assert.throws(
    () => bindEditScenarios(validManifest(), sourceManifest(), TWO_HASH),
    /manifest bytes/
  );
  const missing = validManifest();
  missing.scenarios[0].sourceCaseId = "missing";
  assert.throws(
    () => bindEditScenarios(missing, sourceManifest(), ZERO_HASH),
    /unknown source case/
  );

  const conflicting = validManifest();
  conflicting.scenarios[0].id = "alpha";
  assert.throws(
    () => bindEditScenarios(conflicting, sourceManifest(), ZERO_HASH),
    /conflicts with an existing source case ID/
  );
});

test("results expose an exported source fragment ready for Word registration", () => {
  const summary = {
    fileName: "edited-alpha.docx",
    status: "Loaded edited-alpha.docx",
    semanticModelVersion: 1,
    semanticModelDigest: "0123456789abcdef0123456789abcdef",
    bodyNodeCount: 1,
    bodyParagraphNodeCount: 1,
    bodyTableNodeCount: 0,
    paragraphCount: 1,
    tableCount: 0,
    imageCount: 0,
    formFieldCount: 0,
    trackedChangeCount: 0,
    commentCount: 1,
    resolvedCommentCount: 1,
    sectionCount: 1,
    headerSectionCount: 0,
    footerSectionCount: 0,
    hasTitlePageSection: false,
    canUndo: false,
    canRedo: false,
    firstParagraph: { nodeIndex: 0, textLength: 6 },
  };
  const manifest = buildEditResultsManifest({
    corpus: validManifest().corpus,
    sourceManifestSha256: ZERO_HASH,
    actionManifestSha256: ONE_HASH,
    runner: { name: "runner", version: "1" },
    scenarios: [
      {
        id: "edited-alpha",
        sourceCaseId: "alpha",
        actions: { count: 1, sha256: TWO_HASH, outcomes: [] },
        inputSource: { sha256: ONE_HASH },
        exportedDocx: {
          path: "edited-alpha/edited-alpha.docx",
          sha256: TWO_HASH,
          mediaType:
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
          sizeBytes: 123,
          suggestedFileName: "alpha-edited.docx",
        },
        sourceSummary: summary,
        editedSummary: summary,
        reopenedSummary: summary,
        reopenVerification: {
          importSucceeded: true,
          structuralSummaryMatched: true,
        },
      },
    ],
  });

  assert.deepEqual(manifest.scenarios[0].wordOracleRegistration, {
    caseId: "edited-alpha",
    source: {
      path: "edited-alpha/edited-alpha.docx",
      sha256: TWO_HASH,
    },
  });
  assert.deepEqual(comparableEditSummary(summary), {
    semanticModelVersion: 1,
    semanticModelDigest: "0123456789abcdef0123456789abcdef",
    bodyNodeCount: 1,
    bodyParagraphNodeCount: 1,
    bodyTableNodeCount: 0,
    paragraphCount: 1,
    tableCount: 0,
    imageCount: 0,
    formFieldCount: 0,
    trackedChangeCount: 0,
    commentCount: 1,
    resolvedCommentCount: 1,
    sectionCount: 1,
    headerSectionCount: 0,
    footerSectionCount: 0,
    hasTitlePageSection: false,
    firstParagraph: { nodeIndex: 0, textLength: 6 },
    firstTableCell: undefined,
  });
});
