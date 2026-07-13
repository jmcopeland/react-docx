import { isPortableRelativePath } from "./contract.mjs";

export const EDIT_ACTION_SCHEMA_VERSION = 1;
export const EDIT_RESULTS_SCHEMA_VERSION = 1;

const ID_PATTERN = /^[a-z0-9][a-z0-9._-]{0,127}$/;
const SHA256_PATTERN = /^[a-f0-9]{64}$/;
const COLOR_PATTERN = /^#[a-f0-9]{6}$/;
const SEMANTIC_DIGEST_PATTERN = /^[a-f0-9]{32}$/;
const MAX_INDEX = 1_000_000;
const MAX_OFFSET = 10_000_000;
const MAX_TEXT_LENGTH = 1_000_000;

const ACTION_KEYS = new Map([
  ["select-paragraph", new Set(["type", "nodeIndex"])],
  [
    "select-table-cell",
    new Set(["type", "tableIndex", "rowIndex", "cellIndex"]),
  ],
  ["set-text-range", new Set(["type", "range"])],
  [
    "commit-paragraph-text",
    new Set(["type", "nodeIndex", "text"]),
  ],
  [
    "commit-table-cell-text",
    new Set(["type", "tableIndex", "rowIndex", "cellIndex", "text"]),
  ],
  ["toggle-bold", new Set(["type"])],
  ["toggle-italic", new Set(["type"])],
  ["toggle-underline", new Set(["type"])],
  ["toggle-strike", new Set(["type"])],
  ["set-text-color", new Set(["type", "color"])],
  ["set-highlight", new Set(["type", "color"])],
  ["set-font-family", new Set(["type", "fontFamily"])],
  ["set-font-size", new Set(["type", "fontSizePt"])],
  ["set-alignment", new Set(["type", "alignment"])],
  ["toggle-list", new Set(["type", "listType"])],
  ["set-line-spacing", new Set(["type", "lineMultiple"])],
  [
    "insert-table-row",
    new Set(["type", "tableIndex", "rowIndex", "direction"]),
  ],
  [
    "insert-table-column",
    new Set([
      "type",
      "tableIndex",
      "cellIndex",
      "direction",
      "rowIndex",
    ]),
  ],
  ["delete-table-row", new Set(["type", "tableIndex", "rowIndex"])],
  [
    "delete-table-column",
    new Set(["type", "tableIndex", "cellIndex", "rowIndex"]),
  ],
  ["undo", new Set(["type"])],
  ["redo", new Set(["type"])],
  ["accept-tracked-change", new Set(["type", "index"])],
  ["reject-tracked-change", new Set(["type", "index"])],
  ["create-comment", new Set(["type", "text", "options"])],
  [
    "set-comment-resolved",
    new Set(["type", "commentId", "resolved"]),
  ],
]);

for (const allowedKeys of ACTION_KEYS.values()) {
  allowedKeys.add("expect");
}

const NON_MUTATING_ACTION_TYPES = new Set([
  "select-paragraph",
  "select-table-cell",
  "set-text-range",
]);

const ACTION_EFFECT_KIND = new Map([
  ["select-paragraph", "selection"],
  ["select-table-cell", "selection"],
  ["set-text-range", "active-text-range"],
  ["commit-paragraph-text", "paragraph-text"],
  ["commit-table-cell-text", "table-cell-text"],
  ["toggle-bold", "text-style"],
  ["toggle-italic", "text-style"],
  ["toggle-underline", "text-style"],
  ["toggle-strike", "text-style"],
  ["set-text-color", "text-style"],
  ["set-highlight", "text-style"],
  ["set-font-family", "text-style"],
  ["set-font-size", "text-style"],
  ["set-alignment", "paragraph-style"],
  ["toggle-list", "paragraph-style"],
  ["set-line-spacing", "paragraph-style"],
  ["insert-table-row", "table-shape"],
  ["insert-table-column", "table-shape"],
  ["delete-table-row", "table-shape"],
  ["delete-table-column", "table-shape"],
  ["undo", "history"],
  ["redo", "history"],
  ["accept-tracked-change", "tracked-change"],
  ["reject-tracked-change", "tracked-change"],
  ["create-comment", "created-comment"],
  ["set-comment-resolved", "resolved-comment"],
]);

const TEXT_STYLE_PROPERTIES = new Map([
  ["toggle-bold", "bold"],
  ["toggle-italic", "italic"],
  ["toggle-underline", "underline"],
  ["toggle-strike", "strike"],
  ["set-text-color", "color"],
  ["set-highlight", "highlight"],
  ["set-font-family", "fontFamily"],
  ["set-font-size", "fontSizePt"],
]);

const PARAGRAPH_STYLE_PROPERTIES = new Map([
  ["set-alignment", "alignment"],
  ["toggle-list", "listType"],
  ["set-line-spacing", "lineMultiple"],
]);

export const SUPPORTED_EDIT_ACTION_TYPES = Object.freeze([...ACTION_KEYS.keys()]);

function isRecord(value) {
  return value !== null && typeof value === "object" && !Array.isArray(value);
}

function addIssue(issues, issuePath, message) {
  issues.push({ path: issuePath, message });
}

function requireRecord(issues, value, issuePath) {
  if (!isRecord(value)) {
    addIssue(issues, issuePath, "must be an object");
    return false;
  }
  return true;
}

function checkAllowedKeys(issues, value, issuePath, allowedKeys) {
  if (!isRecord(value)) {
    return;
  }
  for (const key of Object.keys(value)) {
    if (!allowedKeys.has(key)) {
      addIssue(issues, `${issuePath}/${key}`, "is not a supported property");
    }
  }
}

function requireString(issues, value, issuePath, options = {}) {
  if (typeof value !== "string" || value.length === 0) {
    addIssue(issues, issuePath, "must be a non-empty string");
    return false;
  }
  if (options.maxLength !== undefined && value.length > options.maxLength) {
    addIssue(
      issues,
      issuePath,
      `must contain at most ${options.maxLength} characters`
    );
  }
  if (options.pattern && !options.pattern.test(value)) {
    addIssue(issues, issuePath, options.patternMessage ?? "has an invalid format");
  }
  return true;
}

function requireId(issues, value, issuePath) {
  requireString(issues, value, issuePath, {
    pattern: ID_PATTERN,
    patternMessage:
      "must start with a lowercase letter or digit and contain only lowercase letters, digits, dots, underscores, or hyphens",
  });
}

function requireSha256(issues, value, issuePath) {
  requireString(issues, value, issuePath, {
    pattern: SHA256_PATTERN,
    patternMessage: "must be a lowercase 64-character SHA-256 digest",
  });
}

function requireIndex(issues, value, issuePath, maximum = MAX_INDEX) {
  if (!Number.isInteger(value) || value < 0 || value > maximum) {
    addIssue(
      issues,
      issuePath,
      `must be an integer between zero and ${maximum}`
    );
  }
}

function requireFiniteRange(issues, value, issuePath, minimum, maximum) {
  if (
    typeof value !== "number" ||
    !Number.isFinite(value) ||
    value < minimum ||
    value > maximum
  ) {
    addIssue(
      issues,
      issuePath,
      `must be a finite number between ${minimum} and ${maximum}`
    );
  }
}

function validateCorpus(issues, corpus) {
  if (!requireRecord(issues, corpus, "/corpus")) {
    return;
  }
  checkAllowedKeys(
    issues,
    corpus,
    "/corpus",
    new Set(["id", "revision", "description"])
  );
  requireId(issues, corpus.id, "/corpus/id");
  requireString(issues, corpus.revision, "/corpus/revision");
  if (corpus.description !== undefined) {
    requireString(issues, corpus.description, "/corpus/description");
  }
}

function validateLocation(issues, location, issuePath) {
  if (!requireRecord(issues, location, issuePath)) {
    return;
  }
  if (location.kind === "paragraph") {
    checkAllowedKeys(
      issues,
      location,
      issuePath,
      new Set(["kind", "nodeIndex"])
    );
    requireIndex(issues, location.nodeIndex, `${issuePath}/nodeIndex`);
    return;
  }
  if (location.kind === "table-cell") {
    checkAllowedKeys(
      issues,
      location,
      issuePath,
      new Set([
        "kind",
        "tableIndex",
        "rowIndex",
        "cellIndex",
        "paragraphIndex",
      ])
    );
    for (const key of [
      "tableIndex",
      "rowIndex",
      "cellIndex",
      "paragraphIndex",
    ]) {
      requireIndex(issues, location[key], `${issuePath}/${key}`);
    }
    return;
  }
  addIssue(
    issues,
    `${issuePath}/kind`,
    'must be "paragraph" or "table-cell"'
  );
}

function validateBoundary(issues, boundary, issuePath) {
  if (!requireRecord(issues, boundary, issuePath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    boundary,
    issuePath,
    new Set(["location", "offset"])
  );
  validateLocation(issues, boundary.location, `${issuePath}/location`);
  requireIndex(issues, boundary.offset, `${issuePath}/offset`, MAX_OFFSET);
}

function boundaryTuple(boundary) {
  const location = boundary?.location;
  if (location?.kind === "paragraph") {
    return [location.nodeIndex, -1, -1, -1, boundary.offset];
  }
  if (location?.kind === "table-cell") {
    return [
      location.tableIndex,
      location.rowIndex,
      location.cellIndex,
      location.paragraphIndex,
      boundary.offset,
    ];
  }
  return undefined;
}

function compareTuples(left, right) {
  for (let index = 0; index < left.length; index += 1) {
    if (left[index] !== right[index]) {
      return left[index] - right[index];
    }
  }
  return 0;
}

function validateRange(issues, range, issuePath) {
  if (range === null) {
    return;
  }
  if (!requireRecord(issues, range, issuePath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    range,
    issuePath,
    new Set(["start", "end"])
  );
  validateBoundary(issues, range.start, `${issuePath}/start`);
  validateBoundary(issues, range.end, `${issuePath}/end`);
  const startTuple = boundaryTuple(range.start);
  const endTuple = boundaryTuple(range.end);
  if (startTuple && endTuple && compareTuples(startTuple, endTuple) > 0) {
    addIssue(issues, issuePath, "start must not follow end in document order");
  }
}

function validateColor(issues, value, issuePath) {
  if (value === null) {
    return;
  }
  requireString(issues, value, issuePath, {
    pattern: COLOR_PATTERN,
    patternMessage: "must be null or a lowercase #rrggbb color",
  });
}

function validateText(issues, value, issuePath) {
  if (typeof value !== "string") {
    addIssue(issues, issuePath, "must be a string");
    return;
  }
  if (value.length > MAX_TEXT_LENGTH) {
    addIssue(
      issues,
      issuePath,
      `must contain at most ${MAX_TEXT_LENGTH} characters`
    );
  }
}

function validateTableCoordinates(
  issues,
  action,
  issuePath,
  keys,
  optionalKeys = new Set()
) {
  for (const key of keys) {
    if (optionalKeys.has(key) && action[key] === undefined) {
      continue;
    }
    requireIndex(issues, action[key], `${issuePath}/${key}`);
  }
}

function sameJson(left, right) {
  return JSON.stringify(left) === JSON.stringify(right);
}

function requireBoolean(issues, value, issuePath) {
  if (typeof value !== "boolean") {
    addIssue(issues, issuePath, "must be a boolean");
  }
}

function requireSemanticDigest(issues, value, issuePath) {
  requireString(issues, value, issuePath, {
    pattern: SEMANTIC_DIGEST_PATTERN,
    patternMessage: "must be a lowercase 32-character semantic model digest",
  });
}

function validateSelection(issues, selection, issuePath) {
  if (!requireRecord(issues, selection, issuePath)) {
    return;
  }
  if (selection.kind === "paragraph") {
    checkAllowedKeys(
      issues,
      selection,
      issuePath,
      new Set(["kind", "nodeIndex"])
    );
    requireIndex(issues, selection.nodeIndex, `${issuePath}/nodeIndex`);
    return;
  }
  if (selection.kind === "table-cell") {
    checkAllowedKeys(
      issues,
      selection,
      issuePath,
      new Set(["kind", "tableIndex", "rowIndex", "cellIndex"])
    );
    validateTableCoordinates(issues, selection, issuePath, [
      "tableIndex",
      "rowIndex",
      "cellIndex",
    ]);
    return;
  }
  addIssue(
    issues,
    `${issuePath}/kind`,
    'must be "paragraph" or "table-cell"'
  );
}

function validateExpandedSingleLocationRange(issues, range, issuePath) {
  validateRange(issues, range, issuePath);
  if (!isRecord(range)) {
    addIssue(issues, issuePath, "must be a non-null expanded text range");
    return;
  }
  const startLocation = range.start?.location;
  const endLocation = range.end?.location;
  if (startLocation && endLocation && !sameJson(startLocation, endLocation)) {
    addIssue(issues, issuePath, "must target exactly one paragraph location");
  }
  if (
    Number.isInteger(range.start?.offset) &&
    Number.isInteger(range.end?.offset) &&
    range.start.offset >= range.end.offset
  ) {
    addIssue(issues, issuePath, "must be an expanded range with start before end");
  }
}

function validateEffectTextStyle(issues, action, effect, issuePath) {
  checkAllowedKeys(
    issues,
    effect,
    issuePath,
    new Set(["kind", "range", "property", "value"])
  );
  validateExpandedSingleLocationRange(issues, effect.range, `${issuePath}/range`);
  const expectedProperty = TEXT_STYLE_PROPERTIES.get(action.type);
  if (effect.property !== expectedProperty) {
    addIssue(
      issues,
      `${issuePath}/property`,
      `must equal "${expectedProperty}" for ${action.type}`
    );
  }
  if (action.type.startsWith("toggle-")) {
    requireBoolean(issues, effect.value, `${issuePath}/value`);
    return;
  }
  const expectedValue =
    action.type === "set-font-family"
      ? action.fontFamily
      : action.type === "set-font-size"
      ? action.fontSizePt
      : action.color;
  if (!sameJson(effect.value, expectedValue)) {
    addIssue(
      issues,
      `${issuePath}/value`,
      `must exactly match the value requested by ${action.type}`
    );
  }
}

function validateEffectParagraphStyle(issues, action, effect, issuePath) {
  checkAllowedKeys(
    issues,
    effect,
    issuePath,
    new Set(["kind", "location", "property", "value"])
  );
  validateLocation(issues, effect.location, `${issuePath}/location`);
  const expectedProperty = PARAGRAPH_STYLE_PROPERTIES.get(action.type);
  if (effect.property !== expectedProperty) {
    addIssue(
      issues,
      `${issuePath}/property`,
      `must equal "${expectedProperty}" for ${action.type}`
    );
  }
  if (action.type === "toggle-list") {
    if (effect.value !== null && effect.value !== action.listType) {
      addIssue(
        issues,
        `${issuePath}/value`,
        `must be null or "${action.listType}" for toggle-list`
      );
    }
    return;
  }
  const expectedValue =
    action.type === "set-alignment" ? action.alignment : action.lineMultiple;
  if (!sameJson(effect.value, expectedValue)) {
    addIssue(
      issues,
      `${issuePath}/value`,
      `must exactly match the value requested by ${action.type}`
    );
  }
}

function validateActionEffect(issues, action, effect, issuePath) {
  if (!requireRecord(issues, effect, issuePath)) {
    return;
  }
  const requiredKind = ACTION_EFFECT_KIND.get(action.type);
  if (effect.kind !== requiredKind) {
    addIssue(
      issues,
      `${issuePath}/kind`,
      `must equal "${requiredKind}" for ${action.type}`
    );
    return;
  }

  switch (effect.kind) {
    case "selection": {
      checkAllowedKeys(
        issues,
        effect,
        issuePath,
        new Set(["kind", "selection"])
      );
      validateSelection(issues, effect.selection, `${issuePath}/selection`);
      const expectedSelection =
        action.type === "select-paragraph"
          ? { kind: "paragraph", nodeIndex: action.nodeIndex }
          : {
              kind: "table-cell",
              tableIndex: action.tableIndex,
              rowIndex: action.rowIndex,
              cellIndex: action.cellIndex,
            };
      if (!sameJson(effect.selection, expectedSelection)) {
        addIssue(
          issues,
          `${issuePath}/selection`,
          "must exactly match the action target"
        );
      }
      break;
    }
    case "active-text-range":
      checkAllowedKeys(
        issues,
        effect,
        issuePath,
        new Set(["kind", "range"])
      );
      validateRange(issues, effect.range, `${issuePath}/range`);
      if (!sameJson(effect.range, action.range)) {
        addIssue(
          issues,
          `${issuePath}/range`,
          "must exactly match the action range"
        );
      }
      break;
    case "paragraph-text":
      checkAllowedKeys(
        issues,
        effect,
        issuePath,
        new Set(["kind", "nodeIndex", "text"])
      );
      requireIndex(issues, effect.nodeIndex, `${issuePath}/nodeIndex`);
      validateText(issues, effect.text, `${issuePath}/text`);
      if (effect.nodeIndex !== action.nodeIndex || effect.text !== action.text) {
        addIssue(issues, issuePath, "must exactly match the paragraph edit target and text");
      }
      break;
    case "table-cell-text":
      checkAllowedKeys(
        issues,
        effect,
        issuePath,
        new Set(["kind", "tableIndex", "rowIndex", "cellIndex", "text"])
      );
      validateTableCoordinates(issues, effect, issuePath, [
        "tableIndex",
        "rowIndex",
        "cellIndex",
      ]);
      validateText(issues, effect.text, `${issuePath}/text`);
      if (
        effect.tableIndex !== action.tableIndex ||
        effect.rowIndex !== action.rowIndex ||
        effect.cellIndex !== action.cellIndex ||
        effect.text !== action.text
      ) {
        addIssue(issues, issuePath, "must exactly match the table-cell edit target and text");
      }
      break;
    case "text-style":
      validateEffectTextStyle(issues, action, effect, issuePath);
      break;
    case "paragraph-style":
      validateEffectParagraphStyle(issues, action, effect, issuePath);
      break;
    case "table-shape":
      validateExpectedTableShape(issues, effect, issuePath, true);
      if (effect.tableIndex !== action.tableIndex) {
        addIssue(
          issues,
          `${issuePath}/tableIndex`,
          "must exactly match the action tableIndex"
        );
      }
      break;
    case "history":
      checkAllowedKeys(
        issues,
        effect,
        issuePath,
        new Set(["kind", "semanticModelDigest", "canUndo", "canRedo"])
      );
      requireSemanticDigest(
        issues,
        effect.semanticModelDigest,
        `${issuePath}/semanticModelDigest`
      );
      requireBoolean(issues, effect.canUndo, `${issuePath}/canUndo`);
      requireBoolean(issues, effect.canRedo, `${issuePath}/canRedo`);
      break;
    case "tracked-change":
      checkAllowedKeys(
        issues,
        effect,
        issuePath,
        new Set([
          "kind",
          "changeId",
          "revisionId",
          "changeKind",
          "location",
          "changeText",
          "resultText",
          "remainingCount",
        ])
      );
      requireString(issues, effect.changeId, `${issuePath}/changeId`);
      requireString(issues, effect.revisionId, `${issuePath}/revisionId`);
      if (
        ![
          "insertion",
          "deletion",
          "move-from",
          "move-to",
          "format-change",
          "paragraph-format-change",
        ].includes(effect.changeKind)
      ) {
        addIssue(issues, `${issuePath}/changeKind`, "must be a supported tracked-change kind");
      }
      validateLocation(issues, effect.location, `${issuePath}/location`);
      validateText(issues, effect.changeText, `${issuePath}/changeText`);
      validateText(issues, effect.resultText, `${issuePath}/resultText`);
      requireIndex(issues, effect.remainingCount, `${issuePath}/remainingCount`);
      break;
    case "created-comment":
      checkAllowedKeys(
        issues,
        effect,
        issuePath,
        new Set([
          "kind",
          "range",
          "commentId",
          "id",
          "location",
          "text",
          "anchorText",
          "author",
          "initials",
          "date",
          "resolved",
          "commentCount",
          "resolvedCommentCount",
        ])
      );
      validateExpandedSingleLocationRange(issues, effect.range, `${issuePath}/range`);
      requireIndex(issues, effect.commentId, `${issuePath}/commentId`);
      requireString(issues, effect.id, `${issuePath}/id`);
      validateLocation(issues, effect.location, `${issuePath}/location`);
      validateText(issues, effect.text, `${issuePath}/text`);
      validateText(issues, effect.anchorText, `${issuePath}/anchorText`);
      for (const key of ["author", "initials", "date"]) {
        requireString(issues, effect[key], `${issuePath}/${key}`);
      }
      requireBoolean(issues, effect.resolved, `${issuePath}/resolved`);
      requireIndex(issues, effect.commentCount, `${issuePath}/commentCount`);
      requireIndex(
        issues,
        effect.resolvedCommentCount,
        `${issuePath}/resolvedCommentCount`
      );
      if (
        effect.text !== action.text ||
        effect.author !== action.options?.author ||
        effect.initials !== action.options?.initials ||
        effect.date !== action.options?.date ||
        effect.resolved !== false ||
        !sameJson(effect.location, effect.range?.start?.location)
      ) {
        addIssue(
          issues,
          issuePath,
          "must exactly describe the requested comment and its selected anchor"
        );
      }
      break;
    case "resolved-comment":
      checkAllowedKeys(
        issues,
        effect,
        issuePath,
        new Set([
          "kind",
          "commentId",
          "id",
          "location",
          "text",
          "resolved",
          "commentCount",
          "resolvedCommentCount",
        ])
      );
      requireIndex(issues, effect.commentId, `${issuePath}/commentId`);
      requireString(issues, effect.id, `${issuePath}/id`);
      validateLocation(issues, effect.location, `${issuePath}/location`);
      validateText(issues, effect.text, `${issuePath}/text`);
      requireBoolean(issues, effect.resolved, `${issuePath}/resolved`);
      requireIndex(issues, effect.commentCount, `${issuePath}/commentCount`);
      requireIndex(
        issues,
        effect.resolvedCommentCount,
        `${issuePath}/resolvedCommentCount`
      );
      if (
        effect.commentId !== action.commentId ||
        effect.resolved !== action.resolved
      ) {
        addIssue(
          issues,
          issuePath,
          "must exactly match the requested comment resolution target and state"
        );
      }
      break;
    default:
      addIssue(issues, `${issuePath}/kind`, "is not a supported action effect");
  }
}

function validateActionExpectation(issues, action, issuePath) {
  const expectationPath = `${issuePath}/expect`;
  if (!requireRecord(issues, action.expect, expectationPath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    action.expect,
    expectationPath,
    new Set(["semanticModel", "effect"])
  );
  const requiredValue = NON_MUTATING_ACTION_TYPES.has(action.type)
    ? "unchanged"
    : "changed";
  if (action.expect.semanticModel !== requiredValue) {
    addIssue(
      issues,
      `${expectationPath}/semanticModel`,
      `must equal "${requiredValue}" for ${action.type}`
    );
  }
  validateActionEffect(
    issues,
    action,
    action.expect.effect,
    `${expectationPath}/effect`
  );
}

function validateExpectedParagraphText(issues, expectation, issuePath) {
  if (!requireRecord(issues, expectation, issuePath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    expectation,
    issuePath,
    new Set(["nodeIndex", "text"])
  );
  requireIndex(issues, expectation.nodeIndex, `${issuePath}/nodeIndex`);
  validateText(issues, expectation.text, `${issuePath}/text`);
}

function validateExpectedTableCellText(issues, expectation, issuePath) {
  if (!requireRecord(issues, expectation, issuePath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    expectation,
    issuePath,
    new Set(["tableIndex", "rowIndex", "cellIndex", "text"])
  );
  validateTableCoordinates(issues, expectation, issuePath, [
    "tableIndex",
    "rowIndex",
    "cellIndex",
  ]);
  validateText(issues, expectation.text, `${issuePath}/text`);
}

function validateExpectedTableShape(
  issues,
  expectation,
  issuePath,
  allowKind = false
) {
  if (!requireRecord(issues, expectation, issuePath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    expectation,
    issuePath,
    new Set([
      ...(allowKind ? ["kind"] : []),
      "tableIndex",
      "rowCount",
      "columnCounts",
    ])
  );
  requireIndex(issues, expectation.tableIndex, `${issuePath}/tableIndex`);
  requireIndex(issues, expectation.rowCount, `${issuePath}/rowCount`);
  if (
    !Array.isArray(expectation.columnCounts) ||
    expectation.columnCounts.length !== expectation.rowCount
  ) {
    addIssue(
      issues,
      `${issuePath}/columnCounts`,
      "must be an array containing one count for every row"
    );
  } else {
    expectation.columnCounts.forEach((count, index) =>
      requireIndex(issues, count, `${issuePath}/columnCounts/${index}`)
    );
  }
}

function validateScenarioExpectation(issues, expected, issuePath) {
  if (!requireRecord(issues, expected, issuePath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    expected,
    issuePath,
    new Set([
      "semanticModel",
      "semanticModelDigest",
      "paragraphTexts",
      "tableCellTexts",
      "tableShapes",
      "trackedChangeCount",
      "commentCount",
      "resolvedCommentCount",
    ])
  );
  if (
    !["changed-from-source", "same-as-source"].includes(expected.semanticModel)
  ) {
    addIssue(
      issues,
      `${issuePath}/semanticModel`,
      'must be "changed-from-source" or "same-as-source"'
    );
  }
  if (expected.semanticModelDigest !== undefined) {
    requireString(issues, expected.semanticModelDigest, `${issuePath}/semanticModelDigest`, {
      pattern: SEMANTIC_DIGEST_PATTERN,
      patternMessage: "must be a lowercase 32-character semantic model digest",
    });
  }
  for (const field of [
    "trackedChangeCount",
    "commentCount",
    "resolvedCommentCount",
  ]) {
    if (expected[field] !== undefined) {
      requireIndex(issues, expected[field], `${issuePath}/${field}`);
    }
  }
  for (const [field, validator] of [
    ["paragraphTexts", validateExpectedParagraphText],
    ["tableCellTexts", validateExpectedTableCellText],
    ["tableShapes", validateExpectedTableShape],
  ]) {
    if (expected[field] === undefined) {
      continue;
    }
    if (!Array.isArray(expected[field])) {
      addIssue(issues, `${issuePath}/${field}`, "must be an array");
      continue;
    }
    expected[field].forEach((entry, index) =>
      validator(issues, entry, `${issuePath}/${field}/${index}`)
    );
  }
}

function validateAction(issues, action, issuePath) {
  if (!requireRecord(issues, action, issuePath)) {
    return;
  }
  if (typeof action.type !== "string" || !ACTION_KEYS.has(action.type)) {
    addIssue(
      issues,
      `${issuePath}/type`,
      `must be one of: ${SUPPORTED_EDIT_ACTION_TYPES.join(", ")}`
    );
    return;
  }
  checkAllowedKeys(issues, action, issuePath, ACTION_KEYS.get(action.type));
  validateActionExpectation(issues, action, issuePath);

  switch (action.type) {
    case "select-paragraph":
    case "commit-paragraph-text":
      requireIndex(issues, action.nodeIndex, `${issuePath}/nodeIndex`);
      if (action.type === "commit-paragraph-text") {
        validateText(issues, action.text, `${issuePath}/text`);
      }
      break;
    case "select-table-cell":
    case "commit-table-cell-text":
      validateTableCoordinates(issues, action, issuePath, [
        "tableIndex",
        "rowIndex",
        "cellIndex",
      ]);
      if (action.type === "commit-table-cell-text") {
        validateText(issues, action.text, `${issuePath}/text`);
      }
      break;
    case "set-text-range":
      validateRange(issues, action.range, `${issuePath}/range`);
      break;
    case "set-text-color":
    case "set-highlight":
      validateColor(issues, action.color, `${issuePath}/color`);
      break;
    case "set-font-family":
      if (
        requireString(issues, action.fontFamily, `${issuePath}/fontFamily`, {
          maxLength: 128,
        }) &&
        (action.fontFamily.trim() !== action.fontFamily ||
          /[\u0000-\u001f\u007f]/.test(action.fontFamily))
      ) {
        addIssue(
          issues,
          `${issuePath}/fontFamily`,
          "must be trimmed and contain no control characters"
        );
      }
      break;
    case "set-font-size":
      requireFiniteRange(
        issues,
        action.fontSizePt,
        `${issuePath}/fontSizePt`,
        1,
        400
      );
      break;
    case "set-alignment":
      if (
        action.alignment !== null &&
        !["left", "center", "right", "justify"].includes(action.alignment)
      ) {
        addIssue(
          issues,
          `${issuePath}/alignment`,
          'must be null, "left", "center", "right", or "justify"'
        );
      }
      break;
    case "toggle-list":
      if (!["unordered", "ordered"].includes(action.listType)) {
        addIssue(
          issues,
          `${issuePath}/listType`,
          'must be "unordered" or "ordered"'
        );
      }
      break;
    case "set-line-spacing":
      requireFiniteRange(
        issues,
        action.lineMultiple,
        `${issuePath}/lineMultiple`,
        0.1,
        10
      );
      break;
    case "insert-table-row":
      validateTableCoordinates(issues, action, issuePath, [
        "tableIndex",
        "rowIndex",
      ]);
      if (!["above", "below"].includes(action.direction)) {
        addIssue(
          issues,
          `${issuePath}/direction`,
          'must be "above" or "below"'
        );
      }
      break;
    case "insert-table-column":
      validateTableCoordinates(issues, action, issuePath, [
        "tableIndex",
        "cellIndex",
        "rowIndex",
      ], new Set(["rowIndex"]));
      if (!["left", "right"].includes(action.direction)) {
        addIssue(
          issues,
          `${issuePath}/direction`,
          'must be "left" or "right"'
        );
      }
      break;
    case "delete-table-row":
      validateTableCoordinates(issues, action, issuePath, [
        "tableIndex",
        "rowIndex",
      ]);
      break;
    case "delete-table-column":
      validateTableCoordinates(issues, action, issuePath, [
        "tableIndex",
        "cellIndex",
        "rowIndex",
      ], new Set(["rowIndex"]));
      break;
    case "accept-tracked-change":
    case "reject-tracked-change":
      requireIndex(issues, action.index, `${issuePath}/index`);
      break;
    case "create-comment":
      validateText(issues, action.text, `${issuePath}/text`);
      if (action.text?.trim().length === 0) {
        addIssue(issues, `${issuePath}/text`, "must contain non-whitespace text");
      }
      if (requireRecord(issues, action.options, `${issuePath}/options`)) {
        checkAllowedKeys(
          issues,
          action.options,
          `${issuePath}/options`,
          new Set(["author", "initials", "date"])
        );
        for (const key of ["author", "initials", "date"]) {
          requireString(
            issues,
            action.options[key],
            `${issuePath}/options/${key}`
          );
        }
      }
      break;
    case "set-comment-resolved":
      requireIndex(issues, action.commentId, `${issuePath}/commentId`);
      if (typeof action.resolved !== "boolean") {
        addIssue(issues, `${issuePath}/resolved`, "must be a boolean");
      }
      break;
    default:
      break;
  }
}

function validateScenario(issues, scenario, scenarioIndex) {
  const issuePath = `/scenarios/${scenarioIndex}`;
  if (!requireRecord(issues, scenario, issuePath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    scenario,
    issuePath,
    new Set(["id", "sourceCaseId", "description", "expected", "actions"])
  );
  requireId(issues, scenario.id, `${issuePath}/id`);
  requireId(issues, scenario.sourceCaseId, `${issuePath}/sourceCaseId`);
  if (scenario.description !== undefined) {
    requireString(issues, scenario.description, `${issuePath}/description`);
  }
  validateScenarioExpectation(
    issues,
    scenario.expected,
    `${issuePath}/expected`
  );
  if (!Array.isArray(scenario.actions) || scenario.actions.length === 0) {
    addIssue(issues, `${issuePath}/actions`, "must be a non-empty array");
    return;
  }
  if (scenario.actions.length > 10_000) {
    addIssue(
      issues,
      `${issuePath}/actions`,
      "must contain at most 10000 actions"
    );
  }
  scenario.actions.forEach((action, actionIndex) =>
    validateAction(issues, action, `${issuePath}/actions/${actionIndex}`)
  );
}

export function validateEditActionManifest(manifest) {
  const issues = [];
  if (!requireRecord(issues, manifest, "/")) {
    return issues;
  }
  checkAllowedKeys(
    issues,
    manifest,
    "",
    new Set([
      "$schema",
      "schemaVersion",
      "corpus",
      "sourceManifestSha256",
      "scenarios",
    ])
  );
  if (manifest.schemaVersion !== EDIT_ACTION_SCHEMA_VERSION) {
    addIssue(
      issues,
      "/schemaVersion",
      `must equal ${EDIT_ACTION_SCHEMA_VERSION}`
    );
  }
  if (manifest.$schema !== undefined) {
    requireString(issues, manifest.$schema, "/$schema");
  }
  validateCorpus(issues, manifest.corpus);
  requireSha256(
    issues,
    manifest.sourceManifestSha256,
    "/sourceManifestSha256"
  );
  if (!Array.isArray(manifest.scenarios) || manifest.scenarios.length === 0) {
    addIssue(issues, "/scenarios", "must be a non-empty array");
    return issues;
  }
  manifest.scenarios.forEach((scenario, index) =>
    validateScenario(issues, scenario, index)
  );
  const seen = new Map();
  manifest.scenarios.forEach((scenario, index) => {
    if (!isRecord(scenario) || typeof scenario.id !== "string") {
      return;
    }
    if (seen.has(scenario.id)) {
      addIssue(
        issues,
        `/scenarios/${index}/id`,
        `duplicates /scenarios/${seen.get(scenario.id)}/id`
      );
    } else {
      seen.set(scenario.id, index);
    }
  });
  return issues;
}

export function formatEditActionIssues(issues) {
  return issues.map((issue) => `${issue.path}: ${issue.message}`).join("\n");
}

export function bindEditScenarios(
  actionManifest,
  sourceManifest,
  sourceManifestSha256,
  scenarioIds = []
) {
  if (actionManifest.sourceManifestSha256 !== sourceManifestSha256) {
    throw new Error(
      "Edit action manifest was not authored for the supplied Word oracle manifest bytes"
    );
  }
  if (
    actionManifest.corpus.id !== sourceManifest.corpus.id ||
    actionManifest.corpus.revision !== sourceManifest.corpus.revision
  ) {
    throw new Error(
      "Edit action manifest corpus ID/revision does not match the Word oracle manifest"
    );
  }
  const sourceCases = new Map(
    sourceManifest.cases.map((testCase) => [testCase.id, testCase])
  );
  const filter = new Set(scenarioIds);
  const selected = actionManifest.scenarios.filter(
    (scenario) => filter.size === 0 || filter.has(scenario.id)
  );
  if (selected.length === 0) {
    throw new Error("No edit scenarios matched the supplied selectors");
  }
  for (const scenarioId of filter) {
    if (!selected.some((scenario) => scenario.id === scenarioId)) {
      throw new Error(`Unknown edit scenario selector: ${scenarioId}`);
    }
  }
  return selected.map((scenario) => {
    if (sourceCases.has(scenario.id)) {
      throw new Error(
        `Edit scenario ${scenario.id} conflicts with an existing source case ID`
      );
    }
    const sourceCase = sourceCases.get(scenario.sourceCaseId);
    if (!sourceCase) {
      throw new Error(
        `Edit scenario ${scenario.id} references unknown source case ${scenario.sourceCaseId}`
      );
    }
    return { ...scenario, source: sourceCase.source };
  });
}

export function comparableEditSummary(summary) {
  return {
    semanticModelVersion: summary.semanticModelVersion,
    semanticModelDigest: summary.semanticModelDigest,
    bodyNodeCount: summary.bodyNodeCount,
    bodyParagraphNodeCount: summary.bodyParagraphNodeCount,
    bodyTableNodeCount: summary.bodyTableNodeCount,
    paragraphCount: summary.paragraphCount,
    tableCount: summary.tableCount,
    imageCount: summary.imageCount,
    formFieldCount: summary.formFieldCount,
    trackedChangeCount: summary.trackedChangeCount,
    commentCount: summary.commentCount,
    resolvedCommentCount: summary.resolvedCommentCount,
    sectionCount: summary.sectionCount,
    headerSectionCount: summary.headerSectionCount,
    footerSectionCount: summary.footerSectionCount,
    hasTitlePageSection: summary.hasTitlePageSection,
    firstParagraph: summary.firstParagraph,
    firstTableCell: summary.firstTableCell,
  };
}

export function assertSemanticModelExpectation(
  expectation,
  beforeSummary,
  afterSummary
) {
  if (
    beforeSummary?.semanticModelVersion !== 1 ||
    afterSummary?.semanticModelVersion !== 1 ||
    !SEMANTIC_DIGEST_PATTERN.test(beforeSummary?.semanticModelDigest ?? "") ||
    !SEMANTIC_DIGEST_PATTERN.test(afterSummary?.semanticModelDigest ?? "")
  ) {
    throw new Error("semantic model summaries must use version 1 digests");
  }
  const changed =
    beforeSummary.semanticModelDigest !== afterSummary.semanticModelDigest;
  const expectsChange = new Set(["changed", "changed-from-source"]).has(
    expectation
  );
  if (
    !new Set([
      "changed",
      "unchanged",
      "changed-from-source",
      "same-as-source",
    ]).has(expectation)
  ) {
    throw new Error(`unsupported semantic model expectation: ${expectation}`);
  }
  if (changed !== expectsChange) {
    throw new Error(
      `expected semantic model to ${expectsChange ? "change" : "remain unchanged"}`
    );
  }
  return {
    expectation,
    changed,
    beforeDigest: beforeSummary.semanticModelDigest,
    afterDigest: afterSummary.semanticModelDigest,
  };
}

export function buildEditResultsManifest({
  corpus,
  sourceManifestSha256,
  actionManifestSha256,
  runner,
  scenarios,
}) {
  if (!SHA256_PATTERN.test(sourceManifestSha256 ?? "")) {
    throw new Error("sourceManifestSha256 must be a lowercase SHA-256 digest");
  }
  if (!SHA256_PATTERN.test(actionManifestSha256 ?? "")) {
    throw new Error("actionManifestSha256 must be a lowercase SHA-256 digest");
  }
  for (const scenario of scenarios) {
    if (!isPortableRelativePath(scenario.exportedDocx.path)) {
      throw new Error(
        `Scenario ${scenario.id} exported DOCX path is not portable`
      );
    }
    if (!SHA256_PATTERN.test(scenario.exportedDocx.sha256 ?? "")) {
      throw new Error(`Scenario ${scenario.id} has an invalid DOCX digest`);
    }
  }
  return {
    schemaVersion: EDIT_RESULTS_SCHEMA_VERSION,
    kind: "react-docx-edit-roundtrip-results",
    corpus: structuredClone(corpus),
    sourceManifestSha256,
    actionManifestSha256,
    runner: structuredClone(runner),
    scenarios: scenarios.map((scenario) => ({
      ...structuredClone(scenario),
      wordOracleRegistration: {
        caseId: scenario.id,
        source: {
          path: scenario.exportedDocx.path,
          sha256: scenario.exportedDocx.sha256,
        },
      },
    })),
  };
}
