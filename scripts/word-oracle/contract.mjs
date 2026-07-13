import { spawnSync } from "node:child_process";
import { createHash } from "node:crypto";
import { createReadStream, realpathSync } from "node:fs";
import fs from "node:fs/promises";
import path from "node:path";

export const WORD_ORACLE_SCHEMA_VERSION = 1;

const ID_PATTERN = /^[a-z0-9][a-z0-9._-]{0,127}$/;
const SHA256_PATTERN = /^[a-f0-9]{64}$/;
const ISO_8601_UTC_PATTERN =
  /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d{3})?Z$/;
const PAGE_BOXES = new Set(["crop", "media"]);
const EXPORT_MODES = new Set(["print", "screen"]);
const MARKUP_MODES = new Set(["final", "final-with-markup"]);
const FONT_POLICIES = new Set([
  "pinned",
  "document-embedded",
  "provider-managed",
]);

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
  if (options.pattern && !options.pattern.test(value)) {
    addIssue(
      issues,
      issuePath,
      options.patternMessage ?? "has an invalid format"
    );
    return false;
  }
  return true;
}

function optionalString(issues, value, issuePath) {
  if (value !== undefined) {
    requireString(issues, value, issuePath);
  }
}

function requireBoolean(issues, value, issuePath) {
  if (typeof value !== "boolean") {
    addIssue(issues, issuePath, "must be a boolean");
  }
}

function requirePositiveNumber(issues, value, issuePath) {
  if (typeof value !== "number" || !Number.isFinite(value) || value <= 0) {
    addIssue(issues, issuePath, "must be a finite number greater than zero");
    return false;
  }
  return true;
}

function requireId(issues, value, issuePath) {
  return requireString(issues, value, issuePath, {
    pattern: ID_PATTERN,
    patternMessage:
      "must start with a lowercase letter or digit and contain only lowercase letters, digits, dots, underscores, or hyphens",
  });
}

function requireSha256(issues, value, issuePath) {
  return requireString(issues, value, issuePath, {
    pattern: SHA256_PATTERN,
    patternMessage: "must be a lowercase 64-character SHA-256 digest",
  });
}

export function isPortableRelativePath(value) {
  if (typeof value !== "string" || value.length === 0 || value.includes("\\")) {
    return false;
  }
  if (
    path.posix.isAbsolute(value) ||
    /^[a-zA-Z]:\//.test(value) ||
    value.includes("://") ||
    value.includes("\0")
  ) {
    return false;
  }
  const parts = value.split("/");
  if (
    parts.some((part) => part.length === 0 || part === "." || part === "..")
  ) {
    return false;
  }
  return path.posix.normalize(value) === value;
}

function requirePortablePath(issues, value, issuePath, extension) {
  if (!requireString(issues, value, issuePath)) {
    return;
  }
  if (!isPortableRelativePath(value)) {
    addIssue(
      issues,
      issuePath,
      "must be a normalized POSIX path relative to the manifest directory"
    );
  }
  if (!value.toLowerCase().endsWith(extension)) {
    addIssue(issues, issuePath, `must end with ${extension}`);
  }
}

function validateCorpus(issues, corpus) {
  const issuePath = "/corpus";
  if (!requireRecord(issues, corpus, issuePath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    corpus,
    issuePath,
    new Set(["id", "revision", "description"])
  );
  requireId(issues, corpus.id, `${issuePath}/id`);
  requireString(issues, corpus.revision, `${issuePath}/revision`);
  optionalString(issues, corpus.description, `${issuePath}/description`);
}

function validateCaptureProfile(issues, profile, index) {
  const issuePath = `/captureProfiles/${index}`;
  if (!requireRecord(issues, profile, issuePath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    profile,
    issuePath,
    new Set(["id", "description", "pdfExport", "rasterization"])
  );
  requireId(issues, profile.id, `${issuePath}/id`);
  optionalString(issues, profile.description, `${issuePath}/description`);

  if (requireRecord(issues, profile.pdfExport, `${issuePath}/pdfExport`)) {
    checkAllowedKeys(
      issues,
      profile.pdfExport,
      `${issuePath}/pdfExport`,
      new Set(["mode", "markup", "updateFields", "pdfa"])
    );
    if (!EXPORT_MODES.has(profile.pdfExport.mode)) {
      addIssue(
        issues,
        `${issuePath}/pdfExport/mode`,
        'must be "print" or "screen"'
      );
    }
    if (!MARKUP_MODES.has(profile.pdfExport.markup)) {
      addIssue(
        issues,
        `${issuePath}/pdfExport/markup`,
        'must be "final" or "final-with-markup"'
      );
    }
    requireBoolean(
      issues,
      profile.pdfExport.updateFields,
      `${issuePath}/pdfExport/updateFields`
    );
    requireBoolean(
      issues,
      profile.pdfExport.pdfa,
      `${issuePath}/pdfExport/pdfa`
    );
  }

  if (
    requireRecord(issues, profile.rasterization, `${issuePath}/rasterization`)
  ) {
    checkAllowedKeys(
      issues,
      profile.rasterization,
      `${issuePath}/rasterization`,
      new Set(["dpi", "pageBox", "format", "background"])
    );
    if (
      !Number.isInteger(profile.rasterization.dpi) ||
      profile.rasterization.dpi < 72 ||
      profile.rasterization.dpi > 600
    ) {
      addIssue(
        issues,
        `${issuePath}/rasterization/dpi`,
        "must be an integer between 72 and 600"
      );
    }
    if (!PAGE_BOXES.has(profile.rasterization.pageBox)) {
      addIssue(
        issues,
        `${issuePath}/rasterization/pageBox`,
        'must be "crop" or "media"'
      );
    }
    if (profile.rasterization.format !== "png") {
      addIssue(issues, `${issuePath}/rasterization/format`, 'must be "png"');
    }
    if (profile.rasterization.background !== "#ffffff") {
      addIssue(
        issues,
        `${issuePath}/rasterization/background`,
        'must be "#ffffff"'
      );
    }
  }
}

function validatePage(issues, pageMetadata, issuePath) {
  if (!requireRecord(issues, pageMetadata, issuePath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    pageMetadata,
    issuePath,
    new Set(["widthPoints", "heightPoints", "rotation"])
  );
  requirePositiveNumber(
    issues,
    pageMetadata.widthPoints,
    `${issuePath}/widthPoints`
  );
  requirePositiveNumber(
    issues,
    pageMetadata.heightPoints,
    `${issuePath}/heightPoints`
  );
  if (![0, 90, 180, 270].includes(pageMetadata.rotation)) {
    addIssue(
      issues,
      `${issuePath}/rotation`,
      "must be one of 0, 90, 180, or 270"
    );
  }
}

function validateEnvironment(issues, environment, issuePath) {
  if (!requireRecord(issues, environment, issuePath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    environment,
    issuePath,
    new Set(["renderer", "platform", "locale", "timezone", "fontSet"])
  );

  if (requireRecord(issues, environment.renderer, `${issuePath}/renderer`)) {
    checkAllowedKeys(
      issues,
      environment.renderer,
      `${issuePath}/renderer`,
      new Set(["name", "version", "build", "channel"])
    );
    requireString(
      issues,
      environment.renderer.name,
      `${issuePath}/renderer/name`
    );
    requireString(
      issues,
      environment.renderer.version,
      `${issuePath}/renderer/version`
    );
    optionalString(
      issues,
      environment.renderer.build,
      `${issuePath}/renderer/build`
    );
    optionalString(
      issues,
      environment.renderer.channel,
      `${issuePath}/renderer/channel`
    );
  }

  if (environment.platform !== undefined) {
    if (requireRecord(issues, environment.platform, `${issuePath}/platform`)) {
      checkAllowedKeys(
        issues,
        environment.platform,
        `${issuePath}/platform`,
        new Set(["name", "version", "architecture"])
      );
      requireString(
        issues,
        environment.platform.name,
        `${issuePath}/platform/name`
      );
      optionalString(
        issues,
        environment.platform.version,
        `${issuePath}/platform/version`
      );
      optionalString(
        issues,
        environment.platform.architecture,
        `${issuePath}/platform/architecture`
      );
    }
  }

  requireString(issues, environment.locale, `${issuePath}/locale`);
  requireString(issues, environment.timezone, `${issuePath}/timezone`);

  if (requireRecord(issues, environment.fontSet, `${issuePath}/fontSet`)) {
    checkAllowedKeys(
      issues,
      environment.fontSet,
      `${issuePath}/fontSet`,
      new Set(["policy", "fingerprintSha256", "description"])
    );
    if (!FONT_POLICIES.has(environment.fontSet.policy)) {
      addIssue(
        issues,
        `${issuePath}/fontSet/policy`,
        'must be "pinned", "document-embedded", or "provider-managed"'
      );
    }
    if (environment.fontSet.fingerprintSha256 !== undefined) {
      requireSha256(
        issues,
        environment.fontSet.fingerprintSha256,
        `${issuePath}/fontSet/fingerprintSha256`
      );
    }
    if (
      environment.fontSet.policy === "pinned" &&
      environment.fontSet.fingerprintSha256 === undefined
    ) {
      addIssue(
        issues,
        `${issuePath}/fontSet/fingerprintSha256`,
        'is required when policy is "pinned"'
      );
    }
    optionalString(
      issues,
      environment.fontSet.description,
      `${issuePath}/fontSet/description`
    );
  }
}

function validateReference(issues, reference, caseIndex, referenceIndex) {
  const issuePath = `/cases/${caseIndex}/references/${referenceIndex}`;
  if (!requireRecord(issues, reference, issuePath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    reference,
    issuePath,
    new Set([
      "id",
      "provider",
      "captureProfileId",
      "capturedAt",
      "artifact",
      "environment",
      "notes",
    ])
  );
  requireId(issues, reference.id, `${issuePath}/id`);

  if (requireRecord(issues, reference.provider, `${issuePath}/provider`)) {
    checkAllowedKeys(
      issues,
      reference.provider,
      `${issuePath}/provider`,
      new Set(["id", "displayName"])
    );
    requireId(issues, reference.provider.id, `${issuePath}/provider/id`);
    requireString(
      issues,
      reference.provider.displayName,
      `${issuePath}/provider/displayName`
    );
  }

  requireId(
    issues,
    reference.captureProfileId,
    `${issuePath}/captureProfileId`
  );
  const validTimestampShape = requireString(
    issues,
    reference.capturedAt,
    `${issuePath}/capturedAt`,
    {
      pattern: ISO_8601_UTC_PATTERN,
      patternMessage: "must be an ISO 8601 UTC timestamp ending in Z",
    }
  );
  if (validTimestampShape && Number.isNaN(Date.parse(reference.capturedAt))) {
    addIssue(
      issues,
      `${issuePath}/capturedAt`,
      "must be a real calendar date and time"
    );
  }

  if (requireRecord(issues, reference.artifact, `${issuePath}/artifact`)) {
    checkAllowedKeys(
      issues,
      reference.artifact,
      `${issuePath}/artifact`,
      new Set(["path", "sha256", "mediaType", "pageCount", "pages"])
    );
    requirePortablePath(
      issues,
      reference.artifact.path,
      `${issuePath}/artifact/path`,
      ".pdf"
    );
    requireSha256(
      issues,
      reference.artifact.sha256,
      `${issuePath}/artifact/sha256`
    );
    if (reference.artifact.mediaType !== "application/pdf") {
      addIssue(
        issues,
        `${issuePath}/artifact/mediaType`,
        'must be "application/pdf"'
      );
    }
    if (
      !Number.isInteger(reference.artifact.pageCount) ||
      reference.artifact.pageCount < 1
    ) {
      addIssue(
        issues,
        `${issuePath}/artifact/pageCount`,
        "must be an integer greater than zero"
      );
    }
    if (!Array.isArray(reference.artifact.pages)) {
      addIssue(issues, `${issuePath}/artifact/pages`, "must be an array");
    } else {
      if (reference.artifact.pages.length !== reference.artifact.pageCount) {
        addIssue(
          issues,
          `${issuePath}/artifact/pages`,
          "must contain exactly pageCount entries"
        );
      }
      reference.artifact.pages.forEach((pageMetadata, pageIndex) =>
        validatePage(
          issues,
          pageMetadata,
          `${issuePath}/artifact/pages/${pageIndex}`
        )
      );
    }
  }

  validateEnvironment(
    issues,
    reference.environment,
    `${issuePath}/environment`
  );
  optionalString(issues, reference.notes, `${issuePath}/notes`);
}

function validateCase(issues, testCase, caseIndex) {
  const issuePath = `/cases/${caseIndex}`;
  if (!requireRecord(issues, testCase, issuePath)) {
    return;
  }
  checkAllowedKeys(
    issues,
    testCase,
    issuePath,
    new Set(["id", "description", "tags", "source", "references"])
  );
  requireId(issues, testCase.id, `${issuePath}/id`);
  optionalString(issues, testCase.description, `${issuePath}/description`);

  if (testCase.tags !== undefined) {
    if (!Array.isArray(testCase.tags)) {
      addIssue(issues, `${issuePath}/tags`, "must be an array");
    } else {
      testCase.tags.forEach((tag, tagIndex) =>
        requireId(issues, tag, `${issuePath}/tags/${tagIndex}`)
      );
      if (new Set(testCase.tags).size !== testCase.tags.length) {
        addIssue(issues, `${issuePath}/tags`, "must not contain duplicates");
      }
    }
  }

  if (requireRecord(issues, testCase.source, `${issuePath}/source`)) {
    checkAllowedKeys(
      issues,
      testCase.source,
      `${issuePath}/source`,
      new Set(["path", "sha256"])
    );
    requirePortablePath(
      issues,
      testCase.source.path,
      `${issuePath}/source/path`,
      ".docx"
    );
    requireSha256(issues, testCase.source.sha256, `${issuePath}/source/sha256`);
  }

  if (!Array.isArray(testCase.references) || testCase.references.length === 0) {
    addIssue(issues, `${issuePath}/references`, "must be a non-empty array");
  } else {
    testCase.references.forEach((reference, referenceIndex) =>
      validateReference(issues, reference, caseIndex, referenceIndex)
    );
  }
}

function checkUniqueIds(issues, values, issuePath) {
  const seen = new Map();
  values.forEach((value, index) => {
    if (!isRecord(value) || typeof value.id !== "string") {
      return;
    }
    if (seen.has(value.id)) {
      addIssue(
        issues,
        `${issuePath}/${index}/id`,
        `duplicates ${issuePath}/${seen.get(value.id)}/id`
      );
    } else {
      seen.set(value.id, index);
    }
  });
}

export function validateManifest(manifest) {
  const issues = [];
  if (!requireRecord(issues, manifest, "/")) {
    return issues;
  }
  checkAllowedKeys(
    issues,
    manifest,
    "",
    new Set(["$schema", "schemaVersion", "corpus", "captureProfiles", "cases"])
  );
  if (manifest.schemaVersion !== WORD_ORACLE_SCHEMA_VERSION) {
    addIssue(
      issues,
      "/schemaVersion",
      `must equal ${WORD_ORACLE_SCHEMA_VERSION}`
    );
  }
  optionalString(issues, manifest.$schema, "/$schema");
  validateCorpus(issues, manifest.corpus);

  if (
    !Array.isArray(manifest.captureProfiles) ||
    manifest.captureProfiles.length === 0
  ) {
    addIssue(issues, "/captureProfiles", "must be a non-empty array");
  } else {
    manifest.captureProfiles.forEach((profile, index) =>
      validateCaptureProfile(issues, profile, index)
    );
    checkUniqueIds(issues, manifest.captureProfiles, "/captureProfiles");
  }

  if (!Array.isArray(manifest.cases) || manifest.cases.length === 0) {
    addIssue(issues, "/cases", "must be a non-empty array");
  } else {
    manifest.cases.forEach((testCase, index) =>
      validateCase(issues, testCase, index)
    );
    checkUniqueIds(issues, manifest.cases, "/cases");
  }

  if (
    Array.isArray(manifest.captureProfiles) &&
    Array.isArray(manifest.cases)
  ) {
    const profileIds = new Set(
      manifest.captureProfiles
        .filter(isRecord)
        .map((profile) => profile.id)
        .filter((id) => typeof id === "string")
    );
    for (let caseIndex = 0; caseIndex < manifest.cases.length; caseIndex += 1) {
      const references = manifest.cases[caseIndex]?.references;
      if (!Array.isArray(references)) {
        continue;
      }
      checkUniqueIds(issues, references, `/cases/${caseIndex}/references`);
      for (
        let referenceIndex = 0;
        referenceIndex < references.length;
        referenceIndex += 1
      ) {
        const profileId = references[referenceIndex]?.captureProfileId;
        if (typeof profileId === "string" && !profileIds.has(profileId)) {
          addIssue(
            issues,
            `/cases/${caseIndex}/references/${referenceIndex}/captureProfileId`,
            `does not match a capture profile: ${profileId}`
          );
        }
      }
    }
  }

  return issues;
}

export function formatValidationIssues(issues) {
  return issues.map((issue) => `${issue.path}: ${issue.message}`).join("\n");
}

export async function readManifest(manifestPath) {
  let parsed;
  try {
    parsed = JSON.parse(await fs.readFile(manifestPath, "utf8"));
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(
      `Unable to read Word oracle manifest ${manifestPath}: ${message}`
    );
  }
  const issues = validateManifest(parsed);
  if (issues.length > 0) {
    throw new Error(
      `Invalid Word oracle manifest ${manifestPath}:\n${formatValidationIssues(
        issues
      )}`
    );
  }
  return parsed;
}

export async function sha256File(filePath) {
  const hash = createHash("sha256");
  for await (const chunk of createReadStream(filePath)) {
    hash.update(chunk);
  }
  return hash.digest("hex");
}

export function resolveManifestPath(manifestPath, portablePath) {
  if (!isPortableRelativePath(portablePath)) {
    throw new Error(`Unsafe manifest-relative path: ${portablePath}`);
  }
  const root = path.resolve(path.dirname(manifestPath));
  const resolved = path.resolve(root, ...portablePath.split("/"));
  if (resolved !== root && !resolved.startsWith(`${root}${path.sep}`)) {
    throw new Error(`Path escapes the manifest directory: ${portablePath}`);
  }
  let realRoot;
  let realResolved;
  try {
    realRoot = realpathSync(root);
    realResolved = realpathSync(resolved);
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    throw new Error(
      `Cannot resolve manifest-relative input ${portablePath}: ${message}`
    );
  }
  const realRelative = path.relative(realRoot, realResolved);
  if (
    realRelative === ".." ||
    realRelative.startsWith(`..${path.sep}`) ||
    path.isAbsolute(realRelative)
  ) {
    throw new Error(
      `Manifest-relative input resolves outside the manifest directory: ${portablePath}`
    );
  }
  return realResolved;
}

export function parsePdfInfoOutput(output) {
  const pageCountMatch = output.match(/^Pages:\s+(\d+)\s*$/m);
  if (!pageCountMatch) {
    throw new Error("pdfinfo output does not include a page count");
  }
  const pageCount = Number(pageCountMatch[1]);
  const pagesByNumber = new Map();
  const sizePattern =
    /^(?:Page(?:\s+(\d+))?\s+)?size:\s+([0-9.]+)\s+x\s+([0-9.]+)\s+pts(?:\s|$)/gim;
  for (const match of output.matchAll(sizePattern)) {
    const pageNumber = match[1] ? Number(match[1]) : 1;
    pagesByNumber.set(pageNumber, {
      widthPoints: Number(match[2]),
      heightPoints: Number(match[3]),
      rotation: 0,
    });
  }
  const rotationPattern = /^(?:Page(?:\s+(\d+))?\s+)?rot:\s+(-?\d+)\s*$/gim;
  for (const match of output.matchAll(rotationPattern)) {
    const pageNumber = match[1] ? Number(match[1]) : 1;
    const page = pagesByNumber.get(pageNumber);
    if (page) {
      page.rotation = ((Number(match[2]) % 360) + 360) % 360;
    }
  }

  const pages = [];
  for (let pageNumber = 1; pageNumber <= pageCount; pageNumber += 1) {
    const page = pagesByNumber.get(pageNumber);
    if (!page) {
      throw new Error(
        `pdfinfo output does not include geometry for page ${pageNumber}`
      );
    }
    pages.push(page);
  }
  return { pageCount, pages };
}

export function runCommand(command, args, options = {}) {
  const result = spawnSync(command, args, {
    encoding: "utf8",
    maxBuffer: 16 * 1024 * 1024,
    ...options,
  });
  if (result.error || result.status !== 0) {
    const detail =
      result.error?.message || result.stderr?.trim() || result.stdout?.trim();
    throw new Error(
      `${command} failed${
        detail ? `: ${detail}` : ` with status ${result.status}`
      }`
    );
  }
  return result;
}

export function commandVersion(command) {
  const result = runCommand(command, ["-v"]);
  return `${result.stderr || result.stdout}`.split(/\r?\n/, 1)[0].trim();
}

export function inspectPdf(
  pdfPath,
  pdfinfoBin = process.env.PDFINFO_BIN || "pdfinfo"
) {
  const summary = runCommand(pdfinfoBin, ["-box", pdfPath]);
  const countMatch = summary.stdout.match(/^Pages:\s+(\d+)\s*$/m);
  if (!countMatch) {
    throw new Error(`Could not determine the page count for ${pdfPath}`);
  }
  const pageCount = Number(countMatch[1]);
  const detailed = runCommand(pdfinfoBin, [
    "-f",
    "1",
    "-l",
    String(pageCount),
    "-box",
    pdfPath,
  ]);
  return parsePdfInfoOutput(detailed.stdout);
}

function pageGeometryMatches(expected, actual, tolerance = 0.01) {
  return (
    Math.abs(expected.widthPoints - actual.widthPoints) <= tolerance &&
    Math.abs(expected.heightPoints - actual.heightPoints) <= tolerance &&
    expected.rotation === actual.rotation
  );
}

export async function verifyManifestArtifacts(
  manifest,
  manifestPath,
  options = {}
) {
  const issues = [];
  const inspectPdfMetadata = options.inspectPdfMetadata !== false;
  const pdfinfoBin = options.pdfinfoBin ?? process.env.PDFINFO_BIN ?? "pdfinfo";

  for (let caseIndex = 0; caseIndex < manifest.cases.length; caseIndex += 1) {
    const testCase = manifest.cases[caseIndex];
    try {
      const sourcePath = resolveManifestPath(manifestPath, testCase.source.path);
      const actualHash = await sha256File(sourcePath);
      if (actualHash !== testCase.source.sha256) {
        addIssue(
          issues,
          `/cases/${caseIndex}/source/sha256`,
          `does not match ${testCase.source.path}: expected ${testCase.source.sha256}, got ${actualHash}`
        );
      }
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error);
      addIssue(
        issues,
        `/cases/${caseIndex}/source/path`,
        `cannot be read: ${message}`
      );
    }

    for (
      let referenceIndex = 0;
      referenceIndex < testCase.references.length;
      referenceIndex += 1
    ) {
      const reference = testCase.references[referenceIndex];
      const issuePath = `/cases/${caseIndex}/references/${referenceIndex}/artifact`;
      let artifactReadable = true;
      let artifactPath;
      try {
        artifactPath = resolveManifestPath(
          manifestPath,
          reference.artifact.path
        );
        const actualHash = await sha256File(artifactPath);
        if (actualHash !== reference.artifact.sha256) {
          addIssue(
            issues,
            `${issuePath}/sha256`,
            `does not match ${reference.artifact.path}: expected ${reference.artifact.sha256}, got ${actualHash}`
          );
        }
      } catch (error) {
        artifactReadable = false;
        const message = error instanceof Error ? error.message : String(error);
        addIssue(issues, `${issuePath}/path`, `cannot be read: ${message}`);
      }

      if (!inspectPdfMetadata || !artifactReadable || !artifactPath) {
        continue;
      }
      try {
        const actual = inspectPdf(artifactPath, pdfinfoBin);
        if (actual.pageCount !== reference.artifact.pageCount) {
          addIssue(
            issues,
            `${issuePath}/pageCount`,
            `does not match the PDF: expected ${reference.artifact.pageCount}, got ${actual.pageCount}`
          );
          continue;
        }
        for (
          let pageIndex = 0;
          pageIndex < actual.pages.length;
          pageIndex += 1
        ) {
          if (
            !pageGeometryMatches(
              reference.artifact.pages[pageIndex],
              actual.pages[pageIndex]
            )
          ) {
            addIssue(
              issues,
              `${issuePath}/pages/${pageIndex}`,
              `does not match the PDF: expected ${JSON.stringify(
                reference.artifact.pages[pageIndex]
              )}, got ${JSON.stringify(actual.pages[pageIndex])}`
            );
          }
        }
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        addIssue(issues, issuePath, `PDF inspection failed: ${message}`);
      }
    }
  }
  return issues;
}
