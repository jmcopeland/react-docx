import type {
  DocModel,
  DocumentCommentDefinition,
  ParagraphNode,
  TextRunNode,
  TextStyle,
} from "@extend-ai/react-docx-doc-model";
import { cloneDocModel } from "@extend-ai/react-docx-doc-model";

import {
  decodeXmlText,
  directChildElements,
  elementsNamed,
  enclosingElements,
  escapeXmlText,
  scanBalancedXml,
  setXmlAttribute,
  xmlAttribute,
  type XmlElementRange,
  type XmlScanResult,
} from "./xml";

export type AnnotationMutationFailureReason =
  | "stale"
  | "unsupported"
  | "unsafe-xml";

export type AnnotationMutationResult =
  | { ok: true; model: DocModel }
  | { ok: false; reason: AnnotationMutationFailureReason };

export type CommentCreationResult =
  | { ok: true; model: DocModel; commentId: number }
  | { ok: false; reason: AnnotationMutationFailureReason };

export type ParagraphRevisionKind =
  | "insertion"
  | "deletion"
  | "move-from"
  | "move-to"
  | "format-change"
  | "paragraph-format-change";

export interface ParagraphRevisionTarget {
  /** Top-level body node containing the revision. Table locations are unsupported. */
  nodeIndex: number;
  /** OOXML `w:id` from the `w:ins` or `w:del` wrapper. */
  revisionId: string | number;
  kind: ParagraphRevisionKind;
}

export interface CreateParagraphCommentInput {
  /** Top-level body paragraph index. */
  nodeIndex: number;
  /** UTF-16 offsets into the paragraph's plain text, matching DOM selection offsets. */
  startOffset: number;
  endOffset: number;
  text: string;
  author?: string;
  initials?: string;
  date?: string;
}

interface VisibleTextToken {
  range: XmlElementRange;
  text: string;
  enclosingRun: XmlElementRange;
}

interface SafeParagraphContext {
  paragraph: ParagraphNode;
  sourceXml: string;
  scan: XmlScanResult;
  paragraphRange: XmlElementRange;
  visibleTokens: VisibleTextToken[];
}

const FORBIDDEN_PARAGRAPH_ELEMENTS = new Set([
  "w:fldChar",
  "w:instrText",
  "w:fldSimple",
  "w:sdt",
  "w:bookmarkStart",
  "w:bookmarkEnd",
  "w:drawing",
  "w:pict",
  "w:object",
  "w:hyperlink",
  "w:moveFrom",
  "w:moveTo",
  "w:rPrChange",
  "w:pPrChange",
  "w:sectPrChange",
  "w:tblPrChange",
  "w:trPrChange",
  "w:tcPrChange",
  "w:numberingChange",
  "w:footnoteReference",
  "w:endnoteReference",
  "w:tab",
  "w:br",
  "w:cr",
  "mc:AlternateContent",
  "a:t",
]);

function textContent(xml: string, range: XmlElementRange): string | undefined {
  if (range.selfClosing) {
    return "";
  }
  const content = xml.slice(range.openEnd, range.closeStart);
  return content.includes("<") ? undefined : decodeXmlText(content);
}

function parentElement(
  scan: XmlScanResult,
  element: XmlElementRange
): XmlElementRange | undefined {
  return element.parentStart === undefined
    ? undefined
    : scan.elements.find(
        (candidate) => candidate.start === element.parentStart
      );
}

function isInsideNamedElement(
  scan: XmlScanResult,
  element: XmlElementRange,
  tagName: string
): boolean {
  return enclosingElements(scan, element).some(
    (candidate) => candidate.tagName === tagName
  );
}

function visibleTextTokens(
  sourceXml: string,
  scan: XmlScanResult
): VisibleTextToken[] | undefined {
  const tokens: VisibleTextToken[] = [];
  for (const range of elementsNamed(scan, "w:t")) {
    if (isInsideNamedElement(scan, range, "w:del")) {
      continue;
    }
    const text = textContent(sourceXml, range);
    if (text === undefined) {
      return undefined;
    }
    if (text.length === 0) {
      continue;
    }
    const enclosingRun = enclosingElements(scan, range).find(
      (candidate) => candidate.tagName === "w:r"
    );
    if (!enclosingRun) {
      return undefined;
    }
    tokens.push({ range, text, enclosingRun });
  }
  return tokens.sort((left, right) => left.range.start - right.range.start);
}

function modelMatchesVisibleText(
  paragraph: ParagraphNode,
  tokens: VisibleTextToken[]
): boolean {
  if (
    paragraph.children.some(
      (child) =>
        child.type !== "text" ||
        child.link !== undefined ||
        child.noteReference !== undefined
    )
  ) {
    return false;
  }
  const modeled = paragraph.children as TextRunNode[];
  if (tokens.length === 0) {
    return modeled.length === 1 && modeled[0]?.text === "";
  }
  return (
    modeled.length === tokens.length &&
    modeled.every((run, index) => run.text === tokens[index]?.text)
  );
}

function cloneTextStyle(style?: TextStyle): TextStyle | undefined {
  return style
    ? {
        ...style,
        runBorder: style.runBorder ? { ...style.runBorder } : undefined,
      }
    : undefined;
}

function canonicalJsonValue(value: unknown): unknown {
  if (Array.isArray(value)) {
    return value.map(canonicalJsonValue);
  }
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>)
        .filter(([, entry]) => entry !== undefined)
        .sort(([left], [right]) => left.localeCompare(right))
        .map(([key, entry]) => [key, canonicalJsonValue(entry)])
    );
  }
  return value;
}

function metadataValuesEqual(left: unknown, right: unknown): boolean {
  return (
    JSON.stringify(canonicalJsonValue(left)) ===
    JSON.stringify(canonicalJsonValue(right))
  );
}

function sourceRunProvenanceForChildren(
  children: ParagraphNode["children"]
): ParagraphNode["sourceRunProvenance"] | undefined {
  if (children.some((child) => child.type !== "text")) {
    return undefined;
  }
  return {
    runs: (children as TextRunNode[]).map((run) => ({
      style: cloneTextStyle(run.style),
      link: run.link,
      noteReference: run.noteReference ? { ...run.noteReference } : undefined,
    })),
  };
}

function modelMatchesSourceRunProvenance(paragraph: ParagraphNode): boolean {
  const provenance = paragraph.sourceRunProvenance;
  if (!provenance || paragraph.children.length !== provenance.runs.length) {
    return false;
  }
  return paragraph.children.every((child, index) => {
    if (child.type !== "text") {
      return false;
    }
    const source = provenance.runs[index];
    return (
      Boolean(source) &&
      metadataValuesEqual(child.style, source?.style) &&
      child.link === source?.link &&
      metadataValuesEqual(child.noteReference, source?.noteReference)
    );
  });
}

function safeParagraphContext(
  model: DocModel,
  nodeIndex: number
):
  | { ok: true; context: SafeParagraphContext }
  | { ok: false; reason: AnnotationMutationFailureReason } {
  const paragraph = model.nodes[nodeIndex];
  if (!paragraph) {
    return { ok: false, reason: "stale" };
  }
  if (paragraph.type !== "paragraph") {
    return { ok: false, reason: "unsupported" };
  }
  if (!paragraph.sourceXml || paragraph.sourceTextPatch) {
    return { ok: false, reason: "stale" };
  }

  const sourceXml = paragraph.sourceXml;
  const scan = scanBalancedXml(sourceXml);
  if (scan.malformed) {
    return { ok: false, reason: "unsafe-xml" };
  }
  const paragraphs = elementsNamed(scan, "w:p").filter(
    (candidate) => candidate.parentStart === undefined
  );
  if (paragraphs.length !== 1) {
    return { ok: false, reason: "unsafe-xml" };
  }
  if (
    scan.elements.some((element) =>
      FORBIDDEN_PARAGRAPH_ELEMENTS.has(element.tagName)
    )
  ) {
    return { ok: false, reason: "unsupported" };
  }

  const paragraphRange = paragraphs[0]!;
  for (const revision of [
    ...elementsNamed(scan, "w:ins"),
    ...elementsNamed(scan, "w:del"),
  ]) {
    if (revision.parentStart !== paragraphRange.start) {
      return { ok: false, reason: "unsupported" };
    }
  }

  const tokens = visibleTextTokens(sourceXml, scan);
  if (!tokens) {
    return { ok: false, reason: "unsafe-xml" };
  }
  if (!modelMatchesVisibleText(paragraph, tokens)) {
    return { ok: false, reason: "stale" };
  }
  if (!modelMatchesSourceRunProvenance(paragraph)) {
    return { ok: false, reason: "stale" };
  }

  return {
    ok: true,
    context: {
      paragraph,
      sourceXml,
      scan,
      paragraphRange,
      visibleTokens: tokens,
    },
  };
}

function hasNonWhitespaceOutsideChildren(
  xml: string,
  parent: XmlElementRange,
  children: XmlElementRange[]
): boolean {
  let cursor = parent.openEnd;
  for (const child of [...children].sort(
    (left, right) => left.start - right.start
  )) {
    if (xml.slice(cursor, child.start).trim().length > 0) {
      return true;
    }
    cursor = child.end;
  }
  return xml.slice(cursor, parent.closeStart).trim().length > 0;
}

function revisionRangeIsSafe(
  context: SafeParagraphContext,
  revision: XmlElementRange,
  kind: "insertion" | "deletion"
): boolean {
  const runs = directChildElements(context.scan, revision);
  if (
    runs.length === 0 ||
    runs.some((run) => run.tagName !== "w:r") ||
    hasNonWhitespaceOutsideChildren(context.sourceXml, revision, runs)
  ) {
    return false;
  }

  const expectedTextTag = kind === "insertion" ? "w:t" : "w:delText";
  for (const run of runs) {
    const children = directChildElements(context.scan, run);
    if (
      children.some((child) => {
        const name = child.tagName;
        return name !== "w:rPr" && name !== expectedTextTag;
      }) ||
      hasNonWhitespaceOutsideChildren(context.sourceXml, run, children)
    ) {
      return false;
    }
    const textElements = children.filter(
      (child) => child.tagName === expectedTextTag
    );
    if (
      textElements.length === 0 ||
      textElements.some(
        (textElement) =>
          textContent(context.sourceXml, textElement) === undefined
      )
    ) {
      return false;
    }
  }
  return true;
}

function matchingRevisionRanges(
  context: SafeParagraphContext,
  target: ParagraphRevisionTarget,
  kind: "insertion" | "deletion"
): XmlElementRange[] {
  const tagName = kind === "insertion" ? "w:ins" : "w:del";
  const revisionId = String(target.revisionId).trim();
  return elementsNamed(context.scan, tagName).filter(
    (range) =>
      range.parentStart === context.paragraphRange.start &&
      xmlAttribute(range.openTag, "w:id")?.trim() === revisionId
  );
}

function replaceRanges(
  source: string,
  replacements: Array<{ start: number; end: number; value: string }>
): string {
  let updated = source;
  for (const replacement of [...replacements].sort(
    (left, right) => right.start - left.start
  )) {
    updated =
      updated.slice(0, replacement.start) +
      replacement.value +
      updated.slice(replacement.end);
  }
  return updated;
}

function renameBalancedElements(
  xml: string,
  fromTagName: string,
  toTagName: string
): string | undefined {
  const scan = scanBalancedXml(xml);
  if (scan.malformed) {
    return undefined;
  }
  const replacements: Array<{ start: number; end: number; value: string }> = [];
  for (const range of elementsNamed(scan, fromTagName)) {
    replacements.push({
      start: range.start + 1,
      end: range.start + 1 + range.tagName.length,
      value: toTagName,
    });
    if (!range.selfClosing) {
      replacements.push({
        start: range.closeStart + 2,
        end: range.closeStart + 2 + range.tagName.length,
        value: toTagName,
      });
    }
  }
  return replaceRanges(xml, replacements);
}

function booleanElementValue(
  scan: XmlScanResult,
  tagName: string
): boolean | undefined {
  const element = elementsNamed(scan, tagName)[0];
  if (!element) {
    return undefined;
  }
  const raw = xmlAttribute(element.openTag, "w:val")?.trim().toLowerCase();
  return raw === undefined || !["0", "false", "off", "none"].includes(raw);
}

function normalizedColor(raw?: string): string | undefined {
  const value = raw?.trim();
  if (!value || value.toLowerCase() === "auto") {
    return undefined;
  }
  return /^[0-9a-f]{6}$/i.test(value) ? `#${value.toLowerCase()}` : value;
}

function styleFromRunXml(runXml: string): TextStyle | undefined {
  const scan = scanBalancedXml(runXml);
  if (scan.malformed) {
    return undefined;
  }
  const run = elementsNamed(scan, "w:r")[0];
  const properties = run
    ? directChildElements(scan, run).find(
        (candidate) => candidate.tagName === "w:rPr"
      )
    : undefined;
  if (!properties) {
    return undefined;
  }
  const propertyScan = scanBalancedXml(
    runXml.slice(properties.start, properties.end)
  );
  if (propertyScan.malformed) {
    return undefined;
  }

  const fonts = elementsNamed(propertyScan, "w:rFonts")[0]?.openTag;
  const language = elementsNamed(propertyScan, "w:lang")[0]?.openTag;
  const size = elementsNamed(propertyScan, "w:sz")[0]?.openTag;
  const spacing = elementsNamed(propertyScan, "w:spacing")[0]?.openTag;
  const verticalAlign = elementsNamed(propertyScan, "w:vertAlign")[0]?.openTag;
  const color = elementsNamed(propertyScan, "w:color")[0]?.openTag;
  const highlight = elementsNamed(propertyScan, "w:highlight")[0]?.openTag;
  const shading = elementsNamed(propertyScan, "w:shd")[0]?.openTag;
  const border = elementsNamed(propertyScan, "w:bdr")[0]?.openTag;
  const fontFamilyAscii = fonts ? xmlAttribute(fonts, "w:ascii") : undefined;
  const fontFamilyHAnsi = fonts ? xmlAttribute(fonts, "w:hAnsi") : undefined;
  const fontFamilyEastAsia = fonts
    ? xmlAttribute(fonts, "w:eastAsia")
    : undefined;
  const fontFamilyCs = fonts ? xmlAttribute(fonts, "w:cs") : undefined;
  const fontFamily =
    fontFamilyAscii ?? fontFamilyHAnsi ?? fontFamilyEastAsia ?? fontFamilyCs;
  const sizeHalfPoints = Number(size ? xmlAttribute(size, "w:val") : undefined);
  const spacingTwips = Number(
    spacing ? xmlAttribute(spacing, "w:val") : undefined
  );
  const borderSize = Number(border ? xmlAttribute(border, "w:sz") : undefined);
  const borderSpace = Number(
    border ? xmlAttribute(border, "w:space") : undefined
  );
  const borderType = border ? xmlAttribute(border, "w:val") : undefined;

  const style: TextStyle = {
    bold: booleanElementValue(propertyScan, "w:b"),
    italic: booleanElementValue(propertyScan, "w:i"),
    underline: booleanElementValue(propertyScan, "w:u"),
    strike: booleanElementValue(propertyScan, "w:strike"),
    color: normalizedColor(color ? xmlAttribute(color, "w:val") : undefined),
    highlight: highlight ? xmlAttribute(highlight, "w:val") : undefined,
    backgroundColor: normalizedColor(
      shading ? xmlAttribute(shading, "w:fill") : undefined
    ),
    fontSizePt:
      Number.isFinite(sizeHalfPoints) && sizeHalfPoints > 0
        ? sizeHalfPoints / 2
        : undefined,
    fontFamily,
    sourceFontFamily: fontFamily,
    fontFamilyAscii,
    fontFamilyHAnsi,
    fontFamilyEastAsia,
    fontFamilyCs,
    fontThemeAscii: fonts ? xmlAttribute(fonts, "w:asciiTheme") : undefined,
    fontThemeHAnsi: fonts ? xmlAttribute(fonts, "w:hAnsiTheme") : undefined,
    fontThemeEastAsia: fonts
      ? xmlAttribute(fonts, "w:eastAsiaTheme")
      : undefined,
    fontThemeCs: fonts ? xmlAttribute(fonts, "w:csTheme") : undefined,
    fontHint: fonts ? xmlAttribute(fonts, "w:hint") : undefined,
    language: language ? xmlAttribute(language, "w:val") : undefined,
    languageEastAsia: language
      ? xmlAttribute(language, "w:eastAsia")
      : undefined,
    languageBidi: language ? xmlAttribute(language, "w:bidi") : undefined,
    rightToLeft: booleanElementValue(propertyScan, "w:rtl"),
    complexScript: booleanElementValue(propertyScan, "w:cs"),
    characterSpacingTwips: Number.isFinite(spacingTwips)
      ? spacingTwips
      : undefined,
    verticalAlign:
      verticalAlign &&
      ["superscript", "subscript"].includes(
        xmlAttribute(verticalAlign, "w:val") ?? ""
      )
        ? (xmlAttribute(verticalAlign, "w:val") as "superscript" | "subscript")
        : undefined,
    runBorder:
      borderType && borderType !== "none" && borderType !== "nil"
        ? {
            type: borderType,
            color: normalizedColor(xmlAttribute(border!, "w:color")),
            sizeEighthPt: Number.isFinite(borderSize) ? borderSize : undefined,
            spacePt: Number.isFinite(borderSpace) ? borderSpace : undefined,
            frame: border ? xmlAttribute(border, "w:frame") === "1" : undefined,
            shadow: border
              ? xmlAttribute(border, "w:shadow") === "1"
              : undefined,
          }
        : undefined,
  };
  return Object.values(style).some((value) => value !== undefined)
    ? style
    : undefined;
}

function deletedRunsForTargets(
  context: SafeParagraphContext,
  targets: XmlElementRange[]
): Array<{ position: number; run: TextRunNode }> | undefined {
  const deleted: Array<{ position: number; run: TextRunNode }> = [];
  for (const target of targets) {
    const targetRuns = directChildElements(context.scan, target).filter(
      (element) => element.tagName === "w:r"
    );
    for (const run of targetRuns) {
      const style = styleFromRunXml(
        context.sourceXml.slice(run.start, run.end)
      );
      const deletedTexts = directChildElements(context.scan, run).filter(
        (element) => element.tagName === "w:delText"
      );
      for (const deletedText of deletedTexts) {
        const text = textContent(context.sourceXml, deletedText);
        if (text === undefined) {
          return undefined;
        }
        if (text.length > 0) {
          deleted.push({
            position: deletedText.start,
            run: {
              type: "text",
              text,
              style: style ? { ...style } : undefined,
            },
          });
        }
      }
    }
  }
  return deleted;
}

function mutateParagraphRevision(
  model: DocModel,
  target: ParagraphRevisionTarget,
  action: "accept" | "reject"
): AnnotationMutationResult {
  const kind =
    target.kind === "insertion"
      ? "insertion"
      : target.kind === "deletion"
      ? "deletion"
      : undefined;
  if (!kind) {
    return { ok: false, reason: "unsupported" };
  }
  const prepared = safeParagraphContext(model, target.nodeIndex);
  if (!prepared.ok) {
    return prepared;
  }
  const context = prepared.context;
  const targets = matchingRevisionRanges(context, target, kind);
  if (targets.length === 0) {
    return { ok: false, reason: "stale" };
  }
  if (targets.length !== 1) {
    return { ok: false, reason: "unsafe-xml" };
  }
  if (targets.some((range) => !revisionRangeIsSafe(context, range, kind))) {
    return { ok: false, reason: "unsupported" };
  }

  const keepRevisionContent =
    (kind === "insertion" && action === "accept") ||
    (kind === "deletion" && action === "reject");
  const replacements: Array<{ start: number; end: number; value: string }> = [];
  for (const range of targets) {
    let value = keepRevisionContent
      ? context.sourceXml.slice(range.openEnd, range.closeStart)
      : "";
    if (kind === "deletion" && action === "reject") {
      const renamed = renameBalancedElements(value, "w:delText", "w:t");
      if (renamed === undefined) {
        return { ok: false, reason: "unsafe-xml" };
      }
      value = renamed;
    }
    replacements.push({ start: range.start, end: range.end, value });
  }

  const next = cloneDocModel(model);
  const paragraph = next.nodes[target.nodeIndex];
  if (!paragraph || paragraph.type !== "paragraph") {
    return { ok: false, reason: "stale" };
  }
  paragraph.sourceXml = replaceRanges(context.sourceXml, replacements);
  paragraph.sourceTextPatch = undefined;

  if (kind === "insertion" && action === "reject") {
    const targetRanges = targets;
    paragraph.children = context.visibleTokens
      .map((token, index) => ({ token, run: paragraph.children[index] }))
      .filter(
        ({ token }) =>
          !targetRanges.some(
            (range) =>
              token.range.start >= range.start && token.range.end <= range.end
          )
      )
      .map(({ run }) => run!) as TextRunNode[];
  } else if (kind === "deletion" && action === "reject") {
    const deletedRuns = deletedRunsForTargets(context, targets);
    if (!deletedRuns) {
      return { ok: false, reason: "unsafe-xml" };
    }
    const events: Array<
      | { position: number; order: 0; run: TextRunNode }
      | { position: number; order: 1; run: TextRunNode }
    > = context.visibleTokens.map((token, index) => ({
      position: token.range.start,
      order: 1,
      run: paragraph.children[index] as TextRunNode,
    }));
    events.push(
      ...deletedRuns.map(({ position, run }) => ({
        position,
        order: 0 as const,
        run,
      }))
    );
    events.sort(
      (left, right) =>
        left.position - right.position || left.order - right.order
    );
    paragraph.children = events.map((event) => event.run);
  }

  if (paragraph.children.length === 0) {
    paragraph.children = [{ type: "text", text: "" }];
  }
  paragraph.sourceRunProvenance = sourceRunProvenanceForChildren(
    paragraph.children
  );
  return { ok: true, model: next };
}

export function acceptParagraphRevision(
  model: DocModel,
  target: ParagraphRevisionTarget
): AnnotationMutationResult {
  return mutateParagraphRevision(model, target, "accept");
}

export function rejectParagraphRevision(
  model: DocModel,
  target: ParagraphRevisionTarget
): AnnotationMutationResult {
  return mutateParagraphRevision(model, target, "reject");
}

export function setCommentResolved(
  model: DocModel,
  commentId: number,
  resolved: boolean
): AnnotationMutationResult {
  const index = model.metadata.comments?.findIndex(
    (comment) => comment.id === commentId
  );
  if (index === undefined || index < 0) {
    return { ok: false, reason: "stale" };
  }
  if (model.metadata.comments?.[index]?.resolved === resolved) {
    return { ok: true, model };
  }
  const next = cloneDocModel(model);
  const comment = next.metadata.comments?.[index];
  if (!comment) {
    return { ok: false, reason: "stale" };
  }
  comment.resolved = resolved;
  comment.resolutionDirty = comment.sourceResolved !== resolved;
  return { ok: true, model: next };
}

function nextCommentId(comments: DocumentCommentDefinition[]): number {
  return (
    comments.reduce((largest, comment) => Math.max(largest, comment.id), -1) + 1
  );
}

function allocateExtendedParagraphId(
  comments: DocumentCommentDefinition[],
  commentId: number
): string {
  const used = new Set(
    comments
      .map((comment) => comment.extendedParagraphId?.toUpperCase())
      .filter((value): value is string => Boolean(value))
  );
  let candidate = (0xc0000000 + Math.max(0, commentId)) >>> 0;
  while (used.has(candidate.toString(16).padStart(8, "0").toUpperCase())) {
    candidate = (candidate + 1) >>> 0;
  }
  return candidate.toString(16).padStart(8, "0").toUpperCase();
}

function isValidXml10Text(value: string): boolean {
  for (let index = 0; index < value.length; index += 1) {
    const codeUnit = value.charCodeAt(index);
    if (codeUnit >= 0xd800 && codeUnit <= 0xdbff) {
      const next = value.charCodeAt(index + 1);
      if (!(next >= 0xdc00 && next <= 0xdfff)) {
        return false;
      }
      index += 1;
      continue;
    }
    if (codeUnit >= 0xdc00 && codeUnit <= 0xdfff) {
      return false;
    }
    if (
      codeUnit !== 0x09 &&
      codeUnit !== 0x0a &&
      codeUnit !== 0x0d &&
      (codeUnit < 0x20 || codeUnit === 0xfffe || codeUnit === 0xffff)
    ) {
      return false;
    }
  }
  return true;
}

function isLeapYear(year: bigint): boolean {
  return year % 4n === 0n && (year % 100n !== 0n || year % 400n === 0n);
}

function isValidXmlDateTime(value: string): boolean {
  const match =
    /^(-?)(\d{4,})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})(?:\.(\d+))?(?:(Z)|([+-])(\d{2}):(\d{2}))?$/.exec(
      value
    );
  if (!match || /^0+$/.test(match[2] ?? "")) {
    return false;
  }

  const year = BigInt(`${match[1] ?? ""}${match[2]}`);
  const month = Number(match[3]);
  const day = Number(match[4]);
  const hour = Number(match[5]);
  const minute = Number(match[6]);
  const second = Number(match[7]);
  const fraction = match[8];
  const zoneHour = match[11] === undefined ? undefined : Number(match[11]);
  const zoneMinute = match[12] === undefined ? undefined : Number(match[12]);
  const daysInMonth = [
    31,
    isLeapYear(year) ? 29 : 28,
    31,
    30,
    31,
    30,
    31,
    31,
    30,
    31,
    30,
    31,
  ];

  if (
    month < 1 ||
    month > 12 ||
    day < 1 ||
    day > (daysInMonth[month - 1] ?? 0) ||
    minute > 59 ||
    second > 59
  ) {
    return false;
  }
  if (
    hour > 24 ||
    (hour === 24 &&
      (minute !== 0 || second !== 0 || Boolean(fraction?.replace(/0/g, ""))))
  ) {
    return false;
  }
  if (
    zoneHour !== undefined &&
    (zoneHour > 14 ||
      (zoneHour === 14 && zoneMinute !== 0) ||
      zoneMinute === undefined ||
      zoneMinute > 59)
  ) {
    return false;
  }
  return true;
}

function isUtf16Boundary(text: string, offset: number): boolean {
  if (offset <= 0 || offset >= text.length) {
    return true;
  }
  const previous = text.charCodeAt(offset - 1);
  const next = text.charCodeAt(offset);
  return !(
    previous >= 0xd800 &&
    previous <= 0xdbff &&
    next >= 0xdc00 &&
    next <= 0xdfff
  );
}

function textOpeningTagForSegment(openTag: string, text: string): string {
  if (!/^\s|\s$/.test(text)) {
    return openTag;
  }
  return setXmlAttribute(openTag, "xml:space", "preserve");
}

function runXmlWithText(
  sourceXml: string,
  run: XmlElementRange,
  textRange: XmlElementRange,
  text: string
): string {
  const runXml = sourceXml.slice(run.start, run.end);
  const relativeTextStart = textRange.start - run.start;
  const relativeTextEnd = textRange.end - run.start;
  const openTag = textOpeningTagForSegment(textRange.openTag, text);
  const textXml = textRange.selfClosing
    ? `<w:t${
        /^\s|\s$/.test(text) ? ' xml:space="preserve"' : ""
      }>${escapeXmlText(text)}</w:t>`
    : `${openTag}${escapeXmlText(text)}${sourceXml.slice(
        textRange.closeStart,
        textRange.end
      )}`;
  return (
    runXml.slice(0, relativeTextStart) + textXml + runXml.slice(relativeTextEnd)
  );
}

/**
 * Creates a comment over a plain-text range in one top-level paragraph. The
 * operation splits only the boundary runs and preserves their original rPr.
 */
export function createParagraphComment(
  model: DocModel,
  input: CreateParagraphCommentInput
): CommentCreationResult {
  const prepared = safeParagraphContext(model, input.nodeIndex);
  if (!prepared.ok) {
    return prepared;
  }
  const context = prepared.context;
  const stringFields = [
    input.text,
    input.author,
    input.initials,
    input.date,
  ].filter((value): value is string => value !== undefined);
  if (
    !input.text.trim() ||
    stringFields.some((value) => !isValidXml10Text(value)) ||
    (input.date !== undefined && !isValidXmlDateTime(input.date)) ||
    !Number.isInteger(input.startOffset) ||
    !Number.isInteger(input.endOffset) ||
    input.startOffset < 0 ||
    input.endOffset <= input.startOffset
  ) {
    return { ok: false, reason: "unsupported" };
  }

  // Keep creation intentionally narrower than resolution/revision mutation:
  // only direct, one-text-node runs with no existing annotation/revision markup.
  const forbiddenForCreation = new Set([
    "w:ins",
    "w:del",
    "w:commentRangeStart",
    "w:commentRangeEnd",
    "w:commentReference",
  ]);
  if (
    context.scan.elements.some((element) =>
      forbiddenForCreation.has(element.tagName)
    )
  ) {
    return { ok: false, reason: "unsupported" };
  }
  const directChildren = directChildElements(
    context.scan,
    context.paragraphRange
  );
  if (
    directChildren.some((child) => {
      const name = child.tagName;
      return name !== "w:pPr" && name !== "w:r";
    })
  ) {
    return { ok: false, reason: "unsupported" };
  }
  const runs = directChildren.filter((child) => child.tagName === "w:r");
  if (runs.length !== context.visibleTokens.length) {
    return { ok: false, reason: "unsupported" };
  }
  for (let index = 0; index < runs.length; index += 1) {
    const run = runs[index]!;
    const token = context.visibleTokens[index]!;
    const runChildren = directChildElements(context.scan, run);
    const textChildren = runChildren.filter((child) => child.tagName === "w:t");
    const hasCanonicalShape =
      (runChildren.length === 1 && runChildren[0]?.tagName === "w:t") ||
      (runChildren.length === 2 &&
        runChildren[0]?.tagName === "w:rPr" &&
        runChildren[1]?.tagName === "w:t");
    if (
      token.enclosingRun.start !== run.start ||
      !hasCanonicalShape ||
      textChildren.length !== 1 ||
      textChildren[0]!.start !== token.range.start ||
      hasNonWhitespaceOutsideChildren(context.sourceXml, run, runChildren)
    ) {
      return { ok: false, reason: "unsupported" };
    }
  }

  const visibleText = context.visibleTokens.map((token) => token.text).join("");
  const totalLength = visibleText.length;
  if (input.endOffset > totalLength) {
    return { ok: false, reason: "stale" };
  }
  if (
    !isUtf16Boundary(visibleText, input.startOffset) ||
    !isUtf16Boundary(visibleText, input.endOffset)
  ) {
    return { ok: false, reason: "unsupported" };
  }
  const comments = model.metadata.comments ?? [];
  const commentId = nextCommentId(comments);
  const extendedParagraphId = allocateExtendedParagraphId(comments, commentId);
  const startMarker = `<w:commentRangeStart w:id="${commentId}"/>`;
  const endMarker = `<w:commentRangeEnd w:id="${commentId}"/><w:r><w:commentReference w:id="${commentId}"/></w:r>`;
  const replacements: Array<{ start: number; end: number; value: string }> = [];
  const nextChildren: TextRunNode[] = [];
  let offset = 0;
  let insertedStart = false;
  let insertedEnd = false;

  for (let index = 0; index < runs.length; index += 1) {
    const run = runs[index]!;
    const token = context.visibleTokens[index]!;
    const modeledRun = context.paragraph.children[index] as TextRunNode;
    const runStart = offset;
    const runEnd = offset + token.text.length;
    const cuts = [0, token.text.length];
    if (input.startOffset > runStart && input.startOffset < runEnd) {
      cuts.push(input.startOffset - runStart);
    }
    if (input.endOffset > runStart && input.endOffset < runEnd) {
      cuts.push(input.endOffset - runStart);
    }
    cuts.sort((left, right) => left - right);
    const uniqueCuts = cuts.filter(
      (value, cutIndex) => value !== cuts[cutIndex - 1]
    );
    let runReplacement = "";

    for (let cutIndex = 0; cutIndex < uniqueCuts.length - 1; cutIndex += 1) {
      const localStart = uniqueCuts[cutIndex]!;
      const localEnd = uniqueCuts[cutIndex + 1]!;
      const globalStart = runStart + localStart;
      const globalEnd = runStart + localEnd;
      if (!insertedStart && globalStart === input.startOffset) {
        runReplacement += startMarker;
        insertedStart = true;
      }
      const segment = token.text.slice(localStart, localEnd);
      if (segment.length > 0) {
        runReplacement += runXmlWithText(
          context.sourceXml,
          run,
          token.range,
          segment
        );
        nextChildren.push({
          type: "text",
          text: segment,
          style: modeledRun.style ? { ...modeledRun.style } : undefined,
        });
      }
      if (!insertedEnd && globalEnd === input.endOffset) {
        runReplacement += endMarker;
        insertedEnd = true;
      }
    }
    replacements.push({
      start: run.start,
      end: run.end,
      value: runReplacement,
    });
    offset = runEnd;
  }

  if (!insertedStart || !insertedEnd) {
    return { ok: false, reason: "stale" };
  }

  const next = cloneDocModel(model);
  const paragraph = next.nodes[input.nodeIndex];
  if (!paragraph || paragraph.type !== "paragraph") {
    return { ok: false, reason: "stale" };
  }
  paragraph.sourceXml = replaceRanges(context.sourceXml, replacements);
  paragraph.sourceTextPatch = undefined;
  paragraph.children = nextChildren;
  paragraph.sourceRunProvenance = sourceRunProvenanceForChildren(nextChildren);
  const definition: DocumentCommentDefinition = {
    id: commentId,
    author: input.author,
    initials: input.initials,
    date: input.date,
    text: input.text,
    resolved: false,
    extendedParagraphId,
    sourceResolved: false,
    resolutionDirty: false,
    isNew: true,
  };
  next.metadata.comments = [...(next.metadata.comments ?? []), definition];
  return { ok: true, model: next, commentId };
}
