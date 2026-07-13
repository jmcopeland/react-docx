/** A balanced XML element range produced without backtracking regular expressions. */
export interface XmlElementRange {
  tagName: string;
  start: number;
  openEnd: number;
  closeStart: number;
  end: number;
  openTag: string;
  parentStart?: number;
  selfClosing: boolean;
}

export interface XmlScanResult {
  elements: XmlElementRange[];
  malformed: boolean;
}

interface OpenElement {
  tagName: string;
  start: number;
  openEnd: number;
  openTag: string;
  parentStart?: number;
}

function findMarkupEnd(xml: string, start: number): number | undefined {
  let quote: '"' | "'" | undefined;
  for (let index = start; index < xml.length; index += 1) {
    const character = xml[index];
    if (quote) {
      if (character === quote) {
        quote = undefined;
      }
      continue;
    }
    if (character === '"' || character === "'") {
      quote = character;
      continue;
    }
    if (character === ">") {
      return index + 1;
    }
  }
  return undefined;
}

function parseTagName(token: string, closing: boolean): string | undefined {
  let index = closing ? 2 : 1;
  while (index < token.length && /\s/.test(token[index] ?? "")) {
    index += 1;
  }
  const start = index;
  while (index < token.length && !/[\s/>]/.test(token[index] ?? "")) {
    index += 1;
  }
  return index > start ? token.slice(start, index) : undefined;
}

function isSelfClosing(token: string): boolean {
  let index = token.length - 2;
  while (index >= 0 && /\s/.test(token[index] ?? "")) {
    index -= 1;
  }
  return token[index] === "/";
}

/**
 * Scans XML once and pairs opening/closing elements with an explicit stack.
 * This intentionally rejects mismatched or unterminated markup instead of
 * attempting a lossy recovery before a surgical OOXML mutation.
 */
export function scanBalancedXml(xml: string): XmlScanResult {
  const elements: XmlElementRange[] = [];
  const stack: OpenElement[] = [];
  let malformed = false;
  let cursor = 0;

  while (cursor < xml.length) {
    const start = xml.indexOf("<", cursor);
    if (start < 0) {
      break;
    }

    if (xml.startsWith("<!--", start)) {
      const end = xml.indexOf("-->", start + 4);
      if (end < 0) {
        malformed = true;
        break;
      }
      cursor = end + 3;
      continue;
    }
    if (xml.startsWith("<![CDATA[", start)) {
      const end = xml.indexOf("]]>", start + 9);
      if (end < 0) {
        malformed = true;
        break;
      }
      cursor = end + 3;
      continue;
    }
    if (xml.startsWith("<?", start)) {
      const end = xml.indexOf("?>", start + 2);
      if (end < 0) {
        malformed = true;
        break;
      }
      cursor = end + 2;
      continue;
    }

    const end = findMarkupEnd(xml, start + 1);
    if (end === undefined) {
      malformed = true;
      break;
    }
    const token = xml.slice(start, end);
    if (token.startsWith("<!")) {
      cursor = end;
      continue;
    }

    const closing = /^<\s*\//.test(token);
    const tagName = parseTagName(token, closing);
    if (!tagName) {
      malformed = true;
      cursor = end;
      continue;
    }

    if (closing) {
      const opener = stack.pop();
      if (!opener || opener.tagName !== tagName) {
        malformed = true;
        cursor = end;
        continue;
      }
      elements.push({
        tagName: opener.tagName,
        start: opener.start,
        openEnd: opener.openEnd,
        closeStart: start,
        end,
        openTag: opener.openTag,
        parentStart: opener.parentStart,
        selfClosing: false,
      });
      cursor = end;
      continue;
    }

    const parentStart = stack.at(-1)?.start;
    if (isSelfClosing(token)) {
      elements.push({
        tagName,
        start,
        openEnd: end,
        closeStart: end,
        end,
        openTag: token,
        parentStart,
        selfClosing: true,
      });
    } else {
      stack.push({
        tagName,
        start,
        openEnd: end,
        openTag: token,
        parentStart,
      });
    }
    cursor = end;
  }

  if (stack.length > 0) {
    malformed = true;
  }
  elements.sort(
    (left, right) => left.start - right.start || right.end - left.end
  );
  return { elements, malformed };
}

export function elementsNamed(
  scan: XmlScanResult,
  tagName: string
): XmlElementRange[] {
  return scan.elements.filter((element) => element.tagName === tagName);
}

export function directChildElements(
  scan: XmlScanResult,
  parent: XmlElementRange
): XmlElementRange[] {
  return scan.elements.filter(
    (element) => element.parentStart === parent.start
  );
}

export function enclosingElements(
  scan: XmlScanResult,
  range: Pick<XmlElementRange, "start" | "end">
): XmlElementRange[] {
  return scan.elements
    .filter(
      (element) => element.start < range.start && element.end >= range.end
    )
    .sort((left, right) => left.end - left.start - (right.end - right.start));
}

export function xmlAttribute(
  openTag: string,
  name: string
): string | undefined {
  let cursor = 1;
  while (cursor < openTag.length) {
    while (cursor < openTag.length && /[^\s]/.test(openTag[cursor] ?? "")) {
      cursor += 1;
    }
    while (cursor < openTag.length && /\s/.test(openTag[cursor] ?? "")) {
      cursor += 1;
    }
    if (
      cursor >= openTag.length ||
      openTag[cursor] === "/" ||
      openTag[cursor] === ">"
    ) {
      break;
    }
    const nameStart = cursor;
    while (cursor < openTag.length && !/[\s=/>]/.test(openTag[cursor] ?? "")) {
      cursor += 1;
    }
    const attributeName = openTag.slice(nameStart, cursor);
    while (cursor < openTag.length && /\s/.test(openTag[cursor] ?? "")) {
      cursor += 1;
    }
    if (openTag[cursor] !== "=") {
      while (cursor < openTag.length && !/[\s>]/.test(openTag[cursor] ?? "")) {
        cursor += 1;
      }
      continue;
    }
    cursor += 1;
    while (cursor < openTag.length && /\s/.test(openTag[cursor] ?? "")) {
      cursor += 1;
    }
    const quote = openTag[cursor];
    if (quote !== '"' && quote !== "'") {
      continue;
    }
    cursor += 1;
    const valueStart = cursor;
    const valueEnd = openTag.indexOf(quote, valueStart);
    if (valueEnd < 0) {
      return undefined;
    }
    if (attributeName === name) {
      return decodeXmlText(openTag.slice(valueStart, valueEnd));
    }
    cursor = valueEnd + 1;
  }
  return undefined;
}

export function setXmlAttribute(
  openTag: string,
  name: string,
  value: string
): string {
  let cursor = 1;
  while (cursor < openTag.length) {
    while (cursor < openTag.length && /[^\s]/.test(openTag[cursor] ?? "")) {
      cursor += 1;
    }
    while (cursor < openTag.length && /\s/.test(openTag[cursor] ?? "")) {
      cursor += 1;
    }
    if (
      cursor >= openTag.length ||
      openTag[cursor] === "/" ||
      openTag[cursor] === ">"
    ) {
      break;
    }
    const nameStart = cursor;
    while (cursor < openTag.length && !/[\s=/>]/.test(openTag[cursor] ?? "")) {
      cursor += 1;
    }
    const attributeName = openTag.slice(nameStart, cursor);
    while (cursor < openTag.length && /\s/.test(openTag[cursor] ?? "")) {
      cursor += 1;
    }
    if (openTag[cursor] !== "=") {
      continue;
    }
    cursor += 1;
    while (cursor < openTag.length && /\s/.test(openTag[cursor] ?? "")) {
      cursor += 1;
    }
    const quote = openTag[cursor];
    if (quote !== '"' && quote !== "'") {
      continue;
    }
    cursor += 1;
    const valueStart = cursor;
    const valueEnd = openTag.indexOf(quote, valueStart);
    if (valueEnd < 0) {
      return openTag;
    }
    if (attributeName === name) {
      return `${openTag.slice(0, valueStart)}${value}${openTag.slice(
        valueEnd
      )}`;
    }
    cursor = valueEnd + 1;
  }

  let insertAt = openTag.length - 1;
  while (insertAt > 0 && /\s/.test(openTag[insertAt - 1] ?? "")) {
    insertAt -= 1;
  }
  if (openTag[insertAt - 1] === "/") {
    insertAt -= 1;
  }
  return `${openTag.slice(0, insertAt)} ${name}="${value}"${openTag.slice(
    insertAt
  )}`;
}

export function decodeXmlText(text: string): string {
  return text.replace(
    /&(?:#x[0-9a-f]+|#[0-9]+|amp|lt|gt|quot|apos);/gi,
    (entity) => {
      const body = entity.slice(1, -1);
      if (body[0] === "#") {
        const hexadecimal = body[1]?.toLowerCase() === "x";
        const digits = body.slice(hexadecimal ? 2 : 1);
        const codePoint = Number.parseInt(digits, hexadecimal ? 16 : 10);
        if (Number.isFinite(codePoint) && codePoint > 0) {
          try {
            return String.fromCodePoint(codePoint);
          } catch {
            return entity;
          }
        }
        return entity;
      }
      switch (body.toLowerCase()) {
        case "amp":
          return "&";
        case "lt":
          return "<";
        case "gt":
          return ">";
        case "quot":
          return '"';
        case "apos":
          return "'";
        default:
          return entity;
      }
    }
  );
}

export function escapeXmlText(text: string): string {
  return text
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}
