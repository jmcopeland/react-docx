import { expect, test, type Locator } from "@playwright/test";

import { createZip } from "../unit/helpers/zip";

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>`;

const ROOT_RELATIONSHIPS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>Fidelity smoke page one</w:t></w:r></w:p>
    <w:p><w:r><w:br w:type="page"/></w:r></w:p>
    <w:p><w:r><w:t>Fidelity smoke page two</w:t></w:r></w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="708" w:footer="708" w:gutter="0"/>
    </w:sectPr>
  </w:body>
</w:document>`;

function smokeDocx(): Buffer {
  return Buffer.from(
    createZip([
      { name: "[Content_Types].xml", content: CONTENT_TYPES_XML },
      { name: "_rels/.rels", content: ROOT_RELATIONSHIPS_XML },
      { name: "word/document.xml", content: DOCUMENT_XML },
    ])
  );
}

async function selectTextOffsets(
  paragraph: Locator,
  startOffset: number,
  endOffset: number
): Promise<void> {
  await paragraph.evaluate(
    (element, offsets) => {
      const resolvePosition = (targetOffset: number): [Node, number] => {
        const walker = document.createTreeWalker(element, NodeFilter.SHOW_TEXT);
        let cursor = 0;
        let lastTextNode: Text | undefined;
        while (walker.nextNode()) {
          const textNode = walker.currentNode;
          if (!(textNode instanceof Text)) {
            continue;
          }
          lastTextNode = textNode;
          const nextCursor = cursor + textNode.data.length;
          if (targetOffset <= nextCursor) {
            return [textNode, Math.max(0, targetOffset - cursor)];
          }
          cursor = nextCursor;
        }
        if (!lastTextNode) {
          throw new Error("Expected editable paragraph text");
        }
        return [lastTextNode, lastTextNode.data.length];
      };

      const [startNode, start] = resolvePosition(offsets.startOffset);
      const [endNode, end] = resolvePosition(offsets.endOffset);
      element.focus();
      const selection = window.getSelection();
      if (!selection) {
        throw new Error("Selection API is unavailable");
      }
      selection.setBaseAndExtent(startNode, start, endNode, end);
      element.dispatchEvent(new MouseEvent("mouseup", { bubbles: true }));
    },
    { startOffset, endOffset }
  );
}

test("loads a generated DOCX and settles on its explicit page geometry", async ({
  page,
}) => {
  await page.goto("/");

  await page.locator('input[type="file"][accept*=".doc"]').setInputFiles({
    name: "fidelity-smoke.docx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: smokeDocx(),
  });

  await expect(page.getByText("Fidelity smoke page one").first()).toBeVisible({
    timeout: 20_000,
  });
  await expect(page.getByText("Fidelity smoke page two").first()).toBeVisible();

  await page.evaluate(async () => {
    await document.fonts.ready;
  });

  const renderedPages = page.locator("[data-docx-page-index]");
  await expect.poll(() => renderedPages.count()).toBe(2);
  await expect(renderedPages.nth(0)).toContainText("Fidelity smoke page one");
  await expect(renderedPages.nth(1)).toContainText("Fidelity smoke page two");
});

test("collapses the native selection after deleting selected text", async ({
  page,
}) => {
  await page.goto("/");

  await page.locator('input[type="file"][accept*=".doc"]').setInputFiles({
    name: "selection-delete.docx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: smokeDocx(),
  });

  const paragraph = page
    .locator(
      '[data-docx-paragraph-kind="paragraph"][data-docx-paragraph-node-index="0"][contenteditable="true"]'
    )
    .first();
  await expect(paragraph).toBeVisible({ timeout: 20_000 });
  await expect(paragraph).toContainText("Fidelity smoke page one");

  await paragraph.evaluate((element) => {
    const textNode = element.firstChild?.firstChild ?? element.firstChild;
    if (!(textNode instanceof Text)) {
      throw new Error("Expected an editable paragraph text node");
    }
    element.focus();
    const selection = window.getSelection();
    if (!selection) {
      throw new Error("Selection API is unavailable");
    }
    selection.setBaseAndExtent(textNode, 9, textNode, 15);
  });
  await expect
    .poll(() => page.evaluate(() => window.getSelection()?.toString()))
    .toBe("smoke ");

  await page.keyboard.press("Delete");

  await expect(paragraph).toContainText("Fidelity page one");
  await expect
    .poll(() =>
      page.evaluate(() => {
        const selection = window.getSelection();
        return {
          collapsed: selection?.isCollapsed,
          selectedText: selection?.toString(),
        };
      })
    )
    .toEqual({ collapsed: true, selectedText: "" });
});

test("applies and toggles toolbar styles on each new selection", async ({
  page,
}) => {
  await page.goto("/");
  await page.locator('input[type="file"][accept*=".doc"]').setInputFiles({
    name: "toolbar-selection.docx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: smokeDocx(),
  });

  const paragraph = page
    .locator(
      '[data-docx-paragraph-kind="paragraph"][data-docx-paragraph-node-index="0"][contenteditable="true"]'
    )
    .first();
  await expect(paragraph).toContainText("Fidelity smoke page one", {
    timeout: 20_000,
  });

  const rangeState = (start: number, end: number) =>
    page.evaluate(
      ({ start, end }) =>
        (window as any).__DOCX_TEST_HOOKS__.getRangeState({
          start: {
            location: { kind: "paragraph", nodeIndex: 0 },
            offset: start,
          },
          end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: end },
        }),
      { start, end }
    );
  const boldButton = page.getByRole("button", { name: "Bold", exact: true });
  const italicButton = page.getByRole("button", {
    name: "Italic",
    exact: true,
  });

  await selectTextOffsets(paragraph, 0, 8);
  await boldButton.click();
  await expect
    .poll(async () =>
      (
        await rangeState(0, 8)
      )?.styles.every((style: { bold: boolean | null }) => style.bold === true)
    )
    .toBe(true);

  await selectTextOffsets(paragraph, 9, 14);
  await expect
    .poll(() => page.evaluate(() => window.getSelection()?.toString()))
    .toBe("smoke");
  await expect
    .poll(() =>
      page.evaluate(() => {
        const range = (window as any).__DOCX_TEST_HOOKS__.getActionState()
          .activeTextRange;
        const selectedRunStyle = (
          window as any
        ).__DOCX_TEST_HOOKS__.getActionState().selectedRunStyle;
        return range
          ? {
              start: range.start.offset,
              end: range.end.offset,
              nodeIndex: range.start.location.nodeIndex,
              bold: selectedRunStyle.bold,
            }
          : null;
      })
    )
    .toEqual({ start: 9, end: 14, nodeIndex: 0, bold: null });
  await boldButton.click();
  expect(
    await page.evaluate(() => {
      const hooks = (window as any).__DOCX_TEST_HOOKS__;
      const actionRange = hooks.getActionState().activeTextRange;
      const range = (start: number, end: number) =>
        hooks.getRangeState({
          start: {
            location: { kind: "paragraph", nodeIndex: 0 },
            offset: start,
          },
          end: { location: { kind: "paragraph", nodeIndex: 0 }, offset: end },
        });
      return {
        selectedText: window.getSelection()?.toString(),
        actionRange: actionRange
          ? [actionRange.start.offset, actionRange.end.offset]
          : null,
        firstBold: range(0, 8)?.styles.map((style: any) => style.bold),
        secondBold: range(9, 14)?.styles.map((style: any) => style.bold),
      };
    })
  ).toEqual({
    selectedText: "smoke",
    actionRange: [9, 14],
    firstBold: [true],
    secondBold: [true],
  });
  await expect
    .poll(async () =>
      (
        await rangeState(9, 14)
      )?.styles.map((style: { bold: boolean | null }) => style.bold)
    )
    .toEqual([true]);
  await expect
    .poll(() => page.evaluate(() => window.getSelection()?.toString()))
    .toBe("smoke");

  await boldButton.click();
  await expect
    .poll(async () =>
      (
        await rangeState(9, 14)
      )?.styles.every((style: { bold: boolean | null }) => style.bold === false)
    )
    .toBe(true);

  await italicButton.click();
  await expect
    .poll(async () =>
      (
        await rangeState(9, 14)
      )?.styles.every(
        (style: { italic: boolean | null }) => style.italic === true
      )
    )
    .toBe(true);
  await italicButton.click();
  await expect
    .poll(async () =>
      (
        await rangeState(9, 14)
      )?.styles.every(
        (style: { italic: boolean | null }) => style.italic === false
      )
    )
    .toBe(true);

  await expect
    .poll(async () =>
      (
        await rangeState(0, 8)
      )?.styles.every((style: { bold: boolean | null }) => style.bold === true)
    )
    .toBe(true);
});
