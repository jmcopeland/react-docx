import { expect, test, type Page } from "@playwright/test";

import { createZip } from "../unit/helpers/zip";

const CONTENT_TYPES_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
  <Override PartName="/word/comments.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml"/>
  <Override PartName="/word/commentsExtended.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.commentsExtended+xml"/>
</Types>`;

const ROOT_RELATIONSHIPS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>`;

const DOCUMENT_RELATIONSHIPS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments" Target="comments.xml"/>
  <Relationship Id="rId2" Type="http://schemas.microsoft.com/office/2011/relationships/commentsExtended" Target="commentsExtended.xml"/>
</Relationships>`;

const DOCUMENT_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t xml:space="preserve">Before </w:t></w:r>
      <w:ins w:id="17" w:author="Ada" w:date="2026-07-09T12:00:00Z"><w:r><w:t>new</w:t></w:r></w:ins>
      <w:r><w:t xml:space="preserve"> and </w:t></w:r>
      <w:commentRangeStart w:id="1"/>
      <w:r><w:t>commented</w:t></w:r>
      <w:commentRangeEnd w:id="1"/>
      <w:r><w:commentReference w:id="1"/></w:r>
      <w:r><w:t xml:space="preserve"> after.</w:t></w:r>
    </w:p>
    <w:sectPr>
      <w:pgSz w:w="12240" w:h="15840"/>
      <w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/>
    </w:sectPr>
  </w:body>
</w:document>`;

const COMMENTS_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:comment w:id="1" w:author="Ada" w:initials="AL" w:date="2026-07-09T12:00:00Z">
    <w:p w14:paraId="11111111"><w:r><w:t>Please review this phrase.</w:t></w:r></w:p>
  </w:comment>
</w:comments>`;

const COMMENTS_EXTENDED_XML = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w15:commentsEx xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w15:commentEx w15:paraId="11111111" w15:done="0"/>
</w15:commentsEx>`;

type AnnotationSummary = {
  fileName: string;
  semanticModelDigest: string;
  trackedChangeCount: number;
  commentCount: number;
  resolvedCommentCount: number;
  canUndo: boolean;
  canRedo: boolean;
};

function annotationDocx(): Buffer {
  return Buffer.from(
    createZip([
      { name: "[Content_Types].xml", content: CONTENT_TYPES_XML },
      { name: "_rels/.rels", content: ROOT_RELATIONSHIPS_XML },
      { name: "word/document.xml", content: DOCUMENT_XML },
      {
        name: "word/_rels/document.xml.rels",
        content: DOCUMENT_RELATIONSHIPS_XML,
      },
      { name: "word/comments.xml", content: COMMENTS_XML },
      { name: "word/commentsExtended.xml", content: COMMENTS_EXTENDED_XML },
    ])
  );
}

async function summary(page: Page): Promise<AnnotationSummary> {
  return page.evaluate(() => {
    const hooks = (
      window as typeof window & {
        __DOCX_TEST_HOOKS__?: { getSummary: () => AnnotationSummary };
      }
    ).__DOCX_TEST_HOOKS__;
    if (!hooks) {
      throw new Error("DOCX test hooks are unavailable");
    }
    return hooks.getSummary();
  });
}

async function waitForSummary(
  page: Page,
  predicate: (value: AnnotationSummary) => boolean
): Promise<AnnotationSummary> {
  await expect.poll(async () => predicate(await summary(page))).toBe(true);
  return summary(page);
}

test("accepts and resolves annotations through export and fresh reopen", async ({
  context,
  page,
}, testInfo) => {
  await page.goto("/");
  await page.locator('input[type="file"][accept*=".doc"]').setInputFiles({
    name: "annotation-roundtrip.docx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: annotationDocx(),
  });

  const imported = await waitForSummary(
    page,
    (value) =>
      value.fileName === "annotation-roundtrip.docx" &&
      value.trackedChangeCount === 1 &&
      value.commentCount === 1 &&
      value.resolvedCommentCount === 0
  );

  const staleHandles = await page.evaluate(() => ({
    change: (window as any).__DOCX_TEST_HOOKS__.getTrackedChange(0),
    comment: (window as any).__DOCX_TEST_HOOKS__.getComment(1),
  }));

  await page.locator('input[type="file"][accept*=".doc"]').setInputFiles({
    name: "annotation-roundtrip-reloaded.docx",
    mimeType:
      "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    buffer: annotationDocx(),
  });
  const reloaded = await waitForSummary(
    page,
    (value) =>
      value.fileName === "annotation-roundtrip-reloaded.docx" &&
      value.trackedChangeCount === 1 &&
      value.commentCount === 1 &&
      value.resolvedCommentCount === 0
  );
  expect(reloaded.semanticModelDigest).toBe(imported.semanticModelDigest);

  const staleChangeFailure = await page.evaluate((change) => {
    try {
      (window as any).__DOCX_TEST_HOOKS__.acceptTrackedChangeHandle(change);
      return null;
    } catch (error) {
      return error instanceof Error ? error.message : String(error);
    }
  }, staleHandles.change);
  expect(staleChangeFailure).toMatch(/stale/);

  const staleCommentFailure = await page.evaluate((comment) => {
    try {
      (window as any).__DOCX_TEST_HOOKS__.setCommentResolvedHandle(
        comment,
        true
      );
      return null;
    } catch (error) {
      return error instanceof Error ? error.message : String(error);
    }
  }, staleHandles.comment);
  expect(staleCommentFailure).toMatch(/stale/);
  const afterStaleCommands = await summary(page);
  expect(afterStaleCommands).toMatchObject({
    semanticModelDigest: reloaded.semanticModelDigest,
    trackedChangeCount: reloaded.trackedChangeCount,
    commentCount: reloaded.commentCount,
    resolvedCommentCount: reloaded.resolvedCommentCount,
    canUndo: reloaded.canUndo,
    canRedo: reloaded.canRedo,
  });

  await page.evaluate(() => {
    (window as any).__DOCX_TEST_HOOKS__.acceptTrackedChange(0);
  });
  const accepted = await waitForSummary(
    page,
    (value) => value.trackedChangeCount === 0 && value.canUndo
  );
  expect(accepted.semanticModelDigest).not.toBe(imported.semanticModelDigest);

  await page.evaluate(() => (window as any).__DOCX_TEST_HOOKS__.undo());
  const undone = await waitForSummary(
    page,
    (value) => value.trackedChangeCount === 1 && value.canRedo
  );
  expect(undone.semanticModelDigest).toBe(imported.semanticModelDigest);

  await page.evaluate(() => (window as any).__DOCX_TEST_HOOKS__.redo());
  const redone = await waitForSummary(
    page,
    (value) => value.trackedChangeCount === 0
  );
  expect(redone.semanticModelDigest).toBe(accepted.semanticModelDigest);

  await page.evaluate(() => {
    (window as any).__DOCX_TEST_HOOKS__.setCommentResolved(1, true);
  });
  const resolved = await waitForSummary(
    page,
    (value) => value.commentCount === 1 && value.resolvedCommentCount === 1
  );
  expect(resolved.semanticModelDigest).not.toBe(redone.semanticModelDigest);

  const downloadPromise = page.waitForEvent("download");
  await page.evaluate(() => (window as any).__DOCX_TEST_HOOKS__.exportDocx());
  const download = await downloadPromise;
  const exportPath = testInfo.outputPath("annotation-roundtrip-exported.docx");
  await download.saveAs(exportPath);

  const reopenedPage = await context.newPage();
  await reopenedPage.goto("/");
  await reopenedPage
    .locator('input[type="file"][accept*=".doc"]')
    .setInputFiles(exportPath);
  const reopened = await waitForSummary(
    reopenedPage,
    (value) =>
      value.trackedChangeCount === 0 &&
      value.commentCount === 1 &&
      value.resolvedCommentCount === 1
  );
  expect(reopened.semanticModelDigest).toBe(resolved.semanticModelDigest);
  await expect(
    reopenedPage.locator('[data-docx-paragraph-kind="paragraph"]').first()
  ).toContainText("Before new and commented after.");
  await reopenedPage.close();
});
