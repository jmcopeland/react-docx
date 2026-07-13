import { describe, expect, it } from "vitest";

import type { DocModel, ParagraphNode } from "@extend-ai/react-docx-doc-model";
import {
  acceptParagraphRevision,
  createParagraphComment,
  rejectParagraphRevision,
  setCommentResolved,
} from "@extend-ai/react-docx-editor-ops";

function modelWithParagraph(paragraph: ParagraphNode): DocModel {
  if (
    paragraph.sourceXml &&
    !paragraph.sourceRunProvenance &&
    paragraph.children.every((child) => child.type === "text")
  ) {
    paragraph.sourceRunProvenance = {
      runs: paragraph.children.map((child) => {
        if (child.type !== "text") {
          throw new Error("expected text run");
        }
        return {
          style: child.style
            ? {
                ...child.style,
                runBorder: child.style.runBorder
                  ? { ...child.style.runBorder }
                  : undefined,
              }
            : undefined,
          link: child.link,
          noteReference: child.noteReference
            ? { ...child.noteReference }
            : undefined,
        };
      }),
    };
  }
  return {
    nodes: [paragraph],
    metadata: {
      sourceParts: 1,
      warnings: [],
      headerSections: [],
      footerSections: [],
      paragraphStyles: [],
    },
  };
}

describe("annotation editor operations", () => {
  it("accepts a safe insertion by unwrapping only its revision XML", () => {
    const sourceXml =
      '<w:p w:rsidR="AAAA"><w:r><w:t>Before </w:t></w:r><w:ins w:id="7" w:author="Ada"><w:r><w:rPr><w:b/></w:rPr><w:t>new</w:t></w:r></w:ins><w:proofErr w:type="spellStart"/><w:r><w:t> after</w:t></w:r></w:p>';
    const model = modelWithParagraph({
      type: "paragraph",
      sourceXml,
      children: [
        { type: "text", text: "Before " },
        { type: "text", text: "new", style: { bold: true } },
        { type: "text", text: " after" },
      ],
    });

    const result = acceptParagraphRevision(model, {
      nodeIndex: 0,
      revisionId: "7",
      kind: "insertion",
    });

    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const paragraph = result.model.nodes[0] as ParagraphNode;
    expect(paragraph.sourceXml).toBe(
      '<w:p w:rsidR="AAAA"><w:r><w:t>Before </w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>new</w:t></w:r><w:proofErr w:type="spellStart"/><w:r><w:t> after</w:t></w:r></w:p>'
    );
    expect(
      paragraph.children.map((run) => run.type === "text" && run.text)
    ).toEqual(["Before ", "new", " after"]);
    expect(model.nodes[0]).toMatchObject({ sourceXml });
  });

  it("rejects a safe insertion and removes only its modeled text", () => {
    const model = modelWithParagraph({
      type: "paragraph",
      sourceXml:
        '<w:p><w:r><w:t>A</w:t></w:r><w:ins w:id="8"><w:r><w:t>B</w:t></w:r></w:ins><w:r><w:t>C</w:t></w:r></w:p>',
      children: [
        { type: "text", text: "A" },
        { type: "text", text: "B" },
        { type: "text", text: "C" },
      ],
    });
    const result = rejectParagraphRevision(model, {
      nodeIndex: 0,
      revisionId: 8,
      kind: "insertion",
    });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    const paragraph = result.model.nodes[0] as ParagraphNode;
    expect(paragraph.sourceXml).toBe(
      "<w:p><w:r><w:t>A</w:t></w:r><w:r><w:t>C</w:t></w:r></w:p>"
    );
    expect(paragraph.children).toEqual([
      { type: "text", text: "A" },
      { type: "text", text: "C" },
    ]);
  });

  it("accepts and rejects safe deletions without regenerating the paragraph", () => {
    const sourceXml =
      '<w:p custom="keep"><w:r><w:t>A</w:t></w:r><w:del w:id="9"><w:r><w:rPr><w:i/><w:color w:val="FF0000"/></w:rPr><w:delText xml:space="preserve"> old </w:delText></w:r></w:del><w:r><w:t>C</w:t></w:r></w:p>';
    const model = modelWithParagraph({
      type: "paragraph",
      sourceXml,
      children: [
        { type: "text", text: "A" },
        { type: "text", text: "C" },
      ],
    });

    const accepted = acceptParagraphRevision(model, {
      nodeIndex: 0,
      revisionId: "9",
      kind: "deletion",
    });
    expect(accepted.ok).toBe(true);
    if (accepted.ok) {
      expect((accepted.model.nodes[0] as ParagraphNode).sourceXml).toBe(
        '<w:p custom="keep"><w:r><w:t>A</w:t></w:r><w:r><w:t>C</w:t></w:r></w:p>'
      );
    }

    const rejected = rejectParagraphRevision(model, {
      nodeIndex: 0,
      revisionId: "9",
      kind: "deletion",
    });
    expect(rejected.ok).toBe(true);
    if (!rejected.ok) return;
    const paragraph = rejected.model.nodes[0] as ParagraphNode;
    expect(paragraph.sourceXml).toContain(
      '<w:rPr><w:i/><w:color w:val="FF0000"/></w:rPr><w:t xml:space="preserve"> old </w:t>'
    );
    expect(paragraph.sourceXml).not.toContain("<w:del");
    expect(paragraph.children).toEqual([
      { type: "text", text: "A" },
      {
        type: "text",
        text: " old ",
        style: expect.objectContaining({ italic: true, color: "#ff0000" }),
      },
      { type: "text", text: "C" },
    ]);
  });

  it("fails closed for duplicate ids, nested revisions, and stale targets", () => {
    const duplicate = modelWithParagraph({
      type: "paragraph",
      sourceXml:
        '<w:p><w:ins w:id="4"><w:r><w:t>A</w:t></w:r></w:ins><w:ins w:id="4"><w:r><w:t>B</w:t></w:r></w:ins></w:p>',
      children: [
        { type: "text", text: "A" },
        { type: "text", text: "B" },
      ],
    });
    expect(
      acceptParagraphRevision(duplicate, {
        nodeIndex: 0,
        revisionId: 4,
        kind: "insertion",
      })
    ).toEqual({ ok: false, reason: "unsafe-xml" });

    const nested = modelWithParagraph({
      type: "paragraph",
      sourceXml:
        '<w:p><w:ins w:id="5"><w:ins w:id="6"><w:r><w:t>A</w:t></w:r></w:ins></w:ins></w:p>',
      children: [{ type: "text", text: "A" }],
    });
    expect(
      acceptParagraphRevision(nested, {
        nodeIndex: 0,
        revisionId: 5,
        kind: "insertion",
      })
    ).toEqual({ ok: false, reason: "unsupported" });
    expect(
      acceptParagraphRevision(duplicate, {
        nodeIndex: 0,
        revisionId: 999,
        kind: "insertion",
      })
    ).toEqual({ ok: false, reason: "stale" });
    expect(
      acceptParagraphRevision(duplicate, {
        nodeIndex: 0,
        revisionId: 4,
        kind: "move-to",
      })
    ).toEqual({ ok: false, reason: "unsupported" });
  });

  it("fails closed when source XML tag names differ by case", () => {
    const model = modelWithParagraph({
      type: "paragraph",
      sourceXml: '<w:p><w:ins w:id="1"><w:r><w:t>A</w:t></w:r></w:ins></W:P>',
      children: [{ type: "text", text: "A" }],
    });
    expect(
      acceptParagraphRevision(model, {
        nodeIndex: 0,
        revisionId: 1,
        kind: "insertion",
      })
    ).toEqual({ ok: false, reason: "unsafe-xml" });
  });

  it("rejects a stale run-metadata target before touching source XML", () => {
    const model = modelWithParagraph({
      type: "paragraph",
      sourceXml:
        '<w:p><w:ins w:id="2"><w:r><w:rPr><w:b/></w:rPr><w:t>A</w:t></w:r></w:ins></w:p>',
      children: [{ type: "text", text: "A", style: { bold: true } }],
    });
    const paragraph = model.nodes[0] as ParagraphNode;
    const run = paragraph.children[0];
    if (run?.type === "text") {
      run.style = { italic: true };
    }

    expect(
      acceptParagraphRevision(model, {
        nodeIndex: 0,
        revisionId: 2,
        kind: "insertion",
      })
    ).toEqual({ ok: false, reason: "stale" });
  });

  it("updates comment resolution while retaining imported provenance", () => {
    const model = modelWithParagraph({
      type: "paragraph",
      sourceXml: "<w:p><w:r><w:t>A</w:t></w:r></w:p>",
      children: [{ type: "text", text: "A" }],
    });
    model.metadata.comments = [
      {
        id: 2,
        text: "Check this",
        resolved: false,
        sourceResolved: false,
        resolutionDirty: false,
        sourceXml:
          '<w:comment w:id="2"><w:p><w:r><w:t>Check this</w:t></w:r></w:p></w:comment>',
        extendedParagraphId: "1234ABCD",
      },
    ];

    const result = setCommentResolved(model, 2, true);
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.model.metadata.comments?.[0]).toMatchObject({
      resolved: true,
      sourceResolved: false,
      resolutionDirty: true,
      sourceXml: model.metadata.comments[0]?.sourceXml,
    });
    expect(model.metadata.comments[0]?.resolved).toBe(false);
    const unchanged = setCommentResolved(model, 2, false);
    expect(unchanged.ok).toBe(true);
    if (unchanged.ok) {
      expect(unchanged.model).toBe(model);
    }
    expect(setCommentResolved(model, 404, true)).toEqual({
      ok: false,
      reason: "stale",
    });
  });

  it("creates a comment by splitting a plain-text run and fixes xml:space", () => {
    const model = modelWithParagraph({
      type: "paragraph",
      sourceXml:
        '<w:p keep="yes"><w:r><w:rPr><w:b/></w:rPr><w:t xml:space="default">Hello world</w:t></w:r></w:p>',
      children: [{ type: "text", text: "Hello world", style: { bold: true } }],
    });

    const result = createParagraphComment(model, {
      nodeIndex: 0,
      startOffset: 5,
      endOffset: 11,
      text: "Greeting wording",
      author: "Ada",
      initials: "AL",
      date: "2026-07-09T12:00:00Z",
    });
    expect(result.ok).toBe(true);
    if (!result.ok) return;
    expect(result.commentId).toBe(0);
    const paragraph = result.model.nodes[0] as ParagraphNode;
    expect(paragraph.sourceXml).toContain('<w:p keep="yes">');
    expect(paragraph.sourceXml).toContain('<w:commentRangeStart w:id="0"/>');
    expect(paragraph.sourceXml).toContain('<w:commentRangeEnd w:id="0"/>');
    expect(paragraph.sourceXml).toContain(
      '<w:t xml:space="preserve"> world</w:t>'
    );
    expect(paragraph.children).toEqual([
      { type: "text", text: "Hello", style: { bold: true } },
      { type: "text", text: " world", style: { bold: true } },
    ]);
    expect(result.model.metadata.comments?.[0]).toMatchObject({
      id: 0,
      text: "Greeting wording",
      resolved: false,
      sourceResolved: false,
      resolutionDirty: false,
      isNew: true,
      extendedParagraphId: "C0000000",
    });
  });

  it("rejects comment creation when splitting would duplicate non-text run content", () => {
    const model = modelWithParagraph({
      type: "paragraph",
      sourceXml: "<w:p><w:r><w:softHyphen/><w:t>Hello</w:t></w:r></w:p>",
      children: [{ type: "text", text: "Hello" }],
    });

    expect(
      createParagraphComment(model, {
        nodeIndex: 0,
        startOffset: 1,
        endOffset: 4,
        text: "Check this",
      })
    ).toEqual({ ok: false, reason: "unsupported" });
  });

  it("rejects XML-invalid comment fields, invalid dates, and split surrogate pairs", () => {
    const createModel = () =>
      modelWithParagraph({
        type: "paragraph",
        sourceXml: "<w:p><w:r><w:t>A😀B</w:t></w:r></w:p>",
        children: [{ type: "text", text: "A😀B" }],
      });

    expect(
      createParagraphComment(createModel(), {
        nodeIndex: 0,
        startOffset: 0,
        endOffset: 1,
        text: "bad\u0000text",
      })
    ).toEqual({ ok: false, reason: "unsupported" });
    expect(
      createParagraphComment(createModel(), {
        nodeIndex: 0,
        startOffset: 0,
        endOffset: 1,
        text: "Comment",
        author: "bad\ud800",
      })
    ).toEqual({ ok: false, reason: "unsupported" });
    expect(
      createParagraphComment(createModel(), {
        nodeIndex: 0,
        startOffset: 0,
        endOffset: 1,
        text: "Comment",
        date: "2026-02-29T12:00:00Z",
      })
    ).toEqual({ ok: false, reason: "unsupported" });
    expect(
      createParagraphComment(createModel(), {
        nodeIndex: 0,
        startOffset: 2,
        endOffset: 3,
        text: "Comment",
      })
    ).toEqual({ ok: false, reason: "unsupported" });
  });
});
