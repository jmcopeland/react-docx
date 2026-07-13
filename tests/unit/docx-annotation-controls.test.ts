import * as React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";

import type { DocModel } from "@extend-ai/react-docx-doc-model";
import {
  DocxEditorViewer,
  defaultStarterModel,
  useDocxEditor,
  useDocxTrackChanges,
  type DocxTrackedChange,
} from "../../packages/react-viewer/src/editor";

function insertedTextModel(): DocModel {
  return {
    ...defaultStarterModel,
    nodes: [
      {
        type: "paragraph",
        sourceXml:
          '<w:p><w:r><w:t>Before </w:t></w:r><w:ins w:id="17" w:author="Ada"><w:r><w:t>new</w:t></w:r></w:ins></w:p>',
        children: [
          { type: "text", text: "Before " },
          { type: "text", text: "new" },
        ],
      },
    ],
  };
}

describe("annotation editor controls", () => {
  it("exposes revision ids and safe accept/reject actions", () => {
    let changes: DocxTrackedChange[] = [];
    let acceptChange: unknown;
    let rejectChange: unknown;
    const markup = renderToStaticMarkup(
      React.createElement(function TrackedChangesProbe(): React.JSX.Element {
        const editor = useDocxEditor({
          starterModel: insertedTextModel(),
          initialShowTrackedChanges: true,
        });
        const tracked = useDocxTrackChanges(editor);
        changes = tracked.trackedChanges;
        acceptChange = tracked.acceptChange;
        rejectChange = tracked.rejectChange;
        return React.createElement(DocxEditorViewer, { editor });
      })
    );

    expect(changes).toHaveLength(1);
    expect(changes[0]).toMatchObject({
      kind: "insertion",
      revisionId: "17",
      location: { kind: "paragraph", nodeIndex: 0 },
    });
    expect(markup).toContain('data-docx-tracked-change="insertion"');
    expect(acceptChange).toBeTypeOf("function");
    expect(rejectChange).toBeTypeOf("function");
  });
});
