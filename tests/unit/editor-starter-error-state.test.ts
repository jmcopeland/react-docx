import * as React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import {
  defaultStarterModel,
  DocxEditorViewer,
  useDocxEditor,
  type DocxEditorController,
} from "../../packages/react-viewer/src/editor";

function DefaultEditorViewer(props: {
  importError?: Error;
}): React.JSX.Element {
  const editor = useDocxEditor();
  const resolvedEditor: DocxEditorController = props.importError
    ? { ...editor, importError: props.importError }
    : editor;

  return React.createElement(DocxEditorViewer, {
    editor: resolvedEditor,
    mode: "read-only",
  });
}

describe("editor starter and import error state", () => {
  it("keeps the bundled starter model blank", () => {
    expect(JSON.stringify(defaultStarterModel)).not.toContain(
      "React DOCX WYSIWYG"
    );
    expect(JSON.stringify(defaultStarterModel)).not.toContain(
      "Import a .docx"
    );
    expect(defaultStarterModel.nodes).toEqual([
      {
        type: "paragraph",
        children: [{ type: "text", text: "" }],
      },
    ]);
  });

  it("renders an import error state over a blank document", () => {
    const html = renderToStaticMarkup(
      React.createElement(DefaultEditorViewer, {
        importError: new Error("Invalid DOCX archive"),
      })
    );

    expect(html).toContain('data-docx-import-error="true"');
    expect(html).toContain("Failed to load DOCX");
    expect(html).toContain("Invalid DOCX archive");
    expect(html).not.toContain("React DOCX WYSIWYG");
    expect(html).not.toContain("Import a .docx");
  });
});
