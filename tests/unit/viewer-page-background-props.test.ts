import * as React from "react";
import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { cloneDocModel } from "../../packages/doc-model/src";
import {
  defaultStarterModel,
  DocxEditorViewer,
  useDocxEditor,
} from "../../packages/react-viewer/src/editor";

function BackgroundViewer(props: {
  initialDocumentTheme?: "light" | "dark";
  pageBackgroundColor?: string;
  pageGapBackgroundColor?: string;
}): React.JSX.Element {
  const editor = useDocxEditor({
    starterModel: cloneDocModel(defaultStarterModel),
    initialDocumentTheme: props.initialDocumentTheme,
  });

  return React.createElement(DocxEditorViewer, {
    editor,
    mode: "read-only",
    pageBackgroundColor: props.pageBackgroundColor,
    pageGapBackgroundColor: props.pageGapBackgroundColor,
  });
}

describe("viewer page background props", () => {
  it("uses Tailwind neutral 950 for dark-theme page surfaces by default", () => {
    const html = renderToStaticMarkup(
      React.createElement(BackgroundViewer, {
        initialDocumentTheme: "dark",
      })
    );

    expect(html).toContain('data-docx-page-surface="true"');
    expect(html).toContain("background-color:#0a0a0a");
  });

  it("lets consumers override the page surface background color", () => {
    const html = renderToStaticMarkup(
      React.createElement(BackgroundViewer, {
        initialDocumentTheme: "dark",
        pageBackgroundColor: "#123456",
      })
    );

    expect(html).toContain("background-color:#123456");
  });

  it("keeps page gaps transparent by default and supports a custom gap color", () => {
    const defaultHtml = renderToStaticMarkup(
      React.createElement(BackgroundViewer, {})
    );
    const customHtml = renderToStaticMarkup(
      React.createElement(BackgroundViewer, {
        pageGapBackgroundColor: "#fedcba",
      })
    );

    expect(defaultHtml).toContain('data-testid="docx-editor-viewer"');
    expect(defaultHtml).toContain("background-color:transparent");
    expect(customHtml).toContain("background-color:#fedcba");
  });
});
