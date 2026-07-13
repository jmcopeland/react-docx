import { describe, expect, it } from "vitest";

import { buildDocModel } from "@extend-ai/react-docx-doc-model";
import { updateTableCellParagraphText } from "@extend-ai/react-docx-editor-ops";
import {
  createMinimalDocxPackage,
  packageToArrayBuffer,
  parseDocx,
} from "@extend-ai/react-docx-ooxml-core";
import { serializeDocModel } from "@extend-ai/react-docx-serializer";

const DOCUMENT_WITH_HEADER_AND_VERTICAL_MERGE = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tblPr><w:tblStyle w:val="TableGrid"/><w:tblW w:w="0" w:type="auto"/><w:tblCaption w:val="Quarterly totals"/><w:tblDescription w:val="Revenue by region"/><w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0"/></w:tblPr>
      <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
      <w:tr>
        <w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>
        <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>Header</w:t></w:r></w:p></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Top</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vMerge/></w:tcPr><w:p/></w:tc>
        <w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>Bottom</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>`;

describe("edited table round-trip fidelity", () => {
  it("retains repeated header and vertical merge properties after regeneration", async () => {
    const seed = createMinimalDocxPackage(
      DOCUMENT_WITH_HEADER_AND_VERTICAL_MERGE
    );
    const pkg = await parseDocx(await packageToArrayBuffer(seed));
    const model = await buildDocModel(pkg);
    const edited = updateTableCellParagraphText(
      model,
      0,
      0,
      0,
      0,
      "Header edited"
    );
    const editedTable = edited.nodes[0];
    expect(editedTable?.type).toBe("table");
    if (editedTable?.type === "table") {
      expect(editedTable.sourceXml).toBeDefined();
      expect(editedTable.sourceTextPatches).toHaveLength(1);
    }

    const serialized = await serializeDocModel(edited, pkg);
    const xml = serialized.parts.get("word/document.xml")?.content ?? "";

    expect(xml).toContain("Header edited");
    expect(xml).toContain("<w:tblHeader/>");
    expect(xml).toContain("<w:cantSplit/>");
    expect(xml).toContain('<w:vMerge w:val="restart"/>');
    expect(xml).toContain("<w:vMerge/>");
    expect(xml).toContain('<w:tblStyle w:val="TableGrid"/>');
    expect(xml).toContain('<w:tblCaption w:val="Quarterly totals"/>');
    expect(xml).toContain('<w:tblDescription w:val="Revenue by region"/>');
    expect(xml).toContain(
      '<w:tblLook w:val="04A0" w:firstRow="1" w:lastRow="0"/>'
    );
  });
});
