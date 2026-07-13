import { describe, expect, it } from "vitest";
import {
  createMinimalDocxPackage,
  packageToArrayBuffer,
  parseDocx,
} from "@extend-ai/react-docx-ooxml-core";
import { buildDocModel } from "@extend-ai/react-docx-doc-model";
import {
  applyRunStyle,
  toggleRunStyleFlag,
} from "@extend-ai/react-docx-editor-ops";
import { serializeDocModel } from "@extend-ai/react-docx-serializer";
import { textStyleWithExplicitFontFamily } from "../../packages/react-viewer/src/editor";
import { segmentTextByDocxScriptFont } from "../../packages/react-viewer/src/script-fonts";

const DOCUMENT_WITH_SCRIPT_FONTS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r>
        <w:rPr>
          <w:rFonts w:ascii="Aptos" w:hAnsi="Times New Roman" w:eastAsia="Yu Mincho" w:cs="Noto Naskh Arabic" w:asciiTheme="minorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="minorEastAsia" w:csTheme="minorBidi" w:hint="eastAsia"/>
          <w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>
          <w:rtl/>
          <w:cs w:val="0"/>
        </w:rPr>
        <w:t>Latin 日本語 العربية</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>`;

const DOCUMENT_WITH_THEME_SCRIPT_FONTS = DOCUMENT_WITH_SCRIPT_FONTS.replace(
  /<w:rFonts\b[^>]*\/>/,
  '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="minorEastAsia" w:csTheme="majorBidi"/>'
);

const THEME_WITH_SCRIPT_FONTS = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Script theme">
  <a:themeElements>
    <a:fontScheme name="Script fonts">
      <a:majorFont><a:latin typeface="Major Latin"/><a:ea typeface="Major East Asia"/><a:cs typeface="Major Complex"/></a:majorFont>
      <a:minorFont><a:latin typeface="Minor Latin"/><a:ea typeface="Minor East Asia"/><a:cs typeface="Minor Complex"/></a:minorFont>
    </a:fontScheme>
  </a:themeElements>
</a:theme>`;

function firstTextRun(model: Awaited<ReturnType<typeof buildDocModel>>) {
  const paragraph = model.nodes[0];
  expect(paragraph?.type).toBe("paragraph");
  if (paragraph?.type !== "paragraph") {
    throw new Error("expected a paragraph");
  }

  const run = paragraph.children[0];
  expect(run?.type).toBe("text");
  if (run?.type !== "text") {
    throw new Error("expected a text run");
  }
  return run;
}

describe("OOXML run font slots", () => {
  it("imports every script slot without changing the legacy fontFamily fallback", async () => {
    const seed = createMinimalDocxPackage(DOCUMENT_WITH_SCRIPT_FONTS);
    const pkg = await parseDocx(await packageToArrayBuffer(seed));
    const model = await buildDocModel(pkg);
    const style = firstTextRun(model).style;

    expect(style).toMatchObject({
      fontFamily: "Aptos",
      fontFamilyAscii: "Aptos",
      fontFamilyHAnsi: "Times New Roman",
      fontFamilyEastAsia: "Yu Mincho",
      fontFamilyCs: "Noto Naskh Arabic",
      fontThemeAscii: "minorHAnsi",
      fontThemeHAnsi: "majorHAnsi",
      fontThemeEastAsia: "minorEastAsia",
      fontThemeCs: "minorBidi",
      fontHint: "eastAsia",
      language: "en-US",
      languageEastAsia: "ja-JP",
      languageBidi: "ar-SA",
      rightToLeft: true,
      complexScript: false,
    });
  });

  it("resolves theme families independently for every script slot", async () => {
    const seed = createMinimalDocxPackage(DOCUMENT_WITH_THEME_SCRIPT_FONTS);
    seed.parts.set("word/theme/theme1.xml", {
      name: "word/theme/theme1.xml",
      content: THEME_WITH_SCRIPT_FONTS,
    });
    seed.parts.set("word/styles.xml", {
      name: "word/styles.xml",
      content:
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>',
    });
    const pkg = await parseDocx(await packageToArrayBuffer(seed));
    const model = await buildDocModel(pkg);
    const style = firstTextRun(model).style;

    expect(style).toMatchObject({
      fontFamily: "Minor Latin",
      resolvedFontFamilyAscii: "Minor Latin",
      resolvedFontFamilyHAnsi: "Major Latin",
      resolvedFontFamilyEastAsia: "Minor East Asia",
      resolvedFontFamilyCs: "Major Complex",
    });
    expect(
      segmentTextByDocxScriptFont("Aé日ع", style).map((segment) => [
        segment.text,
        segment.fontFamily,
      ])
    ).toEqual([
      ["A", "Minor Latin"],
      ["é", "Major Latin"],
      ["日", "Minor East Asia"],
      ["ع", "Major Complex"],
    ]);
  });

  it("keeps script slots and language metadata after an edit forces run regeneration", async () => {
    const seed = createMinimalDocxPackage(DOCUMENT_WITH_SCRIPT_FONTS);
    const pkg = await parseDocx(await packageToArrayBuffer(seed));
    const model = await buildDocModel(pkg);
    const edited = toggleRunStyleFlag(model, 0, 0, "bold");

    expect(firstTextRun(edited).style?.fontFamilyEastAsia).toBe("Yu Mincho");
    expect(firstTextRun(edited).style?.languageBidi).toBe("ar-SA");

    const serialized = await serializeDocModel(edited, pkg);
    const xml = serialized.parts.get("word/document.xml")?.content ?? "";
    expect(xml).toContain(
      '<w:rFonts w:ascii="Aptos" w:hAnsi="Times New Roman" w:eastAsia="Yu Mincho" w:cs="Noto Naskh Arabic" w:asciiTheme="minorHAnsi" w:hAnsiTheme="majorHAnsi" w:eastAsiaTheme="minorEastAsia" w:csTheme="minorBidi" w:hint="eastAsia"/>'
    );
    expect(xml).toContain(
      '<w:lang w:val="en-US" w:eastAsia="ja-JP" w:bidi="ar-SA"/>'
    );
    expect(xml).toContain("<w:rtl/>");
    expect(xml).toContain('<w:cs w:val="0"/>');

    const reparsed = await parseDocx(await packageToArrayBuffer(serialized));
    const reparsedModel = await buildDocModel(reparsed);
    expect(firstTextRun(reparsedModel).style).toMatchObject({
      fontFamilyAscii: "Aptos",
      fontFamilyHAnsi: "Times New Roman",
      fontFamilyEastAsia: "Yu Mincho",
      fontFamilyCs: "Noto Naskh Arabic",
      languageBidi: "ar-SA",
      rightToLeft: true,
      complexScript: false,
    });
  });

  it("keeps the legacy fontFamily serializer contract for callers without slot metadata", async () => {
    const seed = createMinimalDocxPackage(
      DOCUMENT_WITH_SCRIPT_FONTS.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, "")
    );
    const pkg = await parseDocx(await packageToArrayBuffer(seed));
    const model = await buildDocModel(pkg);
    const edited = applyRunStyle(model, 0, 0, { fontFamily: "Legacy Family" });

    const serialized = await serializeDocModel(edited, pkg);
    const xml = serialized.parts.get("word/document.xml")?.content ?? "";
    expect(xml).toContain(
      '<w:rFonts w:ascii="Legacy Family" w:hAnsi="Legacy Family" w:cs="Legacy Family"/>'
    );
  });

  it("treats a fontFamily edit as an override of imported script slots", async () => {
    const seed = createMinimalDocxPackage(DOCUMENT_WITH_SCRIPT_FONTS);
    const pkg = await parseDocx(await packageToArrayBuffer(seed));
    const model = await buildDocModel(pkg);
    const edited = applyRunStyle(model, 0, 0, {
      fontFamily: "Replacement Family",
    });

    const serialized = await serializeDocModel(edited, pkg);
    const xml = serialized.parts.get("word/document.xml")?.content ?? "";
    expect(xml).toContain(
      '<w:rFonts w:ascii="Replacement Family" w:hAnsi="Replacement Family" w:cs="Replacement Family"/>'
    );
    expect(xml).not.toContain('w:eastAsia="Yu Mincho"');
    expect(xml).not.toContain('w:asciiTheme="minorHAnsi"');
  });

  it("makes an explicit toolbar family authoritative without source provenance", async () => {
    const seed = createMinimalDocxPackage(
      DOCUMENT_WITH_SCRIPT_FONTS.replace(/<w:rPr>[\s\S]*?<\/w:rPr>/, "")
    );
    const pkg = await parseDocx(await packageToArrayBuffer(seed));
    const model = await buildDocModel(pkg);
    const explicitStyle = textStyleWithExplicitFontFamily(
      {
        fontThemeAscii: "minorHAnsi",
        fontThemeEastAsia: "minorEastAsia",
        fontFamilyEastAsia: "Yu Mincho",
        fontHint: "eastAsia",
      },
      "Replacement Family"
    );
    const edited = applyRunStyle(model, 0, 0, explicitStyle);

    expect(explicitStyle).toMatchObject({
      fontFamily: "Replacement Family",
    });
    expect(explicitStyle.fontThemeAscii).toBeUndefined();
    expect(explicitStyle.fontThemeEastAsia).toBeUndefined();
    expect(explicitStyle.fontFamilyEastAsia).toBeUndefined();
    expect(explicitStyle.fontHint).toBeUndefined();

    const serialized = await serializeDocModel(edited, pkg);
    const xml = serialized.parts.get("word/document.xml")?.content ?? "";
    expect(xml).toContain(
      '<w:rFonts w:ascii="Replacement Family" w:hAnsi="Replacement Family" w:cs="Replacement Family"/>'
    );
    expect(xml).not.toContain("minorEastAsia");
    expect(xml).not.toContain("Yu Mincho");
  });
});
