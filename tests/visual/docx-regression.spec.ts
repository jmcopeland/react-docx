import { expect, test } from "@playwright/test";

import { listVisualDocxRegressionCasesSync } from "../helpers/docx-regression";

const cases = listVisualDocxRegressionCasesSync();

if (cases.length === 0) {
  test("requires a populated DOCX fidelity corpus", () => {
    throw new Error(
      "No DOCX fidelity cases are available. Fetch or mount the private corpus under tests/fixtures/docx-regression before running test:docx-regression."
    );
  });
}

for (const fixture of cases) {
  test(`matches Word layout invariants for ${fixture.relativePath}`, async ({
    page,
  }) => {
    test.setTimeout(120_000);
    await page.goto("/");

    await page.locator('input[type="file"][accept*=".doc"]').setInputFiles(
      fixture.absolutePath
    );
    await expect(page.getByText(`Loaded ${fixture.fileName}`).first()).toBeVisible({
      timeout: 120_000,
    });
    await page.evaluate(async () => {
      await document.fonts.ready;
    });

    const pages = page.locator("[data-docx-page-index]");
    if (fixture.wordPageCount !== undefined) {
      await expect
        .poll(() => pages.count(), { timeout: 20_000 })
        .toBe(fixture.wordPageCount);
    } else {
      await expect.poll(() => pages.count()).toBeGreaterThan(0);
    }

    const footerOverlapPages = await page.evaluate(() => {
      return Array.from(
        document.querySelectorAll<HTMLElement>("[data-docx-page-index]")
      ).filter((pageElement) => {
        const footer = pageElement.querySelector<HTMLElement>(
          '[data-docx-header-footer-region="footer"]'
        );
        if (!footer) {
          return false;
        }
        const footerTop = footer.getBoundingClientRect().top;
        const bodyParagraphs = Array.from(
          pageElement.querySelectorAll<HTMLElement>(
            '[data-docx-paragraph-host="true"]'
          )
        ).filter(
          (paragraph) =>
            !paragraph.closest('[data-docx-header-footer-region="footer"]')
        );
        const bodyBottom = bodyParagraphs.reduce(
          (maximum, paragraph) =>
            Math.max(maximum, paragraph.getBoundingClientRect().bottom),
          Number.NEGATIVE_INFINITY
        );
        return Number.isFinite(bodyBottom) && bodyBottom > footerTop;
      }).length;
    });

    expect(footerOverlapPages).toBe(0);
  });
}
