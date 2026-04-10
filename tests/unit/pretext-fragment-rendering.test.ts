import { describe, expect, it } from "vitest";
import { sanitizeRenderedPretextFragmentText } from "../../packages/react-viewer/src/editor";

describe("pretext fragment rendering", () => {
  it("strips hard break characters from already-laid-out fragment slices", () => {
    expect(sanitizeRenderedPretextFragmentText("Alpha\n")).toBe("Alpha");
    expect(sanitizeRenderedPretextFragmentText("\r\nBeta")).toBe("Beta");
    expect(sanitizeRenderedPretextFragmentText("\n")).toBe("");
  });
});
