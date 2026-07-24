import { describe, expect, it } from "vitest";

import { getPresentationBaseName } from "../src/export/presentationFileName";

describe("getPresentationBaseName", () => {
  it("prefers the document file name over PowerPoint's title property", () => {
    expect(
      getPresentationBaseName(
        "Title property in win32",
        "file:///Users/duanyll/Documents/RouteEdit1.pptx",
      ),
    ).toBe("RouteEdit1");
  });

  it("decodes a cloud URL and ignores its query string", () => {
    expect(
      getPresentationBaseName(
        "Metadata title",
        "https://example.sharepoint.com/Documents/My%20Deck.pptm?download=1",
      ),
    ).toBe("My Deck");
  });

  it("falls back to the title when the document has not been saved", () => {
    expect(getPresentationBaseName("Quarterly: Review.pptx", "")).toBe(
      "Quarterly_ Review",
    );
  });

  it("uses a stable fallback when neither source has a name", () => {
    expect(getPresentationBaseName("", "")).toBe("Presentation");
  });
});
