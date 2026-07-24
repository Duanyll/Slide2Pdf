import { describe, expect, it, vi } from "vitest";

import { LocalPdfSaver } from "../scripts/local-save.cjs";

describe("LocalPdfSaver", () => {
  it("remembers a chosen path and overwrites it for the same slide", async () => {
    const choosePath = vi.fn().mockResolvedValue("/tmp/figure.pdf");
    const writeFile = vi.fn().mockResolvedValue(undefined);
    const saver = new LocalPdfSaver({ choosePath, writeFile });
    const firstPdf = new TextEncoder().encode("first");
    const secondPdf = new TextEncoder().encode("second");

    await saver.save({
      slideKey: "presentation:slide-1",
      suggestedName: "Deck_Slide1.pdf",
      forceNewPath: false,
      data: firstPdf,
    });
    await saver.save({
      slideKey: "presentation:slide-1",
      suggestedName: "Deck_Slide1.pdf",
      forceNewPath: false,
      data: secondPdf,
    });

    expect(choosePath).toHaveBeenCalledOnce();
    expect(writeFile).toHaveBeenNthCalledWith(
      1,
      "/tmp/figure.pdf",
      firstPdf,
    );
    expect(writeFile).toHaveBeenNthCalledWith(
      2,
      "/tmp/figure.pdf",
      secondPdf,
    );
  });
});
