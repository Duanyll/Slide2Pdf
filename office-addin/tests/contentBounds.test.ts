import { describe, expect, it } from "vitest";

import { computeContentBounds } from "../src/powerpoint/contentBounds";

describe("computeContentBounds", () => {
  it("returns the union of visible shapes as normalized slide coordinates", () => {
    const result = computeContentBounds(
      [
        { left: 100, top: 50, width: 300, height: 100, visible: true },
        { left: 800, top: 200, width: 100, height: 250, visible: true },
        { left: 0, top: 0, width: 1000, height: 500, visible: false },
      ],
      { width: 1000, height: 500 },
    );

    expect(result).toEqual({
      left: 0.1,
      top: 0.1,
      width: 0.8,
      height: 0.8,
    });
  });

  it("clips content bounds to the slide", () => {
    const result = computeContentBounds(
      [
        { left: -100, top: -50, width: 300, height: 100, visible: true },
        { left: 900, top: 450, width: 200, height: 100, visible: true },
      ],
      { width: 1000, height: 500 },
    );

    expect(result).toEqual({ left: 0, top: 0, width: 1, height: 1 });
  });

  it("rejects slides without visible content", () => {
    expect(() =>
      computeContentBounds(
        [{ left: 10, top: 10, width: 20, height: 20, visible: false }],
        { width: 1000, height: 500 },
      ),
    ).toThrowError("No visible shapes were found on the current slide.");
  });
});
