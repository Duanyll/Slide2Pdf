import { afterEach, describe, expect, it, vi } from "vitest";

import { downloadPdf } from "../src/save/downloadPdf";

describe("downloadPdf", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("clicks a hidden PDF download and revokes its object URL", () => {
    const click = vi.fn();
    const remove = vi.fn();
    const anchor = {
      click,
      download: "",
      hidden: false,
      href: "",
      remove,
    };
    const append = vi.fn();
    const createObjectURL = vi.fn((_blob: Blob) => "blob:slide2pdf");
    const revokeObjectURL = vi.fn();

    vi.stubGlobal("document", {
      body: { append },
      createElement: vi.fn(() => anchor),
    });
    vi.stubGlobal("URL", { createObjectURL, revokeObjectURL });
    vi.stubGlobal("window", {
      setTimeout: (callback: () => void) => {
        callback();
        return 1;
      },
    });

    downloadPdf(new Uint8Array([1, 2, 3]), "slide.pdf");

    expect(createObjectURL).toHaveBeenCalledOnce();
    const blob = createObjectURL.mock.calls[0][0] as Blob;
    expect(blob.type).toBe("application/pdf");
    expect(anchor).toMatchObject({
      download: "slide.pdf",
      hidden: true,
      href: "blob:slide2pdf",
    });
    expect(append).toHaveBeenCalledWith(anchor);
    expect(click).toHaveBeenCalledOnce();
    expect(remove).toHaveBeenCalledOnce();
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:slide2pdf");
  });
});
