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

    const fileName = downloadPdf(new Uint8Array([1, 2, 3]), "slide.pdf");

    expect(createObjectURL).toHaveBeenCalledOnce();
    const blob = createObjectURL.mock.calls[0][0] as Blob;
    expect(blob.type).toBe("application/pdf");
    expect(anchor).toMatchObject({
      download: "slide_1.pdf",
      hidden: true,
      href: "blob:slide2pdf",
    });
    expect(fileName).toBe("slide_1.pdf");
    expect(append).toHaveBeenCalledWith(anchor);
    expect(click).toHaveBeenCalledOnce();
    expect(remove).toHaveBeenCalledOnce();
    expect(revokeObjectURL).toHaveBeenCalledWith("blob:slide2pdf");
  });

  it("increments and persists the file name for repeated downloads", () => {
    const anchors: Array<{ download: string }> = [];
    const values = new Map<string, string>();

    vi.stubGlobal("document", {
      body: { append: vi.fn() },
      createElement: vi.fn(() => {
        const anchor = {
          click: vi.fn(),
          download: "",
          hidden: false,
          href: "",
          remove: vi.fn(),
        };
        anchors.push(anchor);
        return anchor;
      }),
    });
    vi.stubGlobal("URL", {
      createObjectURL: vi.fn(() => "blob:slide2pdf"),
      revokeObjectURL: vi.fn(),
    });
    vi.stubGlobal("window", {
      localStorage: {
        getItem: (key: string) => values.get(key) ?? null,
        setItem: (key: string, value: string) => values.set(key, value),
      },
      setTimeout: vi.fn(),
    });

    expect(downloadPdf(new Uint8Array([1]), "deck_Slide3.pdf")).toBe(
      "deck_Slide3_1.pdf",
    );
    expect(downloadPdf(new Uint8Array([2]), "deck_Slide3.pdf")).toBe(
      "deck_Slide3_2.pdf",
    );
    expect(anchors.map((anchor) => anchor.download)).toEqual([
      "deck_Slide3_1.pdf",
      "deck_Slide3_2.pdf",
    ]);
  });

  it("keeps incrementing in memory when local storage is unavailable", () => {
    const anchors: Array<{ download: string }> = [];

    vi.stubGlobal("document", {
      body: { append: vi.fn() },
      createElement: vi.fn(() => {
        const anchor = {
          click: vi.fn(),
          download: "",
          hidden: false,
          href: "",
          remove: vi.fn(),
        };
        anchors.push(anchor);
        return anchor;
      }),
    });
    vi.stubGlobal("URL", {
      createObjectURL: vi.fn(() => "blob:slide2pdf"),
      revokeObjectURL: vi.fn(),
    });
    vi.stubGlobal("window", {
      get localStorage(): Storage {
        throw new DOMException("Storage is unavailable.", "SecurityError");
      },
      setTimeout: vi.fn(),
    });

    downloadPdf(new Uint8Array([1]), "fallback.pdf");
    downloadPdf(new Uint8Array([2]), "fallback.pdf");

    expect(anchors.map((anchor) => anchor.download)).toEqual([
      "fallback_1.pdf",
      "fallback_2.pdf",
    ]);
  });
});
