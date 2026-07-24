import { PDFDocument } from "pdf-lib";
import { describe, expect, it } from "vitest";

import { transformPresentationPdf } from "../src/pdf/transformPresentationPdf";

async function createPresentationPdf(): Promise<Uint8Array> {
  const pdf = await PDFDocument.create();
  pdf.addPage([320, 180]);
  pdf.addPage([640, 360]);
  pdf.addPage([800, 450]);
  return pdf.save();
}

describe("transformPresentationPdf", () => {
  it("extracts the requested slide as a one-page PDF", async () => {
    const source = await createPresentationPdf();

    const result = await transformPresentationPdf(source, 1);

    const output = await PDFDocument.load(result);
    expect(output.getPageCount()).toBe(1);
    expect(output.getPage(0).getSize()).toEqual({ width: 640, height: 360 });
  });

  it("crops the extracted page to normalized PowerPoint content bounds", async () => {
    const source = await createPresentationPdf();

    const result = await transformPresentationPdf(source, 1, {
      left: 0.1,
      top: 0.2,
      width: 0.5,
      height: 0.4,
    });

    const output = await PDFDocument.load(result);
    const page = output.getPage(0);
    expect(page.getCropBox()).toEqual({
      x: 64,
      y: 144,
      width: 320,
      height: 144,
    });
    expect(page.getTrimBox()).toEqual(page.getCropBox());
  });

  it("rejects a slide index that is not present in the generated PDF", async () => {
    const source = await createPresentationPdf();

    await expect(transformPresentationPdf(source, 3)).rejects.toThrowError(
      "Slide 4 is not present in the generated PDF.",
    );
  });
});
