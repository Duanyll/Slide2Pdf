import {
  PDFArray,
  PDFDocument,
  PDFName,
  PDFRawStream,
  arrayAsString,
  decodePDFRawStream,
} from "pdf-lib";
import { describe, expect, it } from "vitest";

import { transformPresentationPdf } from "../src/pdf/transformPresentationPdf";

async function createPresentationPdf(): Promise<Uint8Array> {
  const pdf = await PDFDocument.create();
  pdf.addPage([320, 180]);
  pdf.addPage([640, 360]);
  pdf.addPage([800, 450]);
  return pdf.save();
}

async function createPdfWithContent(content: string): Promise<Uint8Array> {
  const pdf = await PDFDocument.create();
  const page = pdf.addPage([320, 180]);
  page.node.set(
    PDFName.of("Contents"),
    pdf.context.register(pdf.context.flateStream(content)),
  );
  return pdf.save();
}

async function createPowerPointStylePdf(): Promise<Uint8Array> {
  return createPdfWithContent(
    [
      "q",
      "/Cs1 cs",
      "1 1 1 sc",
      "0 180 m",
      "320 180 l",
      "320 0 l",
      "0 0 l",
      "h",
      "f",
      "0.2 0.7 0.3 sc",
      "40 140 m",
      "120 140 l",
      "120 60 l",
      "40 60 l",
      "h",
      "f",
      "Q",
    ].join("\n"),
  );
}

function getPageContent(pdf: PDFDocument): string {
  const contents = pdf.getPage(0).node.Contents();
  const streams =
    contents instanceof PDFArray
      ? Array.from({ length: contents.size() }, (_, index) =>
          contents.lookup(index, PDFRawStream),
        )
      : [contents];

  return streams
    .map((stream) => {
      if (!(stream instanceof PDFRawStream)) {
        throw new Error("Expected a raw PDF page content stream.");
      }
      return arrayAsString(decodePDFRawStream(stream).decode());
    })
    .join("\n");
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
      crop: {
        left: 0.1,
        top: 0.2,
        width: 0.5,
        height: 0.4,
      },
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

  it("removes the page-sized white fill when transparent background is requested", async () => {
    const source = await createPowerPointStylePdf();

    const result = await transformPresentationPdf(source, 0, {
      transparentBackground: true,
    });

    const content = getPageContent(await PDFDocument.load(result));
    expect(content).not.toContain("0 180 m\n320 180 l\n320 0 l\n0 0 l");
    expect(content).toContain("40 140 m\n120 140 l\n120 60 l\n40 60 l");
  });

  it("removes only leading white page fills expressed with common PDF operators", async () => {
    const source = await createPdfWithContent(
      [
        "1 1 1 rg",
        "0 0 320 180 re",
        "f",
        "0.2 0.7 0.3 rg",
        "40 60 80 80 re",
        "f",
        "1 1 1 sc",
        "0 180 m",
        "320 180 l",
        "320 0 l",
        "0 0 l",
        "h",
        "f",
      ].join("\n"),
    );

    const result = await transformPresentationPdf(source, 0, {
      transparentBackground: true,
    });

    const content = getPageContent(await PDFDocument.load(result));
    expect(content).not.toContain("0 0 320 180 re");
    expect(content).toContain("0 180 m\n320 180 l\n320 0 l\n0 0 l");
  });

  it("rejects a slide index that is not present in the generated PDF", async () => {
    const source = await createPresentationPdf();

    await expect(transformPresentationPdf(source, 3)).rejects.toThrowError(
      "Slide 4 is not present in the generated PDF.",
    );
  });
});
