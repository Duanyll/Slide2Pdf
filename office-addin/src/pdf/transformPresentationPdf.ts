import { PDFDocument } from "pdf-lib";

export interface NormalizedRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

export async function transformPresentationPdf(
  presentationPdf: Uint8Array,
  slideIndex: number,
  crop?: NormalizedRect,
): Promise<Uint8Array> {
  const source = await PDFDocument.load(presentationPdf);
  if (
    !Number.isInteger(slideIndex) ||
    slideIndex < 0 ||
    slideIndex >= source.getPageCount()
  ) {
    throw new Error(
      `Slide ${slideIndex + 1} is not present in the generated PDF.`,
    );
  }

  const output = await PDFDocument.create();
  const [page] = await output.copyPages(source, [slideIndex]);

  output.addPage(page);

  if (crop) {
    const { width: pageWidth, height: pageHeight } = page.getSize();
    const x = crop.left * pageWidth;
    const y = (1 - crop.top - crop.height) * pageHeight;
    const width = crop.width * pageWidth;
    const height = crop.height * pageHeight;

    page.setCropBox(x, y, width, height);
    page.setTrimBox(x, y, width, height);
  }

  return output.save();
}
