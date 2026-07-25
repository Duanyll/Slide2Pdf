import { PDFDocument } from "pdf-lib";

import { removeWhiteBackground } from "./removeWhiteBackground";

export interface NormalizedRect {
  left: number;
  top: number;
  width: number;
  height: number;
}

export interface TransformPresentationPdfOptions {
  crop?: NormalizedRect;
  transparentBackground?: boolean;
}

export async function transformPresentationPdf(
  presentationPdf: Uint8Array,
  slideIndex: number,
  options: TransformPresentationPdfOptions = {},
): Promise<Uint8Array> {
  const source = await PDFDocument.load(presentationPdf);
  if (
    !Number.isInteger(slideIndex) ||
    slideIndex < 0 ||
    slideIndex >= source.getPageCount()
  ) {
    throw new Error(
      `生成的 PDF 中没有第 ${slideIndex + 1} 页。请确认当前幻灯片仍然存在，然后重试。`,
    );
  }

  const output = await PDFDocument.create();
  const [page] = await output.copyPages(source, [slideIndex]);

  output.addPage(page);

  if (options.transparentBackground) {
    removeWhiteBackground(output, page);
  }

  if (options.crop) {
    const crop = options.crop;
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
