import { transformPresentationPdf } from "../pdf/transformPresentationPdf";
import { getCurrentSlide } from "../powerpoint/getCurrentSlide";
import { getPresentationFileUrl } from "../powerpoint/getPresentationFileUrl";
import { getPresentationPdf } from "../powerpoint/getPresentationPdf";
import { downloadPdf } from "../save/downloadPdf";
import { getPresentationBaseName } from "./presentationFileName";

export type ExportMode = "slide" | "content";
export type ExportProgress = "reading-slide" | "creating-pdf" | "processing-pdf" | "saving";

export interface ExportOptions {
  transparentBackground?: boolean;
}

export interface CurrentSlidePdf {
  data: Uint8Array;
  slideId: string;
  suggestedFileName: string;
}

export async function exportCurrentSlide(
  mode: ExportMode,
  options: ExportOptions = {},
  onProgress?: (progress: ExportProgress) => void,
): Promise<string> {
  const output = await createCurrentSlidePdf(mode, options, onProgress);
  onProgress?.("saving");
  return downloadPdf(output.data, output.suggestedFileName);
}

export async function createCurrentSlidePdf(
  mode: ExportMode,
  options: ExportOptions = {},
  onProgress?: (progress: ExportProgress) => void,
): Promise<CurrentSlidePdf> {
  onProgress?.("reading-slide");
  const slide = await getCurrentSlide(mode === "content");
  const documentUrl = await getPresentationFileUrl();
  onProgress?.("creating-pdf");
  const presentationPdf = await getPresentationPdf();
  onProgress?.("processing-pdf");
  const output = await transformPresentationPdf(
    presentationPdf,
    slide.slideIndex,
    {
      crop: slide.contentBounds,
      transparentBackground: options.transparentBackground,
    },
  );

  const title = getPresentationBaseName(slide.presentationTitle, documentUrl);
  const fileName = `${title}_Slide${slide.slideIndex + 1}.pdf`;
  return {
    data: output,
    slideId: slide.slideId,
    suggestedFileName: fileName,
  };
}
