import { transformPresentationPdf } from "../pdf/transformPresentationPdf";
import { getCurrentSlide } from "../powerpoint/getCurrentSlide";
import { getPresentationPdf } from "../powerpoint/getPresentationPdf";
import { savePdf, type SavePdfResult } from "../save/savePdf";

export type ExportMode = "slide" | "content";
export type ExportProgress = "reading-slide" | "creating-pdf" | "processing-pdf" | "saving";

export async function exportCurrentSlide(
  mode: ExportMode,
  forceNewPath: boolean,
  onProgress?: (progress: ExportProgress) => void,
): Promise<SavePdfResult> {
  onProgress?.("reading-slide");
  const slide = await getCurrentSlide(mode === "content");
  onProgress?.("creating-pdf");
  const presentationPdf = await getPresentationPdf();
  onProgress?.("processing-pdf");
  const output = await transformPresentationPdf(
    presentationPdf,
    slide.slideIndex,
    slide.contentBounds,
  );

  const title = sanitizeFileName(slide.presentationTitle || "Presentation");
  onProgress?.("saving");
  return savePdf(
    output,
    `${slide.presentationId}:${slide.slideId}`,
    `${title}_Slide${slide.slideIndex + 1}.pdf`,
    forceNewPath,
  );
}

function sanitizeFileName(fileName: string): string {
  return fileName.replace(/[\\/:*?"<>|]/g, "_").replace(/\.(pptx?|pptm)$/i, "");
}
