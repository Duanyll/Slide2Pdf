import { transformPresentationPdf } from "../pdf/transformPresentationPdf";
import { getCurrentSlide } from "../powerpoint/getCurrentSlide";
import { getPresentationPdf } from "../powerpoint/getPresentationPdf";
import { downloadPdf } from "../save/downloadPdf";

export type ExportMode = "slide" | "content";
export type ExportProgress = "reading-slide" | "creating-pdf" | "processing-pdf" | "saving";

export async function exportCurrentSlide(
  mode: ExportMode,
  onProgress?: (progress: ExportProgress) => void,
): Promise<string> {
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
  const fileName = `${title}_Slide${slide.slideIndex + 1}.pdf`;
  onProgress?.("saving");
  downloadPdf(output, fileName);
  return fileName;
}

function sanitizeFileName(fileName: string): string {
  return fileName.replace(/[\\/:*?"<>|]/g, "_").replace(/\.(pptx?|pptm)$/i, "");
}
