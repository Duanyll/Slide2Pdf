import { transformPresentationPdf } from "../pdf/transformPresentationPdf";
import { getCurrentSlide } from "../powerpoint/getCurrentSlide";
import { getPresentationFileUrl } from "../powerpoint/getPresentationFileUrl";
import { getPresentationPdf } from "../powerpoint/getPresentationPdf";
import { downloadPdf } from "../save/downloadPdf";
import { getPresentationBaseName } from "./presentationFileName";

export type ExportMode = "slide" | "content";
export type ExportProgress = "reading-slide" | "creating-pdf" | "processing-pdf" | "saving";

export async function exportCurrentSlide(
  mode: ExportMode,
  onProgress?: (progress: ExportProgress) => void,
): Promise<string> {
  onProgress?.("reading-slide");
  const slide = await getCurrentSlide(mode === "content");
  const documentUrl = await getPresentationFileUrl();
  onProgress?.("creating-pdf");
  const presentationPdf = await getPresentationPdf();
  onProgress?.("processing-pdf");
  const output = await transformPresentationPdf(
    presentationPdf,
    slide.slideIndex,
    slide.contentBounds,
  );

  const title = getPresentationBaseName(slide.presentationTitle, documentUrl);
  const fileName = `${title}_Slide${slide.slideIndex + 1}.pdf`;
  onProgress?.("saving");
  downloadPdf(output, fileName);
  return fileName;
}
