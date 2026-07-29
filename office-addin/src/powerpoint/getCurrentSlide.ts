import type { NormalizedRect } from "../pdf/transformPresentationPdf";
import { computeContentBounds } from "./contentBounds";

export interface CurrentSlide {
  presentationTitle: string;
  slideId: string;
  slideIndex: number;
  contentBounds?: NormalizedRect;
}

export async function getCurrentSlide(
  includeContentBounds: boolean,
): Promise<CurrentSlide> {
  return PowerPoint.run(async (context) => {
    const presentation = context.presentation;
    const selectedSlides = presentation.getSelectedSlides();

    presentation.load("title");
    selectedSlides.load("items/id,items/index");

    if (includeContentBounds) {
      presentation.pageSetup.load("slideWidth,slideHeight");
    }

    await context.sync();

    const slide = selectedSlides.items[0];
    if (!slide) {
      throw new Error("没有找到当前幻灯片。请先选择一页，再重新导出。");
    }

    let contentBounds: NormalizedRect | undefined;
    if (includeContentBounds) {
      slide.shapes.load("items/left,items/top,items/width,items/height,items/visible");
      await context.sync();

      contentBounds = computeContentBounds(
        slide.shapes.items.map((shape) => ({
          left: shape.left,
          top: shape.top,
          width: shape.width,
          height: shape.height,
          visible: shape.visible,
        })),
        {
          width: presentation.pageSetup.slideWidth,
          height: presentation.pageSetup.slideHeight,
        },
      );
    }

    return {
      presentationTitle: presentation.title,
      slideId: slide.id,
      slideIndex: slide.index,
      contentBounds,
    };
  });
}
