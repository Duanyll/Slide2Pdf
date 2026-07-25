import type { NormalizedRect } from "../pdf/transformPresentationPdf";

export interface ShapeBounds {
  left: number;
  top: number;
  width: number;
  height: number;
  visible: boolean;
}

export interface SlideSize {
  width: number;
  height: number;
}

export function computeContentBounds(
  shapes: ShapeBounds[],
  slide: SlideSize,
): NormalizedRect {
  const visibleShapes = shapes.filter(
    (shape) =>
      shape.visible &&
      shape.left < slide.width &&
      shape.top < slide.height &&
      shape.left + shape.width > 0 &&
      shape.top + shape.height > 0,
  );

  if (visibleShapes.length === 0) {
    throw new Error(
      "当前页没有可裁切的可见对象。请选择“导出整页”，或添加可见对象后重试。",
    );
  }

  const left = Math.max(
    0,
    Math.min(...visibleShapes.map((shape) => shape.left)),
  );
  const top = Math.max(
    0,
    Math.min(...visibleShapes.map((shape) => shape.top)),
  );
  const right = Math.min(
    slide.width,
    Math.max(
      ...visibleShapes.map((shape) => shape.left + shape.width),
    ),
  );
  const bottom = Math.min(
    slide.height,
    Math.max(...visibleShapes.map((shape) => shape.top + shape.height)),
  );

  return {
    left: left / slide.width,
    top: top / slide.height,
    width: (right - left) / slide.width,
    height: (bottom - top) / slide.height,
  };
}
