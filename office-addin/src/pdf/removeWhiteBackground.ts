import {
  PDFArray,
  PDFDocument,
  PDFName,
  PDFPage,
  PDFRawStream,
  arrayAsString,
  decodePDFRawStream,
} from "pdf-lib";

const NUMBER = "([-+]?(?:\\d*\\.\\d+|\\d+\\.?\\d*))";
const WHITE_PAGE_FILL = new RegExp(
  `${NUMBER}\\s+${NUMBER}\\s+${NUMBER}\\s+sc\\s+` +
    `${NUMBER}\\s+${NUMBER}\\s+m\\s+` +
    `${NUMBER}\\s+${NUMBER}\\s+l\\s+` +
    `${NUMBER}\\s+${NUMBER}\\s+l\\s+` +
    `${NUMBER}\\s+${NUMBER}\\s+l\\s+` +
    "h\\s+f\\*?",
  "g",
);

interface Point {
  x: number;
  y: number;
}

interface ContentTransform {
  content: string;
  removed: number;
}

function removePageSizedWhiteFills(
  content: string,
  width: number,
  height: number,
): ContentTransform {
  let removed = 0;
  const tolerance = 0.5;

  const transformed = content.replace(WHITE_PAGE_FILL, (...args: unknown[]) => {
    const match = String(args[0]);
    const values = args.slice(1, 12).map(Number);
    const [red, green, blue, ...coordinates] = values;
    if (
      Math.abs(red - 1) > Number.EPSILON ||
      Math.abs(green - 1) > Number.EPSILON ||
      Math.abs(blue - 1) > Number.EPSILON
    ) {
      return match;
    }

    const points: Point[] = Array.from({ length: 4 }, (_, index) => ({
      x: coordinates[index * 2],
      y: coordinates[index * 2 + 1],
    }));
    const corners: Point[] = [
      { x: 0, y: 0 },
      { x: width, y: 0 },
      { x: width, y: height },
      { x: 0, y: height },
    ];
    const coversPage = corners.every((corner) =>
      points.some(
        (point) =>
          Math.abs(point.x - corner.x) <= tolerance &&
          Math.abs(point.y - corner.y) <= tolerance,
      ),
    );

    if (!coversPage) return match;

    removed += 1;
    return `${red} ${green} ${blue} sc\n`;
  });

  return { content: transformed, removed };
}

function decodePageContents(page: PDFPage): string | undefined {
  const contents = page.node.Contents();
  if (!contents) return undefined;

  const streams =
    contents instanceof PDFArray
      ? Array.from({ length: contents.size() }, (_, index) =>
          contents.lookup(index),
        )
      : [contents];

  return streams
    .map((stream) => {
      if (!(stream instanceof PDFRawStream)) {
        throw new Error("PowerPoint PDF 使用了无法处理的页面内容格式。");
      }
      return arrayAsString(decodePDFRawStream(stream).decode());
    })
    .join("\n");
}

export function removeWhiteBackground(
  pdf: PDFDocument,
  page: PDFPage,
): number {
  const content = decodePageContents(page);
  if (content === undefined) return 0;

  const { width, height } = page.getMediaBox();
  const result = removePageSizedWhiteFills(content, width, height);
  if (result.removed === 0) return 0;

  const replacement = pdf.context.flateStream(result.content);
  page.node.set(PDFName.of("Contents"), pdf.context.register(replacement));
  return result.removed;
}
