import {
  PDFArray,
  PDFDocument,
  PDFName,
  PDFPage,
  PDFRawStream,
  arrayAsString,
  decodePDFRawStream,
} from "pdf-lib";

const PDF_NUMBER_PATTERN = /^[+-]?(?:\d+\.?\d*|\.\d+)$/;
const PDF_OPERATORS = new Set([
  "q",
  "Q",
  "cm",
  "w",
  "J",
  "j",
  "M",
  "d",
  "ri",
  "i",
  "gs",
  "m",
  "l",
  "c",
  "v",
  "y",
  "h",
  "re",
  "S",
  "s",
  "f",
  "F",
  "f*",
  "B",
  "B*",
  "b",
  "b*",
  "n",
  "W",
  "W*",
  "CS",
  "cs",
  "SC",
  "SCN",
  "sc",
  "scn",
  "G",
  "g",
  "RG",
  "rg",
  "K",
  "k",
  "sh",
  "Do",
  "MP",
  "DP",
  "BMC",
  "BDC",
  "EMC",
  "BT",
  "ET",
  "Tc",
  "Tw",
  "Tz",
  "TL",
  "Tf",
  "Tr",
  "Ts",
  "Td",
  "TD",
  "Tm",
  "T*",
  "Tj",
  "TJ",
  "'",
  '"',
  "BX",
  "EX",
  "BI",
  "ID",
  "EI",
  "d0",
  "d1",
]);

const PATH_CONSTRUCTION_OPERATORS = new Set(["m", "l", "c", "v", "y", "h", "re"]);
const PATH_PAINTING_OPERATORS = new Set([
  "S",
  "s",
  "f",
  "F",
  "f*",
  "B",
  "B*",
  "b",
  "b*",
]);
const NON_PATH_PAINTING_OPERATORS = new Set([
  "sh",
  "Do",
  "Tj",
  "TJ",
  "'",
  '"',
  "BI",
  "d0",
  "d1",
]);

interface Point {
  x: number;
  y: number;
}

interface PdfToken {
  value: string;
  start: number;
  end: number;
}

interface PdfOperation {
  operator: string;
  operands: PdfToken[];
  start: number;
  operatorStart: number;
  end: number;
}

interface TextRange {
  start: number;
  end: number;
}

interface GraphicsState {
  fillIsWhite: boolean;
  usesDefaultCoordinates: boolean;
}

interface ContentTransform {
  content: string;
  removed: number;
}

function isWhitespace(character: string): boolean {
  return /[\u0000\u0009\u000a\u000c\u000d\u0020]/.test(character);
}

function isDelimiter(character: string): boolean {
  return "()<>[]{}/%".includes(character);
}

function readPdfTokens(content: string): PdfToken[] {
  const tokens: PdfToken[] = [];
  let index = 0;

  while (index < content.length) {
    const character = content[index];
    if (isWhitespace(character)) {
      index += 1;
      continue;
    }
    if (character === "%") {
      while (index < content.length && !"\r\n".includes(content[index])) index += 1;
      continue;
    }

    const start = index;
    if (character === "(") {
      let depth = 1;
      index += 1;
      while (index < content.length && depth > 0) {
        if (content[index] === "\\") {
          index += 2;
        } else {
          if (content[index] === "(") depth += 1;
          if (content[index] === ")") depth -= 1;
          index += 1;
        }
      }
    } else if (character === "<" && content[index + 1] !== "<") {
      index += 1;
      while (index < content.length && content[index] !== ">") index += 1;
      if (index < content.length) index += 1;
    } else if (
      (character === "<" && content[index + 1] === "<") ||
      (character === ">" && content[index + 1] === ">")
    ) {
      index += 2;
    } else if (character === "/") {
      index += 1;
      while (
        index < content.length &&
        !isWhitespace(content[index]) &&
        !isDelimiter(content[index])
      ) {
        index += 1;
      }
    } else if (isDelimiter(character)) {
      index += 1;
    } else {
      while (
        index < content.length &&
        !isWhitespace(content[index]) &&
        !isDelimiter(content[index])
      ) {
        index += 1;
      }
    }

    const value = content.slice(start, index);
    tokens.push({ value, start, end: index });
    if (value === "BI") break;
  }

  return tokens;
}

function readPdfOperations(content: string): PdfOperation[] {
  const operations: PdfOperation[] = [];
  let operands: PdfToken[] = [];

  for (const token of readPdfTokens(content)) {
    if (!PDF_OPERATORS.has(token.value)) {
      operands.push(token);
      continue;
    }

    operations.push({
      operator: token.value,
      operands,
      start: operands[0]?.start ?? token.start,
      operatorStart: token.start,
      end: token.end,
    });
    operands = [];

    // The bytes after BI contain an inline image and must not be tokenized as PDF operators.
    if (token.value === "BI") break;
  }

  return operations;
}

function numericOperands(operation: PdfOperation): number[] {
  return operation.operands
    .map((token) => token.value)
    .filter((value) => PDF_NUMBER_PATTERN.test(value))
    .map(Number);
}

function isWhiteColor(operator: string, components: number[]): boolean {
  if (operator === "g") return components.length === 1 && components[0] === 1;
  if (operator === "rg") {
    return components.length === 3 && components.every((value) => value === 1);
  }
  if (operator === "k") {
    return components.length === 4 && components.every((value) => value === 0);
  }
  return (
    (components.length === 1 && components[0] === 1) ||
    (components.length === 3 && components.every((value) => value === 1)) ||
    (components.length === 4 && components.every((value) => value === 0))
  );
}

function pathCoversPage(
  points: Point[],
  page: { x: number; y: number; width: number; height: number },
): boolean {
  if (points.length !== 4) return false;

  const tolerance = 0.5;
  const corners: Point[] = [
    { x: page.x, y: page.y },
    { x: page.x + page.width, y: page.y },
    { x: page.x + page.width, y: page.y + page.height },
    { x: page.x, y: page.y + page.height },
  ];
  return corners.every((corner) =>
    points.some(
      (point) =>
        Math.abs(point.x - corner.x) <= tolerance &&
        Math.abs(point.y - corner.y) <= tolerance,
    ),
  );
}

function removeRanges(content: string, ranges: TextRange[]): string {
  const sorted = [...ranges].sort((left, right) => left.start - right.start);
  let cursor = 0;
  let transformed = "";

  for (const range of sorted) {
    transformed += content.slice(cursor, range.start);
    transformed += "\n";
    cursor = Math.max(cursor, range.end);
  }

  return transformed + content.slice(cursor);
}

function removeLeadingPageSizedWhiteFills(
  content: string,
  page: { x: number; y: number; width: number; height: number },
): ContentTransform {
  const ranges: TextRange[] = [];
  const graphicsStack: GraphicsState[] = [];
  let graphics: GraphicsState = {
    fillIsWhite: false,
    usesDefaultCoordinates: true,
  };
  let pathPoints: Point[] = [];
  let pathRanges: TextRange[] = [];
  let pathIsSupported = true;
  let pathWillClip = false;
  let removed = 0;

  const resetPath = (): void => {
    pathPoints = [];
    pathRanges = [];
    pathIsSupported = true;
    pathWillClip = false;
  };

  for (const operation of readPdfOperations(content)) {
    const { operator } = operation;
    const operands = numericOperands(operation);

    if (operator === "q") {
      graphicsStack.push({ ...graphics });
      continue;
    }
    if (operator === "Q") {
      graphics = graphicsStack.pop() ?? graphics;
      continue;
    }
    if (operator === "cm") {
      const isIdentity =
        operands.length === 6 &&
        operands.every((value, index) => value === [1, 0, 0, 1, 0, 0][index]);
      graphics.usesDefaultCoordinates &&= isIdentity;
      continue;
    }
    if (["g", "rg", "k", "sc", "scn"].includes(operator)) {
      graphics.fillIsWhite = isWhiteColor(operator, operands);
      continue;
    }

    if (PATH_CONSTRUCTION_OPERATORS.has(operator)) {
      pathRanges.push({ start: operation.start, end: operation.end });
      if (operator === "m" || operator === "l") {
        if (operands.length !== 2) pathIsSupported = false;
        else pathPoints.push({ x: operands[0], y: operands[1] });
      } else if (operator === "re") {
        if (operands.length !== 4 || pathPoints.length > 0) {
          pathIsSupported = false;
        } else {
          const [x, y, width, height] = operands;
          pathPoints.push(
            { x, y },
            { x: x + width, y },
            { x: x + width, y: y + height },
            { x, y: y + height },
          );
        }
      } else if (operator !== "h") {
        pathIsSupported = false;
      }
      continue;
    }

    if (operator === "W" || operator === "W*") {
      pathWillClip = true;
      continue;
    }
    if (operator === "n") {
      resetPath();
      continue;
    }

    if (PATH_PAINTING_OPERATORS.has(operator)) {
      const isFill = operator === "f" || operator === "F" || operator === "f*";
      const isBackground =
        isFill &&
        graphics.fillIsWhite &&
        graphics.usesDefaultCoordinates &&
        pathIsSupported &&
        !pathWillClip &&
        pathCoversPage(pathPoints, page);

      if (!isBackground) break;

      ranges.push(...pathRanges, {
        start: operation.operatorStart,
        end: operation.end,
      });
      removed += 1;
      resetPath();
      continue;
    }

    if (NON_PATH_PAINTING_OPERATORS.has(operator)) break;
  }

  return { content: removeRanges(content, ranges), removed };
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

export function removeWhiteBackground(pdf: PDFDocument, page: PDFPage): void {
  const content = decodePageContents(page);
  if (content === undefined) return;

  const result = removeLeadingPageSizedWhiteFills(content, page.getMediaBox());
  if (result.removed === 0) return;

  const replacement = pdf.context.flateStream(result.content);
  page.node.set(PDFName.of("Contents"), pdf.context.register(replacement));
}
