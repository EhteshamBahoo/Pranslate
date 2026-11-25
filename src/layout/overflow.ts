/**
 * Pure overflow estimate (spike S3). Given a function that measures text width,
 * estimate wrapped line count and height. Approximate by design: missing fonts and
 * PowerPoint's own line breaking will cause differences; callers must treat the result
 * as a warning signal, never as a guarantee.
 */
export type MeasureWidth = (text: string) => number;

export interface BoxSpec {
  widthPt: number;
  heightPt: number;
  insetLeftPt: number;
  insetRightPt: number;
  insetTopPt: number;
  insetBottomPt: number;
  fontSizePt: number;
  lineSpacing: number; // multiple of font size, e.g. 1.2
}

export function estimateLines(text: string, availableWidth: number, measure: MeasureWidth): number {
  let lines = 0;
  for (const para of text.split(/\r?\n|\v/)) {
    const words = para.split(/(\s+)/).filter((w) => w.length > 0);
    let current = "";
    let count = 1;
    for (const w of words) {
      const candidate = current + w;
      if (measure(candidate.trimEnd()) <= availableWidth || current === "") {
        current = candidate;
      } else {
        count += 1;
        current = w.trimStart();
      }
    }
    lines += count;
  }
  return lines;
}

export function estimateOverflow(text: string, box: BoxSpec, measure: MeasureWidth) {
  const availableWidth = box.widthPt - box.insetLeftPt - box.insetRightPt;
  const availableHeight = box.heightPt - box.insetTopPt - box.insetBottomPt;
  const lines = estimateLines(text, availableWidth, measure);
  const neededHeight = lines * box.fontSizePt * box.lineSpacing;
  return { lines, neededHeight, availableHeight, overflows: neededHeight > availableHeight };
}
