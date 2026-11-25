/**
 * Spike S1 harness: read shape text and font properties from the active slide.
 * Throwaway-quality; real extraction lives in src/extract/ after Phase 0.
 */
export interface ShapeTextInfo {
  id: string;
  name: string;
  text: string;
  fontName: string | null;
  fontSize: number | null;
  bold: boolean | null;
}

export async function readActiveSlideText(): Promise<ShapeTextInfo[]> {
  return PowerPoint.run(async (ctx) => {
    const slides = ctx.presentation.getSelectedSlides();
    slides.load("items");
    await ctx.sync();
    const slide = slides.items[0];
    if (!slide) return [];
    slide.shapes.load("items/id,items/name,items/type");
    await ctx.sync();

    const out: ShapeTextInfo[] = [];
    const ranges: { shape: PowerPoint.Shape; range: PowerPoint.TextRange }[] = [];
    for (const shape of slide.shapes.items) {
      const range = shape.textFrame.textRange;
      range.load("text,font/name,font/size,font/bold");
      ranges.push({ shape, range });
    }
    try {
      await ctx.sync();
    } catch {
      // Shapes without a text frame (images, etc.) throw on load; handled per-shape in the real extractor.
      return out;
    }
    for (const { shape, range } of ranges) {
      out.push({
        id: shape.id,
        name: shape.name,
        text: range.text,
        fontName: range.font.name ?? null,
        fontSize: range.font.size ?? null,
        bold: range.font.bold ?? null,
      });
    }
    return out;
  });
}
