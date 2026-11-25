import { isTranslatable, paragraphSpans, splitPadding } from "../extract/paragraphs";
import type { TranslationProvider } from "../translate/provider";

export async function getSlideCount(): Promise<number> {
  return PowerPoint.run(async (ctx) => {
    ctx.presentation.slides.load("items");
    await ctx.sync();
    return ctx.presentation.slides.items.length;
  });
}

/** Original text per shape, kept in memory for this session so "Restore original" works. Lost if the pane reloads. */
interface Saved { slideId: string; shapeId: string; original: string; translated: string; }
const saved = new Map<string, Saved>();
export const hasSavedOriginals = () => saved.size > 0;

const TEXT_SHAPE_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder"]);

export interface Progress { slideIndex: number; slideCount: number; message: string; }
export interface Summary { slidesDone: number; shapesChanged: number; paragraphsChanged: number; warnings: string[]; }

/**
 * Replaces each paragraph's text in place via getSubstring(...).text (reverse order so offsets stay valid).
 * Formatting is whatever PowerPoint keeps for the replaced range; mixed formatting INSIDE one paragraph may
 * be flattened. That is exactly what spike S1 measures; treat this as unverified until S1 results are in.
 */
async function replaceParagraphs(
  ctx: PowerPoint.RequestContext, tr: PowerPoint.TextRange, text: string, replacements: (string | null)[],
) {
  const spans = paragraphSpans(text);
  for (let i = spans.length - 1; i >= 0; i--) {
    const rep = replacements[i];
    if (rep == null) continue;
    tr.getSubstring(spans[i]!.start, spans[i]!.length).text = rep;
  }
  await ctx.sync();
}

export async function translateSlides(
  slideNumbers: number[], provider: TranslationProvider, sourceLang: string, targetLang: string,
  onProgress: (p: Progress) => void, signal?: AbortSignal,
): Promise<Summary> {
  const summary: Summary = { slidesDone: 0, shapesChanged: 0, paragraphsChanged: 0, warnings: [] };
  for (let n = 0; n < slideNumbers.length; n++) {
    if (signal?.aborted) { summary.warnings.push("Cancelled."); break; }
    const slideNo = slideNumbers[n]!;
    onProgress({ slideIndex: n, slideCount: slideNumbers.length, message: `Slide ${slideNo}: reading text` });

    await PowerPoint.run(async (ctx) => {
      ctx.presentation.slides.load("items");
      await ctx.sync();
      const slide = ctx.presentation.slides.items[slideNo - 1];
      if (!slide) { summary.warnings.push(`Slide ${slideNo} not found.`); return; }
      slide.load("id");
      slide.shapes.load("items/id,items/name,items/type");
      await ctx.sync();

      const candidates = slide.shapes.items.filter((s) => TEXT_SHAPE_TYPES.has(s.type));
      const loaded: { shape: PowerPoint.Shape; tr: PowerPoint.TextRange }[] = [];
      for (const shape of candidates) {
        const tr = shape.textFrame.textRange;
        tr.load("text");
        loaded.push({ shape, tr });
      }
      try { await ctx.sync(); }
      catch { summary.warnings.push(`Slide ${slideNo}: some shapes could not be read and were skipped.`); return; }

      // Collect paragraphs for the whole slide into one request (keeps slide context together, fewer calls).
      type Item = { li: number; pi: number; lead: string; trail: string; core: string };
      const items: Item[] = [];
      const perShape = loaded.map((l) => paragraphSpans(l.tr.text));
      loaded.forEach((l, li) => {
        perShape[li]!.forEach((sp, pi) => {
          const raw = l.tr.text.substr(sp.start, sp.length);
          const { lead, core, trail } = splitPadding(raw);
          if (isTranslatable(core)) items.push({ li, pi, lead, trail, core });
        });
      });
      if (items.length === 0) return;

      onProgress({ slideIndex: n, slideCount: slideNumbers.length, message: `Slide ${slideNo}: translating ${items.length} paragraphs` });
      const result = await provider.translate(
        { segments: items.map((i) => i.core), sourceLang, targetLang }, signal);

      const byShape = new Map<number, (string | null)[]>();
      items.forEach((it, k) => {
        const arr = byShape.get(it.li) ?? new Array(perShape[it.li]!.length).fill(null);
        arr[it.pi] = it.lead + result.segments[k]! + it.trail;
        byShape.set(it.li, arr);
      });

      for (const [li, reps] of byShape) {
        const { shape, tr } = loaded[li]!;
        const original = tr.text;
        await replaceParagraphs(ctx, tr, original, reps);
        tr.load("text");
        await ctx.sync();
        const key = `${slide.id}|${shape.id}`;
        // keep the FIRST original if translated more than once
        saved.set(key, { slideId: slide.id, shapeId: shape.id, original: saved.get(key)?.original ?? original, translated: tr.text });
        summary.shapesChanged++;
        summary.paragraphsChanged += reps.filter((r) => r != null).length;
      }
    });
    summary.slidesDone++;
  }
  return summary;
}

export async function restoreOriginals(): Promise<{ restored: number; warnings: string[] }> {
  const warnings: string[] = [];
  let restored = 0;
  await PowerPoint.run(async (ctx) => {
    for (const [key, s] of saved) {
      try {
        const slide = ctx.presentation.slides.getItem(s.slideId);
        const shape = slide.shapes.getItem(s.shapeId);
        const tr = shape.textFrame.textRange;
        tr.load("text");
        await ctx.sync();
        const curSpans = paragraphSpans(tr.text);
        const origSpans = paragraphSpans(s.original);
        if (curSpans.length !== origSpans.length) { warnings.push(`A text box was edited after translation (paragraph count changed); skipped.`); continue; }
        const reps = origSpans.map((sp) => s.original.substr(sp.start, sp.length));
        await replaceParagraphs(ctx, tr, tr.text, reps);
        saved.delete(key);
        restored++;
      } catch {
        warnings.push("A shape or slide no longer exists; skipped.");
      }
    }
  });
  return { restored, warnings };
}
