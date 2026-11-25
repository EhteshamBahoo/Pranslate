import { isTranslatable, paragraphSpans, splitPadding } from "../extract/paragraphs";
import type { TranslationProvider } from "../translate/provider";
import { translateRobust } from "../translate/robust";

export async function getSlideCount(): Promise<number> {
  return PowerPoint.run(async (ctx) => {
    ctx.presentation.slides.load("items");
    await ctx.sync();
    return ctx.presentation.slides.items.length;
  });
}

/** Location of a piece of text: slide, chain of shape ids (outermost first; groups nest), optional table cell. */
interface Ref { slideId: string; path: string[]; cell?: [number, number]; }
interface Saved { ref: Ref; original: string; }

/** Originals per text location, kept in memory this session so "Restore original" works. Lost if the pane reloads. */
const saved = new Map<string, Saved>();
export const hasSavedOriginals = () => saved.size > 0;

/** Shape types that carry a text frame. Group and Table are handled separately; the rest are skipped and reported. */
const TEXT_TYPES = new Set(["GeometricShape", "TextBox", "Placeholder", "Callout", "Freeform", "Unsupported"]);

export interface Progress { slideIndex: number; slideCount: number; message: string; }
export interface Summary {
  slidesDone: number; textItems: number; paragraphsChanged: number; paragraphsFailed: number;
  skippedByType: Record<string, number>; warnings: string[];
}

interface Target { ref: Ref; text: string; tr?: PowerPoint.TextRange; cell?: PowerPoint.TableCell; }

const jkey = (r: Ref) => JSON.stringify(r);
const bump = (m: Record<string, number>, k: string) => { m[k] = (m[k] ?? 0) + 1; };

/** Walks shapes recursively (groups, tables) and gathers every text location on a slide. */
async function collect(
  ctx: PowerPoint.RequestContext, slideId: string, shapes: PowerPoint.Shape[], parentPath: string[],
  out: Target[], skipped: Record<string, number>,
): Promise<void> {
  const textShapes: PowerPoint.Shape[] = [];
  for (const shape of shapes) {
    const path = [...parentPath, shape.id];
    if (shape.type === "Group") {
      try {
        shape.group.shapes.load("items/id,items/name,items/type");
        await ctx.sync();
        await collect(ctx, slideId, shape.group.shapes.items, path, out, skipped);
      } catch { bump(skipped, "Group (unreadable)"); }
    } else if (shape.type === "Table") {
      try {
        const table = shape.getTable();
        table.load("rowCount,columnCount");
        await ctx.sync();
        const cells: { cell: PowerPoint.TableCell; r: number; c: number }[] = [];
        for (let r = 0; r < table.rowCount; r++) for (let c = 0; c < table.columnCount; c++) {
          const cell = table.getCellOrNullObject(r, c);
          cell.load("text");
          cells.push({ cell, r, c });
        }
        await ctx.sync();
        for (const { cell, r, c } of cells) {
          if (!cell.isNullObject && isTranslatable(cell.text)) out.push({ ref: { slideId, path, cell: [r, c] }, text: cell.text, cell });
        }
      } catch { bump(skipped, "Table (unreadable)"); }
    } else if (TEXT_TYPES.has(shape.type)) {
      textShapes.push(shape);
    } else if (shape.type !== "Image" && shape.type !== "Line") {
      bump(skipped, shape.type); // SmartArt, Chart, Media, ... not editable through this API
    }
  }

  // Batch first for speed; if one shape breaks the batch, retry shapes one by one.
  const attempt = async (list: PowerPoint.Shape[]) => {
    const loaded = list.map((s) => { const tr = s.textFrame.textRange; tr.load("text"); return { s, tr }; });
    await ctx.sync();
    return loaded;
  };
  const push = (s: PowerPoint.Shape, tr: PowerPoint.TextRange) => {
    if (isTranslatable(tr.text)) out.push({ ref: { slideId, path: [...parentPath, s.id] }, text: tr.text, tr });
  };
  try {
    for (const { s, tr } of await attempt(textShapes)) push(s, tr);
  } catch {
    for (const s of textShapes) {
      try { for (const l of await attempt([s])) push(l.s, l.tr); }
      catch { bump(skipped, `${s.type} (no text frame)`); }
    }
  }
}

function rebuild(text: string, replacements: (string | null)[]): string {
  const spans = paragraphSpans(text);
  let out = "";
  let cursor = 0;
  spans.forEach((sp, i) => {
    out += text.slice(cursor, sp.start) + (replacements[i] ?? text.substr(sp.start, sp.length));
    cursor = sp.start + sp.length;
  });
  return out + text.slice(cursor);
}

/**
 * Shapes: replace each paragraph's text via getSubstring(...).text in reverse order so offsets stay valid.
 * Formatting is whatever PowerPoint keeps for the replaced range; mixed formatting INSIDE one paragraph may be
 * flattened (spike S1 measures this). Table cells: whole cell text is rewritten (cell.text), so inline
 * formatting inside a cell is likely flattened to the cell's first style (unverified).
 */
async function replaceParagraphs(ctx: PowerPoint.RequestContext, tr: PowerPoint.TextRange, text: string, reps: (string | null)[]) {
  const spans = paragraphSpans(text);
  for (let i = spans.length - 1; i >= 0; i--) {
    if (reps[i] == null) continue;
    tr.getSubstring(spans[i]!.start, spans[i]!.length).text = reps[i]!;
  }
  await ctx.sync();
}

export async function translateSlides(
  slideNumbers: number[], provider: TranslationProvider, sourceLang: string, targetLang: string,
  onProgress: (p: Progress) => void, signal?: AbortSignal,
): Promise<Summary> {
  const summary: Summary = { slidesDone: 0, textItems: 0, paragraphsChanged: 0, paragraphsFailed: 0, skippedByType: {}, warnings: [] };

  for (let n = 0; n < slideNumbers.length; n++) {
    if (signal?.aborted) { summary.warnings.push("Cancelled."); break; }
    const slideNo = slideNumbers[n]!;
    const say = (message: string) => onProgress({ slideIndex: n, slideCount: slideNumbers.length, message });
    say(`Slide ${slideNo}: reading text`);

    await PowerPoint.run(async (ctx) => {
      ctx.presentation.slides.load("items");
      await ctx.sync();
      const slide = ctx.presentation.slides.items[slideNo - 1];
      if (!slide) { summary.warnings.push(`Slide ${slideNo} not found.`); return; }
      slide.load("id");
      slide.shapes.load("items/id,items/name,items/type");
      await ctx.sync();

      const targets: Target[] = [];
      await collect(ctx, slide.id, slide.shapes.items, [], targets, summary.skippedByType);

      type Item = { ti: number; pi: number; lead: string; trail: string; core: string };
      const items: Item[] = [];
      const spansOf = targets.map((t) => paragraphSpans(t.text));
      targets.forEach((t, ti) => spansOf[ti]!.forEach((sp, pi) => {
        const { lead, core, trail } = splitPadding(t.text.substr(sp.start, sp.length));
        if (isTranslatable(core)) items.push({ ti, pi, lead, trail, core });
      }));
      if (items.length === 0) return;

      say(`Slide ${slideNo}: translating ${items.length} paragraphs`);
      const res = await translateRobust(provider, { sourceLang, targetLang }, items.map((i) => i.core), signal);
      summary.paragraphsFailed += res.failed;

      const byTarget = new Map<number, (string | null)[]>();
      items.forEach((it, k) => {
        const arr = byTarget.get(it.ti) ?? new Array(spansOf[it.ti]!.length).fill(null);
        arr[it.pi] = it.lead + res.segments[k]! + it.trail;
        byTarget.set(it.ti, arr);
      });

      for (const [ti, reps] of byTarget) {
        const t = targets[ti]!;
        const k = jkey(t.ref);
        if (!saved.has(k)) saved.set(k, { ref: t.ref, original: t.text }); // keep the FIRST original
        if (t.tr) await replaceParagraphs(ctx, t.tr, t.text, reps);
        else if (t.cell) { t.cell.text = rebuild(t.text, reps); await ctx.sync(); }
        summary.textItems++;
        summary.paragraphsChanged += reps.filter((r) => r != null).length;
      }
    });
    summary.slidesDone++;
  }
  return summary;
}

function resolveShape(ctx: PowerPoint.RequestContext, ref: Ref): PowerPoint.Shape {
  let shape = ctx.presentation.slides.getItem(ref.slideId).shapes.getItem(ref.path[0]!);
  for (const id of ref.path.slice(1)) shape = shape.group.shapes.getItem(id);
  return shape;
}

export async function restoreOriginals(): Promise<{ restored: number; warnings: string[] }> {
  const warnings: string[] = [];
  let restored = 0;
  await PowerPoint.run(async (ctx) => {
    for (const [key, s] of saved) {
      try {
        const shape = resolveShape(ctx, s.ref);
        if (s.ref.cell) {
          shape.getTable().getCellOrNullObject(s.ref.cell[0], s.ref.cell[1]).text = s.original;
          await ctx.sync();
        } else {
          const tr = shape.textFrame.textRange;
          tr.load("text");
          await ctx.sync();
          const cur = paragraphSpans(tr.text);
          const orig = paragraphSpans(s.original);
          if (cur.length !== orig.length) { warnings.push("A text box was edited after translation (paragraph count changed); skipped."); continue; }
          await replaceParagraphs(ctx, tr, tr.text, orig.map((sp) => s.original.substr(sp.start, sp.length)));
        }
        saved.delete(key);
        restored++;
      } catch {
        warnings.push("A shape or slide no longer exists; skipped.");
      }
    }
  });
  return { restored, warnings };
}
