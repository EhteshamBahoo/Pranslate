/**
 * Spike S1: can we replace text while keeping run-level formatting?
 *
 * The harness builds its own fixture slide (so no .pptx needs to be shared), then applies
 * three replacement strategies to identical copies of each sample and compares per-character
 * formatting before/after.
 *
 *   a: whole-range  `textRange.text = translated`
 *   b: whole-range, then re-apply each run's formatting via getSubstring(...).font
 *   c: per-run      `getSubstring(runStart, runLen).text = translatedSegment` (reverse order)
 *
 * The "translations" are hard-coded German segments aligned 1:1 with the source runs. This
 * tests the Office.js mechanics only; it says nothing about how a model would produce
 * run-aligned output (that is spike S2).
 */

export interface Fmt {
  name: string;
  size: number;
  bold: boolean;
  italic: boolean;
  underline: string;
  color: string;
  superscript: boolean;
  subscript: boolean;
}

export interface RunSpec {
  text: string;
  translated: string;
  fmt: Partial<Fmt>;
}

export interface SampleSpec {
  id: string;
  runs: RunSpec[];
  bullets?: boolean;
}

export type Strategy = "a-whole" | "b-whole+reformat" | "c-per-run";
export const STRATEGIES: Strategy[] = ["a-whole", "b-whole+reformat", "c-per-run"];

const BASE: Fmt = {
  name: "Calibri", size: 18, bold: false, italic: false,
  underline: "None", color: "#000000", superscript: false, subscript: false,
};
const full = (f: Partial<Fmt>): Fmt => ({ ...BASE, ...f });

export const SAMPLES: SampleSpec[] = [
  {
    id: "mixed-runs",
    runs: [
      { text: "Revenue", translated: "Umsatz", fmt: { bold: true } },
      { text: " grew ", translated: " wuchs um ", fmt: {} },
      { text: "5%", translated: "5 %", fmt: { color: "#FF0000", size: 24 } },
      { text: " in Q3", translated: " im Q3", fmt: { italic: true } },
    ],
  },
  {
    id: "single-format",
    runs: [
      { text: "A plain sentence in one font.", translated: "Ein einfacher Satz in einer Schriftart, etwas länger.", fmt: { name: "Georgia", size: 14, color: "#1F4E9A" } },
    ],
  },
  {
    id: "bullets-2para",
    bullets: true,
    runs: [
      { text: "First point\n", translated: "Erster Punkt\n", fmt: { bold: true } },
      { text: "Second point", translated: "Zweiter Punkt", fmt: {} },
    ],
  },
  {
    id: "sub-super",
    runs: [
      { text: "CO", translated: "CO", fmt: {} },
      { text: "2", translated: "2", fmt: { subscript: true } },
      { text: " emissions", translated: " Emissionen", fmt: {} },
    ],
  },
];

// ---------- Office.js helpers ----------

const FONT_PROPS = "font/name,font/size,font/bold,font/italic,font/underline,font/color,font/superscript,font/subscript";

function applyFmt(range: PowerPoint.TextRange, f: Fmt) {
  range.font.name = f.name;
  range.font.size = f.size;
  range.font.bold = f.bold;
  range.font.italic = f.italic;
  range.font.underline = f.underline as PowerPoint.ShapeFontUnderlineStyle;
  range.font.color = f.color;
  range.font.superscript = f.superscript;
  range.font.subscript = f.subscript;
}

/** Values are null when Office.js reports mixed formatting within the range. */
function readFmt(r: PowerPoint.TextRange): Record<string, unknown> {
  return {
    name: r.font.name, size: r.font.size, bold: r.font.bold, italic: r.font.italic,
    underline: r.font.underline, color: r.font.color,
    superscript: r.font.superscript, subscript: r.font.subscript,
  };
}

const shapeName = (sample: string, strat: Strategy) => `S1|${sample}|${strat}`;

async function lastSlide(ctx: PowerPoint.RequestContext) {
  const slides = ctx.presentation.slides;
  slides.load("items");
  await ctx.sync();
  return slides.items[slides.items.length - 1]!;
}

/** Adds one new slide containing a text box per sample x strategy. */
export async function createFixture(): Promise<string> {
  return PowerPoint.run(async (ctx) => {
    ctx.presentation.slides.add();
    await ctx.sync();
    const slide = await lastSlide(ctx);

    let row = 0;
    for (const sample of SAMPLES) {
      let col = 0;
      for (const strat of STRATEGIES) {
        const text = sample.runs.map((r) => r.text).join("");
        const box = slide.shapes.addTextBox(text, {
          left: 10 + col * 235, top: 10 + row * 95, width: 225, height: 85,
        });
        box.name = shapeName(sample.id, strat);
        await ctx.sync();
        let off = 0;
        for (const run of sample.runs) {
          applyFmt(box.textFrame.textRange.getSubstring(off, run.text.length), full(run.fmt));
          off += run.text.length;
        }
        if (sample.bullets) box.textFrame.textRange.paragraphFormat.bulletFormat.visible = true;
        await ctx.sync();
        col++;
      }
      row++;
    }
    return `Fixture slide created (${SAMPLES.length * STRATEGIES.length} text boxes). Select that slide, then run S1.`;
  });
}

interface CharProbe { ch: string; fmt: Record<string, unknown>; bullet: boolean | null; }

async function probe(ctx: PowerPoint.RequestContext, shape: PowerPoint.Shape): Promise<{ text: string; chars: CharProbe[] }> {
  const tr = shape.textFrame.textRange;
  tr.load("text");
  await ctx.sync();
  const text = tr.text;
  const subs: PowerPoint.TextRange[] = [];
  for (let i = 0; i < text.length; i++) {
    const s = tr.getSubstring(i, 1);
    s.load(`text,${FONT_PROPS},paragraphFormat/bulletFormat/visible`);
    subs.push(s);
  }
  await ctx.sync();
  return {
    text,
    chars: subs.map((s) => ({ ch: s.text, fmt: readFmt(s), bullet: s.paragraphFormat.bulletFormat.visible ?? null })),
  };
}

export interface Diff { index: number; char: string; prop: string; expected: unknown; actual: unknown; }
export interface CaseResult {
  sample: string;
  strategy: Strategy;
  expectedText: string;
  actualText: string;
  textOk: boolean;
  fmtMismatchedChars: number;
  totalChars: number;
  bulletOk: boolean | null;
  firstDiffs: Diff[];
  ok: boolean;
  error?: string;
}

const norm = (k: string, v: unknown) => (k === "color" && typeof v === "string" ? v.toLowerCase() : v);

function compare(sample: SampleSpec, actual: { text: string; chars: CharProbe[] }) {
  const expectedText = sample.runs.map((r) => r.translated).join("");
  const diffs: Diff[] = [];
  let mismatched = 0;
  let idx = 0;
  for (const run of sample.runs) {
    const exp = full(run.fmt) as unknown as Record<string, unknown>;
    for (let k = 0; k < run.translated.length; k++, idx++) {
      const got = actual.chars[idx];
      if (!got) continue;
      let bad = false;
      for (const prop of Object.keys(exp)) {
        const a = got.fmt[prop];
        if (norm(prop, a) !== norm(prop, exp[prop])) {
          bad = true;
          if (diffs.length < 6) diffs.push({ index: idx, char: got.ch, prop, expected: exp[prop], actual: a });
        }
      }
      if (bad) mismatched++;
    }
  }
  const bulletOk = sample.bullets === undefined ? null : actual.chars.every((c) => c.bullet === sample.bullets);
  return { expectedText, mismatched, diffs, bulletOk };
}

/** Runs every strategy on the fixture slide currently containing the S1 boxes (last slide). */
export async function runS1(): Promise<CaseResult[]> {
  return PowerPoint.run(async (ctx) => {
    const slide = await lastSlide(ctx);
    slide.shapes.load("items/name");
    await ctx.sync();
    const results: CaseResult[] = [];

    for (const sample of SAMPLES) {
      for (const strat of STRATEGIES) {
        const base: Omit<CaseResult, "ok"> = {
          sample: sample.id, strategy: strat, expectedText: sample.runs.map((r) => r.translated).join(""),
          actualText: "", textOk: false, fmtMismatchedChars: -1, totalChars: 0, bulletOk: null, firstDiffs: [],
        };
        try {
          const shape = slide.shapes.items.find((s) => s.name === shapeName(sample.id, strat));
          if (!shape) throw new Error("fixture shape not found (is the fixture slide the LAST slide?)");
          const tr = shape.textFrame.textRange;

          if (strat === "a-whole" || strat === "b-whole+reformat") {
            tr.text = base.expectedText;
            await ctx.sync();
            if (strat === "b-whole+reformat") {
              let off = 0;
              for (const run of sample.runs) {
                applyFmt(tr.getSubstring(off, run.translated.length), full(run.fmt));
                off += run.translated.length;
              }
              await ctx.sync();
            }
          } else {
            // c: replace each run's text in reverse order so earlier offsets stay valid.
            const starts: number[] = [];
            let off = 0;
            for (const run of sample.runs) { starts.push(off); off += run.text.length; }
            for (let i = sample.runs.length - 1; i >= 0; i--) {
              tr.getSubstring(starts[i]!, sample.runs[i]!.text.length).text = sample.runs[i]!.translated;
            }
            await ctx.sync();
          }

          const after = await probe(ctx, shape);
          const cmp = compare(sample, after);
          const textOk = after.text === base.expectedText;
          results.push({
            ...base, actualText: after.text, textOk, fmtMismatchedChars: cmp.mismatched,
            totalChars: after.chars.length, bulletOk: cmp.bulletOk, firstDiffs: cmp.diffs,
            ok: textOk && cmp.mismatched === 0 && cmp.bulletOk !== false,
          });
        } catch (e) {
          results.push({ ...base, ok: false, error: String(e) });
        }
      }
    }
    return results;
  });
}

/** Sanity check that the fixture itself was created with the intended formatting (before any strategy runs). */
export async function probeSelectedShape(): Promise<unknown> {
  return PowerPoint.run(async (ctx) => {
    const shapes = ctx.presentation.getSelectedShapes();
    shapes.load("items/name");
    await ctx.sync();
    const shape = shapes.items[0];
    if (!shape) return "Select a text box first.";
    const p = await probe(ctx, shape);
    // collapse into runs for readability
    const runs: { text: string; fmt: Record<string, unknown>; bullet: boolean | null }[] = [];
    for (const c of p.chars) {
      const last = runs[runs.length - 1];
      if (last && JSON.stringify(last.fmt) === JSON.stringify(c.fmt) && last.bullet === c.bullet) last.text += c.ch;
      else runs.push({ text: c.ch, fmt: c.fmt, bullet: c.bullet });
    }
    return { name: shape.name, text: p.text, runs };
  });
}
