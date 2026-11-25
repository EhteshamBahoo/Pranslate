import type { TranslateRequest, TranslationProvider } from "./provider";

export interface RobustResult { segments: string[]; failed: number; }

const isAbort = (e: unknown) => e instanceof DOMException && e.name === "AbortError";

/**
 * Translates segments in small chunks. If a chunk fails (e.g. the model returns the wrong item count),
 * it is split in half and retried, down to single items. A single item that still fails is left in its
 * original language and counted in `failed` instead of aborting the whole run.
 */
export async function translateRobust(
  provider: TranslationProvider,
  base: Omit<TranslateRequest, "segments">,
  segments: string[],
  signal?: AbortSignal,
  chunkSize = 8,
): Promise<RobustResult> {
  const out: string[] = new Array(segments.length);
  let failed = 0;

  async function run(from: number, to: number): Promise<void> {
    if (signal?.aborted) throw new DOMException("Aborted", "AbortError");
    const part = segments.slice(from, to);
    try {
      const r = await provider.translate({ ...base, segments: part }, signal);
      if (r.segments.length !== part.length) throw new Error("count mismatch");
      r.segments.forEach((s, i) => { out[from + i] = s; });
    } catch (e) {
      if (isAbort(e)) throw e;
      if (part.length > 1) {
        const mid = from + Math.ceil(part.length / 2);
        await run(from, mid);
        await run(mid, to);
      } else {
        out[from] = part[0]!; // leave untranslated
        failed++;
      }
    }
  }

  for (let i = 0; i < segments.length; i += chunkSize) await run(i, Math.min(i + chunkSize, segments.length));
  return { segments: out, failed };
}
