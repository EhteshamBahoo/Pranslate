import type { TranslateRequest, TranslateResult, TranslationProvider } from "./provider";

export const LANG_NAME: Record<string, string> = { ja: "Japanese", en: "English" };

export interface OllamaOptions {
  /** Base URL. In dev the Vite server proxies "/ollama" to http://localhost:11434 (same-origin, no CORS/mixed content). */
  baseUrl?: string;
  model: string;
}

export class OllamaProvider implements TranslationProvider {
  readonly id = "ollama";
  constructor(private opts: OllamaOptions) {}
  private get base() { return this.opts.baseUrl ?? "/ollama"; }

  async listModels(): Promise<string[]> {
    const r = await fetch(`${this.base}/api/tags`);
    if (!r.ok) throw new Error(`Ollama responded ${r.status}`);
    const j = (await r.json()) as { models?: { name: string }[] };
    return (j.models ?? []).map((m) => m.name);
  }

  async translate(req: TranslateRequest, signal?: AbortSignal): Promise<TranslateResult> {
    const src = req.sourceLang ? LANG_NAME[req.sourceLang] ?? req.sourceLang : "the source language";
    const dst = LANG_NAME[req.targetLang] ?? req.targetLang;
    const system =
      `You translate PowerPoint slide text from ${src} to ${dst}. ` +
      `Translate meaning naturally and concisely (slide text must stay short). ` +
      `Keep numbers, product names, URLs and code unchanged. Do not add explanations. ` +
      `Input is a JSON array of strings. Reply ONLY with JSON of the form {"translations": [...]} ` +
      `with exactly the same number of items in the same order.`;
    const user = JSON.stringify(req.segments) + (req.context ? `\n\nSlide context (do not translate): ${req.context}` : "");
    const r = await fetch(`${this.base}/api/chat`, {
      method: "POST",
      signal,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: this.opts.model,
        stream: false,
        format: "json",
        options: { temperature: 0.2 },
        messages: [{ role: "system", content: system }, { role: "user", content: user }],
      }),
    });
    if (!r.ok) throw new Error(`Ollama responded ${r.status}: ${(await r.text()).slice(0, 200)}`);
    const j = (await r.json()) as { message?: { content?: string } };
    const parsed = JSON.parse(j.message?.content ?? "{}") as { translations?: unknown };
    const arr = parsed.translations;
    if (!Array.isArray(arr) || arr.length !== req.segments.length || !arr.every((x) => typeof x === "string")) {
      throw new Error("Model returned an unexpected shape (count mismatch). Try again or use a larger model.");
    }
    return { segments: arr as string[] };
  }
}
