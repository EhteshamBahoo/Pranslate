/** Provider-independent translation interface (proposal §5). Skeleton only in Phase 0. */
export interface TranslateRequest {
  /** Segments to translate, in document order. Inline markup tags (e.g. <b1>..</b1>) must be preserved. */
  segments: string[];
  sourceLang?: string; // undefined = auto-detect
  targetLang: string;
  /** Short deck summary / neighbouring text to give the model context. */
  context?: string;
  glossary?: Record<string, string>;
}

export interface TranslateResult {
  segments: string[];
  detectedSourceLang?: string;
}

export interface TranslationProvider {
  readonly id: string;
  translate(req: TranslateRequest, signal?: AbortSignal): Promise<TranslateResult>;
}
