export interface Span { start: number; length: number; }

/** Paragraph spans of a text frame's text (paragraph separators: \r, \n, \v). Separators are excluded. */
export function paragraphSpans(text: string): Span[] {
  const spans: Span[] = [];
  let start = 0;
  for (let i = 0; i <= text.length; i++) {
    if (i === text.length || text[i] === "\r" || text[i] === "\n" || text[i] === "\v") {
      spans.push({ start, length: i - start });
      start = i + 1;
    }
  }
  return spans;
}

/** True if the string has at least one letter (skips empty strings, bare numbers and symbols). */
export const isTranslatable = (s: string) => /\p{L}/u.test(s);

/** Splits into leading whitespace, core, trailing whitespace so surrounding spacing can be re-attached. */
export function splitPadding(s: string): { lead: string; core: string; trail: string } {
  const m = /^(\s*)([\s\S]*?)(\s*)$/.exec(s)!;
  return { lead: m[1]!, core: m[2]!, trail: m[3]! };
}
