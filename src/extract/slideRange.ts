/** Parses "all" or "1, 3-5" into sorted unique 1-based slide numbers, clamped-checked against total. */
export function parseSlideSelection(input: string, total: number): { slides: number[]; error?: string } {
  const text = input.trim().toLowerCase();
  if (text === "" || text === "all") return { slides: Array.from({ length: total }, (_, i) => i + 1) };
  const set = new Set<number>();
  for (const part of text.split(/[,\s]+/).filter(Boolean)) {
    const m = /^(\d+)(?:-(\d+))?$/.exec(part);
    if (!m) return { slides: [], error: `Can't read "${part}". Use e.g. 2 or 1, 3-5.` };
    const a = Number(m[1]);
    const b = m[2] ? Number(m[2]) : a;
    if (a < 1 || b < a || b > total) return { slides: [], error: `"${part}" is outside 1-${total}.` };
    for (let n = a; n <= b; n++) set.add(n);
  }
  return { slides: [...set].sort((x, y) => x - y) };
}
