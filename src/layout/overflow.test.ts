import { describe, expect, it } from "vitest";
import { estimateLines, estimateOverflow } from "./overflow";

// Fixed-width fake font: 1pt per character.
const measure = (t: string) => t.length;

describe("estimateLines", () => {
  it("fits a short string on one line", () => {
    expect(estimateLines("hello world", 20, measure)).toBe(1);
  });
  it("wraps at word boundaries", () => {
    expect(estimateLines("aaaa bbbb cccc", 9, measure)).toBe(2);
  });
  it("counts explicit paragraph breaks", () => {
    expect(estimateLines("a\nb\nc", 10, measure)).toBe(3);
  });
  it("does not loop forever on a word longer than the box", () => {
    expect(estimateLines("abcdefghijkl", 5, measure)).toBe(1);
  });
});

describe("estimateOverflow", () => {
  const box = {
    widthPt: 24, heightPt: 30,
    insetLeftPt: 2, insetRightPt: 2, insetTopPt: 1, insetBottomPt: 1,
    fontSizePt: 10, lineSpacing: 1,
  };
  it("flags text needing more height than available", () => {
    // 20pt usable width -> 4 words per line -> 3 lines * 10pt = 30pt > 28pt available
    const r = estimateOverflow("aaaa bbbb cccc dddd eeee ffff gggg hhhh iiii", box, measure);
    expect(r.lines).toBe(3);
    expect(r.overflows).toBe(true);
  });
  it("passes text that fits", () => {
    expect(estimateOverflow("hi", box, measure).overflows).toBe(false);
  });
});
