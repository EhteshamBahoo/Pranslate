import { describe, expect, it } from "vitest";
import { parseSlideSelection } from "./slideRange";
import { paragraphSpans, isTranslatable, splitPadding } from "./paragraphs";

describe("parseSlideSelection", () => {
  it("all", () => expect(parseSlideSelection("all", 3).slides).toEqual([1, 2, 3]));
  it("empty means all", () => expect(parseSlideSelection("", 2).slides).toEqual([1, 2]));
  it("lists and ranges", () => expect(parseSlideSelection("1, 3-5, 4", 10).slides).toEqual([1, 3, 4, 5]));
  it("rejects out of range", () => expect(parseSlideSelection("7", 5).error).toBeTruthy());
  it("rejects junk", () => expect(parseSlideSelection("a-b", 5).error).toBeTruthy());
  it("rejects reversed range", () => expect(parseSlideSelection("5-3", 9).error).toBeTruthy());
});

describe("paragraphSpans", () => {
  it("single paragraph", () => expect(paragraphSpans("abc")).toEqual([{ start: 0, length: 3 }]));
  it("splits on separators", () =>
    expect(paragraphSpans("ab\rcde\nf")).toEqual([
      { start: 0, length: 2 }, { start: 3, length: 3 }, { start: 7, length: 1 },
    ]));
  it("keeps empty paragraphs", () =>
    expect(paragraphSpans("a\r\rb")).toEqual([
      { start: 0, length: 1 }, { start: 2, length: 0 }, { start: 3, length: 1 },
    ]));
});

describe("helpers", () => {
  it("isTranslatable", () => {
    expect(isTranslatable("2024")).toBe(false);
    expect(isTranslatable("売上 2024")).toBe(true);
  });
  it("splitPadding", () => expect(splitPadding("  hi \t")).toEqual({ lead: "  ", core: "hi", trail: " \t" }));
});
