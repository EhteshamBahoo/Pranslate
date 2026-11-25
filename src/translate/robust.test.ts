import { describe, expect, it } from "vitest";
import { translateRobust } from "./robust";
import { coerceTranslations } from "./ollama";
import type { TranslationProvider } from "./provider";

const upper = (s: string) => s.toUpperCase();

describe("translateRobust", () => {
  it("translates in chunks", async () => {
    const calls: number[] = [];
    const p: TranslationProvider = { id: "t", translate: async (r) => { calls.push(r.segments.length); return { segments: r.segments.map(upper) }; } };
    const r = await translateRobust(p, { targetLang: "en" }, ["a", "b", "c", "d", "e"], undefined, 2);
    expect(r.segments).toEqual(["A", "B", "C", "D", "E"]);
    expect(calls).toEqual([2, 2, 1]);
  });

  it("splits a chunk when the model drops items", async () => {
    // Misbehaves for any request with more than 2 items.
    const p: TranslationProvider = { id: "t", translate: async (r) =>
      ({ segments: r.segments.length > 2 ? r.segments.slice(1).map(upper) : r.segments.map(upper) }) };
    const r = await translateRobust(p, { targetLang: "en" }, ["a", "b", "c", "d", "e", "f"], undefined, 6);
    expect(r.segments).toEqual(["A", "B", "C", "D", "E", "F"]);
    expect(r.failed).toBe(0);
  });

  it("keeps the original for an item that always fails and counts it", async () => {
    const p: TranslationProvider = { id: "t", translate: async (r) => {
      if (r.segments.includes("bad")) throw new Error("boom");
      return { segments: r.segments.map(upper) };
    } };
    const r = await translateRobust(p, { targetLang: "en" }, ["a", "bad", "c"], undefined, 3);
    expect(r.segments).toEqual(["A", "bad", "C"]);
    expect(r.failed).toBe(1);
  });

  it("propagates aborts", async () => {
    const ac = new AbortController(); ac.abort();
    const p: TranslationProvider = { id: "t", translate: async (r) => ({ segments: r.segments }) };
    await expect(translateRobust(p, { targetLang: "en" }, ["a"], ac.signal)).rejects.toThrow();
  });
});

describe("coerceTranslations", () => {
  it("accepts {translations: [...]}", () => expect(coerceTranslations({ translations: ["x", "y"] }, 2)).toEqual(["x", "y"]));
  it("accepts a bare array", () => expect(coerceTranslations(["x"], 1)).toEqual(["x"]));
  it("wraps a single string for one segment", () => expect(coerceTranslations({ translations: "x" }, 1)).toEqual(["x"]));
  it("accepts numeric-keyed objects", () => expect(coerceTranslations({ translations: { "1": "a", "2": "b" } }, 2)).toEqual(["a", "b"]));
  it("rejects wrong counts", () => expect(coerceTranslations({ translations: ["x"] }, 2)).toBeNull());
});
