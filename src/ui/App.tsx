import { useEffect, useRef, useState } from "react";
import { getSlideCount, hasSavedOriginals, restoreOriginals, translateSlides, type Progress } from "../office/deckTranslate";
import { parseSlideSelection } from "../extract/slideRange";
import { OllamaProvider } from "../translate/ollama";
import { createFixture, probeSelectedShape, runS1 } from "../../spikes/s1/s1";
import { readActiveSlideText } from "../office/textAccess";
import { scanSelectedSlide } from "../../spikes/scan";

type Dir = "ja-en" | "en-ja";
const DEFAULT_MODEL = "qwen2.5:7b"; // assumption: reasonable ja<->en quality at 7B; unverified on this machine

export function App({ diag }: { diag: string }) {
  const [dir, setDir] = useState<Dir>("ja-en");
  const [scope, setScope] = useState<"all" | "some">("all");
  const [slideInput, setSlideInput] = useState("");
  const [total, setTotal] = useState<number | null>(null);
  const [model, setModel] = useState(DEFAULT_MODEL);
  const [models, setModels] = useState<string[] | null>(null);
  const [ollamaErr, setOllamaErr] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);
  const [progress, setProgress] = useState<Progress | null>(null);
  const [status, setStatus] = useState<{ text: string; kind: "ok" | "err" | "" }>({ text: "", kind: "" });
  const [canRestore, setCanRestore] = useState(false);
  const [devOut, setDevOut] = useState("");
  const abort = useRef<AbortController | null>(null);

  const provider = () => new OllamaProvider({ model });

  const refreshSlides = () => getSlideCount().then(setTotal).catch(() => setTotal(null));
  const checkOllama = () =>
    new OllamaProvider({ model }).listModels()
      .then((m) => { setModels(m); setOllamaErr(null); if (m.length && !m.includes(model)) setModel(m[0]!); })
      .catch((e: unknown) => { setModels(null); setOllamaErr(String(e)); });

  useEffect(() => { void refreshSlides(); void checkOllama(); }, []);

  async function translate() {
    setStatus({ text: "", kind: "" });
    const count = total ?? (await getSlideCount());
    const sel = parseSlideSelection(scope === "all" ? "all" : slideInput, count);
    if (sel.error) return setStatus({ text: sel.error, kind: "err" });
    abort.current = new AbortController();
    setBusy(true);
    const [src, dst] = dir === "ja-en" ? ["ja", "en"] : ["en", "ja"];
    try {
      const s = await translateSlides(sel.slides, provider(), src!, dst!, setProgress, abort.current.signal);
      setCanRestore(hasSavedOriginals());
      const skipped = Object.entries(s.skippedByType).map(([k, v]) => `${v}× ${k}`).join(", ");
      const parts = [
        s.textItems === 0 ? "No translatable text found." : `Done: ${s.paragraphsChanged} paragraphs in ${s.textItems} text items across ${s.slidesDone} slide(s).`,
        s.paragraphsFailed ? `${s.paragraphsFailed} paragraph(s) could not be translated and were left as-is.` : "",
        skipped ? `Skipped (not editable by this add-in): ${skipped}.` : "",
        ...s.warnings,
        s.textItems ? "Check long text for overflow." : "",
      ].filter(Boolean);
      setStatus({ text: parts.join(" "), kind: s.paragraphsFailed || s.warnings.length || skipped ? "" : "ok" });
    } catch (e) {
      setCanRestore(hasSavedOriginals());
      setStatus({ text: `Stopped: ${String(e)}`, kind: "err" });
    } finally { setBusy(false); setProgress(null); }
  }

  async function restore() {
    setBusy(true);
    try {
      const r = await restoreOriginals();
      setCanRestore(hasSavedOriginals());
      setStatus({ text: `Restored ${r.restored} text box(es). ${r.warnings.join(" ")}`, kind: r.warnings.length ? "" : "ok" });
    } catch (e) { setStatus({ text: String(e), kind: "err" }); }
    finally { setBusy(false); }
  }

  const dev = (fn: () => Promise<unknown>) => async () => {
    setBusy(true);
    try { const r = await fn(); setDevOut(typeof r === "string" ? r : JSON.stringify(r, null, 2)); }
    catch (e) { setDevOut(`ERROR: ${String(e)}`); } finally { setBusy(false); }
  };

  const pct = progress ? Math.round((progress.slideIndex / progress.slideCount) * 100) : 0;
  return (
    <div className="app">
      <div><h1>Pranslate</h1><p className="sub">Translate slides in place. Japanese ⇄ English.</p></div>

      <div>
        <label>Direction</label>
        <div className="seg">
          <button aria-pressed={dir === "ja-en"} onClick={() => setDir("ja-en")} disabled={busy}>日本語 → English</button>
          <button aria-pressed={dir === "en-ja"} onClick={() => setDir("en-ja")} disabled={busy}>English → 日本語</button>
        </div>
      </div>

      <div>
        <label>Slides {total != null && <span className="msg">({total} in deck)</span>}</label>
        <div className="seg">
          <button aria-pressed={scope === "all"} onClick={() => setScope("all")} disabled={busy}>All slides</button>
          <button aria-pressed={scope === "some"} onClick={() => { setScope("some"); void refreshSlides(); }} disabled={busy}>Choose…</button>
        </div>
        {scope === "some" && (
          <input type="text" placeholder="e.g. 2 or 1, 3-5" value={slideInput}
            onChange={(e) => setSlideInput(e.target.value)} disabled={busy} style={{ marginTop: 6 }} />
        )}
      </div>

      <button className="primary" onClick={translate} disabled={busy}>{busy ? "Translating…" : "Translate"}</button>
      {busy && progress && (
        <div><div className="bar"><div style={{ width: `${pct}%` }} /></div>
          <p className="msg">{progress.message} ({progress.slideIndex + 1}/{progress.slideCount})</p>
          <button className="secondary" onClick={() => abort.current?.abort()}>Cancel</button></div>
      )}
      {status.text && <p className={`msg ${status.kind}`}>{status.text}</p>}
      {canRestore && !busy && <button className="secondary" onClick={restore}>Restore original text</button>}

      <div className="note">
        Uses a local Ollama model ({models ? `${models.length} installed` : "not reachable"}); nothing leaves your Mac.
        {ollamaErr && <span className="err"> Can't reach Ollama: {ollamaErr}. Install it, run <b>ollama serve</b>, pull a model, then{" "}
          <a href="#" onClick={(e) => { e.preventDefault(); void checkOllama(); }}>retry</a>.</span>}
        {models && models.length > 0 && (
          <div style={{ marginTop: 6 }}><select value={model} onChange={(e) => setModel(e.target.value)} disabled={busy}>
            {models.map((m) => <option key={m}>{m}</option>)}</select></div>
        )}
        {models && models.length === 0 && <span className="err"> No models installed. Run: ollama pull {DEFAULT_MODEL}</span>}
        <div style={{ marginTop: 6 }}>Restore only works until this pane is closed. Save a copy first for important decks.</div>
      </div>

      <details>
        <summary>Developer (spikes)</summary>
        <pre className="msg" style={{ whiteSpace: "pre-wrap" }}>{diag}</pre>
        <div className="dev">
          <button className="secondary" disabled={busy} onClick={dev(createFixture)}>S1 · 1. Create fixture slide</button>
          <button className="secondary" disabled={busy} onClick={dev(probeSelectedShape)}>S1 · 2. Probe selected text box</button>
          <button className="secondary" disabled={busy} onClick={dev(runS1)}>S1 · 3. Run (modifies fixture slide)</button>
          <button className="secondary" disabled={busy} onClick={dev(scanSelectedSlide)}>Scan selected slide structure</button>
          <button className="secondary" disabled={busy} onClick={dev(readActiveSlideText)}>Read selected slide</button>
        </div>
        <textarea readOnly value={devOut} onFocus={(e) => e.currentTarget.select()} />
      </details>
    </div>
  );
}
