# Pranslate

Open-source, in-place PowerPoint translation add-in. *Translate the language, not the layout.*

> **Status: Phase 0 (feasibility spikes). Not usable yet.**
> Several core assumptions (formatting retention on text replacement, local-model connectivity, overflow detection, speaker notes) are being verified before the MVP is built. Results are recorded in [`docs/spikes/`](docs/spikes) and [`docs/decisions.md`](docs/decisions.md).

## Scope for now

- Target: **PowerPoint for Mac (desktop)**. Windows and PowerPoint on the web are not supported or tested yet.
- Translation: local models via Ollama first; other providers are optional and later.
- Licence: MIT.

## Development

Requires Node 20+ (developed on Node 24) and, for the Ollama spike, [Ollama](https://ollama.com).

```bash
npm install
npx office-addin-dev-certs install   # trusts a localhost HTTPS cert (asks for your password)
npm run dev                          # https://localhost:3000
npm run typecheck && npm run lint && npm test
```

Sideload `manifest.xml` into PowerPoint for Mac (see Microsoft's
[sideload guide](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)).

## Layout

| Path | Purpose |
|---|---|
| `src/office/` | PowerPoint access via Office.js |
| `src/extract/` | Text extraction with references to original locations |
| `src/translate/` | Provider-independent translation interface + Ollama |
| `src/layout/` | Overflow estimation |
| `src/memory/` | Translation memory and originals store |
| `src/ui/` | React task pane |
| `spikes/` | Throwaway experiments, kept separate from product code |
