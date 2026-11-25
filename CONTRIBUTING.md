# Contributing

The project is in Phase 0 (feasibility spikes), so the architecture may change. Please open an issue before large changes.

- Run `npm run typecheck && npm run lint && npm test` before opening a PR.
- Keep PowerPoint-specific code in `src/office/`; keep pure logic (segmentation, overflow math) testable without Office.
- Test decks in `spikes/fixtures/` must not contain confidential material.
