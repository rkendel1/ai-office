# Live Office acceptance

Acceptance scenarios exercise each Office app end to end: create fixture content, run a specialist task, then check the application's real state. Each scenario is defined once (`packages/core/src/acceptance/scenarios.ts`) and used three ways:

| Mode | Where | What it proves | Recorded as |
|---|---|---|---|
| **Mocked** | CI (`npm test`) against in-memory hosts | The runtime, specialists, approvals, verification and the scenario's own checks behave as designed | `mode: "mocked"` |
| **Live, scripted** | Inside real Office, from the task pane | The Office.js drivers and adapters really create, read back and verify content in that Office build | `mode: "live"`, `model.kind: "scripted"` |
| **Live, with a model** | Inside real Office, with your local model | The above plus how a specific model performs the task | `mode: "live"`, `model.kind: "ollama"` |
| **Manual / visual** | A person, following the steps listed in the record | What automation cannot see (how slides look, Sent Items, undo) | `manual[].status`, `visual` |

**A mocked pass is never a live pass.** The status matrix keeps them in separate tables, and every scenario without a live record is listed as not run.

## Scenarios

| ID | App | Covers |
|---|---|---|
| `word.writer.structured-document` | Word | Create + format content (Title/Heading/Normal) under a batch-approved plan; read-back of text and style; source notes unchanged |
| `word.editor.scoped-edit` | Word | Edit one paragraph with a diff approval; untargeted paragraphs unchanged; facts kept |
| `excel.dashboard.traceable-metrics` | Excel | Statistics, dashboard sheet with formula metrics; values Excel calculated compared with independently computed totals; chart present; source unchanged |
| `excel.cleaner.preserves-source` | Excel | Profile, cleaned copy on a new sheet, row-count invariant, source unchanged |
| `powerpoint.author.multi-slide-deck` | PowerPoint | Three slides (one chart) under one plan approval; object-level check; rendered check where supported; visual steps for a person |
| `outlook.thread.attribution` | Outlook | Thread split by sender; commitments attributed to their authors |
| `outlook.reply.draft-not-sent` | Outlook | Reply draft created; verified in compose mode, honestly unverified in read mode; no send capability; nothing sent |

## Running live

1. `npm run build && npm start`; sideload the manifest; pair; (for model runs) choose a model.
2. Open a **new blank** document / workbook / presentation. For Outlook, import `acceptance/fixtures/budget-thread.eml` and open it.
3. Task pane ▸ **Settings ▸ Acceptance tests**: tick *This is a blank test document*, expand a scenario, choose **Run (scripted)** or **Run with my model**.
4. Do the listed manual steps and set each to passed/failed in the result.
5. **Save record** (stored by the companion under `<data dir>/acceptance/`) or **Download JSON** and commit it to `acceptance/records/`.
6. Print the matrix: `loai-companion acceptance --out acceptance/STATUS.md`.

Scenarios refuse to run when their preconditions fail (a non-blank Word document, existing `LOAI_` sheets, an existing acceptance deck, a different Outlook message), so they never write into real work.

## Running as a manual checklist

When an Office host is not available to the automated runner, follow the same scenario by hand: its preconditions, the request text, the expected checks (table above and `checks()` in `scenarios.ts`), and the manual steps. Record the result with the template below, `mode: "live"`, `model.kind` set to what you used, and each check's evidence written from what you observed.

## Record format

```json
{
  "schema": 1,
  "scenarioId": "excel.dashboard.traceable-metrics",
  "mode": "live",
  "model": { "kind": "ollama", "name": "qwen2.5:14b" },
  "environment": { "officeApp": "Excel", "officeVersion": "16.0.18129.20000", "platform": "PC", "os": "…", "addinVersion": "0.3.0", "companionVersion": "0.3.0", "ollamaVersion": "0.12.3", "recordedAt": "2026-10-20T14:03:11.000Z" },
  "outcome": "passed",
  "checks": [{ "id": "excel-calculated", "description": "…", "passed": true, "evidence": "Excel: 345, 394; expected 345, 394" }],
  "task": { "outcome": "completed", "summary": "…", "actions": [{ "capability": "workbook.write_range", "status": "succeeded", "approval": "plan", "verified": true, "method": "re-read range values and formulas" }], "completion": [], "limitations": [] },
  "visual": "not-applicable",
  "manual": [{ "step": "Click B2 and B3 …", "status": "passed" }],
  "durationMs": 8123
}
```

## Status

No live runs have been recorded yet: this repository has not been tested in a real Office host. All seven scenarios pass in mocked mode in CI.

### Live Office (real hosts)
_No live runs recorded yet._

Not yet run live: `word.writer.structured-document`, `word.editor.scoped-edit`, `excel.dashboard.traceable-metrics`, `excel.cleaner.preserves-source`, `powerpoint.author.multi-slide-deck`, `outlook.thread.attribution`, `outlook.reply.draft-not-sent`.

### Mocked hosts (in-memory, CI)
All seven pass on every CI run (`packages/core/test/acceptance-suite.test.ts`).
