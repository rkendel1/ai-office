# Architecture

Three deployable pieces: the **add-in** (task pane), the **companion** (local service), and **Ollama**. Everything runs on the user's machine.

## Components (spec §7.1)

| Spec component | Where | Notes |
|---|---|---|
| A. Office Add-in | `packages/addin/src/ui` | React task pane: Chat, Context, Actions, History, Settings; approvals with word-level diff |
| B. Host adapters | `packages/core/src/hosts/*` + `packages/addin/src/drivers/*` | Adapters define capabilities over a small **driver** interface; drivers are the only Office.js code. Tests use in-memory drivers with identical interfaces. |
| C. Context manager | `core/src/context.ts` | What was inspected, at what scope, whether truncated, whether stale; budgeted packaging; untrusted-content fencing |
| D. Agent runtime | `core/src/runtime.ts` | Bounded loop, one action per turn, retry limits, cancellation, honest outcome classification |
| E. Capability registry | `core/src/registry.ts` | Only path from a model proposal to Office state |
| F. Inference adapter | `core/src/inference/ollama.ts` | Model discovery, `/api/show` capabilities and context length, streaming chat, native tool calls; `toolcalls.ts` provides a text protocol for models without tool support |
| G. Local companion | `packages/companion` | Loopback HTTPS; pairing-based auth; Ollama allowlist proxy; serves the built add-in |
| H. Local storage | `companion/src/store.ts` | Built-in `node:sqlite`: history (references, not document copies), settings, hashed client tokens |
| Specialist layer | `core/src/specialists` | Task-specific configurations of the runtime: instructions, capability allowlists, required context, guards, completion criteria; deterministic router. See [specialists.md](specialists.md). |

## Request flow (spec §7.2)

1. User submits a request in the pane.
2. The runtime re-checks freshness of previously inspected items, then asks the adapter for **initial context**: metadata plus the selection. Never the whole document by default.
3. The context manager builds a bounded package (selection first, then fresh items, then stale ones) and fences it as untrusted data.
4. The runtime fits the request into the model's real context window (from `/api/show`, capped by `num_ctx`), shrinking Office content first and failing clearly if it still cannot fit.
5. The model returns text, or one or more capability calls (native `tool_calls` or a fenced ```` ```tool ```` block).
6. **Only the first call is processed.** The registry rejects unknown names, internal capabilities, wrong hosts, invalid arguments (strict schema, no extra properties), and out-of-scope targets; computes risk (may escalate, e.g. large overwrite → `destructive`); applies policy; and asks the user via the approver when required.
7. The handler executes through the driver and **reads the affected state back**. `verified: true` → `succeeded`; `false` → `failed`; `null` → `unverified`.
8. The result (status, verification, errors; read content goes into the fenced context package) returns to the model; the loop continues within limits.
9. The runtime produces a `TaskReport` whose **execution summary is generated from records, not from the model's words**. The UI shows both; the summary is authoritative.

## Specialists

A request can be handled by a specialist chosen by the user or by the deterministic router (see [specialists.md](specialists.md)). The same `AgentRuntime.run` executes it with:

- **tools** = the specialist's allowlist ∩ capabilities registered for the host (enforced again by the registry at dispatch);
- **policy** = the user's policy, made stricter where the specialist asks (never looser);
- **required context** acquired through the registry before the first model call;
- **guards** that may reject individual calls (plan required, writes only to new sheets, only new slides, quote text before editing);
- **completion criteria** evaluated after the task against execution records, inspected context and before/after host snapshots. An unmet required criterion turns a "completed" task into "partial", and is listed in the report.

## Outcome classification

| Outcome | When |
|---|---|
| `completed` | No limits hit; every attempted change verified; all required specialist criteria met |
| `partial` | A limit was hit; some changes failed or could not be verified; or a required specialist criterion is unmet |
| `failed` | Model/inference error before any successful change, or all attempted changes failed |
| `cancelled` | User stopped the task. No new actions start; changes already committed by Office remain and are listed |

User-rejected proposals are not failures; they are listed as "Rejected by you", and an identical re-proposal in the same task is blocked without asking again.

## Undo

Handlers return a *compensation* where one is safe: restoring prior paragraph text (guarded by `expectedText`), deleting inserted paragraphs (guarded), restoring prior cell contents, restoring shape text, deleting an added slide. Undo goes through the same registry, policy and approval path, with `origin: "user"` (internal undo-only capabilities are invisible to the model). No rollback is promised for operations without a compensation; the UI points to Office's own Undo.

## Why the companion exists

- Browsers block or restrict calls from an HTTPS task pane to `http://localhost:11434`, and Ollama's CORS must be configured per origin. The companion provides one authenticated, same-machine HTTPS endpoint instead.
- It holds history and settings in SQLite without involving any cloud service.
- It is not an OS bridge: it forwards exactly four Ollama endpoints (`version`, `tags`, `show`, `chat`), plus a user-initiated, explicitly confirmed model download. The model has no capability that reaches it.

The add-in can also talk to Ollama directly (`transport: direct`) if `OLLAMA_ORIGINS` allows the add-in origin; history is then session-only.

## Deliberate limits

- One capability per model turn, so each action is inspected before the next is chosen.
- No "send email" capability. The user sends drafts from Outlook.
- No reading of other documents or unrelated mail (`read_external` exists in the policy for future capabilities and requires confirmation).
- Speaker notes are not offered in live PowerPoint: Office.js does not expose them. The capability exists for hosts that do; the Presentation Author reports notes it could not add.
- PowerPoint charts are drawn with editable shapes, because Office.js cannot create native PowerPoint charts.
- Layout checks inspect slide objects (geometry, text, formatting), never the rendered slide.
