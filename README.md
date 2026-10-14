# Local Office AI

A local-first AI assistant embedded in Microsoft Word, Excel, PowerPoint and Outlook. It reads the open document, workbook, deck or email, plans multi-step work, proposes changes for your approval, applies them through Office.js, and **verifies the result by reading Office state back** — all with models running on your own machine through [Ollama](https://ollama.com).

No account, no subscription, no cloud AI. If the local model fails, the task fails; there is no fallback.

```
┌──────────── Office task pane (React + Office.js) ────────────┐        ┌─── Local companion (Node) ───┐      ┌──────────┐
│ Chat · Context · Actions · History · Settings                 │ HTTPS  │ loopback only · paired token │ HTTP │  Ollama  │
│ Agent runtime → Capability registry → Host adapter → Office   │───────▶│ Ollama allowlist proxy       │─────▶│ (local)  │
│        ▲ approvals (you)        ▲ read-back verification      │        │ SQLite history & settings    │      └──────────┘
└───────────────────────────────────────────────────────────────┘        └──────────────────────────────┘
```

## Specialists

Fourteen purpose-built assistants share one chat, one runtime and one approval path. Pick one in the task pane or let the router choose from your request (it previews its choice as you type, and asks when a request is ambiguous). See [docs/specialists.md](docs/specialists.md).

| Word | Excel | PowerPoint | Outlook |
|---|---|---|---|
| Document Writer · Document Editor · Document Reviewer · Document Summarizer | Data Analyst · Formula Expert · Data Cleaner · Dashboard Builder | Presentation Author · Slide Designer · Deck Reviewer | Thread Analyst · Reply Writer · Action Extractor |

Each specialist narrows the capabilities available, gathers the context it needs, may require an approved plan, and is checked against concrete completion criteria (e.g. every number in an Excel answer must come from application-computed results; untargeted paragraphs must be unchanged; explicit email commitments must be attributable to their author). A task whose checks fail is reported as *partial*, with the reason.

## What it does

| | Read | Change (with your approval) | Verified by |
|---|---|---|---|
| **Word** | selection, paragraphs (chunked), search, structure | replace selection/paragraphs, insert paragraphs, styles/bold/italic | re-reading the affected paragraphs |
| **Excel** | sheets, used ranges, tables, selection, values + formulas; application-computed statistics and data profiles | write values/formulas, add sheet, cleaned copy to a new sheet, create table, filter, create chart | re-reading cells / listing tables & charts; row-count and source-unchanged checks |
| **PowerPoint** | slides, shapes, positions, formatting, selected slides | slides from layouts, text boxes, shapes, diagrams, bar charts (editable shapes), tables, formatting, geometry, text edits | re-reading slide objects; object-level layout check (not the rendered slide) |
| **Outlook** | open message and quoted thread | reply draft (read mode) or insert into compose body | re-reading the compose body; reply windows are reported *unverified* because Office cannot read them back |

The assistant **cannot send email**, run code, use the shell, read files, or reach the network. Those capabilities do not exist in the registry, so neither a model nor a document can invoke them.

## Quick start (development)

Requirements: Node.js ≥ 22.13, Ollama with at least one model (`ollama pull qwen2.5:7b` — any model works; tool-calling models work best), and Office desktop for live use.

```bash
npm install
npm test                     # 223 automated tests (unit, integration, end-to-end, specialists, security, privacy, UI)
npm run build                # core + companion + add-in + manifests

npm run certs -w @loai/addin # once: trusted https://localhost certificate for Office
npm start                    # companion on https://localhost:47615, serves the built add-in, prints a pairing code
```

Sideload `packages/addin/dist/manifest.xml` into Word/Excel/PowerPoint ([Windows](https://learn.microsoft.com/office/dev/add-ins/testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins), [Mac](https://learn.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac)) and `outlook-manifest.xml` into Outlook. Open **Home ▸ Assistant**, go to **Settings**, enter the pairing code, pick a model.

**Hot-reload UI work:** `npm run dev:addin` serves the pane at https://localhost:3000 (use `manifest.dev.xml`). Open https://localhost:3000/taskpane.html in a normal browser to use **demo mode** — an in-memory Word document (`?demo=excel|powerpoint|outlook` for the others) driven by your real Ollama model.

## Installing a release

```bash
node scripts/release-keys.mjs                       # once; keep keys/release-private.pem secret
LOAI_SIGNING_KEY=keys/release-private.pem npm run package
node scripts/verify-release.mjs release/local-office-ai-0.1.0 keys/release-public.pem
```

Users extract the release and run `./install.sh` (macOS) or `install.ps1` (Windows). The installer verifies checksums, installs the app, creates the localhost certificate, registers the add-in with Word/Excel/PowerPoint, and starts the companion at login. `loai-companion pair` prints a pairing code.

## Repository layout

```
packages/core        Domain logic, no Office/DOM dependency — fully unit-testable
  src/registry.ts      capability registry: schema validation, scope, policy, approval, execution, verification
  src/permissions.ts   approval policy with non-negotiable floors
  src/runtime.ts       bounded agent loop: limits, retries, cancellation, honest outcomes
  src/context.ts       context budgeting, provenance, freshness, untrusted-content fencing
  src/inference/       Ollama adapter (streaming, native tools) + text tool protocol fallback
  src/hosts/           Word / Excel / PowerPoint / Outlook adapters over driver interfaces
  src/testing/         in-memory Office hosts and a scripted model
packages/companion   Loopback HTTPS service: pairing auth, Ollama allowlist proxy, SQLite history, CLI
packages/addin       React task pane, Office.js drivers, manifests
installer/           install/uninstall scripts for macOS and Windows
scripts/             manifest generation, release packaging, signing, verification
docs/                architecture, security, compatibility matrix, testing, troubleshooting
```

## Documentation

- [Architecture](docs/architecture.md) — components, request flow, the proposal/execution boundary
- [Specialists](docs/specialists.md) — the 14 specialists, routing, guarantees and the capabilities they use
- [Security and privacy](docs/security.md) — threat model and how each safeguard is enforced and tested
- [Compatibility matrix](docs/compatibility.md) — hosts, platforms, API sets, models
- [Testing](docs/testing.md) — automated suites and the manual Office acceptance checklist
- [Troubleshooting](docs/troubleshooting.md)

## Status

Version 0.2.0. All automated suites pass, including the spec's acceptance scenarios run against in-memory Office hosts and an HTTP Ollama stand-in, and a UI test that drives the real task pane. The Office.js drivers compile against the official type definitions but **have not yet been exercised inside live Office applications** — this includes every specialist workflow and the new PowerPoint construction capabilities — run the manual checklist in [docs/testing.md](docs/testing.md) on each platform in the compatibility matrix before calling a release production-ready.
