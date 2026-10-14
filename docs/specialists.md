# Specialists

Specialists turn the general Office agent into purpose-built assistants. Each one is a **configuration of the existing runtime**: task instructions, a capability allowlist, required context, constraints and completion criteria. There is no second runtime, no new service, and no separate approval or verification path.

```
request ──► SpecialistRouter (user text + host + explicit choice only)
              │  specialist | clarify | general | rejected
              ▼
AgentRuntime.run(request, { specialist })
  1. acquire required context through the registry (normal policy)
  2. snapshot host state (for preservation checks)
  3. model sees: base prompt + specialist instructions; tools = allowlist ∩ registered
  4. every call: specialist guard → plan requirement → registry (schema, scope, policy, approval, verify)
  5. completion criteria evaluated against records, context and snapshots
  6. "completed" downgraded to "partial" if a required criterion is unmet
```

## What a specialist can and cannot do

| A specialist can | A specialist cannot |
|---|---|
| Narrow the capabilities offered to the model (allowlist ∩ registered for the host) | Add a capability, use one from another host, or use an internal (undo-only) capability |
| Make the approval policy stricter for its task | Loosen the user's policy (`tightenPolicy` keeps the stricter decision; floors still apply) |
| Reject individual calls with a guard (e.g. "write only to sheets created in this task") | Approve anything — approvals still come only from the UI |
| Require an approved plan before any change | Treat plan approval as approval of the changes (each change is still approved per policy) |
| Define completion criteria checked against observed results | Make the runtime report success the records do not support |

The router uses only the user's request text, the active host and the user's explicit choice. Model output and document content never select a specialist. A specialist for another host is rejected by the router and refused by the runtime.

## Routing

1. If the user chose a specialist in the task pane, it is used (if it belongs to the active host).
2. Otherwise each specialist for the host is scored by weighted patterns over the request.
3. Below the minimum score → **general assistant**.
4. A specialist whose workflow includes the top match's work and also matched wins (e.g. *Reply Writer* covers *Thread Analyst* when the user also asks for a reply; *Dashboard Builder* covers *Data Analyst*).
5. If two unrelated specialists are within one point → the user is asked which to use.

The task pane previews the routing decision while you type ("Will use Document Editor · change") so you can switch before anything runs. Switching keeps the conversation; the new specialist's capabilities and required context are re-applied on the next task.

## Catalogue

### Word

| Specialist | Purpose | Capabilities | Plan | Key completion checks |
|---|---|---|---|---|
| **Document Writer** | Substantial documents from instructions, notes or the open document | read, search, `plan.propose`, `document.write_blocks`, insert, replace, format | Required (except one-paragraph tasks) | outline approved before writing · content written and verified · title/headings + body · unrelated paragraphs unchanged |
| **Document Editor** | Clarity, tone, grammar, concision of a passage | read, selection, search, replace, write_blocks, format | — | edit verified · numbers and proper nouns kept (unless removal requested) · unrelated paragraphs unchanged |
| **Document Reviewer** | Findings: ambiguity, contradictions, gaps, unsupported claims | read, search (read-only) | — | document read · `[Observation]` / `[Recommendation]` findings · paragraph citations · partial reads disclosed · nothing modified |
| **Document Summarizer** | Audience-specific summaries | read, selection, search (read-only) | — | content read · partial reads disclosed · every number traced to the text · nothing modified |

### Excel

| Specialist | Purpose | Capabilities | Plan | Key completion checks |
|---|---|---|---|---|
| **Data Analyst** | Analytical questions with application-computed numbers | read, `workbook.compute_stats`, `workbook.profile_data`, add sheet, write | — | computed by the application · every number grounded · Sheet!Range citations · hypotheses labelled (advisory) |
| **Formula Expert** | Create, explain, debug, repair formulas | read, write | — | formula written and read back (when a write was asked) · no Excel error values · references explained · other cells unchanged |
| **Data Cleaner** | Duplicates, missing values, formats, categories, outliers | read, `workbook.profile_data`, `workbook.clean_to_sheet`, add sheet, write (new sheets only) | — | profiled · cleaned copy verified (row-count invariant, source unchanged) · source preserved |
| **Dashboard Builder** | Editable dashboard: linked metrics, tables, charts | read, compute, profile, plan, add sheet, write (formulas only, new sheets), table, chart | Required | plan approved · sheet created · metrics are formulas referencing source ranges · chart verified · numbers grounded · source preserved |

### PowerPoint

| Specialist | Purpose | Capabilities | Plan | Key completion checks |
|---|---|---|---|---|
| **Presentation Author** | Complete, editable decks from a brief or pasted material | read, layouts, `create_slide`, text boxes, shapes, `add_diagram`, `add_bar_chart`, `add_table`*, `set_notes`*, format, edit, geometry, `check_layout` | Required | plan approved · requested slide count created and verified · distinct titles/content · layout checked after the last change · serious issues reported · unsupported notes reported · existing slides unchanged |
| **Slide Designer** | Visual hierarchy, density, alignment, consistency | read, `check_layout`, format, geometry, edit (must quote current text), shapes, text boxes | — | checked before changes · changes verified · numbers kept · re-checked after · untouched shapes unchanged |
| **Deck Reviewer** | Narrative, flow, evidence, consistency review | read, `check_layout` (read-only) | — | deck read · layout checked · prioritised, slide-cited recommendations · serious issues reported · no claim of visual verification · nothing modified |

\* offered only where the host supports it (tables: PowerPointApi 1.8; speaker notes: not available through Office.js today — notes are reported in chat and the slide action is marked *partial*).

### Outlook

| Specialist | Purpose | Capabilities | Key completion checks |
|---|---|---|---|
| **Thread Analyst** | Summary, decisions, open questions, attributed commitments | read, `outlook.split_thread` (read-only) | thread split · each explicit commitment cites a real message, shares its wording and is owned by that message's author · nothing modified |
| **Reply Writer** | Reply draft answering the thread | read, split, `outlook.create_draft` | thread read · draft created (unverified outcomes reported honestly) · nothing sent · placeholders pointed out |
| **Action Extractor** | Tasks, owners, deadlines, dependencies | `outlook.split_thread` (read-only) | ≥1 structured item · explicit items traced to the message and to an owner who wrote or is named in it · nothing modified |

Outlook items use a checked line format:

```
- [explicit] Send the revised budget — owner: Alex Lee — due: Friday — source: M1
- [inferred] Share the agenda — owner: unknown — due: none — source: none
```

## Capabilities added for specialists

| Capability | Host | Risk | Verification |
|---|---|---|---|
| `plan.propose` | all | modify, **always asks** | — (grants nothing) |
| `document.write_blocks` | Word | modify / destructive for large replacements | re-read text and style of every written paragraph; undo |
| `workbook.compute_stats` | Excel | read | deterministic (count, sum, mean, median, min, max, stdev, distinct, missing; group-by) |
| `workbook.profile_data` | Excel | read | deterministic (types, missing, text-numbers, IQR outliers, spelling variants, duplicates) |
| `workbook.clean_to_sheet` | Excel | modify | re-read new sheet; rows in = out + removed; source fingerprint unchanged |
| `presentation.list_layouts` | PowerPoint | read | — |
| `presentation.create_slide` | PowerPoint | modify | re-read title/bullets/notes; unused placeholders removed; undo |
| `presentation.add_text_box`, `add_shape` | PowerPoint | modify | re-read text, position, fill, format; undo |
| `presentation.format_text`, `set_geometry` | PowerPoint | modify | re-read format/geometry; text unchanged; undo |
| `presentation.add_diagram` | PowerPoint | modify | every item shape present with its text; undo |
| `presentation.add_bar_chart` | PowerPoint | modify | labels equal the data; bar heights proportional (±2%); undo |
| `presentation.add_table` | PowerPoint (1.8) | modify | re-read cell values; undo |
| `presentation.set_notes` | PowerPoint (where supported) | modify | re-read notes |
| `presentation.check_layout` | PowerPoint | read | object-level: overflow estimate, overlaps, off-slide, missing titles, empty placeholders, density, small text, inconsistent title sizes. **Rendering is not inspected.** |
| `outlook.split_thread` | Outlook | read | deterministic parse of quoted replies (Outlook and "On … wrote:" headers) |

None of these adds file, network, shell or code-execution ability.
