# Compatibility matrix

Status legend: **Auto** = covered by automated tests against in-memory hosts · **Manual** = must be checked in live Office with the checklist in [testing.md](testing.md) · **N/A** = not supported.

## Office hosts and platforms

| Host | Windows (Microsoft 365, 16.0+) | macOS (Microsoft 365, 16.x) | Office on the web | Required API sets |
|---|---|---|---|---|
| Word | Manual | Manual | Not a launch target¹ | WordApi 1.1 (all); 1.3 for positioned insert, list/table detection, multi-paragraph replace |
| Excel | Manual | Manual | Not a launch target¹ | ExcelApi 1.1; 1.2 filters; 1.4 null-safe used range |
| PowerPoint | Manual | Manual | Not a launch target¹ | PowerPointApi 1.3 slides from layouts; 1.4 text boxes, shapes, geometry, fonts, fills; 1.5 selected slides; 1.8 tables |
| Outlook | Manual | Manual (new and classic) | Not a launch target¹ | Mailbox 1.1 read + reply form; 1.2 insert into compose body |

¹ The architecture permits it, but a web-hosted pane generally cannot reach `localhost` without extra configuration (spec §4).

The add-in checks requirement sets at runtime (`Office.context.requirements.isSetSupported`). Capabilities that need an unavailable set are either not offered to the model or rejected with a clear message; **Settings ▸ Diagnostics** lists each requirement and whether it is supported.

## Capability → API dependency

| Capability | API set | Behaviour if missing |
|---|---|---|
| `document.insert_text` before/after paragraph | WordApi 1.3 | rejected: "requires WordApi 1.3" |
| `document.replace_text` across paragraphs | WordApi 1.3 (`Range.expandTo`) | falls back to replace-first + delete-rest on 1.1 |
| `workbook.filter_table` | ExcelApi 1.2 | not offered |
| `presentation.add_slide` | PowerPointApi 1.3 | not offered |
| selected slides in context | PowerPointApi 1.5 | reported as "not reported by this PowerPoint version" |
| `presentation.list_layouts`, `create_slide` | PowerPointApi 1.3 | not offered |
| `presentation.add_text_box`, `add_shape`, `add_diagram`, `add_bar_chart`, `format_text`, `set_geometry` | PowerPointApi 1.4 | not offered |
| `presentation.add_table` | PowerPointApi 1.8 | not offered |
| speaker notes (`presentation.set_notes`) | — | not offered (no Office.js API); slides are created without notes and the action is marked *partial* |
| slide size | — | not reported by Office.js; 960×540 (16:9) assumed and stated in layout checks |
| `workbook.compute_stats`, `profile_data`, `clean_to_sheet` | ExcelApi 1.1 (+1.4 for empty used ranges) | — |
| `document.write_blocks` | WordApi 1.1 (1.3 for positioned inserts) | — |
| `outlook.split_thread` | Mailbox 1.1 | — |
| sending email | — | intentionally not offered |

## Runtime

| Component | Requirement |
|---|---|
| Node.js (companion) | ≥ 22.13 (uses built-in `node:sqlite`) |
| Ollama | ≥ 0.4 for native tool calling; older versions work via the text tool protocol. `/api/show` `capabilities` is used when present (≥ 0.6), otherwise inferred from the template |
| TLS | `office-addin-dev-certs` certificate for `https://localhost` (required on macOS; recommended everywhere) |

## Models

Any Ollama chat model can be selected. Models that report native tool calling (`/api/show` capabilities) use it; others use the text tool protocol, where malformed calls are fed back and retried within the model-call limit.

**No model has been evaluated yet**, so this project makes no claim about which models work well. Measured results, and how to produce them with `loai-companion eval`, are in [model-compatibility.md](model-compatibility.md). Only models listed there have been tested.

Set **Context size (num_ctx)** to what your RAM/VRAM allows; the runtime uses the smaller of that and the model's maximum, shrinks Office content to fit, and fails clearly when even the minimum does not fit.

## Specialists

All 14 specialists are covered by automated workflow tests against the in-memory hosts. **None has yet been exercised in live Office.** Until each is run through the checklist in [testing.md](testing.md) on a host, treat its live behaviour as unverified.

| Specialist | Word Win | Word Mac | Excel Win | Excel Mac | PPT Win | PPT Mac | Outlook Win | Outlook Mac |
|---|---|---|---|---|---|---|---|---|
| Word specialists (4) | Manual | Manual | | | | | | |
| Excel specialists (4) | | | Manual | Manual | | | | |
| PowerPoint specialists (3) | | | | | Manual | Manual | | |
| Outlook specialists (3) | | | | | | | Manual | Manual |

Record live results for each row (host × OS × Office build × model) in this file as they are verified.
