# Compatibility matrix

Status legend: **Auto** = covered by automated tests against in-memory hosts · **Manual** = must be checked in live Office with the checklist in [testing.md](testing.md) · **N/A** = not supported.

## Office hosts and platforms

| Host | Windows (Microsoft 365, 16.0+) | macOS (Microsoft 365, 16.x) | Office on the web | Required API sets |
|---|---|---|---|---|
| Word | Manual | Manual | Not a launch target¹ | WordApi 1.1 (all); 1.3 for positioned insert, list/table detection, multi-paragraph replace |
| Excel | Manual | Manual | Not a launch target¹ | ExcelApi 1.1; 1.2 filters; 1.4 null-safe used range |
| PowerPoint | Manual | Manual | Not a launch target¹ | PowerPointApi 1.3 add slide; 1.4 read/edit text, text boxes; 1.5 selected slides |
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
| speaker notes | — | not offered (no Office.js API) |
| sending email | — | intentionally not offered |

## Runtime

| Component | Requirement |
|---|---|
| Node.js (companion) | ≥ 22.13 (uses built-in `node:sqlite`) |
| Ollama | ≥ 0.4 for native tool calling; older versions work via the text tool protocol. `/api/show` `capabilities` is used when present (≥ 0.6), otherwise inferred from the template |
| TLS | `office-addin-dev-certs` certificate for `https://localhost` (required on macOS; recommended everywhere) |

## Models

Any Ollama chat model works. Guidance:

| Model class | Tool calling | Notes |
|---|---|---|
| Llama 3.1/3.2/3.3, Qwen 2.5/3, Mistral Nemo/Small, Command-R | Native | Recommended. 7–8B models handle single-step edits well; multi-step Excel work benefits from 14B+ |
| Gemma, Phi, older Llama 2-era models | Text protocol | Usable; more malformed calls (these are fed back and retried within the model-call limit) |
| Reasoning models (DeepSeek-R1 distills, QwQ) | Varies | Higher latency; set a larger timeout |

Set **Context size (num_ctx)** to what your RAM/VRAM allows; the runtime uses the smaller of that and the model's maximum, shrinks Office content to fit, and fails clearly when even the minimum does not fit.

Record live results for each row (host × OS × Office build × model) in this file as they are verified.
