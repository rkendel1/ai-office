# Troubleshooting

Start with **Settings ▸ Run diagnostics** in the task pane. It tests Ollama, lists models, sends a fixed test prompt (no document content), checks the companion and lists Office API support.

| Symptom | Cause | Fix |
|---|---|---|
| Red dot, "Cannot reach the companion" | Companion not running, or certificate not trusted | `loai-companion start`; `npx office-addin-dev-certs install`; open https://localhost:47615/v1/health in a browser once |
| "Not paired" | First run, token revoked, or companion data reset | `loai-companion pair`, enter the code in Settings |
| "Cannot reach Ollama" | Ollama not running or different port | Start Ollama; check `loai-companion config` → `ollamaUrl`; `ollama list` |
| "Model … is not installed" | Selected model removed | Pick another in Settings, or `ollama pull <model>` |
| "This task needs about N tokens but the model is limited to M" | Content + conversation exceed the context window | Select a smaller portion; raise **Context size**; start a new conversation; use a model with a larger window |
| Timeouts on large models | Generation slower than the timeout | Raise **Request timeout**; use a smaller/quantized model |
| Model "talks about" calling tools instead of acting | Model without native tool calling | Diagnostics shows "text tool protocol"; choose a model that reports native tool calling (Diagnostics); lower temperature. See [model-compatibility.md](model-compatibility.md) for measured results |
| Edit refused: "no longer contain the expected text" | Document changed after the assistant read it | Ask again; the assistant re-reads |
| Action shows **Not verified** in Outlook | Reply windows cannot be read back by add-ins | Check the draft window; this is expected |
| Add-in does not appear (Mac) | Manifest not in `wef` folder or Office not restarted | Re-run `install.sh`; quit Office fully (Cmd+Q) |
| Add-in does not appear (Windows) | Developer registry key missing, or Office policy blocks sideloading | Re-run `install.ps1`; or use a shared-folder catalog |
| "Direct" mode fails with CORS errors | Ollama does not allow the add-in origin | Use the companion, or set `OLLAMA_ORIGINS=https://localhost:3000,https://localhost:47615` and restart Ollama |

## Logs and data

| | macOS | Windows | Linux |
|---|---|---|---|
| Data (config, DB, tokens) | `~/Library/Application Support/LocalOfficeAI` | `%APPDATA%\LocalOfficeAI` | `~/.config/local-office-ai` |
| App | same folder `/app` | `%LOCALAPPDATA%\LocalOfficeAI\app` | `~/.local/share/local-office-ai/app` |
| Companion log | `…/LocalOfficeAI/logs/companion.log` | console of the hidden process | console |

Reset everything: `loai-companion revoke`, `loai-companion clear-history`, or uninstall with `--purge` / `-Purge`.
