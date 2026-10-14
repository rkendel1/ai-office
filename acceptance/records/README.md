# Acceptance records

Commit live acceptance records (JSON, schema 1) here, one file per run, named
`<date>-<mode>-<scenarioId>.json` — the format written by the task pane's
**Download JSON** and by the companion. Then regenerate the matrix:

```bash
# from records saved by the companion
loai-companion acceptance --out acceptance/STATUS.md
```

Only records with `"mode": "live"` count as live results. See docs/acceptance.md.
