# Model evaluation runs

Commit run files produced by `loai-companion eval --model <name> --out evals/results` here (both the `.json` and the `.md`).

- Only runs of kind `ollama` (a real local model) appear in the compatibility matrix. Self-test runs (`--self-test`) check the harness and are never counted as model results.
- Do not edit run files by hand. A run with a changed number is no longer a measurement.
- Rebuild the matrix after adding runs: `loai-companion eval-matrix --dir evals/results` and paste the output into `docs/model-compatibility.md`.

This directory has no model runs yet.
