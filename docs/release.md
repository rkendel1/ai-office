# Release procedure

A release is a directory (and `.tar.gz`) containing the built add-in, companion and core, the installers, the docs and the acceptance fixtures, plus `SHA256SUMS` over every file and an Ed25519 signature of that file, `SHA256SUMS.sig`.

## 1. Gates

Run all of these from a clean checkout of the release commit. Every one must pass:

```sh
npm ci
npm run typecheck
npm test                    # includes the invariant gates (core/test/invariants.test.ts)
npm run build
npm run validate:manifests  # needs internet: calls Microsoft's manifest validation service
npm run eval:self-test      # evaluation harness scores its reference and flawed scripts correctly
```

`npm run package` repeats the build and the tests before packaging, unless you pass `--skip-build`.

## 2. Keys (once per project)

```sh
node scripts/release-keys.mjs           # writes keys/release-private.pem (0600) and keys/release-public.pem
```

- Keep `release-private.pem` off the repository (`keys/` is ignored) and out of CI logs. Store it offline or in a secrets manager.
- Publish `release-public.pem` separately from the release itself (project website, README of the distribution channel), so users can tell a genuine release from a modified one.
- If the private key may have leaked, generate a new pair, publish the new public key, and re-sign any release still in use.

## 3. Package and sign

```sh
LOAI_SIGNING_KEY=keys/release-private.pem npm run package
```

Without `LOAI_SIGNING_KEY` the release is checksummed but **not signed**, and the script says so. Do not distribute an unsigned release.

## 4. Verify

```sh
node scripts/verify-release.mjs release/local-office-ai-<version> keys/release-public.pem
```

The script checks the signature over `SHA256SUMS`, every listed file's hash, and that no unlisted file exists. It exits non-zero on any mismatch. Verify the extracted `.tar.gz` as well as the directory, on a different machine if possible.

## 5. Record what was and was not verified

Update before tagging:

- `README.md` → *Status*: version, test count, and the honest state of live Office and model evaluation.
- [acceptance.md](acceptance.md) → *Status*: regenerate with `loai-companion acceptance --out` if live records exist; otherwise state that no live runs exist.
- [model-compatibility.md](model-compatibility.md) → *Evaluated models*: regenerate with `npm run eval:matrix` if run files exist; otherwise state that no model has been evaluated.
- [limitations.md](limitations.md): add anything new that is unsupported or unverified.

A release note must not describe a capability as working in Office, or a model as compatible, unless a committed live record or evaluation run shows it.

## What users run

```sh
node scripts/verify-release.mjs <extracted dir> <published release-public.pem>   # optional, recommended
./install.sh            # macOS
./install.ps1           # Windows (PowerShell)
```

The installers check `SHA256SUMS` before installing.
