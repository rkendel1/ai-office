#!/usr/bin/env bash
# Local Office AI installer for macOS (and Linux, companion only).
# Run from the extracted release directory:  ./install.sh
set -euo pipefail

HERE="$(cd "$(dirname "$0")" && pwd)"
VERSION="$(cat "$HERE/VERSION")"
OS="$(uname -s)"

say() { printf '\033[1m==>\033[0m %s\n' "$*"; }
die() { printf '\033[31mError:\033[0m %s\n' "$*" >&2; exit 1; }

# 1. Requirements -------------------------------------------------------------
command -v node >/dev/null || die "Node.js 22.13 or newer is required (https://nodejs.org)."
node -e 'const [a,b]=process.versions.node.split(".").map(Number); process.exit(a>22||(a===22&&b>=13)?0:1)' \
  || die "Node.js $(node -v) is too old; 22.13 or newer is required."
command -v ollama >/dev/null || say "Ollama was not found. The default local model app is Atomic Chat at http://127.0.0.1:1337/v1; Ollama is optional (https://ollama.com)."

# 2. Integrity ----------------------------------------------------------------
say "Verifying release files"
( cd "$HERE" && if command -v shasum >/dev/null; then shasum -a 256 -c --quiet SHA256SUMS; else sha256sum -c --quiet SHA256SUMS; fi ) \
  || die "Checksum verification failed. Do not install this copy."
if [ -f "$HERE/SHA256SUMS.sig" ]; then say "Release is signed. Verify with scripts/verify-release.mjs and the published public key."; fi

# 3. Install files ------------------------------------------------------------
if [ "$OS" = "Darwin" ]; then
  BASE="$HOME/Library/Application Support/LocalOfficeAI"
else
  BASE="${XDG_DATA_HOME:-$HOME/.local/share}/local-office-ai"
fi
APP="$BASE/app"
say "Installing $VERSION to $APP"
rm -rf "$APP"
mkdir -p "$APP"
cp -R "$HERE/app/." "$APP/"

BIN="$HOME/.local/bin"
mkdir -p "$BIN"
cat > "$BIN/loai-companion" <<EOF
#!/usr/bin/env bash
exec node --no-warnings "$APP/companion/dist/cli.js" "\$@"
EOF
chmod +x "$BIN/loai-companion"

# 4. Local HTTPS certificate (required by Office) ------------------------------
if [ ! -f "$HOME/.office-addin-dev-certs/localhost.crt" ]; then
  say "Creating a trusted certificate for https://localhost (you may be asked for your password)"
  npx --yes office-addin-dev-certs@3 install --days 825 || say "Certificate install failed; run 'npx office-addin-dev-certs install' later."
fi

# 5. Register the add-in with Office (sideload) --------------------------------
if [ "$OS" = "Darwin" ]; then
  for app in Word Excel Powerpoint; do
    WEF="$HOME/Library/Containers/com.microsoft.$app/Data/Documents/wef"
    mkdir -p "$WEF"
    cp "$APP/addin/dist/manifest.xml" "$WEF/local-office-ai.xml"
  done
  say "Registered with Word, Excel and PowerPoint (restart them; open Home ▸ Add-ins)."

  # 6. Start at login --------------------------------------------------------
  PLIST="$HOME/Library/LaunchAgents/ai.localoffice.companion.plist"
  mkdir -p "$HOME/Library/LaunchAgents" "$BASE/logs"
  cat > "$PLIST" <<EOF
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0"><dict>
  <key>Label</key><string>ai.localoffice.companion</string>
  <key>ProgramArguments</key><array>
    <string>$(command -v node)</string><string>--no-warnings</string>
    <string>$APP/companion/dist/cli.js</string><string>start</string>
  </array>
  <key>RunAtLoad</key><true/>
  <key>KeepAlive</key><true/>
  <key>StandardOutPath</key><string>$BASE/logs/companion.log</string>
  <key>StandardErrorPath</key><string>$BASE/logs/companion.log</string>
</dict></plist>
EOF
  launchctl unload "$PLIST" 2>/dev/null || true
  launchctl load "$PLIST"
  sleep 2
  say "Companion running at https://localhost:47615 (starts at login)."
else
  say "Start the companion with: loai-companion start"
fi

cat <<EOF

Next steps
  1. Get a pairing code:     $BIN/loai-companion pair
  2. In Word: Home ▸ Add-ins ▸ Local Office AI ▸ Settings ▸ enter the code.
  3. Choose a model in Settings.
  Outlook: Get Add-ins ▸ My add-ins ▸ Add from file ▸ $APP/addin/dist/outlook-manifest.xml

EOF
