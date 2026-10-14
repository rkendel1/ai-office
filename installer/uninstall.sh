#!/usr/bin/env bash
# Remove Local Office AI. Pass --purge to also delete settings, history and pairing data.
set -euo pipefail
OS="$(uname -s)"
if [ "$OS" = "Darwin" ]; then
  BASE="$HOME/Library/Application Support/LocalOfficeAI"
  PLIST="$HOME/Library/LaunchAgents/ai.localoffice.companion.plist"
  [ -f "$PLIST" ] && { launchctl unload "$PLIST" 2>/dev/null || true; rm -f "$PLIST"; }
  for app in Word Excel Powerpoint; do rm -f "$HOME/Library/Containers/com.microsoft.$app/Data/Documents/wef/local-office-ai.xml"; done
else
  BASE="${XDG_DATA_HOME:-$HOME/.local/share}/local-office-ai"
fi
rm -rf "$BASE/app" "$HOME/.local/bin/loai-companion"
if [ "${1:-}" = "--purge" ]; then
  rm -rf "$BASE"
  node -e 'const p=require("path"),os=require("os");console.log(process.platform==="darwin"?p.join(os.homedir(),"Library/Application Support/LocalOfficeAI"):p.join(process.env.XDG_CONFIG_HOME||p.join(os.homedir(),".config"),"local-office-ai"))' | xargs -I{} rm -rf "{}"
  echo "Removed application and all local data."
else
  echo "Removed application. Settings and history were kept (use --purge to delete them)."
fi
echo "Remove the add-in from Outlook via Get Add-ins ▸ My add-ins if you installed it there."
