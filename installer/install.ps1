# Local Office AI installer for Windows.
# Run from the extracted release folder:  powershell -ExecutionPolicy Bypass -File .\install.ps1
$ErrorActionPreference = "Stop"
$Here = Split-Path -Parent $MyInvocation.MyCommand.Path
$Version = (Get-Content "$Here\VERSION").Trim()

function Say($m) { Write-Host "==> $m" -ForegroundColor Cyan }
function Die($m) { Write-Host "Error: $m" -ForegroundColor Red; exit 1 }

# 1. Requirements
$node = Get-Command node -ErrorAction SilentlyContinue
if (-not $node) { Die "Node.js 22.13 or newer is required (https://nodejs.org)." }
& node -e 'const [a,b]=process.versions.node.split(".").map(Number); process.exit(a>22||(a===22&&b>=13)?0:1)'
if ($LASTEXITCODE -ne 0) { Die "Node.js is too old; 22.13 or newer is required." }
if (-not (Get-Command ollama -ErrorAction SilentlyContinue)) { Say "Ollama was not found. The default local model app is Atomic Chat at http://127.0.0.1:1337/v1; Ollama is optional (https://ollama.com)." }

# 2. Integrity
Say "Verifying release files"
foreach ($line in Get-Content "$Here\SHA256SUMS") {
  $hash, $rel = $line -split "  ", 2
  $file = Join-Path $Here ($rel -replace "/", "\")
  if (-not (Test-Path $file)) { Die "Missing file $rel" }
  if ((Get-FileHash -Algorithm SHA256 $file).Hash.ToLower() -ne $hash) { Die "Checksum mismatch for $rel. Do not install this copy." }
}

# 3. Install files
$Base = Join-Path $env:LOCALAPPDATA "LocalOfficeAI"
$App = Join-Path $Base "app"
Say "Installing $Version to $App"
if (Test-Path $App) { Remove-Item -Recurse -Force $App }
New-Item -ItemType Directory -Force -Path $App | Out-Null
Copy-Item -Recurse -Force "$Here\app\*" $App
$Cli = Join-Path $App "companion\dist\cli.js"
Set-Content -Path (Join-Path $Base "loai-companion.cmd") -Value "@node --no-warnings `"$Cli`" %*"

# 4. Local HTTPS certificate
if (-not (Test-Path "$env:USERPROFILE\.office-addin-dev-certs\localhost.crt")) {
  Say "Creating a trusted certificate for https://localhost (Windows will ask to trust it)"
  & npx --yes office-addin-dev-certs@3 install --days 825
}

# 5. Register with Word, Excel, PowerPoint (developer sideload registry key)
$Manifest = Join-Path $App "addin\dist\manifest.xml"
$Key = "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer"
New-Item -Path $Key -Force | Out-Null
Set-ItemProperty -Path $Key -Name "5602660c-37bc-4809-8bcc-c5d3353356ae" -Value $Manifest
Say "Registered with Word, Excel and PowerPoint (restart them; open Home > Add-ins)."

# 6. Start at sign-in (hidden)
$Startup = [Environment]::GetFolderPath("Startup")
$Vbs = Join-Path $Startup "LocalOfficeAI-companion.vbs"
Set-Content -Path $Vbs -Value "CreateObject(""WScript.Shell"").Run ""node --no-warnings """"$Cli"""" start"", 0, False"
Start-Process -WindowStyle Hidden -FilePath "node" -ArgumentList @("--no-warnings", "`"$Cli`"", "start")
Start-Sleep -Seconds 2
Say "Companion running at https://localhost:47615 (starts at sign-in)."

Write-Host ""
Write-Host "Next steps"
Write-Host "  1. Get a pairing code:   $Base\loai-companion.cmd pair"
Write-Host "  2. In Word: Home > Add-ins > Local Office AI > Settings > enter the code."
Write-Host "  3. Choose a model in Settings."
Write-Host "  Outlook: Get Add-ins > My add-ins > Add from file > $App\addin\dist\outlook-manifest.xml"
