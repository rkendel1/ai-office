# Remove Local Office AI. Pass -Purge to also delete settings, history and pairing data.
param([switch]$Purge)
$Base = Join-Path $env:LOCALAPPDATA "LocalOfficeAI"
Get-CimInstance Win32_Process -Filter "Name='node.exe'" | Where-Object { $_.CommandLine -like "*LocalOfficeAI*cli.js*" } | ForEach-Object { Stop-Process -Id $_.ProcessId -Force }
Remove-ItemProperty -Path "HKCU:\Software\Microsoft\Office\16.0\WEF\Developer" -Name "5602660c-37bc-4809-8bcc-c5d3353356ae" -ErrorAction SilentlyContinue
Remove-Item -Force (Join-Path ([Environment]::GetFolderPath("Startup")) "LocalOfficeAI-companion.vbs") -ErrorAction SilentlyContinue
Remove-Item -Recurse -Force (Join-Path $Base "app"), (Join-Path $Base "loai-companion.cmd") -ErrorAction SilentlyContinue
if ($Purge) {
  Remove-Item -Recurse -Force $Base -ErrorAction SilentlyContinue
  Remove-Item -Recurse -Force (Join-Path $env:APPDATA "LocalOfficeAI") -ErrorAction SilentlyContinue
  Write-Host "Removed application and all local data."
} else {
  Write-Host "Removed application. Settings and history were kept (use -Purge to delete them)."
}
Write-Host "Remove the add-in from Outlook via Get Add-ins > My add-ins if you installed it there."
