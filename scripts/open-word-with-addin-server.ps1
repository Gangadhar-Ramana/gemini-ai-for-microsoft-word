$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$serverStarter = Join-Path $repoRoot "scripts\start-word-addin-server.ps1"

powershell.exe -NoProfile -ExecutionPolicy Bypass -File $serverStarter | Out-Null

$deadline = (Get-Date).AddSeconds(15)
do {
  $listener = Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue
  if ($listener) {
    break
  }
  Start-Sleep -Milliseconds 500
} while ((Get-Date) -lt $deadline)

$word = Get-Command WINWORD.EXE -ErrorAction SilentlyContinue
if ($word) {
  Start-Process -FilePath $word.Source
} else {
  Start-Process -FilePath "winword"
}
