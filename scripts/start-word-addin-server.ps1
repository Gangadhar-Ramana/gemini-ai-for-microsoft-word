$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$logDir = Join-Path $repoRoot "logs"
$outLog = Join-Path $logDir "word-addin-server.out.log"
$errLog = Join-Path $logDir "word-addin-server.err.log"

New-Item -ItemType Directory -Force -Path $logDir | Out-Null

$existing = Get-NetTCPConnection -LocalPort 3000 -State Listen -ErrorAction SilentlyContinue
if ($existing) {
  "[$(Get-Date -Format o)] Port 3000 is already listening. Existing PID(s): $($existing.OwningProcess -join ', ')" |
    Out-File -FilePath $outLog -Append -Encoding utf8
  exit 0
}

$node = Get-Command node.exe -ErrorAction Stop
$serverScript = Join-Path $repoRoot "scripts\serve-dist-https.js"
if (-not (Test-Path $serverScript)) {
  throw "Cannot find local server script at $serverScript."
}

Start-Process `
  -FilePath $node.Source `
  -ArgumentList @($serverScript) `
  -WorkingDirectory $repoRoot `
  -RedirectStandardOutput $outLog `
  -RedirectStandardError $errLog `
  -WindowStyle Hidden
