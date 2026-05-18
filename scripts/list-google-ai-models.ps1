param(
  [string]$ApiKey = $env:GEMINI_API_KEY,
  [string]$OutFile
)

$ErrorActionPreference = 'Stop'

if ([string]::IsNullOrWhiteSpace($ApiKey)) {
  throw 'Provide the API key with -ApiKey or set GEMINI_API_KEY. The script never prints the key.'
}

$uri = "https://generativelanguage.googleapis.com/v1beta/models?key=$([uri]::EscapeDataString($ApiKey))"
$response = Invoke-RestMethod -Uri $uri -Method Get

$models = @($response.models) | ForEach-Object {
  $name = [string]$_.name -replace '^models/', ''
  $methods = @($_.supportedGenerationMethods)
  [pscustomobject]@{
    id = $name
    displayName = $_.displayName
    usableInOfficialAddin = $methods -contains 'generateContent'
    usableViaLiveAdapter = $methods -contains 'bidiGenerateContent'
    supportedMethods = ($methods -join ', ')
  }
} | Where-Object {
  $_.usableInOfficialAddin -or $_.usableViaLiveAdapter
} | Sort-Object `
  @{ Expression = { $_.usableInOfficialAddin }; Descending = $true }, `
  @{ Expression = { $_.usableViaLiveAdapter }; Descending = $true }, `
  @{ Expression = { $_.id }; Descending = $false }

if ($OutFile) {
  $models | ConvertTo-Json -Depth 4 | Set-Content -LiteralPath $OutFile -Encoding UTF8
}

$models | Format-Table -AutoSize
