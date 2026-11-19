<#
One-time cleanup for a scratch org loaded before the inspection redesign (2026-11-19): deletes the four seed
inspections and the three IC-REC Inspection Criteria documents (with their files), so the next load recreates
the documents as Draft with their specifications. Not needed for a new scratch org.

  powershell -ExecutionPolicy Bypass -File .\cleanup-step3.ps1 -Org atlas-seed -DryRun   (preview only)
  powershell -ExecutionPolicy Bypass -File .\cleanup-step3.ps1 -Org atlas-seed
#>
param(
    [Parameter(Mandatory = $true)][string]$Org,
    [switch]$DryRun
)
$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

$template = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'scripts\cleanup-step3.apex'))
$tmpFile = Join-Path ([IO.Path]::GetTempPath()) 'atlas-cleanup-step3.apex'
$apex = $template.Replace('__DRY__', $(if ($DryRun) { 'true' } else { 'false' }))
[IO.File]::WriteAllText($tmpFile, $apex, (New-Object Text.UTF8Encoding $false))
Write-Host $(if ($DryRun) { 'Preview only, nothing is changed.' } else { 'Deleting...' })

$ErrorActionPreference = 'Continue'   # sf may print notices on stderr; only the JSON status counts
$raw = (& sf apex run --file $tmpFile --target-org $Org --json 2>$null) | Out-String
Remove-Item -Path $tmpFile -ErrorAction SilentlyContinue
try { $res = $raw | ConvertFrom-Json } catch { throw "sf apex run returned no JSON:`n$raw" }
if ($res.status -ne 0) { throw "sf apex run failed: $($res.message)" }
if (-not $res.result.success) { throw "Apex error: $($res.result.compileProblem) $($res.result.exceptionMessage)" }
$res.result.logs -split "`n" | Where-Object { $_ -match 'USER_DEBUG' } | ForEach-Object { Write-Host (($_ -split '\|')[-1]).Trim() }
