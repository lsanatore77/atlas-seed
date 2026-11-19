<#
Uploads files\<Document Number>.docx (Word: a document's Primary Content is its editable source) for each controlled document in the seed and links it as that
document's Primary Content. Run after the seed load. Safe to run again: documents that already have
a file are skipped.

  powershell -ExecutionPolicy Bypass -File .\link-document-files.ps1 -Org atlas-seed -DryRun   (preview only)
  powershell -ExecutionPolicy Bypass -File .\link-document-files.ps1 -Org atlas-seed
#>
param(
    [Parameter(Mandatory = $true)][string]$Org,
    [switch]$DryRun
)
$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

# Runs an sf command with --json and returns its result, or stops with sf's own message.
function Invoke-Sf([string[]]$SfArgs) {
    $ErrorActionPreference = 'Continue'   # sf may print notices on stderr; only the JSON status counts
    $raw = (& sf @SfArgs --json 2>$null) | Out-String
    try { $res = $raw | ConvertFrom-Json } catch { throw "sf $($SfArgs -join ' ') returned no JSON:`n$raw" }
    if ($res.status -ne 0) { throw "sf $($SfArgs -join ' ') failed: $($res.message)" }
    return $res.result
}

$template = [IO.File]::ReadAllText((Join-Path $PSScriptRoot 'scripts\link-document-file.apex'))
$tmpFile = Join-Path ([IO.Path]::GetTempPath()) 'atlas-link-document-file.apex'
$dryText = if ($DryRun) { 'true' } else { 'false' }
$docFiles = Get-ChildItem -Path (Join-Path $PSScriptRoot 'files') -Filter '*.docx' | Sort-Object Name
Write-Host "$($docFiles.Count) files. $(if ($DryRun) { 'Preview only, nothing is changed.' } else { 'Uploading and linking...' })"
$notDone = @()
foreach ($docFile in $docFiles) {
    $fileBase64 = [Convert]::ToBase64String([IO.File]::ReadAllBytes($docFile.FullName))
    $apex = $template.Replace('__NUM__', $docFile.BaseName).Replace('__FILE__', $docFile.Name).Replace('__DRY__', $dryText).Replace('__B64__', $fileBase64)
    [IO.File]::WriteAllText($tmpFile, $apex, (New-Object Text.UTF8Encoding $false))
    # The CLI sometimes drops the connection ("fetch failed"). Retry twice, then move on to the next file.
    $run = $null
    for ($attempt = 1; ($attempt -le 3) -and (-not $run); $attempt++) {
        try { $run = Invoke-Sf @('apex', 'run', '--file', $tmpFile, '--target-org', $Org) }
        catch {
            if ($attempt -lt 3) {
                Write-Host "  $($docFile.BaseName): $($_.Exception.Message). Retrying in $($attempt * 10) seconds..."
                Start-Sleep -Seconds ($attempt * 10)
            } else {
                Write-Host "FAILED: $($docFile.BaseName): $($_.Exception.Message)"
                $notDone += $docFile.BaseName
            }
        }
    }
    if (-not $run) { continue }
    if (-not $run.success) { Write-Host "ERROR: $($docFile.BaseName): $($run.compileProblem) $($run.exceptionMessage)"; $notDone += $docFile.BaseName; continue }
    $run.logs -split "`n" | Where-Object { $_ -match 'USER_DEBUG' } | ForEach-Object { Write-Host (($_ -split '\|')[-1]).Trim() }
}
Remove-Item -Path $tmpFile -ErrorAction SilentlyContinue
if ($notDone.Count -gt 0) {
    Write-Host "Not done: $($notDone -join ', '). Run the script again: documents that already have their file are skipped."
} else {
    Write-Host 'Done.'
}
