<#
.SYNOPSIS
  Runs batch Google Drive backups for users listed in the "CopyRunEligible" sheet.

.DESCRIPTION
  This script automates multi-user Drive copy operations:
    - Downloads "CopyRunEligible" as a CSV from Google Sheets using GAM.
    - If eligible users are found, invokes Copy-SingleDriveToShare.ps1 once.
    - The child script reads the same CSV, processes users, and writes results to a local results file.
    - After completion, this script imports that results file for reporting and updates the sheet.

  This version removes JSON-based inter-process communication and uses file-based coordination instead.
  All existing modules and email/reporting behavior are preserved.

.PARAMETER ChildPath
  Path to Copy-SingleDriveToShare.ps1.

.PARAMETER SummaryFrom
  Sender email used for GAM notifications.

.PARAMETER SummaryTo
  Recipient email for the consolidated summary.

.PARAMETER WhatIf
  Executes without making changes.
#>

[CmdletBinding()]
param (
  [Parameter(Mandatory = $false)]
  [ValidateNotNullOrEmpty()]
  [string]$ChildPath = "$env:GITHUB_WORKSPACE\runner\copy\Copy-SingleDriveToShare.ps1",

  [Parameter(Mandatory = $false)]
  [ValidateNotNullOrEmpty()]
  [string]$SummaryFrom = 'google.admin@sonos.com',

  [Parameter(Mandatory = $false)]
  [ValidateNotNullOrEmpty()]
  [string]$SummaryTo = 'r-google-admin@sonos.com',

  [Parameter(Mandatory = $false)]
  [switch]$WhatIf
)

$trackingDir = $env:GITHUB_WORKSPACE

# ---------- Pre-flight ----------
Import-Module "$trackingDir\runner\copy\Send-alert.psm1" -ErrorAction Stop
Import-Module "$trackingDir\runner\copy\Send-EmailReport.psm1" -ErrorAction Stop
Import-Module "$trackingDir\runner\copy\Update-GoogleSheet.psm1" -ErrorAction Stop

$gam = 'gam'
if (-not (Get-Command $gam -ErrorAction SilentlyContinue)) {
  Send-alert -Subject 'Runner failed: GAM not found' -Body 'GAM is not installed or not in PATH.'
  throw 'GAM is not installed or not in PATH; cannot continue.'
}

if (-not (Test-Path $ChildPath)) {
  Send-alert -Subject 'Runner failed: child script missing' -Body ("Child script not found: {0}" -f $ChildPath)
  throw ("Child script not found: {0}" -f $ChildPath)
}

# ---------- Download eligibility list ----------
$eligCsv = "$trackingDir\CopyRunEligible.csv"
$dlArgs = @(
  'user', $SummaryFrom,
  'get', 'drivefile', '16joJuSmXTeh-JdbFLz18hIbP6cA4_qHJTHQR5l51GE0',
  'csvsheet', 'CopyRunEligible',
  'targetfolder', "$trackingDir", 'targetname', 'CopyRunEligible.csv',
  'overwrite', 'true'
)

& $gam @dlArgs | Out-Null
if ($LASTEXITCODE -ne 0 -or -not (Test-Path $eligCsv)) {
  Send-alert -Subject 'Runner failed: sheet download' -Body ("Failed to download CopyRunEligible sheet (exit {0})." -f $LASTEXITCODE)
  exit $LASTEXITCODE
}

# ---------- Check for users ----------
$rows = Import-Csv -Path $eligCsv
$emails = $rows |
Select-Object -ExpandProperty UserEmail |
ForEach-Object { $_.Trim() } |
Where-Object { $_ } |
Sort-Object -Unique

if (-not $emails -or $emails.Count -eq 0) {
  Send-alert -Subject 'Runner notice: no users to copy' -Body ("CopyRunEligible sheet contained no users at {0}." -f (Get-Date -Format 'yyyy-MM-dd HH:mm'))
  Write-Output 'No users to process — exiting cleanly.'
  exit 0
}

# ---------- Invoke batch copy ----------
$pwshExe = 'pwsh'
$resultsFile = "$trackingDir\CopyRunResults-raw.csv"

# Ensure any previous run result is cleared
Remove-Item $resultsFile -Force -ErrorAction SilentlyContinue

$DataArgs = @(
  '-NoProfile', '-ExecutionPolicy', 'Bypass',
  '-File', (Resolve-Path $ChildPath).Path,
  '-NoEmail',
  '-SummaryFrom', $SummaryFrom,
  '-SummaryTo', $SummaryTo
)
if ($WhatIf) { $DataArgs += '-WhatIf' }

& $pwshExe @DataArgs
$exitCode = $LASTEXITCODE

if ($exitCode -ne 0) {
  Send-alert -Subject 'Runner failed: child script nonzero exit' -Body ("Copy-SingleDriveToShare.ps1 exited with code {0}." -f $exitCode)
  exit $exitCode
}

# ---------- Read results from file ----------
if (-not (Test-Path $resultsFile)) {
  Send-alert -Subject 'Runner warning: results file missing' -Body 'Child script completed but no results file was found.'
  exit 1
}

$results = Import-Csv -Path $resultsFile

# ---------- Send summary ----------
if ($WhatIf) {
  $subject = "$((Get-Date).ToString("yyyy-MM-dd")) Drive Copy Report - WhatIf Mode"
}
else {
  $subject = "$((Get-Date).ToString("yyyy-MM-dd")) Drive Copy Report"
}

Send-EmailReport -ReportPath $resultsFile -Title $subject -EmailTo $SummaryTo

# ---------- Update Google Sheet ----------
Update-GoogleSheet -DataArray $results

# ---------- Cleanup ----------
Remove-Item $resultsFile -Force -ErrorAction SilentlyContinue
Remove-Item $eligCsv -Force -ErrorAction SilentlyContinue

# ---------- Reset sheet headers ----------
$UpdateHeaders = "$trackingDir\headers.csv"
"UserEmail,FileCount,EstimatedCopyTimeMin,RotationTime,Deferred" | Out-File -FilePath $UpdateHeaders -Encoding UTF8

$updateArgs = @(
  'user', $SummaryFrom,
  'update', 'drivefile', 'id', '16joJuSmXTeh-JdbFLz18hIbP6cA4_qHJTHQR5l51GE0',
  'retainname', 'localfile', $UpdateHeaders
)
& $gam @updateArgs | Out-Null

Remove-Item $UpdateHeaders -Force -ErrorAction SilentlyContinue
exit 0