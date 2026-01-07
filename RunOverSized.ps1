<#
.SYNOPSIS
    Executes Drive copy operations for oversized users and emails summary reports.

.DESCRIPTION
    This script orchestrates per-user Google Drive copy operations using a child script
    (Copy-SingleDriveToShare.ps1). It:
      1. Downloads the OversizedUsers sheet from Google Drive.
      2. Normalizes and filters the email list.
      3. Executes the child script for each user.
      4. Aggregates JSON results from each run.
      5. Emails the summary report and updates the source sheet.

    Errors are automatically reported using Send-Alert.psm1 and Send-EmailReport.psm1.

.PARAMETER ChildPath
    Path to the child script (Copy-SingleDriveToShare.ps1).

.PARAMETER SummaryFrom
    Admin account used to retrieve and update the tracking sheet via GAM.

.PARAMETER SummaryTo
    Recipient email address for the summary report.

.PARAMETER WhatIf
    When specified, executes all logic in dry-run mode. No actual data copy is performed.

.EXAMPLE
    PS> .\Runner-DriveCopy.ps1
    # Runs the Oversized Drive copy process and emails results.

.EXAMPLE
    PS> .\Runner-DriveCopy.ps1 -WhatIf
    # Simulates the run without performing any actual copies.

.NOTES
    - Requires GAM installed and in PATH.
    - Requires Send-Alert.psm1 and Send-EmailReport.psm1 modules in .\
    - Child script must output valid JSON per user.
    - Designed for daily automation execution.
#>

[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$ChildPath = '.\Copy-SingleDriveToShare.ps1',

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$SummaryFrom = 'google.admin@sonos.com',

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$SummaryTo = 'r-google-admin@sonos.com',

    [Parameter(Mandatory = $false)]
    [switch]$WhatIf
)

# Set-Location -path $(Split-Path -Parent (Split-Path -Parent $PSScriptRoot))

# ---------------------------------------------------------------------
# 0. Pre-flight validation and module import
# ---------------------------------------------------------------------
if (-not (Get-Command -Name gam -ErrorAction SilentlyContinue)) {
    throw "GAM tool is not found. Please make sure it is installed and accessible in the system PATH."
}

try {
    Import-Module '.\Send-Alert.psm1' -ErrorAction Stop
    Import-Module '.\Send-EmailReport.psm1' -ErrorAction Stop
}
catch {
    Write-Error "Failed to import one or more required modules: $($_.Exception.Message)"
    exit 1
}

if (Test-Path '.\gam-Dir.ps1') {
    & .\gam-Dir.ps1
}

$gam = 'gam'
if (-not (Get-Command $gam -ErrorAction SilentlyContinue)) {
    Send-Alert -Subject 'Runner failed: GAM not found' -Body 'GAM is not installed or not in PATH.'
    throw 'GAM is not installed or not in PATH; cannot continue.'
}

if (-not (Test-Path $ChildPath)) {
    Send-Alert -Subject 'Runner failed: child script missing' -Body ("Child script not found: {0}" -f $ChildPath)
    throw ("Child script not found: {0}" -f $ChildPath)
}

# ---------------------------------------------------------------------
# 1. Download OversizedUsers sheet
# ---------------------------------------------------------------------
$eligCsv = '.\OversizedUsers.csv'
$sheetID = '1orGbUWkQmE7ssfHVNI_VbRh1aFyIRgfrRKJ0mJX9OPw'
$dlArgs = @(
    'user', $SummaryFrom,
    'get', 'drivefile', $sheetID,
    'csvsheet', 'OversizedUsers',
    'targetfolder', '.', 'targetname', 'OversizedUsers.csv',
    'overwrite', 'true'
)

& $gam @dlArgs | Out-Null
if ($LASTEXITCODE -ne 0 -or -not (Test-Path $eligCsv)) {
    Send-Alert -Subject 'Runner failed: sheet download' -Body ("Failed to download OversizedUsers sheet (exit {0})." -f $LASTEXITCODE)
    exit $LASTEXITCODE
}

# ---------------------------------------------------------------------
# 2. Extract and normalize user email list
# ---------------------------------------------------------------------
try {
    $rows = Import-Csv -Path $eligCsv -ErrorAction Stop
}
catch {
    Send-Alert -Subject 'Runner failed: invalid OversizedUsers sheet' -Body "Unable to parse OversizedUsers.csv: $($_.Exception.Message)"
    exit 1
}

$emails = $rows |
    Select-Object -ExpandProperty UserEmail -ErrorAction SilentlyContinue |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ } |
    Sort-Object -Unique

if (-not $emails -or $emails.Count -eq 0) {
    Send-Alert -Subject 'Runner Oversize notice: no users to copy' -Body ("OversizedUsers sheet contained no valid users. {0}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm'))
    Write-Output 'No users to process.'
    exit 0
}

# ---------------------------------------------------------------------
# 3. Batch execution: process each user via child script
# ---------------------------------------------------------------------
$pwshExe = 'powershell'  # Use 'pwsh' if running in PowerShell Core
$results = New-Object System.Collections.Generic.List[object]

foreach ($email in $emails) {
    $DataArgs = @(
        '-NoProfile', '-ExecutionPolicy', 'Bypass',
        '-File', (Resolve-Path $ChildPath).Path,
        '-Email', $email,
        '-SummaryFrom', $SummaryFrom,
        '-SummaryTo', $SummaryTo
    )
    if ($WhatIf) { $DataArgs += '-WhatIf' }

    Write-Host ("Processing user: {0}" -f $email)
    $stdout = & $pwshExe @DataArgs

    # Parse JSON output from child script
    try {
        $obj = $stdout | ConvertFrom-Json -ErrorAction Stop
    }
    catch {
        $obj = [pscustomobject]@{
            SourceEmailAddress = $email
            ExitCode            = 999
            TodayDate           = (Get-Date).ToString('yyyy-MM-dd')
            Error               = 'Invalid or non-JSON output received.'
        }
    }

    [void]$results.Add($obj)
}

# ---------------------------------------------------------------------
# 4. Prepare report subject
# ---------------------------------------------------------------------
$dateStr = (Get-Date).ToString("yyyy-MM-dd")
$subject = if ($WhatIf) {
    "$dateStr Drive Copy Report Oversize - WhatIf Mode"
}
else {
    "$dateStr Drive Copy Report Oversize"
}

# ---------------------------------------------------------------------
# 5. Aggregate and email results
# ---------------------------------------------------------------------
$reportPath = ".\CopyRunResults-raw.csv"
$results | Export-Csv -Path $reportPath -NoTypeInformation -Encoding UTF8

Send-EmailReport -ReportPath $reportPath -Title $subject -EmailTo $SummaryTo

Remove-Item -LiteralPath $reportPath -Force -ErrorAction SilentlyContinue
Remove-Item -LiteralPath $eligCsv -Force -ErrorAction SilentlyContinue

# ---------------------------------------------------------------------
# 6. Update summary sheet
# ---------------------------------------------------------------------
$updateHeaders = ".\headers.csv"
"UserEmail,FileCount,EstimatedCopyTimeMin,RotationTime" | Out-File -FilePath $updateHeaders -Encoding UTF8 -Force

$updateArgs = @(
    'user', $SummaryFrom,
    'update', 'drivefile', 'id', $sheetID,
    'retainname', 'localfile', $updateHeaders
)

& $gam @updateArgs | Out-Null
if ($LASTEXITCODE -ne 0) {
    Send-Alert -Subject 'Runner failed: summary sheet update' -Body "Could not upload summary header update (exit $LASTEXITCODE)."
}

Remove-Item -LiteralPath $updateHeaders -Force -ErrorAction SilentlyContinue

Write-Host "Runner completed successfully on $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')."
