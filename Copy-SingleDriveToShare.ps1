<#
.SYNOPSIS
    Copies one or more users’ Google Drive to a designated Shared Drive folder.

.DESCRIPTION
    This script handles the backup of user My Drives into a Shared Drive location.

    Modes:
      1. Single-user mode (Email provided):
         - Validates GAM availability and required modules.
         - Determines the destination Shared Drive folder.
         - Creates a new backup folder for the user.
         - Adjusts the user’s state (archive/OU) via Set-UserCopyState.
         - Applies temporary ACLs required for copy.
         - Performs a full recursive copy with GAM.
         - Removes temporary ACLs and finalizes the user’s state.
         - Optionally grants the manager access and notifies them.
         - Produces a per-user result object.
         - If -NoEmail is used, returns a compact JSON result (backwards compatible).
         - If -NoEmail is not used, updates the tracking sheet and sends a summary email.

      2. Batch mode (Email omitted):
         - Reads 'CopyRunEligible.csv' from the current directory.
         - For each row with a non-empty UserEmail, performs the same copy workflow.
         - Aggregates all per-user result objects into 'CopyRunResults-raw.csv'.
         - Intended to be invoked by a parent runner which will handle reporting and sheet updates.

.PARAMETER Email
    User email address to process. When omitted, the script runs in batch mode using CopyRunEligible.csv.

.PARAMETER AdminEmail
    Admin account used for GAM operations.

.PARAMETER SummaryFrom
    Sender address for notifications and summaries.

.PARAMETER SummaryTo
    Recipient for summaries in single-user mode (when -NoEmail is not used).

.PARAMETER WhatIf
    Executes without making changes, for validation.

.PARAMETER NoEmail
    When used in single-user mode, suppresses per-user emails and instead emits a JSON result.
    In batch mode, per-user emails are suppressed; results are written to 'CopyRunResults-raw.csv'.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [ValidatePattern('^[a-zA-Z0-9](?:[a-zA-Z0-9._%+-]*[a-zA-Z0-9])?@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$')]
    [string]$Email,

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$AdminEmail = 'google.admin@sonos.com',

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$SummaryFrom = "google.admin@sonos.com",

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$SummaryTo = "r-google-admin@sonos.com",

    [Parameter(Mandatory = $false)]
    [switch]$WhatIf,

    [Parameter(Mandatory = $false)]
    [switch]$NoEmail
)

$trackingDir = $env:GITHUB_WORKSPACE

if (-not ($WhatIf -and $NoEmail)) {
    Import-Module $trackingDir\runner\copy\ElapsedTime.psm1 -ErrorAction SilentlyContinue
    ElapsedTime -ErrorAction SilentlyContinue
}

# Normalize inputs
if ($Email) { $Email = $Email.Trim() }
$AdminEmail = $AdminEmail.Trim()
$SummaryFrom = $SummaryFrom.Trim()
$SummaryTo = $SummaryTo.Trim()

$gamExe = 'gam'

# Pre-flight
if (-not (Get-Command $gamExe -ErrorAction SilentlyContinue)) {
    throw 'GAM is not installed or not in your PATH.'
}

# Imports and capacity gate
if (-not $WhatIf) {
    try {
        Import-Module "$trackingDir\runner\copy\Update-GoogleSheet.psm1" -ErrorAction Stop
        Import-Module "$trackingDir\runner\copy\rotate\Get-TotalRemainMain.psm1" -ErrorAction Stop
        Import-Module "$trackingDir\runner\copy\Send-alert.psm1" -ErrorAction Stop
        Import-Module "$trackingDir\runner\copy\SendManagerEmail.psm1" -ErrorAction Stop
        Import-Module "$trackingDir\runner\copy\Set-UserCopyState.psm1" -ErrorAction Stop
        Import-Module "$trackingDir\runner\copy\Get-LegacyDrive.psm1" -ErrorAction Stop
    }
    catch { throw $($_.Exception.Message) }

    $total = Get-TotalRemainMain
    if (-not $total -or $total -le 0) { Send-alert 'TotalRemainMain is zero or undefined.'; exit 1 }
}
else {
    Write-Host 'Running in WhatIf mode — no changes will be made.'
}

$LegacyFolderID = Get-LegacyDrive | Select-Object -ExpandProperty ID -ErrorAction SilentlyContinue
if ([string]::IsNullOrWhiteSpace($LegacyFolderID)) { Send-alert "Failed to retrieve LegacyFolderID." }

function Invoke-DriveCopyForUser {
    param (
        [Parameter(Mandatory = $true)]
        [string]$UserEmail
    )

    # Manager lookup via JSON
    $ManagerEmail = $null
    $ManagerActive = $false
    $ManagerEmailStatus = $false
    $DaysBetween = $null
    $lastLogin = $null

    $mgrInfoJson = & $gamExe info user $UserEmail formatjson 2>$null 
    if ($LASTEXITCODE -ne 0) { Send-alert "Failed to retrieve user info for $UserEmail (exit $LASTEXITCODE)" }
    try {
        $mgrInfoObj = $mgrInfoJson | ConvertFrom-Json -ErrorAction Stop
        $mgrRelation = $mgrInfoObj.relations | Where-Object { $_.type -eq "manager" }
        $LastLoginTime = $mgrInfoObj.lastlogintime

        if ($LastLoginTime -and $LastLoginTime -ne 'Never') {
            $lastLogin = (Get-Date $LastLoginTime).ToString('yyyy-MM-dd')
            $DaysBetween = (New-TimeSpan (Get-Date $LastLoginTime) (Get-Date)).Days 
            if ($mgrRelation) { $ManagerEmail = $mgrRelation.value.Trim() }
        }

    }
    catch { 
        if (-not $NoEmail) { Write-Warning "Failed to parse manager relation JSON: $($_.Exception.Message)" } 
    }

    # Prep names/vars
    $BackupFolderName = "Backup_From_$UserEmail"
    $BackupFolderID = $null
    $WebLink = $null
    $Succeeded = $false
    $ExitCode = -1

    # Remove any existing same-named folder (best-effort)
    $existing = & $gamExe user $AdminEmail show filelist select teamdriveid $LegacyFolderID query:"name = '$BackupFolderName' and '$LegacyFolderID' in parents and mimeType = 'application/vnd.google-apps.folder' and trashed = false" | ConvertFrom-Csv -ErrorAction SilentlyContinue
    if ($LASTEXITCODE -eq 0 -and $existing -and ($existing.name -match $BackupFolderName)) {
        & $gamExe delete shareddrive $($existing.webViewLink -replace '^https:\/\/drive\.google\.com\/drive\/folders\/', '') allowitemdeletion asadmin | Out-Null -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 3
    }

    # Create destination folder
    if ($WhatIf) {
        $BackupFolderID = '<SIMULATED_ID>'
        $WebLink = '<SIMULATED_LINK>'
    }
    else {
        $BackupFolderID = (& $gamExe user $AdminEmail create drivefile drivefilename $BackupFolderName mimetype gfolder parentid $LegacyFolderID inheritedpermissionsdisabled true returnidonly)
        if ($LASTEXITCODE -ne 0 -or [string]::IsNullOrWhiteSpace($BackupFolderID)) { Send-alert "Failed to create destination folder (exit $LASTEXITCODE)" }
        $WebLink = "https://drive.google.com/drive/folders/$BackupFolderID"
    }

    # Pre-copy state (unarchive + staging OU via helper)
    if (-not $WhatIf) {
        try { Set-UserCopyState -Email $UserEmail | Out-Null -ErrorAction SilentlyContinue } 
        catch { Send-alert "Failed during pre-copy user state change: $($_.Exception.Message)" }
    }

    # TEMP ACL on Shared Drive ROOT
    if (-not $WhatIf) {
        $gamArgs = @("user", "$AdminEmail", "add", "drivefileacl", "$LegacyFolderID", "user", "$UserEmail", "role", "manager")
        & $gamExe @gamArgs | Out-Null -ErrorAction SilentlyContinue
        if ($LASTEXITCODE -ne 0) {
            Send-alert "Temporary ACL on Legacy root ($LegacyFolderID) for $UserEmail failed to add (exit $LASTEXITCODE)"
        }
    }

    function Invoke-DriveCopy {
        & $gamExe redirect stdout "$BackupFolderName.txt" multiprocess redirect stderr stdout user $UserEmail copy drivefile root recursive teamdriveparentid $BackupFolderID mergewithparent true copyfilepermissions false copysubfolderpermissions false copytopfolderpermissions false duplicatefiles duplicatename duplicatefolders duplicatename copiedshortcutspointtocopiedfiles true newfilename $UserEmail
        return $LASTEXITCODE
    }

    if ($WhatIf) {
        $ExitCode = 0
        $Succeeded = $true
    }
    else {
        $ExitCode = Invoke-DriveCopy
        if ($ExitCode -in 0, 50) { $Succeeded = $true }
    }

    # Remove TEMP ACL from Shared Drive ROOT
    if (-not $WhatIf) {
        $gamArgs = @("user", "$AdminEmail", "delete", "drivefileacl", "$LegacyFolderID", "$UserEmail")
        & $gamExe @gamArgs | Out-Null -ErrorAction SilentlyContinue 
        if ($LASTEXITCODE -ne 0 -and -not $NoEmail) {
            Write-Warning "Temporary ACL on Legacy root ($LegacyFolderID) for $UserEmail may not have been removed (exit $LASTEXITCODE)"
        }
    }

    # Manager access/notify (only on success and active manager)
    if ($ManagerEmail -and -not $WhatIf -and $Succeeded) {
        $mgrStatus = & $gamExe info user $ManagerEmail 2>$null | Select-String "Account Suspended:\s+false"
        if ($LASTEXITCODE -eq 0 -and $mgrStatus) {
            $ManagerActive = $true
            $gamArgs = @("user", "$AdminEmail", "add", "drivefileacl", "$BackupFolderID", "user", "$ManagerEmail", "role", "contentManager")
            & $gamExe @gamArgs | Out-Null -ErrorAction SilentlyContinue
            if ($LASTEXITCODE -eq 0 -and (Get-Command -Name SendManagerEmail -ErrorAction SilentlyContinue)) {
                $ManagerEmailStatus = SendManagerEmail -Manager $ManagerEmail -UserEmail $UserEmail -WebURL $WebLink -ErrorAction SilentlyContinue
            }
            else { $ManagerEmailStatus = $false }
        }
        else { $ManagerEmailStatus = $false }
    }

    # Finalize user state
    if (-not $WhatIf) {
        try {
            if ($ExitCode -notin (0, 50)) { Set-UserCopyState -Email $UserEmail -Fail | Out-Null -ErrorAction SilentlyContinue }
            else { Set-UserCopyState -Email $UserEmail -Back | Out-Null -ErrorAction SilentlyContinue }
        }
        catch { Send-alert "Failed during finalize user state change: $($_.Exception.Message)" }
    }

    # Build result object
    $resultObj = [pscustomobject]@{
        SourceEmailAddress = $UserEmail.Trim()
        ManagerEmail       = if ($ManagerEmail) { $ManagerEmail.Trim() } else { $null }
        ManagerSuspended   = if ($ManagerEmail) { [bool](-not $ManagerActive) } else { $null }
        BackupFolderID     = $BackupFolderID
        ExitCode           = [int]$ExitCode
        DaysSuspended      = $DaysBetween
        DateSuspended      = $lastLogin
        TodayDate          = (Get-Date).ToString('yyyy-MM-dd')
        WebLink            = $WebLink
        ManagerEmailSent   = [bool]$ManagerEmailStatus
    }

    # Cleanup log file for this user
    Remove-Item "$BackupFolderName.txt" -Force -ErrorAction SilentlyContinue

    return $resultObj
}

# ---------- Mode selection: batch vs single-user ----------

# Batch mode: Email not provided, use CopyRunEligible.csv and write CopyRunResults-raw.csv
if ([string]::IsNullOrWhiteSpace($Email)) {
    $eligCsv = "$trackingDir\CopyRunEligible.csv"
    if (-not (Test-Path $eligCsv)) {
        Send-alert "Batch mode requested but CopyRunEligible.csv was not found."
        exit 1
    }

    $rows = Import-Csv -Path $eligCsv
    $emails = $rows |
    Select-Object -ExpandProperty UserEmail |
    ForEach-Object { $_.Trim() } |
    Where-Object { $_ } |
    Sort-Object -Unique

    if (-not $emails -or $emails.Count -eq 0) {
        Write-Output 'No users to process in CopyRunEligible.csv.'
        exit 0
    }

    $allResults = New-Object System.Collections.Generic.List[object]

    foreach ($user in $emails) {
        $result = Invoke-DriveCopyForUser -UserEmail $user
        [void]$allResults.Add($result)
    }

    $resultsFile = "$trackingDir\CopyRunResults-raw.csv"
    $allResults | Export-Csv -Path $resultsFile -NoTypeInformation -Encoding UTF8

    # In batch mode with -NoEmail, do not send per-user emails or update sheets here.
    # Parent runner handles reporting and Update-GoogleSheet.
    exit 0
}

# Single-user mode: Email provided
$resultObjSingle = Invoke-DriveCopyForUser -UserEmail $Email

if ($NoEmail) {
    # Backwards-compatible JSON return for single-user mode
    $json = $resultObjSingle | ConvertTo-Json -Compress
    Write-Output $json
    exit 0
}

# Upload to tracking sheet in single-user, email-enabled mode
if (-not $WhatIf -and -not $NoEmail) {
    Update-GoogleSheet -DataArray $resultObjSingle
}

# Per-user email (HTML summary) in single-user, email-enabled mode
if (-not $WhatIf) {
    $htmlBody = "<pre>$($resultObjSingle | Format-List | Out-String)</pre>"
    $gamArgs = @(
        "user", "$SummaryFrom",
        "sendemail",
        "to", "$SummaryTo",
        "subject", "Drive backup complete for $Email",
        "message", "$htmlBody",
        "html"
    )
    & $gamExe @gamArgs | Out-Null
    if ($LASTEXITCODE -ne 0) { Send-alert "Failed to send summary email (exit $LASTEXITCODE)" }

    Send-Alert -Body (ElapsedTime -AsString) -Subject "ElapsedTime"
}