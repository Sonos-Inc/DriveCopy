<#
.SYNOPSIS
    Performs automated Shared Drive rotation based on usage thresholds.

.DESCRIPTION
    This script checks Google Shared Drive utilization via a tracking sheet.
    If usage exceeds configured thresholds, it:
      1. Creates a new Shared Drive.
      2. Updates the Google Sheet with the new drive entry.
      3. Grants organizer ACLs to designated admin users.
    All key events and errors trigger email alerts via Send-Alert.psm1.

.PARAMETER AdminUser
    The Google Workspace admin user for GAM operations.
    Default: google.admin@sonos.com

.PARAMETER DocID
    The Google Sheet document ID that tracks Shared Drives.

.PARAMETER SheetName
    The tab name within the sheet that contains drive data.

.EXAMPLE
    PS> .\RotateBackup.ps1 -AdminUser "google.admin@sonos.com" `
        -DocID "1BbBdh3gDW4jcnLn9Alef1KHxoKMBeYG0h4f7Gzksnpg" `
        -SheetName "CountParentFolder"

.NOTES
    - Requires GAM installed and available in PATH.
    - Requires Send-Alert.psm1 module in .\
    - All GAM operations validated via $LASTEXITCODE.
    - Idempotent by design — safe to re-run.
#>
[CmdletBinding()]
param (
    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$AdminUser = "google.admin@sonos.com",

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$DocID = "1BbBdh3gDW4jcnLn9Alef1KHxoKMBeYG0h4f7Gzksnpg",

    [Parameter(Mandatory = $false)]
    [ValidateNotNullOrEmpty()]
    [string]$SheetName = "CountParentFolder"
)

# Set-Location -path $(Split-Path -Parent (Split-Path -Parent $PSScriptRoot))

try {
    # ---------------------------------------------------------------------
    # 0. Import dependencies and validate environment
    # ---------------------------------------------------------------------
    Import-Module "$env:GITHUB_WORKSPACE\Send-Alert.psm1" -Force -ErrorAction Stop

    $trackingDir = $env:GITHUB_WORKSPACE   # or Join-Path $PSScriptRoot '..\..' etc, as you prefer
    $csvPath = Join-Path $trackingDir 'CountParentFolder.csv'


    if (-not (Get-Command -Name gam -ErrorAction SilentlyContinue)) {
        Send-Alert -Subject "[RotateBackup] GAM not found" -Body "GAM not found in PATH. Script aborted."
        throw "GAM tool not found."
    }

    # ---------------------------------------------------------------------
    # 1. Download tracking sheet
    # ---------------------------------------------------------------------
    $downloadCmd = @(
        'gam', 'user', $AdminUser,
        'get', 'drivefile', $DocID,
        'csvsheet', $SheetName,
        'targetfolder', $trackingDir,
        'targetname', 'CountParentFolder.csv',
        'overwrite', 'true'
    )

    & $downloadCmd[0] @($downloadCmd[1..($downloadCmd.Length - 1)]) | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Send-Alert -Subject "[RotateBackup] Download failed" `
            -Body "Unable to download $SheetName from document ID $DocID."
        throw "Sheet download failed (ExitCode=$LASTEXITCODE)."
    }

    if (-not (Test-Path $csvPath)) {
        Send-Alert -Subject "[RotateBackup] CSV missing" -Body "CountParentFolder.csv not found after GAM execution."
        throw "Tracking CSV not found."
    }

    $DriveTable = @()
    try { $DriveTable = Import-Csv "$trackingDir\CountParentFolder.csv" -ErrorAction Stop }
    catch {
        Send-Alert -Subject "[RotateBackup] CSV import failed" -Body "Could not parse CountParentFolder.csv: $($_.Exception.Message)"
        throw
    }

    # ---------------------------------------------------------------------
    # 2. Evaluate usage via CountParentFolder.ps1
    # ---------------------------------------------------------------------
    $CountResult = & "$env:GITHUB_WORKSPACE\Rotate\CountParentFolder.ps1" -AdminUser $AdminUser -DocID $DocID

    if (
        -not $CountResult -or
        -not $CountResult.PSObject.Properties.Match('ItemPercent') -or
        -not $CountResult.PSObject.Properties.Match('FolderPercent') -or
        $CountResult.ItemPercent -eq -1 -or
        $CountResult.FolderPercent -eq -1
    ) {
        Send-Alert -Subject "[RotateBackup] Count failed" `
            -Body "Invalid usage values returned by CountParentFolder."
        return
    }

    # ---------------------------------------------------------------------
    # 3. Evaluate thresholds
    # ---------------------------------------------------------------------
    if ($CountResult.ItemPercent -lt 80 -and $CountResult.FolderPercent -lt 80) {
        Send-Alert -Subject "[RotateBackup] No Rotation Needed" `
            -Body "Usage within limits. Items=$($CountResult.ItemPercent)%, Folders=$($CountResult.FolderPercent)%"
        Write-Host "No rotation required."
        return
    }

    Write-Host "Threshold exceeded. Starting rotation..."

    # ---------------------------------------------------------------------
    # 4. Create new Shared Drive
    # ---------------------------------------------------------------------
    $BaseName = "Legacydrivebackup"
    $ExistingNames = $DriveTable | Select-Object -ExpandProperty DriveName -ErrorAction SilentlyContinue

    $nextSuffix = ($ExistingNames | ForEach-Object {
            if ($_ -match '^Legacydrivebackup(\d+)?$') {
                if ($matches[1]) { [int]$matches[1] } else { 1 }
            }
        } | Sort-Object -Descending | Select-Object -First 1) + 1

    $newDriveName = if ($nextSuffix -eq 1) { $BaseName } else { "$BaseName$nextSuffix" }

    $createOut = & gam user $AdminUser create teamdrive $newDriveName asadmin
    if ($LASTEXITCODE -ne 0) {
        Send-Alert -Subject "[RotateBackup] Create failed" -Body "Failed to create Shared Drive $newDriveName. GAM output: $createOut"
        return
    }

    $NewDriveID = ($createOut | Select-String -Pattern 'id:\s*([^\s,]+)').Matches[0].Groups[1].Value
    if (-not $NewDriveID) {
        Send-Alert -Subject "[RotateBackup] Missing Drive ID" -Body "GAM output did not contain a Drive ID for $newDriveName."
        throw "Drive ID parse failure."
    }

    & gam update teamdrive $NewDriveID driveMembersOnly False | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Send-Alert -Subject "[RotateBackup] Drive update failed" -Body "Failed to set driveMembersOnly=False for $NewDriveID."
        return
    }

    # ---------------------------------------------------------------------
    # 5. Update tracking sheet
    # ---------------------------------------------------------------------
    $UpdatedTable = $DriveTable | ForEach-Object {
        $_.IsFull = 'TRUE'
        $_
    }

    $NewEntry = [PSCustomObject]@{
        DriveName   = $newDriveName
        DriveID     = $NewDriveID
        IsFull      = 'FALSE'
        LastUpdated = (Get-Date).ToString("yyyy-MM-dd HH:mm")
    }

    $FinalTable = @()
    $FinalTable += $UpdatedTable
    $FinalTable += $NewEntry
    $FinalTable | Export-Csv -Path "$trackingDir\NewDriveRow.csv" -NoTypeInformation -Encoding UTF8 -Force

    $uploadCmd = @(
        'gam', 'user', $AdminUser,
        'update', 'drivefile', 'id', $DocID,
        'retainname', 'localfile', "$trackingDir\NewDriveRow.csv"
    )

    & $uploadCmd[0] @($uploadCmd[1..($uploadCmd.Length - 1)]) | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Send-Alert -Subject "[RotateBackup] Sheet update failed" -Body "Unable to upload new entry for $newDriveName."
        return
    }

    Send-Alert -Subject "[RotateBackup] New Drive Created" `
        -Body "$newDriveName successfully created and added to tracking sheet."

    # ---------------------------------------------------------------------
    # 6. Apply organizer ACLs
    # ---------------------------------------------------------------------
    $ACLUsers = @("it-support@sonos.com", "google.admin@sonos.com", "dan.casmas@sonos.com")
    foreach ($user in $ACLUsers) {
        $gamArgs = @('add', 'drivefileacl', $NewDriveID, 'user', $user, 'role', 'organizer')
        & gam @gamArgs | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Send-Alert -Subject "[RotateBackup] ACL failed" -Body "Could not assign organizer role to $user on $NewDriveID"
            return
        }
    }

    Write-Host "Rotation completed successfully for $newDriveName."
}
catch {
    Send-Alert -Subject "[RotateBackup] Fatal error" -Body "Unexpected error: $($_.Exception.Message)"
    Write-Error "Fatal error: $($_.Exception.Message)"
}
