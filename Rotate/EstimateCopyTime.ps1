<#
.SYNOPSIS
    Estimates Drive copy time per user based on file counts and splits users into eligible vs oversized buckets.

.DESCRIPTION
    This script:
      - Downloads the current file count sheet (Sheet #4: FileCountSheet)
      - Estimates Drive copy time per user using a realistic multiplier (~1.2 s/file)
      - Allocates users to the main run list (CopyRunEligible - Sheet #2) or oversized queue (OversizedUsers - Sheet #3)
      - Exports and uploads both CSVs
      - Appends to OversizedUsers sheet with deduplication by UserEmail
      - Moves oversized users to the Manual OU
    Optimized for My Drive → Shared Drive copy estimation under ~6 hours total window.

.PARAMETER AdminUser
    GAM admin account used for all GAM operations.

.PARAMETER UserFileCountsID
    Google Sheet ID for the FileCountSheet (Sheet #4).

.PARAMETER EligibleOutputSheetID
    Sheet ID to upload the filtered eligible users (CopyRunEligible - Sheet #2).

.PARAMETER OversizedOutputSheetID
    Sheet ID to append oversized users (OversizedUsers - Sheet #3).

.PARAMETER OutputCsvPath
    Local path to write the eligible user CSV (default: .\RunPlan.csv).

.PARAMETER OversizedCsvPath
    Local path to write the oversized user CSV (default: .\OversizedUsers.csv).

.PARAMETER MaxMinutes
    Time budget in minutes for the copy run (default: 360 min ≈ 6 hours).

.NOTES
    Sheet #4: FileCountSheet
    Sheet #2: CopyRunEligible
    Sheet #3: OversizedUsers
#>

param (
    [ValidateNotNullOrEmpty()]
    [string]$AdminUser = "google.admin@sonos.com",

    [ValidateNotNullOrEmpty()]
    [string]$UserFileCountsID = "1o4OD7SP5bCFuTaw49hwza3YZgnj_QHx2q2rZi83dALM",

    [ValidateNotNullOrEmpty()]
    [string]$EligibleOutputSheetID = "16joJuSmXTeh-JdbFLz18hIbP6cA4_qHJTHQR5l51GE0",

    [ValidateNotNullOrEmpty()]
    [string]$OversizedOutputSheetID = "1orGbUWkQmE7ssfHVNI_VbRh1aFyIRgfrRKJ0mJX9OPw",

    [ValidateNotNullOrEmpty()]
    [string]$OutputCsvPath = "$env:GITHUB_WORKSPACE\RunPlan.csv",

    [ValidateNotNullOrEmpty()]
    [string]$OversizedCsvPath = "$env:GITHUB_WORKSPACE\OversizedUsers.csv",

    [ValidateRange(0, 10000)]
    [int]$MaxMinutes = 360
)

$trackingDir = $env:GITHUB_WORKSPACE

Import-Module "$trackingDir\runner\copy\Send-Alert.psm1" -Force -ErrorAction Stop

# Estimate copy duration (minutes) based on file count.
# Calibrated for MyDrive → SharedDrive transfers:
#   - ~1.2 seconds per file average observed.
#   - 3600 seconds/hour, so (filecount * 1.2) / 60 = minutes.
function Get-EstimatedMinutesFromFileCount {
    param ([int]$FileCount)
    return [math]::Ceiling(($FileCount * 1.2) / 60)
}

# Step 1: Download the FileCount sheet (Sheet #4)
& gam user $AdminUser get drivefile $UserFileCountsID csvsheet "UserFileCounts" targetfolder "$trackingDir" targetname "UserFileCounts.csv" overwrite true | Out-Null
if ($LASTEXITCODE -ne 0) {
    Send-Alert -Subject "[EstimateCopyTime] Failed to download file count sheet" `
        -Body "Unable to retrieve $UserFileCountsID using GAM."
    exit 1
}

# Step 2: Parse CSV into structured user objects
$UserList = @()
try {
    $rows = Import-Csv "$trackingDir\UserFileCounts.csv" | Sort-Object { [datetime]$_.DateSuspended }
    foreach ($row in $rows) {
        $UserList += [PSCustomObject]@{
            UserEmail     = $row.UserEmail.Trim().ToLower()
            FileCount     = [int]$row.FileCount
            DateSuspended = $row.DateSuspended
        }
    }
}
catch {
    Send-Alert -Subject "[EstimateCopyTime] CSV Parse Error" `
        -Body "Failed to parse .\UserFileCounts.csv. $_"
    exit 1
}

# Step 3: Allocate users based on estimated copy time
$RunList = @()
$Oversized = @()
$CumulativeMinutes = 0

foreach ($user in $UserList) {
    $email = $user.UserEmail
    $fileCount = $user.FileCount
    $estMin = Get-EstimatedMinutesFromFileCount -FileCount $fileCount
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm"

    if ($estMin -gt $MaxMinutes) {
        # Single user exceeds total budget — queue for manual.
        $Oversized += [PSCustomObject]@{
            UserEmail            = $email
            FileCount            = $fileCount
            EstimatedCopyTimeMin = $estMin
            RotationTime         = $timestamp
        }
    }
    elseif (($CumulativeMinutes + $estMin) -le $MaxMinutes) {
        # Fits within available window — include in automated run.
        $RunList += [PSCustomObject]@{
            UserEmail            = $email
            FileCount            = $fileCount
            EstimatedCopyTimeMin = $estMin
            RotationTime         = $timestamp
            Deferred             = $false
        }
        $CumulativeMinutes += $estMin
    }
    else {
        # Would exceed budget — defer for next run.
        $Oversized += [PSCustomObject]@{
            UserEmail            = $email
            FileCount            = $fileCount
            EstimatedCopyTimeMin = $estMin
            RotationTime         = $timestamp
        }
    }
}

# Step 4: Export and upload eligible user list
if ($RunList) {
    try {
        $RunList | Export-Csv -Path $OutputCsvPath -NoTypeInformation -Force
    }
    catch {
        Send-Alert -Subject "[EstimateCopyTime] Failed to export RunPlan" `
            -Body "Could not write $OutputCsvPath. $_"
        exit 1
    }

    & gam user $AdminUser update drivefile id $EligibleOutputSheetID `
        retainname localfile $OutputCsvPath | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Send-Alert -Subject "[EstimateCopyTime] Upload Failed (RunPlan)" `
            -Body "GAM failed to upload $OutputCsvPath to Sheet $EligibleOutputSheetID."
        exit 2
    }
}

# Step 5: Export and append Oversized users (dedupe by UserEmail)
if ($Oversized) {
    $existingOversized = @()

    & gam user $AdminUser get drivefile $OversizedOutputSheetID csvsheet "OversizedUsers" targetfolder "$trackingDir" targetname "OversizedUsers_Existing.csv" overwrite true | Out-Null
    if ($LASTEXITCODE -eq 0 -and (Test-Path "$trackingDir\OversizedUsers_Existing.csv")) {
        try {
            $existingOversized = Import-Csv "$trackingDir\OversizedUsers_Existing.csv" -ErrorAction Stop
        }
        catch {
            Send-Alert -Subject "[EstimateCopyTime] Parse Error (OversizedUsers)" `
                -Body "Failed to parse OversizedUsers_Existing.csv. $_"
            exit 1
        }
    }

    $mergedOversized = @($existingOversized) + @($Oversized) | Sort-Object -Property UserEmail -Unique

    try {
        $mergedOversized | Export-Csv -Path $OversizedCsvPath -NoTypeInformation -Force
    }
    catch {
        Send-Alert -Subject "[EstimateCopyTime] Failed to export OversizedUsers" `
            -Body "Could not write $OversizedCsvPath. $_"
        exit 1
    }

    & gam user $AdminUser update drivefile id $OversizedOutputSheetID `
        retainname localfile $OversizedCsvPath | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Send-Alert -Subject "[EstimateCopyTime] Upload Failed (OversizedUsers)" `
            -Body "GAM failed to upload $OversizedCsvPath to Sheet $OversizedOutputSheetID."
        exit 3
    }

    Remove-Item "$trackingDir\OversizedUsers_Existing.csv" -Force -ErrorAction SilentlyContinue

    # Step 6: Move oversized users to Manual OU
    foreach ($entry in $Oversized) {
        $email = $entry.UserEmail
        & gam update user $email org '/Sonos Inc/Copy/Manual'
        if ($LASTEXITCODE -ne 0) {
            Send-Alert -Subject "[EstimateCopyTime] Failed to move $email to Manual OU" `
                -Body "GAM failed to update org for oversized user: $email"
        }
        Start-Sleep -Seconds 5
    }
}
