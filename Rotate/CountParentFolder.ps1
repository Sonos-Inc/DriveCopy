<#
.SYNOPSIS
    Counts items/folders in the current shared drive and estimates projected usage after next batch of suspended user Drives.
.DESCRIPTION
    See original description.
.RETURNS
    [PSCustomObject] with ItemPercent and FolderPercent
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
    [ValidateRange(1, [int]::MaxValue)]
    [int]$ItemLimit = 500000,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, [int]::MaxValue)]
    [int]$FolderLimit = 500000
)

# Set-Location -path $(Split-Path -Parent (Split-Path -Parent $PSScriptRoot))

$CTError = [PSCustomObject]@{
    ItemPercent   = -1
    FolderPercent = -1
}

try {
    Import-Module "$env:GITHUB_WORKSPACE\Send-Alert.psm1" -Force -ErrorAction Stop
    Import-Module "$env:GITHUB_WORKSPACE\rotate\Get-LongTermSuspendedAccounts.psm1" -Force -ErrorAction Stop

    $trackingDir = $env:GITHUB_WORKSPACE   # or Join-Path $PSScriptRoot '..\..' etc, as you prefer

    # Step 1: Download current Drive tracking sheet
    $downloadCmd = @(
        'gam', 'user', $AdminUser, 'get', 'drivefile', $DocID,
        'csvsheet', 'CountParentFolder', 'targetfolder', $trackingDir,
        'targetname', 'CountParentFolder.csv', 'overwrite', 'true'
    )
    & $downloadCmd[0] @($downloadCmd[1..($downloadCmd.Length - 1)]) | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Send-Alert -Subject "[CountParentFolder] Download error" -Body "Could not download CountParentFolder.csv"
        return $CTError
    }

    $DriveTable = Import-Csv "$trackingDir\CountParentFolder.csv" -ErrorAction Stop
    $CurrentDriveID = (@($DriveTable) | Where-Object { $_.IsFull -ne 'TRUE' })[0].DriveID
    if (-not $CurrentDriveID) {
        Send-Alert -Subject "[CountParentFolder] No active drive" -Body "No IsFull=FALSE row found in CountParentFolder.csv"
        return $CTError
    }

    # Step 2: Count current Drive usage
    # (Wrap in @() so .Count is reliable even for a single row)
    $currentCmd = @('gam', 'user', $AdminUser, 'print', 'filelist', 'select', 'teamdriveid', $CurrentDriveID, 'fields', 'id,mimeType')
    $CurrentList = @(& $currentCmd[0] @($currentCmd[1..($currentCmd.Length - 1)]) | ConvertFrom-Csv)
    if ($LASTEXITCODE -ne 0 -or -not $CurrentList) {
        Send-Alert -Subject "[CountParentFolder] Drive scan failed" -Body "GAM failed to list files for DriveID $CurrentDriveID"
        return $CTError
    }

    [int]$CurrentItemCount = @($CurrentList).Count
    [int]$CurrentFolderCount = @($CurrentList | Where-Object { $_.mimeType -eq 'application/vnd.google-apps.folder' }).Count

    # Step 3: Gather next 100 long-term suspended users
    $IncomingUsers = @(Get-LongTermSuspendedAccounts)
    if (-not $IncomingUsers -or @($IncomingUsers).Count -eq 0) {
        Send-Alert -Subject "[CountParentFolder] No incoming users found" -Body "Skipping projection."
        return [PSCustomObject]@{
            ItemPercent   = [math]::Round(($CurrentItemCount / $ItemLimit) * 100, 2)
            FolderPercent = [math]::Round(($CurrentFolderCount / $FolderLimit) * 100, 2)
        }
    }

    # Step 4: Count contents per user
    [int]$IncomingItemCount = 0
    [int]$IncomingFolderCount = 0
    $FileCountOutput = @()

    foreach ($User in $IncomingUsers) {
        $Email = $User.Email
        # NOTE: Convert to array before counting; handle empty/headers-only cases.
        $userCmd = @('gam', 'user', $Email, 'print', 'filelist', 'fields', 'id,mimeType')
        $UserList = @(& $userCmd[0] @($userCmd[1..($userCmd.Length - 1)]) | ConvertFrom-Csv)

        if ($LASTEXITCODE -ne 0 -or -not $UserList) {
            Write-Warning "Could not retrieve files for $Email"
            continue
        }

        [int]$FileCount = @($UserList).Count
        [int]$FolderCount = @($UserList | Where-Object { $_.mimeType -eq 'application/vnd.google-apps.folder' }).Count

        # Use unary comma to ensure array append even when first element
        $FileCountOutput += , ([PSCustomObject]@{
                UserEmail     = $Email
                FileCount     = $FileCount
                DateSuspended = $User.DateSuspended
            })

        $IncomingItemCount += $FileCount
        $IncomingFolderCount += $FolderCount
    }

    # Step 5: Upload UserFileCounts (Sheet #4)
    $UserFileCsv = "$trackingDir\UserFileCounts.csv"
    $FileCountOutput | Export-Csv -Path $UserFileCsv -NoTypeInformation -Force

    $uploadCmd = @(
        'gam', 'user', $AdminUser, 'update', 'drivefile', 'id', '1o4OD7SP5bCFuTaw49hwza3YZgnj_QHx2q2rZi83dALM',
        'retainname', 'localfile', $UserFileCsv
    )
    & $uploadCmd[0] @($uploadCmd[1..($uploadCmd.Length - 1)]) | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Send-Alert -Subject "[CountParentFolder] Upload failed" -Body "Unable to upload UserFileCounts.csv to Sheet #4"
        # continue; not fatal for projection
    }

    # Step 6: Call EstimateCopyTime.ps1
    $EstimateScript = "$trackingDir\runner\Copy\Rotate\EstimateCopyTime.ps1"

    if (Test-Path $EstimateScript) {
        & $EstimateScript `
            -AdminUser $AdminUser `
            -UserFileCountsID "1o4OD7SP5bCFuTaw49hwza3YZgnj_QHx2q2rZi83dALM" `
            -EligibleOutputSheetID "16joJuSmXTeh-JdbFLz18hIbP6cA4_qHJTHQR5l51GE0" `
            -OversizedOutputSheetID "1orGbUWkQmE7ssfHVNI_VbRh1aFyIRgfrRKJ0mJX9OPw"
        if ($LASTEXITCODE -ne 0) {
            Send-Alert -Subject "[CountParentFolder] EstimateCopyTime.ps1 failed" -Body "EstimateCopyTime.ps1 returned non-zero exit code."
        }
    }
    else {
        Send-Alert -Subject "[CountParentFolder] Estimate script missing" -Body "Could not locate EstimateCopyTime.ps1 at $EstimateScript"
    }

    # Step 7: Projected usage
    [int]$IncomingItemCount = $IncomingItemCount
    [int]$IncomingFolderCount = $IncomingFolderCount

    [int]$ProjectedItemCount = ($CurrentItemCount + $IncomingItemCount)
    [int]$ProjectedFolderCount = ($CurrentFolderCount + $IncomingFolderCount)

    return [PSCustomObject]@{
        ItemPercent   = [math]::Round(($ProjectedItemCount / $ItemLimit) * 100, 2)
        FolderPercent = [math]::Round(($ProjectedFolderCount / $FolderLimit) * 100, 2)
    }
}
catch {
    Send-Alert -Subject "[CountParentFolder] Fatal error" -Body "Unexpected error: $($_.Exception.Message)"
    return $CTError
}
