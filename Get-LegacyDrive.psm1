<#
.SYNOPSIS
    Retrieves the first available (non-full) legacy Shared Drive from the tracking sheet.

.DESCRIPTION
    The Get-LegacyDrive function uses GAM to download a Google Sheet containing Shared Drive
    tracking data (e.g., drive names, IDs, and fullness flags). It parses the sheet to locate
    the first drive marked as "FALSE" in the 'IsFull' column, returning its ID and name.

    If no available drive is found or GAM fails, the function returns $null and logs an error.

.PARAMETER AdminUser
    The Google Workspace admin user for GAM operations.
    Default: google.admin@sonos.com

.PARAMETER DocID
    The Google Sheet document ID containing Shared Drive tracking information.

.PARAMETER SheetName
    The sheet/tab name within the document that holds the drive data.

.EXAMPLE
    PS> Get-LegacyDrive
    ID                                   Name
    --                                   ----
    0AExampleDriveID12345PVA             Legacydrivebackup3

.EXAMPLE
    PS> Get-LegacyDrive -AdminUser "admin@company.com" -DocID "1Abc123..." -SheetName "DriveTracker"
    # Retrieves the first Shared Drive that is not marked as full.

.NOTES
    - Requires GAM installed and available in PATH.
    - Temporary CSV ('CountParentFolder.csv') is removed after processing.
    - Returns a PSCustomObject or $null if none are available.
#>
function Get-LegacyDrive {
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

    # ---------------------------------------------------------------------
    # 0. Validate GAM installation
    # ---------------------------------------------------------------------
    if (-not (Get-Command -Name gam -ErrorAction SilentlyContinue)) {
        Write-Error "GAM not found in PATH. Please install or update the GAM tool."
        return $null
    }

    # ---------------------------------------------------------------------
    # 1. Download the Google Sheet using GAM
    # ---------------------------------------------------------------------
    $csvFile = "$env:GITHUB_WORKSPACE\CountParentFolder.csv"
    $downloadCmd = @(
        'gam', 'user', $AdminUser,
        'get', 'drivefile', $DocID,
        'csvsheet', $SheetName,
        'targetfolder', "$trackingDir",
        'targetname', 'CountParentFolder.csv',
        'overwrite', 'true'
    )

    & $downloadCmd[0] @($downloadCmd[1..($downloadCmd.Length - 1)]) | Out-Null

    if ($LASTEXITCODE -ne 0) {
        Write-Error "GAM failed to download the sheet. Exit code: $LASTEXITCODE"
        return $null
    }

    # ---------------------------------------------------------------------
    # 2. Validate and import CSV data
    # ---------------------------------------------------------------------
    if (-not (Test-Path -LiteralPath $csvFile)) {
        Write-Error "Expected CSV file not found after GAM execution: $csvFile"
        return $null
    }

    $DriveTable = @()
    try {
        $DriveTable = Import-Csv -Path $csvFile -ErrorAction Stop
    }
    catch {
        Write-Error "Failed to import CSV file $csvFile : $($_.Exception.Message)"
        return $null
    }

    # ---------------------------------------------------------------------
    # 3. Search for the first available (non-full) drive
    # ---------------------------------------------------------------------
    $availableDrive = $DriveTable | Where-Object { $_.IsFull -eq 'FALSE' } | Select-Object -First 1

    if (-not $availableDrive) {
        Write-Warning "No available legacy drives found in $SheetName."
        Remove-Item -LiteralPath $csvFile -Force -ErrorAction SilentlyContinue
        return $null
    }

    $result = [PSCustomObject]@{
        ID   = $availableDrive.DriveID
        Name = $availableDrive.DriveName
    }

    # ---------------------------------------------------------------------
    # 4. Cleanup and return result
    # ---------------------------------------------------------------------
    Remove-Item -LiteralPath $csvFile -Force -ErrorAction SilentlyContinue
    return $result
}

Export-ModuleMember -Function Get-LegacyDrive