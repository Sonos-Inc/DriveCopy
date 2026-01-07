<#
.SYNOPSIS
    Updates a specified Google Sheet with new Drive data using GAM.

.DESCRIPTION
    This script downloads an existing Google Sheet as CSV, merges it with new data,
    sorts the result, and re-uploads the updated CSV back to the same Google Sheet.
    It uses GAM for Google Workspace API calls and relies on an alert module for
    error notifications.

.PARAMETER DataArray
    The new data to append (array of objects).

.PARAMETER DocID
    The Google Sheet document ID. Defaults to a known DriveCopy sheet.

.PARAMETER SheetName
    The specific sheet (tab) within the document. Defaults to 'DriveCopy'.

.PARAMETER AdminUser
    The administrative user under which GAM executes. Defaults to 'google.admin@sonos.com'.

.NOTES
    Author: Dan Casmas
    Purpose: Sonos IT automation
    Requires: GAM CLI, Send-Alert.psm1, Get-LegacyDrive.psm1
#>
function Update-GoogleSheet {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [Object[]]$DataArray,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$DocID = "1_0WIDlZriHpb1_YqP1nPGs0lcOhliweGt-klJSPg2C4",

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$SheetName = "DriveCopy",

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$AdminUser = "google.admin@sonos.com"
    )

    $trackingDir = $env:GITHUB_WORKSPACE 

    # Define constants
    $gamExe = 'gam'
    $DownloadPath = "$trackingDir\DriveCopy.csv"
    $UploadPath = "$trackingDir\NewDriveCopy.csv"

    # Cleanup routine to ensure temporary files are deleted
    $cleanup = {
        Remove-Item -Path $UploadPath -Force -ErrorAction SilentlyContinue
        Remove-Item -Path $DownloadPath -Force -ErrorAction SilentlyContinue
    }

    try {
        # Load the alerting module
        Import-Module "$trackingDir\Send-Alert.psm1" -ErrorAction SilentlyContinue

        # Import drive info module and get current drive name
        try {
            Import-Module "$trackingDir\runner\Copy\Get-LegacyDrive.psm1" -Force -ErrorAction Stop
            $currentDrive = Get-LegacyDrive | Select-Object -ExpandProperty Name -First 1
        }
        catch {
            Send-Alert -Subject "UpdateSheet Error" -Body "Failed to get current Legacy Drive. Error: $_"
            & $cleanup
            throw $_
        }

        # --- Download the current Google Sheet as CSV ---
        & $gamExe user $AdminUser get drivefile $DocID csvsheet $SheetName targetfolder "$trackingDir" targetname "DriveCopy.csv" overwrite true 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0 -or -not (Test-Path -Path $DownloadPath)) {
            Send-Alert -Subject "UpdateSheet Error" -Body "Error downloading Google Sheet or file missing."
            & $cleanup
            throw "Download failed"
        }

        # --- Import the old data from the downloaded sheet ---
        try {
            $OldData = Import-Csv -Path $DownloadPath -ErrorAction Stop
        }
        catch {
            Send-Alert -Subject "UpdateSheet Error" -Body "Unable to import DriveCopy.csv. Error: $_"
            & $cleanup
            throw $_
        }

        # --- Append new data with current drive name ---
        $DataArray = $DataArray | Select-Object *, @{Name = 'TeamDriveParentName'; Expression = { $currentDrive } }

        # --- Merge and sort full dataset ---
        $FullData = @($OldData + $DataArray) |
        Where-Object { $_ } |
        Sort-Object -Property TodayDate -Descending

        # --- Export updated data to new CSV ---
        try {
            $FullData | Export-Csv -Path $UploadPath -NoTypeInformation -Encoding UTF8 -Force -ErrorAction Stop
        }
        catch {
            Send-Alert -Subject "UpdateSheet Error" -Body "Failed to export updated CSV. Error: $_"
            & $cleanup
            throw $_
        }

        # --- Upload the updated CSV back to Google Sheet ---
        & $gamExe user $AdminUser update drivefile id $DocID retainname localfile $UploadPath 2>&1 | Out-Null
        if ($LASTEXITCODE -ne 0) {
            Send-Alert -Subject "UpdateSheet Error" -Body "Error uploading updated Google Sheet."
            & $cleanup
            throw "Upload failed"
        }

        # Cleanup on success
        & $cleanup
    }
    catch {
        Send-Alert -Subject "UpdateSheet Error" -Body "Unexpected error during Google Sheet update: $_"
        & $cleanup
        throw $_
    }
}
Export-ModuleMember -Function Update-GoogleSheet