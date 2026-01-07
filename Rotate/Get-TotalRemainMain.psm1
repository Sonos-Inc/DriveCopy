<#
.SYNOPSIS
    Retrieves the remaining Google Workspace licenses from a specified Google Sheet.

.DESCRIPTION
    The Get-TotalRemainMain function uses the GAM tool to download a Google Sheet containing
    license data, parses it for the "TotalRemainMain" value, and subtracts a defined buffer.
    If the resulting license count is below or equal to zero, an alert is sent via Send-Alert.psm1.

.PARAMETER AdminUser
    The administrator user account used by GAM for accessing Google Drive.
    Default is "google.admin@sonos.com".

.PARAMETER DocID
    The Google Sheet document ID containing license data.

.PARAMETER SheetName
    The sheet/tab name within the document to extract.

.PARAMETER LicenseBuffer
    The number of buffer licenses to retain before alerting.
    Must be between 5 and 1000. Default is 10.

.RETURNS
    [int] The adjusted remaining license count, or 0 if an error or low-license condition occurs.

.NOTES
    - Requires GAM installed and accessible in PATH.
    - Requires Send-Alert.psm1 module for alerting.
    - Designed for automation pipelines that rely on exit codes.

.EXAMPLE
    PS> Get-TotalRemainMain -AdminUser "admin@company.com" -DocID "1Abc123..." -SheetName "License Data"
#>
function Get-TotalRemainMain {
    [CmdletBinding()]
    param(
        [ValidateNotNullOrEmpty()]
        [string]$AdminUser = "google.admin@sonos.com",

        [ValidateNotNullOrEmpty()]
        [string]$DocID = "1eBwVTMikAeM74r1uWx1C_wXWJXJMaAIZnZyqGiXDGj4",
        
        [ValidateNotNullOrEmpty()]
        [string]$SheetName = "Google License Numbers",

        [Parameter(Mandatory = $false)]
        [ValidateRange(5, 1000)]
        [int]$LicenseBuffer = 10
    )

    try {
        if (-not (Get-Command -Name gam -ErrorAction SilentlyContinue)) {
            throw "GAM tool is not found. Please make sure it is installed and accessible in the system PATH."
        }

        Import-Module "$env:GITHUB_WORKSPACE\runner\copy\Send-alert.psm1" -ErrorAction Stop

        $trackingDir = $env:GITHUB_WORKSPACE

        $gamExe = 'gam'
        $csvFile = "$trackingDir\LicenseNumbers.csv"
        $gamArgs = @("user", "$AdminUser", "get", "drivefile", "$DocID", "csvsheet", "$SheetName", "targetfolder", "$trackingDir", "targetname", "LicenseNumbers.csv", "overwrite", "true")
        & $gamExe @gamArgs | Out-Null
        if ($LASTEXITCODE -ne 0) {
            # If licenses are low,send alert
            $date = (Get-Date).ToString("yyyy-MM-dd")
            $msg = "Function Get-TotalRemainMain had a problem."
            Send-Alert -Subject "$date Error Function Get-TotalRemainMain" -Body $msg
            return 0
        }

        [int]$Remaining = 0
        $DriveTable = Import-Csv $csvFile
        foreach ($Drive in $DriveTable) {
            if ($Drive.'License Type' -eq "TotalRemainMain") { $Remaining = $($Drive.Total) }
        }

        $Remaining = $Remaining - $LicenseBuffer

        # If licenses are low,send alert
        if ($Remaining -le 0) {
            $date = (Get-Date).ToString("yyyy-MM-dd")
            $msg = "Function Get-TotalRemainMain detected low licenses: $Remaining remaining after buffer."
            Send-Alert -Subject "$date [Low License Alert]" -Body $msg
            return 0
        }

        return $Remaining
    }
    catch {
        $date = (Get-Date).ToString("yyyy-MM-dd")
        $msg = "Function Get-TotalRemainMain had a problem."
        Send-Alert -Subject "$date Error Function Get-TotalRemainMain" -Body $msg
        return 0
    }
    finally {
        Remove-Item $csvFile -Force -ErrorAction SilentlyContinue
    }
}