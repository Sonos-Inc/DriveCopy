<#
.SYNOPSIS
    Retrieves detailed information about long-term suspended accounts using the GAM tool.

.DESCRIPTION
    This function identifies long-term suspended accounts within a specified organizational unit using the GAM tool. 
    It retrieves detailed information about disabled accounts and combines the data for reporting purposes.

.PARAMETER Days
    Specifies the number of days for considering an account as long-term suspended. Default is 31 days.

.PARAMETER ServiceAccount
    Specifies the email address for the service account. Default is 'google.admin@sonos.com'.

.PARAMETER Fail
    Switch to specify Transfer Fail OU instead of default Sonos Inc OU.

.EXAMPLE
    Get-LongTermSuspendedAccounts -Days 60 -ServiceAccount "admin@example.com" -Fail
    Retrieves detailed information about accounts suspended for 60 or more days,considering transfer failures,
    and uses "admin@example.com" as the service account email.

.NOTES
    Ensure that the GAM tool is installed and accessible in the system PATH for this function to work properly.

https://sonosinc.atlassian.net/wiki/spaces/PEOPLEPROD/pages/1346830354/Drive+Copy+Automation Architectural
https://sonosinc.atlassian.net/wiki/spaces/ITKB/pages/1348370617/Legacy+Drive+Backup ITKB
https://docs.google.com/spreadsheets/d/1_0WIDlZriHpb1_YqP1nPGs0lcOhliweGt-klJSPg2C4/edit?gid=1966817793#gid=1966817793 Spreadsheet
https://drive.google.com/drive/folders/0ADVPlcaDSgiTUk9PVA Legacy Drive Backup Folder
#>
function Get-LongTermSuspendedAccounts {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param (
        [Parameter(Mandatory = $false)]
        [ValidateRange(15,1000)]
        [int]$Days = 31,

        [Parameter(Mandatory = $false,HelpMessage = "Email address for the service account.")]
        [ValidatePattern("^[a-zA-Z0-9](?:[a-zA-Z0-9._%+-]*[a-zA-Z0-9])?(?:@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,})?$")]
        [ValidateNotNullOrEmpty()]
        [string]$ServiceAccount = 'google.admin@sonos.com',

        [switch]$Fail
    )
    try {
        # Step 1: Get all suspended accounts not in excluded OUs
        $AccountsSuspend = & gam print users fields orgUnitPath,suspended,relations,primaryEmail,archived | ConvertFrom-Csv | Where-Object {
            ($_.suspended -eq $true -or $_.archived -eq $true ) -and $_.orgUnitPath -notin @(
                '/Sonos Inc/Legal Hold',
                '/Sonos Inc/Copy',
                '/Sonos Inc/Drive Transferred',
                '/Sonos Inc/Drive Transferred/Service Account',
                '/Sonos Inc/Copy/Manual'
            )
        }

        $TodayDate = Get-Date -Format "yyyy-MM-dd"

        # Step 2: Get detailed reports for suspended accounts
        $GoogleReport = & gam report users parameters accounts:drive_used_quota_in_mb,accounts:disabled,accounts:disabled_reason fulldatarequired accounts | ConvertFrom-Csv | Where-Object {
            $_.email -in $AccountsSuspend.primaryEmail
        } | ForEach-Object {
            # Use today's date if disabled_time is missing or 'Never'
            $DateSuspended = if ([string]::IsNullOrWhiteSpace($_.'accounts:disabled_time') -or $_.'accounts:disabled_time' -eq 'Never') {
                $TodayDate
            }
            else {
                (Get-Date -Date $_.'accounts:disabled_time' -Format "yyyy-MM-dd")
            }

            [PSCustomObject]@{
                Email         = $_.email
                DateSuspended = $DateSuspended
                Data          = [int]$_.'accounts:drive_used_quota_in_mb'
                DaysSuspended = [int](New-TimeSpan -Start ([DateTime]::Parse($DateSuspended)) -End ([DateTime]::Parse($TodayDate))).Days
            }
        }

        # Step 3: Filter out accounts with zero data or not suspended long enough
        $GoogleReport = $GoogleReport | Where-Object { $_.Data -ne 0 -and $_.DaysSuspended -ge $Days }

        # Step 4: If -Fail is set,filter for the "Copy Failed" OU; otherwise,use all suspended accounts
        $OuPath = if ($Fail) { '/Sonos Inc/Copy/Copy Failed' } else { $null }
        $OUSuspend = if ($OuPath) {
            $AccountsSuspend | Where-Object { $_.orgUnitPath -eq $OuPath }
        }
        else {
            $AccountsSuspend
        }

        # Step 5: Merge GoogleReport with manager/org info
        $combinedArray = @()
        foreach ($item1 in $OUSuspend) {
            $email = $item1.primaryEmail
            $item2 = $GoogleReport | Where-Object { $_.Email -eq $email }

            if ($item2) {
                # Determine manager; if manager is suspended or missing,use service account
                $ManagerSuspended = $false
                $manager = $item1.'relations.0.value'
                if ([string]::IsNullOrWhiteSpace($manager) -or ($manager -in $AccountsSuspend.primaryEmail)) {
                    $manager = $ServiceAccount
                    $ManagerSuspended = $true
                }

                $combinedArray += [PSCustomObject]@{
                    Email            = $email
                    Manager          = $manager
                    ManagerSuspended = $ManagerSuspended
                    TodaysDate       = $TodayDate
                    orgUnitPath      = $item1.orgUnitPath
                    DateSuspended    = $item2.DateSuspended
                    DaysSuspended    = [int]$item2.DaysSuspended
                }
            }
        }

        # Step 6: Sort and output the final array (unique by Email,then by DaysSuspended)
        $FinalArray = $combinedArray | Sort-Object -Property Email -Unique | Sort-Object -Property DaysSuspended

        return $FinalArray
    }
    catch {
        Write-Error "An error occurred while retrieving long-term suspended accounts: $_"
        return @()
    }
}