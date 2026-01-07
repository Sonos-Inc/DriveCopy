<#
.SYNOPSIS
    Toggles a user’s archive state and moves them to the correct OU.

.DESCRIPTION
    - Default (no switch): prepares the user for copy
        * Unarchives user
        * Moves to staging OU (/Unused Accounts by default)
    - With -Finalize: finalizes the user after copy
        * Archives user
        * Moves to final OU (/Sonos Inc/Copy/ by default)

.PARAMETER Email
    Target user’s email address.

.PARAMETER AdminEmail
    GAM super admin account.

.PARAMETER Back
    Switch: perform archive + move to final OU instead of /Unused Accounts.

.PARAMETER Fail
    Switch: perform archive + move to failed OU instead of /Unused Accounts.

.EXAMPLE
    .\Set-UserCopyState.ps1 -Email "user@example.com"
    # Unarchives and moves user to /Unused Accounts

.EXAMPLE
    .\Set-UserCopyState.ps1 -Email "user@example.com" -Back
    # Archives and moves user to /Sonos Inc/Copy/
#>
function Set-UserCopyState {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]$Email,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$AdminEmail = 'google.admin@sonos.com',

        [Parameter(Mandatory = $false)]
        [switch]$Back,

        [Parameter(Mandatory = $false)]
        [switch]$Fail
    )

      $gamExe = 'gam'

    if (-not (Get-Command -Name $gamExe -ErrorAction SilentlyContinue)) {
        throw "GAM tool is not found. Please make sure it is installed and accessible in the system PATH."
    }

    function Fail {
        param([string]$Message)
        throw $Message
    }

    if (-not (Get-Command $gamExe -ErrorAction SilentlyContinue)) {
        Fail 'GAM is not installed or not in your PATH.'
    }

    if ($Back -or $Fail) {
        $archiveState = 'on'

        if ($Fail) {
            $targetOU = '/Sonos Inc/Copy Failed/'
        }
        else {
            $targetOU = '/Sonos Inc/Copy/'
        }
    }
    else {
        $archiveState = 'off'
        $targetOU = '/Unused Accounts'
    }

    $gamArgs = @("update", "user", "$Email", "archived", "$archiveState")
    & $gamExe @gamArgs | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Write-Warning "Failed to set archive=$archiveState for $Email (exit $LASTEXITCODE)"
    }

    $gamArgs = @("user", "$Email", "update", "user", "ou", "$targetOU")
    & $gamExe @gamArgs | Out-Null
    if ($LASTEXITCODE -ne 0) {
        Fail "Failed to move $Email to OU $targetOU (exit $LASTEXITCODE)"
    }

    return "Set-UserCopyState complete for $Email (archive=$archiveState, OU=$targetOU)"
}