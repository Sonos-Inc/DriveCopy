function ValidateEmail {
    [CmdletBinding()]
    [OutputType([int])]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [alias("EmailAddress")]
        [string[]]$Email,

        [Parameter(Mandatory = $false, Position = 1)]
        [string]$Requestor,

        [Parameter(Mandatory = $false, Position = 2)]
        [switch]$nr
    )

    try {
        Import-Module '.\Send-alert.psm1' -ErrorAction Stop

        # Validate GAM presence
        if (-not (Get-Command -Name gam -ErrorAction SilentlyContinue)) {
            Send-Alert -Body "GAM not found on system." -Subject "ValidateEmail Error" 
            return 1
        }

        # Default requestor
        if ([string]::IsNullOrWhiteSpace($Requestor) -or -not $nr) {
            $Requestor = "r-google-admin@sonos.com"
        }

        if ([string]::IsNullOrWhiteSpace($Requestor)) {
            return 1
        }

        # Requestor authorization check
        if ($Requestor -ne 'r-google-admin@sonos.com') {
            $rawGroup = & gam print group-members group "GoogleACL@sonos.com" recursive
            if ($LASTEXITCODE -ne 0 -or $null -eq $rawGroup) {
                Send-Alert -Body "Unable to read GoogleACL group for requestor validation." `
                    -Subject "ValidateEmail Error"
                return 1
            }

            $result = $rawGroup | ConvertFrom-Csv | Select-Object -Property email -Unique

            if ($null -eq $result -or $Requestor -notin $result.email) {
                Send-Alert -Body "Unauthorized requestor: $Requestor" `
                    -Subject "ValidateEmail Warning" -EmailTo $Requestor
                return 1
            }
        }

        # RFC-lite regex
        $pattern = '^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'

        # Normalize Email array safely
        if ($null -eq $Email -or $Email.Count -eq 0) {
            return 1
        }

        # Append requestor so we validate them too
        $Email = @($Email + $Requestor)

        foreach ($entry in $Email) {

            if ($null -eq $entry) {
                Send-Alert -Body "Null email entry detected." `
                    -Subject "ValidateEmail Warning" -EmailTo $Requestor
                return 1
            }

            $trimmed = $entry.Trim()
            if ([string]::IsNullOrWhiteSpace($trimmed)) {
                Send-Alert -Body "Empty or whitespace-only email address detected." `
                    -Subject "ValidateEmail Warning" -EmailTo $Requestor
                return 1
            }

            if (-not ($trimmed -match $pattern)) {
                Send-Alert -Body "Invalid email address detected: '$entry'" `
                    -Subject "ValidateEmail Warning" -EmailTo $Requestor
                return 1
            }

            # Check existence via GAM (skip only the service account)
            if ($trimmed -ne 'r-google-admin@sonos.com') {
                & gam whatis $trimmed 2>$null | Out-Null
                if ($LASTEXITCODE -ne 20 -and $LASTEXITCODE -ne 22) {
                    Send-Alert -Body "Non-existent email detected: '$entry'" `
                        -Subject "ValidateEmail Warning" -EmailTo $Requestor
                    return 1
                }
            }
        }

        return 0
    }
    catch {
        $errorMessage = $_.Exception.Message
        Send-Alert -Body "Unexpected error during email validation: $errorMessage" `
            -Subject "ValidateEmail Error" -EmailTo $Requestor
        return 1
    }
}