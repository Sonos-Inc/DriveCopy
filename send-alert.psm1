<#
.SYNOPSIS
    Sends an alert email using configurable SMTP settings.

.DESCRIPTION
    This function sends an HTML-formatted alert email using the provided or default SMTP parameters.
    It requires a subject and body,while allowing customization of recipient,sender,SMTP server,and port.

.PARAMETER Subject
    The subject line of the alert email. This is required.

.PARAMETER Body
    The HTML-formatted body content of the alert email. This is required.

.PARAMETER EmailTo
    The recipient email address. Defaults to 'dan.casmas@sonos.com'.

.PARAMETER From
    The sender email address. Defaults to 'Google Reports@sonos.com'.

.EXAMPLE
    Send-Alert -Subject "Backup Error" -Body "<p>Backup failed on server01.</p>"

.EXAMPLE
    Send-Alert -Subject "Notice" -Body "<p>All systems normal.</p>" -To "admin@company.com" -From "monitor@company.com"

.NOTES
https://sonosinc.atlassian.net/wiki/spaces/PEOPLEPROD/pages/1346830354/Drive+Copy+Automation Architectural
https://sonosinc.atlassian.net/wiki/spaces/ITKB/pages/1348370617/Legacy+Drive+Backup ITKB
https://docs.google.com/spreadsheets/d/1_0WIDlZriHpb1_YqP1nPGs0lcOhliweGt-klJSPg2C4/edit?gid=1966817793#gid=1966817793 Spreadsheet
https://drive.google.com/drive/folders/0ADVPlcaDSgiTUk9PVA Legacy Drive Backup Folder

#>
function Send-Alert {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$Subject = "Alert",

        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$Body,

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$EmailTo = "r-google-admin@sonos.com",

        [Parameter(Mandatory = $false)]
        [ValidateNotNullOrEmpty()]
        [string]$From = "google.admin@sonos.com"
    )

    try {
        $gamCommand = @(
            'user', $From,
            'sendemail', 'to', $EmailTo,
            'subject', $Subject,
            'textmessage', $Body,
            'html'
        )

        $gamexe = 'gam'

        & $gamexe @gamCommand | Out-Null

        if ($LASTEXITCODE -ne 0) {
            throw "GAM command failed with exit code $LASTEXITCODE"
        }
    }
    catch {
        # Output a warning if the mail fails to send
        throw "Failed to send alert email: $_"
    }
}